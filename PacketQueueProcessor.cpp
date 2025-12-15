#include "PacketQueueProcessor.h"

#include "Commands.h"

#include <msclr/marshal_cppstd.h>
// Marshal is used only to bridge between managed queue payload and existing std::wstring GUID map.

using namespace System;
using namespace System::Runtime::InteropServices;
using namespace System::Threading;

namespace ProjectServerW {
	Object^ PacketQueueProcessor::GetOrCreateSendGate(IntPtr socketKey)
	{
		Object^ gate = nullptr;
		if (s_sendGates->TryGetValue(socketKey, gate)) {
			return gate;
		}

		Object^ candidate = gcnew Object();
		if (s_sendGates->TryAdd(socketKey, candidate)) {
			return candidate;
		}

		// Another thread won the race.
		if (s_sendGates->TryGetValue(socketKey, gate) && gate != nullptr) {
			return gate;
		}

		// Last resort: never return null (callers use it for locking).
		return candidate;
	}

	void PacketQueueProcessor::EnsureWorker()
	{
		Monitor::Enter(s_workerSync);
		try {
			if (s_workerStarted && s_workerThread != nullptr) {
				return;
			}

			s_workerThread = gcnew Thread(gcnew ThreadStart(&PacketQueueProcessor::WorkerLoop));
			s_workerThread->IsBackground = true;
			s_workerThread->Start();
			s_workerStarted = true;
		}
		finally {
			Monitor::Exit(s_workerSync);
		}
	}

	bool PacketQueueProcessor::TrySendTelemetryAck(SOCKET socket, uint8_t code)
	{
		Command ackCmd = CreateTelemetryAckCommand(code);
		uint8_t ackBuffer[MAX_COMMAND_SIZE];
		size_t ackSize = BuildCommandBuffer(ackCmd, ackBuffer, sizeof(ackBuffer));
		if (ackSize == 0) {
			return false;
		}

		const IntPtr socketKey(reinterpret_cast<void*>(socket));
		Object^ gate = GetOrCreateSendGate(socketKey);

		Monitor::Enter(gate);
		try {
			const int bytesSent = send(socket, reinterpret_cast<const char*>(ackBuffer), static_cast<int>(ackSize), 0);
			return bytesSent != SOCKET_ERROR;
		}
		finally {
			Monitor::Exit(gate);
		}
	}

	bool PacketQueueProcessor::ValidateTelemetryCrc(const uint8_t* data, int size)
	{
		// Telemetry CRC is evaluated outside the recv() loop to keep the network thread fast and deterministic.
		if (data == nullptr || size < 3) {
			return false;
		}

		const int crcOffset = size - 2;
		uint16_t received = 0;
		memcpy(&received, data + crcOffset, 2);

		const uint16_t calculated = CalculateCommandCRC(data, static_cast<size_t>(crcOffset));
		return received == calculated;
	}

	void PacketQueueProcessor::WorkerLoop()
	{
		while (true) {
			s_telemetryAvailable->WaitOne();

			TelemetryWorkItem^ item = nullptr;
			while (s_telemetryQueue->TryDequeue(item)) {
				if (item == nullptr || item->packet == nullptr) {
					continue;
				}

				if (item->size <= 0 || item->packet->Length < item->size) {
					continue;
				}

				// WinSock SOCKET is an integer handle (UINT_PTR). IntPtr stores it as a pointer-sized value,
				// so we must use reinterpret_cast instead of static_cast when converting from void*.
				const SOCKET socket = reinterpret_cast<SOCKET>(item->clientSocket.ToPointer());

				pin_ptr<System::Byte> pinned = &item->packet[0];
				const uint8_t* raw = reinterpret_cast<const uint8_t*>(pinned);

				const bool crcOk = ValidateTelemetryCrc(raw, item->size);
				const uint8_t ackCode = crcOk ? CmdTelemetry::DATA_OK : CmdTelemetry::DATA_FALSE;

				// ACK is sent regardless of whether UI is still alive; it is part of device protocol behavior.
				TrySendTelemetryAck(socket, ackCode);

				if (!crcOk) {
					continue;
				}

				if (String::IsNullOrEmpty(item->formGuid)) {
					continue;
				}

				msclr::interop::marshal_context ctx;
				const std::wstring guidW = ctx.marshal_as<std::wstring>(item->formGuid);
				DataForm^ form = DataForm::GetFormByGuid(guidW);
				if (form == nullptr || form->IsDisposed || form->Disposing || !form->IsHandleCreated) {
					continue;
				}

				// Keep per-form connection details fresh for command sending.
				form->ClientSocket = socket;
				form->ClientIP = item->clientIP;

				// DataTable/UI mutation must run on the form thread.
				form->BeginInvoke(
					gcnew Action<cli::array<System::Byte>^, int, int>(form, &DataForm::AddDataToTableThreadSafe),
					item->packet,
					item->size,
					item->port);
			}
		}
	}

	void PacketQueueProcessor::EnqueueTelemetry(cli::array<System::Byte>^ packet,
		int size,
		int port,
		System::String^ formGuid,
		SOCKET clientSocket,
		System::String^ clientIP)
	{
		if (packet == nullptr || size <= 0) {
			return;
		}

		EnsureWorker();

		TelemetryWorkItem^ item = gcnew TelemetryWorkItem();
		item->packet = packet;
		item->size = size;
		item->port = port;
		item->formGuid = formGuid;
		item->clientIP = clientIP;
		item->clientSocket = IntPtr(reinterpret_cast<void*>(clientSocket));

		s_telemetryQueue->Enqueue(item);
		s_telemetryAvailable->Release();
	}

	System::Object^ PacketQueueProcessor::GetSendGate(SOCKET clientSocket)
	{
		return GetOrCreateSendGate(IntPtr(reinterpret_cast<void*>(clientSocket)));
	}

	// ============================
	// Command response queue (moved out of recv loop translation unit)
	// ============================

	void DataForm::EnqueueResponse(cli::array<System::Byte>^ response)
	{
		if (response == nullptr) {
			return;
		}

		// Keep this method strictly O(1): it is called from the recv() loop.
		responseQueue->Enqueue(response);
		responseAvailable->Release();
	}

	bool DataForm::ReceiveResponse(CommandResponse& response, int timeoutMs)
	{
		if (clientSocket == INVALID_SOCKET) {
			return false;
		}

		if (!responseAvailable->WaitOne(timeoutMs)) {
			return false;
		}

		cli::array<System::Byte>^ responseBuffer = nullptr;
		if (!responseQueue->TryDequeue(responseBuffer) || responseBuffer == nullptr) {
			return false;
		}

		if (responseBuffer->Length <= 0 || responseBuffer->Length > static_cast<int>(MAX_COMMAND_SIZE)) {
			return false;
		}

		uint8_t buffer[MAX_COMMAND_SIZE];
		pin_ptr<System::Byte> pinnedBuffer = &responseBuffer[0];
		memcpy(buffer, pinnedBuffer, responseBuffer->Length);

		return ParseResponseBuffer(buffer, static_cast<size_t>(responseBuffer->Length), response);
	}

	bool DataForm::ReceiveResponse(CommandResponse& response)
	{
		return ReceiveResponse(response, 1000);
	}

}

