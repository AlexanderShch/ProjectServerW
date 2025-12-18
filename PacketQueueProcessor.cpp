#include "PacketQueueProcessor.h"

#include "Commands.h"

#include <msclr/marshal_cppstd.h>

#include <unordered_map>
#include <string>
#include <sstream>
#include <iomanip>
// Marshal is used only to bridge between managed queue payload and existing std::wstring GUID map.

using namespace System;
using namespace System::Runtime::InteropServices;
using namespace System::Threading;

namespace ProjectServerW {
	struct TelemetryDuplicateGuardState
	{
		std::string lastBadFrame;
	};

	static std::unordered_map<UINT_PTR, TelemetryDuplicateGuardState> s_telemetryDuplicateGuardBySocket;

	static std::string BytesToHex(const uint8_t* data, int length)
	{
		// Why: log raw frames to diagnose routing/CRC/length issues without attaching binary dumps.
		if (data == nullptr || length <= 0) {
			return std::string();
		}

		std::ostringstream oss;
		oss << std::hex << std::uppercase << std::setfill('0');
		for (int i = 0; i < length; i++) {
			if (i != 0) {
				oss << ' ';
			}
			oss << std::setw(2) << static_cast<unsigned int>(data[i]);
		}
		return oss.str();
	}
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

				//  ритично: контроллер при DATA_FALSE повтор€ет пакет. ≈сли сервер продолжит слать DATA_FALSE на неизменившийс€ пакет,
				// можно получить бесконечную "карусель" повторов.
				// ѕолитика сервера: если пакет телеметрии с неверным CRC пришЄл повторно и байты кадра не изменились,
				// отправл€ем DATA_OK и просто отбрасываем пакет (без доставки в UI).
				const UINT_PTR socketKey = static_cast<UINT_PTR>(socket);
				TelemetryDuplicateGuardState& guard = s_telemetryDuplicateGuardBySocket[socketKey];

				if (!crcOk) {
					const std::string frameBytes(reinterpret_cast<const char*>(raw), static_cast<size_t>(item->size));
					if (!guard.lastBadFrame.empty() && guard.lastBadFrame == frameBytes) {
						TrySendTelemetryAck(socket, CmdTelemetry::DATA_OK);
						guard.lastBadFrame.clear();
						continue;
					}

					guard.lastBadFrame = frameBytes;
					TrySendTelemetryAck(socket, CmdTelemetry::DATA_FALSE);
					continue;
				}

				guard.lastBadFrame.clear();
				TrySendTelemetryAck(socket, CmdTelemetry::DATA_OK);

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

		//  ритично: метод должен быть строго O(1), потому что вызываетс€ из recv()-цикла.
		responseQueue->Enqueue(response);
		responseAvailable->Release();
	}

	bool DataForm::ReceiveResponse(CommandResponse& response, int timeoutMs)
	{
		cli::array<System::Byte>^ ignored = nullptr;
		return ReceiveResponse(response, timeoutMs, ignored);
	}

	bool DataForm::ReceiveResponse(CommandResponse& response, int timeoutMs, cli::array<System::Byte>^% rawResponse)
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
			// Why: this buffer is already removed from the queue; log it to explain why it was dropped.
			GlobalLogger::LogMessage(ConvertToStdString(String::Format(
				"Warning: Dropping invalid response frame (len={0})", responseBuffer->Length)));
			return false;
		}

		uint8_t buffer[MAX_COMMAND_SIZE];
		pin_ptr<System::Byte> pinnedBuffer = &responseBuffer[0];
		memcpy(buffer, pinnedBuffer, responseBuffer->Length);

		rawResponse = responseBuffer;

		const bool ok = ParseResponseBuffer(buffer, static_cast<size_t>(responseBuffer->Length), response);
		if (!ok) {
			GlobalLogger::LogMessage(ConvertToStdString(String::Format(
				"Warning: Dropping unparseable response frame (len={0}): {1}",
				responseBuffer->Length,
				gcnew String(BytesToHex(buffer, responseBuffer->Length).c_str()))));
			ScheduleCommandInfoProbe("dropped unparseable response frame");
		}
		return ok;
	}

	bool DataForm::ReceiveResponse(CommandResponse& response)
	{
		return ReceiveResponse(response, 1000);
	}

}

