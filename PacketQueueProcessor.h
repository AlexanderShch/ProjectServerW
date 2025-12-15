#pragma once

#include "SServer.h" // Provides SOCKET and DataForm declarations via existing include chain

namespace ProjectServerW {
	// This type exists to keep the recv() loop minimal and predictable:
	// the network thread only enqueues complete packets; processing happens on a background worker.
	public ref class PacketQueueProcessor abstract sealed
	{
	private:
		ref class TelemetryWorkItem sealed
		{
		public:
			cli::array<System::Byte>^ packet;
			int size;
			int port;
			System::String^ formGuid;
			System::String^ clientIP;
			System::IntPtr clientSocket;
		};

		static initonly System::Collections::Concurrent::ConcurrentQueue<TelemetryWorkItem^>^ s_telemetryQueue =
			gcnew System::Collections::Concurrent::ConcurrentQueue<TelemetryWorkItem^>();
		// Critical: recv() thread must never observe SemaphoreFullException from EnqueueTelemetry->Release().
		static initonly System::Threading::Semaphore^ s_telemetryAvailable =
			gcnew System::Threading::Semaphore(0, System::Int32::MaxValue);
		static initonly System::Object^ s_workerSync = gcnew System::Object();
		static System::Threading::Thread^ s_workerThread;
		static bool s_workerStarted = false;

		// Critical: both the telemetry worker and UI thread can send on the same socket.
		// A per-socket gate prevents byte-stream interleaving.
		static initonly System::Collections::Concurrent::ConcurrentDictionary<System::IntPtr, System::Object^>^ s_sendGates =
			gcnew System::Collections::Concurrent::ConcurrentDictionary<System::IntPtr, System::Object^>();

		static void EnsureWorker();
		static void WorkerLoop();
		static System::Object^ GetOrCreateSendGate(System::IntPtr socketKey);
		static bool TrySendTelemetryAck(SOCKET socket, uint8_t code);
		static bool ValidateTelemetryCrc(const uint8_t* data, int size);

	public:
		// Critical: telemetry ACKs and UI-originated commands share the same socket.
		// Serializing send() per-socket prevents interleaving at the TCP stream level.
		static System::Object^ GetSendGate(SOCKET clientSocket);

		static void EnqueueTelemetry(cli::array<System::Byte>^ packet,
			int size,
			int port,
			System::String^ formGuid,
			SOCKET clientSocket,
			System::String^ clientIP);
	};
}

