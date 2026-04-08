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

	Object^ PacketQueueProcessor::GetOrCreateReceivingGate(IntPtr socketKey)
	{
		Object^ gate = nullptr;
		if (s_receivingGates->TryGetValue(socketKey, gate)) {
			return gate;
		}
		Object^ candidate = gcnew Object();
		if (s_receivingGates->TryAdd(socketKey, candidate)) {
			return candidate;
		}
		if (s_receivingGates->TryGetValue(socketKey, gate) && gate != nullptr) {
			return gate;
		}
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
		// Полудуплекс: не отправлять, пока recv()-поток обрабатывает принятые данные.
		Object^ recvGate = GetOrCreateReceivingGate(socketKey);
		Monitor::Enter(recvGate);
		try {
			Object^ sendGate = GetOrCreateSendGate(socketKey);
			Monitor::Enter(sendGate);
			try {
				const int bytesSent = send(socket, reinterpret_cast<const char*>(ackBuffer), static_cast<int>(ackSize), 0);
				return bytesSent != SOCKET_ERROR;
			}
			finally {
				Monitor::Exit(sendGate);
			}
		}
		finally {
			Monitor::Exit(recvGate);
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
			// Ожидание наступления события в очереди телеметрии.
			/* Что такое s_telemetryAvailable?
				Статический семафор Semaphore(0, Int32::MaxValue) — изначально счётчик 0, то есть воркер сразу блокируется на WaitOne().		
				Semaphore(0, Int32::MaxValue) — изначально счётчик 0, то есть воркер сразу блокируется на WaitOne().
			Когда разблокируется?
				После EnqueueTelemetry / EnqueueControlLog: пакет кладут в s_telemetryQueue, затем вызывают s_telemetryAvailable->Release() 
				— счётчик увеличивается, один ожидающий WaitOne() проходит.
			Зачем?	
				Поток приёма (recv) только ставит задачи в очередь; WorkerLoop в бесконечном цикле:
				ждёт сигнал семафора;
				выгребает всё из s_telemetryQueue через TryDequeue и обрабатывает (CRC, ACK, UI и т.д.).
			*/
			s_telemetryAvailable->WaitOne();

			// Обработка элементов из очереди телеметрии.
			TelemetryWorkItem^ item = nullptr;
			/* TryDequeue — это метод потокобезопасной очереди ConcurrentQueue<T> в .NET (у вас в PacketQueueProcessor.h: ConcurrentQueue<TelemetryWorkItem^>).
				Что делает: пытается снять один элемент с начала очереди (FIFO).
				Если очередь не пуста — забирает первый элемент, записывает его в переданный аргумент (item), возвращает true.
				Если очередь пуста — ничего не извлекает, item обычно остаётся как был (или сбрасывается в зависимости от реализации вызова), возвращает false.
			*/
			while (s_telemetryQueue->TryDequeue(item)) {
				if (item == nullptr || item->packet == nullptr) {	// Если элемент не найден или пакет пуст, пропускаем.
					continue;
				}

				if (item->size <= 0 || item->packet->Length < item->size) {	// Если размер пакета некорректный, пропускаем.
					continue;
				}

				// Пакет лога (Type 0x01): доставка на форму в потоке UI; строка для лога берётся из очереди rowsPendingLog (1:1 с телеметрией).
				if (item->itemType == 1) {
					if (String::IsNullOrEmpty(item->formGuid)) {	// Если GUID пуст, логируем предупреждение и пропускаем.
						GlobalLogger::LogMessage("Предупреждение: пакет лога (itemType=1): formGuid пуст, запись отброшена");
					}
					else {
						// Преобразование GUID из строки в wstring для поиска формы.
						msclr::interop::marshal_context ctx;	// Контекст для маршалинга строк.
						const std::wstring guidW = ctx.marshal_as<std::wstring>(item->formGuid);
						DataForm^ form = DataForm::GetFormByGuid(guidW);	// Поиск формы по GUID.
						if (form == nullptr) {	// Если форма не найдена, логируем предупреждение и пропускаем.
							GlobalLogger::LogMessage("Предупреждение: пакет лога: GetFormByGuid вернул null, запись отброшена");
						}
						else if (form->IsDisposed || form->Disposing || !form->IsHandleCreated) {	// Если форма уничтожена, пропускаем.
							GlobalLogger::LogMessage(String::Format(
								"Предупреждение: пакет лога: форма недействительна (IsDisposed={0}, Disposing={1}, IsHandleCreated={2}), запись отброшена",
								form->IsDisposed, form->Disposing, form->IsHandleCreated));
						}
						else {
							// Вызов метода AppendControlLogToDataRow на форме в потоке UI.
							form->BeginInvoke(
								gcnew Action<cli::array<System::Byte>^, int>(form, &DataForm::AppendControlLogToDataRow),
								item->packet,
								item->size);
						}
					}
					continue;
				}

				// Пакет телеметрии: CRC и доставка в UI.
				// WinSock SOCKET — целочисленный дескриптор; IntPtr хранит его как pointer-sized, поэтому reinterpret_cast.
				const SOCKET clientSock = reinterpret_cast<SOCKET>(item->clientSocket.ToPointer());

				pin_ptr<System::Byte> pinned = &item->packet[0];	// Получение указателя на пакет.
				const uint8_t* raw = reinterpret_cast<const uint8_t*>(pinned);	// Получение указателя на пакет в виде uint8_t*.

				const bool crcOk = ValidateTelemetryCrc(raw, item->size);	// Проверка CRC пакета.

				// Критично: контроллер при DATA_FALSE повторяет пакет. Если сервер продолжит слать DATA_FALSE на неизменившийся пакет,
				// можно получить бесконечную "карусель" повторов.
				// Политика сервера: если пакет телеметрии с неверным CRC пришёл повторно и байты кадра не изменились,
				// отправляем DATA_OK и просто отбрасываем пакет (без доставки в UI).
				const UINT_PTR socketKey = static_cast<UINT_PTR>(clientSock);	// Преобразование SOCKET в UINT_PTR.
				TelemetryDuplicateGuardState& guard = s_telemetryDuplicateGuardBySocket[socketKey];	// Получение состояния дубликата пакета для сокета.
				/*TelemetryDuplicateGuardState хранит одно поле: lastBadFrame — последний кадр телеметрии, у которого не сошёлся CRC (сырые байты пакета).
				* operator[] у unordered_map: по ключу socketKey либо находит уже существующую запись, либо создаёт новую (пустой TelemetryDuplicateGuardState).
				* Ссылка guard — это работа именно с этой записью для данного сокета, без лишнего копирования структуры.
				*/
				if (!crcOk) {	// Если CRC не сошёлся, пропускаем.
					const std::string frameBytes(reinterpret_cast<const char*>(raw), static_cast<size_t>(item->size));
					if (!guard.lastBadFrame.empty() && guard.lastBadFrame == frameBytes) {	
						// Если последний кадр не пуст и равен текущему (с битым CRC), отправляем DATA_OK и очищаем состояние дубликата.
						TrySendTelemetryAck(clientSock, CmdTelemetry::DATA_OK);
						guard.lastBadFrame.clear();
						continue;
					}

					guard.lastBadFrame = frameBytes;	// Обновляем последний кадр с битым CRC.
					TrySendTelemetryAck(clientSock, CmdTelemetry::DATA_FALSE);	// Отправляем DATA_FALSE, будем ждать повторный пакет
					continue;
				}

				guard.lastBadFrame.clear();	// Очищаем состояние дубликата.
				TrySendTelemetryAck(clientSock, CmdTelemetry::DATA_OK);	// Отправляем DATA_OK.

				/* ОБРАБАТЫВАЕМ ПАКЕТ ТЕЛЕМЕТРИИ ------------------------------------------------------------
				ДОСТАВЛЯЕМ В UI
				*/
				if (String::IsNullOrEmpty(item->formGuid)) {	// Если GUID пуст, т.е. нет формы для доставки, пропускаем.
					continue;
				}

				msclr::interop::marshal_context ctx;	// Контекст для маршалинга строк.
				const std::wstring guidW = ctx.marshal_as<std::wstring>(item->formGuid);	// Преобразование GUID из строки в wstring.
				DataForm^ form = DataForm::GetFormByGuid(guidW);	// Поиск формы по GUID.
				if (form == nullptr || form->IsDisposed || form->Disposing || !form->IsHandleCreated) {	// Если форма не найдена или уничтожена, пропускаем.
					continue;
				}

				// Обновляем соединение формы с сокетом и IP клиента для отправки команд.
				form->ClientSocket = clientSock;	// Сохраняем сокет формы.
				form->ClientIP = item->clientIP;	// Сохраняем IP клиента.

				// Изменение DataTable/UI должно выполняться в потоке формы.
				form->BeginInvoke(
					// Вызов метода AddDataToTableThreadSafe на форме в потоке UI.
					gcnew Action<cli::array<System::Byte>^, int, int>(form, &DataForm::AddDataToTableThreadSafe),
					item->packet,	// Пакет телеметрии.
					item->size,	// Размер пакета.
					item->port);	// Порт.
			}
		}
	}

	// Добавляем пакет телеметрии в очередь.
	void PacketQueueProcessor::EnqueueTelemetry(cli::array<System::Byte>^ packet,
		int size,
		int port,
		System::String^ formGuid,
		SOCKET clientSocket,
		System::String^ clientIP)
	{
		if (packet == nullptr || size <= 0) {	// Если пакет пуст или размер некорректный, пропускаем.
			return;
		}

		// EnsureWorker гарантирует, что фоновый поток обработки очереди телеметрии запущен ровно один раз.
		EnsureWorker();	// Запускаем воркер, если он не запущен.

		TelemetryWorkItem^ item = gcnew TelemetryWorkItem();	// Создаём новый элемент очереди.
		item->itemType = 0;	// Тип элемента - телеметрия.
		item->packet = packet;	// Пакет телеметрии.
		item->size = size;	// Размер пакета.
		item->port = port;	// Порт.
		item->formGuid = formGuid;	// GUID формы.
		item->clientIP = clientIP;	// IP клиента.
		item->clientSocket = IntPtr(reinterpret_cast<void*>(clientSocket));	// Сокет клиента.

		s_telemetryQueue->Enqueue(item);	// Добавляем элемент в очередь.
		s_telemetryAvailable->Release();	// Увеличиваем счётчик семафора, чтобы один из ожидающих WaitOne() прошёл.
	}

	// Добавляем пакет лога в очередь.
	void PacketQueueProcessor::EnqueueControlLog(cli::array<System::Byte>^ packet,
		int size,
		int port,
		System::String^ formGuid,
		SOCKET clientSocket,
		System::String^ clientIP)
	{
		if (packet == nullptr || size <= 0) {	// Если пакет пуст или размер некорректный, пропускаем.
			return;
		}
		EnsureWorker();	// Запускаем воркер, если он не запущен.
		TelemetryWorkItem^ item = gcnew TelemetryWorkItem();	// Создаём новый элемент очереди.
		item->itemType = 1;
		item->packet = packet;	// Пакет лога.
		item->size = size;	// Размер пакета.
		item->port = port;	// Порт.
		item->formGuid = formGuid;	// GUID формы.
		item->clientIP = clientIP;	// IP клиента.
		item->clientSocket = IntPtr(reinterpret_cast<void*>(clientSocket));	// Сокет клиента.
		s_telemetryQueue->Enqueue(item);	// Добавляем элемент в очередь.
		s_telemetryAvailable->Release();	// Увеличиваем счётчик семафора, чтобы один из ожидающих WaitOne() прошёл.
	}

	// Получаем или создаём семафор для отправки команд на сокет.
	System::Object^ PacketQueueProcessor::GetSendGate(SOCKET clientSocket)
	{
		return GetOrCreateSendGate(IntPtr(reinterpret_cast<void*>(clientSocket)));
	}

	// Получаем или создаём семафор для приёма команд от сокета.
	System::Object^ PacketQueueProcessor::GetReceivingGate(SOCKET clientSocket)
	{
		return GetOrCreateReceivingGate(IntPtr(reinterpret_cast<void*>(clientSocket)));
	}

	// ============================
	// Command response queue (moved out of recv loop translation unit)
	// ============================

	// Добавляем ответ команды в очередь.
	void DataForm::EnqueueResponse(cli::array<System::Byte>^ response)
	{
		if (response == nullptr) {	// Если ответ пуст, пропускаем.
			return;
		}

		// Критично: метод должен быть строго O(1), потому что вызывается из recv()-цикла.
		responseQueue->Enqueue(response);	// Добавляем ответ в очередь.
		responseAvailable->Release();	// Увеличиваем счётчик семафора, чтобы один из ожидающих WaitOne() прошёл.
	}

	// Получаем ответ команды из очереди.
	bool DataForm::ReceiveResponse(CommandResponse& response, int timeoutMs)
	{
		cli::array<System::Byte>^ ignored = nullptr;	// Игнорируемый пакет.
		return ReceiveResponse(response, timeoutMs, ignored);
	}

	// Получаем ответ команды из очереди.
	bool DataForm::ReceiveResponse(CommandResponse& response, int timeoutMs, cli::array<System::Byte>^% rawResponse)
	{
		if (clientSocket == INVALID_SOCKET) {	// Если сокет недействительный, пропускаем.
			return false;
		}

		if (!responseAvailable->WaitOne(timeoutMs)) {	// Если нет ответа по тайм-ауту, пропускаем.
			return false;
		}

		cli::array<System::Byte>^ responseBuffer = nullptr;
		if (!responseQueue->TryDequeue(responseBuffer) || responseBuffer == nullptr) {	// Если нет ответа из очереди, пропускаем.
			return false;
		}

		if (responseBuffer->Length <= 0 || responseBuffer->Length > static_cast<int>(MAX_COMMAND_SIZE)) {	// Если размер пакета некорректный, пропускаем.
			// Why: this buffer is already removed from the queue; log it to explain why it was dropped.
			GlobalLogger::LogMessage(String::Format(
				"Предупреждение: отбрасываем некорректный ответный пакет (len={0})", responseBuffer->Length));
			return false;
		}

		uint8_t buffer[MAX_COMMAND_SIZE];	// Буфер для ответа.
		pin_ptr<System::Byte> pinnedBuffer = &responseBuffer[0];	// Получение указателя на пакет в виде uint8_t*.
		memcpy(buffer, pinnedBuffer, responseBuffer->Length);	// Копирование данных из пакета в буфер.

		rawResponse = responseBuffer;	// Сохранение пакета в rawResponse.

		// Разбор пакета ответа.
		const bool ok = ParseResponseBuffer(buffer, static_cast<size_t>(responseBuffer->Length), response);	// Разбор пакета ответа.
		if (!ok) {	// Если разбор не удался, логируем предупреждение и пропускаем.
			GlobalLogger::LogMessage(String::Format(
				"Предупреждение: отбрасываем неразборный ответный пакет (len={0}): {1}",
				responseBuffer->Length,
				gcnew String(BytesToHex(buffer, responseBuffer->Length).c_str())));
		}
		return ok;	// Возвращаем результат разбора.
	}

	// Получаем ответ команды из очереди.
	bool DataForm::ReceiveResponse(CommandResponse& response)
	{
		return ReceiveResponse(response, 1000);	// Получаем ответ команды из очереди с тайм-аутом 1000 мс.
	}

}

