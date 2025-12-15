#include "SServer.h" // Включает DataForm.h внутри себя
#include "PacketQueueProcessor.h"
#include "Commands.h" // MAX_COMMAND_SIZE is used for framed payload length heuristics.
#include <fstream>
#include <iostream>
#include <chrono>
#include <ctime>
#include <msclr/marshal_cppstd.h>
#include <cstdio>

// Добавьте эту строку для доступа к Marshal
using namespace System::Runtime::InteropServices;

using namespace System::Windows::Forms;
using namespace ProjectServerW;

std::map<std::wstring, std::thread> formThreads;  // Хранение потоков форм

/*CRC16-CITT tables*/
const static uint16_t crc16_table[] =
{
  0x0000, 0xc0c1, 0xc181, 0x0140, 0xc301, 0x03c0, 0x0280, 0xc241,0xc601, 0x06c0, 0x0780, 0xc741, 0x0500, 0xc5c1, 0xc481, 0x0440,
  0xcc01, 0x0cc0, 0x0d80, 0xcd41, 0x0f00, 0xcfc1, 0xce81, 0x0e40,0x0a00, 0xcac1, 0xcb81, 0x0b40, 0xc901, 0x09c0, 0x0880, 0xc841,
  0xd801, 0x18c0, 0x1980, 0xd941, 0x1b00, 0xdbc1, 0xda81, 0x1a40,0x1e00, 0xdec1, 0xdf81, 0x1f40, 0xdd01, 0x1dc0, 0x1c80, 0xdc41,
  0x1400, 0xd4c1, 0xd581, 0x1540, 0xd701, 0x17c0, 0x1680, 0xd641,0xd201, 0x12c0, 0x1380, 0xd341, 0x1100, 0xd1c1, 0xd081, 0x1040,
  0xf001, 0x30c0, 0x3180, 0xf141, 0x3300, 0xf3c1, 0xf281, 0x3240,0x3600, 0xf6c1, 0xf781, 0x3740, 0xf501, 0x35c0, 0x3480, 0xf441,
  0x3c00, 0xfcc1, 0xfd81, 0x3d40, 0xff01, 0x3fc0, 0x3e80, 0xfe41,0xfa01, 0x3ac0, 0x3b80, 0xfb41, 0x3900, 0xf9c1, 0xf881, 0x3840,
  0x2800, 0xe8c1, 0xe981, 0x2940, 0xeb01, 0x2bc0, 0x2a80, 0xea41,0xee01, 0x2ec0, 0x2f80, 0xef41, 0x2d00, 0xedc1, 0xec81, 0x2c40,
  0xe401, 0x24c0, 0x2580, 0xe541, 0x2700, 0xe7c1, 0xe681, 0x2640,0x2200, 0xe2c1, 0xe381, 0x2340, 0xe101, 0x21c0, 0x2080, 0xe041,
  0xa001, 0x60c0, 0x6180, 0xa141, 0x6300, 0xa3c1, 0xa281, 0x6240,0x6600, 0xa6c1, 0xa781, 0x6740, 0xa501, 0x65c0, 0x6480, 0xa441,
  0x6c00, 0xacc1, 0xad81, 0x6d40, 0xaf01, 0x6fc0, 0x6e80, 0xae41,0xaa01, 0x6ac0, 0x6b80, 0xab41, 0x6900, 0xa9c1, 0xa881, 0x6840,
  0x7800, 0xb8c1, 0xb981, 0x7940, 0xbb01, 0x7bc0, 0x7a80, 0xba41,0xbe01, 0x7ec0, 0x7f80, 0xbf41, 0x7d00, 0xbdc1, 0xbc81, 0x7c40,
  0xb401, 0x74c0, 0x7580, 0xb541, 0x7700, 0xb7c1, 0xb681, 0x7640,0x7200, 0xb2c1, 0xb381, 0x7340, 0xb101, 0x71c0, 0x7080, 0xb041,
  0x5000, 0x90c1, 0x9181, 0x5140, 0x9301, 0x53c0, 0x5280, 0x9241,0x9601, 0x56c0, 0x5780, 0x9741, 0x5500, 0x95c1, 0x9481, 0x5440,
  0x9c01, 0x5cc0, 0x5d80, 0x9d41, 0x5f00, 0x9fc1, 0x9e81, 0x5e40,0x5a00, 0x9ac1, 0x9b81, 0x5b40, 0x9901, 0x59c0, 0x5880, 0x9841,
  0x8801, 0x48c0, 0x4980, 0x8941, 0x4b00, 0x8bc1, 0x8a81, 0x4a40,0x4e00, 0x8ec1, 0x8f81, 0x4f40, 0x8d01, 0x4dc0, 0x4c80, 0x8c41,
  0x4400, 0x84c1, 0x8581, 0x4540, 0x8701, 0x47c0, 0x4680, 0x8641,0x8201, 0x42c0, 0x4380, 0x8341, 0x4100, 0x81c1, 0x8081, 0x4040,
};

// Определение конструктора класса SServer
SServer::SServer() : port(0), this_s(INVALID_SOCKET) {
}

// Определение деструктора класса SServer
SServer::~SServer() {
    if (this_s != INVALID_SOCKET) {
        closesocket(this_s);
    }
    WSACleanup();
	GlobalLogger::Shutdown();
}

void SServer::startServer() {
	MyForm^ form = safe_cast<MyForm^>(Application::OpenForms["MyForm"]);
	GlobalLogger::Initialize();

	if (WSAStartup(MAKEWORD(2, 2), &wData) == 0) {
		form->SetWSA_TextValue("WSA Startup success");
	} else {
		form->SetWSA_TextValue("WSA Startup failed: " + WSAGetLastError());
        return;
    }

	SOCKADDR_IN addr{};
	int addrl = sizeof(addr);
	addr.sin_addr.S_un.S_addr = INADDR_ANY;
	addr.sin_port = htons(port);
	addr.sin_family = AF_INET;

	this_s = socket(AF_INET, SOCK_STREAM, 0);
	if (this_s == SOCKET_ERROR) {
		form->SetSocketState_TextValue("Error when creating a socket");
	} else {
		form->SetSocketState_TextValue("Socket is created");
	}

	if (bind(this_s, (struct sockaddr*)&addr, sizeof(addr)) != SOCKET_ERROR) {
		form->SetSocketBind_TextValue("Socket successfully binded");
	}else {
		form->SetSocketBind_TextValue("Socket bind failed: " + WSAGetLastError());
        closesocket(this_s);
        return;
    }

	if (listen(this_s, SOMAXCONN) != SOCKET_ERROR) 	{
		// Установите текст в textBox_ListPort
		unsigned short port = ntohs(addr.sin_port);
		System::String^ portString = port.ToString();
		form->SetTextValue("Start listenin at port " + portString);
		GlobalLogger::LogMessage(ConvertToStdString("Start listenin at port " + portString));
	}
	
	if (form->InvokeRequired) {
		form->Invoke(gcnew Action(form, &MyForm::Refresh));
	}
	else {
		form->Refresh();
	}
	handle();
}

void SServer::closeServer() {
	MyForm^ form = safe_cast<MyForm^>(Application::OpenForms["MyForm"]);
	closesocket(this_s);
	WSACleanup();
	if (form != nullptr && !form->IsDisposed) {
		form->SetWSA_TextValue("Server was stopped");
		form->SetSocketState_TextValue(" ");
		form->SetSocketBind_TextValue(" ");
		form->SetTextValue(" ");
	}
}

// Преобразование IP-адреса в строку
String^ GetIPString(const SOCKADDR_IN& addr_c) {
	String^ ipString = String::Format("{0}.{1}.{2}.{3}",
		addr_c.sin_addr.S_un.S_un_b.s_b1,
		addr_c.sin_addr.S_un.S_un_b.s_b2,
		addr_c.sin_addr.S_un.S_un_b.s_b3,
		addr_c.sin_addr.S_un.S_un_b.s_b4);

	// Добавляем порт
	int port = ntohs(addr_c.sin_port);
	String^ fullAddress = String::Format("{0}:{1}", ipString, port);

	return fullAddress;
}

void SServer::handle() {
	while (true)
	{
		SOCKET acceptS;
		SOCKADDR_IN addr_c{};
		// откроем форму MyForm для вывода сообщений о клиенте и ошибок
		MyForm^ form = safe_cast<MyForm^>(Application::OpenForms["MyForm"]);

		// Проверяем, существует ли форма
		if (form == nullptr || form->IsDisposed) {
			break; // Выходим из цикла, если форма закрыта
		}

		int addrlen = sizeof(addr_c);
		if ((acceptS = accept(this_s, (struct sockaddr*)&addr_c, &addrlen)) != INVALID_SOCKET) {
			int ClientPort = ntohs(addr_c.sin_port);
			String^ address = GetIPString(addr_c);
			if (form != nullptr && !form->IsDisposed) {
				form->SetClientAddr_TextValue(address);
			}

            // Создание нового потока для обработки клиента
			DWORD threadId;
			HANDLE hThread;

			hThread = CreateThread(
				NULL,                   // Дескриптор безопасности
				0,                      // Начальный размер стека
				ClientHandler,          // Функция потока
				(LPVOID)acceptS,		// Параметр функции потока
				0,                      // Флаги создания
				&threadId);             // Идентификатор потока

            if (hThread == NULL) {
				if (form != nullptr && !form->IsDisposed) {
					form->SetMessage_TextValue("Error: Failed to create thread");
				}
				GlobalLogger::LogMessage("Error: Failed to create thread");
				closesocket(acceptS);
			} else {
				if (form != nullptr && !form->IsDisposed) {
					form->SetMessage_TextValue("Information: New thread was created");
					// Визуальный разделитель для нового соединения
					GlobalLogger::LogMessage("");
					GlobalLogger::LogMessage("================================================================================");
					GlobalLogger::LogMessage(ConvertToStdString("Information: New thread was created, port " + ClientPort));
				}
				CloseHandle(hThread); // Закрытие дескриптора потока
            }
		}
		Sleep(200);
	}
}

DWORD WINAPI SServer::ClientHandler(LPVOID lpParam) {
	SOCKADDR_IN clientAddr = {};
	int addrLen = sizeof(clientAddr);
	int clientPort = 0;
	int SclientPort = 900000; // Порт Sclient

	SOCKET clientSocket = (SOCKET)lpParam;
	int bytesReceived;
	// откроем форму MyForm для записи сообщений об ошибках
	MyForm^ form = safe_cast<MyForm^>(Application::OpenForms["MyForm"]);

	// Получаем информацию о клиенте
	if (getpeername(clientSocket, (SOCKADDR*)&clientAddr, &addrLen) != SOCKET_ERROR) {
		// Преобразуем порт из сетевого в локальный формат
		clientPort = ntohs(clientAddr.sin_port);
	}

	// Получаем IP-адрес клиента
	String^ clientIPAddress = String::Format("{0}.{1}.{2}.{3}",
		clientAddr.sin_addr.S_un.S_un_b.s_b1,
		clientAddr.sin_addr.S_un.S_un_b.s_b2,
		clientAddr.sin_addr.S_un.S_un_b.s_b3,
		clientAddr.sin_addr.S_un.S_un_b.s_b4);

	// GUID формы (используется и для нового подключения, и для реконнекта)
	std::wstring guid;

	// Проверяем, есть ли уже активное соединение от этого IP
	std::wstring existingGuid = ProjectServerW::DataForm::FindFormByClientIP(clientIPAddress);
	if (!existingGuid.empty()) {
		// Reconnect scenario: device can be power-cycled. Reuse the existing DataForm and continue collecting into the same table.
		guid = existingGuid;
		String^ msg = "Information: Повторное подключение клиента " + clientIPAddress + ", переиспользуем существующую форму";
		if (form != nullptr && !form->IsDisposed) {
			form->SetMessage_TextValue(msg);
		}
		GlobalLogger::LogMessage(ConvertToStdString(msg));
	}
	else {
		// Открываем форму DataForm для демонстрации принятых данных в отдельном потоке
		// Идентификатор формы возвращаем обратно через очередь сообщений
		std::queue<std::wstring> messageQueue;
		std::mutex mtx;
		std::condition_variable cv;
			
		try {	// открываем форму DataForm в новом потоке, передаём в поток ссылки на очередь сообщений, мьютекс и условную переменную
			std::thread formThread([&messageQueue, &mtx, &cv]() {
				ProjectServerW::DataForm::CreateAndShowDataFormInThread(messageQueue, mtx, cv);
				});
			// Получение идентификатора формы данных - guid
			// область видимости нужна для мьютекса, при выходе из области видимости мьютекс разблокируется
			{
				std::unique_lock<std::mutex> lock(mtx);
				cv.wait(lock, [&messageQueue] { return !messageQueue.empty(); });
				guid = messageQueue.front();
				messageQueue.pop();
			}
			// Сохраняем поток формы данных в хранилище с GUID формы в качестве ключа
			ThreadStorage::StoreThread(guid, formThread);
			GlobalLogger::LogMessage(ConvertToStdString("Information: The DataForm has been opened successfully!"));
		}
		catch (const std::exception& e) {	// Обработка исключения, выводим сообщение об ошибке
			String^ errorMessage = gcnew String(e.what());
			if (form != nullptr && !form->IsDisposed) {
				form->SetMessage_TextValue("Error: Couldn't create a form in a new thread " + errorMessage);
				GlobalLogger::LogMessage(ConvertToStdString("Error: Couldn't create a DataForm in a new thread " + errorMessage));
			}
			return 1;
		}
	}

	// Cache managed GUID once; allocating it per packet would add avoidable pressure in the recv() loop.
	String^ guidManaged = gcnew String(guid.c_str());

	int timeout = 30*60*1000;	// Тайм-аут в миллисекундах (1000 мс = 1 секунда), если сообщения нет, то соединение разрывается

	// Установка тайм-аута для операций чтения (recv)
	setsockopt(clientSocket, SOL_SOCKET, SO_RCVTIMEO, (const char*)&timeout, sizeof(timeout));

	// Буфер для накопления данных (для обработки нескольких пакетов)
	char accumulatedBuffer[1024];  // Буфер для накопления данных
	int accumulatedBytes = 0;      // Количество накопленных байт
	const int TELEMETRY_PACKET_SIZE = 48;  // Размер одного пакета телеметрии
	// Critical: some devices can switch to framed packets (AA 55 ... 55 AA).
	// Once framing is detected we permanently parse by markers for this connection.
	bool useMarkedFrames = false;

	// Бесконечный цикл считывания данных
	while (true) {
		// Принимаем данные в конец накопительного буфера
		int spaceAvailable = sizeof(accumulatedBuffer) - accumulatedBytes;
		if (spaceAvailable <= 0) {
			// Буфер переполнен - это ошибка
			GlobalLogger::LogMessage("Error: Accumulated buffer overflow! Resetting buffer.");
			accumulatedBytes = 0;
			continue;
		}
		
		bytesReceived = recv(clientSocket, accumulatedBuffer + accumulatedBytes, spaceAvailable, 0);

		/* ОБРАБОТКА ОШИБОК
		Если функция recv возвращает 0, это означает, что соединение было закрыто клиентом.
		Если функция recv возвращает SOCKET_ERROR, проверяется код ошибки с помощью функции WSAGetLastError.
		Если код ошибки равен WSAETIMEDOUT, это означает, что операция чтения завершилась по тайм-ауту.
		*/
		if (bytesReceived == SOCKET_ERROR) {
			int error = WSAGetLastError();
			if (error == WSAETIMEDOUT) {
				if (form != nullptr && !form->IsDisposed) {
					form->SetMessage_TextValue("Attention: Recv timed out");
					GlobalLogger::LogMessage("Attention: Recv timed out");
				}
			}
			else {
				if (form != nullptr && !form->IsDisposed) {
					form->SetMessage_TextValue("Attention: Recv failed: " + error);
					GlobalLogger::LogMessage(ConvertToStdString("Attention: Recv failed: " + error));
				}
			}
			break;	// Выходим из цикла
		}
		if (bytesReceived == 0) {
			if (form != nullptr && !form->IsDisposed) {
				form->SetMessage_TextValue("Attention: Connection closed by client");
				GlobalLogger::LogMessage("Attention: Connection closed by client");
			}
			break;	// Выходим из цикла
		}

		// Добавляем полученные байты к накопленным
		accumulatedBytes += bytesReceived;
		
		// Обрабатываем все полные пакеты в буфере
		int processedBytes = 0;
		while (accumulatedBytes - processedBytes > 0) {
			// Указатель на начало текущего пакета
			char* buffer = accumulatedBuffer + processedBytes;
			int remainingInBuffer = accumulatedBytes - processedBytes;
			if (remainingInBuffer <= 0) {
				break;
			}

			// ============================
			// Framed mode: AA 55 <payload> 55 AA
			// ============================
			if (useMarkedFrames) {
				const uint8_t SYNC_START_0 = 0xAA;
				const uint8_t SYNC_START_1 = 0x55;
				const uint8_t SYNC_END_0 = 0x55;
				const uint8_t SYNC_END_1 = 0xAA;

				// Need at least 2 bytes to detect the header marker.
				if (remainingInBuffer < 2) {
					break;
				}

				// Resync to the header marker (AA 55). Keep one byte to not drop a split marker.
				if ((uint8_t)buffer[0] != SYNC_START_0 || (uint8_t)buffer[1] != SYNC_START_1) {
					int headerOffset = -1;
					for (int i = 0; i + 1 < remainingInBuffer; i++) {
						if ((uint8_t)buffer[i] == SYNC_START_0 && (uint8_t)buffer[i + 1] == SYNC_START_1) {
							headerOffset = i;
							break;
						}
					}

					if (headerOffset < 0) {
						processedBytes += (remainingInBuffer - 1);
						continue;
					}

					if (headerOffset > 0) {
						processedBytes += headerOffset;
						continue;
					}
				}

				// We are aligned to AA 55.
				// The separator between frames is expected to be SYNC_END + SYNC_START (55 AA AA 55),
				// but either marker can be damaged. We therefore accept the earliest boundary found:
				// - a valid SYNC_END (55 AA), OR
				// - the next SYNC_START (AA 55).
				int footerPos = -1;
				int nextHeaderPos = -1;
				for (int i = 2; i + 1 < remainingInBuffer; i++) {
					const uint8_t b0 = (uint8_t)buffer[i];
					const uint8_t b1 = (uint8_t)buffer[i + 1];

					if (footerPos < 0 && b0 == SYNC_END_0 && b1 == SYNC_END_1) {
						footerPos = i;
					}
					if (nextHeaderPos < 0 && b0 == SYNC_START_0 && b1 == SYNC_START_1) {
						nextHeaderPos = i;
					}

					if (footerPos >= 0 && nextHeaderPos >= 0) {
						break;
					}
				}

				if (footerPos < 0 && nextHeaderPos < 0) {
					break; // Wait for more recv() data.
				}

				// Choose the earliest boundary in the stream:
				// - SYNC_END ends the current frame
				// - SYNC_START begins the next frame and can serve as an end boundary if SYNC_END is damaged
				const bool endByNextHeader = (nextHeaderPos >= 0 && (footerPos < 0 || nextHeaderPos < footerPos));
				const int endPos = endByNextHeader ? nextHeaderPos : footerPos;
				const int payloadLen = endPos - 2;
				if (payloadLen <= 0) {
					// Empty/degenerate frame: advance to the chosen boundary.
					processedBytes += endByNextHeader ? nextHeaderPos : (footerPos + 2);
					continue;
				}

				uint8_t* payload = reinterpret_cast<uint8_t*>(buffer + 2);
				const uint8_t framedType = payload[0];

				// Route by explicit type when possible, otherwise use length as an additional signal.
				// Critical: telemetry routing triggers ACK behavior in the worker; prefer routing to telemetry only when likely.
				const bool looksLikeTelemetryByLen = (payloadLen == TELEMETRY_PACKET_SIZE);
				const bool looksLikeCmdResponseByLen = (payloadLen >= 6 && payloadLen <= static_cast<int>(MAX_COMMAND_SIZE));

				const bool isCmdResponseType = (framedType >= 0x01 && framedType <= 0x04);
				const bool isTelemetryType = (framedType == 0x00);

				if (isCmdResponseType || (!isTelemetryType && looksLikeCmdResponseByLen && !looksLikeTelemetryByLen)) {
					cli::array<System::Byte>^ responseBuffer = gcnew cli::array<System::Byte>(payloadLen);
					Marshal::Copy(IntPtr(payload), responseBuffer, 0, payloadLen);
					DataForm::EnqueueResponse(responseBuffer);
				}
				else {
					cli::array<System::Byte>^ dataBuffer = gcnew cli::array<System::Byte>(payloadLen);
					Marshal::Copy(IntPtr(payload), dataBuffer, 0, payloadLen);
					PacketQueueProcessor::EnqueueTelemetry(dataBuffer, payloadLen, clientPort, guidManaged, clientSocket, clientIPAddress);
				}

				// Consume up to the chosen boundary.
				// If we end by nextHeaderPos, it becomes the new current position and the next loop sees SYNC_START immediately.
				processedBytes += endByNextHeader ? nextHeaderPos : (footerPos + 2);
				continue;
			}

			// ============================
			// Unframed mode (legacy): first byte is packet type
			// ============================

			// Detect a switch to framed packets as early as possible.
			if (remainingInBuffer >= 2 && (uint8_t)buffer[0] == 0xAA && (uint8_t)buffer[1] == 0x55) {
				useMarkedFrames = true;
				continue;
			}

			// If we have only the first marker byte, wait for the next recv() to avoid skipping a valid header.
			if (remainingInBuffer == 1 && (uint8_t)buffer[0] == 0xAA) {
				break;
			}

			// ===== РАЗЛИЧЕНИЕ ТИПА ПАКЕТА ПО ПЕРВОМУ БАЙТУ =====
			uint8_t packetType = buffer[0];
			int bytesInPacket = 0;
			
			// Определяем размер пакета
			if (packetType == 0x00) {
				// ОБРАБОТКА АНОМАЛИИ: Контроллер иногда присылает "ответы" с Type=0x00
				// Проверяем, не 6-байтовый ли это мусорный пакет перед телеметрией
				if (remainingInBuffer >= 6 && remainingInBuffer >= 54) {
					// Достаточно данных для проверки
					// Проверяем: есть ли второй Type=0x00 на позиции 6?
					if (buffer[6] == 0x00) {
						// Вероятно, первые 6 байт - мусор, пропускаем их
						processedBytes += 6;
						continue;  // Начинаем обработку заново со следующего пакета
					}
				}
				
				// Телеметрия - фиксированный размер 48 байт
				bytesInPacket = TELEMETRY_PACKET_SIZE;
			} else if (packetType >= 0x01 && packetType <= 0x04) {
				// Ответ на команду - переменный размер
				// Минимум 4 байта для чтения DataLen: Type + Code + Status + DataLen
				if (remainingInBuffer < 4) {
					break; // Недостаточно данных для определения размера
				}
				uint8_t dataLen = buffer[3]; // DataLen находится в 4-м байте (индекс 3)
				bytesInPacket = 4 + dataLen + 2; // Type+Code+Status+DataLen + Data + CRC
			} else {
				// Unknown type: drop a single byte to resync quickly, but do not drop a potential AA 55 header start.
				if (packetType == 0xAA) {
					if (remainingInBuffer < 2) {
						break;
					}
					if ((uint8_t)buffer[1] == 0x55) {
						useMarkedFrames = true;
						continue;
					}
				}
				processedBytes += 1;
				continue;
			}
			
			// Проверяем, достаточно ли данных для полного пакета
			if (remainingInBuffer < bytesInPacket) {
				break; // Пакет неполный, ждем следующего recv()
			}
			
			// ===== ОБРАБОТКА ПАКЕТА В ЗАВИСИМОСТИ ОТ ТИПА =====
			if (packetType >= 0x01 && packetType <= 0x04) {
				// ===== ЭТО ОТВЕТ НА КОМАНДУ =====
				// Пакеты с Type = 0x01-0x04 это ответы от контроллера на команды сервера
				
				// Создаем управляемый массив для ответа
				cli::array<System::Byte>^ responseBuffer = gcnew cli::array<System::Byte>(bytesInPacket);
				Marshal::Copy(IntPtr(buffer), responseBuffer, 0, bytesInPacket);
				
				// Добавляем в очередь ответов
				DataForm::EnqueueResponse(responseBuffer);
				
				// Переходим к следующему пакету
				processedBytes += bytesInPacket;
			} // конец if (packetType >= 0x01 && packetType <= 0x04)
			else if (packetType == 0x00) {
				// Telemetry processing (CRC, UI dispatch, ACK) is offloaded to a background worker.
				cli::array<System::Byte>^ dataBuffer = gcnew cli::array<System::Byte>(bytesInPacket);
				Marshal::Copy(IntPtr(buffer), dataBuffer, 0, bytesInPacket);
				PacketQueueProcessor::EnqueueTelemetry(dataBuffer, bytesInPacket, clientPort, guidManaged, clientSocket, clientIPAddress);

				// Переходим к следующему пакету
				processedBytes += bytesInPacket;
			} // else if (packetType == 0x00)
		} // конец while (обработка пакетов из буфера)
		
		// Переносим необработанные байты в начало буфера
		int remainingBytes = accumulatedBytes - processedBytes;
		if (remainingBytes > 0) {
			memmove(accumulatedBuffer, accumulatedBuffer + processedBytes, remainingBytes);
		}
		accumulatedBytes = remainingBytes;
	}	// конец while (основной цикл)

exitMainLoop:  // Метка для выхода из всех циклов
	// Соединение оборвалось. Форму НЕ закрываем: устройство может вернуться в течение 30 минут,
	// тогда новое подключение переиспользует текущую форму и продолжит заполнение таблицы.
	closesocket(clientSocket);
	// Помечаем сокет формы как недействительный (команды на отправку должны перестать проходить)
	try {
		DataForm^ form2 = DataForm::GetFormByGuid(guid);
		if (form2 != nullptr && !form2->IsDisposed && !form2->Disposing) {
			form2->ClientSocket = INVALID_SOCKET;
		}
	}
	catch (...) {}

	return 0;
}

std::string ConvertToStdString(System::String^ managedString) {

	if (managedString == nullptr)
		return std::string();

	// Используем UTF8 кодировку для конвертации
//	cli::array<System::Byte>^ bytes = System::Text::Encoding::UTF8->GetBytes(managedString);
	// Используем 1251 кодировку для конвертации
	cli::array<System::Byte>^ bytes = System::Text::Encoding::GetEncoding(1251)->GetBytes(managedString);

	std::string result(bytes->Length, 0);

	{
		cli::pin_ptr<System::Byte> pinnedBytes = &bytes[0];
		memcpy(&result[0], pinnedBytes, bytes->Length);
	}
	return result;
}

// Unicode::snprintfFloat
namespace Unicode {
    int snprintfFloat(char* buffer, size_t size, const char* format, float value) {
        return snprintf(buffer, size, format, value);
    }
}