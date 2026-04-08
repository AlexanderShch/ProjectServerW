#include "SServer.h" // Включает DataForm.h внутри себя
#include "PacketQueueProcessor.h"
#include "Commands.h" // MAX_COMMAND_SIZE is used for framed payload length heuristics.
#include <fstream>
#include <iostream>
#include <chrono>
#include <ctime>
#include <msclr/marshal_cppstd.h>
#include <cstdio>
#include <set>
#include <mutex>
#include <atomic>

// Добавьте эту строку для доступа к Marshal
using namespace System::Runtime::InteropServices;

using namespace System::Windows::Forms;
using namespace ProjectServerW;

std::map<std::wstring, std::thread> formThreads;  // Хранение потоков форм
namespace {
	// Глобальное состояние останова сервера: используется и в accept-loop, и в client threads.
	std::atomic<bool> g_serverStopping(false);
	std::mutex g_clientsMx;
	std::set<SOCKET> g_activeClients;

	void RegisterClientSocket(SOCKET s) {
		if (s == INVALID_SOCKET) return;
		std::lock_guard<std::mutex> lk(g_clientsMx);
		g_activeClients.insert(s);
	}

	void UnregisterClientSocket(SOCKET s) {
		if (s == INVALID_SOCKET) return;
		std::lock_guard<std::mutex> lk(g_clientsMx);
		g_activeClients.erase(s);
	}

	void ShutdownAllClientSockets() {
		std::lock_guard<std::mutex> lk(g_clientsMx);
		for (SOCKET s : g_activeClients) {
			if (s == INVALID_SOCKET) continue;
			shutdown(s, SD_BOTH);
			closesocket(s);
		}
		g_activeClients.clear();
	}
}

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
	g_serverStopping.store(false);
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
		GlobalLogger::LogMessage("Start listenin at port " + portString);
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
	g_serverStopping.store(true);
	if (this_s != INVALID_SOCKET) {
		shutdown(this_s, SD_BOTH);
		closesocket(this_s);
		this_s = INVALID_SOCKET;
	}
	// Важно: иначе client recv-потоки могут жить до SO_RCVTIMEO и держать процесс.
	ShutdownAllClientSockets();
	WSACleanup();
	if (form != nullptr && !form->IsDisposed) {
		form->SetWSA_TextValue("Server was stopped");
		form->SetSocketState_TextValue(" ");
		form->SetSocketBind_TextValue(" ");
		form->SetTextValue(" ");
	}
}

// Преобразование IP-адреса в строку (внутренняя функция файла)
static String^ GetIPString(const SOCKADDR_IN& addr_c) {
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
	while (!g_serverStopping.load())
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
			RegisterClientSocket(acceptS);
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
				UnregisterClientSocket(acceptS);
				closesocket(acceptS);
			} else {
				if (form != nullptr && !form->IsDisposed) {
					form->SetMessage_TextValue("Information: New thread was created");
					// Визуальный разделитель для нового соединения
					GlobalLogger::LogMessage("");
					GlobalLogger::LogMessage("================================================================================");
					GlobalLogger::LogMessage("Information: New thread was created, port " + ClientPort);
				}
				CloseHandle(hThread); // Закрытие дескриптора потока
            }
		}
		else {
			// accept прерван закрытием listen-сокета во время выхода.
			if (g_serverStopping.load() || this_s == INVALID_SOCKET) {
				break;
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
	const bool isReconnect = !existingGuid.empty();
	if (!existingGuid.empty()) {
		// Reconnect scenario: device can be power-cycled. Reuse the existing DataForm and continue collecting into the same table.
		guid = existingGuid;
		String^ msg = "Information: Reconnect " + clientIPAddress;
		if (form != nullptr && !form->IsDisposed) {
			form->SetMessage_TextValue(msg);
		}
		GlobalLogger::LogMessage(msg);
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
			GlobalLogger::LogMessage("Information: The DataForm has been opened successfully!");
		}
		catch (const std::exception& e) {	// Обработка исключения, выводим сообщение об ошибке
			String^ errorMessage = gcnew String(e.what());
			if (form != nullptr && !form->IsDisposed) {
				form->SetMessage_TextValue("Error: Couldn't create a form in a new thread " + errorMessage);
				GlobalLogger::LogMessage("Error: Couldn't create a DataForm in a new thread " + errorMessage);
			}
			return 1;
		}
	}

	// Критично: кэшируем managed GUID один раз; создание строки на каждый пакет добавляет лишнюю нагрузку в recv()-цикле.
	String^ guidManaged = gcnew String(guid.c_str());

	// Критично: FindFormByClientIP() матчится по DataForm::ClientIP. Если соединение оборвётся до первой телеметрии,
	// ClientIP останется null (обычно он обновляется из telemetry worker), и при переподключении можно получить дубли пустых форм.
	// Поэтому выставляем ClientIP/ClientSocket сразу, чтобы переподключения переиспользовали тот же DataForm даже без телеметрии.
	DataForm^ df = nullptr;
	try {
		df = DataForm::GetFormByGuid(guid);
		if (df != nullptr && !df->IsDisposed && !df->Disposing) {
			const SOCKET prevSocket = df->ClientSocket;
			if (prevSocket != INVALID_SOCKET && prevSocket != clientSocket) {
				// Причина: новое подключение с тем же IP должно мгновенно "выбивать" старый recv()-поток,
				// иначе он может висеть до SO_RCVTIMEO и держать ресурсы/путать диагностику.
				// Важно: closesocket(prevSocket) здесь не делаем, чтобы избежать двойного closesocket
				// (старый recv()-поток сам закроет свой сокет на выходе).
				shutdown(prevSocket, SD_BOTH);
				GlobalLogger::LogMessage(String::Format(
					"Information: Reconnect: drop old socket {0}",
					clientIPAddress));
			}

			df->ClientIP = clientIPAddress;
			df->ClientSocket = clientSocket;

			if (df->IsHandleCreated) {
				// На любом подключении запускаем стартовые команды через отложенный механизм после установки сокета.
				GlobalLogger::LogMessage(String::Format(
					"Information: Deferred startup scheduled for {0}",
					clientIPAddress));
				df->BeginInvoke(gcnew System::Windows::Forms::MethodInvoker(df, &DataForm::ScheduleDeferredStartupOnReconnect));
			}
		}
	}
	catch (...) {}

	// Причина: SO_RCVTIMEO используется как "тик", чтобы recv()-поток мог выйти, если этот сокет уже не актуален
	// (например, пришло переподключение и форма получила новый ClientSocket). Закрытие формы по истечении 5 минут
	// без телеметрии делает DataForm::inactivityTimer.
	int timeout = 5 * 60 * 1000;	// 5 минут

	// Установка тайм-аута для операций чтения (recv)
	setsockopt(clientSocket, SOL_SOCKET, SO_RCVTIMEO, (const char*)&timeout, sizeof(timeout));

	// Буфер для накопления данных (для обработки нескольких пакетов)
	char accumulatedBuffer[1024] = {};  // Буфер для накопления данных
	int accumulatedBytes = 0;            // Количество накопленных байт
	// Контроллер передаёт на линию фреймы с маркером начала кадра:
	//   [AA 55][Type][Cmd][Status][DataLen][Data...][CRC16]
	// CRC16 считается по блоку: [Type][Cmd][Status][DataLen][Data...].
	const int TELEMETRY_DATA_LEN = 46;
	const int CONTROL_LOG_DATA_LEN = 89; // sizeof(ControlLogPayload_t) for T_filt_C[0..5]
	// Максимальный размер блока после AA55: 4 + DataLen + 2.
	const int MAX_FRAMED_PAYLOAD_LEN = 128;

	// Бесконечный цикл считывания данных
	// Важно: recv() вызываем БЕЗ удержания recvGate, иначе UI/таймер SEND_STATE не смогут отправить команду —
	// контроллер шлёт данные только по запросу SEND_STATE, получается deadlock (пустая шина, окно не реагирует).
	while (!g_serverStopping.load()) {
		int spaceAvailable = sizeof(accumulatedBuffer) - accumulatedBytes;
		if (spaceAvailable <= 0) {
			GlobalLogger::LogMessage("Error: Accumulated buffer overflow! Resetting buffer.");
			accumulatedBytes = 0;
			continue;
		}

		bytesReceived = recv(clientSocket, accumulatedBuffer + accumulatedBytes, spaceAvailable, 0);

		if (bytesReceived == SOCKET_ERROR) {
			int error = WSAGetLastError();
			if (error == WSAETIMEDOUT) {
				try {
					DataForm^ currentForm = DataForm::GetFormByGuid(guid);
					if (currentForm == nullptr || currentForm->IsDisposed || currentForm->Disposing) {
						break;
					}
					if (currentForm->ClientSocket != clientSocket) {
						break;
					}
				}
				catch (...) {
					break;
				}
				continue;
			}
			else {
				if (form != nullptr && !form->IsDisposed) {
					form->SetMessage_TextValue("Attention: Recv failed: " + error);
					GlobalLogger::LogMessage("Attention: Recv failed: " + error);
				}
			}
			break;
		}
		if (bytesReceived == 0) {
			if (form != nullptr && !form->IsDisposed) {
				form->SetMessage_TextValue("Attention: Connection closed by client");
				GlobalLogger::LogMessage("Attention: Connection closed by client");
			}
			break;
		}

		accumulatedBytes += bytesReceived;

		// Захватываем recvGate только на время разбора буфера, чтобы не пересекаться с отправкой команд.
		System::Object^ recvGate = PacketQueueProcessor::GetReceivingGate(clientSocket);
		System::Threading::Monitor::Enter(recvGate);
		try {
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
			// Режим фреймов: AA 55 + Type + Len + Payload + CRC16
			// ============================
				const uint8_t SYNC_START_0 = 0xAA;
				const uint8_t SYNC_START_1 = 0x55;

				// Нужно минимум 2 байта, чтобы проверить маркер заголовка.
				if (remainingInBuffer < 2) {
					break;
				}

				// Выполним ресинхронизацию потока в режиме “фреймированных” пакетов (где кадр ожидается как AA 55 <payload> 55 AA). 
				// Если первые 2 байта текущей позиции (buffer[0], buffer[1]) не равны AA 55, значит парсер потерял границу кадра и нам нужно найти её заново.
				if ((uint8_t)buffer[0] != SYNC_START_0 || (uint8_t)buffer[1] != SYNC_START_1) {
					int headerOffset = -1;
					// Пробежимся по всему буферу, ищем первую встречу AA 55
					for (int i = 0; i + 1 < remainingInBuffer; i++) {
						if ((uint8_t)buffer[i] == SYNC_START_0 && (uint8_t)buffer[i + 1] == SYNC_START_1) {
							headerOffset = i;
							break;
						}
					}

					if (headerOffset < 0) {
						// Если не нашли, значит парсер потерял границу кадра и нам нужно пропустить все байты до конца буфера.
						processedBytes += (remainingInBuffer - 1);
						continue;
					}

					if (headerOffset > 0) {
						// Если нашли, значит мы нашли границу кадра и нам нужно пропустить все байты до неё.
						processedBytes += headerOffset;
						continue;
					}
				}

				// Мы выровнены на AA 55.
				// [AA 55][Type][Cmd][Status][DataLen][Data...][CRC16]
				if (remainingInBuffer < 6) {
					// Нет даже [AA 55][Type][Cmd][Status][DataLen]
					break;
				}

				const uint8_t* frame = reinterpret_cast<const uint8_t*>(buffer);
				const uint8_t* payload = frame + 2; // payload = [Type][Cmd][Status][DataLen][Data...][CRC]

				const uint8_t framedType = payload[0];
				const uint8_t framedCmd = payload[1];
				const uint8_t framedStatus = payload[2];
				const uint8_t payloadDataLenByte = payload[3];

				const int payloadNoCrcLen = 4 + static_cast<int>(payloadDataLenByte); // [Type..DataLen][Data...]
				const int payloadWithCrcLen = payloadNoCrcLen + 2;
				const int totalFrameLen = 2 + payloadWithCrcLen;

				if (payloadWithCrcLen <= 0 || payloadWithCrcLen > MAX_FRAMED_PAYLOAD_LEN) {
					GlobalLogger::LogMessage(String::Format(
						"Warning: Dropping framed packet with invalid length. Type=0x{0:X2}, Cmd=0x{1:X2}, DataLen={2}",
						framedType, framedCmd, payloadDataLenByte));
					processedBytes += 1;
					continue;
				}

				if (remainingInBuffer < totalFrameLen) {
					break;
				}

				uint16_t receivedCrc = 0;
				memcpy(&receivedCrc, payload + payloadNoCrcLen, 2);
				const uint16_t calculatedCrc = CalculateCommandCRC(payload, static_cast<size_t>(payloadNoCrcLen));

				const bool isTelemetryFrame =
					(framedType == 0x03 && framedCmd == 0x08 && framedStatus == 0x00);
				const bool isLogFrame =
					(framedType == 0x00 && framedCmd == 0x01 && framedStatus == 0x00);

				if (receivedCrc != calculatedCrc) {
					// Для телеметрии CRC-ошибка должна повлиять на DATA_OK/DATA_FALSE, поэтому телеметрию
					// направляем в worker даже при неравенстве CRC.
					if (isTelemetryFrame && payloadDataLenByte == TELEMETRY_DATA_LEN) {
						cli::array<System::Byte>^ dataBuffer = gcnew cli::array<System::Byte>(payloadWithCrcLen);
						Marshal::Copy(IntPtr(const_cast<uint8_t*>(payload)), dataBuffer, 0, payloadWithCrcLen);
						PacketQueueProcessor::EnqueueTelemetry(dataBuffer, payloadWithCrcLen, clientPort, guidManaged, clientSocket, clientIPAddress);
						processedBytes += totalFrameLen;
						continue;
					}

					GlobalLogger::LogMessage(String::Format(
						"Warning: Dropping framed packet due CRC mismatch. Type=0x{0:X2}, Cmd=0x{1:X2}, DataLen={2}",
						framedType, framedCmd, payloadDataLenByte));
					processedBytes += 1;
					continue;
				}

				if (isTelemetryFrame) {
					if (payloadDataLenByte != TELEMETRY_DATA_LEN) {
						GlobalLogger::LogMessage(String::Format(
							"Warning: Dropping framed telemetry with unexpected DataLen={0}", payloadDataLenByte));
					}
					else {
						cli::array<System::Byte>^ dataBuffer = gcnew cli::array<System::Byte>(payloadWithCrcLen);
						Marshal::Copy(IntPtr(const_cast<uint8_t*>(payload)), dataBuffer, 0, payloadWithCrcLen);
						PacketQueueProcessor::EnqueueTelemetry(dataBuffer, payloadWithCrcLen, clientPort, guidManaged, clientSocket, clientIPAddress);
					}
				}
				else if (isLogFrame) {
					if (payloadDataLenByte != CONTROL_LOG_DATA_LEN) {
						GlobalLogger::LogMessage(String::Format(
							"Warning: Dropping framed control log with unexpected DataLen={0}", payloadDataLenByte));
					}
					else {
						cli::array<System::Byte>^ logBuffer = gcnew cli::array<System::Byte>(payloadWithCrcLen);
						Marshal::Copy(IntPtr(const_cast<uint8_t*>(payload)), logBuffer, 0, payloadWithCrcLen);
						PacketQueueProcessor::EnqueueControlLog(logBuffer, payloadWithCrcLen, clientPort, guidManaged, clientSocket, clientIPAddress);
					}
				}
				else {
					cli::array<System::Byte>^ responseBuffer = gcnew cli::array<System::Byte>(payloadWithCrcLen);
					Marshal::Copy(IntPtr(const_cast<uint8_t*>(payload)), responseBuffer, 0, payloadWithCrcLen);
					if (df != nullptr && !df->IsDisposed && !df->Disposing) {
						df->EnqueueResponse(responseBuffer);
					}
				}

				// Поглощаем весь кадр целиком: [AA 55][Type][Cmd][Status][DataLen][Data...][CRC16]
				processedBytes += totalFrameLen;
				continue;
		} // конец while (обработка пакетов из буфера)
		
		// Переносим необработанные байты в начало буфера
		int remainingBytes = accumulatedBytes - processedBytes;
		if (remainingBytes > 0) {
			memmove(accumulatedBuffer, accumulatedBuffer + processedBytes, remainingBytes);
		}
		accumulatedBytes = remainingBytes;

		}
		finally {
			System::Threading::Monitor::Exit(recvGate);
		}
		// Уступаем планировщику: поток отправки (UI/таймер SEND_STATE) может сразу захватить recvGate и отправить команды. 
		// Один мастер — задержка не нужна.
		System::Threading::Thread::Sleep(0);
	}	// конец while (основной цикл)

	// Соединение оборвалось. Форму НЕ закрываем: устройство может вернуться в течение 30 минут,
	// тогда новое подключение переиспользует текущую форму и продолжит заполнение таблицы.
	closesocket(clientSocket);
	UnregisterClientSocket(clientSocket);
	// Помечаем сокет формы как недействительный (команды на отправку должны перестать проходить)
	try {
		DataForm^ form2 = DataForm::GetFormByGuid(guid);
		if (form2 != nullptr && !form2->IsDisposed && !form2->Disposing) {
			// Важно: не трогаем форму, если за время работы этого потока пришло переподключение
			// и форма уже указывает на новый сокет.
			if (form2->ClientSocket == clientSocket) {
				form2->ClientSocket = INVALID_SOCKET;
			}
		}
	}
	catch (...) {}

	return 0;
}

std::string ConvertToStdString(System::String^ managedString) {

	if (managedString == nullptr)
		return std::string();

	// Используем Windows 1251 кодировку для конвертации
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