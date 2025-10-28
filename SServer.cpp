#include "SServer.h" // Включает DataForm.h внутри себя
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
					GlobalLogger::LogMessage(ConvertToStdString("Information: New thread was created, port " + ClientPort));
				}
				CloseHandle(hThread); // Закрытие дескриптора потока
            }
		}
		Sleep(200);
	}
}

// Табличное определение CRC
static uint16_t MB_GetCRC(char* buf, uint16_t len)
{
	uint16_t crc_16 = 0xffff;
	for (uint16_t i = 0; i < len; i++)
	{
		crc_16 = (crc_16 >> 8) ^ crc16_table[(buf[i] ^ crc_16) & 0xff];
	}
	return crc_16;
}

DWORD WINAPI SServer::ClientHandler(LPVOID lpParam) {
	SOCKADDR_IN clientAddr = {};
	int addrLen = sizeof(clientAddr);
	int clientPort = 0;
	int SclientPort = 900000; // Порт Sclient

	SOCKET clientSocket = (SOCKET)lpParam;
    char buffer[512];
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

	// Проверяем, есть ли уже активное соединение от этого IP
	std::wstring existingGuid = ProjectServerW::DataForm::FindFormByClientIP(clientIPAddress);
	if (!existingGuid.empty()) {
		// Уже есть активное соединение от этого IP
		String^ warningMsg = "Внимание: Уже существует активное соединение от клиента " + clientIPAddress +
			". Новое соединение не будет создано.";
		if (form != nullptr && !form->IsDisposed) {
			form->SetMessage_TextValue(warningMsg);
			GlobalLogger::LogMessage(ConvertToStdString(warningMsg));
		}
		// Закрываем новое соединение
		closesocket(clientSocket);
		return 1;
	}

	// Открываем форму DataForm для демонстрации принятых данных в отдельном потоке
	// Идентификатор формы возвращаем обратно через очередь сообщений
	std::wstring guid;
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

	int timeout = 30*60*1000;	// Тайм-аут в миллисекундах (1000 мс = 1 секунда), если сообщения нет, то соединение разрывается

	// Установка тайм-аута для операций чтения (recv)
	setsockopt(clientSocket, SOL_SOCKET, SO_RCVTIMEO, (const char*)&timeout, sizeof(timeout));

	// Бесконечный цикл считывания данных
	while (true) {
		bytesReceived = recv(clientSocket, buffer, sizeof(buffer), 0);

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
		};

		// ===== РАЗЛИЧЕНИЕ ТИПА ПАКЕТА ПО ПЕРВОМУ БАЙТУ =====
		uint8_t packetType = buffer[0];
		
		if (packetType >= 0x01 && packetType <= 0x04) {
			// ===== ЭТО ОТВЕТ НА КОМАНДУ =====
			// Пакеты с Type = 0x01-0x04 это ответы на команды от контроллера
			
			// Создаем управляемый массив для ответа
			cli::array<System::Byte>^ responseBuffer = gcnew cli::array<System::Byte>(bytesReceived);
			Marshal::Copy(IntPtr(buffer), responseBuffer, 0, bytesReceived);
			
		// Добавляем в очередь ответов
		DataForm::EnqueueResponse(responseBuffer);
		
		GlobalLogger::LogMessage(ConvertToStdString(String::Format(
			"Response received and enqueued: Type=0x{0:X2}, Size={1} bytes", 
			packetType, bytesReceived)));
		}
		else if (packetType == 0x00) {
			// ===== ЭТО ТЕЛЕМЕТРИЯ =====
			// Пакеты с Type = 0x00 это телеметрия от контроллера
			
			// Длина посылки вместе с CRC
			uint8_t LengthOfPackage = 48;
			// Вычислим CRC16 для первых 46 байт
			uint16_t dataCRC = MB_GetCRC(buffer, LengthOfPackage - 2);
			uint16_t DatCRC;
			memcpy(&DatCRC, &buffer[LengthOfPackage - 2], 2);

			if (DatCRC == dataCRC) {
				// Отправляем зеркальную посылку клиенту
				// send(clientSocket, buffer, bytesReceived, 0);

				// Найдём форму по идентификатору
				DataForm^ form2 = DataForm::GetFormByGuid(guid);
				if (form2 != nullptr && !form2->IsDisposed && form2->IsHandleCreated && !form2->Disposing) 
				{
					
					// Передадим в форму сокет клиента этой формы и IP-адрес
					if (form2 != nullptr) {
						form2->ClientSocket = clientSocket;
						form2->ClientIP = clientIPAddress;
					}
					
					if (clientPort < SclientPort) {
						// Создаем копию данных для безопасной передачи в другой поток
						cli::array<System::Byte>^ dataBuffer = gcnew cli::array<System::Byte>(bytesReceived);
						Marshal::Copy(IntPtr(buffer), dataBuffer, 0, bytesReceived);

						// Вызываем AddDataToTable через Invoke для выполнения в потоке формы
						form2->BeginInvoke(gcnew Action<cli::array <System::Byte>^, int, int>
							(form2, &DataForm::AddDataToTableThreadSafe),
							dataBuffer, bytesReceived, clientPort);
						// Refresh вызывать отдельно уже не нужно - он будет вызван в AddDataToTableThreadSafe
					} else {
						// здесь обработка для другого типа порта
					}
				}
				else {
					// Форма отсутствует, завершаем поток
					break;
		} // конец if (form2 != nullptr...
	} // if (DatCRC == dataCRC)
	else {
		// Ошибка CRC телеметрии
		GlobalLogger::LogMessage(ConvertToStdString(String::Format(
			"Telemetry CRC error: Expected={0:X4}, Received={1:X4}", 
			dataCRC, DatCRC)));
	}
} // else if (packetType == 0x00)
else {
	// Неизвестный тип пакета
	GlobalLogger::LogMessage(ConvertToStdString(String::Format(
		"Unknown packet type: 0x{0:X2}, Size={1} bytes", 
		packetType, bytesReceived)));
		}
	}	// конец while

	// Цикл приёма данных прерван по ошибке приёма, закрываем форрму приёма данных
	closesocket(clientSocket);
	// Найдём форму данных по идентификатору и закроем её
	DataForm::CloseForm(guid);
	// Закрытие потока формы данных по guid формы
	ThreadStorage::StopThread(guid);

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