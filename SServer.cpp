#include "SServer.h" // 
#include <map>						// Для использования std::map - структуры, сохраняющей соответствие ID и ссылки на потоки

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
}

void SServer::startServer() {
	MyForm^ form = safe_cast<MyForm^>(Application::OpenForms["MyForm"]);

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
	form->SetWSA_TextValue("Server was stopped");
	form->SetSocketState_TextValue(" ");
	form->SetSocketBind_TextValue(" ");
	form->SetTextValue(" ");
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

		int addrlen = sizeof(addr_c);
		if ((acceptS = accept(this_s, (struct sockaddr*)&addr_c, &addrlen)) != INVALID_SOCKET) {
			int ClientPort = ntohs(addr_c.sin_port);
			String^ address = GetIPString(addr_c);
			form->SetClientAddr_TextValue(address);

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
				form->SetMessage_TextValue("Error: Failed to create thread");
				closesocket(acceptS);
            } else {
				form->SetMessage_TextValue("Information: New thread was created");
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

// Эта функция принимает буфер и его длину, затем преобразует каждый байт буфера в шестнадцатеричный формат и добавляет его в строку.
String^ bufferToHex(const char* buffer, int length) {
	// Использование std::stringstream позволяет преобразовывать данные в строку и форматировать их.
    std::stringstream ss;
    ss << std::hex << std::setfill('0');
    for (int i = 0; i < length; ++i) {
        ss << std::setw(2) << static_cast<unsigned int>(static_cast<unsigned char>(buffer[i])) << " ";
    }
    return gcnew String(ss.str().c_str());
}

DWORD WINAPI SServer::ClientHandler(LPVOID lpParam) {
	SOCKADDR_IN clientAddr;
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
	// Открываем форму DataForm для демонстрации принятых данных в отдельном потоке
	// Идентификатор формы возвращаем обратно через очередь сообщений
	std::queue<std::wstring> messageQueue;
	std::mutex mtx;
	std::condition_variable cv;
		
	// Идентификатор потока формы
	std::wstring id;

	try {	// открываем форму DataForm в новом потоке, передаём в поток ссылки на очередь сообщений, мьютекс и условную переменную
		std::thread formThread([&messageQueue, &mtx, &cv]() {
			ProjectServerW::DataForm::CreateAndShowDataFormInThread(messageQueue, mtx, cv);
			});
		formThread.detach(); // Отсоединяем поток формы от основного потока, чтобы он работал независимо

		// Получение Windows thread ID
		DWORD threadId = GetCurrentThreadId();
		// Преобразование thread ID в строку
		id = std::to_wstring(threadId);
		// Сохраним поток с идентификатором в map
		ThreadStorage::StoreThread(id, formThread);
	}
	catch (const std::exception& e) {	// Обработка исключения, выводим сообщение об ошибке
		String^ errorMessage = gcnew String(e.what());
		form->SetMessage_TextValue("Error: Couldn't create a form in a new thread "+ errorMessage);
		return 1;
	}

	// Ожидание идентификатора из очереди сообщений потока формы
	std::wstring guid;		// это переменная для идентификатора потока
	// область видимости нужна для мьютекса, при выходе из области видимости мьютекс разблокируется
	{
		std::unique_lock<std::mutex> lock(mtx);
		cv.wait(lock, [&messageQueue] { return !messageQueue.empty(); });
		guid = messageQueue.front();
		messageQueue.pop();
	}
		
	int timeout = 600*1000;	// Тайм-аут в миллисекундах (1000 мс = 1 секунда), если сообщения нет, то соединение разрывается

	// Установка тайм-аута для операций чтения (recv)
	setsockopt(clientSocket, SOL_SOCKET, SO_RCVTIMEO, (const char*)&timeout, sizeof(timeout));
	String^ managedString{};
	String^ dataCRC_String{};

	// Бесконечный цикл считывания данных
	while ((bytesReceived = recv(clientSocket, buffer, sizeof(buffer), 0)) > 0) {
		uint16_t dataCRC = MB_GetCRC(buffer, 40);
		dataCRC_String = bufferToHex((const char*) &dataCRC, 2);
		uint16_t DatCRC;
		memcpy(&DatCRC, &buffer[40], 2);

		if (DatCRC == dataCRC) {
			send(clientSocket, buffer, bytesReceived, 0);

			// Найдём форму по идентификатору
			DataForm^ form2 = DataForm::GetFormByGuid(guid);
			if (form2 != nullptr) {
				if (clientPort < SclientPort) {
					// Создаем копию данных для безопасной передачи в другой поток
					cli::array<System::Byte>^ dataBuffer = gcnew cli::array<System::Byte>(bytesReceived);
					Marshal::Copy(IntPtr(buffer), dataBuffer, 0, bytesReceived);

					// Вызываем AddDataToTable через Invoke для выполнения в потоке формы
					form2->Invoke(gcnew Action<cli::array<System::Byte>^, int>(
						form2, &DataForm::AddDataToTableThreadSafe),
						dataBuffer, bytesReceived);

					// Refresh вызывать отдельно уже не нужно - он будет вызван в AddDataToTableThreadSafe
				}
			}
		}
	}	// конец while
	/*
	Если функция recv возвращает 0, это означает, что соединение было закрыто клиентом.
	Если функция recv возвращает SOCKET_ERROR, проверяется код ошибки с помощью функции WSAGetLastError.
	Если код ошибки равен WSAETIMEDOUT, это означает, что операция чтения завершилась по тайм-ауту.
	*/
	if (bytesReceived == 0) {
		form->SetMessage_TextValue("Attention: Connection closed by client");
	}
	else if (bytesReceived == SOCKET_ERROR) {
		int error = WSAGetLastError();
		if (error == WSAETIMEDOUT) {
			form->SetMessage_TextValue("Attention: Recv timed out");
		}
		else {
			form->SetMessage_TextValue("Attention: Recv failed: " + error);
		}
	}

	closesocket(clientSocket);
	// Найдём форму по идентификатору и закроем её
	DataForm::CloseForm(guid);
	// Закрытие потока
	ThreadStorage::StopThread(id);
	return 0;
}
