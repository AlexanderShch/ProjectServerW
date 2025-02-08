#include "SServer.h" // 
#include <map>						// Для использования std::map - структуры, сохраняющей соответствие ID и ссылки на потоки

using namespace System::Windows::Forms;
using namespace ProjectServerW;

std::map<std::wstring, std::thread> formThreads;  // Хранение потоков форм

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

void printCurrentTime() {
    //std::time_t now = std::time(nullptr);
    //std::tm* localTime = std::localtime(&now);
    //std::cout << "Current time: " << std::put_time(localTime, "%Y-%m-%d %H:%M:%S") << std::endl;
}

// Эта функция принимает буфер и его длину, затем преобразует каждый байт буфера в шестнадцатеричный формат и добавляет его в строку.
String^ bufferToHex(const char* buffer, int length) {
	// Использование std::stringstream позволяет легко преобразовывать данные в строку и форматировать их.
    std::stringstream ss;
    ss << std::hex << std::setfill('0');
	// << std::hex: Этот манипулятор устанавливает формат вывода чисел в шестнадцатеричном формате. После этого все последующие числа, 
	// записанные в поток ss, будут выводиться в шестнадцатеричном формате.
	// << std::setfill('0'): Этот манипулятор устанавливает символ заполнения для вывода чисел. В данном случае, символом заполнения является '0'. 
	// Это означает, что если число занимает меньше символов, чем указано в std::setw, то пустые места будут заполнены нулями.
    for (int i = 0; i < length; ++i) {
        ss << std::setw(2) << static_cast<unsigned int>(static_cast<unsigned char>(buffer[i])) << " ";
    }
    return gcnew String(ss.str().c_str());
}
// std::setw(2): Устанавливает ширину поля вывода в 2 символа.
// static_cast<unsigned int>(static_cast<unsigned char>(buffer[i])): Преобразует байт в беззнаковое 
// целое число, чтобы корректно отобразить его в шестнадцатеричном формате.
// " ": Добавляет пробел между шестнадцатеричными значениями для удобства чтения.


    DWORD WINAPI SServer::ClientHandler(LPVOID lpParam) {
		SOCKADDR_IN clientAddr;
		int addrLen = sizeof(clientAddr);
		int clientPort = 0;
		int SclientPort = 9000; // Порт Sclient

		SOCKET clientSocket = (SOCKET)lpParam;
        char buffer[512];
        int bytesReceived;
		MyForm^ form = safe_cast<MyForm^>(Application::OpenForms["MyForm"]);

		// Получаем информацию о клиенте
		if (getpeername(clientSocket, (SOCKADDR*)&clientAddr, &addrLen) != SOCKET_ERROR) {
			// Преобразуем порт из сетевого в локальный формат
			clientPort = ntohs(clientAddr.sin_port);
		}
		// Открываем форму для демонстрации принятых данных в отдельном потоке
		// Идентификатор формы передаём через очередь сообщений
		std::queue<std::wstring> messageQueue;
		std::mutex mtx;
		std::condition_variable cv;
		
		// Идентификатор потока формы
		std::wstring id;

		try {	// открываем форму в новом потоке, передаём в поток ссылки на очередь сообщений, мьютекс и условную переменную
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
			form->SetMessage_TextValue("Error: Couldn't create a form in a new thread");
			return 1;
		}

		// Ожидание идентификатора из очереди сообщений потока формы
		std::wstring guid;
		{
			std::unique_lock<std::mutex> lock(mtx);
			cv.wait(lock, [&messageQueue] { return !messageQueue.empty(); });
			guid = messageQueue.front();
			messageQueue.pop();
		}
		
		int timeout = 60*1000;	// Тайм-аут в миллисекундах (1000 мс = 1 секунда), если сообщения нет, то соединение разрывается

		// Установка тайм-аута для операций чтения (recv)
		setsockopt(clientSocket, SOL_SOCKET, SO_RCVTIMEO, (const char*)&timeout, sizeof(timeout));
		String^ managedString{};

		while ((bytesReceived = recv(clientSocket, buffer, sizeof(buffer), 0)) > 0) {
            printCurrentTime();
			send(clientSocket, buffer, bytesReceived, 0);
			ChartForm::ParseBuffer(buffer, bytesReceived);
			
			if (clientPort < SclientPort) {
				// Это приём от удалённого устройства, преобразуем в HEX формат
				managedString = bufferToHex(buffer, bytesReceived);
			} else {
				// Это приём от локального клиента, встроенного в компьютер, оставляем в текстовом формате
				// Конвертация char* в String^
				managedString = gcnew String(buffer, 0, bytesReceived);
			}	// конец if

			// Найдём форму по идентификатору и обновим её
			DataForm^ form2 = DataForm::GetFormByGuid(guid);
			if (form2 != nullptr) {
				form2->Invoke(gcnew Action<String^>(form2, &DataForm::SetData_TextValue), managedString);
				form2->Invoke(gcnew Action(form2, &DataForm::Refresh));
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
		//form->SetMessage_TextValue("Attention: The thread and form was closed");
		return 0;
   }
