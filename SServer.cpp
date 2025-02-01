#include "SServer.h" // 

using namespace System::Windows::Forms;
using namespace ProjectServerW;

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

	SOCKADDR_IN addr;
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

void SServer::handle() {
	while (true)
	{
		SOCKET acceptS;
		SOCKADDR_IN addr_c;
		int SclientPort = 9000; // Порт Sclient
		int ClientPort = 0;
		int addrlen = sizeof(addr_c);
		if ((acceptS = accept(this_s, (struct sockaddr*)&addr_c, &addrlen)) != INVALID_SOCKET) {
			ClientPort = ntohs(addr_c.sin_port);
			//std::cout << "Client connected from "
			//<<	to_string(addr_c.sin_addr.S_un.S_un_b.s_b1) << "." 
			//<<	to_string(addr_c.sin_addr.S_un.S_un_b.s_b2) << "."
			//<<	to_string(addr_c.sin_addr.S_un.S_un_b.s_b3) << "."
			//<<	to_string(addr_c.sin_addr.S_un.S_un_b.s_b4) << ":"
			//<<	ClientPort << std::endl;

            // Создание нового потока для обработки клиента
			DWORD threadId;
			HANDLE hThread;
			if (ClientPort >= SclientPort) {
				hThread = CreateThread(
					NULL,                   // Дескриптор безопасности
					0,                      // Начальный размер стека
					ClientTextHandler,      // Функция потока текстового клиента
					(LPVOID)acceptS,        // Параметр функции потока
					0,                      // Флаги создания
					&threadId);             // Идентификатор потока
			} else {
				hThread = CreateThread(
					NULL,                   // Дескриптор безопасности
					0,                      // Начальный размер стека
					ClientHandler,          // Функция потока
					(LPVOID)acceptS,        // Параметр функции потока
					0,                      // Флаги создания
					&threadId);             // Идентификатор потока

			}
            if (hThread == NULL) {
                std::cout << "Failed to create thread: " << GetLastError() << std::endl;
                closesocket(acceptS);
            } else {
                std::cout << "New thread is created" << std::endl 
						  << "=====================" 
						  << std::endl;
				CloseHandle(hThread); // Закрытие дескриптора потока
            }
		}
		Sleep(200);
	}
}

void printCurrentTime() {
    std::time_t now = std::time(nullptr);
    std::tm* localTime = std::localtime(&now);
    std::cout << "Current time: " << std::put_time(localTime, "%Y-%m-%d %H:%M:%S") << std::endl;
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
        SOCKET clientSocket = (SOCKET)lpParam;
        char buffer[512];
        int bytesReceived;

        // Открываем форму для демонстрации принятых данных в отдельном потоке
		// Идентификатор формы передаём через очередь сообщений
		std::queue<std::wstring> messageQueue;
		std::mutex mtx;
		std::condition_variable cv;

		try {
			std::thread formThread([&messageQueue, &mtx, &cv]() {
				ProjectServerW::DataForm::CreateAndShowDataFormInThread(messageQueue, mtx, cv);
				});
			formThread.detach(); // Отсоединяем поток от основного потока, чтобы он работал независимо
		}
		catch (const std::exception& e) {
			std::cerr << "Error creating thread: " << e.what() << std::endl;
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

		//std::wcout << L"GUID: " << guid << std::endl;

		while ((bytesReceived = recv(clientSocket, buffer, sizeof(buffer), 0)) > 0) {
            printCurrentTime();
			String^ hexStr = bufferToHex(buffer, bytesReceived);

			// Найдём форму по идентификатору и обновим её
			DataForm^ form2 = DataForm::GetFormByGuid(guid);
			if (form2 != nullptr) {
				form2->Invoke(gcnew Action<String^>(form2, &DataForm::SetData_TextValue), hexStr);
				form2->Invoke(gcnew Action(form2, &DataForm::Refresh));
			}
        }

        if (bytesReceived == 0) {
            std::cout << "Connection closed by client" << std::endl;
        } else if (bytesReceived == SOCKET_ERROR) {
            int error = WSAGetLastError();
            if (error == WSAETIMEDOUT) {
                std::cout << "Recv timed out" << std::endl;
            } else {
                std::cout << "Recv failed: " << error << std::endl;
            }
        }
        closesocket(clientSocket);
        std::cout << "The thread is closed" << std::endl 
                  << "********************" << std::endl;
        return 0;
    }
/*     int timeout = 60000;	// Тайм-аут в миллисекундах (1000 мс = 1 секунда), если сообщения нет, то соединение разрывается
	
	// Установка тайм-аута для операций чтения (recv)
	setsockopt(clientSocket, SOL_SOCKET, SO_RCVTIMEO, (const char*)&timeout, sizeof(timeout));
 */	

DWORD WINAPI SServer::ClientTextHandler(LPVOID lpParam) {
 //   SOCKET clientSocket = (SOCKET)lpParam;
 //   char buffer[512];
 //   int bytesReceived;

 //   int timeout = 60000;	// Тайм-аут в миллисекундах (1000 мс = 1 секунда), если сообщения нет, то соединение разрывается
	//
	//// Установка тайм-аута для операций чтения (recv)
	//setsockopt(clientSocket, SOL_SOCKET, SO_RCVTIMEO, (const char*)&timeout, sizeof(timeout));
	//
	//// Обработка клиента
	//	while ((bytesReceived = recv(clientSocket, buffer, sizeof(buffer), 0)) > 0) {
	//		printCurrentTime();
	//		std::cout << "Received from TEXT client: " << std::string(buffer, bytesReceived) << std::endl;
	//		send(clientSocket, buffer, bytesReceived, 0);
	//	}

	//	/*
	//	Если функция recv возвращает 0, это означает, что соединение было закрыто клиентом. 
	//	Если функция recv возвращает SOCKET_ERROR, проверяется код ошибки с помощью функции WSAGetLastError. 
	//	Если код ошибки равен WSAETIMEDOUT, это означает, что операция чтения завершилась по тайм-ауту.
	//	*/
	//	if (bytesReceived == 0) {
	//		std::cout << "Connection closed by client" << std::endl;
	//	} else if (bytesReceived == SOCKET_ERROR) {
	//		int error = WSAGetLastError();
	//		if (error == WSAETIMEDOUT) {
	//			std::cout << "Recv timed out" << std::endl;
	//		} else {
	//			std::cout << "Recv failed: " << error << std::endl;
	//		}
 //   }
	//closesocket(clientSocket);
	//std::cout << "The thread is closed" << std::endl 
	//		  << "********************" << std::endl;
    return 0;
}