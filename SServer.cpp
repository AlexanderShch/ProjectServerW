#include "SServer.h" // �������� DataForm.h ������ ����
#include <fstream>
#include <iostream>
#include <chrono>
#include <ctime>
#include <msclr/marshal_cppstd.h>
#include <cstdio>

// �������� ��� ������ ��� ������� � Marshal
using namespace System::Runtime::InteropServices;

using namespace System::Windows::Forms;
using namespace ProjectServerW;

std::map<std::wstring, std::thread> formThreads;  // �������� ������� ����

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

// ����������� ������������ ������ SServer
SServer::SServer() : port(0), this_s(INVALID_SOCKET) {
}

// ����������� ����������� ������ SServer
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
		// ���������� ����� � textBox_ListPort
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

// �������������� IP-������ � ������
String^ GetIPString(const SOCKADDR_IN& addr_c) {
	String^ ipString = String::Format("{0}.{1}.{2}.{3}",
		addr_c.sin_addr.S_un.S_un_b.s_b1,
		addr_c.sin_addr.S_un.S_un_b.s_b2,
		addr_c.sin_addr.S_un.S_un_b.s_b3,
		addr_c.sin_addr.S_un.S_un_b.s_b4);

	// ��������� ����
	int port = ntohs(addr_c.sin_port);
	String^ fullAddress = String::Format("{0}:{1}", ipString, port);

	return fullAddress;
}

void SServer::handle() {
	while (true)
	{
		SOCKET acceptS;
		SOCKADDR_IN addr_c{};
		// ������� ����� MyForm ��� ������ ��������� � ������� � ������
		MyForm^ form = safe_cast<MyForm^>(Application::OpenForms["MyForm"]);

		// ���������, ���������� �� �����
		if (form == nullptr || form->IsDisposed) {
			break; // ������� �� �����, ���� ����� �������
		}

		int addrlen = sizeof(addr_c);
		if ((acceptS = accept(this_s, (struct sockaddr*)&addr_c, &addrlen)) != INVALID_SOCKET) {
			int ClientPort = ntohs(addr_c.sin_port);
			String^ address = GetIPString(addr_c);
			if (form != nullptr && !form->IsDisposed) {
				form->SetClientAddr_TextValue(address);
			}

            // �������� ������ ������ ��� ��������� �������
			DWORD threadId;
			HANDLE hThread;

			hThread = CreateThread(
				NULL,                   // ���������� ������������
				0,                      // ��������� ������ �����
				ClientHandler,          // ������� ������
				(LPVOID)acceptS,		// �������� ������� ������
				0,                      // ����� ��������
				&threadId);             // ������������� ������

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
				CloseHandle(hThread); // �������� ����������� ������
            }
		}
		Sleep(200);
	}
}

// ��������� ����������� CRC
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
	int SclientPort = 900000; // ���� Sclient

	SOCKET clientSocket = (SOCKET)lpParam;
    char buffer[512];
    int bytesReceived;
	// ������� ����� MyForm ��� ������ ��������� �� �������
	MyForm^ form = safe_cast<MyForm^>(Application::OpenForms["MyForm"]);

	// �������� ���������� � �������
	if (getpeername(clientSocket, (SOCKADDR*)&clientAddr, &addrLen) != SOCKET_ERROR) {
		// ����������� ���� �� �������� � ��������� ������
		clientPort = ntohs(clientAddr.sin_port);
	}

	// �������� IP-����� �������
	String^ clientIPAddress = String::Format("{0}.{1}.{2}.{3}",
		clientAddr.sin_addr.S_un.S_un_b.s_b1,
		clientAddr.sin_addr.S_un.S_un_b.s_b2,
		clientAddr.sin_addr.S_un.S_un_b.s_b3,
		clientAddr.sin_addr.S_un.S_un_b.s_b4);

	// ���������, ���� �� ��� �������� ���������� �� ����� IP
	std::wstring existingGuid = ProjectServerW::DataForm::FindFormByClientIP(clientIPAddress);
	if (!existingGuid.empty()) {
		// ��� ���� �������� ���������� �� ����� IP
		String^ warningMsg = "��������: ��� ���������� �������� ���������� �� ������� " + clientIPAddress +
			". ����� ���������� �� ����� �������.";
		if (form != nullptr && !form->IsDisposed) {
			form->SetMessage_TextValue(warningMsg);
			GlobalLogger::LogMessage(ConvertToStdString(warningMsg));
		}
		// ��������� ����� ����������
		closesocket(clientSocket);
		return 1;
	}

	// ��������� ����� DataForm ��� ������������ �������� ������ � ��������� ������
	// ������������� ����� ���������� ������� ����� ������� ���������
	std::wstring guid;
	std::queue<std::wstring> messageQueue;
	std::mutex mtx;
	std::condition_variable cv;
		
	try {	// ��������� ����� DataForm � ����� ������, ������� � ����� ������ �� ������� ���������, ������� � �������� ����������
		std::thread formThread([&messageQueue, &mtx, &cv]() {
			ProjectServerW::DataForm::CreateAndShowDataFormInThread(messageQueue, mtx, cv);
			});
		// ��������� �������������� ����� ������ - guid
		// ������� ��������� ����� ��� ��������, ��� ������ �� ������� ��������� ������� ��������������
		{
			std::unique_lock<std::mutex> lock(mtx);
			cv.wait(lock, [&messageQueue] { return !messageQueue.empty(); });
			guid = messageQueue.front();
			messageQueue.pop();
		}
		// ��������� ����� ����� ������ � ��������� � GUID ����� � �������� �����
		ThreadStorage::StoreThread(guid, formThread);
		GlobalLogger::LogMessage(ConvertToStdString("Information: The DataForm has been opened successfully!"));
	}
	catch (const std::exception& e) {	// ��������� ����������, ������� ��������� �� ������
		String^ errorMessage = gcnew String(e.what());
		if (form != nullptr && !form->IsDisposed) {
			form->SetMessage_TextValue("Error: Couldn't create a form in a new thread " + errorMessage);
			GlobalLogger::LogMessage(ConvertToStdString("Error: Couldn't create a DataForm in a new thread " + errorMessage));
		}
		return 1;
	}

	int timeout = 30*60*1000;	// ����-��� � ������������� (1000 �� = 1 �������), ���� ��������� ���, �� ���������� �����������

	// ��������� ����-���� ��� �������� ������ (recv)
	setsockopt(clientSocket, SOL_SOCKET, SO_RCVTIMEO, (const char*)&timeout, sizeof(timeout));

	// ����������� ���� ���������� ������
	while (true) {
		bytesReceived = recv(clientSocket, buffer, sizeof(buffer), 0);

		/* ��������� ������
		���� ������� recv ���������� 0, ��� ��������, ��� ���������� ���� ������� ��������.
		���� ������� recv ���������� SOCKET_ERROR, ����������� ��� ������ � ������� ������� WSAGetLastError.
		���� ��� ������ ����� WSAETIMEDOUT, ��� ��������, ��� �������� ������ ����������� �� ����-����.
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
			break;	// ������� �� �����
		}
		if (bytesReceived == 0) {
			if (form != nullptr && !form->IsDisposed) {
				form->SetMessage_TextValue("Attention: Connection closed by client");
				GlobalLogger::LogMessage("Attention: Connection closed by client");
			}
			break;	// ������� �� �����
		};

		// ===== ���������� ���� ������ �� ������� ����� =====
		uint8_t packetType = buffer[0];
		
		if (packetType >= 0x01 && packetType <= 0x04) {
			// ===== ��� ����� �� ������� =====
			// ������ � Type = 0x01-0x04 ��� ������ �� ������� �� �����������
			
			// ������� ����������� ������ ��� ������
			cli::array<System::Byte>^ responseBuffer = gcnew cli::array<System::Byte>(bytesReceived);
			Marshal::Copy(IntPtr(buffer), responseBuffer, 0, bytesReceived);
			
		// ��������� � ������� �������
		DataForm::EnqueueResponse(responseBuffer);
		
		GlobalLogger::LogMessage(ConvertToStdString(String::Format(
			"Response received and enqueued: Type=0x{0:X2}, Size={1} bytes", 
			packetType, bytesReceived)));
		}
		else if (packetType == 0x00) {
			// ===== ��� ���������� =====
			// ������ � Type = 0x00 ��� ���������� �� �����������
			
			// ����� ������� ������ � CRC
			uint8_t LengthOfPackage = 48;
			// �������� CRC16 ��� ������ 46 ����
			uint16_t dataCRC = MB_GetCRC(buffer, LengthOfPackage - 2);
			uint16_t DatCRC;
			memcpy(&DatCRC, &buffer[LengthOfPackage - 2], 2);

			if (DatCRC == dataCRC) {
				// ���������� ���������� ������� �������
				// send(clientSocket, buffer, bytesReceived, 0);

				// ����� ����� �� ��������������
				DataForm^ form2 = DataForm::GetFormByGuid(guid);
				if (form2 != nullptr && !form2->IsDisposed && form2->IsHandleCreated && !form2->Disposing) 
				{
					
					// ��������� � ����� ����� ������� ���� ����� � IP-�����
					if (form2 != nullptr) {
						form2->ClientSocket = clientSocket;
						form2->ClientIP = clientIPAddress;
					}
					
					if (clientPort < SclientPort) {
						// ������� ����� ������ ��� ���������� �������� � ������ �����
						cli::array<System::Byte>^ dataBuffer = gcnew cli::array<System::Byte>(bytesReceived);
						Marshal::Copy(IntPtr(buffer), dataBuffer, 0, bytesReceived);

						// �������� AddDataToTable ����� Invoke ��� ���������� � ������ �����
						form2->BeginInvoke(gcnew Action<cli::array <System::Byte>^, int, int>
							(form2, &DataForm::AddDataToTableThreadSafe),
							dataBuffer, bytesReceived, clientPort);
						// Refresh �������� �������� ��� �� ����� - �� ����� ������ � AddDataToTableThreadSafe
					} else {
						// ����� ��������� ��� ������� ���� �����
					}
				}
				else {
					// ����� �����������, ��������� �����
					break;
		} // ����� if (form2 != nullptr...
	} // if (DatCRC == dataCRC)
	else {
		// ������ CRC ����������
		GlobalLogger::LogMessage(ConvertToStdString(String::Format(
			"Telemetry CRC error: Expected={0:X4}, Received={1:X4}", 
			dataCRC, DatCRC)));
	}
} // else if (packetType == 0x00)
else {
	// ����������� ��� ������
	GlobalLogger::LogMessage(ConvertToStdString(String::Format(
		"Unknown packet type: 0x{0:X2}, Size={1} bytes", 
		packetType, bytesReceived)));
		}
	}	// ����� while

	// ���� ����� ������ ������� �� ������ �����, ��������� ������ ����� ������
	closesocket(clientSocket);
	// ����� ����� ������ �� �������������� � ������� �
	DataForm::CloseForm(guid);
	// �������� ������ ����� ������ �� guid �����
	ThreadStorage::StopThread(guid);

	return 0;
}

std::string ConvertToStdString(System::String^ managedString) {

	if (managedString == nullptr)
		return std::string();

	// ���������� UTF8 ��������� ��� �����������
//	cli::array<System::Byte>^ bytes = System::Text::Encoding::UTF8->GetBytes(managedString);
	// ���������� 1251 ��������� ��� �����������
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