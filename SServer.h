#ifndef SSERVER_H
#define SSERVER_H

#pragma once
#pragma comment(lib, "Ws2_32.lib")

#include <winsock2.h>       // ������ ���� ������!
#include <ws2tcpip.h>       // ����� winsock2.h
#include <windows.h>        // Windows.h ������ ���� ����� winsock2.h ��� CreateThread

#include <ctime>			// ��� std::time
#include <iomanip>			// ��� std::put_time
#include <iostream>
#include <sstream>			// ��� std::stringstream
#include <thread>
#include <msclr\marshal_cppstd.h>  // ��� ������ � �����������

// ���������� ��� ������ � Unicode � ���������������
namespace Unicode {
    int snprintfFloat(char* buffer, size_t size, const char* format, float value);
}

#include "MyForm.h"
#include "DataForm.h"

using namespace std;
using namespace System;
using namespace System::IO;
using namespace System::Threading;
using namespace System::Windows::Forms;

// ����� ��� �������������� ��������
extern char ValueCoreT1SmallBuffer[256];

class SServer
{
public:
    SServer();			// ���������� ������������
    ~SServer();			// ���������� �����������
    void startServer(); // ���������� �������
    void closeServer();
	void handle();
    int port;
private:
	SOCKET this_s;
	WSAData wData;
    static DWORD WINAPI ClientHandler(LPVOID lpParam); // ���������� ������� ������
};

ref class GlobalLogger abstract sealed { // abstract sealed = static class
private:
    static StreamWriter^ writer;
    static Object^ lockObject;
    static bool isInitialized;

    // ����������� ����������� ��� ������������� �����
    static GlobalLogger() {
        writer = nullptr;
        lockObject = gcnew Object();
        isInitialized = false;
    }

public:
    static void Initialize() {
        if (!isInitialized) {
            try {
                String^ appPath = System::IO::Path::GetDirectoryName(System::Windows::Forms::Application::ExecutablePath);
                String^ logFilePath = System::IO::Path::Combine(appPath, "log.txt");

                FileStream^ fs = gcnew FileStream(
                    logFilePath,
                    FileMode::Append,
                    FileAccess::Write,
                    FileShare::Read
                );

                writer = gcnew StreamWriter(fs, System::Text::Encoding::UTF8);
                isInitialized = true;
            }
            catch (Exception^ ex) {
                MessageBox::Show("Failed to initialize logger: " + ex->Message);
            }
        }
    }

    static void LogMessage(const std::string& message) {
        if (!isInitialized || writer == nullptr) {
            Initialize();
        }

        try {
            Monitor::Enter(lockObject);
            try {
                String^ timestamp = DateTime::Now.ToString("dd.MM.yy HH:mm:ss");
                String^ managedMessage = gcnew String(message.c_str());
                writer->WriteLine(timestamp + " : " + managedMessage);
                writer->Flush();
            }
            finally {
                Monitor::Exit(lockObject);
            }
        }
        catch (Exception^ ex) {
            MessageBox::Show("Error writing to log: " + ex->Message);
        }
    }

    static void Shutdown() {
        if (writer != nullptr) {
            try {
                writer->Flush();
                writer->Close();
                delete writer;
                writer = nullptr;
                isInitialized = false;
            }
            catch (Exception^ ex) {
                MessageBox::Show("Error closing log file: " + ex->Message);
            }
        }
    }
};

std::string ConvertToStdString(System::String^ managedString);

#endif // SSERVER_H

