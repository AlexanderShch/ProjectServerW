#ifndef SSERVER_H
#define SSERVER_H

#pragma once
#pragma comment(lib, "Ws2_32.lib")

#include <winsock2.h>       // Должен быть первым!
#include <ws2tcpip.h>       // После winsock2.h
#include <windows.h>        // Windows.h должен быть после winsock2.h Для CreateThread

#include <climits>
#include <cstring>
#include <ctime>			// Для std::time
#include <iomanip>			// Для std::put_time
#include <iostream>
#include <sstream>			// Для std::stringstream
#include <thread>
#include <msclr\marshal_cppstd.h>  // Для работы с кодировками

// Объявления для работы с Unicode и форматированием
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

// Буфер для форматирования значений
extern char ValueCoreT1SmallBuffer[256];

class SServer
{
public:
    SServer();			// Объявление конструктора
    ~SServer();			// Объявление деструктора
    void startServer(); // Объявление функции
    void closeServer();
	void handle();
    int port;
private:
	SOCKET this_s;
	WSAData wData;
    static DWORD WINAPI ClientHandler(LPVOID lpParam); // Объявление функции потока
};

ref class GlobalLogger abstract sealed { // abstract sealed = static class
private:
    static StreamWriter^ writer;
    static Object^ lockObject;
    static bool isInitialized;

    // Статический конструктор для инициализации полей
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
                // std::string: если из ConvertToStdString — байты в CP1251; декодируем в Unicode, файл пишется в UTF-8
                const size_t msgLenSize = message.size();
                const int msgLen = (msgLenSize <= static_cast<size_t>(INT_MAX)) ? static_cast<int>(msgLenSize) : INT_MAX;
                cli::array<System::Byte>^ msgBytes = gcnew cli::array<System::Byte>(msgLen);
                for (int i = 0; i < msgLen; ++i)
                    msgBytes[i] = static_cast<System::Byte>(message[static_cast<size_t>(i)]);
                String^ managedMessage = System::Text::Encoding::GetEncoding(1251)->GetString(msgBytes);
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

    // Запись в лог управляемой строки (Unicode). Файл log.txt в UTF-8.
    static void LogMessage(String^ managedMessage) {
        if (!isInitialized || writer == nullptr) {
            Initialize();
        }
        if (managedMessage == nullptr) return;
        try {
            Monitor::Enter(lockObject);
            try {
                String^ timestamp = DateTime::Now.ToString("dd.MM.yy HH:mm:ss");
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

    // Узкий литерал в UTF-8 (исходник в UTF-8): байты → String^ → запись в лог (файл UTF-8).
    static void LogMessage(const char* utf8Message) {
        if (utf8Message == nullptr) return;
        size_t len = std::strlen(utf8Message);
        if (len == 0) return;
        if (len > static_cast<size_t>(INT_MAX)) len = static_cast<size_t>(INT_MAX);
        cli::array<System::Byte>^ bytes = gcnew cli::array<System::Byte>(static_cast<int>(len));
        System::Runtime::InteropServices::Marshal::Copy(IntPtr(const_cast<char*>(utf8Message)), bytes, 0, static_cast<int>(len));
        String^ s = System::Text::Encoding::UTF8->GetString(bytes);
        LogMessage(s);
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

