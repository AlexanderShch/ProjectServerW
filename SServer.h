#ifndef SSERVER_H
#define SSERVER_H

#pragma once

//#pragma comment(lib, "ws2_32.lib")
#include <winsock2.h>       // Должен быть первым!
#include <ws2tcpip.h>       // После winsock2.h
#include <windows.h>        // Windows.h должен быть после winsock2.h Для CreateThread

#include <ctime>			// Для std::time
#include <iomanip>			// Для std::put_time
#include <iostream>
//#include <sstream>		// Для std::stringstream

#include "MyForm.h"
#include "DataForm.h"

using namespace std;

class SServer
{
public:
	SServer();			// Объявление конструктора
	~SServer();			// Объявление деструктора
	void startServer(); // Объявление функции
	void closeServer();
	void handle();
    int port;
    void ShowDataForm(); // Объявление функции
private:
	SOCKET this_s;
		WSAData wData;
    static DWORD WINAPI ClientHandler(LPVOID lpParam); // Объявление функции потока
    static DWORD WINAPI ClientTextHandler(LPVOID lpParam); // Объявление функции потока
};

#endif // SSERVER_H