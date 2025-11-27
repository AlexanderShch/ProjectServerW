#include "MyForm.h"
#include "SServer.h"
#include <thread>

using namespace ProjectServerW;		// название проекта
SServer server;

// Global exception handler for unhandled exceptions
void UnhandledExceptionHandler(Object^ sender, UnhandledExceptionEventArgs^ e) {
	try {
		Exception^ ex = dynamic_cast<Exception^>(e->ExceptionObject);
		if (ex != nullptr) {
			GlobalLogger::LogMessage(ConvertToStdString(String::Format(
				"FATAL UNHANDLED EXCEPTION: {0}\nStackTrace: {1}\nTerminating: {2}",
				ex->Message, ex->StackTrace, e->IsTerminating)));
		}
		else {
			GlobalLogger::LogMessage("FATAL UNHANDLED NON-MANAGED EXCEPTION");
		}
		GlobalLogger::Shutdown();
	}
	catch (...) {
		// Last resort - do nothing
	}
}

// Thread exception handler for UI thread exceptions
void ThreadExceptionHandler(Object^ sender, System::Threading::ThreadExceptionEventArgs^ e) {
	try {
		GlobalLogger::LogMessage(ConvertToStdString(String::Format(
			"THREAD EXCEPTION: {0}\nStackTrace: {1}",
			e->Exception->Message, e->Exception->StackTrace)));
	}
	catch (...) {
		// Last resort - do nothing
	}
}

// Server thread function with exception handling
void ServerThreadFunction() {
	try {
		server.startServer();
	}
	catch (Exception^ ex) {
		GlobalLogger::LogMessage(ConvertToStdString(String::Format(
			"CRITICAL: Managed exception in server thread: {0}\nStackTrace: {1}",
			ex->Message, ex->StackTrace)));
	}
	catch (const std::exception& ex) {
		GlobalLogger::LogMessage(std::string("CRITICAL: Native exception in server thread: ") + ex.what());
	}
	catch (...) {
		GlobalLogger::LogMessage("CRITICAL: Unknown exception in server thread");
	}
}

int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nCmdShow) {
	try {
		// Initialize global logger early
		GlobalLogger::Initialize();
		
		// Set up global exception handlers
		System::AppDomain::CurrentDomain->UnhandledException += 
			gcnew UnhandledExceptionEventHandler(&UnhandledExceptionHandler);
		Application::ThreadException += 
			gcnew System::Threading::ThreadExceptionEventHandler(&ThreadExceptionHandler);
		
		// Enable unhandled exception mode to catch all exceptions
		Application::SetUnhandledExceptionMode(UnhandledExceptionMode::CatchException);
		
		GlobalLogger::LogMessage("======================================================================");
		GlobalLogger::LogMessage("Application starting...");
		GlobalLogger::LogMessage("======================================================================");
		
		Application::EnableVisualStyles();
		Application::SetCompatibleTextRenderingDefault(false);
		Application::Run(gcnew MyForm);

		GlobalLogger::LogMessage("======================================================================");
		GlobalLogger::LogMessage("Application closing normally...");
		GlobalLogger::LogMessage("======================================================================");
		
		server.closeServer();
		GlobalLogger::Shutdown();
		return 0;
	}
	catch (Exception^ ex) {
		try {
			GlobalLogger::LogMessage(ConvertToStdString(String::Format(
				"FATAL: Unhandled managed exception in WinMain: {0}\nStackTrace: {1}",
				ex->Message, ex->StackTrace)));
			GlobalLogger::Shutdown();
		}
		catch (...) {}
		return 1;
	}
	catch (const std::exception& ex) {
		try {
			GlobalLogger::LogMessage(std::string("FATAL: Unhandled native exception in WinMain: ") + ex.what());
			GlobalLogger::Shutdown();
		}
		catch (...) {}
		return 1;
	}
	catch (...) {
		try {
			GlobalLogger::LogMessage("FATAL: Unhandled unknown exception in WinMain");
			GlobalLogger::Shutdown();
		}
		catch (...) {}
		return 1;
	}
}

System::Void ProjectServerW::MyForm::выходToolStripMenuItem_Click(System::Object^ sender, System::EventArgs^ e)
{
	// Сначала закрываем сервер
	server.closeServer();
	// Даем время потокам завершиться
	System::Threading::Thread::Sleep(500);
	// Теперь закрываем приложение
	Application::Exit();
	return System::Void();
}

System::Void ProjectServerW::MyForm::button_Listen_Click(System::Object^ sender, System::EventArgs^ e)
{
	try {
		server.port = 3487;
		
		// Запустим сервер в отдельном потоке, чтобы не блокировать поток основного окна
		std::thread serverThread(&ServerThreadFunction);
		serverThread.detach(); // Отсоединяем поток сервера, чтобы он работал независимо
	}
	catch (Exception^ ex) {
		GlobalLogger::LogMessage(ConvertToStdString(String::Format(
			"Error starting server thread: {0}\nStackTrace: {1}",
			ex->Message, ex->StackTrace)));
	}
	catch (const std::exception& ex) {
		GlobalLogger::LogMessage(std::string("Error starting server thread: ") + ex.what());
	}
	catch (...) {
		GlobalLogger::LogMessage("Unknown error starting server thread");
	}
}

System::Void ProjectServerW::MyForm::MyForm_Load(System::Object^ sender, System::EventArgs^ e)
{
	// Запускаем сервер после полной загрузки формы
	button_Listen_Click(nullptr, nullptr);
}