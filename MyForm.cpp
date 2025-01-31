#include "MyForm.h"
#include "SServer.h"

using namespace ProjectServerW;		// название проекта
SServer server;

int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nCmdShow) {
	Application::EnableVisualStyles();
	Application::SetCompatibleTextRenderingDefault(false);
	Application::Run(gcnew MyForm);

	server.closeServer();
	return 0;

}

System::Void ProjectServerW::MyForm::выходToolStripMenuItem_Click(System::Object^ sender, System::EventArgs^ e)
{
	Application::Exit();
	return System::Void();
}

System::Void ProjectServerW::MyForm::button_Listen_Click(System::Object^ sender, System::EventArgs^ e)
{
	server.port = 3487;
	
	// Запустим сервер в отдельном потоке, чтобы не блокировать поток основного окна
	std::thread serverThread(&SServer::startServer, &server);
	serverThread.detach(); // Отсоединяем поток сервера, чтобы он работал независимо
}


