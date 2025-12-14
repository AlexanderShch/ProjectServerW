#include "MyForm.h"
#include "SServer.h"

using namespace ProjectServerW;		// ˜˜˜˜˜˜˜˜ ˜˜˜˜˜˜˜
SServer server;

int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nCmdShow) {
	Application::EnableVisualStyles();
	Application::SetCompatibleTextRenderingDefault(false);
	Application::Run(gcnew MyForm);

	server.closeServer();
	return 0;

}

System::Void ProjectServerW::MyForm::˜˜˜˜˜ToolStripMenuItem_Click(System::Object^ sender, System::EventArgs^ e)
{
	Application::Exit();
	return System::Void();
}

System::Void ProjectServerW::MyForm::button_Listen_Click(System::Object^ sender, System::EventArgs^ e)
{
	server.port = 3487;
	
	// ˜˜˜˜˜˜˜˜ ˜˜˜˜˜˜ ˜ ˜˜˜˜˜˜˜˜˜ ˜˜˜˜˜˜, ˜˜˜˜˜ ˜˜ ˜˜˜˜˜˜˜˜˜˜˜ ˜˜˜˜˜ ˜˜˜˜˜˜˜˜˜ ˜˜˜˜
	std::thread serverThread(&SServer::startServer, &server);
	serverThread.detach(); // ˜˜˜˜˜˜˜˜˜˜˜ ˜˜˜˜˜ ˜˜˜˜˜˜˜, ˜˜˜˜˜ ˜˜ ˜˜˜˜˜˜˜ ˜˜˜˜˜˜˜˜˜˜
}


