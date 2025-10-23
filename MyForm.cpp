#include "MyForm.h"
#include "SServer.h"
#include <thread>

using namespace ProjectServerW;		// �������� �������
SServer server;

int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nCmdShow) {
	Application::EnableVisualStyles();
	Application::SetCompatibleTextRenderingDefault(false);
	Application::Run(gcnew MyForm);

	server.closeServer();
	return 0;

}

System::Void ProjectServerW::MyForm::�����ToolStripMenuItem_Click(System::Object^ sender, System::EventArgs^ e)
{
	// ������� ��������� ������
	server.closeServer();
	// ���� ����� ������� �����������
	System::Threading::Thread::Sleep(500);
	// ������ ��������� ����������
	Application::Exit();
	return System::Void();
}

System::Void ProjectServerW::MyForm::button_Listen_Click(System::Object^ sender, System::EventArgs^ e)
{
	server.port = 3487;
	
	// �������� ������ � ��������� ������, ����� �� ����������� ����� ��������� ����
	std::thread serverThread(&SServer::startServer, &server);
	serverThread.detach(); // ����������� ����� �������, ����� �� ������� ����������
}

System::Void ProjectServerW::MyForm::MyForm_Load(System::Object^ sender, System::EventArgs^ e)
{
	// ��������� ������ ����� ������ �������� �����
	button_Listen_Click(nullptr, nullptr);
}
