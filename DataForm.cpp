#include "DataForm.h"
#include <objbase.h>                // ��� CoCreateGuid - ��������� ����������� ��������������
#include <map>						// ��� ������������� std::map - ���������, ����������� ������������ ID � ������ �� �����
#include <string>
#include <vcclr.h>					// ��� ������������� gcroot

using namespace ProjectServerW; // ��������� ������������ ����

std::map<std::wstring, gcroot<DataForm^>> formData_Map; // ����������� ���������� formData_Map

System::Void ProjectServerW::DataForm::�����ToolStripMenuItem_Click(System::Object^ sender, System::EventArgs^ e)
{
	Application::Exit();
	return System::Void();
}

void ProjectServerW::DataForm::ShowDataForm() {
    DataForm^ form = gcnew DataForm();
    form->ShowDialog();
}

