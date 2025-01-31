#include "DataForm.h"
#include <objbase.h>                // Для CoCreateGuid - генерация уникального идентификатора
#include <map>						// Для использования std::map - структуры, сохраняющей соответствие ID и ссылки на форму
#include <string>
#include <vcclr.h>					// Для использования gcroot

using namespace ProjectServerW; // Добавлено пространство имен

std::map<std::wstring, gcroot<DataForm^>> formData_Map; // Определение переменной formData_Map

System::Void ProjectServerW::DataForm::выходToolStripMenuItem_Click(System::Object^ sender, System::EventArgs^ e)
{
	Application::Exit();
	return System::Void();
}

void ProjectServerW::DataForm::ShowDataForm() {
    DataForm^ form = gcnew DataForm();
    form->ShowDialog();
}

