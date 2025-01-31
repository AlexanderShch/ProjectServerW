#include "DataForm.h"

using namespace ProjectServerW; // Äîáàâëåíî ïðîñòðàíñòâî èìåí

System::Void ProjectServerW::DataForm::âûõîäToolStripMenuItem_Click(System::Object^ sender, System::EventArgs^ e)
{
	Application::Exit();
	return System::Void();
}

void ProjectServerW::DataForm::ShowDataForm() {
    DataForm^ form2 = gcnew DataForm();
    form2->ShowDialog();
}

