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

// "������" ��������� �������� �����
void ProjectServerW::DataForm::ShowDataForm() {
    DataForm^ form = gcnew DataForm();
    form->ShowDialog();
}

// ��������� �������� ����� � ������� �������������� � Label_ID
void ProjectServerW::DataForm::CreateAndShowDataForm() {
    DataForm^ form = gcnew DataForm();
    form->ShowDialog();

    // ������������� ���������� COM
    HRESULT hr = CoInitialize(NULL);
    if (SUCCEEDED(hr)) {
        // ��������� ����������� �������������� �����
        GUID guid;
        HRESULT result = CoCreateGuid(&guid);
        if (result == S_OK) {
            // �������������� GUID � ������
            wchar_t guidString[40] = { 0 };
            int simb_N = StringFromGUID2(guid, guidString, 40);

            String^ formId = gcnew String(guidString);

            // ��������� ����������� �������������� � Label_ID
            form->SetData_FormID_value(formId);

            // ���������� ����� � �����
            formData_Map[guidString] = form;
        }
        // ������������ ���������� COM
        CoUninitialize();
    }
    else {
        // ��������� ������ ������������� COM
        MessageBox::Show("������ ������������� COM ����������", "������", MessageBoxButtons::OK, MessageBoxIcon::Error);
    }
}

void ProjectServerW::DataForm::CreateAndShowDataFormInThread(std::queue<std::wstring>& messageQueue,
                                                             std::mutex& mtx, 
                                                             std::condition_variable& cv) {
        DataForm^ form = gcnew DataForm();
    form->ShowDialog();

    // ������������� ���������� COM
    HRESULT hr = CoInitialize(NULL);
    if (SUCCEEDED(hr)) {
        // ��������� ����������� �������������� �����
        GUID guid;
        HRESULT result = CoCreateGuid(&guid);
        if (result == S_OK) {
            // �������������� GUID � ������
            wchar_t guidString[40] = { 0 };
            int simb_N = StringFromGUID2(guid, guidString, 40);

            String^ formId = gcnew String(guidString);

            // ��������� ����������� �������������� � Label_ID
            form->SetData_FormID_value(formId);

            // ���������� ����� � �����
            formData_Map[guidString] = form;

            // ��������� �������� � ������� ���������
            {
                std::lock_guard<std::mutex> lock(mtx);
                messageQueue.push(guidString);
            }
            cv.notify_one();
        }
        // ������������ ���������� COM
        CoUninitialize();
    }
    else {
         // ��������� ������ ������������� COM
         MessageBox::Show("������ ������������� COM ����������", "������", MessageBoxButtons::OK, MessageBoxIcon::Error);
    }
}