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

// "Старая" процедура открытия формы
void ProjectServerW::DataForm::ShowDataForm() {
    DataForm^ form = gcnew DataForm();
    form->ShowDialog();
}

// Процедура создания формы с записью идентификатора в Label_ID
void ProjectServerW::DataForm::CreateAndShowDataForm() {
    DataForm^ form = gcnew DataForm();
    form->ShowDialog();

    // Инициализация библиотеки COM
    HRESULT hr = CoInitialize(NULL);
    if (SUCCEEDED(hr)) {
        // Генерация уникального идентификатора формы
        GUID guid;
        HRESULT result = CoCreateGuid(&guid);
        if (result == S_OK) {
            // Преобразование GUID в строку
            wchar_t guidString[40] = { 0 };
            int simb_N = StringFromGUID2(guid, guidString, 40);

            String^ formId = gcnew String(guidString);

            // Установка уникального идентификатора в Label_ID
            form->SetData_FormID_value(formId);

            // Сохранение формы в карте
            formData_Map[guidString] = form;
        }
        // Освобождение библиотеки COM
        CoUninitialize();
    }
    else {
        // Обработка ошибки инициализации COM
        MessageBox::Show("Ошибка инициализации COM библиотеки", "Ошибка", MessageBoxButtons::OK, MessageBoxIcon::Error);
    }
}

void ProjectServerW::DataForm::CreateAndShowDataFormInThread(std::queue<std::wstring>& messageQueue,
                                                             std::mutex& mtx, 
                                                             std::condition_variable& cv) {
        DataForm^ form = gcnew DataForm();
    form->ShowDialog();

    // Инициализация библиотеки COM
    HRESULT hr = CoInitialize(NULL);
    if (SUCCEEDED(hr)) {
        // Генерация уникального идентификатора формы
        GUID guid;
        HRESULT result = CoCreateGuid(&guid);
        if (result == S_OK) {
            // Преобразование GUID в строку
            wchar_t guidString[40] = { 0 };
            int simb_N = StringFromGUID2(guid, guidString, 40);

            String^ formId = gcnew String(guidString);

            // Установка уникального идентификатора в Label_ID
            form->SetData_FormID_value(formId);

            // Сохранение формы в карте
            formData_Map[guidString] = form;

            // Установка значения в очередь сообщений
            {
                std::lock_guard<std::mutex> lock(mtx);
                messageQueue.push(guidString);
            }
            cv.notify_one();
        }
        // Освобождение библиотеки COM
        CoUninitialize();
    }
    else {
         // Обработка ошибки инициализации COM
         MessageBox::Show("Ошибка инициализации COM библиотеки", "Ошибка", MessageBoxButtons::OK, MessageBoxIcon::Error);
    }
}