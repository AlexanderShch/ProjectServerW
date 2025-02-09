#include "DataForm.h"
#include <objbase.h>                // Для CoCreateGuid - генерация уникального идентификатора
#include <string>
#include <vcclr.h>					// Для использования gcroot

using namespace ProjectServerW; // Добавлено пространство имен

std::map<std::wstring, gcroot<DataForm^>> formData_Map; // Определение переменной formData_Map

typedef struct   // object data for Server type из STM32
{
    uint16_t Time;				// Количество секунд с момента включения
    uint8_t SensorQuantity;		// Количество сенсоров
    uint8_t SensorType[SQ];		// Тип сенсора
    uint8_t Active[SQ];			// Активность сенсора
    uint16_t T[SQ];				// Значение 1 сенсора (температура)
    uint16_t H[SQ];				// Значение 2 сенсора (влажность)
    uint16_t CRC_SUM;			// Контрольное значение
} MSGQUEUE_OBJ_t;

// Преобразование данных из буфера в структуру типа MSGQUEUE_OBJ_t из STM32
void ProjectServerW::DataForm::ParseBuffer(const char* buffer, size_t size) {

    MSGQUEUE_OBJ_t data;

    if (size < sizeof(data)) {
        return; // Принятых данных слишком мало
    }
    memcpy(&data, buffer, sizeof(data));
}

System::Void ProjectServerW::DataForm::выходToolStripMenuItem_Click(System::Object^ sender, System::EventArgs^ e)
{
	Application::Exit();
	return System::Void();
}

void ProjectServerW::DataForm::CreateAndShowDataFormInThread(std::queue<std::wstring>& messageQueue,
                                                             std::mutex& mtx, 
                                                             std::condition_variable& cv) {
    DataForm^ form = gcnew DataForm();

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
			form->Refresh();

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

	form->ShowDialog();
}

// Метод закрытия формы
void ProjectServerW::DataForm::CloseForm(const std::wstring& guid) {
    // Находим форму
    ProjectServerW::DataForm^ form = ProjectServerW::DataForm::GetFormByGuid(guid);

    if (form != nullptr) {
        // Проверяем, нужен ли Invoke
        if (form->InvokeRequired) {
            // Закрываем форму через Invoke
            form->Invoke(gcnew Action(form, &ProjectServerW::DataForm::Close));
        }
        else {
            form->Close();
        }
    }
}
        
// Метод для получения формы по её GUID
DataForm^ ProjectServerW::DataForm::GetFormByGuid(const std::wstring& guid) {
    auto it = formData_Map.find(guid);
    if (it != formData_Map.end()) {
        return it->second;
    }
    return nullptr;
}

// Записываем пару guid и поток в map
void ThreadStorage::StoreThread(const std::wstring& guid, std::thread& thread) {
    std::lock_guard<std::mutex> lock(GetMutex());
    GetThreadMap()[guid] = std::move(thread);
}

// Останавливаем поток по guid
void ThreadStorage::StopThread(const std::wstring& guid)
{
    std::lock_guard<std::mutex> lock(GetMutex());

    // Ищем поток
    auto it = GetThreadMap().find(guid);
    if (it != GetThreadMap().end() && it->second.joinable()) {
        it->second.join();  // Ждем завершения
        GetThreadMap().erase(it);  // Удаляем из карты
    }
}

// Функция для определения статической переменной map для потока
std::map<std::wstring, std::thread>& ThreadStorage::GetThreadMap() {
    static std::map<std::wstring, std::thread> threadMap;
    return threadMap;
}

// Функция для определения статической переменной Mutex для потока
std::mutex& ThreadStorage::GetMutex() {
    static std::mutex mtx;
    return mtx;
}

/*
*********** Создаём таблицу *******************
*/

// 1. Создаём таблицу данных
void ProjectServerW::DataForm::InitializeDataTable() {
    // Создаем таблицу данных
    dataTable = gcnew DataTable("SensorData");
    dataTable->Columns->Add("Time", uint16_t::typeid);
    dataTable->Columns->Add("SQ", uint8_t::typeid);

    //DataGridViewCellStyle^ headerStyle = gcnew DataGridViewCellStyle();
    //headerStyle->Alignment = DataGridViewContentAlignment::MiddleLeft;
    //headerStyle->WrapMode = DataGridViewTriState::True;


    for (uint8_t i = 0; i < SQ; i++)
    {
        dataTable->Columns->Add("Typ" + i, uint8_t::typeid);
        dataTable->Columns->Add("Act" + i, uint8_t::typeid);
        dataTable->Columns->Add("T" + i, uint16_t::typeid);
        dataTable->Columns->Add("H" + i, uint16_t::typeid);
    }
    dataGridView->DataSource = dataTable;
}

// 2. Добавление данных
void DataForm::AddDataToTable(DataTable^ table) {
    // Создание новой строки
    DataRow^ row = table->NewRow();
    row["Time"] = 1010;// data.Time;
    row["SQ"] = 8080;

    // Добавление строки в таблицу
    table->Rows->Add(row);
}

// 3. Сохраняем таблицу в EXCEL
// Установите NuGet пакет EPPlus
//using namespace OfficeOpenXml;
//
//void SaveToExcelEPPlus(DataTable^ table, String^ filePath) {
//    ExcelPackage^ package = gcnew ExcelPackage();
//    ExcelWorksheet^ worksheet = package->Workbook->Worksheets->Add("Sheet1");
//
//    // Записываем заголовки
//    for (int i = 0; i < table->Columns->Count; i++) {
//        worksheet->Cells[1, i + 1]->Value = table->Columns[i]->ColumnName;
//    }
//
//    // Записываем данные
//    for (int i = 0; i < table->Rows->Count; i++) {
//        for (int j = 0; j < table->Columns->Count; j++) {
//            worksheet->Cells[i + 2, j + 1]->Value =
//                table->Rows[i][j]->ToString();
//        }
//    }
//
//    package->SaveAs(gcnew System::IO::FileInfo(filePath));
//}