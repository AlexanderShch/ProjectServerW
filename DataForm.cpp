#include "DataForm.h"
#include "Chart.h"
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
    ProjectServerW::DataForm::Close();
    System::Windows::Forms::Application::Exit();
	return System::Void();
}

System::Void ProjectServerW::DataForm::buttonEXCEL_Click(System::Object^ sender, System::EventArgs^ e)
{
    // Отключаем кнопку на время обработки
    buttonExcel->Enabled = false;

    // Создаем и запускаем поток
    excelThread = gcnew Thread(gcnew ThreadStart(this, &DataForm::AddDataToExcel));
    excelThread->SetApartmentState(ApartmentState::STA);  // Обязательно для Excel!
    excelThread->IsBackground = true;
    excelThread->Start();
}

void ProjectServerW::DataForm::CreateAndShowDataFormInThread(std::queue<std::wstring>& messageQueue,
                                                             std::mutex& mtx, 
                                                             std::condition_variable& cv) {
    DataForm^ form = gcnew DataForm();
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
            form->Invoke(gcnew System::Action(form, &ProjectServerW::DataForm::Close));
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
    dataTable = gcnew System::Data::DataTable("SensorData");
 
    dataTable->Columns->Add("RealTime", String::typeid);
    dataTable->Columns->Add("Time", uint16_t::typeid);
    dataTable->Columns->Add("SQ", uint8_t::typeid);

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
void ProjectServerW::DataForm::AddDataToTable(const char* buffer, size_t size, System::Data::DataTable^ table) {
    MSGQUEUE_OBJ_t data;
    DateTime now = DateTime::Now;   // Получение текущего времени

    if (size < sizeof(data)) {
        return; // Принятых данных слишком мало
    }
    memcpy(&data, buffer, sizeof(data));

    // Создание новой строки
    DataRow^ row = table->NewRow();
    row["RealTime"] = now.ToString("HH:mm:ss");
    row["Time"] = data.Time;
    row["SQ"] = data.SensorQuantity;
    for (uint8_t i = 0; i < SQ; i++)
    {
        row["Typ" + i] = data.SensorType[i];
        row["Act" + i] = data.Active[i];
        row["T" + i] = data.T[i];
        row["H" + i] = data.H[i];
    }

    // Добавление строки в таблицу
    table->Rows->Add(row);
}

// 3. Сохраняем таблицу в EXCEL
void ProjectServerW::DataForm::AddDataToExcel() {
    try {
        // DataTable не является потокобезопасной структурой, в Excel будем перегонять копию
        System::Data::DataTable^ copiedTable = dataTable->Copy();
        // Последняя строка может быть не запонена полностью, её удалим
        if (copiedTable->Rows->Count > 0) {
            copiedTable->Rows->RemoveAt(copiedTable->Rows->Count - 1);
        }

        ExcelHelper^ excel = gcnew ExcelHelper();
        if (excel->CreateNewWorkbook()) {
            Microsoft::Office::Interop::Excel::Worksheet^ ws = excel->GetWorksheet();

            // Заголовки
            ws->Cells[1, 1] = "RealTime";
            ws->Cells[1, 2] = "Time";
            ws->Cells[1, 3] = "SQ";
            for (uint8_t i = 0; i < SQ; i++)
            {
                ws->Cells[1, 4 + 4*i] = "Typ" + i;
                ws->Cells[1, 5 + 4*i] = "Act" + i;
                ws->Cells[1, 6 + 4*i] = "T" + i;
                ws->Cells[1, 7 + 4*i] = "H" + i;
            }

            // Данные
            int row = 2;
            for each (DataRow ^ dr in copiedTable->Rows)
            {
                ws->Cells[row, 1] = dr["RealTime"]->ToString();
                ws->Cells[row, 2] = Convert::ToInt32(dr["Time"]);
                ws->Cells[row, 3] = Convert::ToUInt16(dr["SQ"]);
                for (uint8_t i = 0; i < SQ; i++)
                {
                    ws->Cells[row, 4 + 4*i] = Convert::ToUInt16(dr["Typ" + i]);
                    ws->Cells[row, 5 + 4*i] = Convert::ToUInt16(dr["Act" + i]);
                    ws->Cells[row, 6 + 4*i] = Convert::ToInt32(dr["T" + i]);
                    ws->Cells[row, 7 + 4*i] = Convert::ToInt32(dr["H" + i]);
                }
                row++;
            }

            // Сохранение
            excel->SaveAs("D:\\SensorData.xlsx");
            excel->Close();
            // Освободим память от копии таблицы
            delete copiedTable;
            // Правильное освобождение COM-объектов
            if (ws != nullptr) {
                Marshal::ReleaseComObject(ws);
                ws = nullptr;
            }
        }
    }
    catch (Exception^ ex) {
        MessageBox::Show("Excel error: " + ex->Message);
    }
    finally
    {
        CoUninitialize();
        //GC::WaitForPendingFinalizers();
        //for (int i = 0; i < 3; i++) {
            //GC::Collect();
        GC::SuppressFinalize(this);
        Thread::Sleep(100);  // Короткая пауза
        //}
        // Включаем кнопку обратно
        this->BeginInvoke(gcnew MethodInvoker(this, &DataForm::EnableButton));
        // Сборка мусора только при необходимости
        GC::Collect();
    }
}

void DataForm::EnableButton()
{
    buttonExcel->Enabled = true;
}