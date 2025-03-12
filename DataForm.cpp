#include "DataForm.h"
#include "Chart.h"
#include <objbase.h>                // Для CoCreateGuid - генерация уникального идентификатора
#include <string>

// Добавьте эту строку для доступа к Process
using namespace System::Diagnostics;
using namespace ProjectServerW; // Добавлено пространство имен
using namespace Microsoft::Office::Interop::Excel;

std::map<std::wstring, gcroot<DataForm^>> formData_Map; // Определение переменной formData_Map
//
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
    // Важно! Установка как основного, не фонового потока. 
    // Даем потоку завершиться правильно, блокирует завершение программы до своего завершения
    excelThread->IsBackground = false; 
    excelThread->Start();

}

System::Void ProjectServerW::DataForm::buttonBrowse_Click(System::Object^ sender, System::EventArgs^ e)
{
    // Создаем диалог выбора папки
    FolderBrowserDialog^ folderDialog = gcnew FolderBrowserDialog();

    // Настройка диалога
    folderDialog->Description = "Выберите папку для сохранения Excel файлов";
    folderDialog->ShowNewFolderButton = true;

    // Устанавливаем начальную директорию из текстового поля
    if (textBoxExcelDirectory->Text->Length > 0) {
        folderDialog->SelectedPath = textBoxExcelDirectory->Text;
    }

    // Показываем диалог и проверяем результат
    if (folderDialog->ShowDialog() == System::Windows::Forms::DialogResult::OK) {
        // Обновляем текстовое поле выбранным путем
        textBoxExcelDirectory->Text = folderDialog->SelectedPath;

        // Сохраняем выбранный путь
        excelSavePath = folderDialog->SelectedPath;
        // Сохранение настроек
        SaveSettings();

    }
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
    MessageBox::Show("DataForm will be closed!");

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

// Эта функция принимает буфер и его длину, затем преобразует каждый байт буфера в шестнадцатеричный формат и добавляет его в строку.
String^ bufferToHex(const char* buffer, size_t length) {
    // Использование std::stringstream позволяет преобразовывать данные в строку и форматировать их.
    std::stringstream ss;
    ss << std::hex << std::setfill('0');
    for (int i = 0; i < length; ++i) {
        ss << std::setw(2) << static_cast<unsigned int>(static_cast<unsigned char>(buffer[i])) << " ";
    }
    return gcnew String(ss.str().c_str());
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

    for (uint8_t i = 0; i < (SQ - 1); i++)
    {
        dataTable->Columns->Add("Typ" + i, uint8_t::typeid);
        dataTable->Columns->Add("Act" + i, uint8_t::typeid);
        dataTable->Columns->Add("T" + i, uint16_t::typeid);
        dataTable->Columns->Add("H" + i, uint16_t::typeid);
    }

    // Обработка сигналов с устройства ввода-вывода
    dataTable->Columns->Add("Typ" + (SQ - 1), uint8_t::typeid);
    dataTable->Columns->Add("Act" + (SQ - 1), uint8_t::typeid);

    cli::array<cli::array<String^>^>^ bitNames = GetBitFieldNames();
    if (bitNames[0] != nullptr) {
        // Работа с массивом имен сигналов состояния дефростера
        for (uint8_t i = 0; i < 16; i++)
        {
            dataTable->Columns->Add(bitNames[0][i], uint8_t::typeid);
        }
        // Работа с массивом имен сигналов управления дефростером
        for (uint8_t i = 0; i < 16; i++)
        {
            dataTable->Columns->Add(bitNames[1][i], uint8_t::typeid);
        }
    }

    dataGridView->DataSource = dataTable;
}

// 2. Добавление данных
void ProjectServerW::DataForm::AddDataToTable(const char* buffer, size_t size, System::Data::DataTable^ table) {
    
    MSGQUEUE_OBJ_t data;
    DateTime now = DateTime::Now;   // Получение текущего времени
    String^ data_String{};

    if (size < sizeof(data)) {
        return; // Принятых данных слишком мало
    }
    memcpy(&data, buffer, sizeof(data));

    data_String = bufferToHex(buffer, size);
    // Выводим значение строки data_String в элемент Label_Data
    DataForm::SetData_TextValue("Принятые данные: " + data_String);

    // Создание новой строки
    DataRow^ row = table->NewRow();
    row["RealTime"] = now.ToString("HH:mm:ss");
    row["Time"] = data.Time;
    row["SQ"] = data.SensorQuantity;
    for (uint8_t i = 0; i < (SQ - 1); i++)
    {
        row["Typ" + i] = data.SensorType[i];
        row["Act" + i] = data.Active[i];
        row["T" + i] = data.T[i];
        row["H" + i] = data.H[i];
    }
 
    cli::array<cli::array<String^>^>^ bitNames = GetBitFieldNames();

    // Находим индекс бита "Work" в массиве bitNames[0]
    int workBitIndex = -1;
    // Найдем индекс бита "Work" в массиве имен
    for (int i = 0; i < bitNames[0]->Length; i++) {
        if (bitNames[0][i] == "Work") {
            workBitIndex = i;
            break;
        }
    }

    uint16_t bitField;
    bitField = data.T[SQ - 1];
    if (workBitIndex != -1) {
        // Проверяем текущее состояние бита "Work"
        bool currentWorkBitState = (bitField & (1 << workBitIndex)) != 0;

        // Если бит "Work" переходит из 0 в 1 (дефростер is ON)
        if (!workBitDetected && currentWorkBitState) {
            // Создаем имя файла на основе текущего времени
            excelFileName = "WorkData_" + now.ToString("yyyy-MM-dd_HH-mm-ss") + ".xlsx";
            //this->BeginInvoke(gcnew System::Action<String^>(this, &DataForm::UpdateFileNameLabel), excelFileName);
        }

        // Если бит "Work" переходит из 1 в 0 (дефростер is OFF)
        if (workBitDetected && !currentWorkBitState) {
            // Запускаем запись в Excel в отдельном потоке
            this->BeginInvoke(gcnew MethodInvoker(this, &DataForm::TriggerExcelExport));
        }

        // Обновляем состояние флага
        workBitDetected = currentWorkBitState;
    }

    // Извлекаем каждый бит и добавляем в таблицу сигналов состояния дефростера
    for (int bit = 0; bit < 16; bit++) {
        bool bitValue = (bitField & (1 << bit)) != 0;
        row[bitNames[0][bit]] = bitValue;
    }
    bitField = data.H[SQ - 1];
    // Извлекаем каждый бит и добавляем в таблицу сигналов управления дефростером
    for (int bit = 0; bit < 16; bit++) {
        bool bitValue = (bitField & (1 << bit)) != 0;
        row[bitNames[1][bit]] = bitValue;
    }

    // Добавление строки в таблицу только во время работы дефростера
    if (workBitDetected)
    {
        table->Rows->Add(row);
    }
}

// 3. Сохраняем таблицу в EXCEL
void ProjectServerW::DataForm::AddDataToExcel() {
    // Создаём объект Excel в STA-потоке
    ExcelHelper^ excel = nullptr;
    excel = gcnew ExcelHelper();

    try {
        // DataTable не является потокобезопасной структурой, в Excel будем перегонять копию
        System::Data::DataTable^ copiedTable = dataTable->Copy();
        // Последняя строка может быть не запонена полностью, её удалим
        if (copiedTable->Rows->Count > 0) {
            copiedTable->Rows->RemoveAt(copiedTable->Rows->Count - 1);
        }

        if (excel->CreateNewWorkbook()) {
            Microsoft::Office::Interop::Excel::Worksheet^ ws = excel->GetWorksheet();

            // Заголовки
            ws->Cells[1, 1] = "RealTime";
            ws->Cells[1, 2] = "Time";
            ws->Cells[1, 3] = "SQ";
            for (uint8_t i = 0; i < (SQ - 1); i++)
            {
                ws->Cells[1, 4 + 4*i] = "Typ" + i;
                ws->Cells[1, 5 + 4*i] = "Act" + i;
                ws->Cells[1, 6 + 4*i] = "T" + i;
                ws->Cells[1, 7 + 4*i] = "H" + i;
            }
            // Обработка сигналов с устройства ввода-вывода
            ws->Cells[1, 4 + 4 * (SQ - 1)] = "Typ" + (SQ - 1);
            ws->Cells[1, 5 + 4 * (SQ - 1)] = "Act" + (SQ - 1);
            uint8_t ColomnNumber = 6 + 4 * (SQ - 1);

            cli::array<cli::array<String^>^>^ bitNames = GetBitFieldNames();
            // Работа с массивом имен сигналов состояния дефростера
            if (bitNames[0] != nullptr) {
                for (uint8_t i = 0; i < 16; i++)
                {
                    ws->Cells[1, ColomnNumber++] = bitNames[0][i];
                }
            }
            // Работа с массивом имен сигналов управления дефростером
            if (bitNames[1] != nullptr) {
                for (uint8_t i = 0; i < 16; i++)
                {
                    ws->Cells[1, ColomnNumber++] = bitNames[1][i];
                }
            }

            // Данные
            int row = 2;
            for each (DataRow ^ dr in copiedTable->Rows)
            {
                ws->Cells[row, 1] = dr["RealTime"]->ToString();
                ws->Cells[row, 2] = Convert::ToInt32(dr["Time"]);
                ws->Cells[row, 3] = Convert::ToUInt16(dr["SQ"]);
                for (uint8_t i = 0; i < (SQ - 1); i++)
                {
                    ws->Cells[row, 4 + 4*i] = Convert::ToUInt16(dr["Typ" + i]);
                    ws->Cells[row, 5 + 4*i] = Convert::ToUInt16(dr["Act" + i]);
                    ws->Cells[row, 6 + 4*i] = Math::Round(Convert::ToSingle(dr["T" + i]) / 10.0, 1);
                    ws->Cells[row, 7 + 4*i] = Math::Round(Convert::ToSingle(dr["H" + i]) / 10.0, 1);
                }

                ColomnNumber = 6 + 4 * (SQ - 1);
                // Работа с массивом имен сигналов состояния дефростера
                if (bitNames[0] != nullptr) {
                    for (uint8_t i = 0; i < 16; i++)
                    {
                        ws->Cells[row, ColomnNumber++] = Convert::ToUInt16(dr[bitNames[0][i]]);
                    }
                }
                // Работа с массивом имен сигналов управления дефростером
                if (bitNames[1] != nullptr) {
                    for (uint8_t i = 0; i < 16; i++)
                    {
                        ws->Cells[row, ColomnNumber++] = Convert::ToUInt16(dr[bitNames[1][i]]);
                    }
                }

                row++;
            }

            // Сохранение файла
            // Формируем полный путь к файлу
            String^ filePath = excelSavePath;

            // Добавляем разделитель в конец пути, если его нет
            if (!filePath->EndsWith("\\")) {
                filePath += "\\";
            }

            // Добавляем имя файла с текущей датой и временем
            if (excelFileName != nullptr) {
                // Используем имя файла, созданное при обнаружении бита "Work"
                filePath += excelFileName;
            }
            else {
                // Используем стандартное имя с текущей датой и временем
                DateTime now = DateTime::Now;
                filePath += "SensorData_" + now.ToString("yyyy-MM-dd_HH-mm-ss") + ".xlsx";
            }

            // Сохранение по выбранному пути
            excel->SaveAs(filePath);

            // Сообщаем об успешном создании файла ДО закрытия COM-объектов
            Invoke(gcnew MethodInvoker(this, &DataForm::ShowSuccess));

            // Освободим память от copiedTable
            delete copiedTable;
            copiedTable = nullptr;
        }
    }
    catch (System::Runtime::InteropServices::COMException^ comEx) {
        // Специальная обработка COM-исключений
        String^ errorMsg = "COM error: " + comEx->Message + " (Code: " +
            comEx->ErrorCode.ToString("X8") + ")";
        Invoke(gcnew System::Action<String^>(this, &DataForm::ShowError), errorMsg);
    }
    catch (Exception^ ex) {
        // В случае ошибки
        String^ errorMsg = "Excel error: " + ex->Message;
        Invoke(gcnew System::Action<String^>(this, &DataForm::ShowError), errorMsg);
        // Необходимо освободить ресурсы даже при ошибке
        try {
            if (excel != nullptr) {
                excel->Close();
                delete excel;
                excel = nullptr;
            }
        }
        catch (...) { /* Игнорируем ошибки при очистке */ }
    }
    finally {
        try {
            // Важно! Сначала включить кнопку - до освобождения COM-объектов
            Invoke(gcnew MethodInvoker(this, &DataForm::EnableButton));

            // Затем небольшая пауза для обработки BeginInvoke
            Thread::Sleep(50);

            // Только после этого освобождаем Excel
            if (excel != nullptr) {
                excel->Close();
                delete excel;
                excel = nullptr;
            }

            // Запускаем отложенную сборку мусора
            ThreadPool::QueueUserWorkItem(gcnew WaitCallback(DataForm::DelayedGarbageCollection));
        }
        catch (...) {}
    }    

}

// Отложенная сборка мусора
void DataForm::DelayedGarbageCollection(Object^ state) {
    try {
        // Пауза перед сборкой мусора
        Thread::Sleep(500);

        // Сборка мусора в отдельном потоке
        GC::Collect();
        GC::WaitForPendingFinalizers();
    }
    catch (...) {}
}

void DataForm::EnableButton()
{
    buttonExcel->Enabled = true;
}

// Вспомогательные методы для UI-операций
void DataForm::ShowSuccess() {
    MessageBox::Show("Excel file was recorded successfully!");
}

void DataForm::ShowError(String^ message) {
    MessageBox::Show(message);
}

void ProjectServerW::DataForm::AddDataToTableThreadSafe(cli::array<System::Byte>^ buffer, int size) {
    // Этот метод уже выполняется в потоке формы благодаря Invoke

    // Преобразуем управляемый массив байтов в неуправляемый буфер
    pin_ptr<Byte> pinnedBuffer = &buffer[0];
    char* rawBuffer = reinterpret_cast<char*>(pinnedBuffer);

    // Вызываем стандартный метод добавления данных
    AddDataToTable(rawBuffer, size, dataTable);

    // Автоматически обновляем форму после добавления данных
    if (dataGridView->RowCount > 0) {
        // Прокрутка к последней строке
        dataGridView->FirstDisplayedScrollingRowIndex = dataGridView->RowCount - 1;
    }

    // Обновляем UI (это безопасно, т.к. мы уже в потоке формы)
    this->Refresh();
}

/*
********************** методы для сохранения и загрузки настроек **********************
*/

void ProjectServerW::DataForm::SaveSettings() {
    try {
        // Создаем или открываем файл настроек
        // Получаем директорию, где находится исполняемый файл
        String^ appPath = System::IO::Path::GetDirectoryName(System::Windows::Forms::Application::ExecutablePath);
        String^ settingsPath = System::IO::Path::Combine(appPath, "ExcelSettings.txt");

        System::IO::StreamWriter^ writer = gcnew System::IO::StreamWriter(settingsPath);
        // Записываем путь
        writer->WriteLine(textBoxExcelDirectory->Text);

        // Закрываем файл
        writer->Close();
    }
    catch (Exception^ ex) {
        // Обработка ошибок
        MessageBox::Show("Не удалось сохранить настройки: " + ex->Message);
    }
}

void ProjectServerW::DataForm::LoadSettings() {
    try {
        // Проверяем существование файла настроек
        String^ appPath = System::IO::Path::GetDirectoryName(System::Windows::Forms::Application::ExecutablePath);
        String^ settingsPath = System::IO::Path::Combine(appPath, "ExcelSettings.txt");
        if (System::IO::File::Exists(settingsPath)) {
            // Открываем и читаем файл
            System::IO::StreamReader^ reader = gcnew System::IO::StreamReader("ExcelSettings.txt");

            // Читаем первую строку (путь)
            String^ path = reader->ReadLine();

            // Закрываем файл
            reader->Close();

            // Обновляем текстовое поле и переменную пути
            if (path != nullptr && path->Length > 0) {
                textBoxExcelDirectory->Text = path;
                excelSavePath = path;
            }
        }
    }
    catch (Exception^ ex) {
        // Обработка ошибок
        MessageBox::Show("Не удалось загрузить настройки: " + ex->Message);
    }
}

//*******************************************************************************************
// Реализация метода инициализации названий битовых полей
void ProjectServerW::DataForm::InitializeBitFieldNames(gcroot<cli::array<cli::array<String^>^>^>& namesRef) {
    namesRef = gcnew cli::array<cli::array<String^>^>(10);

    // Инициализация массива состояния устройств дефростера
    namesRef[0] = gcnew cli::array<String^>(16) {
        "Vent0", "Vent1", "Vent2", "Vent3", // Работа циркуляционного вентилятора 1..4
        "Heat0", "Heat1", "Heat2", "Heat3", // Работа нагревателя  (ТЭНа) 1..4
        "OutA",     // Работа вытяжного вентилятора
        "InjW",     // Работа водяных форсунок
        "Work",     // Лампа РАБОТА 
        "Alrm",     // Лампа АВАРИЯ 
        "Open",     // Сигнал «Ворота открыты»
        "Clse",     // Сигнал «Ворота закрыты»
        "StUP",     // Сигнал «Промежуточное положение ворот при движении вверх»
        "StDN"      // Сигнал «Промежуточное положение ворот при движении вниз»
    };

    // Инициализация массива cигналов управления устройствами дефростера
    namesRef[1] = gcnew cli::array<String^>(16) {
        "_V0", "_V1", "_V2", "_V3", // Циркуляционный вентилятор 1..4 включить
        "_H0", "_H1", "_H2", "_H3", // Нагреватель (ТЭН) 1..4 включить
        "_Out",     // Вытяжной вентилятор включить
        "_Inj",     // Водяную форсунку включить
        "_Flp",     // Закрыть защитную заслонку вытяжного вентилятора 
        "_Opn",     // Ворота открыть
        "_Stp",     // Ворота остановить
        "_Cls",     // Ворота закрыть
        "_Snd",     // Звуковой сигнал включить
        "_Wrk"      // Включить зелёную лампу РАБОТА
    };

    // Добавьте другие типы сенсоров по мере необходимости
}

////******************** Обработка завершения формы ***************************************
//System::Void ProjectServerW::DataForm::DataForm_FormClosing(System::Object^ sender, System::Windows::Forms::FormClosingEventArgs^ e) {
//    try {
//        // Проверяем, есть ли данные в таблице
//        if (dataTable != nullptr && dataTable->Rows->Count > 0) {
//            // Сначала отключаем обработчик события FormClosing, чтобы предотвратить повторный вызов
//            this->FormClosing -= gcnew System::Windows::Forms::FormClosingEventHandler(this, &DataForm::DataForm_FormClosing);
//            // Блокируем закрытие формы на время сохранения
//            e->Cancel = true;
//
//            // Создаем имя файла, если оно еще не создано
//            if (excelFileName == nullptr) {
//                DateTime now = DateTime::Now;
//                excelFileName = "CloseData_" + now.ToString("yyyy-MM-dd_HH-mm-ss") + ".xlsx";
//            }
//
//            // Обновляем метку
//            if (Label_Data != nullptr) {
//                Label_Data->Text = "Сохранение данных в Excel...";
//                Label_Data->Refresh();
//            }
//
//            // Создаем и запускаем поток с STA-моделью для работы с Excel
//            Thread^ closeExcelThread = gcnew Thread(gcnew ThreadStart(this, &DataForm::AddDataToExcel));
//            closeExcelThread->SetApartmentState(ApartmentState::STA);
//            closeExcelThread->IsBackground = false;
//
//            // Добавляем метод, который будет вызван после завершения потока
//            this->BeginInvoke(gcnew MethodInvoker(this, &DataForm::CloseFormAfterExport));
//
//            closeExcelThread->Start();
//        }
//    }
//    catch (Exception^ ex) {
//        MessageBox::Show("Ошибка при попытке сохранения данных: " + ex->Message);
//        // В случае ошибки разрешаем закрытие формы
//        e->Cancel = false;
//    }
//}
//
//// метод для экспорта и закрытия формы
//void ProjectServerW::DataForm::ExportAndClose() {
//    try {
//        // Вызываем метод экспорта напрямую, не через кнопку
//        AddDataToExcel();
//
//        // После завершения экспорта закрываем форму из потока UI
//        this->BeginInvoke(gcnew MethodInvoker(this, &DataForm::CloseFormAfterExport));
//    }
//    catch (Exception^ ex) {
//        this->BeginInvoke(gcnew System::Action<String^>(this, &DataForm::ShowError),
//            "Ошибка при экспорте: " + ex->Message);
//
//        // В случае ошибки все равно закрываем форму
//        this->BeginInvoke(gcnew MethodInvoker(this, &DataForm::CloseFormAfterExport));
//    }
//}

//// Метод для закрытия формы после экспорта
//void ProjectServerW::DataForm::CloseFormAfterExport() {
//    // Проверяем, завершился ли поток экспорта в Excel
//    // Подождем несколько секунд для завершения операции
//    int timeoutSeconds = 5;
//    DateTime startTime = DateTime::Now;
//
//    while (DateTime::Now.Subtract(startTime).TotalSeconds < timeoutSeconds) {
//        if (buttonExcel->Enabled) {
//            // Кнопка Excel включена обратно, значит экспорт завершен
//            break;
//        }
//        Thread::Sleep(100);
//    }
//
//    // Отключаем обработчик закрытия формы, чтобы избежать повторного вызова
//    this->FormClosing -= gcnew System::Windows::Forms::FormClosingEventHandler(this, &DataForm::DataForm_FormClosing);
//
//    // Закрываем форму
//    this->Close();
//}
//// Метод для безопасного вызова BeginInvoke
//void ProjectServerW::DataForm::SafeBeginInvoke(MethodInvoker^ method) {
//    try {
//        // Проверяем, что форма открыта и создана
//        if (!this->IsDisposed && this->IsHandleCreated && !this->Disposing) {
//            this->BeginInvoke(method);
//        }
//    }
//    catch (...) {
//        // Игнорируем любые ошибки
//    }
//}

//// Перегрузка для методов с параметром
//void ProjectServerW::DataForm::SafeBeginInvoke(System::Action<String^>^ method, String^ param) {
//    try {
//        // Проверяем, что форма открыта и создана
//        if (!this->IsDisposed && this->IsHandleCreated && !this->Disposing) {
//            this->BeginInvoke(method, param);
//        }
//    }
//    catch (...) {
//        // Игнорируем любые ошибки
//    }
//}
