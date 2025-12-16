#include "DataForm.h"
#include "Chart.h"
#include "Commands.h"               // Для работы с командами управления
#include "PacketQueueProcessor.h"   // Shared per-socket send gate (telemetry ACKs vs UI commands)
#include <objbase.h>                // Для CoCreateGuid - генерация уникального идентификатора
#include <string>
#include <vcclr.h>  // Для gcnew
#include <msclr/marshal_cppstd.h>

// Добавьте эту строку для доступа к Process
using namespace System::Diagnostics;
using namespace ProjectServerW; // Добавлено пространство имен
using namespace Microsoft::Office::Interop::Excel;

std::map<std::wstring, gcroot<DataForm^>> formData_Map; // Определение переменной formData_Map

// Структура телеметрии от контроллера (упакованная, без выравнивания байтов)
// Контроллер использует __attribute__((packed)), здесь используем #pragma pack(1)
#pragma pack(push, 1)
typedef struct   // object data for Server type из STM32
{
    uint8_t DataType;           // Байт типа передаваемых данных (0х00 для телеметрии)
    uint16_t Time;				// Количество секунд с момента включения
    uint8_t SensorQuantity;		// Количество сенсоров
    uint8_t SensorType[SQ];		// Тип сенсора
    uint8_t Active[SQ];			// Активность сенсора
    short T[SQ];				// Значение 1 сенсора (температура)
    short H[SQ];				// Значение 2 сенсора (влажность)
    uint16_t CRC_SUM;			// Контрольное значение
} MSGQUEUE_OBJ_t;
#pragma pack(pop)

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
    // Найдем GUID текущей формы в карте formData_Map
    std::wstring currentFormGuid;
    bool formFound = false;

    for (auto it = formData_Map.begin(); it != formData_Map.end(); ++it) {
        // Извлекаем управляемый указатель из gcroot
        ProjectServerW::DataForm^ formPtr = it->second;

        // Теперь сравниваем указатели
        if (formPtr == this) {
            currentFormGuid = it->first;
            formFound = true;
            break;
        }
    }

    try {
        if (formFound) {
            // Закрываем соединение с клиентом
            // Находим форму по GUID
            DataForm^ form = GetFormByGuid(currentFormGuid);

            if (form != nullptr && form->ClientSocket != INVALID_SOCKET) {
                // Закрываем сокет клиента
                closesocket(form->ClientSocket);
                form->ClientSocket = INVALID_SOCKET;
                GlobalLogger::LogMessage("Information: Закрыто соединение с клиентом по кнопке Выход");
            }
        }
    }
    catch (Exception^ ex) {
        MessageBox::Show("Ошибка при закрытии сокета: " + ex->Message);
        GlobalLogger::LogMessage(ConvertToStdString("Ошибка при закрытии сокета: " + ex->Message));
    }

	return System::Void();
}

System::Void ProjectServerW::DataForm::buttonEXCEL_Click(System::Object^ sender, System::EventArgs^ e)
{
    try {
        StartExcelExportThread(false);
    }
    catch (Exception^ ex) {
        GlobalLogger::LogMessage(ConvertToStdString("Error: Exception in buttonEXCEL_Click: " + ex->ToString()));
    }
}

bool ProjectServerW::DataForm::StartExcelExportThread(bool isEmergency) {
    try {
        // Critical: prevent duplicate enqueue per form (UI can trigger export from multiple places).
        System::Threading::Monitor::Enter(excelExportSync);
        try {
            if (excelExportInProgress != 0) {
                GlobalLogger::LogMessage("Information: Excel export already pending/in progress (per-form guard)");
                return true;
            }
            excelExportInProgress = 1;
        }
        finally {
            System::Threading::Monitor::Exit(excelExportSync);
        }

        EnsureExcelExportWorker();

        // Disable the button only for non-emergency exports to avoid touching disposed UI during shutdown.
        if (!isEmergency) {
            try {
                if (buttonExcel != nullptr && !buttonExcel->IsDisposed) {
                    buttonExcel->Enabled = false;
                }
            }
            catch (...) {}
        }

        System::Data::DataTable^ snapshot = nullptr;
        System::Threading::Monitor::Enter(dataTableSync);
        try {
            snapshot = dataTable->Copy();
        }
        finally {
            System::Threading::Monitor::Exit(dataTableSync);
        }

        if (snapshot != nullptr && snapshot->Rows->Count > 0) {
            snapshot->Rows->RemoveAt(snapshot->Rows->Count - 1);
        }

        ExcelExportJob^ job = gcnew ExcelExportJob();
        job->tableSnapshot = snapshot;
        job->saveDirectory = excelSavePath;
        job->sessionStart = dataCollectionStartTime;
        job->sessionEnd = dataCollectionEndTime;
        job->clientPort = clientPort;
        job->clientIP = this->ClientIP;
        job->formGuid = this->FormGuid;
        job->formRef = gcnew System::WeakReference(this);
        job->enableButtonOnComplete = !isEmergency;

        excelExportQueue->Enqueue(job);
        excelExportQueueEvent->Set();

        if (isEmergency) {
            GlobalLogger::LogMessage("Information: Excel export queued (emergency background)");
        }
        else {
            GlobalLogger::LogMessage("Information: Excel export queued");
        }

        return true;
    }
    catch (Exception^ ex) {
        System::Threading::Monitor::Enter(excelExportSync);
        try {
            excelExportInProgress = 0;
        }
        finally {
            System::Threading::Monitor::Exit(excelExportSync);
        }
        GlobalLogger::LogMessage(ConvertToStdString("Error: Failed to start Excel export thread: " + ex->ToString()));
        return false;
    }
}

void ProjectServerW::DataForm::EnsureExcelExportWorker() {
    System::Threading::Monitor::Enter(excelExportWorkerSync);
    try {
        if (excelExportWorkerStarted && excelExportWorkerThread != nullptr) {
            return;
        }

        excelExportWorkerThread = gcnew System::Threading::Thread(gcnew System::Threading::ThreadStart(&ProjectServerW::DataForm::ExcelExportWorkerLoop));
        excelExportWorkerThread->IsBackground = true;
        excelExportWorkerThread->SetApartmentState(System::Threading::ApartmentState::STA);
        excelExportWorkerThread->Start();
        excelExportWorkerStarted = true;
    }
    finally {
        System::Threading::Monitor::Exit(excelExportWorkerSync);
    }
}

void ProjectServerW::DataForm::ExcelExportWorkerLoop() {
    while (true) {
        excelExportQueueEvent->WaitOne();

        ExcelExportJob^ job = nullptr;
        while (excelExportQueue->TryDequeue(job)) {
            ProcessExcelExportJob(job);
        }
    }
}

System::String^ ProjectServerW::DataForm::ResolveExcelSaveDirectory(System::String^ preferredDirectory) {
    try {
        if (!String::IsNullOrEmpty(preferredDirectory) && System::IO::Directory::Exists(preferredDirectory)) {
            return preferredDirectory;
        }
    }
    catch (...) {}

    System::String^ appPath = System::IO::Path::GetDirectoryName(System::Windows::Forms::Application::ExecutablePath);
    System::String^ dir = System::IO::Path::Combine(appPath, "SensorData");
    if (!System::IO::Directory::Exists(dir)) {
        System::IO::Directory::CreateDirectory(dir);
    }
    return dir;
}

void ProjectServerW::DataForm::ProcessExcelExportJob(ExcelExportJob^ job) {
    if (job == nullptr || job->tableSnapshot == nullptr) {
        return;
    }

    bool mutexAcquired = false;
    try {
        const int timeoutMs = 5 * 60 * 1000; // 5 minutes
        try {
            mutexAcquired = excelGlobalMutex->WaitOne(timeoutMs);
        }
        catch (System::Threading::AbandonedMutexException^) {
            mutexAcquired = true;
            GlobalLogger::LogMessage("Warning: Abandoned Excel mutex detected; continuing export.");
        }

        if (!mutexAcquired) {
            // Critical: keep retrying via the queue; Excel can be busy for long exports.
            excelExportQueue->Enqueue(job);
            excelExportQueueEvent->Set();
            System::Threading::Thread::Sleep(1000);
            return;
        }

        ExcelHelper^ excel = gcnew ExcelHelper();
        if (!excel->CreateNewWorkbook()) {
            return;
        }

        Microsoft::Office::Interop::Excel::Worksheet^ ws = excel->GetWorksheet();
        Microsoft::Office::Interop::Excel::Application^ excelApp = safe_cast<Microsoft::Office::Interop::Excel::Application^>(ws->Application);

        try {
            excelApp->ScreenUpdating = false;
            excelApp->Calculation = Microsoft::Office::Interop::Excel::XlCalculation::xlCalculationManual;
            excelApp->EnableEvents = false;

            // Info sheet (metadata)
            try {
                Microsoft::Office::Interop::Excel::Workbook^ wb = safe_cast<Microsoft::Office::Interop::Excel::Workbook^>(ws->Parent);
                System::Object^ missing = System::Type::Missing;
                Microsoft::Office::Interop::Excel::Worksheet^ infoSheet =
                    safe_cast<Microsoft::Office::Interop::Excel::Worksheet^>(wb->Worksheets->Add(missing, ws, 1, Microsoft::Office::Interop::Excel::XlSheetType::xlWorksheet));
                infoSheet->Name = "Info";

                DateTime sessionStart = (job->sessionStart == DateTime::MinValue) ? DateTime::Now : job->sessionStart;
                DateTime sessionEnd = (job->sessionEnd == DateTime::MinValue) ? DateTime::Now : job->sessionEnd;

                infoSheet->Cells[1, 1] = "SessionStart";
                infoSheet->Cells[1, 2] = sessionStart.ToString("yyyy-MM-dd HH:mm:ss");
                infoSheet->Cells[2, 1] = "SessionEnd";
                infoSheet->Cells[2, 2] = sessionEnd.ToString("yyyy-MM-dd HH:mm:ss");
                infoSheet->Cells[3, 1] = "ClientPort";
                infoSheet->Cells[3, 2] = job->clientPort.ToString();
                infoSheet->Cells[4, 1] = "ClientIP";
                infoSheet->Cells[4, 2] = (job->clientIP != nullptr ? job->clientIP : "");
                infoSheet->Cells[5, 1] = "FormGuid";
                infoSheet->Cells[5, 2] = (job->formGuid != nullptr ? job->formGuid : "");

                Marshal::ReleaseComObject(infoSheet);
                Marshal::ReleaseComObject(wb);
            }
            catch (...) {}

            int colCount = job->tableSnapshot->Columns->Count;
            int rowCount = job->tableSnapshot->Rows->Count;

            cli::array<System::Object^>^ headerArray = gcnew cli::array<System::Object^>(colCount);
            for (int c = 0; c < colCount; c++) {
                headerArray[c] = job->tableSnapshot->Columns[c]->ColumnName;
            }
            Microsoft::Office::Interop::Excel::Range^ headerRange = ws->Range[ws->Cells[1, 1], ws->Cells[1, colCount]];
            headerRange->Value2 = headerArray;
            Marshal::ReleaseComObject(headerRange);

            if (rowCount > 0) {
                cli::array<System::Object^, 2>^ dataArray = gcnew cli::array<System::Object^, 2>(rowCount, colCount);
                for (int r = 0; r < rowCount; r++) {
                    System::Data::DataRow^ dr = job->tableSnapshot->Rows[r];
                    for (int c = 0; c < colCount; c++) {
                        dataArray[r, c] = dr[c];
                    }
                }

                Microsoft::Office::Interop::Excel::Range^ startCell = safe_cast<Microsoft::Office::Interop::Excel::Range^>(ws->Cells[2, 1]);
                Microsoft::Office::Interop::Excel::Range^ endCell = safe_cast<Microsoft::Office::Interop::Excel::Range^>(ws->Cells[rowCount + 1, colCount]);
                Microsoft::Office::Interop::Excel::Range^ dataRange = ws->Range[startCell, endCell];
                dataRange->Value2 = dataArray;
                Marshal::ReleaseComObject(startCell);
                Marshal::ReleaseComObject(endCell);
                Marshal::ReleaseComObject(dataRange);
            }
        }
        finally {
            try {
                excelApp->Calculation = Microsoft::Office::Interop::Excel::XlCalculation::xlCalculationAutomatic;
                excelApp->ScreenUpdating = true;
                excelApp->EnableEvents = true;
            }
            catch (...) {}
            Marshal::ReleaseComObject(excelApp);
        }

        System::String^ dir = ResolveExcelSaveDirectory(job->saveDirectory);
        if (!dir->EndsWith("\\")) {
            dir += "\\";
        }

        DateTime start = (job->sessionStart == DateTime::MinValue) ? DateTime::Now : job->sessionStart;
        DateTime end = (job->sessionEnd == DateTime::MinValue) ? DateTime::Now : job->sessionEnd;
        System::String^ finalFileName = String::Format(
            "WorkData_Start_{0}_End_{1}_Port{2}.xlsx",
            start.ToString("yyyy-MM-dd_HH-mm-ss"),
            end.ToString("yyyy-MM-dd_HH-mm-ss"),
            job->clientPort.ToString());

        excel->SaveAs(dir + finalFileName);
        excel->Close();
        delete excel;
        excel = nullptr;

        GlobalLogger::LogMessage(ConvertToStdString("Information: Excel file saved: " + finalFileName));
    }
    catch (Exception^ ex) {
        GlobalLogger::LogMessage(ConvertToStdString("Error: Excel export job failed: " + ex->ToString()));
    }
    finally {
        if (mutexAcquired) {
            try { excelGlobalMutex->ReleaseMutex(); } catch (...) {}
        }

        try {
            DataForm^ form = nullptr;
            if (job != nullptr && job->formRef != nullptr && job->formRef->IsAlive) {
                form = dynamic_cast<DataForm^>(job->formRef->Target);
            }

            if (form != nullptr && !form->IsDisposed && !form->Disposing) {
                System::Threading::Monitor::Enter(form->excelExportSync);
                try {
                    form->excelExportInProgress = 0;
                }
                finally {
                    System::Threading::Monitor::Exit(form->excelExportSync);
                }

                if (job->enableButtonOnComplete) {
                    try {
                        form->BeginInvoke(gcnew MethodInvoker(form, &DataForm::EnableButton));
                    }
                    catch (...) {}
                }
            }
        }
        catch (...) {}
    }
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
        form->FormGuid = formId;

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

// Метод закрытия формы, применяется при закрытии сокета
void ProjectServerW::DataForm::CloseForm(const std::wstring& guid) {
    // Находим форму
    ProjectServerW::DataForm^ form = ProjectServerW::DataForm::GetFormByGuid(guid);
    //MessageBox::Show("DataForm will be closed!");
    GlobalLogger::LogMessage("DataForm will be closed!");

    if (form != nullptr) {
        // ВАЖНО: Удаляем форму из карты СРАЗУ, до закрытия
        // Это позволяет новым соединениям от того же клиента создавать новые формы немедленно
        auto it = formData_Map.find(guid);
        if (it != formData_Map.end()) {
            formData_Map.erase(it);
            GlobalLogger::LogMessage("Information: Форма удалена из карты, разрешено новое соединение");
        }
        
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
        DataForm^ form = it->second;

        // Проверяем, что форма существует и не закрыта
        if (form != nullptr && !form->IsDisposed && !form->Disposing) {
            return form;
        }
        else {
            // Форма закрыта или недействительна, удаляем её из карты
            formData_Map.erase(it);
            return nullptr;
        }

    }
    return nullptr;
}

// Метод для поиска формы по IP-адресу клиента
std::wstring ProjectServerW::DataForm::FindFormByClientIP(String^ clientIP) {
    // Проходим по всем формам в карте
    for (auto it = formData_Map.begin(); it != formData_Map.end(); ++it) {
        DataForm^ form = it->second;
        
        // Проверяем, что форма существует и не закрыта
        if (form != nullptr && !form->IsDisposed && !form->Disposing) {
            // Сравниваем IP-адреса
            if (form->ClientIP != nullptr && form->ClientIP->Equals(clientIP)) {
                // Нашли форму с таким же IP
                return it->first; // Возвращаем GUID
            }
        }
    }
    // Не нашли форму с таким IP
    return std::wstring();
}

// Записываем пару guid и поток в map
void ThreadStorage::StoreThread(const std::wstring& guid, std::thread& thread) {
    std::lock_guard<std::mutex> lock(GetMutex());
    GetThreadMap()[guid] = std::move(thread);
}

// Останавливаем поток по guid
void ThreadStorage::StopThread(const std::wstring& guid)
{
    std::thread threadToStop;
    {
        std::lock_guard<std::mutex> lock(GetMutex());

        // Move the thread out of the map first. Holding the mutex while waiting can deadlock shutdown paths.
        auto it = GetThreadMap().find(guid);
        if (it == GetThreadMap().end()) {
            return;
        }

        threadToStop = std::move(it->second);
        GetThreadMap().erase(it);
    }

    if (!threadToStop.joinable()) {
        return;
    }

    // Closing forms must never block on thread joins. In this app a DataForm is hosted on its own UI thread
    // (ShowDialog inside std::thread). Joining from any UI-related close path can deadlock the modal loop.
    threadToStop.detach();
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
static String^ bufferToHex(const char* buffer, size_t length) {
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
        dataTable->Columns->Add("T" + i, float::typeid);
        dataTable->Columns->Add("H" + i, float::typeid);
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
    GlobalLogger::LogMessage("Information: Таблица данных создана");
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
    // Why: used to convert device-relative time in command audit replies into an approximate wall-clock timestamp.
    lastTelemetryDeviceSeconds = data.Time;
    row["SQ"] = data.SensorQuantity;
    for (uint8_t i = 0; i < (SQ - 1); i++)
    {
        row["Typ" + i] = data.SensorType[i];
        row["Act" + i] = data.Active[i];
        row["T" + i] = data.T[i] / 10.0;
        row["H" + i] = data.H[i] / 10.0;
    }
    // Обновляем значения температуры в элементах интерфейса
    cli::array<double>^ temperatures = gcnew cli::array<double>(5);
    temperatures[0] = data.T[0] / 10.0;  // T_def_left
    temperatures[1] = data.T[1] / 10.0;  // T_def_right
    temperatures[2] = data.T[2] / 10.0;  // T_def_center
    temperatures[3] = data.T[3] / 10.0;  // T_product_left
    temperatures[4] = data.T[4] / 10.0;  // T_product_right
    UpdateAllTemperatureValues(temperatures);

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

        if (reconnectFixationLogPending) {
            reconnectFixationLogPending = false;
            GlobalLogger::LogMessage(ConvertToStdString(String::Format(
                "Information: Начало фиксации данных после повторного подключения клиента {0} (Port {1})",
                (this->ClientIP != nullptr ? this->ClientIP : ""),
                clientPort)));
        }

        // ===== УСТАНОВКА НАЧАЛЬНОГО СОСТОЯНИЯ КНОПОК ПРИ ПЕРВОМ ПРИЁМЕ ДАННЫХ =====
        if (!firstDataReceived) {
            // Это первый пакет данных - устанавливаем начальное состояние кнопок
            if (currentWorkBitState) {
                // Бит WORK активен - работа уже идёт
                buttonSTOPstate_TRUE();
                GlobalLogger::LogMessage("Information: Первый пакет данных - бит WORK активен, кнопка STOP активна");
            } else {
                // Бит WORK не активен - можно запустить
                buttonSTARTstate_TRUE();
                GlobalLogger::LogMessage("Information: Первый пакет данных - бит WORK не активен, кнопка START активна");
            }
            firstDataReceived = true;
            
            // Автоматически запрашиваем версию прошивки контроллера после первого пакета
            SendVersionRequest();
        }

        // Если бит "Work" переходит из 0 в 1 (дефростер is ON)
        if (!workBitDetected && currentWorkBitState) {
            // Проверяем: это возврат после мигания или новый запуск?
            if (workBitZeroTimerActive) {
                // Это мигание лампы - бит вернулся в 1 во время отслеживания
                // НЕ логируем, только сбрасываем таймер отслеживания
                workBitZeroTimerActive = false;
                // НЕ сбрасываем workBitZeroLogged - чтобы не логировать повторно при следующем мигании
                if (delayedExcelTimer != nullptr && delayedExcelTimer->Enabled) {
                    delayedExcelTimer->Stop();
                }
            } else {
                // Это новый запуск работы (не мигание)
                // Store the session start time. The final filename will include both start and finish timestamps.
                dataCollectionStartTime = now;
                dataCollectionEndTime = DateTime::MinValue;
                excelFileName = "WorkData_Port" + clientPort.ToString();
                GlobalLogger::LogMessage(ConvertToStdString(String::Format(
                    "Information: СТАРТ фиксации данных: {0} (Port {1})",
                    dataCollectionStartTime.ToString("yyyy-MM-dd HH:mm:ss"), clientPort)));
                // Сбрасываем флаги при новом запуске работы
                workBitZeroLogged = false;
                dataExportedToExcel = false;  // Новый цикл - данные еще не экспортированы
            }
            // Меняем состояние кнопок СТАРТ и СТОП
            buttonSTOPstate_TRUE();
        }

        // Если бит "Work" переходит из 1 в 0 (дефростер is OFF)
        if (workBitDetected && !currentWorkBitState) {
            // Начинаем отслеживать время, когда бит стал нулём
            workBitZeroStartTime = now;
            workBitZeroTimerActive = true;
            
            // Создаем и запускаем таймер для проверки состояния каждую секунду
            if (delayedExcelTimer == nullptr) {
                delayedExcelTimer = gcnew System::Windows::Forms::Timer();
                delayedExcelTimer->Interval = 1000; // Проверяем каждую секунду
                delayedExcelTimer->Tick += gcnew EventHandler(this, &DataForm::OnDelayedExcelTimerTick);
            }
            delayedExcelTimer->Start();
            
            // Логируем только первый переход в ноль
            if (!workBitZeroLogged) {
                GlobalLogger::LogMessage(ConvertToStdString("Information: Бит Work стал нулём, начато отслеживание времени..."));
                workBitZeroLogged = true;
            }
        }

        // Cancel auto-restart only by a large safety timeout.
        // Why: after STOP the device may "blink" Work (1<->0) for minutes; readiness is detected by
        // "no blinking for 60s" (same logic that enables the START button), not by Work being 1 or 0.
        if (autoRestartPending && autoRestartStopIssuedTime != DateTime::MinValue) {
            TimeSpan waited = now.Subtract(autoRestartStopIssuedTime);
            const double safetyTimeoutMinutes = 5;
            if (waited.TotalMinutes >= safetyTimeoutMinutes) {
                autoRestartPending = false;
                autoRestartStopIssuedTime = DateTime::MinValue;
                GlobalLogger::LogMessage(ConvertToStdString(String::Format(
                    "Warning: Автоперезапуск отменён: не дождались остановки (нет 60с без моргания Work) за {0} минут после STOP",
                    safetyTimeoutMinutes)));
                if (Label_Commands != nullptr && !Label_Commands->IsDisposed) {
                    Label_Commands->Text = "[!] Автоперезапуск отменён: таймаут ожидания остановки";
                    Label_Commands->ForeColor = System::Drawing::Color::Orange;
                }
            }
        }

        // Если бит "Work" остаётся в нуле и таймер активен, проверяем, прошла ли минута
        if (!currentWorkBitState && workBitZeroTimerActive) {
            TimeSpan elapsed = now.Subtract(workBitZeroStartTime);
            if (elapsed.TotalSeconds >= 60) {
                // Прошло не менее 1 минуты непрерывного нахождения в нуле
                workBitZeroTimerActive = false;
                workBitZeroLogged = false;  // Сбрасываем флаг логирования
                delayedExcelTimer->Stop();
                
                // Сохраняем время окончания сбора данных
                dataCollectionEndTime = now;
                
                GlobalLogger::LogMessage(ConvertToStdString("Information: СТОП фиксации данных, запись в файл " + excelFileName));
                // Меняем состояние кнопок СТАРТ и СТОП
                buttonSTARTstate_TRUE();
                
                // Устанавливаем флаг, что данные экспортированы (чтобы не дублировать при закрытии формы)
                dataExportedToExcel = true;
                
                // Выполняем запись в Excel
                this->BeginInvoke(gcnew MethodInvoker(this, &DataForm::TriggerExcelExport));

                // Why: START/STOP commands are synchronous (wait for controller response). Queue START to the UI
                // message loop to keep telemetry processing short and to allow Excel export to start in parallel.
                if (autoRestartPending) {
                    autoRestartPending = false;
                    autoRestartStopIssuedTime = DateTime::MinValue;
                    this->BeginInvoke(gcnew MethodInvoker(this, &DataForm::ExecuteAutoRestartStart));
                }
            }
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

    // Добавление строки в таблицу данных во время работы устройства
    // Добавляем строку если:
    // 1. Work = 1 (устройство активно работает)
    // 2. Work = 0, но ещё идёт отслеживание переходного периода (workBitZeroTimerActive = true)
    //    Это нужно для записи данных во время "мигания" бита Work перед окончательным завершением
    if (workBitDetected || workBitZeroTimerActive)
    {
        table->Rows->Add(row);
        
        // Сбрасываем флаг экспорта, т.к. появились новые данные, которые ещё не записаны в Excel
        // Это важно для случая, когда после ручной записи продолжают поступать данные
        dataExportedToExcel = false;
    }
}


// 3. Сохраняем таблицу в EXCEL
void ProjectServerW::DataForm::AddDataToExcel() {
    // Create Excel COM objects only after acquiring the global mutex.
    ExcelHelper^ excel = nullptr;

    bool mutexAcquired = false;
    try {
        // Глобальная сериализация экспорта в Excel между всеми DataForm (и даже между процессами).
        // Excel COM-автоматизация часто нестабильна при параллельных экспортерах.
        const int timeoutMs = 5 * 60 * 1000; // 5 минут
        try {
            mutexAcquired = excelGlobalMutex->WaitOne(timeoutMs);
        }
        catch (System::Threading::AbandonedMutexException^) {
            // Предыдущий владелец "умер" не освободив mutex. В .NET он считается захваченным текущим потоком.
            mutexAcquired = true;
            GlobalLogger::LogMessage("Warning: Abandoned Excel mutex detected; continuing export.");
        }

        if (!mutexAcquired) {
            GlobalLogger::LogMessage("Error: Не удалось захватить глобальный Excel mutex (таймаут). Экспорт отменён.");
            return;
        }

        excel = gcnew ExcelHelper();

        // Замеряем время выполнения для мониторинга производительности
        DateTime startTime = DateTime::Now;
        GlobalLogger::LogMessage("Information: Начало записи в Excel...");

        // DataTable не является потокобезопасной структурой.
        // Берём консистентный "снимок" под lock, потому что в это же время UI-поток может добавлять строки.
        System::Data::DataTable^ copiedTable = nullptr;
        System::Threading::Monitor::Enter(dataTableSync);
        try {
            copiedTable = dataTable->Copy();
        }
        finally {
            System::Threading::Monitor::Exit(dataTableSync);
        }
        // Последняя строка может быть не заполнена полностью, её удалим
        if (copiedTable->Rows->Count > 0) {
            copiedTable->Rows->RemoveAt(copiedTable->Rows->Count - 1);
        }

        GlobalLogger::LogMessage(ConvertToStdString(String::Format(
            "Information: Количество строк для записи: {0}, столбцов: {1}", 
            copiedTable->Rows->Count, copiedTable->Columns->Count)));

        if (excel->CreateNewWorkbook()) {
            Microsoft::Office::Interop::Excel::Worksheet^ ws = excel->GetWorksheet();

            // ОПТИМИЗАЦИЯ: отключаем обновление экрана и автоматические вычисления
            Microsoft::Office::Interop::Excel::Application^ excelApp = safe_cast<Microsoft::Office::Interop::Excel::Application^>(ws->Application);
            
            try {
                excelApp->ScreenUpdating = false;
                excelApp->Calculation = Microsoft::Office::Interop::Excel::XlCalculation::xlCalculationManual;
                excelApp->EnableEvents = false;
                GlobalLogger::LogMessage("Information: Excel оптимизирован для быстрой записи");

            // Write session metadata to a dedicated sheet to avoid shifting headers/data and breaking charts.
            try {
                Microsoft::Office::Interop::Excel::Workbook^ wb = safe_cast<Microsoft::Office::Interop::Excel::Workbook^>(ws->Parent);
                System::Object^ missing = System::Type::Missing;

                Microsoft::Office::Interop::Excel::Worksheet^ infoSheet =
                    safe_cast<Microsoft::Office::Interop::Excel::Worksheet^>(wb->Worksheets->Add(missing, ws, 1, Microsoft::Office::Interop::Excel::XlSheetType::xlWorksheet));
                infoSheet->Name = "Info";

                DateTime sessionStart = dataCollectionStartTime;
                if (sessionStart == DateTime::MinValue) {
                    sessionStart = DateTime::Now;
                }
                DateTime sessionEnd = dataCollectionEndTime;
                if (sessionEnd == DateTime::MinValue) {
                    sessionEnd = DateTime::Now;
                }

                infoSheet->Cells[1, 1] = "SessionStart";
                infoSheet->Cells[1, 2] = sessionStart.ToString("yyyy-MM-dd HH:mm:ss");
                infoSheet->Cells[2, 1] = "SessionEnd";
                infoSheet->Cells[2, 2] = sessionEnd.ToString("yyyy-MM-dd HH:mm:ss");
                infoSheet->Cells[3, 1] = "ClientPort";
                infoSheet->Cells[3, 2] = clientPort.ToString();
                infoSheet->Cells[4, 1] = "ClientIP";
                infoSheet->Cells[4, 2] = (this->ClientIP != nullptr ? this->ClientIP : "");
                infoSheet->Cells[5, 1] = "FormGuid";
                infoSheet->Cells[5, 2] = (this->FormGuid != nullptr ? this->FormGuid : "");

                Marshal::ReleaseComObject(infoSheet);
                Marshal::ReleaseComObject(wb);
            }
            catch (...) {
                // Ignore metadata sheet errors to avoid breaking export.
            }

            // Заголовки - ОПТИМИЗИРОВАННАЯ ВЕРСИЯ: записываем все заголовки сразу через массив
            cli::array<cli::array<String^>^>^ bitNames = GetBitFieldNames();
            
            // Определяем количество столбцов
            int colCount = 6 + 4 * (SQ - 1); // базовые столбцы
            if (bitNames[0] != nullptr) colCount += 16;
            if (bitNames[1] != nullptr) colCount += 16;

            // Создаем одномерный массив для заголовков
            cli::array<System::Object^>^ headerArray = gcnew cli::array<System::Object^>(colCount);
            
            int col = 0;
            headerArray[col++] = "RealTime";
            headerArray[col++] = "Time";
            headerArray[col++] = "SQ";
            
            for (uint8_t i = 0; i < (SQ - 1); i++)
            {
                headerArray[col++] = "Typ" + i;
                headerArray[col++] = "Act" + i;
                // Добавляем имя датчика в заголовок T-столбца, если оно задано
                if (sensorNames != nullptr && i < sensorNames->Length && sensorNames[i] != nullptr && sensorNames[i]->Length > 0) {
                    headerArray[col++] = "T" + i + " " + sensorNames[i];
                    headerArray[col++] = "H" + i + " " + sensorNames[i];
                }
                else {
                    headerArray[col++] = "T" + i;
                    headerArray[col++] = "H" + i;
                }
            }
            
            // Обработка сигналов с устройства ввода-вывода
            headerArray[col++] = "Typ" + (SQ - 1);
            headerArray[col++] = "Act" + (SQ - 1);

            // Работа с массивом имен сигналов состояния дефростера
            if (bitNames[0] != nullptr) {
                for (uint8_t i = 0; i < 16; i++)
                {
                    headerArray[col++] = bitNames[0][i];
                }
            }
            // Работа с массивом имен сигналов управления дефростером
            if (bitNames[1] != nullptr) {
                for (uint8_t i = 0; i < 16; i++)
                {
                    headerArray[col++] = bitNames[1][i];
                }
            }

            // КЛЮЧЕВАЯ ОПТИМИЗАЦИЯ: записываем весь ряд заголовков одной операцией
            Microsoft::Office::Interop::Excel::Range^ headerRange = ws->Range[ws->Cells[1, 1], ws->Cells[1, colCount]];
            headerRange->Value2 = headerArray;
            Marshal::ReleaseComObject(headerRange);

            // Данные - ОПТИМИЗИРОВАННАЯ ВЕРСИЯ: записываем все данные сразу через массив
            int rowCount = copiedTable->Rows->Count;
            if (rowCount > 0) {
                // Определяем количество столбцов
                int colCount = 6 + 4 * (SQ - 1); // базовые столбцы
                if (bitNames[0] != nullptr) colCount += 16;
                if (bitNames[1] != nullptr) colCount += 16;

                // Создаем двумерный массив для всех данных
                cli::array<System::Object^, 2>^ dataArray = gcnew cli::array<System::Object^, 2>(rowCount, colCount);

                // Заполняем массив данными из таблицы
                int arrayRow = 0;
                for each (DataRow ^ dr in copiedTable->Rows)
                {
                    int col = 0;
                    dataArray[arrayRow, col++] = dr["RealTime"]->ToString();
                    dataArray[arrayRow, col++] = Convert::ToInt32(dr["Time"]);
                    dataArray[arrayRow, col++] = Convert::ToUInt16(dr["SQ"]);
                    
                    for (uint8_t i = 0; i < (SQ - 1); i++)
                    {
                        dataArray[arrayRow, col++] = Convert::ToUInt16(dr["Typ" + i]);
                        dataArray[arrayRow, col++] = Convert::ToUInt16(dr["Act" + i]);
                        dataArray[arrayRow, col++] = dr["T" + i];
                        dataArray[arrayRow, col++] = dr["H" + i];
                    }

                    // Работа с массивом имен сигналов состояния дефростера
                    if (bitNames[0] != nullptr) {
                        for (uint8_t i = 0; i < 16; i++)
                        {
                            dataArray[arrayRow, col++] = Convert::ToUInt16(dr[bitNames[0][i]]);
                        }
                    }
                    // Работа с массивом имен сигналов управления дефростером
                    if (bitNames[1] != nullptr) {
                        for (uint8_t i = 0; i < 16; i++)
                        {
                            dataArray[arrayRow, col++] = Convert::ToUInt16(dr[bitNames[1][i]]);
                        }
                    }

                    arrayRow++;
                }

                // КЛЮЧЕВАЯ ОПТИМИЗАЦИЯ: записываем весь массив одной операцией
                Microsoft::Office::Interop::Excel::Range^ startCell = safe_cast<Microsoft::Office::Interop::Excel::Range^>(ws->Cells[2, 1]);
                Microsoft::Office::Interop::Excel::Range^ endCell = safe_cast<Microsoft::Office::Interop::Excel::Range^>(ws->Cells[rowCount + 1, colCount]);
                Microsoft::Office::Interop::Excel::Range^ dataRange = ws->Range[startCell, endCell];
                dataRange->Value2 = dataArray;

                // Освобождаем временные объекты
                Marshal::ReleaseComObject(startCell);
                Marshal::ReleaseComObject(endCell);
                Marshal::ReleaseComObject(dataRange);
            }
            int row = rowCount + 2; // Для дальнейшего использования в графике
            // Добавление листа с графиком температур
            try {
                int lastRow = row - 1;
                if (lastRow >= 2) {
                    // Получаем рабочую книгу из рабочего листа
                    Microsoft::Office::Interop::Excel::Workbook^ wb = safe_cast<Microsoft::Office::Interop::Excel::Workbook^>(ws->Parent);
                    System::Object^ missing = System::Type::Missing;

                    // Создаём новый лист для диаграммы
                    Microsoft::Office::Interop::Excel::Worksheet^ chartSheet = safe_cast<Microsoft::Office::Interop::Excel::Worksheet^>(wb->Worksheets->Add(missing, ws, 1, Microsoft::Office::Interop::Excel::XlSheetType::xlWorksheet));
                    chartSheet->Name = "Chart";

                    // Размещаем объект диаграммы
                    Microsoft::Office::Interop::Excel::ChartObjects^ chartObjects = safe_cast<Microsoft::Office::Interop::Excel::ChartObjects^>(chartSheet->ChartObjects(missing));
                    Microsoft::Office::Interop::Excel::ChartObject^ chartObject = chartObjects->Add(20, 20, 900, 500);
                    Microsoft::Office::Interop::Excel::Chart^ chart = chartObject->Chart;
                    chart->ChartType = Microsoft::Office::Interop::Excel::XlChartType::xlLine;

                    // Диапазон X (время) из первого листа: колонка 1 (RealTime)
                    Microsoft::Office::Interop::Excel::Range^ xRange = ws->Range[ws->Cells[2, 1], ws->Cells[lastRow, 1]];

                    // Добавляем серии T0..T(SQ-2)
                    Microsoft::Office::Interop::Excel::SeriesCollection^ seriesCollection = safe_cast<Microsoft::Office::Interop::Excel::SeriesCollection^>(chart->SeriesCollection(missing));
                    for (uint8_t i = 0; i < (SQ - 1); i++)
                    {
                        int seriesCol = 6 + 4 * i; // Колонка T{i}
                        Microsoft::Office::Interop::Excel::Range^ yRange = ws->Range[ws->Cells[2, seriesCol], ws->Cells[lastRow, seriesCol]];
                        Microsoft::Office::Interop::Excel::Series^ s = seriesCollection->NewSeries();
                        s->XValues = xRange;
                        s->Values = yRange;
                        if (sensorNames != nullptr && i < sensorNames->Length && sensorNames[i] != nullptr && sensorNames[i]->Length > 0) {
                            s->Name = "T" + Convert::ToString(i) + " " + sensorNames[i];
                        }
                        else {
                            s->Name = "T" + Convert::ToString(i);
                        }
                        // Применяем цвет к серии, если задан
                        if (sensorColors != nullptr && i < sensorColors->Length) {
                            System::Drawing::Color c = sensorColors[i];
                            int oleColor = System::Drawing::ColorTranslator::ToOle(c);
                            // Color у линии в Chart.Series.Format.Line.ForeColor / или Border.Color
                            try {
                                s->Border->Color = oleColor;
                            }
                            catch (...) {}
                            try {
                                s->Format->Line->ForeColor->RGB = oleColor;
                            }
                            catch (...) {}
                        }
                    }

                    // Заголовки осей
                    chart->HasTitle = true;
                    chart->ChartTitle->Text = "Температуры";
                    Microsoft::Office::Interop::Excel::Axis^ categoryAxis = safe_cast<Microsoft::Office::Interop::Excel::Axis^>(
                        chart->Axes(Microsoft::Office::Interop::Excel::XlAxisType::xlCategory, Microsoft::Office::Interop::Excel::XlAxisGroup::xlPrimary));
                    categoryAxis->HasTitle = true;
                    categoryAxis->AxisTitle->Text = "RealTime";
                    Microsoft::Office::Interop::Excel::Axis^ valueAxis = safe_cast<Microsoft::Office::Interop::Excel::Axis^>(
                        chart->Axes(Microsoft::Office::Interop::Excel::XlAxisType::xlValue, Microsoft::Office::Interop::Excel::XlAxisGroup::xlPrimary));
                    valueAxis->HasTitle = true;
                    valueAxis->AxisTitle->Text = "T, °C";
                }
            }
            catch (Exception^) {
                // Пропускаем создание графика при ошибке, чтобы не сорвать сохранение файла
            }

            } // конец try блока для работы с Excel
            finally {
                // ОПТИМИЗАЦИЯ: ВСЕГДА восстанавливаем настройки Excel перед сохранением
                try {
                    if (excelApp != nullptr) {
                        excelApp->Calculation = Microsoft::Office::Interop::Excel::XlCalculation::xlCalculationAutomatic;
                        excelApp->ScreenUpdating = true;
                        excelApp->EnableEvents = true;
                        Marshal::ReleaseComObject(excelApp);
                        GlobalLogger::LogMessage("Information: Настройки Excel восстановлены");
                    }
                }
                catch (...) {
                    GlobalLogger::LogMessage("Warning: Ошибка при восстановлении настроек Excel");
                }
            }

            // Сохранение файла
            // Формируем полный путь к файлу
            String^ filePath = excelSavePath;
            // Проверяем, существует ли указанный путь
            if (String::IsNullOrEmpty(filePath) || !System::IO::Directory::Exists(filePath)) {
                // Если путь не существует или пустой, создаем директорию в папке запуска программы
                String^ appPath = System::IO::Path::GetDirectoryName(System::Windows::Forms::Application::ExecutablePath);
                filePath = System::IO::Path::Combine(appPath, "SensorData");

                // Создаем директорию, если она не существует
                if (!System::IO::Directory::Exists(filePath)) {
                    System::IO::Directory::CreateDirectory(filePath);
                }

            // Обновляем переменную excelSavePath для будущих сохранений
            excelSavePath = filePath;

            // Обновляем текстовое поле, если оно существует и форма не закрыта
            if (textBoxExcelDirectory != nullptr && !textBoxExcelDirectory->IsDisposed && 
                !this->IsDisposed && this->IsHandleCreated) {
                try {
                    if (textBoxExcelDirectory->InvokeRequired) {
                        textBoxExcelDirectory->BeginInvoke(gcnew System::Action<String^>(this, &DataForm::UpdateDirectoryTextBox), filePath);
                    }
                    else {
                        textBoxExcelDirectory->Text = filePath;
                    }
                }
                catch (...) {
                    // Форма закрыта во время обновления - игнорируем
                }
            }
            // Сохраняем в настройках
            try {
                DataForm::SaveSettings();
            }
            catch (...) {
                // Игнорируем ошибки сохранения настроек при закрытии формы
            }
            }

            // Добавляем разделитель в конец пути, если его нет
            if (!filePath->EndsWith("\\")) {
                filePath += "\\";
            }

            // Build a deterministic filename with session start and finish timestamps (including the date).
            DateTime sessionStart = dataCollectionStartTime;
            if (sessionStart == DateTime::MinValue) {
                sessionStart = DateTime::Now;
            }
            DateTime sessionEnd = dataCollectionEndTime;
            if (sessionEnd == DateTime::MinValue) {
                sessionEnd = DateTime::Now;
            }
            String^ finalFileName = String::Format(
                "WorkData_Start_{0}_End_{1}_Port{2}.xlsx",
                sessionStart.ToString("yyyy-MM-dd_HH-mm-ss"),
                sessionEnd.ToString("yyyy-MM-dd_HH-mm-ss"),
                clientPort.ToString());
            filePath += finalFileName;

            // Сохранение по выбранному пути
            excel->SaveAs(filePath);
            
            // Вычисляем время выполнения
            DateTime endTime = DateTime::Now;
            TimeSpan elapsed = endTime.Subtract(startTime);
            GlobalLogger::LogMessage(ConvertToStdString(String::Format(
                "Information: Файл Excel успешно сохранен: {0}\nВремя записи: {1} секунд ({2} строк)", 
                finalFileName, elapsed.TotalSeconds.ToString("F2"), copiedTable->Rows->Count)));

            // ВАЖНО: НЕ устанавливаем dataExportedToExcel = true здесь!
            // Причина: во время записи в Excel могли поступить НОВЫЕ данные в исходную таблицу,
            // которые ещё не записаны. Флаг управляется в других местах:
            // - При автоматическом экспорте (Work=0, 60 сек) флаг устанавливается ДО вызова записи
            // - При добавлении новых данных флаг сбрасывается в false

            // Освободим память от copiedTable
            delete copiedTable;
            copiedTable = nullptr;
        }
    }
    catch (System::Runtime::InteropServices::COMException^ comEx) {
        // Специальная обработка COM-исключений
        String^ errorMsg = "COM error: " + comEx->Message + " (Code: " +
            comEx->ErrorCode.ToString("X8") + ")";
        //MessageBox::Show(errorMsg);
        GlobalLogger::LogMessage(ConvertToStdString(errorMsg));
    }
    catch (Exception^ ex) {
        // В случае ошибки
        String^ errorMsg = "Excel error: " + ex->Message;
        //MessageBox::Show(errorMsg);
        GlobalLogger::LogMessage(ConvertToStdString("Excel error: Не могу создать файл..." + ex->Message));
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
            // Проверяем, не закрывается ли форма и существует ли она
    // Проверяем состояние формы
            bool canInvoke = false;
            try {
                canInvoke = !this->IsDisposed && this->IsHandleCreated && this->Visible;
            }
            catch (...) {
                canInvoke = false;
            }

            // Включаем кнопку только если можно безопасно вызвать Invoke
            if (canInvoke) {
                try {
                    this->BeginInvoke(gcnew MethodInvoker(this, &DataForm::EnableButton));
                }
                catch (...) {
                    // Игнорируем любые исключения при вызове BeginInvoke
                }
            }
            else {
                // Если форма закрывается, просто пропускаем включение кнопки
            }

            // Только после этого освобождаем Excel
            if (excel != nullptr) {
                excel->Close();
                delete excel;
                excel = nullptr;

                // Затем небольшая пауза для обработки
                Thread::Sleep(50);
                GlobalLogger::LogMessage("Information: Закрываю EXCEL...");
            }

            // Запускаем отложенную сборку мусора
            ThreadPool::QueueUserWorkItem(gcnew WaitCallback(DataForm::DelayedGarbageCollection));
            // Сигнализируем о завершении экспорта
            exportSuccessful = true;
            if (exportCompletedEvent != nullptr) {
                exportCompletedEvent->Set();
            }

            // Critical: allow subsequent exports for this form even if Excel failed.
            System::Threading::Monitor::Enter(excelExportSync);
            try {
                excelExportInProgress = 0;
            }
            finally {
                System::Threading::Monitor::Exit(excelExportSync);
            }

            // Освобождаем глобальный mutex Excel, если он был захвачен
            if (mutexAcquired) {
                try {
                    excelGlobalMutex->ReleaseMutex();
                }
                catch (...) {}
                mutexAcquired = false;
            }
        }
        catch (...) {
            // Critical: allow subsequent exports for this form even if finalization failed.
            System::Threading::Monitor::Enter(excelExportSync);
            try {
                excelExportInProgress = 0;
            }
            finally {
                System::Threading::Monitor::Exit(excelExportSync);
            }

            if (mutexAcquired) {
                try { excelGlobalMutex->ReleaseMutex(); } catch (...) {}
            }
        }
    }    

}

// Отложенная сборка мусора
void DataForm::DelayedGarbageCollection(Object^ state) {
    try {
        // Пауза перед сборкой мусора
        //MessageBox::Show("Excel file was recorded successfully!");
        GlobalLogger::LogMessage(ConvertToStdString("Excel file was recorded successfully!"));
        Thread::Sleep(500);

        // Сборка мусора в отдельном потоке
        GC::Collect();
        GC::WaitForPendingFinalizers();
    }
    catch (...) {}
}

void DataForm::EnableButton()
{
    try {
        if (buttonExcel == nullptr || buttonExcel->IsDisposed) {
            return;
        }
        buttonExcel->Enabled = true;
    }
    catch (...) {}
}

void ProjectServerW::DataForm::OnInactivityTimerTick(Object^ sender, EventArgs^ e) {
    try {
        if (inactivityCloseRequested || !hasTelemetry || lastTelemetryTime == DateTime::MinValue) {
            return;
        }

        TimeSpan idle = DateTime::Now.Subtract(lastTelemetryTime);
        if (idle.TotalMinutes < 30) {
            return;
        }

        // Critical: finalize the current session file (do not create EmergencyData files).
        if (dataCollectionEndTime == DateTime::MinValue) {
            dataCollectionEndTime = DateTime::Now;
        }

        GlobalLogger::LogMessage(ConvertToStdString(String::Format(
            "Information: No telemetry for {0:F1} minutes. Finalizing and closing DataForm.",
            idle.TotalMinutes)));

        // Start export first; only then mark as exported and request close to avoid losing the export due to a transient guard.
        if (!StartExcelExportThread(true)) {
            GlobalLogger::LogMessage("Warning: Export did not start on inactivity timeout; will retry on next tick.");
            return;
        }

        inactivityCloseRequested = true;
        dataExportedToExcel = true;

        // Allow new sessions to create a new form after timeout.
        this->Close();
    }
    catch (Exception^ ex) {
        GlobalLogger::LogMessage(ConvertToStdString("Error: Exception in OnInactivityTimerTick: " + ex->ToString()));
    }
}

void ProjectServerW::DataForm::AddDataToTableThreadSafe(cli::array<System::Byte>^ buffer, int size, int port) {
    // Этот метод выполняется в потоке формы благодаря Invoke
    // ВАЖНО! Методы, работающие с UI, должны выполняться в потоке, создавшем объект с UI!
    // Сохраняем порт клиента в форму DataForm
    this->clientPort = port;

    // Преобразуем управляемый массив байтов в неуправляемый буфер
    pin_ptr<Byte> pinnedBuffer = &buffer[0];
    char* rawBuffer = reinterpret_cast<char*>(pinnedBuffer);

    try {
        if (this == nullptr || this->IsDisposed || this->Disposing) {
            return;
        }
        if (dataGridView == nullptr || dataGridView->IsDisposed) {
            return;
        }

        DateTime telemetryTime = DateTime::Now;
        const SOCKET currentSocket = clientSocket;
        if (hasTelemetry && lastTelemetrySocket != INVALID_SOCKET && currentSocket != INVALID_SOCKET &&
            currentSocket != lastTelemetrySocket) {
            // Why: reconnection is already logged in SServer. This flag adds a correlated "data collection start"
            // log on the first valid telemetry packet of the new TCP connection.
            reconnectFixationLogPending = true;
        }

        hasTelemetry = true;
        lastTelemetryTime = telemetryTime;
        lastTelemetrySocket = currentSocket;

        // Critical: if the controller never toggles Work->1 (or we started mid-session), we still need stable session timestamps.
        if (dataCollectionStartTime == DateTime::MinValue) {
            dataCollectionStartTime = telemetryTime;
        }

        dataGridView->SuspendLayout();
        try {
            // Critical: keep Add and Snapshot mutually exclusive.
            System::Threading::Monitor::Enter(dataTableSync);
            try {
                AddDataToTable(rawBuffer, size, dataTable);
            }
            finally {
                System::Threading::Monitor::Exit(dataTableSync);
            }
        }
        finally {
            dataGridView->ResumeLayout(false);
        }
    }
    catch (Exception^ ex) {
        GlobalLogger::LogMessage(ConvertToStdString("Error: Exception in AddDataToTableThreadSafe: " + ex->ToString()));
    }

    // УБРАНО: Автоматическая прокрутка к последней строке
    // Причина: мешает пользователю просматривать данные в середине таблицы
    // Пользователь может самостоятельно прокрутить к последней строке при необходимости
    
    // УБРАНО: this->Refresh() - полное обновление формы
    // Причина: DataGridView обновляется автоматически при изменении DataSource
    // Полное обновление всей формы замедляет работу
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
        // Line 1 is kept as the legacy Excel directory for backward compatibility.
        writer->WriteLine(textBoxExcelDirectory->Text);

        // Why: the app is frequently deployed without a full installer; a simple text file is easy to carry/backup.
        writer->WriteLine("AutoStartEnabled=" + (checkBoxAutoStart->Checked ? "1" : "0"));
        writer->WriteLine("AutoStartTime=" + dateTimePickerAutoStart->Value.ToString("HH:mm"));
        writer->WriteLine("AutoRestartEnabled=" + (checkBoxAutoRestart->Checked ? "1" : "0"));
        writer->WriteLine("AutoRestartTime=" + dateTimePickerAutoRestart->Value.ToString("HH:mm"));

        // Закрываем файл
        writer->Close();
    }
    catch (Exception^ ex) {
        // Обработка ошибок
        MessageBox::Show("Не удалось сохранить настройки: " + ex->Message);
        GlobalLogger::LogMessage(ConvertToStdString("Не удалось сохранить настройки: " + ex->Message));
    }
}

void ProjectServerW::DataForm::LoadSettings() {
    try {
        settingsLoading = true;
        // Проверяем существование файла настроек
        String^ appPath = System::IO::Path::GetDirectoryName(System::Windows::Forms::Application::ExecutablePath);
        String^ settingsPath = System::IO::Path::Combine(appPath, "ExcelSettings.txt");
        if (System::IO::File::Exists(settingsPath)) {
            cli::array<String^>^ lines = System::IO::File::ReadAllLines(settingsPath);
            if (lines != nullptr && lines->Length > 0) {
                String^ path = lines[0];
                if (path != nullptr && path->Length > 0) {
                    textBoxExcelDirectory->Text = path;
                    excelSavePath = path;
                }
            }

            for (int i = 1; i < lines->Length; i++) {
                String^ line = lines[i];
                if (String::IsNullOrWhiteSpace(line)) {
                    continue;
                }
                int eq = line->IndexOf('=');
                if (eq <= 0 || eq >= line->Length - 1) {
                    continue;
                }

                String^ key = line->Substring(0, eq)->Trim();
                String^ value = line->Substring(eq + 1)->Trim();

                if (key->Equals("AutoStartEnabled", StringComparison::OrdinalIgnoreCase)) {
                    checkBoxAutoStart->Checked = (value == "1" || value->Equals("true", StringComparison::OrdinalIgnoreCase));
                }
                else if (key->Equals("AutoStartTime", StringComparison::OrdinalIgnoreCase)) {
                    DateTime t = DateTime::ParseExact(
                        value, "HH:mm",
                        System::Globalization::CultureInfo::InvariantCulture,
                        System::Globalization::DateTimeStyles::None);
                    DateTime baseDate = dateTimePickerAutoStart->Value;
                    dateTimePickerAutoStart->Value = DateTime(baseDate.Year, baseDate.Month, baseDate.Day, t.Hour, t.Minute, 0);
                }
                else if (key->Equals("AutoRestartEnabled", StringComparison::OrdinalIgnoreCase)) {
                    checkBoxAutoRestart->Checked = (value == "1" || value->Equals("true", StringComparison::OrdinalIgnoreCase));
                }
                else if (key->Equals("AutoRestartTime", StringComparison::OrdinalIgnoreCase)) {
                    DateTime t = DateTime::ParseExact(
                        value, "HH:mm",
                        System::Globalization::CultureInfo::InvariantCulture,
                        System::Globalization::DateTimeStyles::None);
                    DateTime baseDate = dateTimePickerAutoRestart->Value;
                    dateTimePickerAutoRestart->Value = DateTime(baseDate.Year, baseDate.Month, baseDate.Day, t.Hour, t.Minute, 0);
                }
            }
        }
    }
    catch (Exception^ ex) {
        // Обработка ошибок
        MessageBox::Show("Не удалось загрузить настройки: " + ex->Message);
        GlobalLogger::LogMessage(ConvertToStdString("Не удалось загрузить настройки: " + ex->Message));
    }
    finally {
        settingsLoading = false;
    }
}

// Обновление текстового поля директории сохранения файлов EXCEL
void ProjectServerW::DataForm::UpdateDirectoryTextBox(String^ path) {
    if (textBoxExcelDirectory != nullptr && !textBoxExcelDirectory->IsDisposed) {
        textBoxExcelDirectory->Text = path;
    }
}

//*******************************************************************************************
// Реализация метода инициализации названий битовых полей
void ProjectServerW::DataForm::InitializeBitFieldNames(gcroot<cli::array<cli::array<String^>^>^>& namesRef) {
    namesRef = gcnew cli::array<cli::array<String^>^>(10);

    // Инициализация массива состояния устройств дефростера
    namesRef[0] = gcnew cli::array<String^>(16) {
        "Heat0", "Heat1", "Heat2", "Heat3", // Работа нагревателя  (ТЭНа) 1..4
        "Vent0", "Vent1", "Vent2", "Vent3", // Работа циркуляционного вентилятора 1..4
        "InjW",     // Работа водяных форсунок
        "UP",       // Сигнал «Двигать ворота вверх»
        "DOWN",     // Сигнал «Двигать ворота вниз»
        "STOP",     // Сигнал «Остановить ворота»
        "Clse",     // Сигнал «Ворота закрыты»
        "Open",     // Сигнал «Ворота открыты»
        "Work",     // Лампа РАБОТА 
        "Alrm",     // Лампа АВАРИЯ 
        
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
        "_Wrk",     // Включить зелёную лампу РАБОТА
        "_Red"      // Включить красную лампу СТОП
    };

    // Добавьте другие типы сенсоров по мере необходимости
}

////******************** Обработка завершения формы ***************************************
System::Void ProjectServerW::DataForm::DataForm_FormClosing(System::Object^ sender, System::Windows::Forms::FormClosingEventArgs^ e) {
    try {
        // Проверяем, есть ли данные в таблице И не были ли они уже экспортированы
        if (dataTable != nullptr && dataTable->Rows->Count > 0) {
            
            // Проверяем флаг: были ли данные уже записаны в Excel
            if (dataExportedToExcel) {
                // Данные уже были экспортированы - не дублируем запись
                GlobalLogger::LogMessage("Information: Данные уже были экспортированы в Excel, повторная запись не требуется");
            }
            else {
                // Data was not exported yet. If we are closing due to inactivity timeout, we export to the current session file.
                // No "EmergencyData" filenames should be created.
                if (excelFileName == nullptr) {
                    DateTime now = DateTime::Now;
                    if (dataCollectionStartTime == DateTime::MinValue) {
                        dataCollectionStartTime = now;
                    }
                    if (dataCollectionEndTime == DateTime::MinValue) {
                        dataCollectionEndTime = now;
                    }
                    excelFileName = "WorkData_Port" + clientPort.ToString();
                }

                // Логируем начало сохранения
                GlobalLogger::LogMessage("Information: Finalizing data to Excel on form close... " + ConvertToStdString(excelFileName));

                // Start background export without touching UI controls; the form can close immediately.
                StartExcelExportThread(true);
                
                // НЕ ЖДЁМ завершения записи - форма закрывается сразу
                // Это позволяет клиенту переподключиться немедленно
                GlobalLogger::LogMessage("Information: Форма закрывается, запись в Excel продолжается в фоновом режиме");
            }
        }
        
        // Разрешаем закрытие формы немедленно
        e->Cancel = false;
    }
    catch (Exception^ ex) {
        GlobalLogger::LogMessage("Ошибка при попытке сохранения данных: " + ConvertToStdString(ex->Message));
        // В случае ошибки также разрешаем закрытие формы
        e->Cancel = false;
    }
}

// Обработчик события после закрытия формы
System::Void DataForm::DataForm_FormClosed(Object^ sender, FormClosedEventArgs^ e)
{
    // Удаляем форму из карты (если она ещё там есть)
    // Форма могла быть удалена ранее в CloseForm()
    bool formFoundInMap = false;
    for (auto it = formData_Map.begin(); it != formData_Map.end(); ++it) {
        // Извлекаем управляемый указатель из gcroot
        ProjectServerW::DataForm^ formPtr = it->second;

        // Теперь сравниваем указатели
        if (formPtr == this) {
            // Нашли текущую форму, удаляем её из карты
            formData_Map.erase(it);
            formFoundInMap = true;
            GlobalLogger::LogMessage("Information: Форма удалена из карты в FormClosed");
            break;
        }
    }
    
    if (!formFoundInMap) {
        GlobalLogger::LogMessage("Information: Форма уже была удалена из карты ранее");
    }
    
    // Закрываем сокет клиента
    closesocket(this->ClientSocket);

    // Critical: the form can close itself (e.g., inactivity timeout). Ensure the owning native thread is joined/removed.
    try {
        if (this->FormGuid != nullptr && this->FormGuid->Length > 0) {
            msclr::interop::marshal_context ctx;
            std::wstring guidW = ctx.marshal_as<std::wstring>(this->FormGuid);
            ThreadStorage::StopThread(guidW);
        }
    }
    catch (...) {}
}

// Обработчик уничтожения дескриптора окна
System::Void DataForm::DataForm_HandleDestroyed(Object^ sender, EventArgs^ e)
{
    try {
        // Отписываемся от всех событий
        this->FormClosing -= gcnew FormClosingEventHandler(this, &DataForm::DataForm_FormClosing);
        this->FormClosed -= gcnew FormClosedEventHandler(this, &DataForm::DataForm_FormClosed);
        this->HandleDestroyed -= gcnew EventHandler(this, &DataForm::DataForm_HandleDestroyed);

        // Принудительная сборка мусора для освобождения COM-ресурсов
        GC::Collect();
        GC::WaitForPendingFinalizers();
        GC::Collect();
    }
    catch (...) {
        // Игнорируем исключения в деструкторе
    }

    // Очищаем таймер
    if (delayedExcelTimer != nullptr) {
        delayedExcelTimer->Stop();
        delete delayedExcelTimer;
        delayedExcelTimer = nullptr;
    }

    // Stop export retry timer to avoid Tick firing after controls are disposed.
    if (exportTimer != nullptr) {
        try {
            exportTimer->Stop();
            exportTimer->Tick -= gcnew EventHandler(this, &DataForm::CheckExcelButtonStatus);
        }
        catch (...) {}
        exportTimer = nullptr;
    }

    if (inactivityTimer != nullptr) {
        try {
            inactivityTimer->Stop();
            inactivityTimer->Tick -= gcnew EventHandler(this, &DataForm::OnInactivityTimerTick);
        }
        catch (...) {}
        inactivityTimer = nullptr;
    }
}
// Перегруженный метод - автоматическое определение имени команды
bool ProjectServerW::DataForm::SendCommand(const Command& cmd) {
    // Получаем имя команды автоматически
    const char* cmdName = GetCommandName(cmd);
    String^ commandName = gcnew String(cmdName);
    
    // Вызываем основной метод
    return SendCommand(cmd, commandName);
}

// Универсальный метод для отправки команды клиенту с явным именем
bool ProjectServerW::DataForm::SendCommand(const Command& cmd, String^ commandName) {
    try {
        // Проверяем, что сокет клиента валиден
        if (clientSocket == INVALID_SOCKET) {
            MessageBox::Show("Нет активного соединения с клиентом!");
            GlobalLogger::LogMessage("Error: Не могу отправить команду " + ConvertToStdString(commandName) + 
                                   ", нет активного соединения с клиентом!");
            return false;
        }

        // Формируем буфер команды
        uint8_t buffer[MAX_COMMAND_SIZE];
        size_t commandLength = BuildCommandBuffer(cmd, buffer, sizeof(buffer));

        if (commandLength == 0) {
            String^ errorMsg = "Ошибка формирования команды " + commandName;
            MessageBox::Show(errorMsg);
            GlobalLogger::LogMessage("Error: " + ConvertToStdString(errorMsg));
            return false;
        }

        // Critical: serialize send() per socket to prevent byte-stream interleaving with telemetry ACKs.
        System::Object^ sendGate = PacketQueueProcessor::GetSendGate(clientSocket);
        System::Threading::Monitor::Enter(sendGate);
        int bytesSent = SOCKET_ERROR;
        try {
            bytesSent = send(clientSocket, reinterpret_cast<const char*>(buffer),
                static_cast<int>(commandLength), 0);
        }
        finally {
            System::Threading::Monitor::Exit(sendGate);
        }

        if (bytesSent == SOCKET_ERROR) {
            int error = WSAGetLastError();
            String^ errorMsg = "Ошибка отправки команды " + commandName + ": " + error.ToString();
            MessageBox::Show(errorMsg);
            GlobalLogger::LogMessage("Error: " + ConvertToStdString(errorMsg));
            return false;
        }
        else if (bytesSent == commandLength) {
            // Команда успешно отправлена
            Label_Commands->Text = "Команда " + commandName + " отправлена клиенту";
            GlobalLogger::LogMessage("Information: Команда " + ConvertToStdString(commandName) + 
                                   " отправлена клиенту");
            return true;
        }
        else {
            // Отправлено меньше байт, чем ожидалось
            String^ errorMsg = "Отправлено только " + bytesSent.ToString() + " из " + 
                             commandLength.ToString() + " байт для команды " + commandName;
            MessageBox::Show(errorMsg);
            GlobalLogger::LogMessage("Error: Частичная отправка команды " + 
                                   ConvertToStdString(commandName) + ": " + 
                                   ConvertToStdString(errorMsg));
            return false;
        }
    }
    catch (Exception^ ex) {
        String^ errorMsg = "Исключение при отправке команды " + commandName + ": " + ex->Message;
        MessageBox::Show(errorMsg);
        GlobalLogger::LogMessage("Error: " + ConvertToStdString(errorMsg));
        return false;
    }
}

// Метод для отправки команды START клиенту
void ProjectServerW::DataForm::SendStartCommand() {
    // Создаем команду START
    Command cmd = CreateControlCommand(CmdProgControl::START);
    CommandResponse response;
    
    // Отправляем команду и ждем ответ
    if (SendCommandAndWaitResponse(cmd, response)) {
        // Команда успешно выполнена на контроллере
        buttonSTOPstate_TRUE();
        Label_Commands->Text = "[OK] Программа запущена";
        Label_Commands->ForeColor = System::Drawing::Color::Green;
        GlobalLogger::LogMessage("Information: Команда START успешно выполнена контроллером");
        
        // Восстанавливаем цвет через 3 секунды с помощью таймера
        System::Windows::Forms::Timer^ colorTimer = gcnew System::Windows::Forms::Timer();
        colorTimer->Interval = 3000;
        colorTimer->Tick += gcnew EventHandler(this, &DataForm::RestoreLabelCommandsColor);
        colorTimer->Start();
    } else {
        // Ошибка выполнения команды - детали уже обработаны в ProcessResponse
        GlobalLogger::LogMessage(ConvertToStdString(String::Format(
            "Error: Команда START не выполнена. Статус: 0x{0:X2} ({1})",
            response.status, gcnew String(GetStatusName(response.status)))));
        
        // Дополнительная обработка специфичных ошибок для START
        switch (response.status) {
            case CmdStatus::EXECUTION_ERROR:
                // Возможно, программа уже запущена или контроллер не готов
                Label_Commands->Text = "[!] Невозможно запустить программу. Проверьте состояние контроллера";
                Label_Commands->ForeColor = System::Drawing::Color::Orange;
                break;
                
            case CmdStatus::TIMEOUT:
                // Контроллер не успел выполнить запуск
                Label_Commands->Text = "[!] Таймаут запуска программы";
                Label_Commands->ForeColor = System::Drawing::Color::Orange;
                break;
                
            default:
                // Другие ошибки уже отображены в ProcessResponse
                break;
        }
    }
}

// Метод для отправки команды STOP клиенту
void ProjectServerW::DataForm::SendStopCommand() {
    // Создаем команду STOP
    Command cmd = CreateControlCommand(CmdProgControl::STOP);
    CommandResponse response;
    
    // Отправляем команду и ждем ответ
    if (SendCommandAndWaitResponse(cmd, response)) {
        // Команда успешно выполнена на контроллере
        buttonSTARTstate_TRUE();
        Label_Commands->Text = "[OK] Программа остановлена";
        Label_Commands->ForeColor = System::Drawing::Color::Green;
        GlobalLogger::LogMessage("Information: Команда STOP успешно выполнена контроллером");
        
        // Восстанавливаем цвет через 3 секунды с помощью таймера
        System::Windows::Forms::Timer^ colorTimer = gcnew System::Windows::Forms::Timer();
        colorTimer->Interval = 3000;
        colorTimer->Tick += gcnew EventHandler(this, &DataForm::RestoreLabelCommandsColor);
        colorTimer->Start();
    } else {
        // Ошибка выполнения команды - детали уже обработаны в ProcessResponse
        GlobalLogger::LogMessage(ConvertToStdString(String::Format(
            "Error: Команда STOP не выполнена. Статус: 0x{0:X2} ({1})",
            response.status, gcnew String(GetStatusName(response.status)))));
        
        // Дополнительная обработка специфичных ошибок для STOP
        switch (response.status) {
            case CmdStatus::EXECUTION_ERROR:
                // Возможно, программа уже остановлена или контроллер не готов
                Label_Commands->Text = "[!] Невозможно остановить программу. Проверьте состояние контроллера";
                Label_Commands->ForeColor = System::Drawing::Color::Orange;
                break;
                
            case CmdStatus::TIMEOUT:
                // Контроллер не успел выполнить остановку
                Label_Commands->Text = "[!] Таймаут остановки программы";
                Label_Commands->ForeColor = System::Drawing::Color::Orange;
                break;
                
            default:
                // Другие ошибки уже отображены в ProcessResponse
                break;
        }
    }
}

// Метод для отправки команды RESET клиенту
void ProjectServerW::DataForm::SendResetCommand() {
    // Создаем команду RESET
    Command cmd = CreateControlCommand(CmdProgControl::RESET);
    CommandResponse response;
    
    // Отправляем команду и ждем ответ
    if (SendCommandAndWaitResponse(cmd, response)) {
        // Команда успешно выполнена на контроллере
        Label_Commands->Text = "[OK] Контроллер сброшен";
        Label_Commands->ForeColor = System::Drawing::Color::Blue;
        GlobalLogger::LogMessage("Information: Команда RESET успешно выполнена контроллером");
        
        // Восстанавливаем цвет через 3 секунды с помощью таймера
        System::Windows::Forms::Timer^ colorTimer = gcnew System::Windows::Forms::Timer();
        colorTimer->Interval = 3000;
        colorTimer->Tick += gcnew EventHandler(this, &DataForm::RestoreLabelCommandsColor);
        colorTimer->Start();
    } else {
        // Ошибка выполнения команды - детали уже обработаны в ProcessResponse
        GlobalLogger::LogMessage(ConvertToStdString(String::Format(
            "Error: Команда RESET не выполнена. Статус: 0x{0:X2} ({1})",
            response.status, gcnew String(GetStatusName(response.status)))));
        
        // Дополнительная обработка специфичных ошибок для RESET
        switch (response.status) {
            case CmdStatus::EXECUTION_ERROR:
                // Контроллер не может выполнить сброс
                Label_Commands->Text = "[!] Невозможно выполнить сброс контроллера";
                Label_Commands->ForeColor = System::Drawing::Color::Orange;
                break;
                
            case CmdStatus::TIMEOUT:
                // Контроллер не успел выполнить сброс
                Label_Commands->Text = "[!] Таймаут сброса контроллера";
                Label_Commands->ForeColor = System::Drawing::Color::Orange;
                break;
                
            default:
                // Другие ошибки уже отображены в ProcessResponse
                break;
        }
    }
}

// Метод для запроса версии прошивки контроллера
void ProjectServerW::DataForm::SendVersionRequest() {
    // Создаем команду GET_VERSION
    Command cmd;
    cmd.commandType = CmdType::REQUEST;
    cmd.commandCode = CmdRequest::GET_VERSION;
    cmd.dataLength = 0;
    
    CommandResponse response;
    
    // Отправляем команду и ждем ответ
    if (SendCommandAndWaitResponse(cmd, response)) {
        // Команда успешно выполнена контроллером
        if (response.dataLength > 0) {
            // Извлекаем строку версии из ответа
            String^ version = gcnew String(
                reinterpret_cast<const char*>(response.data), 
                0, static_cast<int>(response.dataLength), 
                System::Text::Encoding::ASCII);
            
            // Обновляем label_Version на форме (поточно-безопасно)
            pendingVersion = version; // Сохраняем версию во временное поле
            if (label_Version != nullptr && !label_Version->IsDisposed) {
                if (label_Version->InvokeRequired) {
                    label_Version->BeginInvoke(gcnew System::Windows::Forms::MethodInvoker(
                        this, &DataForm::UpdateVersionLabelInternal));
                } else {
                    label_Version->Text = version;
                }
            }
            
            Label_Commands->Text = "Версия прошивки получена: " + version;
            Label_Commands->ForeColor = System::Drawing::Color::Green;
            GlobalLogger::LogMessage(ConvertToStdString("Information: Версия прошивки контроллера: " + version));
        } else {
            Label_Commands->Text = "Версия получена, но данные пусты";
            Label_Commands->ForeColor = System::Drawing::Color::Orange;
        }
        
        // Восстанавливаем цвет через 3 секунды в обратном таймере
        System::Windows::Forms::Timer^ colorTimer = gcnew System::Windows::Forms::Timer();
        colorTimer->Interval = 3000;
        colorTimer->Tick += gcnew EventHandler(this, &DataForm::RestoreLabelCommandsColor);
        colorTimer->Start();
    } else {
        // Ошибка выполнения команды
        GlobalLogger::LogMessage(ConvertToStdString(String::Format(
            "Error: Команда GET_VERSION не выполнена. Статус: 0x{0:X2} ({1})",
            response.status, gcnew String(GetStatusName(response.status)))));
        
        Label_Commands->Text = "[!] Не удалось получить версию прошивки";
        Label_Commands->ForeColor = System::Drawing::Color::Red;
    }
}

void ProjectServerW::DataForm::SendCommandInfoRequest() {
    // Why: this is a command-processing audit, not device operational state (telemetry).
    // The firmware is expected to remember the last received command and expose it via this request.
    Command cmd = CreateRequestCommand(CmdRequest::GET_CMD_INFO);
    CommandResponse response;

    const DateTime requestTime = DateTime::Now;
    bool received = false;

    // Retry a couple of times: firmware can be busy right after reconnect/start.
    for (int attempt = 0; attempt < 2; attempt++) {
        if (SendCommandAndWaitResponse(cmd, response, "GET_CMD_INFO")) {
            received = true;
            break;
        }
        if (response.status != CmdStatus::TIMEOUT) {
            break;
        }
    }

    if (received) {
        return;
    }

    const bool telemetryRecent = hasTelemetry && lastTelemetryTime != DateTime::MinValue &&
        DateTime::Now.Subtract(lastTelemetryTime).TotalSeconds < 5;

    if (telemetryRecent) {
        GlobalLogger::LogMessage("Warning: GET_CMD_INFO: no response, but telemetry is still arriving (firmware may not be ready to answer yet)");
        if (Label_Commands != nullptr && !Label_Commands->IsDisposed) {
            Label_Commands->Text = "[!] GET_CMD_INFO: нет ответа, но телеметрия идёт (прошивка может быть не готова)";
            Label_Commands->ForeColor = System::Drawing::Color::Orange;
        }
        return;
    }

    TimeSpan waited = DateTime::Now.Subtract(requestTime);
    GlobalLogger::LogMessage(ConvertToStdString(String::Format(
        "Warning: GET_CMD_INFO: no response for {0:F1}s; device may be OFF and the connection can be breaking",
        waited.TotalSeconds)));
    if (Label_Commands != nullptr && !Label_Commands->IsDisposed) {
        Label_Commands->Text = "[!] GET_CMD_INFO: нет ответа (возможно устройство выключено)";
        Label_Commands->ForeColor = System::Drawing::Color::Orange;
    }
}

void ProjectServerW::DataForm::ScheduleCommandInfoProbe(System::String^ reason) {
    // Why: when a response is dropped or a command times out, ask the device what it last received/processed.
    // Guarded to avoid recursion and log spam.
    try {
        if (this == nullptr || this->IsDisposed || this->Disposing || !this->IsHandleCreated) {
            return;
        }

        if (cmdInfoProbeInProgress) {
            return;
        }

        DateTime now = DateTime::Now;
        if (lastCmdInfoProbeTime != DateTime::MinValue) {
            TimeSpan since = now.Subtract(lastCmdInfoProbeTime);
            if (since.TotalSeconds < 2) {
                return;
            }
        }
        lastCmdInfoProbeTime = now;

        GlobalLogger::LogMessage(ConvertToStdString("Information: Scheduling GET_CMD_INFO probe. Reason: " + (reason != nullptr ? reason : "")));

        if (this->InvokeRequired) {
            this->BeginInvoke(gcnew MethodInvoker(this, &DataForm::ExecuteCommandInfoProbe));
        }
        else {
            ExecuteCommandInfoProbe();
        }
    }
    catch (Exception^ ex) {
        GlobalLogger::LogMessage(ConvertToStdString("Error: Exception in ScheduleCommandInfoProbe: " + ex->Message));
    }
}

void ProjectServerW::DataForm::ExecuteCommandInfoProbe() {
    // Why: run probe on UI thread; it uses synchronous command wait.
    if (cmdInfoProbeInProgress) {
        return;
    }

    cmdInfoProbeInProgress = true;
    try {
        SendCommandInfoRequest();
    }
    catch (Exception^ ex) {
        GlobalLogger::LogMessage(ConvertToStdString("Error: Exception in ExecuteCommandInfoProbe: " + ex->ToString()));
    }
    finally {
        cmdInfoProbeInProgress = false;
    }
}

// Вспомогательный метод для обновления label_Version из UI потока
void ProjectServerW::DataForm::UpdateVersionLabelInternal() {
    if (label_Version != nullptr && !label_Version->IsDisposed) {
        label_Version->Text = pendingVersion;
    }
}

// Обработка полученного ответа
void ProjectServerW::DataForm::ProcessResponse(const CommandResponse& response) {
    try {
        // Получаем имя команды и статус
        const char* statusName = GetStatusName(response.status);
        const char* statusDescription = GetStatusDescription(response.status);
        
        // Формируем сообщение о результате выполнения команды
        String^ message;
        
        if (response.status == CmdStatus::OK) {
            message = String::Format(
                "[OK] Команда Type=0x{0:X2}, Code=0x{1:X2} успешно выполнена",
                response.commandType, response.commandCode);

            // Если есть данные в ответе (для команд REQUEST)
            if (response.dataLength > 0) {
                // Обрабатываем данные в зависимости от типа команды
                if (response.commandType == CmdType::REQUEST) {
                    switch (response.commandCode) {
                        case CmdRequest::GET_CMD_INFO: {
                            // Expected payload (firmware contract):
                            // [0] lastCmdType (uint8)
                            // [1] lastCmdCode (uint8)
                            // [2..3] lastCmdDeviceTimeSeconds (uint16, same units as telemetry MSGQUEUE_OBJ_t.Time)
                            // [4] ackSent (uint8, 0/1)
                            // [5] lastCmdStatus (uint8, CmdStatus)
                            if (response.dataLength >= 6) {
                                const uint8_t lastCmdType = response.data[0];
                                const uint8_t lastCmdCode = response.data[1];
                                uint16_t lastCmdSeconds = 0;
                                memcpy(&lastCmdSeconds, &response.data[2], 2);
                                const uint8_t ackSent = response.data[4];
                                const uint8_t lastCmdStatus = response.data[5];

                                // Why: device time is relative; map it to wall-clock using the latest telemetry timestamp.
                                // This is approximate and assumes telemetry and command receiver use the same time base.
                                DateTime approxWallTime = DateTime::MinValue;
                                if (hasTelemetry && lastTelemetryTime != DateTime::MinValue) {
                                    uint16_t cur = lastTelemetryDeviceSeconds;
                                    uint16_t prev = lastCmdSeconds;
                                    int delta = static_cast<int>(cur) - static_cast<int>(prev);
                                    if (delta < 0) {
                                        // Handle uint16 wrap-around (device seconds counter).
                                        delta += 65536;
                                    }
                                    approxWallTime = lastTelemetryTime.Subtract(TimeSpan::FromSeconds(delta));
                                }

                                message += String::Format(
                                    "\nПоследняя команда (устройство): Type=0x{0:X2}, Code=0x{1:X2}"
                                    "\nВремя устройства: {2} сек"
                                    "\nACK отправлен: {3}"
                                    "\nСтатус обработки: 0x{4:X2} ({5})",
                                    lastCmdType,
                                    lastCmdCode,
                                    lastCmdSeconds,
                                    (ackSent != 0 ? "ДА" : "НЕТ"),
                                    lastCmdStatus,
                                    gcnew String(GetStatusName(lastCmdStatus)));

                                if (approxWallTime != DateTime::MinValue) {
                                    message += "\nПримерное реальное время приёма: " + approxWallTime.ToString("yyyy-MM-dd HH:mm:ss");
                                }
                            }
                            else {
                                message += "\nGET_CMD_INFO: недостаточно данных ответа";
                            }
                            break;
                        }
                        case CmdRequest::GET_VERSION: {
                            // Получаем версию (строка)
                            String^ version = gcnew String(
                                reinterpret_cast<const char*>(response.data), 
                                0, static_cast<int>(response.dataLength), System::Text::Encoding::ASCII);
                            message += "\nВерсия прошивки: " + version;
                            break;
                        }
                        case CmdRequest::GET_DATA: {
                            // Получаем данные конфигурации
                            if (response.dataLength >= 1) {
                                message += String::Format("\nРежим работы: {0}", 
                                    response.data[0] == 0 ? "Автоматический" : "Ручной");
                            }
                            break;
                        }
                    }
                }
            }
            
            // Отображаем успешный результат
            Label_Commands->Text = message;
            GlobalLogger::LogMessage(ConvertToStdString(message));
            
        } else {
            // ===== ОБРАБОТКА ОШИБОК =====
            
            // Формируем детальное сообщение об ошибке
            message = String::Format(
                "[ОШИБКА] ОШИБКА выполнения команды\n\n"
                "Команда:\n"
                "  Тип: 0x{0:X2}\n"
                "  Код: 0x{1:X2}\n\n"
                "Статус ошибки:\n"
                "  Код: 0x{2:X2} ({3})\n\n"
                "Описание:\n"
                "  {4}",
                response.commandType, 
                response.commandCode, 
                response.status,
                gcnew String(statusName),
                gcnew String(statusDescription));
            
            // Добавляем специфичные для ошибки рекомендации
            String^ recommendation = "";
            switch (response.status) {
                case CmdStatus::CRC_ERROR:
                    recommendation = "\n\nРекомендация:\n"
                                   "  • Проверьте качество соединения\n"
                                   "  • Проверьте экранирование кабеля\n"
                                   "  • Уменьшите скорость передачи данных";
                    break;
                    
                case CmdStatus::INVALID_TYPE:
                case CmdStatus::INVALID_CODE:
                    recommendation = "\n\nРекомендация:\n"
                                   "  • Обновите прошивку контроллера\n"
                                   "  • Проверьте совместимость версий\n"
                                   "  • Убедитесь, что команда поддерживается";
                    break;
                    
                case CmdStatus::INVALID_LENGTH:
                    recommendation = "\n\nРекомендация:\n"
                                   "  • Проверьте параметры команды\n"
                                   "  • Команда может требовать другой набор данных";
                    break;
                    
                case CmdStatus::EXECUTION_ERROR:
                    recommendation = "\n\nРекомендация:\n"
                                   "  • Проверьте текущее состояние контроллера\n"
                                   "  • Команда может быть недоступна в текущем режиме\n"
                                   "  • Проверьте наличие необходимых условий для выполнения\n"
                                   "  • Попробуйте выполнить команду позже";
                    break;
                    
                case CmdStatus::TIMEOUT:
                    recommendation = "\n\nРекомендация:\n"
                                   "  • Команда требует длительного выполнения\n"
                                   "  • Увеличьте таймаут ожидания\n"
                                   "  • Проверьте, не занят ли контроллер другой операцией";
                    break;
                    
                case CmdStatus::UNKNOWN_ERROR:
                    recommendation = "\n\nРекомендация:\n"
                                   "  • Перезагрузите контроллер\n"
                                   "  • Проверьте журнал ошибок контроллера\n"
                                   "  • Обратитесь в техническую поддержку";
                    break;
            }
            
            message += recommendation;
            
            // Отображаем сообщение об ошибке
            Label_Commands->Text = "[ОШИБКА] " + gcnew String(statusDescription);
            Label_Commands->ForeColor = System::Drawing::Color::Red;
            
            MessageBox::Show(message, "Ошибка выполнения команды", 
                           MessageBoxButtons::OK, MessageBoxIcon::Error);
            GlobalLogger::LogMessage(ConvertToStdString(message));
            
            // Восстанавливаем цвет текста через 3 секунды с помощью таймера
            System::Windows::Forms::Timer^ colorTimer = gcnew System::Windows::Forms::Timer();
            colorTimer->Interval = 3000;
            colorTimer->Tick += gcnew EventHandler(this, &DataForm::RestoreLabelCommandsColor);
            colorTimer->Start();
        }

    } catch (Exception^ ex) {
        String^ errorMsg = "Исключение в ProcessResponse: " + ex->Message;
        MessageBox::Show(errorMsg, "Критическая ошибка", 
                       MessageBoxButtons::OK, MessageBoxIcon::Error);
        GlobalLogger::LogMessage(ConvertToStdString(errorMsg));
    }
}

// Восстановление цвета Label_Commands (вызывается таймером)
void ProjectServerW::DataForm::RestoreLabelCommandsColor(System::Object^ sender, System::EventArgs^ e) {
    try {
        // Останавливаем таймер
        System::Windows::Forms::Timer^ timer = safe_cast<System::Windows::Forms::Timer^>(sender);
        timer->Stop();
        timer->Tick -= gcnew EventHandler(this, &DataForm::RestoreLabelCommandsColor);
        
        // Восстанавливаем цвет
        Label_Commands->ForeColor = System::Drawing::SystemColors::ControlText;
    }
    catch (Exception^ ex) {
        GlobalLogger::LogMessage(ConvertToStdString("Error in RestoreLabelCommandsColor: " + ex->Message));
    }
}

// Отправка команды и ожидание ответа
bool ProjectServerW::DataForm::SendCommandAndWaitResponse(
    const Command& cmd, CommandResponse& response, System::String^ commandName) {
    
    try {
        // Отправляем команду
        bool sendResult;
        if (commandName != nullptr) {
            sendResult = SendCommand(cmd, commandName);
        } else {
            sendResult = SendCommand(cmd);
        }

        if (!sendResult) {
            GlobalLogger::LogMessage("Error: Failed to send command");
            return false;
        }

        // Wait for the matching response and tolerate unrelated responses in the shared queue.
        // Why: the recv() loop can enqueue responses from earlier commands; rejecting on first mismatch
        // breaks auto flows and can incorrectly mark a valid command as failed.
        const bool isCmdInfoRequest = (cmd.commandType == CmdType::REQUEST && cmd.commandCode == CmdRequest::GET_CMD_INFO);

        const int defaultTimeoutMs = 2000;
        const int totalTimeoutMs = defaultTimeoutMs;
        DateTime deadline = DateTime::Now.AddMilliseconds(totalTimeoutMs);
        while (true) {
            TimeSpan remaining = deadline.Subtract(DateTime::Now);
            if (remaining.TotalMilliseconds <= 0) {
                response.commandType = cmd.commandType;
                response.commandCode = cmd.commandCode;
                response.status = CmdStatus::TIMEOUT;
                response.dataLength = 0;
                GlobalLogger::LogMessage("Error: No response received from controller");
                if (!isCmdInfoRequest) {
                    ScheduleCommandInfoProbe(String::Format(
                        "timeout waiting response Type=0x{0:X2}, Code=0x{1:X2}",
                        cmd.commandType, cmd.commandCode));
                }
                return false;
            }

            CommandResponse candidate{};
            cli::array<System::Byte>^ rawCandidate = nullptr;
            if (!ReceiveResponse(candidate, static_cast<int>(remaining.TotalMilliseconds), rawCandidate)) {
                continue;
            }

            if (candidate.commandType != cmd.commandType || candidate.commandCode != cmd.commandCode) {
                String^ hex = "";
                if (rawCandidate != nullptr && rawCandidate->Length > 0) {
                    System::Text::StringBuilder^ sb = gcnew System::Text::StringBuilder(rawCandidate->Length * 3);
                    for (int i = 0; i < rawCandidate->Length; i++) {
                        if (i != 0) sb->Append(" ");
                        sb->Append(rawCandidate[i].ToString("X2"));
                    }
                    hex = sb->ToString();
                }

                GlobalLogger::LogMessage(ConvertToStdString(String::Format(
                    "Warning: Discarding unrelated response. Expected Type=0x{0:X2}, Code=0x{1:X2}; Got Type=0x{2:X2}, Code=0x{3:X2}; Raw={4}",
                    cmd.commandType, cmd.commandCode, candidate.commandType, candidate.commandCode, hex)));
                if (!isCmdInfoRequest) {
                    ScheduleCommandInfoProbe(String::Format(
                        "discarded unrelated response while waiting Type=0x{0:X2}, Code=0x{1:X2}",
                        cmd.commandType, cmd.commandCode));
                }
                continue;
            }

            response = candidate;
            break;
        }

        // Обрабатываем полученный ответ
        ProcessResponse(response);

        return (response.status == CmdStatus::OK);

    } catch (Exception^ ex) {
        String^ errorMsg = "Exception in SendCommandAndWaitResponse: " + ex->Message;
        MessageBox::Show(errorMsg);
        GlobalLogger::LogMessage(ConvertToStdString(errorMsg));
        return false;
    }
}

// Перегрузка SendCommandAndWaitResponse без указания имени команды
bool ProjectServerW::DataForm::SendCommandAndWaitResponse(
    const Command& cmd, CommandResponse& response) {
    return SendCommandAndWaitResponse(cmd, response, nullptr);
}

void ProjectServerW::DataForm::buttonSTARTstate_TRUE()
{
    buttonSTART->Enabled = true;
    labelSTART->BackColor = System::Drawing::Color::Snow;
    labelSTART->Text = "0";
    buttonSTOP->Enabled = false;
    labelSTOP->BackColor = System::Drawing::Color::OrangeRed;
    labelSTOP->Text = "1";
}

void ProjectServerW::DataForm::buttonSTOPstate_TRUE()
{
    buttonSTART->Enabled = false;
    labelSTART->BackColor = System::Drawing::Color::Lime;
    labelSTART->Text = "1";
    buttonSTOP->Enabled = true;
    labelSTOP->BackColor = System::Drawing::Color::Snow;
    labelSTOP->Text = "0";
}

// Обработчик события таймера для отложенной записи в Excel
void ProjectServerW::DataForm::OnDelayedExcelTimerTick(Object^ sender, EventArgs^ e) {
    // Проверяем, активен ли таймер отслеживания нуля
    if (!workBitZeroTimerActive) {
        delayedExcelTimer->Stop();
        return;
    }

    // Вычисляем прошедшее время с момента установки бита Work в ноль
    DateTime now = DateTime::Now;
    TimeSpan elapsed = now.Subtract(workBitZeroStartTime);
    
    if (elapsed.TotalSeconds >= 60) {
        // Прошло не менее 1 минуты непрерывного нахождения в нуле
        workBitZeroTimerActive = false;
        workBitZeroLogged = false;  // Сбрасываем флаг логирования
        delayedExcelTimer->Stop();
        
        // Сохраняем время окончания сбора данных
        dataCollectionEndTime = now;
        
        GlobalLogger::LogMessage(ConvertToStdString("Information: СТОП фиксации данных (по таймеру), запись в файл " + excelFileName));
        // Меняем состояние кнопок СТАРТ и СТОП
        buttonSTARTstate_TRUE();
        
        // Устанавливаем флаг, что данные экспортированы (чтобы не дублировать при закрытии формы)
        dataExportedToExcel = true;
        
        // Выполняем запись в Excel
        this->BeginInvoke(gcnew MethodInvoker(this, &DataForm::TriggerExcelExport));
    }
}

// ????????????????????????????????????????????????????????????????????????????
// Реализация методов записи в Excel (отложенный запуск)
// ????????????????????????????????????????????????????????????????????????????

// Триггер автоматического экспорта в Excel
void ProjectServerW::DataForm::TriggerExcelExport() {
    // Проверяем, вызывается ли метод из правильного потока
    if (this->InvokeRequired) {
        // Вызываем из потока UI
        try {
            this->Invoke(gcnew MethodInvoker(this, &DataForm::TriggerExcelExport));
        }
        catch (Exception^ ex) {
            GlobalLogger::LogMessage("Ошибка при вызове TriggerExcelExport: " + ConvertToStdString(ex->Message));
        }
        return;
    }
    
    // Теперь мы в потоке UI, можем безопасно обращаться к элементам управления
    try {
        if (StartExcelExportThread(false)) {
            pendingExcelExport = false;
            if (exportTimer != nullptr && exportTimer->Enabled) {
                exportTimer->Stop();
            }
            return;
        }

        pendingExcelExport = true;
        if (exportTimer == nullptr) {
            exportTimer = gcnew System::Windows::Forms::Timer();
            exportTimer->Interval = 500;
            exportTimer->Tick += gcnew EventHandler(this, &DataForm::CheckExcelButtonStatus);
        }
        exportTimer->Start();
        if (Label_Data != nullptr && !Label_Data->IsDisposed) {
            Label_Data->Text = "Ожидание возможности записи в Excel...";
        }
    }
    catch (Exception^ ex) {
        GlobalLogger::LogMessage("Ошибка в TriggerExcelExport: " + ConvertToStdString(ex->Message));
    }
}

// Обработчик события таймера проверки доступности кнопки экспорта в Excel
void ProjectServerW::DataForm::CheckExcelButtonStatus(Object^ sender, EventArgs^ e) {
    if (!pendingExcelExport) {
        return;
    }
    if (!StartExcelExportThread(false)) {
        return;
    }

    pendingExcelExport = false;
    if (exportTimer != nullptr && exportTimer->Enabled) {
        exportTimer->Stop();
    }
    if (Label_Data != nullptr && !Label_Data->IsDisposed) {
        Label_Data->Text = "Данные записываются в Excel...";
    }
}

// ????????????????????????????????????????????????????????????????????????????
// Реализация методов автозапуска по времени
// ????????????????????????????????????????????????????????????????????????????

// Обработчик изменения состояния чекбокса автозапуска
System::Void ProjectServerW::DataForm::checkBoxAutoStart_CheckedChanged(System::Object^ sender, System::EventArgs^ e) {
    // Включение/выключение автозапуска
    if (checkBoxAutoStart->Checked) {
        // Включаем таймер
        timerAutoStart->Start();
        GlobalLogger::LogMessage(ConvertToStdString(String::Format(
            "Information: Автозапуск включен на {0}",
            dateTimePickerAutoStart->Value.ToString("HH:mm"))));
        
        // Визуальная индикация
        labelAutoStart->ForeColor = System::Drawing::Color::Green;
        dateTimePickerAutoStart->Enabled = true;
    }
    else {
        // Выключаем таймер
        timerAutoStart->Stop();
        GlobalLogger::LogMessage("Information: Автозапуск отключен");
        
        // Визуальная индикация
        labelAutoStart->ForeColor = System::Drawing::SystemColors::ControlText;
    }
}

// Обработчик тика таймера автозапуска
System::Void ProjectServerW::DataForm::timerAutoStart_Tick(System::Object^ sender, System::EventArgs^ e) {
    // Проверка времени для автозапуска
    if (!checkBoxAutoStart->Checked) {
        return; // Если автозапуск выключен, ничего не делаем
    }
    
    DateTime now = DateTime::Now;
    DateTime targetTime = dateTimePickerAutoStart->Value;
    
    // Сравниваем часы и минуты
    if (now.Hour == targetTime.Hour && now.Minute == targetTime.Minute) {
        // Время совпало!
        GlobalLogger::LogMessage(ConvertToStdString(String::Format(
            "Information: Автозапуск сработал в {0} (установлено: {1})",
            now.ToString("HH:mm:ss"),
            targetTime.ToString("HH:mm"))));
        
        // Проверяем, что кнопка START доступна (программа не запущена)
        if (buttonSTART->Enabled) {
            // Визуальная индикация
            labelAutoStart->ForeColor = System::Drawing::Color::Blue;
            Label_Commands->Text = "[АВТОЗАПУСК] Автоматический запуск программы...";
            Label_Commands->ForeColor = System::Drawing::Color::Blue;
            
            CommandResponse startResp{};
            CommandAckResult startResult = SendControlCommandWithAck(CmdProgControl::START, "START", 2000, 2, startResp);

            if (startResult == CommandAckResult::NoResponse) {
                // Why: controller can start but ACK can be lost; verify via Work bit in telemetry.
                GlobalLogger::LogMessage("Warning: Автозапуск: START без подтверждения, проверка по биту Work");
                Label_Commands->Text = "[АВТОЗАПУСК] START без подтверждения, проверка по Work...";
                Label_Commands->ForeColor = System::Drawing::Color::Orange;
            }
            else if (startResult != CommandAckResult::Ok) {
                GlobalLogger::LogMessage("Warning: Автозапуск: START не выполнен (ошибка ответа/отправки)");
                Label_Commands->Text = "[!] Автозапуск: START не выполнен";
                Label_Commands->ForeColor = System::Drawing::Color::Orange;
            }
            
            // Отключаем автозапуск после выполнения
            checkBoxAutoStart->Checked = false;
            
            // Восстанавливаем цвет через 3 секунды
            System::Windows::Forms::Timer^ colorTimer = gcnew System::Windows::Forms::Timer();
            colorTimer->Interval = 3000;
            colorTimer->Tick += gcnew EventHandler(this, &DataForm::RestoreAutoStartColor);
            colorTimer->Start();
        }
        else {
            // Программа уже запущена
            GlobalLogger::LogMessage("Warning: Автозапуск не выполнен - программа уже запущена");
            Label_Commands->Text = "[!] Автозапуск пропущен - программа уже работает";
            Label_Commands->ForeColor = System::Drawing::Color::Orange;
            
            // Отключаем автозапуск
            checkBoxAutoStart->Checked = false;
        }
    }
}

// Обработчик восстановления цвета надписи автозапуска
System::Void ProjectServerW::DataForm::RestoreAutoStartColor(System::Object^ sender, System::EventArgs^ e) {
    try {
        // Останавливаем таймер
        System::Windows::Forms::Timer^ timer = safe_cast<System::Windows::Forms::Timer^>(sender);
        timer->Stop();
        timer->Tick -= gcnew EventHandler(this, &DataForm::RestoreAutoStartColor);
        
        // Восстанавливаем цвет
        labelAutoStart->ForeColor = System::Drawing::SystemColors::ControlText;
    }
    catch (Exception^ ex) {
        GlobalLogger::LogMessage(ConvertToStdString("Error in RestoreAutoStartColor: " + ex->Message));
    }
}

// ????????????????????????????????????????????????????????????????????????????
// Реализация методов автоперезапуска по времени
// ????????????????????????????????????????????????????????????????????????????

System::Void ProjectServerW::DataForm::checkBoxAutoRestart_CheckedChanged(System::Object^ sender, System::EventArgs^ e) {
    if (checkBoxAutoRestart->Checked) {
        timerAutoRestart->Start();
        GlobalLogger::LogMessage(ConvertToStdString(String::Format(
            "Information: Автоперезапуск включен на {0}",
            dateTimePickerAutoRestart->Value.ToString("HH:mm"))));

        labelAutoRestart->ForeColor = System::Drawing::Color::Green;
    }
    else {
        timerAutoRestart->Stop();
        if (autoRestartInternalUncheck) {
            autoRestartInternalUncheck = false;
        }
        else {
            autoRestartPending = false;
        }
        GlobalLogger::LogMessage("Information: Автоперезапуск отключен");

        labelAutoRestart->ForeColor = System::Drawing::SystemColors::ControlText;
    }
    if (!settingsLoading) {
        SaveSettings();
    }
}

System::Void ProjectServerW::DataForm::dateTimePickerAutoStart_ValueChanged(System::Object^ sender, System::EventArgs^ e) {
    if (!settingsLoading) {
        SaveSettings();
    }
}

System::Void ProjectServerW::DataForm::dateTimePickerAutoRestart_ValueChanged(System::Object^ sender, System::EventArgs^ e) {
    if (!settingsLoading) {
        SaveSettings();
    }
}

System::Void ProjectServerW::DataForm::timerAutoRestart_Tick(System::Object^ sender, System::EventArgs^ e) {
    if (!checkBoxAutoRestart->Checked) {
        return;
    }

    DateTime now = DateTime::Now;
    DateTime targetTime = dateTimePickerAutoRestart->Value;

    if (now.Hour != targetTime.Hour || now.Minute != targetTime.Minute) {
        return;
    }

    GlobalLogger::LogMessage(ConvertToStdString(String::Format(
        "Information: Автоперезапуск сработал в {0} (установлено: {1})",
        now.ToString("HH:mm:ss"),
        targetTime.ToString("HH:mm"))));

    // If we don't know device state yet (no telemetry), do not consume the schedule.
    if (!buttonSTOP->Enabled && !buttonSTART->Enabled) {
        GlobalLogger::LogMessage("Warning: Автоперезапуск: состояние устройства неизвестно (нет телеметрии), ожидание...");
        return;
    }

    labelAutoRestart->ForeColor = System::Drawing::Color::Blue;

    // If the device is currently running (STOP is available), issue STOP and wait for stable Work=0.
    if (buttonSTOP->Enabled) {
        Label_Commands->Text = "[АВТОПЕРЕЗАПУСК] Отправка команды STOP...";
        Label_Commands->ForeColor = System::Drawing::Color::Blue;

        CommandResponse stopResp{};
        CommandAckResult stopResult = SendControlCommandWithAck(CmdProgControl::STOP, "STOP", 2000, 2, stopResp);

        if (stopResult == CommandAckResult::Ok) {
            autoRestartPending = true;
            autoRestartStopIssuedTime = DateTime::Now;
            Label_Commands->Text = "[АВТОПЕРЕЗАПУСК] STOP подтверждён, ожидание окончания работы...";
            Label_Commands->ForeColor = System::Drawing::Color::Blue;
        }
        else if (stopResult == CommandAckResult::NoResponse) {
            // Why: controller can execute STOP but ACK can be lost; keep the flow and verify via Work bit.
            autoRestartPending = true;
            autoRestartStopIssuedTime = DateTime::Now;
            GlobalLogger::LogMessage("Warning: Автоперезапуск: STOP без подтверждения, ожидание по биту Work");
            Label_Commands->Text = "[АВТОПЕРЕЗАПУСК] STOP без подтверждения, ожидание по Work...";
            Label_Commands->ForeColor = System::Drawing::Color::Orange;
        }
        else {
            autoRestartPending = false;
            autoRestartStopIssuedTime = DateTime::MinValue;
            GlobalLogger::LogMessage("Warning: Автоперезапуск: STOP не выполнен (ошибка ответа/отправки)");
            Label_Commands->Text = "[!] Автоперезапуск: STOP не выполнен";
            Label_Commands->ForeColor = System::Drawing::Color::Orange;
        }

        autoRestartInternalUncheck = true;
        checkBoxAutoRestart->Checked = false; // one-shot, same UX as AutoStart
    }
    else if (buttonSTART->Enabled) {
        // Device is already stopped; behave as scheduled start.
        Label_Commands->Text = "[АВТОПЕРЕЗАПУСК] Устройство уже остановлено, отправка START...";
        Label_Commands->ForeColor = System::Drawing::Color::Blue;

        CommandResponse startResp{};
        CommandAckResult startResult = SendControlCommandWithAck(CmdProgControl::START, "START", 2000, 2, startResp);

        if (startResult == CommandAckResult::Ok) {
            Label_Commands->Text = "[АВТОПЕРЕЗАПУСК] START подтверждён";
            Label_Commands->ForeColor = System::Drawing::Color::Blue;
        }
        else if (startResult == CommandAckResult::NoResponse) {
            // Why: keep reacting to missing ACK, but allow telemetry Work to confirm actual state.
            GlobalLogger::LogMessage("Warning: Автоперезапуск: START без подтверждения, ожидание по биту Work");
            Label_Commands->Text = "[АВТОПЕРЕЗАПУСК] START без подтверждения, проверка по Work...";
            Label_Commands->ForeColor = System::Drawing::Color::Orange;
        }
        else {
            GlobalLogger::LogMessage("Warning: Автоперезапуск: START не выполнен (ошибка ответа/отправки)");
            Label_Commands->Text = "[!] Автоперезапуск: START не выполнен";
            Label_Commands->ForeColor = System::Drawing::Color::Orange;
        }
        autoRestartInternalUncheck = true;
        checkBoxAutoRestart->Checked = false;
    }

    // Restore color through a short-lived timer to avoid a permanent "armed" look.
    System::Windows::Forms::Timer^ colorTimer = gcnew System::Windows::Forms::Timer();
    colorTimer->Interval = 3000;
    colorTimer->Tick += gcnew EventHandler(this, &DataForm::RestoreAutoRestartColor);
    colorTimer->Start();
}

System::Void ProjectServerW::DataForm::RestoreAutoRestartColor(System::Object^ sender, System::EventArgs^ e) {
    try {
        System::Windows::Forms::Timer^ timer = safe_cast<System::Windows::Forms::Timer^>(sender);
        timer->Stop();
        timer->Tick -= gcnew EventHandler(this, &DataForm::RestoreAutoRestartColor);
        labelAutoRestart->ForeColor = System::Drawing::SystemColors::ControlText;
    }
    catch (Exception^ ex) {
        GlobalLogger::LogMessage(ConvertToStdString("Error in RestoreAutoRestartColor: " + ex->Message));
    }
}

void ProjectServerW::DataForm::ExecuteAutoRestartStart() {
    try {
        if (buttonSTART == nullptr || buttonSTART->IsDisposed) {
            return;
        }
        if (!buttonSTART->Enabled) {
            GlobalLogger::LogMessage("Warning: Автоперезапуск: START пропущен (программа уже запущена)");
            return;
        }

        Label_Commands->Text = "[АВТОПЕРЕЗАПУСК] Автоматический запуск программы...";
        Label_Commands->ForeColor = System::Drawing::Color::Blue;
        SendStartCommand();
    }
    catch (Exception^ ex) {
        GlobalLogger::LogMessage(ConvertToStdString("Error: Exception in ExecuteAutoRestartStart: " + ex->ToString()));
    }
}

bool ProjectServerW::DataForm::TrySendControlCommandFireAndForget(uint8_t controlCode, System::String^ commandName) {
    // Why: scheduled operations must not depend on a synchronous command response because some firmware builds
    // can execute START/STOP but skip/lose the ACK. We confirm the actual state via the Work bit telemetry.
    try {
        Command cmd = CreateControlCommand(controlCode);
        return SendCommand(cmd, commandName);
    }
    catch (Exception^ ex) {
        GlobalLogger::LogMessage(ConvertToStdString("Error: Exception in TrySendControlCommandFireAndForget: " + ex->Message));
        return false;
    }
}

ProjectServerW::DataForm::CommandAckResult ProjectServerW::DataForm::SendControlCommandWithAck(
    uint8_t controlCode,
    System::String^ commandName,
    int timeoutMs,
    int retries,
    CommandResponse% lastResponse)
{
    // Why: keep the protocol behavior (wait for ACK and react) but remain resilient when ACK is lost.
    // Fallback state verification is performed via telemetry Work bit in higher-level logic.
    lastResponse.commandType = CmdType::PROG_CONTROL;
    lastResponse.commandCode = controlCode;
    lastResponse.status = CmdStatus::TIMEOUT;
    lastResponse.dataLength = 0;
    bool hadTimeout = false;

    for (int attempt = 0; attempt < retries; attempt++) {
        Command cmd = CreateControlCommand(controlCode);
        if (!SendCommand(cmd, commandName)) {
            return CommandAckResult::SendFailed;
        }

        CommandResponse response{};
        cli::array<System::Byte>^ rawResponse = nullptr;
        if (!ReceiveResponse(response, timeoutMs, rawResponse)) {
            hadTimeout = true;
            lastResponse.status = CmdStatus::TIMEOUT;
            continue;
        }

        // Some responses can be for other commands; emulate SendCommandAndWaitResponse matching behavior here as well.
        if (response.commandType != cmd.commandType || response.commandCode != cmd.commandCode) {
            String^ hex = "";
            if (rawResponse != nullptr && rawResponse->Length > 0) {
                System::Text::StringBuilder^ sb = gcnew System::Text::StringBuilder(rawResponse->Length * 3);
                for (int i = 0; i < rawResponse->Length; i++) {
                    if (i != 0) sb->Append(" ");
                    sb->Append(rawResponse[i].ToString("X2"));
                }
                hex = sb->ToString();
            }
            GlobalLogger::LogMessage(ConvertToStdString(String::Format(
                "Warning: Discarding unrelated response in SendControlCommandWithAck. Expected Type=0x{0:X2}, Code=0x{1:X2}; Got Type=0x{2:X2}, Code=0x{3:X2}; Raw={4}",
                cmd.commandType, cmd.commandCode, response.commandType, response.commandCode, hex)));
            ScheduleCommandInfoProbe(String::Format(
                "discarded unrelated response while waiting ACK Type=0x{0:X2}, Code=0x{1:X2}",
                cmd.commandType, cmd.commandCode));
            lastResponse = response;
            continue;
        }

        lastResponse = response;
        ProcessResponse(response);
        return (response.status == CmdStatus::OK) ? CommandAckResult::Ok : CommandAckResult::ErrorResponse;
    }

    if (hadTimeout) {
        ScheduleCommandInfoProbe(String::Format(
            "timeout waiting ACK Type=0x{0:X2}, Code=0x{1:X2}",
            CmdType::PROG_CONTROL, controlCode));
    }
    return CommandAckResult::NoResponse;
}

System::Void ProjectServerW::DataForm::button_CMDINFO_Click(System::Object^ sender, System::EventArgs^ e) {
    try {
        SendCommandInfoRequest();
    }
    catch (Exception^ ex) {
        GlobalLogger::LogMessage(ConvertToStdString("Error: Exception in button_CMDINFO_Click: " + ex->ToString()));
    }
}
