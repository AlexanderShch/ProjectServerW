#include "DataForm.h"
#include "Chart.h"
#include "FormExcel.h"
#include "Commands.h"               // Для работы с командами управления
#include "PacketQueueProcessor.h"   // Общий per-socket "затвор" отправки (ACK телеметрии vs команды из UI)
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
typedef struct   // Данные телеметрии (формат STM32)
{
    uint8_t DataType;           // Байт типа передаваемых данных (0х00 для телеметрии)
    uint8_t Len;                // Длина полезной части (байты после Len и до CRC), включается в CRC
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
        // Критично: защита от повторной постановки экспорта в очередь для одной формы (UI может триггерить экспорт из разных мест).
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

        // Отключаем кнопку только для не-аварийного экспорта, чтобы не трогать disposed-UI во время завершения приложения.
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

        ProjectServerW::FormExcel::ExcelExportJob^ job = gcnew ProjectServerW::FormExcel::ExcelExportJob();
        job->tableSnapshot = snapshot;
        job->saveDirectory = excelSavePath;
        job->sessionStart = dataCollectionStartTime;
        job->sessionEnd = dataCollectionEndTime;
        job->clientPort = clientPort;
        job->clientIP = this->ClientIP;
        job->formGuid = this->FormGuid;
        job->formRef = gcnew System::WeakReference(this);
        job->enableButtonOnComplete = !isEmergency;

        ProjectServerW::FormExcel::EnqueueExport(job);

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

void ProjectServerW::DataForm::OnExcelExportCompleted(bool enableButtonOnComplete) {
    try {
        if (this == nullptr || this->IsDisposed || this->Disposing) {
            return;
        }

        System::Threading::Monitor::Enter(excelExportSync);
        try {
            excelExportInProgress = 0;
        }
        finally {
            System::Threading::Monitor::Exit(excelExportSync);
        }

        if (enableButtonOnComplete) {
            try {
                this->BeginInvoke(gcnew MethodInvoker(this, &DataForm::EnableButton));
                    }
                    catch (...) {}
            }
        }
        catch (...) {}
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
    //MessageBox::Show("Форма DataForm будет закрыта!");
    GlobalLogger::LogMessage("Форма DataForm будет закрыта!");

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

        // Сначала выносим поток из map. Если держать mutex во время ожидания, можно получить deadlock при завершении.
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

    // Закрытие форм не должно блокироваться на join потоков. В этом приложении DataForm живёт в своём UI-потоке
    // (ShowDialog внутри std::thread). Join из UI-пути закрытия может повесить модальный цикл.
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
    // Зачем: нужно, чтобы переводить относительное время устройства в ответах аудита команд в примерное реальное время.
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
    temperatures[0] = data.T[0] / 10.0;  // Температура дефростера слева (T_def_left)
    temperatures[1] = data.T[1] / 10.0;  // Температура дефростера справа (T_def_right)
    temperatures[2] = data.T[2] / 10.0;  // Температура дефростера по центру (T_def_center)
    temperatures[3] = data.T[3] / 10.0;  // Температура продукта слева (T_product_left)
    temperatures[4] = data.T[4] / 10.0;  // Температура продукта справа (T_product_right)
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

        // для тестовой отладки - принудительно устанавливаем Work=1
        // currentWorkBitState = true;

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

            // После запроса версии синхронизируем интервал измерений из настроек.
            // Почему: устройство могло быть перезапущено/перепрошито и вернуться к дефолту, а UI ожидает применение настроек при подключении.
            int intervalSeconds = 10;
            if (numericUpDownMeasurementInterval != nullptr && !numericUpDownMeasurementInterval->IsDisposed) {
                intervalSeconds = System::Decimal::ToInt32(numericUpDownMeasurementInterval->Value);
            }
            SendSetIntervalCommand(intervalSeconds);
        }

        // Если бит "Work" переходит из 0 в 1 (дефростер is ON)
        if (!workBitDetected && currentWorkBitState) {
            // Проверяем: это возврат после попытки остановки или новый запуск?
            if (workBitZeroTimerActive) {
                // Это отмена отложенной финализации: устройство снова в работе, дозапись после Work=0 не нужна.
                workBitZeroTimerActive = false;
                workStopFinalizeDelayElapsed = false;
                if (workStopFinalizeTimer != nullptr && workStopFinalizeTimer->Enabled) {
                    workStopFinalizeTimer->Stop();
                }
            } else {
                // Это новый запуск работы (не мигание)
                // Сохраняем время начала сессии: финальное имя файла включает и старт, и завершение.
                dataCollectionStartTime = now;
                dataCollectionEndTime = DateTime::MinValue;
                excelFileName = "WorkData_Port" + clientPort.ToString();
                GlobalLogger::LogMessage(ConvertToStdString(String::Format(
                    "Information: СТАРТ фиксации данных: {0} (Port {1})",
                    dataCollectionStartTime.ToString("yyyy-MM-dd HH:mm:ss"), clientPort)));
                // Сбрасываем флаги при новом запуске работы
                dataExportedToExcel = false;  // Новый цикл - данные еще не экспортированы
            }
            // Меняем состояние кнопок СТАРТ и СТОП
            buttonSTOPstate_TRUE();
        }

        // Если бит "Work" переходит из 1 в 0 (дефростер is OFF)
        if (workBitDetected && !currentWorkBitState) {
            // Начинаем 5-минутную дозапись после Work=0. Финализацию делаем только на приходе телеметрии.
            workBitZeroTimerActive = true;
            workStopFinalizeDelayElapsed = false;
            if (workStopFinalizeTimer != nullptr) {
                workStopFinalizeTimer->Stop();
                workStopFinalizeTimer->Start();
            }
            GlobalLogger::LogMessage("Information: Work=0, старт 5-минутной дозаписи перед финализацией");
        }

        // Отменяем автоперезапуск только по большому страховочному таймауту с момента начала автоперезапуска.
        // Зачем: команда STOP может не привести к полноценной остановке (или телеметрия может быть потеряна).
        // Готовность к START определяется по завершению 5-минутной дозаписи и последующему пакету телеметрии с Work=0.
        if (autoRestartPending && autoRestartStopIssuedTime != DateTime::MinValue) {
            TimeSpan waited = now.Subtract(autoRestartStopIssuedTime);
            const double safetyTimeoutMinutes = 20;
            if (waited.TotalMinutes >= safetyTimeoutMinutes) { 
                // ожидаем уже больше 20 минут после STOP, тогда отменяем автоперезапуск
                autoRestartPending = false;
                autoRestartStopIssuedTime = DateTime::MinValue;
                GlobalLogger::LogMessage(ConvertToStdString(String::Format(
                    "Warning: Автоперезапуск отменён: не дождались финализации после Work=0 за {0} минут после STOP",
                    safetyTimeoutMinutes)));
                if (Label_Commands != nullptr && !Label_Commands->IsDisposed) {
                    Label_Commands->Text = "[!] Автоперезапуск отменён: таймаут ожидания остановки";
                    Label_Commands->ForeColor = System::Drawing::Color::Orange;
                    GlobalLogger::LogMessage(ConvertToStdString(Label_Commands->Text));
                }
            }
        }

        // Финализация: только по приходу телеметрии, когда Work=0 и прошли 5 минут дозаписи.
        if (workBitZeroTimerActive && workStopFinalizeDelayElapsed && !currentWorkBitState) {
                workBitZeroTimerActive = false;
                workStopFinalizeDelayElapsed = false;
                
                // Сохраняем время окончания сбора данных
                dataCollectionEndTime = now;
                
            GlobalLogger::LogMessage(ConvertToStdString("Information: СТОП фиксации данных (Work=0, +5 минут дозаписи), запись в файл " + excelFileName));
                // Меняем состояние кнопок СТАРТ и СТОП
                buttonSTARTstate_TRUE();
                
                // Устанавливаем флаг, что данные экспортированы (чтобы не дублировать при закрытии формы)
                dataExportedToExcel = true;
                
                // Выполняем запись в Excel
                this->BeginInvoke(gcnew MethodInvoker(this, &DataForm::TriggerExcelExport));

            // Зачем: команды START/STOP синхронные (ждём ответ контроллера). Ставим START в UI message loop,
            // чтобы обработка телеметрии была короткой и чтобы экспорт в Excel мог стартовать параллельно.
                if (autoRestartPending) {
                    autoRestartPending = false;
                autoRestartStopIssuedTime = DateTime::MinValue;
                    this->BeginInvoke(gcnew MethodInvoker(this, &DataForm::ExecuteAutoRestartStart));
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
    // 2. Work = 0, но ещё идёт ожидание финального условия остановки (workBitZeroTimerActive = true)
    //    Это нужно для записи данных между Work->0 и Open->1
    if (workBitDetected || workBitZeroTimerActive)
    {
        table->Rows->Add(row);
        
        // Сбрасываем флаг экспорта, т.к. появились новые данные, которые ещё не записаны в Excel
        // Это важно для случая, когда после ручной записи продолжают поступать данные
        dataExportedToExcel = false;
    }
}


// Запись Excel вынесена в FormExcel.cpp

// Отложенная сборка мусора
void DataForm::DelayedGarbageCollection(Object^ state) {
    try {
        // Пауза перед сборкой мусора
        //MessageBox::Show("Файл Excel успешно записан!");
        GlobalLogger::LogMessage(ConvertToStdString("Файл Excel успешно записан!"));
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

void ProjectServerW::DataForm::OnWorkStopFinalizeTimerTick(Object^ sender, EventArgs^ e) {
    try {
        if (workStopFinalizeTimer != nullptr) {
            workStopFinalizeTimer->Stop();
        }

        // Зачем: таймер может отработать без новых пакетов телеметрии. Экспорт делаем только на телеметрии,
        // поэтому здесь лишь разрешаем финализацию на следующем пакете с Work=0.
        workStopFinalizeDelayElapsed = true;
        GlobalLogger::LogMessage("Information: 5 минут дозаписи после Work=0 истекли, ожидание телеметрии для финализации");
    }
    catch (Exception^ ex) {
        GlobalLogger::LogMessage(ConvertToStdString("Error: Exception in OnWorkStopFinalizeTimerTick: " + ex->ToString()));
    }
}

void ProjectServerW::DataForm::OnInactivityTimerTick(Object^ sender, EventArgs^ e) {
    try {
        if (inactivityCloseRequested) {
            return;
        }

        DateTime now = DateTime::Now;

        // 1) Режим ожидания переподключения "своего" клиента (по IP) после потери сокета.
        // Требование: если форма открыта и за 30 минут не пришёл клиент с тем же IP — закрыть форму,
        // чтобы новый клиент с другим IP получал новую форму, а "чужая" форма не зависала бесконечно.
        if (ClientSocket == INVALID_SOCKET && !String::IsNullOrEmpty(ClientIP)) {
            const long long ticks = System::Threading::Interlocked::Read(disconnectedSinceTicks);
            if (ticks != 0) {
                DateTime disconnectedSince(ticks);
                TimeSpan waited = now.Subtract(disconnectedSince);
                if (waited.TotalMinutes >= 30) {
                    GlobalLogger::LogMessage(ConvertToStdString(String::Format(
                        "Information: Нет переподключения клиента {0} {1:F1} минут. Закрываем ожидающую DataForm.",
                        ClientIP,
                        waited.TotalMinutes)));

                    // Если телеметрии не было — экспортировать нечего, просто закрываем.
                    if (!hasTelemetry) {
                        this->Close();
                        return;
                    }

                    // Иначе используем тот же путь, что и для таймаута неактивности: экспорт + закрытие.
                    // Сначала запускаем экспорт; затем помечаем как экспортированное и просим закрытие,
                    // чтобы не потерять экспорт из-за временного guard.
                    if (!StartExcelExportThread(true)) {
                        GlobalLogger::LogMessage("Warning: Экспорт не стартовал по таймауту ожидания переподключения; повторим на следующем тике.");
                        return;
                    }

                    inactivityCloseRequested = true;
                    dataExportedToExcel = true;
                    this->Close();
                    return;
                }
            }
        }

        // 2) Таймаут по отсутствию телеметрии (для активной ранее сессии).
        if (!hasTelemetry || lastTelemetryTime == DateTime::MinValue) {
            return;
        }

        TimeSpan idle = now.Subtract(lastTelemetryTime);
        if (idle.TotalMinutes < 30) {
            return;
        }

        // Критично: финализируем текущий файл сессии (не создаём EmergencyData-файлы).
        if (dataCollectionEndTime == DateTime::MinValue) {
            dataCollectionEndTime = now;
        }

        GlobalLogger::LogMessage(ConvertToStdString(String::Format(
            "Information: Нет телеметрии {0:F1} минут. Финализация и закрытие DataForm.",
            idle.TotalMinutes)));

        // Сначала запускаем экспорт; затем помечаем как экспортированное и просим закрытие, чтобы не потерять экспорт из-за временного guard.
        if (!StartExcelExportThread(true)) {
            GlobalLogger::LogMessage("Warning: Экспорт не стартовал по таймауту неактивности; повторим на следующем тике.");
            return;
        }

        inactivityCloseRequested = true;
        dataExportedToExcel = true;

        // Разрешаем новым сессиям создавать новую форму после таймаута.
        this->Close();
    }
    catch (Exception^ ex) {
        GlobalLogger::LogMessage(ConvertToStdString("Error: Exception in OnInactivityTimerTick: " + ex->ToString()));
    }
}

void ProjectServerW::DataForm::OnReconnectSendStartupCommands() {
    try {
        if (this == nullptr || this->IsDisposed || this->Disposing || !this->IsHandleCreated) {
            return;
        }

        if (ClientSocket == INVALID_SOCKET) {
            return;
        }

        // Защита: при бурстах переподключений не спамим одинаковыми командами на один и тот же сокет.
        // Важно: фиксируем сокет только после успешного выполнения обоих команд, чтобы разрешить повтор,
        // если устройство ещё не готово читать UART4.
        if (lastStartupCommandsSocket == ClientSocket) {
            return;
        }

        // Запрашиваем версию и настраиваем интервал измерений из UI-настроек.
        const bool okVersion = SendVersionRequest();

        int intervalSeconds = 10;
        if (numericUpDownMeasurementInterval != nullptr && !numericUpDownMeasurementInterval->IsDisposed) {
            intervalSeconds = System::Decimal::ToInt32(numericUpDownMeasurementInterval->Value);
        }
        const bool okInterval = SendSetIntervalCommand(intervalSeconds);

        if (okVersion && okInterval) {
            lastStartupCommandsSocket = ClientSocket;
            postResetInitPending = false;
        }
    }
    catch (Exception^ ex) {
        GlobalLogger::LogMessage(ConvertToStdString("Error: Exception in OnReconnectSendStartupCommands: " + ex->ToString()));
    }
    catch (...) {
        GlobalLogger::LogMessage("Error: Unknown exception in OnReconnectSendStartupCommands");
    }
}

void ProjectServerW::DataForm::SchedulePostResetInit() {
    if (this == nullptr || this->IsDisposed || this->Disposing) {
        return;
    }

    postResetInitPending = true;
    postResetInitAttempt = 0;
    postResetInitDeadline = DateTime::Now.AddSeconds(60);

    // Важно: даже если сокет после RESET не поменялся, мы должны повторно отправить версию/интервал.
    lastStartupCommandsSocket = INVALID_SOCKET;

    if (postResetInitTimer == nullptr) {
        postResetInitTimer = gcnew System::Windows::Forms::Timer();
        postResetInitTimer->Tick += gcnew EventHandler(this, &DataForm::OnPostResetInitTimerTick);
    }

    postResetInitTimer->Stop();
    postResetInitTimer->Interval = 7000; // 7 секунд на перезапуск контроллера до чтения UART4
    postResetInitTimer->Start();
}

void ProjectServerW::DataForm::OnPostResetInitTimerTick(System::Object^ sender, System::EventArgs^ e) {
    try {
        if (postResetInitTimer != nullptr) {
            postResetInitTimer->Stop();
        }

        if (!postResetInitPending) {
            return;
        }

        if (this == nullptr || this->IsDisposed || this->Disposing || !this->IsHandleCreated) {
            postResetInitPending = false;
            return;
        }

        DateTime now = DateTime::Now;
        if (postResetInitDeadline != DateTime::MinValue && now >= postResetInitDeadline) {
            postResetInitPending = false;
            return;
        }

        if (ClientSocket == INVALID_SOCKET) {
            // Ждём переподключение TCP (контроллер может перезагружаться через мост/АП).
            if (postResetInitTimer != nullptr) {
                postResetInitTimer->Interval = 2000;
                postResetInitTimer->Start();
            }
            return;
        }

        postResetInitAttempt++;

        const bool okVersion = SendVersionRequest();
        int intervalSeconds = 10;
        if (numericUpDownMeasurementInterval != nullptr && !numericUpDownMeasurementInterval->IsDisposed) {
            intervalSeconds = System::Decimal::ToInt32(numericUpDownMeasurementInterval->Value);
        }
        const bool okInterval = SendSetIntervalCommand(intervalSeconds);

        if (okVersion && okInterval) {
            postResetInitPending = false;
            lastStartupCommandsSocket = ClientSocket;
            return;
        }

        // Повторяем до дедлайна; контроллер может поднять UART позже, чем TCP.
        if (postResetInitTimer != nullptr) {
            postResetInitTimer->Interval = 5000;
            postResetInitTimer->Start();
        }
    }
    catch (Exception^ ex) {
        GlobalLogger::LogMessage(ConvertToStdString("Error: Exception in OnPostResetInitTimerTick: " + ex->ToString()));
    }
    catch (...) {
        GlobalLogger::LogMessage("Error: Unknown exception in OnPostResetInitTimerTick");
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
            // Зачем: переподключение уже логируется в SServer. Этот флаг добавляет связанное сообщение
            // "начало фиксации данных" на первом валидном пакете телеметрии нового TCP-соединения.
            reconnectFixationLogPending = true;
        }

        hasTelemetry = true;
        lastTelemetryTime = telemetryTime;
        lastTelemetrySocket = currentSocket;

        // Критично: даже если контроллер не переключил Work->1 (или мы стартовали "в середине" сессии), нужны стабильные времена сессии.
        if (dataCollectionStartTime == DateTime::MinValue) {
            dataCollectionStartTime = telemetryTime;
        }

        dataGridView->SuspendLayout();
        try {
            // Критично: операции Add и Snapshot должны быть взаимоисключающими.
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
        // Строка 1 сохраняется как "старый" каталог Excel для обратной совместимости.
        writer->WriteLine(textBoxExcelDirectory->Text);

        // Зачем: приложение часто разворачивают без полноценного инсталлятора; простой текстовый файл легко переносить/бэкапить.
        writer->WriteLine("AutoStartEnabled=" + (checkBoxAutoStart->Checked ? "1" : "0"));
        writer->WriteLine("AutoStartTime=" + dateTimePickerAutoStart->Value.ToString("HH:mm"));
        writer->WriteLine("AutoRestartEnabled=" + (checkBoxAutoRestart->Checked ? "1" : "0"));
        writer->WriteLine("AutoRestartTime=" + dateTimePickerAutoRestart->Value.ToString("HH:mm"));
        int intervalSeconds = 10;
        if (numericUpDownMeasurementInterval != nullptr && !numericUpDownMeasurementInterval->IsDisposed) {
            intervalSeconds = System::Decimal::ToInt32(numericUpDownMeasurementInterval->Value);
        }
        // Why: keep settings file portable and stable across locale changes.
        writer->WriteLine("MeasurementIntervalSeconds=" +
            intervalSeconds.ToString(System::Globalization::CultureInfo::InvariantCulture));

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
                else if (key->Equals("MeasurementIntervalSeconds", StringComparison::OrdinalIgnoreCase)) {
                    int parsed = 0;
                    if (Int32::TryParse(value, System::Globalization::NumberStyles::Integer,
                        System::Globalization::CultureInfo::InvariantCulture, parsed)) {
                        if (numericUpDownMeasurementInterval != nullptr && !numericUpDownMeasurementInterval->IsDisposed) {
                            int minVal = System::Decimal::ToInt32(numericUpDownMeasurementInterval->Minimum);
                            int maxVal = System::Decimal::ToInt32(numericUpDownMeasurementInterval->Maximum);
                            if (parsed < minVal) parsed = minVal;
                            if (parsed > maxVal) parsed = maxVal;
                            numericUpDownMeasurementInterval->Value = System::Decimal(parsed);
                        }
                    }
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

// Default table: Parameter -> (WarmUp, Plateau, Finish) from attached spec. Each Default goes into column matching Phase.
void ProjectServerW::DataForm::LoadDataGridView1Defaults() {
    if (dataGridView1 == nullptr || dataGridView1->IsDisposed) return;
    dataGridView1->Rows->Clear();
    cli::array<Object^>^ row = gcnew cli::array<Object^>(4);
    row[0] = "fishHotMax_C";         row[1] = "20";   row[2] = "20";   row[3] = "20";   dataGridView1->Rows->Add(row);
    row[0] = "supplySet_C";          row[1] = "30";   row[2] = "26";   row[3] = "22";   dataGridView1->Rows->Add(row);
    row[0] = "supplyMax_C";          row[1] = "35";   row[2] = "30";   row[3] = "26";   dataGridView1->Rows->Add(row);
    row[0] = "fishDeltaMax_C";       row[1] = "6";    row[2] = "5";    row[3] = "4";    dataGridView1->Rows->Add(row);
    row[0] = "fishHotRateMax_Cps";  row[1] = "0.02"; row[2] = "0.015"; row[3] = "0.01"; dataGridView1->Rows->Add(row);
    row[0] = "returnTargetRH_percent"; row[1] = "85"; row[2] = "92"; row[3] = "85"; dataGridView1->Rows->Add(row);
}

// Default table: Parameter2, Description, Value from attached config parameters.
void ProjectServerW::DataForm::LoadDataGridView2Defaults() {
    if (dataGridView2 == nullptr || dataGridView2->IsDisposed) return;
    dataGridView2->Rows->Clear();
    dataGridView2->Rows->Add("Sensor0 use in defrost", "Use sensor 0 in algorithm", "1");
    dataGridView2->Rows->Add("Sensor1 use in defrost", "Use sensor 1 in algorithm", "1");
    dataGridView2->Rows->Add("Sensor2 use in defrost", "Use sensor 2 in algorithm", "1");
    dataGridView2->Rows->Add("Sensor3 use in defrost", "Use sensor 3 in algorithm", "1");
    dataGridView2->Rows->Add("Sensor4 use in defrost", "Use sensor 4 in algorithm", "1");
    dataGridView2->Rows->Add("Sensor5 use in defrost", "Use sensor 5 in algorithm", "1");
    dataGridView2->Rows->Add("Sensor6 use in defrost", "Use sensor 6 in algorithm", "1");
    dataGridView2->Rows->Add("leftRightTrimGain", "Heater trim per degC difference", "0.08");
    dataGridView2->Rows->Add("leftRightTrimMaxEq", "Max trim B± heater equivalent", "0.6");
    dataGridView2->Rows->Add("wDeadband_kgkg", "Humidity deadband", "0.0008");
    dataGridView2->Rows->Add("outDamperTimer_s", "Damper open time", "10");
    dataGridView2->Rows->Add("outFanDelay_s", "Fan delay after damper open", "5");
    dataGridView2->Rows->Add("tenMinHold_s", "Heater min hold between switches", "10");
    dataGridView2->Rows->Add("injMinHold_s", "Injector min hold between switches", "5");
    dataGridView2->Rows->Add("outHold_s", "Damper and exhaust fan min hold", "15");
}

void ProjectServerW::DataForm::LoadDataGridView2FromFile() {
    if (dataGridView2 == nullptr || dataGridView2->IsDisposed) return;
    try {
        String^ appPath = System::IO::Path::GetDirectoryName(System::Windows::Forms::Application::ExecutablePath);
        String^ settingsPath = System::IO::Path::Combine(appPath, "ExcelSettings.txt");
        if (!System::IO::File::Exists(settingsPath)) {
            LoadDataGridView2Defaults();
            dataGridView2Dirty = false;
            return;
        }
        cli::array<String^>^ lines = System::IO::File::ReadAllLines(settingsPath);
        bool inSection = false;
        dataGridView2->Rows->Clear();
        for (int i = 0; i < lines->Length; i++) {
            String^ line = lines[i]->Trim();
            if (line->Equals("[DataGridView2]", StringComparison::OrdinalIgnoreCase)) {
                inSection = true;
                continue;
            }
            if (inSection) {
                if (line->StartsWith("[")) break;
                if (String::IsNullOrEmpty(line)) continue;
                cli::array<String^>^ parts = line->Split('\t');
                if (parts->Length >= 3) {
                    dataGridView2->Rows->Add(parts[0]->Trim(), parts[1]->Trim(), parts[2]->Trim());
                }
            }
        }
        if (dataGridView2->Rows->Count == 0) {
            LoadDataGridView2Defaults();
        }
    }
    catch (Exception^) {
        LoadDataGridView2Defaults();
    }
    dataGridView2Dirty = false;
}

void ProjectServerW::DataForm::SaveDataGridView2ToFile() {
    if (dataGridView2 == nullptr || dataGridView2->IsDisposed) return;
    try {
        String^ appPath = System::IO::Path::GetDirectoryName(System::Windows::Forms::Application::ExecutablePath);
        String^ settingsPath = System::IO::Path::Combine(appPath, "ExcelSettings.txt");
        cli::array<String^>^ allLines = System::IO::File::Exists(settingsPath) ? System::IO::File::ReadAllLines(settingsPath) : gcnew cli::array<String^>(0);
        System::Collections::Generic::List<String^>^ outLines = gcnew System::Collections::Generic::List<String^>();
        bool sectionReplaced = false;
        for (int i = 0; i < allLines->Length; i++) {
            String^ line = allLines[i];
            String^ trimmed = line->Trim();
            if (trimmed->Equals("[DataGridView2]", StringComparison::OrdinalIgnoreCase)) {
                outLines->Add("[DataGridView2]");
                for (int r = 0; r < dataGridView2->Rows->Count; r++) {
                    DataGridViewRow^ row = dataGridView2->Rows[r];
                    if (row->IsNewRow) continue;
                    Object^ v0 = row->Cells["Parameter2"]->Value;
                    Object^ v1 = row->Cells["Description"]->Value;
                    Object^ v2 = row->Cells["Value"]->Value;
                    String^ s0 = v0 != nullptr ? v0->ToString() : "";
                    String^ s1 = v1 != nullptr ? v1->ToString() : "";
                    String^ s2 = v2 != nullptr ? v2->ToString() : "";
                    outLines->Add(s0 + "\t" + s1 + "\t" + s2);
                }
                sectionReplaced = true;
                i++;
                while (i < allLines->Length && !allLines[i]->Trim()->StartsWith("[")) i++;
                i--;
                continue;
            }
            outLines->Add(line);
        }
        if (!sectionReplaced) {
            outLines->Add("[DataGridView2]");
            for (int r = 0; r < dataGridView2->Rows->Count; r++) {
                DataGridViewRow^ row = dataGridView2->Rows[r];
                if (row->IsNewRow) continue;
                Object^ v0 = row->Cells["Parameter2"]->Value;
                Object^ v1 = row->Cells["Description"]->Value;
                Object^ v2 = row->Cells["Value"]->Value;
                String^ s0 = v0 != nullptr ? v0->ToString() : "";
                String^ s1 = v1 != nullptr ? v1->ToString() : "";
                String^ s2 = v2 != nullptr ? v2->ToString() : "";
                outLines->Add(s0 + "\t" + s1 + "\t" + s2);
            }
        }
        System::IO::File::WriteAllLines(settingsPath, outLines->ToArray());
        dataGridView2Dirty = false;
    }
    catch (Exception^ ex) {
        MessageBox::Show("Не удалось сохранить данные таблицы: " + ex->Message);
    }
}

void ProjectServerW::DataForm::LoadDataGridView1FromFile() {
    if (dataGridView1 == nullptr || dataGridView1->IsDisposed) return;
    try {
        String^ appPath = System::IO::Path::GetDirectoryName(System::Windows::Forms::Application::ExecutablePath);
        String^ settingsPath = System::IO::Path::Combine(appPath, "ExcelSettings.txt");
        if (!System::IO::File::Exists(settingsPath)) {
            LoadDataGridView1Defaults();
            return;
        }
        cli::array<String^>^ lines = System::IO::File::ReadAllLines(settingsPath);
        bool inSection = false;
        dataGridView1->Rows->Clear();
        for (int i = 0; i < lines->Length; i++) {
            String^ line = lines[i]->Trim();
            if (line->Equals("[DataGridView1]", StringComparison::OrdinalIgnoreCase)) {
                inSection = true;
                continue;
            }
            if (inSection) {
                if (line->StartsWith("[")) break;
                if (String::IsNullOrEmpty(line)) continue;
                cli::array<String^>^ parts = line->Split('\t');
                if (parts->Length >= 4) {
                    dataGridView1->Rows->Add(parts[0]->Trim(), parts[1]->Trim(), parts[2]->Trim(), parts[3]->Trim());
                }
            }
        }
        if (dataGridView1->Rows->Count == 0) {
            LoadDataGridView1Defaults();
        }
    }
    catch (Exception^) {
        LoadDataGridView1Defaults();
    }
    dataGridView1Dirty = false;
}

void ProjectServerW::DataForm::SaveDataGridView1ToFile() {
    if (dataGridView1 == nullptr || dataGridView1->IsDisposed) return;
    try {
        String^ appPath = System::IO::Path::GetDirectoryName(System::Windows::Forms::Application::ExecutablePath);
        String^ settingsPath = System::IO::Path::Combine(appPath, "ExcelSettings.txt");
        cli::array<String^>^ allLines = System::IO::File::Exists(settingsPath) ? System::IO::File::ReadAllLines(settingsPath) : gcnew cli::array<String^>(0);
        System::Collections::Generic::List<String^>^ outLines = gcnew System::Collections::Generic::List<String^>();
        bool sectionReplaced = false;
        for (int i = 0; i < allLines->Length; i++) {
            String^ line = allLines[i];
            String^ trimmed = line->Trim();
            if (trimmed->Equals("[DataGridView1]", StringComparison::OrdinalIgnoreCase)) {
                outLines->Add("[DataGridView1]");
                for (int r = 0; r < dataGridView1->Rows->Count; r++) {
                    DataGridViewRow^ row = dataGridView1->Rows[r];
                    if (row->IsNewRow) continue;
                    Object^ v0 = row->Cells["Parameter"]->Value;
                    Object^ v1 = row->Cells["WarmUP"]->Value;
                    Object^ v2 = row->Cells["Plateau"]->Value;
                    Object^ v3 = row->Cells["Finish"]->Value;
                    String^ s0 = v0 != nullptr ? v0->ToString() : "";
                    String^ s1 = v1 != nullptr ? v1->ToString() : "";
                    String^ s2 = v2 != nullptr ? v2->ToString() : "";
                    String^ s3 = v3 != nullptr ? v3->ToString() : "";
                    outLines->Add(s0 + "\t" + s1 + "\t" + s2 + "\t" + s3);
                }
                sectionReplaced = true;
                i++;
                while (i < allLines->Length && !allLines[i]->Trim()->StartsWith("[")) i++;
                i--;
                continue;
            }
            outLines->Add(line);
        }
        if (!sectionReplaced) {
            outLines->Add("[DataGridView1]");
            for (int r = 0; r < dataGridView1->Rows->Count; r++) {
                DataGridViewRow^ row = dataGridView1->Rows[r];
                if (row->IsNewRow) continue;
                Object^ v0 = row->Cells["Parameter"]->Value;
                Object^ v1 = row->Cells["WarmUP"]->Value;
                Object^ v2 = row->Cells["Plateau"]->Value;
                Object^ v3 = row->Cells["Finish"]->Value;
                String^ s0 = v0 != nullptr ? v0->ToString() : "";
                String^ s1 = v1 != nullptr ? v1->ToString() : "";
                String^ s2 = v2 != nullptr ? v2->ToString() : "";
                String^ s3 = v3 != nullptr ? v3->ToString() : "";
                outLines->Add(s0 + "\t" + s1 + "\t" + s2 + "\t" + s3);
            }
        }
        System::IO::File::WriteAllLines(settingsPath, outLines->ToArray());
        dataGridView1Dirty = false;
    }
    catch (Exception^ ex) {
        MessageBox::Show("Не удалось сохранить данные таблицы: " + ex->Message);
    }
}

// On entering tabPage3: load both grids from file. On leaving: ask to save if either grid was changed.
System::Void ProjectServerW::DataForm::tabControl1_SelectedIndexChanged(System::Object^ sender, System::EventArgs^ e) {
    TabControl^ tc = safe_cast<TabControl^>(sender);
    if (tabControl1PrevTab == tabPage3 && tc->SelectedTab != tabPage3) {
        if (dataGridView1Dirty || dataGridView2Dirty) {
            System::Windows::Forms::DialogResult dr = MessageBox::Show(
                "Сохранить изменения в ExcelSettings.txt?",
                "Параметры",
                MessageBoxButtons::YesNoCancel,
                MessageBoxIcon::Question);
            if (dr == System::Windows::Forms::DialogResult::Yes) {
                SaveDataGridView1ToFile();
                SaveDataGridView2ToFile();
            }
            else if (dr == System::Windows::Forms::DialogResult::Cancel) {
                tc->SelectedTab = tabPage3;
                tabControl1PrevTab = tabPage3;
                return;
            }
        }
        dataGridView1Dirty = false;
        dataGridView2Dirty = false;
    }
    if (tc->SelectedTab == tabPage3) {
        LoadDataGridView1FromFile();
        LoadDataGridView2FromFile();
    }
    tabControl1PrevTab = tc->SelectedTab;
}

System::Void ProjectServerW::DataForm::dataGridView1_CellValueChanged(System::Object^ sender, System::Windows::Forms::DataGridViewCellEventArgs^ e) {
    if (e->RowIndex >= 0 && !dataGridView1->Rows[e->RowIndex]->IsNewRow)
        dataGridView1Dirty = true;
}

System::Void ProjectServerW::DataForm::dataGridView1_RowChanged(System::Object^ sender, System::Windows::Forms::DataGridViewRowEventArgs^ e) {
    dataGridView1Dirty = true;
}

System::Void ProjectServerW::DataForm::dataGridView2_CellValueChanged(System::Object^ sender, System::Windows::Forms::DataGridViewCellEventArgs^ e) {
    if (e->RowIndex >= 0 && dataGridView2 != nullptr && !dataGridView2->IsDisposed && e->RowIndex < dataGridView2->Rows->Count && !dataGridView2->Rows[e->RowIndex]->IsNewRow)
        dataGridView2Dirty = true;
}

System::Void ProjectServerW::DataForm::dataGridView2_RowChanged(System::Object^ sender, System::Windows::Forms::DataGridViewRowEventArgs^ e) {
    dataGridView2Dirty = true;
}

System::Void ProjectServerW::DataForm::buttonLoadFromFile_Click(System::Object^ sender, System::EventArgs^ e) {
    LoadDataGridView1FromFile();
    LoadDataGridView2FromFile();
}

System::Void ProjectServerW::DataForm::buttonSaveToFile_Click(System::Object^ sender, System::EventArgs^ e) {
    SaveDataGridView1ToFile();
    SaveDataGridView2ToFile();
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
        "DBlk",     // Сигнал «Разблокировать ворота»
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
        "_Dbl",     // Ворота разблокировать
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
                // Данные ещё не экспортированы. Если закрываемся из-за таймаута неактивности, экспортируем в текущий файл сессии.
                // Имена файлов вида "EmergencyData" создавать не должны.
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
                GlobalLogger::LogMessage("Information: Финализация данных в Excel при закрытии формы... " + ConvertToStdString(excelFileName));

                // Запускаем фоновый экспорт без обращения к UI-контролам — форма может закрыться сразу.
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

    // Критично: форма может закрыться сама (например, по таймауту неактивности). Владеющий native-поток должен быть корректно завершён/удалён.
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

    if (inactivityTimer != nullptr) {
        try {
            inactivityTimer->Stop();
            inactivityTimer->Tick -= gcnew EventHandler(this, &DataForm::OnInactivityTimerTick);
        }
        catch (...) {}
        inactivityTimer = nullptr;
    }

    if (workStopFinalizeTimer != nullptr) {
        try {
            workStopFinalizeTimer->Stop();
            workStopFinalizeTimer->Tick -= gcnew EventHandler(this, &DataForm::OnWorkStopFinalizeTimerTick);
        }
        catch (...) {}
        workStopFinalizeTimer = nullptr;
    }
}

// Реализация методов управления дефростером вынесена в CommandsDefroster.cpp

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

//  ===========================================================================
// Реализация методов записи в Excel (отложенный запуск)
//  ===========================================================================

// Триггер автоматического экспорта в Excel
void ProjectServerW::DataForm::TriggerExcelExport() {
    if (this->InvokeRequired) {
        try {
            this->BeginInvoke(gcnew MethodInvoker(this, &DataForm::TriggerExcelExport));
        }
        catch (Exception^ ex) {
            GlobalLogger::LogMessage("Error invoking TriggerExcelExport: " + ConvertToStdString(ex->Message));
        }
        return;
    }
    
    try {
        StartExcelExportThread(false);
    }
    catch (Exception^ ex) {
        GlobalLogger::LogMessage("Error in TriggerExcelExport: " + ConvertToStdString(ex->Message));
    }
}

//  ===========================================================================
// Реализация методов автозапуска по времени
//  ===========================================================================

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
            GlobalLogger::LogMessage(ConvertToStdString(Label_Commands->Text));
            
            CommandResponse startResp{};
            CommandAckResult startResult = SendControlCommandWithAck(CmdProgControl::START, "START", 2000, 2, startResp);

            if (startResult == CommandAckResult::NoResponse) {
                // Зачем: контроллер мог стартовать, но ACK потерялся; проверяем по биту Work в телеметрии.
                GlobalLogger::LogMessage("Warning: Автозапуск: START без подтверждения, проверка по биту Work");
                Label_Commands->Text = "[АВТОЗАПУСК] START без подтверждения, проверка по Work...";
                Label_Commands->ForeColor = System::Drawing::Color::Orange;
                GlobalLogger::LogMessage(ConvertToStdString(Label_Commands->Text));
            }
            else if (startResult != CommandAckResult::Ok) {
                GlobalLogger::LogMessage("Warning: Автозапуск: START не выполнен (ошибка ответа/отправки)");
                Label_Commands->Text = "[!] Автозапуск: START не выполнен";
                Label_Commands->ForeColor = System::Drawing::Color::Orange;
                GlobalLogger::LogMessage(ConvertToStdString(Label_Commands->Text));
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
            GlobalLogger::LogMessage(ConvertToStdString(Label_Commands->Text));
            
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

//  ===========================================================================
// Реализация методов автоперезапуска по времени
//  ===========================================================================

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

    // Если мы ещё не знаем состояние устройства (нет телеметрии), не "съедаем" расписание.
    if (!buttonSTOP->Enabled && !buttonSTART->Enabled) {
        GlobalLogger::LogMessage("Warning: Автоперезапуск: состояние устройства неизвестно (нет телеметрии), ожидание...");
        return;
    }

    labelAutoRestart->ForeColor = System::Drawing::Color::Blue;

    // Если устройство сейчас работает (доступен STOP) — отправляем STOP и ждём стабильного Work=0.
    if (buttonSTOP->Enabled) {
        Label_Commands->Text = "[АВТОПЕРЕЗАПУСК] Отправка команды STOP...";
        Label_Commands->ForeColor = System::Drawing::Color::Blue;
        GlobalLogger::LogMessage(ConvertToStdString(Label_Commands->Text));

        CommandResponse stopResp{};
        CommandAckResult stopResult = SendControlCommandWithAck(CmdProgControl::STOP, "STOP", 2000, 2, stopResp);

        if (stopResult == CommandAckResult::Ok) {
            autoRestartPending = true;
            autoRestartStopIssuedTime = DateTime::Now;
            Label_Commands->Text = "[АВТОПЕРЕЗАПУСК] STOP подтверждён, ожидание окончания работы...";
            Label_Commands->ForeColor = System::Drawing::Color::Blue;
            GlobalLogger::LogMessage(ConvertToStdString(Label_Commands->Text));
        }
        else if (stopResult == CommandAckResult::NoResponse) {
            // Зачем: контроллер мог выполнить STOP, но ACK потерялся; продолжаем сценарий и проверяем по биту Work.
            autoRestartPending = true;
            autoRestartStopIssuedTime = DateTime::Now;
            GlobalLogger::LogMessage("Warning: Автоперезапуск: STOP без подтверждения, ожидание по биту Work");
            Label_Commands->Text = "[АВТОПЕРЕЗАПУСК] STOP без подтверждения, ожидание по Work...";
            Label_Commands->ForeColor = System::Drawing::Color::Orange;
            GlobalLogger::LogMessage(ConvertToStdString(Label_Commands->Text));
        }
        else {
            autoRestartPending = false;
            autoRestartStopIssuedTime = DateTime::MinValue;
            GlobalLogger::LogMessage("Warning: Автоперезапуск: STOP не выполнен (ошибка ответа/отправки)");
            Label_Commands->Text = "[!] Автоперезапуск: STOP не выполнен";
            Label_Commands->ForeColor = System::Drawing::Color::Orange;
            GlobalLogger::LogMessage(ConvertToStdString(Label_Commands->Text));
        }

            autoRestartInternalUncheck = true;
        checkBoxAutoRestart->Checked = false; // одноразово, UX такой же как у AutoStart
    }
    else if (buttonSTART->Enabled) {
        // Устройство уже остановлено — ведём себя как при плановом старте.
        Label_Commands->Text = "[АВТОПЕРЕЗАПУСК] Устройство уже остановлено, отправка START...";
        Label_Commands->ForeColor = System::Drawing::Color::Blue;
        GlobalLogger::LogMessage(ConvertToStdString(Label_Commands->Text));

        CommandResponse startResp{};
        CommandAckResult startResult = SendControlCommandWithAck(CmdProgControl::START, "START", 2000, 2, startResp);

        if (startResult == CommandAckResult::Ok) {
            Label_Commands->Text = "[АВТОПЕРЕЗАПУСК] START подтверждён";
            Label_Commands->ForeColor = System::Drawing::Color::Blue;
            GlobalLogger::LogMessage(ConvertToStdString(Label_Commands->Text));
        }
        else if (startResult == CommandAckResult::NoResponse) {
            // Зачем: продолжаем реагировать на отсутствие ACK, но разрешаем телеметрии Work подтвердить реальное состояние.
            GlobalLogger::LogMessage("Warning: Автоперезапуск: START без подтверждения, ожидание по биту Work");
            Label_Commands->Text = "[АВТОПЕРЕЗАПУСК] START без подтверждения, проверка по Work...";
            Label_Commands->ForeColor = System::Drawing::Color::Orange;
            GlobalLogger::LogMessage(ConvertToStdString(Label_Commands->Text));
        }
        else {
            GlobalLogger::LogMessage("Warning: Автоперезапуск: START не выполнен (ошибка ответа/отправки)");
            Label_Commands->Text = "[!] Автоперезапуск: START не выполнен";
            Label_Commands->ForeColor = System::Drawing::Color::Orange;
            GlobalLogger::LogMessage(ConvertToStdString(Label_Commands->Text));
        }
        autoRestartInternalUncheck = true;
        checkBoxAutoRestart->Checked = false;
    }

    // Возвращаем цвет через короткоживущий таймер, чтобы не оставлять "взведённый" вид навсегда.
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
        GlobalLogger::LogMessage(ConvertToStdString(Label_Commands->Text));
        SendStartCommand();
    }
    catch (Exception^ ex) {
        GlobalLogger::LogMessage(ConvertToStdString("Error: Exception in ExecuteAutoRestartStart: " + ex->ToString()));
    }
}

// Реализация команд управления дефростером вынесена в CommandsDefroster.cpp

