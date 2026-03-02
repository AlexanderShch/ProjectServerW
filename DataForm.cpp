#include "DataForm.h"
#include "Chart.h"
#include "FormExcel.h"
#include "Commands.h"               // Command, DefrostParam Рё РѕР±РјРµРЅ СЃ РєРѕРЅС‚СЂРѕР»Р»РµСЂРѕРј
#include "PacketQueueProcessor.h"   // per-socket РѕС‡РµСЂРµРґСЊ РєРѕРјР°РЅРґ (ACK РѕС‚ РєРѕРЅС‚СЂРѕР»Р»РµСЂР° vs РѕС‚РІРµС‚С‹ РІ UI)
#include <objbase.h>                // РґР»СЏ CoCreateGuid вЂ” РіРµРЅРµСЂР°С†РёСЏ СѓРЅРёРєР°Р»СЊРЅРѕРіРѕ РёРґРµРЅС‚РёС„РёРєР°С‚РѕСЂР° РѕРєРЅР°
#include <string>
#include <vcclr.h>  // РґР»СЏ gcnew
#include <msclr/marshal_cppstd.h>

// РћР±С‰Р°СЏ Р»РѕРіРёРєР° РґР»СЏ РІСЃРµС… РѕРєРѕРЅ РІ Process
using namespace System::Diagnostics;
using namespace ProjectServerW; // РїСЂРѕСЃС‚СЂР°РЅСЃС‚РІРѕ РёРјС‘РЅ РїСЂРёР»РѕР¶РµРЅРёСЏ
using namespace Microsoft::Office::Interop::Excel;

std::map<std::wstring, gcroot<DataForm^>> formData_Map; // РіР»РѕР±Р°Р»СЊРЅР°СЏ РєР°СЂС‚Р° formData_Map

// Структура пакета по контракту (совпадает с контроллером)
// На сервере использовать __attribute__((packed)), на Win — #pragma pack(1)
#pragma pack(push, 1)
typedef struct   // Формат пакета (как на STM32)
{
    uint8_t DataType;           // тип пакета телеметрии (0x00 или служебный)
    uint8_t Len;                // длина полезной нагрузки (байты после Len и до CRC), не включая CRC
    uint16_t Time;				// время в секундах с включения
    uint8_t SensorQuantity;		// количество датчиков
    uint8_t SensorType[SQ];		// тип датчика
    uint8_t Active[SQ];			// маска активных
    short T[SQ];				// температура 1 группы (испаритель)
    short H[SQ];				// влажность 2 группы (относительная)
    uint16_t CRC_SUM;			// контрольная сумма
} MSGQUEUE_OBJ_t;
#pragma pack(pop)

// Совпадает с форматом пакета в типе MSGQUEUE_OBJ_t на STM32
void ProjectServerW::DataForm::ParseBuffer(const char* buffer, size_t size) {

    MSGQUEUE_OBJ_t data;

    if (size < sizeof(data)) {
        return; // недостаточно байт для пакета
    }
    memcpy(&data, buffer, sizeof(data));
}

System::Void ProjectServerW::DataForm::выходToolStripMenuItem_Click(System::Object^ sender, System::EventArgs^ e)
{
    // Ищем GUID текущей формы в formData_Map
    std::wstring currentFormGuid;
    bool formFound = false;

    for (auto it = formData_Map.begin(); it != formData_Map.end(); ++it) {
        // Получаем указатель формы из gcroot
        ProjectServerW::DataForm^ formPtr = it->second;

        // Проверяем, это наша форма
        if (formPtr == this) {
            currentFormGuid = it->first;
            formFound = true;
            break;
        }
    }

    try {
        if (formFound) {
            // Закрываем сокет формы
            // Получаем форму по GUID
            DataForm^ form = GetFormByGuid(currentFormGuid);

            if (form != nullptr && form->ClientSocket != INVALID_SOCKET) {
                // Закрываем сокет формы
                closesocket(form->ClientSocket);
                form->ClientSocket = INVALID_SOCKET;
                GlobalLogger::LogMessage("Information: Форма успешно добавлена в список по текущему GUID");
            }
        }
    }
    catch (Exception^ ex) {
        MessageBox::Show("Не удалось сохранить данные: " + ex->Message);
        GlobalLogger::LogMessage(ConvertToStdString("Не удалось сохранить данные: " + ex->Message));
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
        // Замечание: блок не по потоку — захват монитора в контексте этого потока (UI поток захватывает монитор по кнопке).
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

        // Отключаем кнопку только для не-аварийного вызова, чтобы при disposed-UI не падать при обращении.
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
    // Выбор папки для сохранения файлов
    FolderBrowserDialog^ folderDialog = gcnew FolderBrowserDialog();

    // Описание диалога
    folderDialog->Description = "Р’С‹Р±РµСЂРёС‚Рµ РїР°РїРєСѓ РґР»СЏ СЃРѕС…СЂР°РЅРµРЅРёСЏ Excel С„Р°Р№Р»РѕРІ";
    folderDialog->ShowNewFolderButton = true;

    // Предзаполнение текущей папки по умолчанию
    if (textBoxExcelDirectory->Text->Length > 0) {
        folderDialog->SelectedPath = textBoxExcelDirectory->Text;
    }

    // Открываем диалог и при OK сохраняем путь
    if (folderDialog->ShowDialog() == System::Windows::Forms::DialogResult::OK) {
        // Записываем выбранную папку в поле
        textBoxExcelDirectory->Text = folderDialog->SelectedPath;

        // Сохраняем выбранный путь
        excelSavePath = folderDialog->SelectedPath;
        // Сохраняем настройки
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
        // Преобразуем GUID в строку
        wchar_t guidString[40] = { 0 };
        int simb_N = StringFromGUID2(guid, guidString, 40);

        String^ formId = gcnew String(guidString);
        form->FormGuid = formId;

        // Сохраняем форму в карту
        formData_Map[guidString] = form;

        // Кладём GUID в очередь сообщений
        {
            std::lock_guard<std::mutex> lock(mtx);
            messageQueue.push(guidString);
        }
        cv.notify_one();
    }
    form->ShowDialog();

 }

// Закрытие формы по GUID (вызывается при закрытии окна)
void ProjectServerW::DataForm::CloseForm(const std::wstring& guid) {
    // Получаем форму
    ProjectServerW::DataForm^ form = ProjectServerW::DataForm::GetFormByGuid(guid);
    //MessageBox::Show("Ошибка DataForm уже уничтожена!");
    GlobalLogger::LogMessage("Ошибка DataForm уже уничтожена!");

    if (form != nullptr) {
        // Важно: удаляем форму из карты до закрытия, иначе
        // при повторном вызове форма уже уничтожена и обращение к ней даст исключение
        auto it = formData_Map.find(guid);
        if (it != formData_Map.end()) {
            formData_Map.erase(it);
            GlobalLogger::LogMessage("Information: Форма закрыта по кнопке, данные не сохранились");
        }
        
        // Закрываем потокобезопасно, через Invoke
        if (form->InvokeRequired) {
            // Закрываем форму через Invoke
            form->Invoke(gcnew System::Action(form, &ProjectServerW::DataForm::Close));
        }
        else {
            form->Close();
        }
    }
}
        
// Получение формы по GUID
DataForm^ ProjectServerW::DataForm::GetFormByGuid(const std::wstring& guid) {
    auto it = formData_Map.find(guid);
    if (it != formData_Map.end()) {
        DataForm^ form = it->second;

        // Проверяем, что форма жива и не уничтожается
        if (form != nullptr && !form->IsDisposed && !form->Disposing) {
            return form;
        }
        else {
            // Удаляем мёртвую форму из карты, возвращаем nullptr
            formData_Map.erase(it);
            return nullptr;
        }

    }
    return nullptr;
}

// Поиск формы по IP-адресу клиента
std::wstring ProjectServerW::DataForm::FindFormByClientIP(String^ clientIP) {
    // Перебор всех форм в карте
    for (auto it = formData_Map.begin(); it != formData_Map.end(); ++it) {
        DataForm^ form = it->second;
        
        // Проверяем, что форма жива и не уничтожается
        if (form != nullptr && !form->IsDisposed && !form->Disposing) {
            // Сравниваем IP-адрес
            if (form->ClientIP != nullptr && form->ClientIP->Equals(clientIP)) {
                // Нашли форму по IP
                return it->first; // Возвращаем GUID
            }
        }
    }
    // Не нашли форму по IP
    return std::wstring();
}

// Сохранение потока по guid в общую map
void ThreadStorage::StoreThread(const std::wstring& guid, std::thread& thread) {
    std::lock_guard<std::mutex> lock(GetMutex());
    GetThreadMap()[guid] = std::move(thread);
}

// Остановка потока по guid
void ThreadStorage::StopThread(const std::wstring& guid)
{
    std::thread threadToStop;
    {
        std::lock_guard<std::mutex> lock(GetMutex());

        // Достаём поток из map. Не держать mutex при join ниже, иначе возможен deadlock при обращении.
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

    // Не делаем join здесь — блокировка по join потока. Т.к. создание DataForm идёт в том же UI-потоке
    // (ShowDialog из std::thread). Join из UI-потока приведёт к взаимной блокировке потока.
    threadToStop.detach();
}

// Единственный экземпляр общей карты map для потоков
std::map<std::wstring, std::thread>& ThreadStorage::GetThreadMap() {
    static std::map<std::wstring, std::thread> threadMap;
    return threadMap;
}

// Единственный экземпляр общего Mutex для потоков
std::mutex& ThreadStorage::GetMutex() {
    static std::mutex mtx;
    return mtx;
}

// Вспомогательная функция конвертации строки в std::string (для логов и т.д.), т.к. managed-строка живёт в куче.
static String^ bufferToHex(const char* buffer, size_t length) {
    // Используем std::stringstream для конвертации managed-строки в std::string с кодировкой.
    std::stringstream ss;
    ss << std::hex << std::setfill('0');
    for (int i = 0; i < length; ++i) {
        ss << std::setw(2) << static_cast<unsigned int>(static_cast<unsigned char>(buffer[i])) << " ";
    }
    return gcnew String(ss.str().c_str());
}

/*
*********** Блок таблицы данных *******************
*/

// 1. Инициализация таблицы данных
void ProjectServerW::DataForm::InitializeDataTable() {
    // Создаём таблицу данных
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

    // Последний столбец без H (только тип/активность)
    dataTable->Columns->Add("Typ" + (SQ - 1), uint8_t::typeid);
    dataTable->Columns->Add("Act" + (SQ - 1), uint8_t::typeid);

    cli::array<cli::array<String^>^>^ bitNames = GetBitFieldNames();
    if (bitNames[0] != nullptr) {
        // Добавляем в таблицу колонки битовых полей первой группы
        for (uint8_t i = 0; i < 16; i++)
        {
            dataTable->Columns->Add(bitNames[0][i], uint8_t::typeid);
        }
        // Добавляем в таблицу колонки битовых полей второй группы
        for (uint8_t i = 0; i < 16; i++)
        {
            dataTable->Columns->Add(bitNames[1][i], uint8_t::typeid);
        }
    }

    dataGridView->DataSource = dataTable;
    GlobalLogger::LogMessage("Information: Получены новые данные");
}

// 2. Добавление строки
void ProjectServerW::DataForm::AddDataToTable(const char* buffer, size_t size, System::Data::DataTable^ table) {
    
    MSGQUEUE_OBJ_t data;
    DateTime now = DateTime::Now;   // Текущее время сервера
    String^ data_String{};

    if (size < sizeof(data)) {
        return; // недостаточно байт для пакета
    }
    memcpy(&data, buffer, sizeof(data));

    data_String = bufferToHex(buffer, size);
    // Выводим содержимое пакета data_String в поле Label_Data
    DataForm::SetData_TextValue("Текущие данные: " + data_String);

    // Новая строка таблицы
    DataRow^ row = table->NewRow();
    row["RealTime"] = now.ToString("HH:mm:ss");
    row["Time"] = data.Time;
    // Замечание: время, когда пришла телеметрия с устройства — в секундах с включения устройства.
    lastTelemetryDeviceSeconds = data.Time;
    row["SQ"] = data.SensorQuantity;
    for (uint8_t i = 0; i < (SQ - 1); i++)
    {
        row["Typ" + i] = data.SensorType[i];
        row["Act" + i] = data.Active[i];
        row["T" + i] = data.T[i] / 10.0;
        row["H" + i] = data.H[i] / 10.0;
    }
    // Температуры датчиков в градусах Цельсия
    cli::array<double>^ temperatures = gcnew cli::array<double>(5);
    temperatures[0] = data.T[0] / 10.0;  // Температура испарителя слева (T_def_left)
    temperatures[1] = data.T[1] / 10.0;  // Температура испарителя справа (T_def_right)
    temperatures[2] = data.T[2] / 10.0;  // Температура испарителя по центру (T_def_center)
    temperatures[3] = data.T[3] / 10.0;  // Температура продукта слева (T_product_left)
    temperatures[4] = data.T[4] / 10.0;  // Температура продукта справа (T_product_right)
    UpdateAllTemperatureValues(temperatures);

    cli::array<cli::array<String^>^>^ bitNames = GetBitFieldNames();

    // Индекс бита имени "Work" в массиве bitNames[0]
    int workBitIndex = -1;
    // Ищем индекс имени "Work" в массиве
    for (int i = 0; i < bitNames[0]->Length; i++) {
        if (bitNames[0][i] == "Work") {
            workBitIndex = i;
            break;
        }
    }

    uint16_t bitField;
    bitField = data.T[SQ - 1];
    if (workBitIndex != -1) {
        // Текущее состояние бита "Work"
        bool currentWorkBitState = (bitField & (1 << workBitIndex)) != 0;

        // Для отладки — принудительно считаем Work=1
        // currentWorkBitState = true;

        if (reconnectFixationLogPending) {
            reconnectFixationLogPending = false;
            GlobalLogger::LogMessage(ConvertToStdString(String::Format(
                "Information: Устройство подключено к порту (Port {1}), GUID: {0}",
                (this->ClientIP != nullptr ? this->ClientIP : ""),
                clientPort)));
        }

        // ===== Синхронизация состояния кнопок при первом приходе данных =====
        if (!firstDataReceived) {
            // При первом пакете данных — синхронизируем кнопки с состоянием устройства
            if (currentWorkBitState) {
                // Если WORK включён — показываем STOP
                buttonSTOPstate_TRUE();
                GlobalLogger::LogMessage("Information: Получен бит Work - флаг WORK установлен, отправлен STOP");
            } else {
                // Если WORK не включён — показываем START
                buttonSTARTstate_TRUE();
                GlobalLogger::LogMessage("Information: Получен бит Work - флаг WORK не установлен, отправлен START");
            }
            firstDataReceived = true;
            
            // Отправляем запрос версии и интервала после синхронизации состояния кнопок
            SendVersionRequest();

            // Отправляем текущий интервал измерения на устройство.
            // Замечание: интервал задаётся пользователем в настройках формы, в UI отображается актуальное значение для отладки.
            int intervalSeconds = 10;
            if (numericUpDownMeasurementInterval != nullptr && !numericUpDownMeasurementInterval->IsDisposed) {
                intervalSeconds = System::Decimal::ToInt32(numericUpDownMeasurementInterval->Value);
            }
            SendSetIntervalCommand(intervalSeconds);
        }

        // Когда бит "Work" переходит с 0 на 1 (устройство is ON)
        if (!workBitDetected && currentWorkBitState) {
            // Отмена: если был запланирован переход в состояние "остановлено"
            if (workBitZeroTimerActive) {
                // Если устройство включилось снова: отменяем таймер, сбрасываем бит Work=0 в ожидание.
                workBitZeroTimerActive = false;
                workStopFinalizeDelayElapsed = false;
                if (workStopFinalizeTimer != nullptr && workStopFinalizeTimer->Enabled) {
                    workStopFinalizeTimer->Stop();
                }
            } else {
                // Если это новый цикл работы (не переподключение)
                // Запоминаем время начала цикла: момент когда устройство включилось в работу, в секундах.
                dataCollectionStartTime = now;
                dataCollectionEndTime = DateTime::MinValue;
                excelFileName = "WorkData_Port" + clientPort.ToString();
                GlobalLogger::LogMessage(ConvertToStdString(String::Format(
                    "Information: Устройство отключено: {0} (Port {1})",
                    dataCollectionStartTime.ToString("yyyy-MM-dd HH:mm:ss"), clientPort)));
                // Сбрасываем флаг — данные ещё не экспортированы
                dataExportedToExcel = false;  // при экспорте — сбросим после выгрузки
            }
            // Показываем кнопку STOP в нажатом виде
            buttonSTOPstate_TRUE();
        }

        // Когда бит "Work" переходит с 1 на 0 (устройство is OFF)
        if (workBitDetected && !currentWorkBitState) {
            // Запуск 5-секундного таймера после бита Work=0. Ожидаем стабилизации перед экспортом.
            workBitZeroTimerActive = true;
            workStopFinalizeDelayElapsed = false;
            if (workStopFinalizeTimer != nullptr) {
                workStopFinalizeTimer->Stop();
                workStopFinalizeTimer->Start();
            }
            GlobalLogger::LogMessage("Information: Work=0, запущен 5-секундный таймер перед экспортом");
        }

        // Обработка автоперезапуска: ждём подтверждения от устройства о выполнении команды в текущем цикле дефроста.
        // Замечание: кнопка STOP отправлена на устройство в текущем цикле (при нажатии пользователь нажал кнопку).
        // Повторный на START отправится после ожидания 5-секундного таймера и экспорта данных в Excel с Work=0.
        if (autoRestartPending && autoRestartStopIssuedTime != DateTime::MinValue) {
            TimeSpan waited = now.Subtract(autoRestartStopIssuedTime);
            const double safetyTimeoutMinutes = 20;
            if (waited.TotalMinutes >= safetyTimeoutMinutes) { 
                // Таймаут: прошло 20 минут после STOP, отменяем автоперезапуск
                autoRestartPending = false;
                autoRestartStopIssuedTime = DateTime::MinValue;
                GlobalLogger::LogMessage(ConvertToStdString(String::Format(
                    "Warning: Автоперезапуск отменён: не получено подтверждение бита Work=0 за {0} минут после STOP",
                    safetyTimeoutMinutes)));
                if (Label_Commands != nullptr && !Label_Commands->IsDisposed) {
                    Label_Commands->Text = "[!] Неожиданное состояние: проверьте подключение устройства";
                    Label_Commands->ForeColor = System::Drawing::Color::Orange;
                    GlobalLogger::LogMessage(ConvertToStdString(Label_Commands->Text));
                }
            }
        }

        // Условие: прошло по таймеру, бит Work=0 и прошло 5 секунд ожидания.
        if (workBitZeroTimerActive && workStopFinalizeDelayElapsed && !currentWorkBitState) {
                workBitZeroTimerActive = false;
                workStopFinalizeDelayElapsed = false;
                
                // Запоминаем время окончания цикла данных
                dataCollectionEndTime = now;
                
            GlobalLogger::LogMessage(ConvertToStdString("Information: Цикл завершён (Work=0, +5 сек ожидания), экспорт в файл " + excelFileName));
                // Показываем кнопку START в нажатом виде
                buttonSTARTstate_TRUE();
                
                // Считаем, что данные экспортированы (файл не перезаписывается при следующем цикле)
                dataExportedToExcel = true;
                
                // Экспорт данных в Excel
                this->BeginInvoke(gcnew MethodInvoker(this, &DataForm::TriggerExcelExport));

            // Замечание: кнопки START/STOP обрабатываются (при нажатии пользователем). Вызов START в UI message loop,
            // затем отправка команды на устройство и экспорт данных в Excel при следующем подключении.
                if (autoRestartPending) {
                    autoRestartPending = false;
                autoRestartStopIssuedTime = DateTime::MinValue;
                    this->BeginInvoke(gcnew MethodInvoker(this, &DataForm::ExecuteAutoRestartStart));
            }
        }

        // Запоминаем текущее состояние бита Work
        workBitDetected = currentWorkBitState;
    }

    // Записываем биты первой группы в строку таблицы
    for (int bit = 0; bit < 16; bit++) {
        bool bitValue = (bitField & (1 << bit)) != 0;
        row[bitNames[0][bit]] = bitValue;
    }
    bitField = data.H[SQ - 1];
    // Записываем биты второй группы в строку таблицы
    for (int bit = 0; bit < 16; bit++) {
        bool bitValue = (bitField & (1 << bit)) != 0;
        row[bitNames[1][bit]] = bitValue;
    }

    // Добавляем строку в таблицу только при сборе данных
    // Условия добавления строки: 
    // 1. Work = 1 (устройство в процессе работы)
    // 2. Work = 0, но ещё не истекло время ожидания экспорта (workBitZeroTimerActive = true)
    //    или уже истекло — тогда бит Work->0 и Open->1
    if (workBitDetected || workBitZeroTimerActive)
    {
        table->Rows->Add(row);
        
        // Сбрасываем флаг экспорта, т.к. добавлена новая строка, которая ещё не выгружена в Excel
        // при следующем цикле данные этой сессии будут экспортированы в файл
        dataExportedToExcel = false;
    }
}


// Связка Excel реализована в FormExcel.cpp

// Отложенная сборка мусора
void DataForm::DelayedGarbageCollection(Object^ state) {
    try {
        // Пауза перед сборкой мусора
        //MessageBox::Show("Ошибка Excel не удалось запуститься!");
        GlobalLogger::LogMessage(ConvertToStdString("Ошибка Excel не удалось запуститься!"));
        Thread::Sleep(500);

        // Сборка мусора и ожидание финализаторов
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

        // Замечание: 5 секунд прошло после бита Work=0. Устанавливаем флаг ожидания экспорта.
        // Таймер больше не перезапускается, данные считаются стабилизированными по таймеру с Work=0.
        workStopFinalizeDelayElapsed = true;
        GlobalLogger::LogMessage("Information: 5 секунд после бита Work=0 истекло, разрешаем экспорт для пользователя");
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

        // 1) Нет связи с устройством "текущего" окна (по IP) уже 30 минут.
        // Условие: если нет связи более 30 минут с момента отключения по этому IP и адресу порта,
        // то форма считается мёртвой по этому IP адресу порта, и "окно" можно закрыть автоматически.
        if (ClientSocket == INVALID_SOCKET && !String::IsNullOrEmpty(ClientIP)) {
            const long long ticks = System::Threading::Interlocked::Read(disconnectedSinceTicks);
            if (ticks != 0) {
                DateTime disconnectedSince(ticks);
                TimeSpan waited = now.Subtract(disconnectedSince);
                if (waited.TotalMinutes >= 30) {
                    GlobalLogger::LogMessage(ConvertToStdString(String::Format(
                        "Information: Нет соединения с устройством {0} {1:F1} мин. Закрытие формы DataForm.",
                        ClientIP,
                        waited.TotalMinutes)));

                    // Если телеметрии не было в аварийном экспорте, закрываем форму.
                    if (!hasTelemetry) {
                        this->Close();
                        return;
                    }

                    // Вызов экспорта не из формы, т.к. при закрытии формы: Emergency + аварийный.
                    // Повторный захват монитора; поток вызывающий для аварийного экспорта в том же потоке,
                    // иначе не блокируем поток из-за экспортного guard.
                    if (!StartExcelExportThread(true)) {
                        GlobalLogger::LogMessage("Warning: Не удалось запустить аварийный экспорт; закрытие без сохранения данных.");
                        return;
                    }

                    inactivityCloseRequested = true;
                    dataExportedToExcel = true;
                    this->Close();
                    return;
                }
            }
        }

        // 2) Проверка по последней телеметрии (если устройство отключилось давно).
        if (!hasTelemetry || lastTelemetryTime == DateTime::MinValue) {
            return;
        }

        TimeSpan idle = now.Subtract(lastTelemetryTime);
        if (idle.TotalMinutes < 30) {
            return;
        }

        // Замечание: фиксируем время окончания цикла сейчас (как при EmergencyData-экспорте).
        if (dataCollectionEndTime == DateTime::MinValue) {
            dataCollectionEndTime = now;
        }

        GlobalLogger::LogMessage(ConvertToStdString(String::Format(
            "Information: Нет телеметрии {0:F1} мин. Закрытие формы DataForm.",
            idle.TotalMinutes)));

        // Повторный захват монитора; поток вызывающий для аварийного экспорта в том же потоке, иначе не блокируем поток из-за экспортного guard.
        if (!StartExcelExportThread(true)) {
            GlobalLogger::LogMessage("Warning: Не удалось запустить аварийный экспорт; закрытие без сохранения данных.");
            return;
        }

        inactivityCloseRequested = true;
        dataExportedToExcel = true;

        // Закрываем форму после аварийного экспорта данных в файл.
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

        // Замечание: если сокет переподключился с того же устройства по порту и тот же сокет.
        // Замечание: после сброса устройства сокет мог переподключиться к новому сокету, но порт тот же,
        // поэтому сбрасываем флаг по номеру порта UART4.
        if (lastStartupCommandsSocket == ClientSocket) {
            return;
        }

        // Отправляем версию и интервал измерения команды на UI-устройство.
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

    // Замечание: т.к. устройство могло получить RESET от контроллера, сбрасываем признак отправленных команд/параметров.
    lastStartupCommandsSocket = INVALID_SOCKET;

    if (postResetInitTimer == nullptr) {
        postResetInitTimer = gcnew System::Windows::Forms::Timer();
        postResetInitTimer->Tick += gcnew EventHandler(this, &DataForm::OnPostResetInitTimerTick);
    }

    postResetInitTimer->Stop();
    postResetInitTimer->Interval = 7000; // 7 секунд на ожидание ответа устройства по порту UART4
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
            // Нет соединения TCP (ожидание установки соединения сокет ещё/повтор).
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

        // Повтор через таймер; ожидание ответа по порту UART дольше, чем TCP.
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
    // Вызов приходит из потока приёмника, выполняется через Invoke
    // Важно! Данные, пришедшие в UI, обрабатываются в контексте формы, обновляем данные в UI!
    // Устанавливаем порт клиента в контексте DataForm
    this->clientPort = port;

    // Блокировка управляемого массива в неуправляемой памяти (pin_ptr для передачи в native-код)
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
            // Замечание: переподключение после отключения в SServer. Когда сокет переподключился к другому
            // "текущему устройству" по новому сокету отображаем сообщение о переподключении TCP-соединения.
            reconnectFixationLogPending = true;
        }

        hasTelemetry = true;
        lastTelemetryTime = telemetryTime;
        lastTelemetrySocket = currentSocket;

        // Замечание: если время пришло с момента переключения Work->1 (как при переподключении "к устройству" окна), то обновляем время начала цикла.
        if (dataCollectionStartTime == DateTime::MinValue) {
            dataCollectionStartTime = telemetryTime;
        }

        dataGridView->SuspendLayout();
        try {
            // Замечание: вызов Add в Snapshot потоке уже потокобезопасен.
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
********************** Блок сохранения настроек в файл Excel **********************
*/

void ProjectServerW::DataForm::SaveSettings() {
    try {
        // Получаем путь к папке настроек
        // Сохраняем настройки, куда писать экспорт
        String^ appPath = System::IO::Path::GetDirectoryName(System::Windows::Forms::Application::ExecutablePath);
        String^ settingsPath = System::IO::Path::Combine(appPath, "ExcelSettings.txt");

        System::IO::StreamWriter^ writer = gcnew System::IO::StreamWriter(settingsPath);
        // Строка 1 — путь к "папке" экспорта Excel для сохранения данных.
        writer->WriteLine(textBoxExcelDirectory->Text);

        // Замечание: настройки формы сохраняются в файл при закрытии; при открытии форма загружает настройки из файла.
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
        // Ошибка при записи настроек в файл
        MessageBox::Show("Не удалось сохранить настройки в файл: " + ex->Message);
        GlobalLogger::LogMessage(ConvertToStdString("Не удалось сохранить настройки в файл: " + ex->Message));
    }
}

void ProjectServerW::DataForm::LoadSettings() {
    try {
        settingsLoading = true;
        // Загружаем настройки из файла
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
        // Ошибка при загрузке настроек
        MessageBox::Show("Не удалось загрузить настройки: " + ex->Message);
        GlobalLogger::LogMessage(ConvertToStdString("Не удалось загрузить настройки: " + ex->Message));
    }
    finally {
        settingsLoading = false;
    }
}

// Загрузка сохранённых настроек папки экспорта при открытии EXCEL
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
    dataGridView2->Rows->Add("leftRightTrimMaxEq", "Max trim heater equivalent", "0.6");
    dataGridView2->Rows->Add("piKp", "PI proportional gain (supply temp)", "0.18");
    dataGridView2->Rows->Add("piKi", "PI integral gain (supply temp)", "0.02");
    dataGridView2->Rows->Add("wDeadband_kgkg", "Humidity deadband", "0.0008");
    dataGridView2->Rows->Add("injGain", "Injection gain (duty per kg/kg humidity error)", "900");
    dataGridView2->Rows->Add("outDamperTimer_s", "Damper open time", "10");
    dataGridView2->Rows->Add("outFanDelay_s", "Fan delay after damper open", "5");
    dataGridView2->Rows->Add("tenMinHold_s", "Heater min hold between switches", "10");
    dataGridView2->Rows->Add("injMinHold_s", "Injector min hold between switches", "5");
    dataGridView2->Rows->Add("outHold_s", "Damper and exhaust fan min hold", "15");
    dataGridView2->Rows->Add("airOnlyPhaseWarmUp_s", "Air-only: WarmUp phase duration (s)", "600");
    dataGridView2->Rows->Add("airOnlyPhasePlateau_s", "Air-only: Plateau end time from start (s)", "1800");
    dataGridView2->Rows->Add("maxRuntime_s", "Air-only: Max process duration (s)", "7200");
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
                "Сохранение",
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

// ВВВВВВВВВВВ ВВВ ВВВВВВВВВ (Parameter2) В ВВВВВВВВВВВВВВ ВВВВВВВВВ: groupId, paramId, ВВВ (1=U8, 2=U16, 3=F32)
static bool GetDefrostParamId(System::String^ paramName, uint8_t% outGroupId, uint8_t% outParamId, uint8_t% outValueType) {
    if (paramName == nullptr) return false;
    String^ p = paramName->Trim();
    if (String::IsNullOrEmpty(p)) return false;
    // ВВВВВВ: 1=SENSORS, 2=TEMPERATURE, 3=HUMIDITY, 4=PWM (ВВВВВВВВВВВВ В DefrostControl.h)
    if (p->StartsWith("Sensor", StringComparison::OrdinalIgnoreCase) && p->Contains("use in defrost")) {
        int i = -1; if (p->Length >= 7) Int32::TryParse(p->Substring(6, 1), i);
        if (i >= 0 && i <= 6) { outGroupId = 1; outParamId = (uint8_t)i; outValueType = DefrostParamType::U8; return true; }
    }
    if (p->Equals("leftRightTrimGain", StringComparison::OrdinalIgnoreCase)) { outGroupId = 2; outParamId = 15; outValueType = DefrostParamType::F32; return true; }
    if (p->Equals("leftRightTrimMaxEq", StringComparison::OrdinalIgnoreCase)) { outGroupId = 2; outParamId = 16; outValueType = DefrostParamType::F32; return true; }
    if (p->Equals("piKp", StringComparison::OrdinalIgnoreCase)) { outGroupId = 2; outParamId = 17; outValueType = DefrostParamType::F32; return true; }
    if (p->Equals("piKi", StringComparison::OrdinalIgnoreCase)) { outGroupId = 2; outParamId = 18; outValueType = DefrostParamType::F32; return true; }
    if (p->Equals("wDeadband_kgkg", StringComparison::OrdinalIgnoreCase)) { outGroupId = 3; outParamId = 3; outValueType = DefrostParamType::F32; return true; }
    if (p->Equals("injGain", StringComparison::OrdinalIgnoreCase)) { outGroupId = 3; outParamId = 6; outValueType = DefrostParamType::F32; return true; }
    if (p->Equals("outDamperTimer_s", StringComparison::OrdinalIgnoreCase)) { outGroupId = 3; outParamId = 4; outValueType = DefrostParamType::U16; return true; }
    if (p->Equals("outFanDelay_s", StringComparison::OrdinalIgnoreCase)) { outGroupId = 3; outParamId = 5; outValueType = DefrostParamType::U16; return true; }
    if (p->Equals("tenMinHold_s", StringComparison::OrdinalIgnoreCase)) { outGroupId = 4; outParamId = 0; outValueType = DefrostParamType::U16; return true; }
    if (p->Equals("injMinHold_s", StringComparison::OrdinalIgnoreCase)) { outGroupId = 4; outParamId = 1; outValueType = DefrostParamType::U16; return true; }
    if (p->Equals("outHold_s", StringComparison::OrdinalIgnoreCase)) { outGroupId = 4; outParamId = 2; outValueType = DefrostParamType::U16; return true; }
    if (p->Equals("airOnlyPhaseWarmUp_s", StringComparison::OrdinalIgnoreCase)) { outGroupId = 4; outParamId = 3; outValueType = DefrostParamType::U16; return true; }
    if (p->Equals("airOnlyPhasePlateau_s", StringComparison::OrdinalIgnoreCase)) { outGroupId = 4; outParamId = 4; outValueType = DefrostParamType::U16; return true; }
    if (p->Equals("maxRuntime_s", StringComparison::OrdinalIgnoreCase)) { outGroupId = 4; outParamId = 5; outValueType = DefrostParamType::U16; return true; }
    return false;
}

// dataGridView1: Parameter + phase (0=WarmUP, 1=Plateau, 2=Finish) -> groupId, paramId. DefrostControl TEMPERATURE 0-14, HUMIDITY 0-2.
static bool GetDefrostParamIdGrid1(System::String^ paramName, int phaseIndex, uint8_t% outGroupId, uint8_t% outParamId) {
    if (paramName == nullptr || phaseIndex < 0 || phaseIndex > 2) return false;
    String^ p = paramName->Trim();
    if (String::IsNullOrEmpty(p)) return false;
    uint8_t ph = (uint8_t)phaseIndex;
    if (p->Equals("fishHotMax_C", StringComparison::OrdinalIgnoreCase)) { outGroupId = 2; outParamId = ph; return true; }
    if (p->Equals("supplySet_C", StringComparison::OrdinalIgnoreCase)) { outGroupId = 2; outParamId = (uint8_t)(3 + ph); return true; }
    if (p->Equals("supplyMax_C", StringComparison::OrdinalIgnoreCase)) { outGroupId = 2; outParamId = (uint8_t)(6 + ph); return true; }
    if (p->Equals("fishDeltaMax_C", StringComparison::OrdinalIgnoreCase)) { outGroupId = 2; outParamId = (uint8_t)(9 + ph); return true; }
    if (p->Equals("fishHotRateMax_Cps", StringComparison::OrdinalIgnoreCase)) { outGroupId = 2; outParamId = (uint8_t)(12 + ph); return true; }
    if (p->Equals("returnTargetRH_percent", StringComparison::OrdinalIgnoreCase)) { outGroupId = 3; outParamId = ph; return true; }
    return false;
}

System::Void ProjectServerW::DataForm::buttonLoadFromFile_Click(System::Object^ sender, System::EventArgs^ e) {
    LoadDataGridView1FromFile();
    LoadDataGridView2FromFile();
}

System::Void ProjectServerW::DataForm::buttonSaveToFile_Click(System::Object^ sender, System::EventArgs^ e) {
    SaveDataGridView1ToFile();
    SaveDataGridView2ToFile();
}

System::Void ProjectServerW::DataForm::buttonReadParameters_Click(System::Object^ sender, System::EventArgs^ e) {
    if ((dataGridView1 == nullptr || dataGridView1->IsDisposed) && (dataGridView2 == nullptr || dataGridView2->IsDisposed)) return;
    if (ClientSocket == INVALID_SOCKET) {
        MessageBox::Show("Нет соединения с устройством. Подключитесь к устройству.");
        if (Label_Commands != nullptr && !Label_Commands->IsDisposed) {
            Label_Commands->Text = "Соединение с устройством: нет соединения";
            Label_Commands->ForeColor = System::Drawing::Color::Orange;
        }
        return;
    }
    int ok = 0, fail = 0;
    if (dataGridView1 != nullptr && !dataGridView1->IsDisposed) {
        for (int r = 0; r < dataGridView1->Rows->Count; r++) {
            DataGridViewRow^ row = dataGridView1->Rows[r];
            if (row->IsNewRow) continue;
            Object^ v0 = row->Cells["Parameter"]->Value;
            String^ paramName = v0 != nullptr ? v0->ToString() : "";
            for (int ph = 0; ph < 3; ph++) {
                uint8_t g, id;
                if (!GetDefrostParamIdGrid1(paramName, ph, g, id)) continue;
                DefrostParamValue val;
                if (!GetDefrostParam(g, id, &val)) { fail++; continue; }
                String^ valueStr = (val.valueType == DefrostParamType::F32) ? val.value.f32.ToString(System::Globalization::CultureInfo::InvariantCulture) : "";
                String^ colName = (ph == 0) ? "WarmUP" : (ph == 1) ? "Plateau" : "Finish";
                row->Cells[colName]->Value = valueStr;
                ok++;
            }
        }
    }
    if (dataGridView2 != nullptr && !dataGridView2->IsDisposed) {
        for (int r = 0; r < dataGridView2->Rows->Count; r++) {
            DataGridViewRow^ row = dataGridView2->Rows[r];
            if (row->IsNewRow) continue;
            Object^ v0 = row->Cells["Parameter2"]->Value;
            String^ paramName = v0 != nullptr ? v0->ToString() : "";
            uint8_t g, id, vt;
            if (!GetDefrostParamId(paramName, g, id, vt)) continue;
            DefrostParamValue val;
            if (!GetDefrostParam(g, id, &val)) { fail++; continue; }
            String^ valueStr = "";
            if (val.valueType == DefrostParamType::U8) valueStr = val.value.u8.ToString();
            else if (val.valueType == DefrostParamType::U16) valueStr = val.value.u16.ToString();
            else if (val.valueType == DefrostParamType::F32) valueStr = val.value.f32.ToString(System::Globalization::CultureInfo::InvariantCulture);
            row->Cells["Value"]->Value = valueStr;
            ok++;
        }
    }
    if (Label_Commands != nullptr && !Label_Commands->IsDisposed) {
        Label_Commands->Text = String::Format("Считать с устройства: прочитано {0} параметров" + (fail > 0 ? ", ошибок: " + fail : ""), ok);
        Label_Commands->ForeColor = System::Drawing::Color::DarkGreen;
    }
    GlobalLogger::LogMessage(ConvertToStdString(String::Format("Information: Read {0} params from defroster" + (fail > 0 ? ", {1} failed" : ""), ok, fail)));
}

System::Void ProjectServerW::DataForm::buttonWriteParameters_Click(System::Object^ sender, System::EventArgs^ e) {
    if ((dataGridView1 == nullptr || dataGridView1->IsDisposed) && (dataGridView2 == nullptr || dataGridView2->IsDisposed)) return;
    if (ClientSocket == INVALID_SOCKET) {
        MessageBox::Show("Нет соединения с устройством. Подключитесь к устройству.");
        if (Label_Commands != nullptr && !Label_Commands->IsDisposed) {
            Label_Commands->Text = "Запись в устройство: нет соединения";
            Label_Commands->ForeColor = System::Drawing::Color::Orange;
        }
        return;
    }
    int ok = 0, fail = 0;
    if (dataGridView1 != nullptr && !dataGridView1->IsDisposed) {
        for (int r = 0; r < dataGridView1->Rows->Count; r++) {
            DataGridViewRow^ row = dataGridView1->Rows[r];
            if (row->IsNewRow) continue;
            Object^ v0 = row->Cells["Parameter"]->Value;
            String^ paramName = v0 != nullptr ? v0->ToString() : "";
            for (int ph = 0; ph < 3; ph++) {
                uint8_t g, id;
                if (!GetDefrostParamIdGrid1(paramName, ph, g, id)) continue;
                String^ colName = (ph == 0) ? "WarmUP" : (ph == 1) ? "Plateau" : "Finish";
                Object^ vc = row->Cells[colName]->Value;
                String^ valueStr = vc != nullptr ? vc->ToString() : "";
                float f;
                if (!Single::TryParse(valueStr->Trim()->Replace(",", "."), System::Globalization::NumberStyles::Float, System::Globalization::CultureInfo::InvariantCulture, f)) { fail++; continue; }
                DefrostParamValue val;
                val.valueType = DefrostParamType::F32;
                val.value.f32 = f;
                if (!SetDefrostParam(g, id, val)) { fail++; continue; }
                ok++;
            }
        }
    }
    if (dataGridView2 != nullptr && !dataGridView2->IsDisposed) {
        for (int r = 0; r < dataGridView2->Rows->Count; r++) {
            DataGridViewRow^ row = dataGridView2->Rows[r];
            if (row->IsNewRow) continue;
            Object^ v0 = row->Cells["Parameter2"]->Value;
            Object^ v2 = row->Cells["Value"]->Value;
            String^ paramName = v0 != nullptr ? v0->ToString() : "";
            String^ valueStr = v2 != nullptr ? v2->ToString() : "";
            uint8_t g, id, vt;
            if (!GetDefrostParamId(paramName, g, id, vt)) continue;
            DefrostParamValue val;
            val.valueType = vt;
            if (vt == DefrostParamType::U8) {
                unsigned int u; if (!UInt32::TryParse(valueStr->Trim(), u)) { fail++; continue; }
                val.value.u8 = (uint8_t)(u & 0xFF);
            } else if (vt == DefrostParamType::U16) {
                unsigned int u; if (!UInt32::TryParse(valueStr->Trim(), u)) { fail++; continue; }
                val.value.u16 = (uint16_t)(u & 0xFFFF);
            } else if (vt == DefrostParamType::F32) {
                float f; if (!Single::TryParse(valueStr->Trim()->Replace(",", "."), System::Globalization::NumberStyles::Float, System::Globalization::CultureInfo::InvariantCulture, f)) { fail++; continue; }
                val.value.f32 = f;
            } else { fail++; continue; }
            if (!SetDefrostParam(g, id, val)) { fail++; continue; }
            ok++;
        }
    }
    if (Label_Commands != nullptr && !Label_Commands->IsDisposed) {
        Label_Commands->Text = String::Format("Запись в устройство: записано {0} параметров" + (fail > 0 ? ", ошибок: " + fail : ""), ok);
        Label_Commands->ForeColor = System::Drawing::Color::DarkGreen;
    }
    GlobalLogger::LogMessage(ConvertToStdString(String::Format("Information: Wrote {0} params to defroster" + (fail > 0 ? ", {1} failed" : ""), ok, fail)));
}

//*******************************************************************************************
// Инициализация имён битовых полей для таблицы данных
void ProjectServerW::DataForm::InitializeBitFieldNames(gcroot<cli::array<cli::array<String^>^>^>& namesRef) {
    namesRef = gcnew cli::array<cli::array<String^>^>(10);

    // Первая группа битов (управление/статус)
    namesRef[0] = gcnew cli::array<String^>(16) {
        "Heat0", "Heat1", "Heat2", "Heat3", // Тэны испарителя  (нагрев) 1..4
        "Vent0", "Vent1", "Vent2", "Vent3", // Вентиляторы испарителя 1..4
        "InjW",     // Инжектор воды увлажнения
        "UP",       // Клапан подъёма шторки вверх
        "DOWN",     // Клапан опускания шторки вниз
        "DBlk",     // Блокировка двери
        "Clse",     // Клапан закрытия заслонки
        "Open",     // Клапан открытия заслонки
        "Work",     // Режим работы 
        "Alrm",     // Режим аварии 
        
    };

    // Вторая группа битов (статусы/отклики)
    namesRef[1] = gcnew cli::array<String^>(16) {
        "_V0", "_V1", "_V2", "_V3", // Вентиляторы статус 1..4 отклик
        "_H0", "_H1", "_H2", "_H3", // Тэны (нагрев) 1..4 отклик
        "_Out",     // Заслонка открыта отклик
        "_Inj",     // Инжектор включён отклик
        "_Flp",     // Инжектор флоп открыт для увлажнения 
        "_Opn",     // Открытие шторки
        "_Dbl",     // Дверь заблокирована
        "_Cls",     // Шторка закрыта
        "_Wrk",     // Устройство в режиме работы цикла
        "_Red"      // Устройство в аварии
    };

    // Остальные группы при необходимости расширяются
}

////******************** Обработка закрытия формы ***************************************
System::Void ProjectServerW::DataForm::DataForm_FormClosing(System::Object^ sender, System::Windows::Forms::FormClosingEventArgs^ e) {
    try {
        // Проверяем, есть ли данные в таблице и не были ли они уже экспортированы
        if (dataTable != nullptr && dataTable->Rows->Count > 0) {
            
            // Условие: если данные уже выгружены в Excel
            if (dataExportedToExcel) {
                // Данные за цикл экспортированы — нет предупреждения
                GlobalLogger::LogMessage("Information: Данные за цикл экспортированы в Excel, закрытие формы без предупреждения");
            }
            else {
                // Данные за цикл не экспортированы. Запускаем экспорт из-за закрытия формы, выгружаем в файл перед закрытием.
                // Имя файла типа "EmergencyData" экспорт при закрытии.
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

                // Аварийный экспорт при закрытии
                GlobalLogger::LogMessage("Information: Экспортируем данные в Excel при закрытии формы... " + ConvertToStdString(excelFileName));

                // Вызов экспорта в потоке для выполнения в UI-потоке при закрытии формы
                StartExcelExportThread(true);
                
                // При закрытии формы — аварийный экспорт
                // для сохранения данных перед закрытием
                GlobalLogger::LogMessage("Information: Форма закрывается, данные в Excel экспортируются перед закрытием формы");
            }
        }
        
        // Разрешаем закрытие формы
        e->Cancel = false;
    }
    catch (Exception^ ex) {
        GlobalLogger::LogMessage("Ошибка при записи параметров в устройство: " + ConvertToStdString(ex->Message));
        // При ошибке всё равно разрешаем закрытие формы
        e->Cancel = false;
    }
}

// Обработчик события закрытия формы
System::Void DataForm::DataForm_FormClosed(Object^ sender, FormClosedEventArgs^ e)
{
    // Удаляем форму из карты (если ещё не удалена)
    // Форма могла быть закрыта кнопкой и удалена в CloseForm()
    bool formFoundInMap = false;
    for (auto it = formData_Map.begin(); it != formData_Map.end(); ++it) {
        // Получаем указатель формы из gcroot
        ProjectServerW::DataForm^ formPtr = it->second;

        // Проверяем, это наша форма
        if (formPtr == this) {
            // Удаляем форму из карты, возвращаем nullptr
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

    // Замечание: сокет управляемый код (managed), при закрытии формы. Не вызывать native-код после этого в деструкторе/финализаторе.
    try {
        if (this->FormGuid != nullptr && this->FormGuid->Length > 0) {
            msclr::interop::marshal_context ctx;
            std::wstring guidW = ctx.marshal_as<std::wstring>(this->FormGuid);
            ThreadStorage::StopThread(guidW);
        }
    }
    catch (...) {}
}

// Инициализация последовательного порта
System::Void DataForm::DataForm_HandleDestroyed(Object^ sender, EventArgs^ e)
{
    try {
        // Открываем по имени порта
        this->FormClosing -= gcnew FormClosingEventHandler(this, &DataForm::DataForm_FormClosing);
        this->FormClosed -= gcnew FormClosedEventHandler(this, &DataForm::DataForm_FormClosed);
        this->HandleDestroyed -= gcnew EventHandler(this, &DataForm::DataForm_HandleDestroyed);

        // Устанавливаем параметры порта для соединения COM-портом
        GC::Collect();
        GC::WaitForPendingFinalizers();
        GC::Collect();
    }
    catch (...) {
        // Сохраняем настройки в переменной
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

// Отправка команды дефроста на устройство (реализация в CommandsDefroster.cpp)

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
// Экспорт данных таблицы в Excel (аварийный экспорт)
//  ===========================================================================

// Запрос версии прошивки с устройства
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
// Отправка команды интервала на устройство
//  ===========================================================================

// Вызов команды START на устройство
System::Void ProjectServerW::DataForm::checkBoxAutoStart_CheckedChanged(System::Object^ sender, System::EventArgs^ e) {
    // Включение/выключение автозапуска
    if (checkBoxAutoStart->Checked) {
        // Запуск таймера
        timerAutoStart->Start();
        GlobalLogger::LogMessage(ConvertToStdString(String::Format(
            "Information: Автозапуск установлен на {0}",
            dateTimePickerAutoStart->Value.ToString("HH:mm"))));
        
        // Включаем элементы (метка зелёная, выбор времени доступен)
        labelAutoStart->ForeColor = System::Drawing::Color::Green;
        dateTimePickerAutoStart->Enabled = true;
    }
    else {
        // Остановка таймера
        timerAutoStart->Stop();
        GlobalLogger::LogMessage("Information: Автозапуск отключён");
        
        // Выключение элементов
        labelAutoStart->ForeColor = System::Drawing::SystemColors::ControlText;
    }
}

// Таймер автозапуска по времени
System::Void ProjectServerW::DataForm::timerAutoStart_Tick(System::Object^ sender, System::EventArgs^ e) {
    // Проверка флажка автозапуска
    if (!checkBoxAutoStart->Checked) {
        return; // если чекбокс снят, выходим
    }
    
    DateTime now = DateTime::Now;
    DateTime targetTime = dateTimePickerAutoStart->Value;
    
    // Сравниваем текущее время с целевым (час и минута)
    if (now.Hour == targetTime.Hour && now.Minute == targetTime.Minute) {
        // Время пришло — срабатывание автозапуска
        GlobalLogger::LogMessage(ConvertToStdString(String::Format(
            "Information: Автозапуск сработал в {0} (целевое: {1})",
            now.ToString("HH:mm:ss"),
            targetTime.ToString("HH:mm"))));
        
        // Проверяем, что кнопка START доступна (нажатие по таймеру)
        if (buttonSTART->Enabled) {
            // Включение элементов
            labelAutoStart->ForeColor = System::Drawing::Color::Blue;
            Label_Commands->Text = "[Ожидание] Выполняется команда остановки...";
            Label_Commands->ForeColor = System::Drawing::Color::Blue;
            GlobalLogger::LogMessage(ConvertToStdString(Label_Commands->Text));
            
            CommandResponse startResp{};
            CommandAckResult startResult = SendControlCommandWithAck(CmdProgControl::START, "START", 2000, 2, startResp);

            if (startResult == CommandAckResult::NoResponse) {
                // Замечание: команда отправлена на устройство, но ACK не получен; ожидание по биту Work в следующем пакете.
                GlobalLogger::LogMessage("Warning: Автозапуск: START не подтверждён, ожидание по биту Work");
                Label_Commands->Text = "[Ожидание] START не отправлен, ожидание по Work...";
                Label_Commands->ForeColor = System::Drawing::Color::Orange;
                GlobalLogger::LogMessage(ConvertToStdString(Label_Commands->Text));
            }
            else if (startResult != CommandAckResult::Ok) {
                GlobalLogger::LogMessage("Warning: Автозапуск: START не выполнен (ошибка связи/таймаут)");
                Label_Commands->Text = "[!] Ошибка: START не отправлен";
                Label_Commands->ForeColor = System::Drawing::Color::Orange;
                GlobalLogger::LogMessage(ConvertToStdString(Label_Commands->Text));
            }
            
            // Отключаем чекбокс автозапуска
            checkBoxAutoStart->Checked = false;
            
            // Таймер сброса цвета через 3 секунды
            System::Windows::Forms::Timer^ colorTimer = gcnew System::Windows::Forms::Timer();
            colorTimer->Interval = 3000;
            colorTimer->Tick += gcnew EventHandler(this, &DataForm::RestoreAutoStartColor);
            colorTimer->Start();
        }
        else {
            // Кнопка недоступна
            GlobalLogger::LogMessage("Warning: Автозапуск не выполнен — кнопка недоступна");
            Label_Commands->Text = "[!] Таймаут операции - проверьте подключение";
            Label_Commands->ForeColor = System::Drawing::Color::Orange;
            GlobalLogger::LogMessage(ConvertToStdString(Label_Commands->Text));
            
            // Отключаем автозапуск
            checkBoxAutoStart->Checked = false;
        }
    }
}

// Восстановление цвета метки автозапуска
System::Void ProjectServerW::DataForm::RestoreAutoStartColor(System::Object^ sender, System::EventArgs^ e) {
    try {
        // Останавливаем таймер
        System::Windows::Forms::Timer^ timer = safe_cast<System::Windows::Forms::Timer^>(sender);
        timer->Stop();
        timer->Tick -= gcnew EventHandler(this, &DataForm::RestoreAutoStartColor);
        
        // Восстановление цвета
        labelAutoStart->ForeColor = System::Drawing::SystemColors::ControlText;
    }
    catch (Exception^ ex) {
        GlobalLogger::LogMessage(ConvertToStdString("Error in RestoreAutoStartColor: " + ex->Message));
    }
}

//  ===========================================================================
// Отправка команды автоперезапуска по времени
//  ===========================================================================

System::Void ProjectServerW::DataForm::checkBoxAutoRestart_CheckedChanged(System::Object^ sender, System::EventArgs^ e) {
    if (checkBoxAutoRestart->Checked) {
        timerAutoRestart->Start();
        GlobalLogger::LogMessage(ConvertToStdString(String::Format(
            "Information: Автоперезапуск установлен на {0}",
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
        GlobalLogger::LogMessage("Information: Автоперезапуск отключён");

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
        "Information: Автоперезапуск сработал в {0} (целевое: {1})",
        now.ToString("HH:mm:ss"),
        targetTime.ToString("HH:mm"))));

    // Если ни STOP ни START недоступны (устройство отключено), то "автоперезапуск" отменяется.
    if (!buttonSTOP->Enabled && !buttonSTART->Enabled) {
        GlobalLogger::LogMessage("Warning: Автоперезапуск: устройство отключено (нет связи), отмена...");
        return;
    }

    labelAutoRestart->ForeColor = System::Drawing::Color::Blue;

    // Если нажата кнопка STOP доступна (отправка STOP) — отправляем STOP и ждём подтверждения Work=0.
    if (buttonSTOP->Enabled) {
        Label_Commands->Text = "[Выполняется] Отправлена команда STOP...";
        Label_Commands->ForeColor = System::Drawing::Color::Blue;
        GlobalLogger::LogMessage(ConvertToStdString(Label_Commands->Text));

        CommandResponse stopResp{};
        CommandAckResult stopResult = SendControlCommandWithAck(CmdProgControl::STOP, "STOP", 2000, 2, stopResp);

        if (stopResult == CommandAckResult::Ok) {
            autoRestartPending = true;
            autoRestartStopIssuedTime = DateTime::Now;
            Label_Commands->Text = "[Выполняется] STOP отправлена, ожидание экспорта данных...";
            Label_Commands->ForeColor = System::Drawing::Color::Blue;
            GlobalLogger::LogMessage(ConvertToStdString(Label_Commands->Text));
        }
        else if (stopResult == CommandAckResult::NoResponse) {
            // Замечание: команда отправлена STOP, но ACK не получен; считаем отправленной и ожидаем по биту Work.
            autoRestartPending = true;
            autoRestartStopIssuedTime = DateTime::Now;
            GlobalLogger::LogMessage("Warning: Автоперезапуск: STOP не подтверждён, ожидание по биту Work");
            Label_Commands->Text = "[Выполняется] STOP не отправлен, ожидание по Work...";
            Label_Commands->ForeColor = System::Drawing::Color::Orange;
            GlobalLogger::LogMessage(ConvertToStdString(Label_Commands->Text));
        }
        else {
            autoRestartPending = false;
            autoRestartStopIssuedTime = DateTime::MinValue;
            GlobalLogger::LogMessage("Warning: Автоперезапуск: STOP не выполнен (ошибка связи/таймаут)");
            Label_Commands->Text = "[!] Ошибка выполнения: STOP не отправлен";
            Label_Commands->ForeColor = System::Drawing::Color::Orange;
            GlobalLogger::LogMessage(ConvertToStdString(Label_Commands->Text));
        }

            autoRestartInternalUncheck = true;
        checkBoxAutoRestart->Checked = false; // сбрасываем, UX чекбокс не связан с AutoStart
    }
    else if (buttonSTART->Enabled) {
        // Устройство уже остановлено в этот момент — отправляем START.
        Label_Commands->Text = "[Выполняется] Ожидание отключения, отправка START...";
        Label_Commands->ForeColor = System::Drawing::Color::Blue;
        GlobalLogger::LogMessage(ConvertToStdString(Label_Commands->Text));

        CommandResponse startResp{};
        CommandAckResult startResult = SendControlCommandWithAck(CmdProgControl::START, "START", 2000, 2, startResp);

        if (startResult == CommandAckResult::Ok) {
            Label_Commands->Text = "[Выполняется] START отправлен";
            Label_Commands->ForeColor = System::Drawing::Color::Blue;
            GlobalLogger::LogMessage(ConvertToStdString(Label_Commands->Text));
        }
        else if (startResult == CommandAckResult::NoResponse) {
            // Замечание: устройство ответило по приёму ACK, но подтверждение бита Work придет в следующем пакете.
            GlobalLogger::LogMessage("Warning: Автоперезапуск: START не подтверждён, ожидание по биту Work");
            Label_Commands->Text = "[Выполняется] START не отправлен, ожидание по Work...";
            Label_Commands->ForeColor = System::Drawing::Color::Orange;
            GlobalLogger::LogMessage(ConvertToStdString(Label_Commands->Text));
        }
        else {
            GlobalLogger::LogMessage("Warning: Автоперезапуск: START не выполнен (ошибка связи/таймаут)");
            Label_Commands->Text = "[!] Ошибка выполнения: START не отправлен";
            Label_Commands->ForeColor = System::Drawing::Color::Orange;
            GlobalLogger::LogMessage(ConvertToStdString(Label_Commands->Text));
        }
        autoRestartInternalUncheck = true;
        checkBoxAutoRestart->Checked = false;
    }

    // Восстановление цвета метки автоперезапуска через 3 сек, как для "автозапуска".
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
            GlobalLogger::LogMessage("Warning: Автоперезапуск: START не выполнен (кнопка недоступна)");
            return;
        }

        Label_Commands->Text = "[Выполняется] Выполняется команда остановки...";
        Label_Commands->ForeColor = System::Drawing::Color::Blue;
        GlobalLogger::LogMessage(ConvertToStdString(Label_Commands->Text));
        SendStartCommand();
    }
    catch (Exception^ ex) {
        GlobalLogger::LogMessage(ConvertToStdString("Error: Exception in ExecuteAutoRestartStart: " + ex->ToString()));
    }
}

// Отправка команды старта дефроста (реализация в CommandsDefroster.cpp)


