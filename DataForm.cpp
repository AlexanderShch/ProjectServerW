#include "DataForm.h"
#include "Chart.h"
#include "FormExcel.h"
#include "Commands.h"               // Command, DefrostParam и обмен с контроллером
#include "PacketQueueProcessor.h"   // per-socket Очередь команд (ACK от контроллера vs ответы в UI)
#include <objbase.h>                // для CoCreateGuid — генерация уникального идентификатора окна
#include <string>
#include <vcclr.h>  // для gcnew
#include <msclr/marshal_cppstd.h>

// общая логика для всех окон в PI Process
using namespace System::Diagnostics;
using namespace ProjectServerW; // пространство имён приложения
using namespace Microsoft::Office::Interop::Excel;

std::map<std::wstring, gcroot<DataForm^>> formData_Map; // глобальная карта окон formData_Map

// Пакет лога алгоритма (Type 0x01): сначала отфильтрованные температуры 6 датчиков 0..5 (°C),
// затем текущая фаза + группа 3 — переменные алгоритма (совпадает с ControlLogPayload_t на контроллере).
#pragma pack(push, 1)
typedef struct {
    float T_filt_C[6];
    uint8_t phase;
    float eT_common, heatScale01;
    float uCommon_TEN, trim_TEN, uLeft_TEN, uRight_TEN;
    float leftTen1Duty, leftTen2Duty, rightTen1Duty, rightTen2Duty;
    float w_sup_avg, wErr, injDuty;
    float rate_Cps;
    float fishHot_C, fishCold_C;
} ControlLogPayload_t;
static_assert(sizeof(ControlLogPayload_t) == 89, "ControlLogPayload_t size must be 89 bytes");

/* Ответ GET_DEFROST_GROUP(groupId=5): структура совпадает с DefrostLogPhasePayload_t на контроллере. */
namespace { constexpr int DEFROST_PHASE_COUNT_SERVER = 3; }
typedef struct {
    float fishHotMax_C[DEFROST_PHASE_COUNT_SERVER];
    float fishHotRateMax_Cps[DEFROST_PHASE_COUNT_SERVER];
    float fishDeltaMax_C[DEFROST_PHASE_COUNT_SERVER];
    float supplySet_C[DEFROST_PHASE_COUNT_SERVER];
    float supplyMax_C[DEFROST_PHASE_COUNT_SERVER];
    float returnTargetRH_percent[DEFROST_PHASE_COUNT_SERVER];
} DefrostLogPhasePayload_t;

/* Ответ GET_DEFROST_GROUP(groupId=6): структура совпадает с DefrostLogGlobalPayload_t на контроллере. */
namespace { constexpr int DEFROST_MAX_SENSOR_COUNT_SERVER = 6; }   /* совпадает с DEFROST_MAX_SENSOR_COUNT на контроллере (датчики 0..5) */
typedef struct {
    float leftRightTrimGain;
    float leftRightTrimMaxEq;
    float piKp;
    float piKi;
    float wDeadband_kgkg;
    float injGain;
    uint16_t outDamperTimer_s;
    uint16_t outFanDelay_s;
    uint16_t outHold_s;
    uint16_t tenMinHold_s;
    uint16_t injMinHold_s;
    uint16_t airOnlyPhaseWarmUp_s;
    uint16_t airOnlyPhasePlateau_s;
    uint16_t maxRuntime_s;
    float fishColdTarget_C;          /* целевая мин. Т рыбы °C; при достижении — автоостанов алгоритма */
    uint8_t debugDisableTargetTStop; /* 1 = отладка: отключить автостоп по целевой Т, 0 = автостоп включен */
    uint8_t debugDisableDeviceSwitchCheck; /* 1 = отладка: отключить проверку соответствия входов/выходов, 0 = проверка включена */
    uint8_t sensorUseInDefrost[DEFROST_MAX_SENSOR_COUNT_SERVER];
} DefrostLogGlobalPayload_t;
static_assert(sizeof(DefrostLogGlobalPayload_t) + 2 <= (MAX_COMMAND_SIZE - 6),
    "GET_DEFROST_GROUP(group6): response payload+header exceeds CommandResponse::data capacity");
#pragma pack(pop)

// Совпадает с форматом пакета в типе MSGQUEUE_OBJ_t на STM32
void ProjectServerW::DataForm::ParseBuffer(const char* buffer, size_t size) {

    MSGQUEUE_OBJ_t data;

    // buffer содержит: [Type][Cmd][Status][DataLen][Data...][CRC16]
    if (size < 4 + sizeof(data) + 2) {
        return; // недостаточно байт для пакета
    }
    memcpy(&data, buffer + 4, sizeof(data));
}

// Обновление внутренних флагов автомата по текущей телеметрии. 
void ProjectServerW::DataForm::UpdateModeFlagsFromTelemetry(
    DateTime now,                       // Текущее время
    uint16_t deviceTime,                // Время из структуры пакета
    bool wrkBit,                        // Флаг _Wrk
    bool shdBit,                        // Флаг _Shd
    int& outDeltaCounts)                // Дельта времени
{
    const bool wasAuto = controllerAutoModeActive;   // Флаг автомата
    bool suppressStopAfterRecentStart = false;      // Флаг отключения останова после недавнего запуска (5 секунд)

    // Проверяем, есть ли время последнего успешного запуска
    if (lastStartSuccessTime != DateTime::MinValue) {   // Если время последнего успешного запуска не равно минимальному значению
        TimeSpan sinceStart = now.Subtract(lastStartSuccessTime);   // Вычисляем время с момента последнего успешного запуска
        suppressStopAfterRecentStart = (sinceStart.TotalSeconds >= 0.0 && sinceStart.TotalSeconds < 5.0);   // Устанавливаем флаг suppressStopAfterRecentStart
    }

    // Если флаг _Wrk установлен, сбрасываем счётчик нулевых последовательных отсчётов
    // И устанавливаем флаг автомата
    if (wrkBit) {
        wrkZeroConsecutiveCounts = 0;
        controllerAutoModeActive = true;   // Устанавливаем флаг автомата
    }
    // Если флаг _Wrk не установлен, но автомат был активен, увеличиваем счётчик нулевых последовательных отсчётов
    else if (wasAuto) {
        wrkZeroConsecutiveCounts += outDeltaCounts;
    }
    else {
        wrkZeroConsecutiveCounts = 0;
    }

    // Проверяем, нужно ли остановить автомат
    const bool stopConfirmedByWrk =
        (wasAuto && !wrkBit && shdBit && !suppressStopAfterRecentStart && wrkZeroConsecutiveCounts >= 5);
        // был в автомате, 
        // но в автомате сейчас не работает, 
        // И находится в режиме post-shutdown, 
        // И нет отключения останова после недавнего запуска, 
        // И счётчик нулевых _Wrk последовательных отсчётов >= 5
        // Тогда нужно остановить автомат
    
    // Если нужно остановить автомат, т.е. сработало условие остановки, выводим сообщение в лог и готовимфлаги
    if (stopConfirmedByWrk) {   
        // Переход в остановку: пишем в лог дату/время остановки и достигнутую мин. Т рыбы.
        DateTime stopTime = now;   // Время остановки

        String^ stopMessage = nullptr;   // Сообщение остановки
        // Температура в сообщении — только fishCold_C из лога параметров (контроллер).
        if (lastFishCold_C_Valid) {
            stopMessage = String::Format(
                "Information: Останов автоматического алгоритма дефростации. Время: {0:dd.MM.yyyy HH:mm:ss}.\nДостигнутая мин. Т рыбы (fishCold_C): {1:F1} °C",
                stopTime, lastFishCold_C);
        }
        else {
            stopMessage = String::Format(
                "Information: Останов автоматического алгоритма дефростации. Время: {0:dd.MM.yyyy HH:mm:ss}. Нет данных fishCold_C из лога параметров.",
                stopTime);
        }

        // Выводим сообщение в лог
        GlobalLogger::LogMessage(stopMessage);
        if (labelDefrosterState != nullptr && !labelDefrosterState->IsDisposed) {   // Выводим сообщение в окне состояния дефростера
            labelDefrosterState->Text = stopMessage->Replace("Information: ", "");
            labelDefrosterState->ForeColor = System::Drawing::Color::Blue;
        }
        
        // Экспорт в EXCEL запустим только после завершения post-shutdown И 10 стабильных отсчётов _Shd=0 && _Wrk=0.
        stopExportPending = true;   // Устанавливаем флаг экспорта данных в Excel
        shdWrkZeroStableCounts = 0;   // Сбрасываем счётчик стабильных отсчётов _Shd=0 && _Wrk=0
        controllerAutoModeActive = false;   // Сбрасываем флаг автомата
        wrkZeroConsecutiveCounts = 0;   // Сбрасываем счётчик нулевых последовательных отсчётов _Wrk
        // Как при ответе на СТОП с панели: запоздалый лог не должен затирать кнопки в течение 10 с.
        lastStopSuccessTime = now;
    }

    // Условие останова автомата не выполнено, продолжаем запись в таблицу данных, пока активен post-shutdown (_Shd=1).
    if (controllerAutoModeActive) {   // Если автомат активен
        postStopCaptureActive = false;   // Сбрасываем флаг активного post-shutdown
        postStopGateOpenSeen = false;   // Сбрасываем флаг открытия ворот
        postStopCaptureGraceUntil = DateTime::MinValue;   // Сбрасываем время grace-периода
    }
    else {                              // Если автомат не активен
        postStopCaptureActive = shdBit;   // Копируем флаг _Shd в флаг активного post-shutdown
        if (!postStopCaptureActive) {   // Если post-shutdown ещё не активен
            postStopGateOpenSeen = false;   // Сбрасываем флаг открытия ворот
            postStopCaptureGraceUntil = DateTime::MinValue;   // Сбрасываем время grace-периода
        }
    }

    // Обновляем UI программы
    if (this->InvokeRequired) {   // Если вызов требуется в другом потоке
        // Вызываем метод SetProgramStateUi в другом потоке
        this->BeginInvoke(gcnew System::Action<bool>(this, &DataForm::SetProgramStateUi), controllerAutoModeActive);    
    }
    else {   // Если вызов не требуется в другом потоке
        // Вызываем метод SetProgramStateUi в текущем потоке
        SetProgramStateUi(controllerAutoModeActive);
    }
}

// Обработка установки соединения при получении телеметрии (например, при подключении устройства).
void ProjectServerW::DataForm::HandleConnectionSetupOnTelemetry(DateTime now)   // Текущее время
{
    if (reconnectFixationLogPending) {   // Если есть запрос на запись в лог при подключении устройства
        reconnectFixationLogPending = false;   // Сбрасываем флаг запроса на запись в лог при подключении устройства
        GlobalLogger::LogMessage(String::Format(
            "Information: Устройство подключено к порту (Port {1}), GUID: {0}",
            (this->ClientIP != nullptr ? this->ClientIP : ""),
            clientPort));
    }

    if (!firstDataReceived) {   // Если первые данные ещё не получены
        firstDataReceived = true;   // Устанавливаем флаг первых данных получены
        dataCollectionStartTime = now;   // Устанавливаем время начала сбора данных
        dataCollectionEndTime = DateTime::MinValue;   // Сбрасываем время окончания сбора данных
        excelFileName = "WorkData_Port" + clientPort.ToString();   // Устанавливаем имя файла Excel
        dataExportedToExcel = false;   // Сбрасываем флаг экспорта данных в Excel
        SetProgramStateUi(false);   // Устанавливаем UI программы в режим неавтоматического алгоритма
    }

    // Таймаут автоперезапуска: если после STOP прошло 20 минут без экспорта — отменяем ожидание.
    if (autoRestartPending && autoRestartStopIssuedTime != DateTime::MinValue) {
        // Если после STOP прошло 20 минут без экспорта — отменяем ожидание.
        TimeSpan waited = now.Subtract(autoRestartStopIssuedTime);   // Вычисляем время ожидания
        const double safetyTimeoutMinutes = 20;   // Таймаут ожидания в минутах
        if (waited.TotalMinutes >= safetyTimeoutMinutes) {   // Если время ожидания >= таймаута
            autoRestartPending = false;   // Сбрасываем флаг автоперезапуска
            autoRestartStopIssuedTime = DateTime::MinValue;   // Сбрасываем время выдачи STOP
        }
    }
    // Если автоперезапуск не запущен, или время ожидания < таймаута, продолжаем ожидание.
}

// Определение режимов по _Wrk/_Shd и установка внутренних флагов автомата.
void ProjectServerW::DataForm::ResolveModesAndFlushPendingRow(
    DateTime now,
    const MSGQUEUE_OBJ_t& data,             // Структура пакета по контракту (совпадает с контроллером) из ParseBuffer
    bool& outWrkBit,                        // Флаг _Wrk
    bool& outShdBit,                        // Флаг _Shd
    int& outDeltaCounts)                    // Дельта времени из UpdateModeFlagsFromTelemetry
{
    // Определение режимов по _Wrk/_Shd и установка внутренних флагов автомата.
    const uint16_t doBits = data.H[SQ - 1];   // Получаем флаги _Wrk/_Shd из структуры пакета
    outWrkBit = (doBits & (1u << 14)) != 0;   // Флаг _Wrk
    outShdBit = (data.ShutdownActive != 0);   // Флаг _Shd
    outDeltaCounts = 1;   // Дельта времени
    // Обновляем внутренние флаги автомата
    UpdateModeFlagsFromTelemetry(   
        now,                         // Текущее время
        data.Time,                   // Время из структуры пакета
        outWrkBit,                   // Флаг _Wrk
        outShdBit,                   // Флаг _Shd
        outDeltaCounts);              // Дельта времени

    // Если строка предыдущей секунды ждала control-log, но пришла уже новая телеметрия (другое Time),
    // фиксируем старую строку как есть, чтобы не терять телеметрию на переходах _Wrk/_Shd.
    System::Threading::Monitor::Enter(dataTableSync);
    try {
        if (pendingRow != nullptr) {
            bool shouldFlushPending = true;
            try {
                Object^ prevTimeObj = pendingRow["Time"];
                if (prevTimeObj != nullptr && prevTimeObj != System::DBNull::Value) {
                    const uint16_t prevTime = Convert::ToUInt16(prevTimeObj);
                    shouldFlushPending = (prevTime != data.Time);
                }
            }
            catch (...) {
                shouldFlushPending = true;
            }

            if (shouldFlushPending) {
                if (dataGridView != nullptr && !dataGridView->IsDisposed) {
                    dataGridView->SuspendLayout();
                }
                try {
                    dataTable->Rows->Add(pendingRow);
                    dataExportedToExcel = false;
                }
                finally {
                    if (dataGridView != nullptr && !dataGridView->IsDisposed) {
                        dataGridView->ResumeLayout(false);
                    }
                }
                pendingRow = nullptr;
                pendingRowWrkBit = false;
            }
        }
    }
    finally {
        System::Threading::Monitor::Exit(dataTableSync);
    }
}

// Обновление флагов экспорта данных в EXCEL при завершении post-shutdown.
void ProjectServerW::DataForm::UpdateShutdownExportFlags(
    DateTime now,   // Текущее время
    bool wrkBit,   // Флаг _Wrk
    bool shdBit,   // Флаг _Shd
    int deltaCounts)   // Дельта времени
{
    // Если автомат активен, сбрасываем флаг экспорта данных в Excel и счётчик стабильных отсчётов _Shd=0 && _Wrk=0
    if (controllerAutoModeActive) {   // Если автомат активен
        stopExportPending = false;   // Сбрасываем флаг экспорта данных в Excel
        shdWrkZeroStableCounts = 0;   // Сбрасываем счётчик стабильных отсчётов _Shd=0 && _Wrk=0
    }
    else if (stopExportPending) {   // Если флаг экспорта данных в Excel установлен
        if (!postStopCaptureActive && !wrkBit && !shdBit) {   // Если post-shutdown не активен, _Wrk=0 и _Shd=0
            shdWrkZeroStableCounts += deltaCounts;   // Увеличиваем счётчик стабильных отсчётов _Shd=0 && _Wrk=0
        }
        else {   // Если post-shutdown активен, _Wrk=1 или _Shd=1
            shdWrkZeroStableCounts = 0;   // Сбрасываем счётчик стабильных отсчётов _Shd=0 && _Wrk=0
        }

        // Если счётчик стабильных отсчётов _Shd=0 && _Wrk=0 >= 10, запускаем экспорт данных в Excel
        if (shdWrkZeroStableCounts >= 10) {
            if (dataCollectionEndTime == DateTime::MinValue) {   // Если время окончания сбора данных не установлено
                dataCollectionEndTime = now;   // Устанавливаем время окончания сбора данных
            }
            StartExcelExportThread(true);   // Запускаем экспорт данных в Excel
            stopExportPending = false;   // Сбрасываем флаг экспорта данных в Excel
            shdWrkZeroStableCounts = 0;   // Сбрасываем счётчик стабильных отсчётов _Shd=0 && _Wrk=0
        }
    }
}

void ProjectServerW::DataForm::ApplyStopLikeSessionStateReset()
{
    // Сначала снимаем ожидание экспорта с телеметрии, чтобы не запустить второй экспорт в UpdateShutdownExportFlags.
    const bool flushExcelFromPendingStop = stopExportPending;   // Флаг ожидания экспорта данных в Excel
    if (flushExcelFromPendingStop) {
        stopExportPending = false;   // Сбрасываем флаг ожидания экспорта данных в Excel
        shdWrkZeroStableCounts = 0;   // Сбрасываем счётчик стабильных отсчётов _Shd=0 && _Wrk=0
    }

    if (InvokeRequired) {
        Invoke(gcnew System::Action<bool>(this, &DataForm::SetProgramStateUi), false);   // Вызываем метод SetProgramStateUi в другом потоке
    }
    else {
        SetProgramStateUi(false);   // Вызываем метод SetProgramStateUi в текущем потоке
    }
    controllerAutoModeActive = false;   // Сбрасываем флаг автомата
    lastControlLogTime = DateTime::MinValue;   // Сбрасываем время последнего логирования
    wrkZeroConsecutiveCounts = 0;   // Сбрасываем счётчик нулевых последовательных отсчётов
    wrkLastSampleValid = false;   // Сбрасываем признак последнего валидного отсчёта
    lastStopSuccessTime = DateTime::Now;   // Сбрасываем время последнего успешного останова
    postStopCaptureActive = false;   // Сбрасываем флаг активного post-shutdown
    postStopGateOpenSeen = false;   // Сбрасываем флаг открытия ворот
    postStopCaptureGraceUntil = DateTime::MinValue;   // Сбрасываем время grace-периода

    if (flushExcelFromPendingStop) {
        const DateTime now = DateTime::Now;   // Текущее время
        if (dataCollectionEndTime == DateTime::MinValue) {
            dataCollectionEndTime = now;   // Устанавливаем время окончания сбора данных
        }
        try {
            StartExcelExportThread(true);   // Запускаем экспорт данных в Excel
        }
        catch (Exception^ ex) {
            GlobalLogger::LogMessage("Ошибка: экспорт в Excel при STOP/RESET (ожидался по stopExportPending): " + ex->ToString());   // Логируем ошибку
        }
    }
}

void ProjectServerW::DataForm::ResetProgramStateAfterLinkLoss()
{
    if (this == nullptr || this->IsDisposed || this->Disposing)
        return;

    const bool hadAutoSession = controllerAutoModeActive;

    controllerAutoModeActive = false;
    stopExportPending = false;
    shdWrkZeroStableCounts = 0;
    postStopCaptureActive = false;
    postStopGateOpenSeen = false;
    postStopCaptureGraceUntil = DateTime::MinValue;
    wrkZeroConsecutiveCounts = 0;
    lastControlLogTime = DateTime::MinValue;

    System::Threading::Monitor::Enter(dataTableSync);
    try {
        pendingRow = nullptr;
        pendingRowWrkBit = false;
    }
    finally {
        System::Threading::Monitor::Exit(dataTableSync);
    }

    SetProgramStateUi(false);

    if (hadAutoSession) {
        GlobalLogger::LogMessage(
            "Информация: сброс состояния ПУСК/СТОП — нет приёма телеметрии или потерян TCP (контроллер сброшен / не отвечает).");
    }
}

// Обработка выхода из программы.
System::Void ProjectServerW::DataForm::выходToolStripMenuItem_Click(System::Object^ sender, System::EventArgs^ e)
{
    // Ищем GUID текущей формы в formData_Map
    std::wstring currentFormGuid;
    bool formFound = false;   // Флаг найденной формы

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
        GlobalLogger::LogMessage("Не удалось сохранить данные: " + ex->Message);
    }

	return System::Void();
}

// Обработка нажатия кнопки EXCEL.
System::Void ProjectServerW::DataForm::buttonEXCEL_Click(System::Object^ sender, System::EventArgs^ e)
{
    try {
        StartExcelExportThread(false);   // Запускаем экспорт данных в Excel
    }
    catch (Exception^ ex) {
        GlobalLogger::LogMessage("Error: Exception in buttonEXCEL_Click: " + ex->ToString());
    }
}

// Запуск потока экспорта данных в EXCEL.
bool ProjectServerW::DataForm::StartExcelExportThread(bool isEmergency) {   // Флаг аварийного экспорта
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

        // Перед выгрузкой в Excel обновляем аварийные флаги с контроллера (если соединение доступно).
        // Это гарантирует, что лист «АВАРИЯ» содержит актуальное состояние на момент завершения цикла.
        try {
            if (!isFormClosingNow && ClientSocket != INVALID_SOCKET) {
                AlarmFlagsPayload flags{};
                if (GetAlarmFlags(&flags)) {
                    PopulateEquipmentAlarmGrid(flags.deviceAlarmFlags, flags.sensorAlarmFlags);
                }
                else {
                    GlobalLogger::LogMessage("Предупреждение: GET_ALARM_FLAGS перед экспортом в Excel не выполнен, используется текущий снимок таблицы аварий");
                }
            }
        }
        catch (Exception^ ex) {
            GlobalLogger::LogMessage("Предупреждение: не удалось обновить флаги аварий перед экспортом в Excel: " + ex->Message);
        }

        // Создаём задачу экспорта данных в EXCEL
        ProjectServerW::FormExcel::ExcelExportJob^ job = gcnew ProjectServerW::FormExcel::ExcelExportJob();
        job->tableSnapshot = snapshot;

        // Снимок параметров для листа «Параметры»
        job->paramsPhase = BuildPhaseParamsDataTableForExcel();
        job->paramsGlobal = BuildGlobalParamsDataTableForExcel();
        job->equipmentAlarms = BuildEquipmentAlarmDataTableForExcel();
        job->saveDirectory = excelSavePath;
        job->sessionStart = dataCollectionStartTime;
        job->sessionEnd = dataCollectionEndTime;
        job->clientPort = clientPort;
        job->clientIP = this->ClientIP;
        job->formGuid = this->FormGuid;
        job->formRef = gcnew System::WeakReference(this);
        job->enableButtonOnComplete = !isEmergency;

        // Помещаем задачу в очередь экспорта данных в EXCEL
        ProjectServerW::FormExcel::EnqueueExport(job);

        // Выводим сообщение в лог
        if (isEmergency) {
            GlobalLogger::LogMessage("Информация: экспорт данных в Excel поставлен в очередь (аварийное фоновое задание)");
        }
        else {
            GlobalLogger::LogMessage("Информация: экспорт данных в Excel поставлен в очередь");
        }

        // Возвращаем true, если экспорт данных в Excel поставлен в очередь
        return true;
    }
    catch (Exception^ ex) {
        GlobalLogger::LogMessage("Ошибка: не удалось запустить поток экспорта данных в Excel: " + ex->ToString());
        System::Threading::Monitor::Enter(excelExportSync);
        try {
            excelExportInProgress = 0;
        }
        finally {
            System::Threading::Monitor::Exit(excelExportSync);
        }
        return false;
    }
}

// Обработка завершения экспорта данных в EXCEL.
void ProjectServerW::DataForm::OnExcelExportCompleted(bool enableButtonOnComplete) {
    // Снимаем per-form guard — иначе после первого Enqueue все следующие StartExcelExportThread блокируются навсегда.
    try {
        System::Threading::Monitor::Enter(excelExportSync);
        try {
            excelExportInProgress = 0;
        }
        finally {
            System::Threading::Monitor::Exit(excelExportSync);
        }
    }
    catch (...) {}

    if (this == nullptr || this->IsDisposed || this->Disposing) {
        return;
    }

    if (enableButtonOnComplete) {
        try {
            this->BeginInvoke(gcnew MethodInvoker(this, &DataForm::EnableButton));
        }
        catch (Exception^ ex) {
            GlobalLogger::LogMessage("Ошибка: не удалось включить кнопку экспорта данных в Excel: " + ex->ToString());
        }
        catch (...) {
            GlobalLogger::LogMessage("Ошибка: не удалось включить кнопку экспорта данных в Excel (неизвестное исключение)");
        }
    }
}

// Обработка нажатия кнопки Browse.
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
       
// Создание и отображение формы в потоке.
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
        for (int i = 0; i < bitNames[0]->Length; i++)
        {
            dataTable->Columns->Add(bitNames[0][i], uint8_t::typeid);
        }
        // Добавляем в таблицу колонки битовых полей второй группы
        for (int i = 0; i < bitNames[1]->Length; i++)
        {
            dataTable->Columns->Add(bitNames[1][i], uint8_t::typeid);
        }
    }

    // Колонки лога параметров процесса (группа 3): сначала отфильтрованные температуры,
    // затем переменные алгоритма (имена — как в ControlLogPayload_t на контроллере).
    dataTable->Columns->Add("T_filt_0", float::typeid);
    dataTable->Columns->Add("T_filt_1", float::typeid);
    dataTable->Columns->Add("T_filt_2", float::typeid);
    dataTable->Columns->Add("T_filt_3", float::typeid);
    dataTable->Columns->Add("T_filt_4", float::typeid);
    dataTable->Columns->Add("T_filt_5", float::typeid);

    dataTable->Columns->Add("phase", int::typeid);
    dataTable->Columns->Add("eT_common", float::typeid);
    dataTable->Columns->Add("heatScale01", float::typeid);
    dataTable->Columns->Add("uCommon_TEN", float::typeid);
    dataTable->Columns->Add("trim_TEN", float::typeid);
    dataTable->Columns->Add("uLeft_TEN", float::typeid);
    dataTable->Columns->Add("uRight_TEN", float::typeid);
    dataTable->Columns->Add("leftTen1Duty", float::typeid);
    dataTable->Columns->Add("leftTen2Duty", float::typeid);
    dataTable->Columns->Add("rightTen1Duty", float::typeid);
    dataTable->Columns->Add("rightTen2Duty", float::typeid);
    dataTable->Columns->Add("w_sup_avg", float::typeid);
    dataTable->Columns->Add("wErr", float::typeid);
    dataTable->Columns->Add("injDuty", float::typeid);
    dataTable->Columns->Add("rate_Cps", float::typeid);
    dataTable->Columns->Add("fishHot_C", float::typeid);
    dataTable->Columns->Add("fishCold_C", float::typeid);

    // Fix: на некоторых сборках/сценариях Designer может держать "хвост" колонок (например T_filt_6),
    // даже если DataTable содержит только T_filt_0..T_filt_5. Очищаем и заставляем AutoGenerateColumns
    // пересоздать колонки строго по схеме DataTable.
    if (dataGridView != nullptr) {
        dataGridView->AutoGenerateColumns = true;
        dataGridView->Columns->Clear();
        dataGridView->DataSource = nullptr;
    }
    dataGridView->DataSource = dataTable;
    GlobalLogger::LogMessage("Information: DataGridView привязан к DataTable");
}

// 2. Добавление строки
void ProjectServerW::DataForm::AddDataToTable(const char* buffer, size_t size, System::Data::DataTable^ table) {
    
    MSGQUEUE_OBJ_t data;
    DateTime now = DateTime::Now;   // Текущее время сервера
    String^ data_String{};

    // На вход приходит фрейм от контроллера (без AA55):
    // [Type][Cmd][Status][DataLen][Data...][CRC16]
    // Сам массив Data (MSGQUEUE_OBJ_t) начинается с offset=4.
    if (size < 4 + sizeof(data) + 2) {
        return; // недостаточно байт для пакета
    }
    const uint8_t* raw = reinterpret_cast<const uint8_t*>(buffer);
    const uint8_t dataLenByte = raw[3];
    if (dataLenByte != sizeof(data)) {
        // DataLen должен совпадать с размером MSGQUEUE_OBJ_t на сервере.
        return;
    }

    // Всё хорошо, данные в буфере корректны, копируем данные из буфера в структуру data
    memcpy(&data, raw + 4, sizeof(data));

    // Перезапуск контроллера: data.Time (сек с включения МК) не должен резко уменьшаться.
    // Иначе при _Wrk=0 сервер остаётся в controllerAutoModeActive и ветка (active && !wrk) продолжает писать таблицу.
    if (hasTelemetry) {
        const uint16_t prevDeviceSec = lastTelemetryDeviceSeconds;
        if (prevDeviceSec >= 15u) {
            const int dTime = (int)data.Time - (int)prevDeviceSec;
            if (dTime < -8) {
                GlobalLogger::LogMessage(
                    "Информация: перезапуск контроллера (откат поля Time телеметрии). Сброс авторежима; запись в таблицу без новой команды ПУСК отключена.");
                controllerAutoModeActive = false;
                stopExportPending = false;
                shdWrkZeroStableCounts = 0;
                wrkZeroConsecutiveCounts = 0;
                postStopCaptureActive = false;
                postStopGateOpenSeen = false;
                postStopCaptureGraceUntil = DateTime::MinValue;
                lastControlLogTime = DateTime::MinValue;
                System::Threading::Monitor::Enter(dataTableSync);
                try {
                    pendingRow = nullptr;
                    pendingRowWrkBit = false;
                }
                finally {
                    System::Threading::Monitor::Exit(dataTableSync);
                }
                SetProgramStateUi(false);
            }
        }
    }

    data_String = bufferToHex(buffer, size);   // Преобразуем буфер в строку в hex формате
    // Выводим содержимое пакета data_String в поле Label_Data
    DataForm::SetData_TextValue("Текущие данные: " + data_String);

    // Новая строка таблицы
    DataRow^ row = table->NewRow();
    
    // ВРЕМЯ -------------------------------------------------------------
    row["RealTime"] = now.ToString("HH:mm:ss");
    row["Time"] = data.Time;
    // Замечание: время, когда пришла телеметрия с устройства — в секундах с включения устройства.
    lastTelemetryDeviceSeconds = data.Time;
    
    // ТЕМПЕРАТУРЫ
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
    UpdateAllTemperatureValues(temperatures);   // Обновляем температуры в окне

    // БИТОВЫЕ ПОЛЯ -------------------------------------------------------------
    cli::array<cli::array<String^>^>^ bitNames = GetBitFieldNames();   // Получаем имена битовых полей из контроллера
    uint16_t bitField;
    
    // Первая группа (Vent1_Left..But_Stop) — входы DI модуля ввода-вывода. В контроллере: Read_Data_2 → T.
    bitField = data.T[SQ - 1];                           // Получаем биты первой группы (Vent1_Left..But_Stop)
    const bool alrmBit = (bitField & (1u << 15)) != 0;   // Флаг аварии
    // Записываем биты первой группы (Vent1_Left..But_Stop) — DI, тот же битовый порядок, что в модуле.
    for (int bit = 0; bit < 16; bit++) {
        bool bitValue = (bitField & (1 << bit)) != 0;   // Получаем значение бита
        row[bitNames[0][bit]] = bitValue;
    }

    // Вторая группа (_V0.._Alr) — выходы DO модуля ввода-вывода. В контроллере: Read_Data_1 → H.
    bitField = data.H[SQ - 1];                           // Получаем биты второй группы (_V0.._Alr)
    // Записываем биты второй группы (DO) + дополнительный флаг post-shutdown.
    // Порядок полей в таблице: _V0.._Cls, _Wrk, _Shd, _Alr.
    for (int bit = 0; bit <= 14; bit++) {
        bool bitValue = (bitField & (1 << bit)) != 0;
        row[bitNames[1][bit]] = bitValue;
    }
    row[bitNames[1][15]] = (data.ShutdownActive != 0) ? 1 : 0;    // _Shd
    row[bitNames[1][16]] = (bitField & (1 << 15)) != 0;            // _Alr
    //--------------------------------------------------------------------

    // Определяем режимы работы и флаги
    bool wrkBit = false;   // Флаг режима работы
    bool shdBit = false;   // Флаг режима shutdown
    int deltaCounts = 1;   // Количество отсчётов
    ResolveModesAndFlushPendingRow(now, data, wrkBit, shdBit, deltaCounts);   // Определяем режимы работы и флаги
    
    // Если флаг аварии изменился, обновляем флаги аварий
    if (!lastTelemetryAlrmBit && alrmBit) {
        RefreshAlarmFlagsFromController();   // Обновляем флаги аварий из контроллера
    }
    lastTelemetryAlrmBit = alrmBit;   // Сохраняем флаг аварии
    

    // ЗАПИСЬ СТРОКИ -------------------------------------------------------------
    // - при активном _Wrk ждём пакет лога (дополним pendingRow в AppendControlLogToDataRow),
    // - в окне подтверждения STOP (_Wrk=0, но controllerAutoModeActive ещё true) пишем телеметрию сразу,
    // - в послеостановочном хвосте (_Shd=1) пишем телеметрию сразу,
    // - после _Shd=0 пишем ещё 10 стабильных отсчётов (_Wrk=0 && _Shd=0) до автосохранения.
    bool rowTakenForTableOrPending = false;   // Строка уходит в таблицу или в pending (ведётся запись сессии)
    if (controllerAutoModeActive && wrkBit) {               // Авторежим: ждём пакет лога (дополним pendingRow в AppendControlLogToDataRow)
        // pendingRow — это «текущая незавершённая» строка телеметрии (DataRow^), которая ещё не добавлена в dataTable.
        // pendingRow связывает пару «телеметрия + лог алгоритма» в одну строку и страхует от потери данных при смене секунды.
        rowTakenForTableOrPending = true;                   // Строка уходит в таблицу или в pending (ведётся запись сессии)
        pendingRow = row;                                   // Сохраняем строку в pendingRow
        pendingRowWrkBit = true;                            // Сохраняем флаг режима работы в pendingRowWrkBit
    }
    else if ((controllerAutoModeActive && !wrkBit) ||
             postStopCaptureActive ||
             (stopExportPending && !wrkBit && !shdBit)) {   // Если в режиме STOP или послеостановочном хвосте, или в режиме STOP_EXPORT
        rowTakenForTableOrPending = true;
        System::Threading::Monitor::Enter(dataTableSync);   // Захватываем мьютекс для синхронизации доступа к таблице данных
        try {
            if (dataGridView != nullptr && !dataGridView->IsDisposed) {   // Если dataGridView не nullptr и не уничтожена
                // dataGridView->SuspendLayout() — это вызов WinForms: у DataGridView временно отключается пересчёт раскладки и лишние перерисовки, 
                // пока не вызовут ResumeLayout.
                dataGridView->SuspendLayout();              // Приостанавливаем layout (для ускорения добавления строки в таблицу)
            }
            try {   
                dataTable->Rows->Add(row);                  // Добавляем строку без лога алгоритма в таблицу данных
                dataExportedToExcel = false;                // Сбрасываем флаг экспорта данных в Excel (если экспорт был запущен)
            }
            finally {
                if (dataGridView != nullptr && !dataGridView->IsDisposed) {
                    dataGridView->ResumeLayout(false);      // Возобновляем layout
                }
            }
        }
        finally {                                            // Освобождаем мьютекс для синхронизации доступа к таблице данных
            System::Threading::Monitor::Exit(dataTableSync);
        }
        // В этом режиме строка уже зафиксирована в таблице, отложенное ожидание лога не нужно.
        pendingRow = nullptr;
        pendingRowWrkBit = false;
    }
    else {
        // До первого запуска алгоритма (и вне post-shutdown) новые строки в таблицу не пишем.
        // Также не оставляем "висящую" pendingRow, чтобы не было отложенной записи на следующем тике.
        pendingRow = nullptr;
        pendingRowWrkBit = false;
    }
    // Если строка в таблицу не принимается, а UpdateModeFlagsFromTelemetry уже поднял «авторежим» по _Wrk=1,
    // кнопки ПУСК/СТОП остаются в режиме «работа» — приводим UI и флаг к фактическому отсутствию записи.
    if (!rowTakenForTableOrPending) {
        controllerAutoModeActive = false;
        if (this->InvokeRequired) {
            this->BeginInvoke(gcnew System::Action<bool>(this, &DataForm::SetProgramStateUi), false);
        }
        else {
            SetProgramStateUi(false);
        }
    }
    //--------------------------------------------------------------------
    // ОБНОВЛЕНИЕ ФЛАГОВ ЭКСПОРТА ПОСЛЕОСТАНОВОЧНОГО ХВОСТА 
    UpdateShutdownExportFlags(now, wrkBit, shdBit, deltaCounts);   
    //--------------------------------------------------------------------
}


// Связка Excel реализована в FormExcel.cpp

// Отложенная сборка мусора
void DataForm::DelayedGarbageCollection(Object^ state) {
    try {
        // Пауза перед сборкой мусора
        //MessageBox::Show("Ошибка Excel не удалось запуститься!");
        GlobalLogger::LogMessage("Information: delayed garbage collection started");
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

// Обработка таймера простоя формы DataForm без связи с устройством
void ProjectServerW::DataForm::OnInactivityTimerTick(Object^ sender, EventArgs^ e) {
    try {
        if (inactivityCloseRequested) {   // Если запрошено закрытие формы
            return;
        }

        DateTime now = DateTime::Now;   // Текущее время

        // 1) Нет связи с устройством "текущего" окна (по IP) уже 30 минут.
        // Условие: если нет связи более 30 минут с момента отключения по этому IP и адресу порта,
        // то форма считается мёртвой по этому IP адресу порта, и "окно" можно закрыть автоматически.
        if (ClientSocket == INVALID_SOCKET && !String::IsNullOrEmpty(ClientIP)) {
            const long long ticks = System::Threading::Interlocked::Read(disconnectedSinceTicks);
            if (ticks != 0) {
                DateTime disconnectedSince(ticks);
                TimeSpan waited = now.Subtract(disconnectedSince);
                if (waited.TotalMinutes >= 30) {
                    GlobalLogger::LogMessage(String::Format(
                        "Information: Нет соединения с устройством {0} {1:F1} мин. Закрытие формы DataForm.",
                        ClientIP,
                        waited.TotalMinutes));

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

        GlobalLogger::LogMessage(String::Format(
            "Information: Нет телеметрии {0:F1} мин. Закрытие формы DataForm.",
            idle.TotalMinutes));

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
        GlobalLogger::LogMessage("Error: Exception in OnInactivityTimerTick: " + ex->ToString());
    }
}

// Обработка переподключения устройства
void ProjectServerW::DataForm::OnReconnectSendStartupCommands() {
    try {
        if (this == nullptr || this->IsDisposed || this->Disposing || !this->IsHandleCreated) {   // Если форма уничтожена или уничтожается или не создана
            return;
        }

        if (ClientSocket == INVALID_SOCKET) {   // Если сокет невалидный
            return;
        }

        // Замечание: если сокет переподключился с того же устройства по порту и тот же сокет.
        // Замечание: после сброса устройства сокет мог переподключиться к новому сокету, но порт тот же,
        // поэтому сбрасываем флаг по номеру порта UART4.
        if (lastStartupCommandsSocket == ClientSocket) {
            return;
        }

        if (ExecuteStartupCommandSequence()) {
            lastStartupCommandsSocket = ClientSocket;
            postResetInitPending = false;
        }
    }
    catch (Exception^ ex) {
        GlobalLogger::LogMessage("Error: Exception in OnReconnectSendStartupCommands: " + ex->ToString());
    }
    catch (...) {
        GlobalLogger::LogMessage("Error: Unknown exception in OnReconnectSendStartupCommands");
    }
}

// Выполнение последовательности команд запуска устройства
bool ProjectServerW::DataForm::ExecuteStartupCommandSequence() {
    if (ClientSocket == INVALID_SOCKET) {   // Если сокет невалидный, то последовательность команд запуска не выполняется
        startupSequenceCompleted = false;
        return false;
    }

    // ПОЛУЧЕНИЕ ВЕРСИИ УСТРОЙСТВА -------------------------------------------------------------
    GlobalLogger::LogMessage("Отладка: Последовательность запуска: шаг 1/3 GET_VERSION начат");
    const bool okVersion = SendVersionRequest();
    if (!okVersion) {
        startupSequenceCompleted = false;
        if (sendStateTimer != nullptr) sendStateTimer->Stop();
        GlobalLogger::LogMessage("Предупреждение: Последовательность запуска прервана на шаге GET_VERSION");
        return false;
    }
    GlobalLogger::LogMessage("Отладка: Последовательность запуска: шаг 1/3 GET_VERSION выполнен");

    // УСТАНОВКА ИНТЕРВАЛА ИЗМЕРЕНИЙ -------------------------------------------------------------
    int intervalSeconds = 10;
    if (numericUpDownMeasurementInterval != nullptr && !numericUpDownMeasurementInterval->IsDisposed) {
        intervalSeconds = System::Decimal::ToInt32(numericUpDownMeasurementInterval->Value);
    }
    GlobalLogger::LogMessage(String::Format("Отладка: Последовательность запуска: шаг 2/3 SET_INTERVAL({0}) начат", intervalSeconds));
    const bool okInterval = SendSetIntervalCommand(intervalSeconds);
    if (!okInterval) {
        startupSequenceCompleted = false;
        if (sendStateTimer != nullptr) sendStateTimer->Stop();
        GlobalLogger::LogMessage("Предупреждение: Последовательность запуска прервана на шаге SET_INTERVAL");
        return false;
    }
    GlobalLogger::LogMessage("Отладка: Последовательность запуска: шаг 2/3 SET_INTERVAL выполнен");

    // ПОЛУЧЕНИЕ ГРУППЫ ПАРАМЕТРОВ 5 И 6 -------------------------------------------------------------
    GlobalLogger::LogMessage("Отладка: Последовательность запуска: шаг 3/3 GET_DEFROST_GROUP(5,6) начат");
    uint8_t buffer[256];   // Буфер для данных
    const uint8_t kPayloadCapacity = 255;   // Максимальная длина данных
    uint8_t len = 0;   // Длина данных
    bool loaded1 = false;   // Флаг загрузки данных группы 5
    bool loaded2 = false;   // Флаг загрузки данных группы 6

    // ПОЛУЧЕНИЕ ДАННЫХ ГРУППЫ 5 -------------------------------------------------------------
    if (dataGridView1 != nullptr && !dataGridView1->IsDisposed) {   // Если dataGridView1 не уничтожена и не уничтожается
        const bool ok1 = GetDefrostGroup(5, 0, buffer, kPayloadCapacity, &len);   // Получаем данные группы 5
        if (ok1 && len >= (uint8_t)sizeof(DefrostLogPhasePayload_t)) {
            FillDataGridView1FromGroup5Payload(buffer, len);   // Заполняем dataGridView1 данными группы 5
            loaded1 = true;
        }
        else {
            GlobalLogger::LogMessage(String::Format("Предупреждение: Стартовая загрузка группы 5 не выполнена (transport={0}, len={1})", ok1 ? "ok" : "fail", (int)len));
        }
    }
    // ПОЛУЧЕНИЕ ДАННЫХ ГРУППЫ 6 -------------------------------------------------------------
    if (dataGridView2 != nullptr && !dataGridView2->IsDisposed) {   // Если dataGridView2 не уничтожена и не уничтожается
        const bool ok2 = GetDefrostGroup(6, 0, buffer, kPayloadCapacity, &len);   // Получаем данные группы 6
        if (ok2 && len >= (uint8_t)sizeof(DefrostLogGlobalPayload_t)) {
            FillDataGridView2FromGroup6Payload(buffer, len);   // Заполняем dataGridView2 данными группы 6
            loaded2 = true;
        }
        else {
            GlobalLogger::LogMessage(String::Format("Предупреждение: Стартовая загрузка группы 6 не выполнена (transport={0}, len={1})", ok2 ? "ok" : "fail", (int)len));
        }
    }

    // ОПРЕДЕЛЕНИЕ ФЛАГА ВЫПОЛНЕНИЯ ПОСЛЕОСТАНОВОЧНОЙ ПОСЛЕДОВАТЕЛЬНОСТИ -------------------------------------------------------------
    startupSequenceCompleted = (loaded1 && loaded2);
    paramsLoadedFromDevice = startupSequenceCompleted;
    if (!startupSequenceCompleted) {
        if (sendStateTimer != nullptr) sendStateTimer->Stop();
        GlobalLogger::LogMessage("Предупреждение: Последовательность запуска прервана на шаге GET_DEFROST_GROUP(5,6)");
        return false;
    }

    // ЗАПУСК ТАЙМЕРА СОСТОЯНИЯ -------------------------------------------------------------   
    if (sendStateTimer != nullptr) {
        sendStateTimer->Start();
    }
    GlobalLogger::LogMessage("Информация: Последовательность запуска выполнена (GET_VERSION -> SET_INTERVAL -> GET_DEFROST_GROUP[5,6])");
    return true;
}

// Запланировать отложенный запуск после переподключения устройства
void ProjectServerW::DataForm::ScheduleDeferredStartupOnReconnect() {
    SchedulePostResetInit();
}

// Запланировать отложенный запуск после переподключения устройства
void ProjectServerW::DataForm::SchedulePostResetInit() {
    if (this == nullptr || this->IsDisposed || this->Disposing) {   // Если форма уничтожена или уничтожается или не создана
        return;
    }

    postResetInitPending = true;   // Флаг запланированного отложенного запуска
    postResetInitAttempt = 0;
    postResetInitDeadline = DateTime::Now.AddSeconds(60);   // Срок выполнения отложенного запуска

    // Замечание: т.к. устройство могло получить RESET от контроллера, сбрасываем признак отправленных команд/параметров.
    lastStartupCommandsSocket = INVALID_SOCKET;   // Сбрасываем признак отправленных команд/параметров

    if (postResetInitTimer == nullptr) {
        postResetInitTimer = gcnew System::Windows::Forms::Timer();   // Создаём таймер для отложенного запуска
        postResetInitTimer->Tick += gcnew EventHandler(this, &DataForm::OnPostResetInitTimerTick);
    }

    postResetInitTimer->Stop();   // Останавливаем таймер для отложенного запуска
    // После RESET контроллер уже через ~1 с шлёт телеметрию, поэтому не тянем время: первая попытка через 1 с.
    postResetInitTimer->Interval = 1000;
    postResetInitTimer->Start();   // Запускаем таймер для отложенного запуска
}

// Обработка таймера отложенного запуска после переподключения устройства
void ProjectServerW::DataForm::OnPostResetInitTimerTick(System::Object^ sender, System::EventArgs^ e) {
    try {
        if (postResetInitTimer != nullptr) {   // Если таймер для отложенного запуска существует
            postResetInitTimer->Stop();
        }

        if (!postResetInitPending) {   // Если отложенный запуск не запланирован
            return;
        }

        // ПРОВЕРКА ФОРМЫ -------------------------------------------------------------
        if (this == nullptr || this->IsDisposed || this->Disposing || !this->IsHandleCreated) {
            postResetInitPending = false;   // Сбрасываем флаг запланированного отложенного запуска
            return;
        }

        // ПРОВЕРКА СРОКА ВЫПОЛНЕНИЯ ОТЛОЖЕННОГО ЗАПУСКА -------------------------------------------------------------
        DateTime now = DateTime::Now;
        if (postResetInitDeadline != DateTime::MinValue && now >= postResetInitDeadline) {
            postResetInitPending = false;
            return;
        }

        // ПРОВЕРКА СОЕДИНЕНИЯ TCP -------------------------------------------------------------
        if (ClientSocket == INVALID_SOCKET) {
            // Нет соединения TCP (ожидание установки соединения сокет ещё/повтор).
            if (postResetInitTimer != nullptr) {
                postResetInitTimer->Interval = 2000;
                postResetInitTimer->Start();
            }
            return;
        }

        // ПРОВЕРКА ПОПЫТОК ВЫПОЛНЕНИЯ ОТЛОЖЕННОГО ЗАПУСКА -------------------------------------------------------------
        postResetInitAttempt++;   // Увеличиваем счётчик попыток выполнения отложенного запуска

        if (ExecuteStartupCommandSequence()) {   // Выполняем последовательность команд запуска устройства
            postResetInitPending = false;   // Сбрасываем флаг запланированного отложенного запуска
            lastStartupCommandsSocket = ClientSocket;   // Сохраняем сокет для отправки команд/параметров
            return;
        }

        // ПОВТОР ЧЕРЕЗ ТАЙМЕР; ОЖИДАНИЕ ОТВЕТА ПО ПОРТУ UART ДОЛЬШЕ, ЧЕМ TCP. -------------------------------------------------------------
        if (postResetInitTimer != nullptr) {
            postResetInitTimer->Interval = 5000;
            postResetInitTimer->Start();
        }
    }
    // ОБРАБОТКА ИСКЛЮЧЕНИЙ -------------------------------------------------------------
    catch (Exception^ ex) {
        GlobalLogger::LogMessage("Ошибка: Исключение в OnPostResetInitTimerTick: " + ex->ToString());
    }
    catch (...) {
        GlobalLogger::LogMessage("Ошибка: Неизвестное исключение в OnPostResetInitTimerTick");
    }
}

/* Добавление данных в таблицу в потоке приёмника SServer::ClientHandler
* Потоки ClientHandler (приёмник данных) и DataForm (сохранение данных в таблицу) создаются для каждого нового Socket клиента, 
* это разные потоки, поэтому данные добавляются в таблицу из потока приёмника в поток формы через Invoke.
* Эта функция — потокобезопасный входной шлюз для добавления телеметрии в таблицу формы.
* По сути: это “обёртка” между фоновым приёмом данных и UI-таблицей, чтобы данные добавлялись корректно и без гонок потоков.
*/
void ProjectServerW::DataForm::AddDataToTableThreadSafe(cli::array<System::Byte>^ buffer, int size, int port) {
    // Вызов приходит из потока приёмника, выполняется через Invoke
    // Важно! Данные, пришедшие в UI, обрабатываются в контексте формы, обновляем данные в UI!
    // Устанавливаем порт клиента в контексте DataForm
    this->clientPort = port;

    // Блокировка управляемого массива в неуправляемой памяти (pin_ptr для передачи в native-код)
    pin_ptr<Byte> pinnedBuffer = &buffer[0];   // Блокируем массив байт в неуправляемой памяти
    char* rawBuffer = reinterpret_cast<char*>(pinnedBuffer);   // Преобразуем массив байт в строку

    try {
        if (this == nullptr || this->IsDisposed || this->Disposing) {   // Если форма уничтожена или уничтожается или не создана
            return;
        }
        if (dataGridView == nullptr || dataGridView->IsDisposed) {   // Если dataGridView не существует или уничтожена
            return;
        }

        DateTime telemetryTime = DateTime::Now;   // Время приёма данных
        const SOCKET currentSocket = clientSocket;   // Текущий сокет клиента
        if (hasTelemetry && lastTelemetrySocket != INVALID_SOCKET && currentSocket != INVALID_SOCKET &&
            currentSocket != lastTelemetrySocket) {   
            // Если есть телеметрия и сокет невалидный или текущий сокет не равен последнему сокету телеметрии
            // Замечание: переподключение после отключения в SServer. Когда сокет переподключился к другому
            // "текущему устройству" по новому сокету отображаем сообщение о переподключении TCP-соединения.
            reconnectFixationLogPending = true;   // Устанавливаем флаг ожидания переподключения
        }

        hasTelemetry = true;   // Устанавливаем флаг наличия телеметрии
        lastTelemetryTime = telemetryTime;   // Сохраняем время последнего приёма телеметрии
        lastTelemetrySocket = currentSocket;   // Сохраняем сокет последнего приёма телеметрии
        HandleConnectionSetupOnTelemetry(telemetryTime);   // Обработка установки соединения на телеметрию

        // Обновляем время начала цикла (новая сессия сбора данных) при первом приёме после переподключения.
        // Это время потом используется как «Start» в имени/метаданных Excel-экспорта для текущего цикла данных.
        if (dataCollectionStartTime == DateTime::MinValue) {   // Если время начала цикла не установлено
            dataCollectionStartTime = telemetryTime;   // Устанавливаем время начала цикла
        }

        dataGridView->SuspendLayout();   // Приостанавливаем обновление таблицы
        // Зачем: нужна, чтобы временно отключить перерасчёт/перерисовку UI таблицы на время пачки операций с таблицей.
        try {
            // Замечание: вызов Add в Snapshot потоке уже потокобезопасен.
            // Входим в синхронизатор таблицы
            // Это обычный объект-замок (System::Object^), который используется как mutex/lock для доступа к dataTable.
            System::Threading::Monitor::Enter(dataTableSync);   
            // Входим в блок синхронизации
            try {
                // Добавление в таблицу выполняется в UI-потоке формы DataForm для конкретного соединения
                AddDataToTable(rawBuffer, size, dataTable);   // Добавляем данные в таблицу
            }
            finally {
                System::Threading::Monitor::Exit(dataTableSync);   // Выходим из синхронизатора таблицы
            }
        }
        finally {
            dataGridView->ResumeLayout(false);   // Возобновляем обновление таблицы
        }
    }
    catch (Exception^ ex) {
        GlobalLogger::LogMessage("Ошибка: Исключение в AddDataToTableThreadSafe: " + ex->ToString());
    }

}

// ЛОГ ПАРАМЕТРОВ (ГРУППА 3) И ДОБАВЛЕНИЕ ЛОГА В СТРОКУ ДАННЫХ -------------------------------------------------------------
// Парсинг лога параметров (группа 3) и добавление лога в строку данных: дополняет накопленную строку
// телеметрии параметрами лога и добавляет эту строку в таблицу (только при приходе лога, авто-режим).
void ProjectServerW::DataForm::AppendControlLogToDataRow(cli::array<System::Byte>^ packet, int size) {
    const int expectedPayloadSize = (int)sizeof(ControlLogPayload_t);   // Ожидаемый размер полезной нагрузки
    // Входной массив содержит: [Type][Cmd][Status][DataLen][Data...][CRC16]
    const int minPacketSize = 4 + expectedPayloadSize + 2;   // Минимальный размер пакета
    if (packet == nullptr || size < minPacketSize) {   // Если пакет отсутствует или размер меньше минимального размера пакета
        GlobalLogger::LogMessage(String::Format(
            "Предупреждение: AppendControlLogToDataRow: пакет отсутствует или размер {0} < {1}", size, minPacketSize));   // Лог ошибки
        return;
    }
    const int dataLenByte = (int)packet[3];
    if (dataLenByte != expectedPayloadSize) {   // Если байт DataLen не равен ожидаемому размеру полезной нагрузки
        GlobalLogger::LogMessage(String::Format(
            "Предупреждение: AppendControlLogToDataRow: байт DataLen {0} != sizeof(ControlLogPayload_t)={1}, пакет отбрасывается", dataLenByte, expectedPayloadSize));   // Лог ошибки
        return;
    }

    // Блокировка управляемого массива в неуправляемой памяти (pin_ptr для передачи в native-код)
    pin_ptr<System::Byte> pinned = &packet[0];   // Блокируем массив байт в неуправляемой памяти
    const uint8_t* raw = reinterpret_cast<const uint8_t*>(pinned);   // Преобразуем массив байт в строку
    ControlLogPayload_t pl;   // Структура для хранения данных лога
    memcpy(&pl, raw + 4, sizeof(ControlLogPayload_t));   // Копируем данные лога в структуру pl начиная с 4-го байта
    lastFishCold_C = pl.fishCold_C;   // Сохраняем последнее значение параметра fishCold_C
    lastFishCold_C_Valid = true;   // Устанавливаем флаг валидности последнего значения параметра fishCold_C

    // Дополняем накопленную строку телеметрии (pendingRow) параметрами лога; при _Wrk=1 добавляем строку в таблицу.
    // Входим в синхронизатор таблицы
    // Это обычный объект-замок (System::Object^), который используется как mutex/lock для доступа к dataTable.
    System::Threading::Monitor::Enter(dataTableSync);
    try {
        if (dataTable == nullptr) {   // Если таблица не существует
            GlobalLogger::LogMessage("Warning: AppendControlLogToDataRow: dataTable == nullptr");
            return;
        }
        // Получаем строку для обновления
        DataRow^ rowToUpdate = pendingRow;
        const bool shouldAddToTable = pendingRowWrkBit;   // Флаг добавления строки в таблицу
        pendingRow = nullptr;   // Сбрасываем строку, освобождаем для новой строки
        pendingRowWrkBit = false;   // Сбрасываем флаг добавления строки в таблицу

        if (rowToUpdate == nullptr) {   // Если строка для обновления не существует
            // Лог пришёл без телеметрии: создаём строку только с логом.
            rowToUpdate = dataTable->NewRow();   // Создаём новую строку
            for (int c = 0; c < dataTable->Columns->Count; c++) {
                rowToUpdate[c] = System::DBNull::Value;   // Заполняем строку пустыми значениями
            }
        }
        // Заполняем параметрами из лога
        rowToUpdate["T_filt_0"] = pl.T_filt_C[0];
        rowToUpdate["T_filt_1"] = pl.T_filt_C[1];
        rowToUpdate["T_filt_2"] = pl.T_filt_C[2];
        rowToUpdate["T_filt_3"] = pl.T_filt_C[3];
        rowToUpdate["T_filt_4"] = pl.T_filt_C[4];
        rowToUpdate["T_filt_5"] = pl.T_filt_C[5];

        rowToUpdate["phase"] = (int)pl.phase;
        rowToUpdate["eT_common"] = pl.eT_common;
        rowToUpdate["heatScale01"] = pl.heatScale01;
        rowToUpdate["uCommon_TEN"] = pl.uCommon_TEN;
        rowToUpdate["trim_TEN"] = pl.trim_TEN;
        rowToUpdate["uLeft_TEN"] = pl.uLeft_TEN;
        rowToUpdate["uRight_TEN"] = pl.uRight_TEN;
        rowToUpdate["leftTen1Duty"] = pl.leftTen1Duty;
        rowToUpdate["leftTen2Duty"] = pl.leftTen2Duty;
        rowToUpdate["rightTen1Duty"] = pl.rightTen1Duty;
        rowToUpdate["rightTen2Duty"] = pl.rightTen2Duty;
        rowToUpdate["w_sup_avg"] = pl.w_sup_avg;
        rowToUpdate["wErr"] = pl.wErr;
        rowToUpdate["injDuty"] = pl.injDuty;
        rowToUpdate["rate_Cps"] = pl.rate_Cps;
        rowToUpdate["fishHot_C"] = pl.fishHot_C;
        rowToUpdate["fishCold_C"] = pl.fishCold_C;

        // Если флаг добавления строки в таблицу установлен, добавляем строку в таблицу
        if (shouldAddToTable) {
            if (dataGridView != nullptr && !dataGridView->IsDisposed) {   // Если dataGridView существует И НЕ уничтожена
                dataGridView->SuspendLayout();   // Приостанавливаем обновление таблицы
            }
            try {
                dataTable->Rows->Add(rowToUpdate);   // Добавляем строку в таблицу
                dataExportedToExcel = false;   // Сбрасываем флаг экспорта данных в Excel
            }
            finally {
                if (dataGridView != nullptr && !dataGridView->IsDisposed) {   // Если dataGridView существует И НЕ уничтожена
                    dataGridView->ResumeLayout(false);   // Возобновляем обновление таблицы
                }
            }
        }
    }
    // Выходим из синхронизатора таблицы
    catch (Exception^ ex) {
        GlobalLogger::LogMessage("Error: AppendControlLogToDataRow exception: " + ex->ToString());   // Лог ошибки
    }
    finally {
        System::Threading::Monitor::Exit(dataTableSync);   // Выходим из синхронизатора таблицы
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
        GlobalLogger::LogMessage("Не удалось сохранить настройки в файл: " + ex->Message);
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
                            if (sendStateTimer != nullptr)
                                sendStateTimer->Interval = Math::Max(1000, parsed * 1000);
                        }
                    }
                }
            }
        }
    }
    catch (Exception^ ex) {
        // Ошибка при загрузке настроек
        MessageBox::Show("Не удалось загрузить настройки: " + ex->Message);
        GlobalLogger::LogMessage("Не удалось загрузить настройки: " + ex->Message);
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
    dataGridView2->Rows->Add("fishColdTarget_C", "Target min fish temp °C (auto-stop when reached)", "6");
    dataGridView2->Rows->Add("debugDisableTargetTStop", "Debug: disable auto-stop by fishColdTarget_C (0/1)", "0");
    dataGridView2->Rows->Add("debugDisableDeviceSwitchCheck", "Debug: disable DO->DI switch check (0/1)", "0");
}

// Имена параметров группы 5 (LOG_PHASE): порядок полей в DefrostLogPhasePayload_t
static const char* const kGroup5ParamNames[] = {
    "fishHotMax_C", "fishHotRateMax_Cps", "fishDeltaMax_C", "supplySet_C", "supplyMax_C", "returnTargetRH_percent"
};

/** Заполнить payload группы 5 из dataGridView1 (порядок строк 0..5 = fishHotMax_C, fishHotRateMax_Cps, ...). */
static bool BuildPhasePayloadFromGrid1(DataGridView^ grid, DefrostLogPhasePayload_t* outPayload) {
    if (grid == nullptr || grid->IsDisposed || outPayload == nullptr) return false;
    memset(outPayload, 0, sizeof(DefrostLogPhasePayload_t));
    System::Globalization::CultureInfo^ inv = System::Globalization::CultureInfo::InvariantCulture;
    for (int r = 0; r < 6 && r < grid->Rows->Count; r++) {
        DataGridViewRow^ row = grid->Rows[r];
        if (row->IsNewRow) continue;
        for (int ph = 0; ph < 3; ph++) {
            String^ colName = (ph == 0) ? "WarmUP" : (ph == 1) ? "Plateau" : "Finish";
            Object^ v = row->Cells[colName]->Value;
            String^ s = v != nullptr ? v->ToString()->Trim()->Replace(",", ".") : "";
            float f = 0.0f;
            Single::TryParse(s, System::Globalization::NumberStyles::Float, inv, f);
            if (r == 0) outPayload->fishHotMax_C[ph] = f;
            else if (r == 1) outPayload->fishHotRateMax_Cps[ph] = f;
            else if (r == 2) outPayload->fishDeltaMax_C[ph] = f;
            else if (r == 3) outPayload->supplySet_C[ph] = f;
            else if (r == 4) outPayload->supplyMax_C[ph] = f;
            else if (r == 5) outPayload->returnTargetRH_percent[ph] = f;
        }
    }
    return true;
}

void ProjectServerW::DataForm::FillDataGridView1FromGroup5Payload(const uint8_t* payload, uint8_t payloadLen) {
    if (dataGridView1 == nullptr || dataGridView1->IsDisposed || payload == nullptr) return;
    const size_t expectedSize = sizeof(DefrostLogPhasePayload_t);
    if ((size_t)payloadLen < expectedSize) {
        GlobalLogger::LogMessage(String::Format(
            "Warning: Group5 payload too short: got {0}, expected {1}. Keep previous table values.",
            (int)payloadLen, (int)expectedSize));
        return;
    }
    dataGridView1->Rows->Clear();
    DefrostLogPhasePayload_t s;
    memcpy(&s, payload, expectedSize);
    const float* phaseRows[6] = {
        s.fishHotMax_C, s.fishHotRateMax_Cps, s.fishDeltaMax_C,
        s.supplySet_C, s.supplyMax_C, s.returnTargetRH_percent
    };
    System::Globalization::CultureInfo^ inv = System::Globalization::CultureInfo::InvariantCulture;
    for (int row = 0; row < 6; row++) {
        cli::array<Object^>^ rowData = gcnew cli::array<Object^>(4);
        rowData[0] = gcnew System::String(kGroup5ParamNames[row]);
        rowData[1] = phaseRows[row][0].ToString(inv);
        rowData[2] = phaseRows[row][1].ToString(inv);
        rowData[3] = phaseRows[row][2].ToString(inv);
        dataGridView1->Rows->Add(rowData);
    }
    dataGridView1Dirty = false;
}

// Имена и описания для группы 6 (LOG_GLOBAL): порядок полей в DefrostLogGlobalPayload_t
struct Group6Row { const char* name; const char* desc; };
static const Group6Row kGroup6Rows[] = {
    {"leftRightTrimGain", "Heater trim per degC difference"},
    {"leftRightTrimMaxEq", "Max trim heater equivalent"},
    {"piKp", "PI proportional gain (supply temp)"},
    {"piKi", "PI integral gain (supply temp)"},
    {"wDeadband_kgkg", "Humidity deadband"},
    {"injGain", "Injection gain (duty per kg/kg humidity error)"},
    {"outDamperTimer_s", "Damper open time"},
    {"outFanDelay_s", "Fan delay after damper open"},
    {"outHold_s", "Damper and exhaust fan min hold"},
    {"tenMinHold_s", "Heater min hold between switches"},
    {"injMinHold_s", "Injector min hold between switches"},
    {"airOnlyPhaseWarmUp_s", "Air-only: WarmUp phase duration (s)"},
    {"airOnlyPhasePlateau_s", "Air-only: Plateau end time from start (s)"},
    {"maxRuntime_s", "Air-only: Max process duration (s)"},
    {"fishColdTarget_C", "Target min fish temp °C (auto-stop when reached)"},
    {"debugDisableTargetTStop", "Debug: disable auto-stop by fishColdTarget_C (0/1)"},
    {"debugDisableDeviceSwitchCheck", "Debug: disable DO->DI switch check (0/1)"}
};

/** Заполнить payload группы 6 из dataGridView2 по имени параметра (Parameter2). */
static bool BuildGlobalPayloadFromGrid2(DataGridView^ grid, DefrostLogGlobalPayload_t* outPayload) {
    if (grid == nullptr || grid->IsDisposed || outPayload == nullptr) return false;
    memset(outPayload, 0, sizeof(DefrostLogGlobalPayload_t));
    System::Globalization::CultureInfo^ inv = System::Globalization::CultureInfo::InvariantCulture;
    for (int r = 0; r < grid->Rows->Count; r++) {
        DataGridViewRow^ row = grid->Rows[r];
        if (row->IsNewRow) continue;
        Object^ nameObj = row->Cells["Parameter2"]->Value;
        Object^ valueObj = row->Cells["Value"]->Value;
        String^ name = nameObj != nullptr ? nameObj->ToString()->Trim() : "";
        String^ valueStr = valueObj != nullptr ? valueObj->ToString()->Trim()->Replace(",", ".") : "";
        if (String::IsNullOrEmpty(name)) continue;
        if (name->Equals("leftRightTrimGain", StringComparison::OrdinalIgnoreCase)) { float f; if (Single::TryParse(valueStr, System::Globalization::NumberStyles::Float, inv, f)) outPayload->leftRightTrimGain = f; continue; }
        if (name->Equals("leftRightTrimMaxEq", StringComparison::OrdinalIgnoreCase)) { float f; if (Single::TryParse(valueStr, System::Globalization::NumberStyles::Float, inv, f)) outPayload->leftRightTrimMaxEq = f; continue; }
        if (name->Equals("piKp", StringComparison::OrdinalIgnoreCase)) { float f; if (Single::TryParse(valueStr, System::Globalization::NumberStyles::Float, inv, f)) outPayload->piKp = f; continue; }
        if (name->Equals("piKi", StringComparison::OrdinalIgnoreCase)) { float f; if (Single::TryParse(valueStr, System::Globalization::NumberStyles::Float, inv, f)) outPayload->piKi = f; continue; }
        if (name->Equals("wDeadband_kgkg", StringComparison::OrdinalIgnoreCase)) { float f; if (Single::TryParse(valueStr, System::Globalization::NumberStyles::Float, inv, f)) outPayload->wDeadband_kgkg = f; continue; }
        if (name->Equals("injGain", StringComparison::OrdinalIgnoreCase)) { float f; if (Single::TryParse(valueStr, System::Globalization::NumberStyles::Float, inv, f)) outPayload->injGain = f; continue; }
        if (name->Equals("outDamperTimer_s", StringComparison::OrdinalIgnoreCase)) { unsigned int u; if (UInt32::TryParse(valueStr, u)) outPayload->outDamperTimer_s = (uint16_t)(u & 0xFFFF); continue; }
        if (name->Equals("outFanDelay_s", StringComparison::OrdinalIgnoreCase)) { unsigned int u; if (UInt32::TryParse(valueStr, u)) outPayload->outFanDelay_s = (uint16_t)(u & 0xFFFF); continue; }
        if (name->Equals("outHold_s", StringComparison::OrdinalIgnoreCase)) { unsigned int u; if (UInt32::TryParse(valueStr, u)) outPayload->outHold_s = (uint16_t)(u & 0xFFFF); continue; }
        if (name->Equals("tenMinHold_s", StringComparison::OrdinalIgnoreCase)) { unsigned int u; if (UInt32::TryParse(valueStr, u)) outPayload->tenMinHold_s = (uint16_t)(u & 0xFFFF); continue; }
        if (name->Equals("injMinHold_s", StringComparison::OrdinalIgnoreCase)) { unsigned int u; if (UInt32::TryParse(valueStr, u)) outPayload->injMinHold_s = (uint16_t)(u & 0xFFFF); continue; }
        if (name->Equals("airOnlyPhaseWarmUp_s", StringComparison::OrdinalIgnoreCase)) { unsigned int u; if (UInt32::TryParse(valueStr, u)) outPayload->airOnlyPhaseWarmUp_s = (uint16_t)(u & 0xFFFF); continue; }
        if (name->Equals("airOnlyPhasePlateau_s", StringComparison::OrdinalIgnoreCase)) { unsigned int u; if (UInt32::TryParse(valueStr, u)) outPayload->airOnlyPhasePlateau_s = (uint16_t)(u & 0xFFFF); continue; }
        if (name->Equals("maxRuntime_s", StringComparison::OrdinalIgnoreCase)) { unsigned int u; if (UInt32::TryParse(valueStr, u)) outPayload->maxRuntime_s = (uint16_t)(u & 0xFFFF); continue; }
        if (name->Equals("fishColdTarget_C", StringComparison::OrdinalIgnoreCase)) { float f; if (Single::TryParse(valueStr, System::Globalization::NumberStyles::Float, inv, f)) outPayload->fishColdTarget_C = f; continue; }
        if (name->Equals("debugDisableTargetTStop", StringComparison::OrdinalIgnoreCase)) { unsigned int u; if (UInt32::TryParse(valueStr, u)) outPayload->debugDisableTargetTStop = (uint8_t)(u & 0xFF); continue; }
        if (name->Equals("debugDisableDeviceSwitchCheck", StringComparison::OrdinalIgnoreCase)) { unsigned int u; if (UInt32::TryParse(valueStr, u)) outPayload->debugDisableDeviceSwitchCheck = (uint8_t)(u & 0xFF); continue; }
        for (int i = 0; i <= 5; i++) {
            String^ sensorName = System::String::Format("Sensor{0} use in defrost", i);
            if (name->Equals(sensorName, StringComparison::OrdinalIgnoreCase)) {
                unsigned int u; if (UInt32::TryParse(valueStr, u)) outPayload->sensorUseInDefrost[i] = (uint8_t)(u & 0xFF);
                break;
            }
        }
    }
    return true;
}

void ProjectServerW::DataForm::FillDataGridView2FromGroup6Payload(const uint8_t* payload, uint8_t payloadLen) {
    if (dataGridView2 == nullptr || dataGridView2->IsDisposed || payload == nullptr) return;
    const size_t expectedSize = sizeof(DefrostLogGlobalPayload_t);
    if ((size_t)payloadLen < expectedSize) {
        GlobalLogger::LogMessage(String::Format(
            "Warning: Group6 payload too short: got {0}, expected {1}. Keep previous table values.",
            (int)payloadLen, (int)expectedSize));
        return;
    }
    dataGridView2->Rows->Clear();
    DefrostLogGlobalPayload_t s;
    memcpy(&s, payload, expectedSize);
    System::Globalization::CultureInfo^ inv = System::Globalization::CultureInfo::InvariantCulture;
    const float* fFields[] = { &s.leftRightTrimGain, &s.leftRightTrimMaxEq, &s.piKp, &s.piKi, &s.wDeadband_kgkg, &s.injGain };
    const uint16_t* u16Fields[] = { &s.outDamperTimer_s, &s.outFanDelay_s, &s.outHold_s, &s.tenMinHold_s, &s.injMinHold_s, &s.airOnlyPhaseWarmUp_s, &s.airOnlyPhasePlateau_s, &s.maxRuntime_s };
    for (int i = 0; i < 6; i++)
        dataGridView2->Rows->Add(gcnew System::String(kGroup6Rows[i].name), gcnew System::String(kGroup6Rows[i].desc), (*fFields[i]).ToString(inv));
    for (int i = 0; i < 8; i++)
        dataGridView2->Rows->Add(gcnew System::String(kGroup6Rows[6 + i].name), gcnew System::String(kGroup6Rows[6 + i].desc), System::Convert::ToString((int)(*u16Fields[i])));
    dataGridView2->Rows->Add(gcnew System::String(kGroup6Rows[14].name), gcnew System::String(kGroup6Rows[14].desc), s.fishColdTarget_C.ToString(inv));
    dataGridView2->Rows->Add(gcnew System::String(kGroup6Rows[15].name), gcnew System::String(kGroup6Rows[15].desc), System::Convert::ToString((int)s.debugDisableTargetTStop));
    dataGridView2->Rows->Add(gcnew System::String(kGroup6Rows[16].name), gcnew System::String(kGroup6Rows[16].desc), System::Convert::ToString((int)s.debugDisableDeviceSwitchCheck));
    for (int i = 0; i <= 5; i++) {
        dataGridView2->Rows->Add(
            System::String::Format("Sensor{0} use in defrost", i),
            System::String::Format("Use sensor {0} in algorithm", i),
            System::Convert::ToString((int)s.sensorUseInDefrost[i]));
    }
    dataGridView2Dirty = false;
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
        // Важно: DataGridView обычно содержит placeholder "new row",
        // поэтому Rows->Count может быть > 0, даже если реальных данных нет.
        bool hasNonNewRows = false;
        for (int r = 0; r < dataGridView2->Rows->Count; r++) {
            if (!dataGridView2->Rows[r]->IsNewRow) { hasNonNewRows = true; break; }
        }
        if (!hasNonNewRows) LoadDataGridView2Defaults();
    }
    catch (Exception^) {
        LoadDataGridView2Defaults();
    }
    dataGridView2Dirty = false;
}

void ProjectServerW::DataForm::SaveDataGridView2ToFile() {
    // Параметры больше не сохраняются в ExcelSettings.txt.
    // Текущее состояние таблицы параметров попадает в лист \"Параметры\" Excel при экспорте данных.
    dataGridView2Dirty = false;
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
        // Важно: DataGridView обычно содержит placeholder "new row",
        // поэтому Rows->Count может быть > 0, даже если реальных данных нет.
        bool hasNonNewRows = false;
        for (int r = 0; r < dataGridView1->Rows->Count; r++) {
            if (!dataGridView1->Rows[r]->IsNewRow) { hasNonNewRows = true; break; }
        }
        if (!hasNonNewRows) LoadDataGridView1Defaults();
    }
    catch (Exception^) {
        LoadDataGridView1Defaults();
    }
    dataGridView1Dirty = false;
}

void ProjectServerW::DataForm::SaveDataGridView1ToFile() {
    // Параметры больше не сохраняются в ExcelSettings.txt.
    // Текущее состояние таблицы параметров попадает в лист \"Параметры\" Excel при экспорте данных.
    dataGridView1Dirty = false;
}

System::Data::DataTable^ ProjectServerW::DataForm::BuildPhaseParamsDataTableForExcel() {
    System::Data::DataTable^ t = gcnew System::Data::DataTable("PhaseParams");
    t->Columns->Add("Parameter", System::String::typeid);
    t->Columns->Add("WarmUP", System::String::typeid);
    t->Columns->Add("Plateau", System::String::typeid);
    t->Columns->Add("Finish", System::String::typeid);
    if (dataGridView1 != nullptr && !dataGridView1->IsDisposed) {
        for (int r = 0; r < dataGridView1->Rows->Count; r++) {
            System::Windows::Forms::DataGridViewRow^ row = dataGridView1->Rows[r];
            if (row->IsNewRow) continue;
            System::Data::DataRow^ dr = t->NewRow();
            dr[0] = row->Cells["Parameter"]->Value;
            dr[1] = row->Cells["WarmUP"]->Value;
            dr[2] = row->Cells["Plateau"]->Value;
            dr[3] = row->Cells["Finish"]->Value;
            t->Rows->Add(dr);
        }
    }
    return t;
}

System::Data::DataTable^ ProjectServerW::DataForm::BuildGlobalParamsDataTableForExcel() {
    System::Data::DataTable^ t = gcnew System::Data::DataTable("GlobalParams");
    t->Columns->Add("Parameter", System::String::typeid);
    t->Columns->Add("Description", System::String::typeid);
    t->Columns->Add("Value", System::String::typeid);
    if (dataGridView2 != nullptr && !dataGridView2->IsDisposed) {
        for (int r = 0; r < dataGridView2->Rows->Count; r++) {
            System::Windows::Forms::DataGridViewRow^ row = dataGridView2->Rows[r];
            if (row->IsNewRow) continue;
            System::Data::DataRow^ dr = t->NewRow();
            dr[0] = row->Cells["Parameter2"]->Value;
            dr[1] = row->Cells["Description"]->Value;
            dr[2] = row->Cells["Value"]->Value;
            t->Rows->Add(dr);
        }
    }
    return t;
}

System::Data::DataTable^ ProjectServerW::DataForm::BuildEquipmentAlarmDataTableForExcel() {
    System::Data::DataTable^ t = gcnew System::Data::DataTable("EquipmentAlarms");
    t->Columns->Add("Оборудование", System::String::typeid);
    t->Columns->Add("Описание аварии", System::String::typeid);

    if (dataGridEquipmentAlarm == nullptr || dataGridEquipmentAlarm->IsDisposed) {
        return t;
    }

    for (int r = 0; r < dataGridEquipmentAlarm->Rows->Count; r++) {
        System::Windows::Forms::DataGridViewRow^ row = dataGridEquipmentAlarm->Rows[r];
        if (row->IsNewRow) continue;

        String^ equipment = (row->Cells->Count > 0 && row->Cells[0]->Value != nullptr)
            ? System::Convert::ToString(row->Cells[0]->Value)
            : "";
        String^ description = (row->Cells->Count > 1 && row->Cells[1]->Value != nullptr)
            ? System::Convert::ToString(row->Cells[1]->Value)
            : "";

        if (String::IsNullOrWhiteSpace(equipment)) continue;
        if (equipment->Equals("Нет активных аварий")) continue;

        System::Data::DataRow^ dr = t->NewRow();
        dr[0] = equipment;
        dr[1] = description;
        t->Rows->Add(dr);
    }

    return t;
}

void ProjectServerW::DataForm::RefreshAlarmFlagsFromController() {
    try {
        if (ClientSocket == INVALID_SOCKET) {
            return;
        }
        AlarmFlagsPayload flags{};
        if (GetAlarmFlags(&flags)) {
            PopulateEquipmentAlarmGrid(flags.deviceAlarmFlags, flags.sensorAlarmFlags);
        }
    }
    catch (Exception^ ex) {
        GlobalLogger::LogMessage("Warning: RefreshAlarmFlagsFromController failed: " + ex->Message);
    }
}

// При переходе с вкладки параметров (tabPage3) не спрашиваем про сохранение.
// Раньше при входе на tabPage3 таблицы перечитывались из файла, из-за чего значения,
// загруженные с контроллера по кнопке «Чтение параметров», терялись при переключении вкладок.
// Теперь таблицы живут в памяти формы и не очищаются при переходах между вкладками.
System::Void ProjectServerW::DataForm::tabControl1_SelectedIndexChanged(System::Object^ sender, System::EventArgs^ e) {
    TabControl^ tc = safe_cast<TabControl^>(sender);

    // Авточтение параметров по вкладке отключено:
    // чтение выполняется в общем порядке один раз после установки периода опроса при соединении.

    if (tabControl1PrevTab == tabPage3 && tc->SelectedTab != tabPage3) {
        // После выхода с вкладки считаем текущее состояние базовым (без незаписанных изменений).
        dataGridView1Dirty = false;
        dataGridView2Dirty = false;
    }
    if (tc->SelectedTab == tabPage4) {
        RefreshAlarmFlagsFromController();
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

System::Void ProjectServerW::DataForm::buttonCheckAlarm_Click(System::Object^ sender, System::EventArgs^ e) {
    RefreshAlarmFlagsFromController();
}

// Преобразование имени параметра (колонка Parameter2) в идентификаторы: groupId, paramId, тип значения (1=U8, 2=U16, 3=F32).
static bool GetDefrostParamId(System::String^ paramName, uint8_t% outGroupId, uint8_t% outParamId, uint8_t% outValueType) {
    if (paramName == nullptr) return false;
    String^ p = paramName->Trim();
    if (String::IsNullOrEmpty(p)) return false;
    // Группы: 1=SENSORS, 2=TEMPERATURE, 3=HUMIDITY, 4=PWM (как в DefrostControl.h)
    if (p->StartsWith("Sensor", StringComparison::OrdinalIgnoreCase) && p->Contains("use in defrost")) {
        int i = -1; if (p->Length >= 7) Int32::TryParse(p->Substring(6, 1), i);
        if (i >= 0 && i <= 5) { outGroupId = 1; outParamId = (uint8_t)i; outValueType = DefrostParamType::U8; return true; }
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
    OpenFileDialog^ openDialog = gcnew OpenFileDialog();
    openDialog->Title = "Выберите файл данных (WorkData_...xlsx)";
    openDialog->Filter = "Файлы Excel (*.xlsx)|*.xlsx|Все файлы (*.*)|*.*";
    openDialog->DefaultExt = "xlsx";
    openDialog->RestoreDirectory = true;
    if (textBoxExcelDirectory != nullptr && !textBoxExcelDirectory->IsDisposed && textBoxExcelDirectory->Text->Length > 0) {
        openDialog->InitialDirectory = textBoxExcelDirectory->Text;
    }
    if (openDialog->ShowDialog() != System::Windows::Forms::DialogResult::OK)
        return;
    LoadParamsFromExcelFile(openDialog->FileName);
}

void ProjectServerW::DataForm::LoadParamsFromExcelFile(System::String^ filePath) {
    if (System::String::IsNullOrEmpty(filePath) || !System::IO::File::Exists(filePath)) {
        MessageBox::Show("Файл не выбран или не существует.", "Загрузка параметров");
        return;
    }
    if ((dataGridView1 == nullptr || dataGridView1->IsDisposed) && (dataGridView2 == nullptr || dataGridView2->IsDisposed)) {
        MessageBox::Show("Таблицы параметров недоступны.", "Загрузка параметров");
        return;
    }
    bool mutexAcquired = false;
    try {
        System::Threading::Mutex^ excelMutex = FormExcel::GetExcelGlobalMutex();
        mutexAcquired = excelMutex->WaitOne(15000);
        if (!mutexAcquired) {
            MessageBox::Show("Не удалось получить доступ к Excel (таймаут). Закройте другие операции с Excel и повторите.", "Загрузка параметров");
            return;
        }
        Microsoft::Office::Interop::Excel::Application^ app = gcnew Microsoft::Office::Interop::Excel::ApplicationClass();
        app->Visible = false;
        app->DisplayAlerts = false;
        System::Object^ missing = System::Type::Missing;
        Microsoft::Office::Interop::Excel::Workbook^ wb = nullptr;
        try {
            // Workbooks.Open имеет 15 параметров; в C++/CLI нужно передать все (остальные — Type::Missing)
            wb = app->Workbooks->Open(
                filePath,
                missing, missing, missing, missing, missing, missing, missing,
                missing, missing, missing, missing, missing, missing, missing);
            Microsoft::Office::Interop::Excel::Sheets^ sheets = wb->Worksheets;
            bool foundPhase = false, foundGlobal = false;
            for (int s = 1; s <= sheets->Count; s++) {
                Microsoft::Office::Interop::Excel::Worksheet^ sh = safe_cast<Microsoft::Office::Interop::Excel::Worksheet^>(sheets->Item[s]);
                System::String^ name = safe_cast<System::String^>(sh->Name);
                if (name != nullptr && name->Equals("Параметры по фазам", System::StringComparison::OrdinalIgnoreCase) &&
                    dataGridView1 != nullptr && !dataGridView1->IsDisposed) {
                    foundPhase = true;
                    dataGridView1->Rows->Clear();
                    Microsoft::Office::Interop::Excel::Range^ used = sh->UsedRange;
                    int lastRow = used->Rows->Count;
                    for (int r = 2; r <= lastRow; r++) {
                        cli::array<Object^>^ rowData = gcnew cli::array<Object^>(4);
                        for (int c = 1; c <= 4; c++) {
                            try {
                                Microsoft::Office::Interop::Excel::Range^ cell = safe_cast<Microsoft::Office::Interop::Excel::Range^>(sh->Cells[r, c]);
                                System::Object^ val = cell->Value2;
                                rowData[c - 1] = (val != nullptr) ? val->ToString() : "";
                            }
                            catch (...) { rowData[c - 1] = ""; }
                        }
                        dataGridView1->Rows->Add(rowData);
                    }
                    dataGridView1Dirty = false;
                }
                else if (name != nullptr && name->Equals("Параметры общие", System::StringComparison::OrdinalIgnoreCase) &&
                    dataGridView2 != nullptr && !dataGridView2->IsDisposed) {
                    foundGlobal = true;
                    dataGridView2->Rows->Clear();
                    Microsoft::Office::Interop::Excel::Range^ used = sh->UsedRange;
                    int lastRow = used->Rows->Count;
                    for (int r = 2; r <= lastRow; r++) {
                        cli::array<Object^>^ rowData = gcnew cli::array<Object^>(3);
                        for (int c = 1; c <= 3; c++) {
                            try {
                                Microsoft::Office::Interop::Excel::Range^ cell = safe_cast<Microsoft::Office::Interop::Excel::Range^>(sh->Cells[r, c]);
                                System::Object^ val = cell->Value2;
                                rowData[c - 1] = (val != nullptr) ? val->ToString() : "";
                            }
                            catch (...) { rowData[c - 1] = ""; }
                        }
                        dataGridView2->Rows->Add(rowData);
                    }
                    dataGridView2Dirty = false;
                }
                System::Runtime::InteropServices::Marshal::ReleaseComObject(sh);
            }
            if (!foundPhase && !foundGlobal) {
                MessageBox::Show("В выбранном файле не найдены листы «Параметры по фазам» и «Параметры общие».", "Загрузка параметров");
            }
            wb->Close(false, missing, missing);
            System::Runtime::InteropServices::Marshal::ReleaseComObject(wb);
        }
        catch (System::Exception^ ex) {
            MessageBox::Show("Ошибка при чтении файла: " + ex->Message, "Загрузка параметров");
            if (wb != nullptr) {
                try { wb->Close(false, missing, missing); } catch (...) {}
                try { System::Runtime::InteropServices::Marshal::ReleaseComObject(wb); } catch (...) {}
            }
        }
        app->Quit();
        System::Runtime::InteropServices::Marshal::ReleaseComObject(app);
    }
    catch (System::Exception^ ex) {
        MessageBox::Show("Ошибка: " + ex->Message, "Загрузка параметров");
        GlobalLogger::LogMessage("Error: LoadParamsFromExcelFile: " + ex->ToString());
    }
    finally {
        if (mutexAcquired && FormExcel::GetExcelGlobalMutex() != nullptr) {
            try { FormExcel::GetExcelGlobalMutex()->ReleaseMutex(); } catch (...) {}
        }
    }
    if (Label_Commands != nullptr && !Label_Commands->IsDisposed) {
        Label_Commands->Text = "Параметры загружены из файла";
        Label_Commands->ForeColor = System::Drawing::Color::DarkGreen;
    }
}

System::Void ProjectServerW::DataForm::buttonSaveToFile_Click(System::Object^ sender, System::EventArgs^ e) {
    SaveDataGridView1ToFile();
    SaveDataGridView2ToFile();
}

System::Void ProjectServerW::DataForm::AutoReadParametersFromController() {
    // Вспомогательный колбэк без параметров для MethodInvoker:
    // вызывает стандартный обработчик чтения параметров.
    buttonReadParameters_Click(nullptr, nullptr);
}

System::Void ProjectServerW::DataForm::buttonReadParameters_Click(System::Object^ sender, System::EventArgs^ e) {
    try {
        if ((dataGridView1 == nullptr || dataGridView1->IsDisposed) && (dataGridView2 == nullptr || dataGridView2->IsDisposed)) {
            MessageBox::Show("Таблицы параметров недоступны. Закройте и снова откройте вкладку \"Параметры\" или перезапустите приложение.", "Считать параметры");
            return;
        }
        if (ClientSocket == INVALID_SOCKET) {
            MessageBox::Show("Нет соединения с устройством. Подключитесь к устройству.");
            if (Label_Commands != nullptr && !Label_Commands->IsDisposed) {
                Label_Commands->Text = "Соединение с устройством: нет соединения";
                Label_Commands->ForeColor = System::Drawing::Color::Orange;
            }
            return;
        }
        uint8_t buffer[256];
        const uint8_t kPayloadCapacity = 255; // максимум для uint8_t, буфер 256 байт
        uint8_t len = 0;
        bool ok1 = false, ok2 = false;
        bool loaded1 = false, loaded2 = false;
        // Запросить лог по группе 5 (параметры по фазам) и загрузить в dataGridView1
        if (dataGridView1 != nullptr && !dataGridView1->IsDisposed) {
            ok1 = GetDefrostGroup(5, 0, buffer, kPayloadCapacity, &len);
            if (ok1 && len >= (uint8_t)sizeof(DefrostLogPhasePayload_t)) {
                FillDataGridView1FromGroup5Payload(buffer, len);
                loaded1 = true;
            } else if (ok1) {
                GlobalLogger::LogMessage(String::Format(
                    "Warning: Group5 read succeeded but payload too short: {0} bytes (expected >= {1})",
                    (int)len, (int)sizeof(DefrostLogPhasePayload_t)));
            }
        }
        // Запросить лог по группе 6 (общие параметры) и загрузить в dataGridView2
        if (dataGridView2 != nullptr && !dataGridView2->IsDisposed) {
            ok2 = GetDefrostGroup(6, 0, buffer, kPayloadCapacity, &len);
            if (ok2 && len >= (uint8_t)sizeof(DefrostLogGlobalPayload_t)) {
                FillDataGridView2FromGroup6Payload(buffer, len);
                loaded2 = true;
            } else if (ok2) {
                GlobalLogger::LogMessage(String::Format(
                    "Warning: Group6 read succeeded but payload too short: {0} bytes (expected >= {1})",
                    (int)len, (int)sizeof(DefrostLogGlobalPayload_t)));
            }
        }
        if (Label_Commands != nullptr && !Label_Commands->IsDisposed) {
            if (loaded1 && loaded2) {
                Label_Commands->Text = "Параметры загружены с устройства (группа 5 → таблица 1, группа 6 → таблица 2)";
                Label_Commands->ForeColor = System::Drawing::Color::DarkGreen;
            }
            else if (loaded1 || loaded2) {
                Label_Commands->Text = "Часть параметров загружена с устройства (группа 5/6). Проверьте обе таблицы.";
                Label_Commands->ForeColor = System::Drawing::Color::Orange;
            }
            else {
                Label_Commands->Text = "Не удалось получить параметры с устройства";
                Label_Commands->ForeColor = System::Drawing::Color::Red;
            }
        }
        // Mark that we have successfully loaded both tables from controller.
        // If only one group loaded, keep paramsLoadedFromDevice=false to allow another auto-read
        // on the next entry to the "Параметры" tab (if the remaining table is still empty).
        paramsLoadedFromDevice = (loaded1 && loaded2);
        GlobalLogger::LogMessage(String::Format(
            "Information: Read params from defroster: group5 transport={0}, group6 transport={1}, group5 loaded={2}, group6 loaded={3}",
            ok1 ? "OK" : "fail",
            ok2 ? "OK" : "fail",
            loaded1 ? "yes" : "no",
            loaded2 ? "yes" : "no"));
    }
    catch (Exception^ ex) {
        String^ msg = "Ошибка при считывании параметров: " + (ex->Message != nullptr ? ex->Message : "");
        MessageBox::Show(msg, "Считать параметры", MessageBoxButtons::OK, MessageBoxIcon::Warning);
        GlobalLogger::LogMessage("Error: buttonReadParameters_Click: " + msg);
        if (Label_Commands != nullptr && !Label_Commands->IsDisposed) {
            Label_Commands->Text = "Ошибка при считывании параметров";
            Label_Commands->ForeColor = System::Drawing::Color::Red;
        }
    }
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
    // Проверка: если обе таблицы параметров пусты (нет ни одной строки с данными), не отправлять команды.
    int dataRows1 = 0, dataRows2 = 0;
    if (dataGridView1 != nullptr && !dataGridView1->IsDisposed) {
        for (int r = 0; r < dataGridView1->Rows->Count; r++)
            if (!dataGridView1->Rows[r]->IsNewRow) dataRows1++;
    }
    if (dataGridView2 != nullptr && !dataGridView2->IsDisposed) {
        for (int r = 0; r < dataGridView2->Rows->Count; r++)
            if (!dataGridView2->Rows[r]->IsNewRow) dataRows2++;
    }
    if (dataRows1 == 0 && dataRows2 == 0) {
        MessageBox::Show("Таблицы параметров пусты. Заполните таблицы или загрузите параметры из файла.", "Запись параметров");
        if (Label_Commands != nullptr && !Label_Commands->IsDisposed) {
            Label_Commands->Text = "Запись в устройство: нечего записывать (таблицы пусты)";
            Label_Commands->ForeColor = System::Drawing::Color::Orange;
        }
        return;
    }
    bool ok5 = false, ok6 = false;
    if (dataRows1 > 0 && dataGridView1 != nullptr && !dataGridView1->IsDisposed) {
        DefrostLogPhasePayload_t payload5;
        if (BuildPhasePayloadFromGrid1(dataGridView1, &payload5)) {
            ok5 = SetDefrostGroup(5, (const uint8_t*)&payload5, (uint8_t)sizeof(DefrostLogPhasePayload_t));
        }
    }
    if (dataRows2 > 0 && dataGridView2 != nullptr && !dataGridView2->IsDisposed) {
        DefrostLogGlobalPayload_t payload6;
        if (BuildGlobalPayloadFromGrid2(dataGridView2, &payload6)) {
            ok6 = SetDefrostGroup(6, (const uint8_t*)&payload6, (uint8_t)sizeof(DefrostLogGlobalPayload_t));
        }
    }
    if (Label_Commands != nullptr && !Label_Commands->IsDisposed) {
        if (ok5 && ok6) {
            Label_Commands->Text = "Запись в устройство: группы 5 и 6 записаны";
            Label_Commands->ForeColor = System::Drawing::Color::DarkGreen;
        } else if (ok5 || ok6) {
            Label_Commands->Text = String::Format("Запись в устройство: группа 5 = {0}, группа 6 = {1}", ok5 ? "OK" : "ошибка", ok6 ? "OK" : "ошибка");
            Label_Commands->ForeColor = System::Drawing::Color::Orange;
        } else {
            Label_Commands->Text = "Запись в устройство: ошибка отправки групп параметров";
            Label_Commands->ForeColor = System::Drawing::Color::Red;
        }
    }
    GlobalLogger::LogMessage(String::Format("Information: Wrote defrost params by group: group5={0}, group6={1}", ok5 ? "OK" : "fail", ok6 ? "OK" : "fail"));
}

//*******************************************************************************************
// Инициализация имён битовых полей для таблицы данных
void ProjectServerW::DataForm::InitializeBitFieldNames(gcroot<cli::array<cli::array<String^>^>^>& namesRef) {
    namesRef = gcnew cli::array<cli::array<String^>^>(10);

    // Первая группа битов (входы DI, порядок как в DI_DFR_REGISTERS_t контроллера)
    namesRef[0] = gcnew cli::array<String^>(16) {
        "Vent1_Left", "Vent2_Left", "Vent1_Right", "Vent2_Right", // IN0..IN3
        "Ten1_Left", "Ten2_Left", "Ten1_Right", "Ten2_Right",     // IN4..IN7
        "Vent_Out",  // IN8: вытяжной вентилятор работает
        "Air_Open",  // IN9: заслонка открыта
        "Air_Close", // IN10: заслонка закрыта
        "Gate_Alarm",// IN11: авария ворот
        "Gate_Close",// IN12: ворота закрыты (концевик)
        "Gate_Open", // IN13: ворота открыты (концевик)
        "But_Start", // IN14: кнопка ПУСК
        "But_Stop",  // IN15: кнопка СТОП
        
    };

    // Вторая группа битов (управление устройствами)
    namesRef[1] = gcnew cli::array<String^>(17) {
        "_V0", "_V1", "_V2", "_V3", // Вентиляторы статус 1..4 включить
        "_H0", "_H1", "_H2", "_H3", // Тэны (нагрев) 1..4 включить
        "_Out",     // Вентилятор вытяжки включить
        "_Inj",     // Инжектор включить
        "_Flp",     // Заслонку открыть 
        "_Opn",     // Подъём ворот
        "_Dbl",     // Разблокировать ручное управление воротами
        "_Cls",     // Опускание ворот
        "_Wrk",     // Устройство в режиме работы автоматического цикла (зелёная лампа ПУСК)
        "_Shd",     // Процесс post-shutdown (завершение после STOP)
        "_Alr"      // Красная лампа АВАРИЯ
    };

    // Остальные группы при необходимости расширяются
}

////******************** Обработка закрытия формы ***************************************
System::Void ProjectServerW::DataForm::DataForm_FormClosing(System::Object^ sender, System::Windows::Forms::FormClosingEventArgs^ e) {
    try {
        isFormClosingNow = true;
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
                GlobalLogger::LogMessage("Information: \u042D\u043A\u0441\u043F\u043E\u0440\u0442\u0438\u0440\u0443\u0435\u043C \u0434\u0430\u043D\u043D\u044B\u0435 \u0432 Excel \u043F\u0440\u0438 \u0437\u0430\u043A\u0440\u044B\u0442\u0438\u0438 \u0444\u043E\u0440\u043C\u044B... " + excelFileName);

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
        GlobalLogger::LogMessage("Ошибка при записи параметров в устройство: " + ex->Message);
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
        GlobalLogger::LogMessage(gcnew String(L"Information: \u0424\u043E\u0440\u043C\u0430 \u0443\u0436\u0435 \u0431\u044B\u043B\u0430 \u0443\u0434\u0430\u043B\u0435\u043D\u0430 \u0438\u0437 \u043A\u0430\u0440\u0442\u044B \u0440\u0430\u043D\u0435\u0435"));
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
    if (controlLogAbsenceTimer != nullptr) {
        try {
            controlLogAbsenceTimer->Stop();
            controlLogAbsenceTimer->Tick -= gcnew EventHandler(this, &DataForm::OnControlLogAbsenceTimerTick);
        }
        catch (...) {}
        controlLogAbsenceTimer = nullptr;
    }
    if (sendStateTimer != nullptr) {
        try {
            sendStateTimer->Stop();
            sendStateTimer->Tick -= gcnew EventHandler(this, &DataForm::OnSendStateTimerTick);
        }
        catch (...) {}
        sendStateTimer = nullptr;
    }

}

// Отправка команды дефроста на устройство (реализация в CommandsDefroster.cpp)

// Единый метод установки состояния ПУСК/СТОП (кнопки и лампы).
void ProjectServerW::DataForm::SetProgramStateUi(bool isRunning)
{
    if (buttonSTART == nullptr || buttonSTOP == nullptr || labelSTART == nullptr || labelSTOP == nullptr)
        return;
    // Не дёргаем Refresh, если состояние кнопок уже соответствует (частые пакеты телеметрии вне записи в таблицу).
    if (isRunning) {
        if (!buttonSTART->Enabled && buttonSTOP->Enabled)
            return;
    }
    else {
        if (buttonSTART->Enabled && !buttonSTOP->Enabled)
            return;
    }
    if (isRunning) {
        buttonSTART->Enabled = false;
        labelSTART->BackColor = System::Drawing::Color::Lime;
        labelSTART->Text = "1";
        buttonSTOP->Enabled = true;
        labelSTOP->BackColor = System::Drawing::Color::Red;
        labelSTOP->Text = "0";
    }
    else {
        buttonSTART->Enabled = true;
        labelSTART->BackColor = System::Drawing::Color::Red;
        labelSTART->Text = "0";
        buttonSTOP->Enabled = false;
        labelSTOP->BackColor = System::Drawing::Color::Lime;
        labelSTOP->Text = "1";
    }
    if (this->IsHandleCreated && !this->IsDisposed)
        this->Refresh();
}

void ProjectServerW::DataForm::EnsureProgramRunningStateFromLog()
{
    // Лог параметров идёт только при работе в автоматическом режиме — выставляем «программа запущена».
    // Не перезаписывать только если недавно (10 с) получен ответ на СТОП — запоздалый лог не должен затирать кнопки.
    if (buttonSTART == nullptr || buttonSTOP == nullptr)
        return;
    if (buttonSTART->Enabled && !buttonSTOP->Enabled) {
        if (lastStopSuccessTime != DateTime::MinValue) {
            TimeSpan elapsed = DateTime::Now.Subtract(lastStopSuccessTime);
            if (elapsed.TotalSeconds < 10.0)
                return;
        }
    }
    SetProgramStateUi(true);
}

void ProjectServerW::DataForm::OnControlLogAbsenceTimerTick(System::Object^ sender, System::EventArgs^ e)
{
    try {
        // Авто-останов по таймауту отключён.
        // Единый критерий останова алгоритма: только подтверждённый _Wrk == 0
        // не менее 5 отсчётов подряд (см. AddDataToTable / stopConfirmedByWrk).
        if (!controllerAutoModeActive || lastControlLogTime == DateTime::MinValue) {
            controlLogAbsenceStrikeCount = 0;
            return;
        }
        TimeSpan elapsed = DateTime::Now.Subtract(lastControlLogTime);
        // Таймаут отсутствия телеметрии с _Wrk=1 = интервал измерений (с вкладки «Настройки») + 1 с
        int intervalSec = 10;
        if (numericUpDownMeasurementInterval != nullptr && !numericUpDownMeasurementInterval->IsDisposed)
            intervalSec = System::Decimal::ToInt32(numericUpDownMeasurementInterval->Value);
        double absenceThresholdSec = (double)intervalSec + 1.0;
        if (elapsed.TotalSeconds < absenceThresholdSec) {
            controlLogAbsenceStrikeCount = 0;
            return;
        }
        // Оставляем только диагностический счётчик пропусков _Wrk=1.
        controlLogAbsenceStrikeCount++;
        return;
    }
    catch (Exception^ ex) {
        GlobalLogger::LogMessage("Error: Exception in OnControlLogAbsenceTimerTick: " + ex->ToString());
    }
}

void ProjectServerW::DataForm::OnSendStateTimerTick(System::Object^ sender, System::EventArgs^ e)
{
    try {
        if (ClientSocket == INVALID_SOCKET)
            return;
        if (!startupSequenceCompleted)
            return;
        // TCP ещё открыт, но телеметрия не приходит (сброс МК, зависание, обрыв без закрытия сокета) —
        // иначе AddDataToTable не вызывается и кнопки остаются в режиме «автомат».
        if (hasTelemetry && lastTelemetryTime != DateTime::MinValue && controllerAutoModeActive) {
            int intervalSec = 10;
            if (numericUpDownMeasurementInterval != nullptr && !numericUpDownMeasurementInterval->IsDisposed) {
                intervalSec = System::Math::Max(1, System::Decimal::ToInt32(numericUpDownMeasurementInterval->Value));
            }
            double stallThresholdSec = (double)intervalSec * 2.0 + 10.0;
            if (stallThresholdSec < 20.0)
                stallThresholdSec = 20.0;
            if (DateTime::Now.Subtract(lastTelemetryTime).TotalSeconds >= stallThresholdSec) {
                ResetProgramStateAfterLinkLoss();
                return;
            }
        }
        if (!System::Threading::Monitor::TryEnter(commandPipelineGate))
            return;
        try {
        Command cmd = CreateRequestCommandSendState();
        uint8_t buffer[MAX_COMMAND_SIZE];
        size_t commandLength = BuildCommandBuffer(cmd, buffer, sizeof(buffer));
        if (commandLength == 0)
            return;
        System::Object^ recvGate = PacketQueueProcessor::GetReceivingGate(ClientSocket);
        System::Threading::Monitor::Enter(recvGate);
        try {
            System::Object^ sendGate = PacketQueueProcessor::GetSendGate(ClientSocket);
            System::Threading::Monitor::Enter(sendGate);
            try {
                int bytesSent = send(ClientSocket, reinterpret_cast<const char*>(buffer),
                    static_cast<int>(commandLength), 0);
                if (bytesSent == SOCKET_ERROR)
                    GlobalLogger::LogMessage("Warning: SEND_STATE send failed");
            }
            finally {
                System::Threading::Monitor::Exit(sendGate);
            }
        }
        finally {
            System::Threading::Monitor::Exit(recvGate);
        }
        }
        finally {
            System::Threading::Monitor::Exit(commandPipelineGate);
        }
    }
    catch (Exception^ ex) {
        GlobalLogger::LogMessage("Error: Exception in OnSendStateTimerTick: " + ex->ToString());
    }
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
            GlobalLogger::LogMessage("Error invoking TriggerExcelExport: " + ex->Message);
        }
        return;
    }
    
    try {
        StartExcelExportThread(false);
    }
    catch (Exception^ ex) {
        GlobalLogger::LogMessage("Error in TriggerExcelExport: " + ex->Message);
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
        GlobalLogger::LogMessage(String::Format(
            "Information: Автозапуск установлен на {0}",
            dateTimePickerAutoStart->Value.ToString("HH:mm")));
        
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
        GlobalLogger::LogMessage(String::Format(
            "Information: Автозапуск сработал в {0} (целевое: {1})",
            now.ToString("HH:mm:ss"),
            targetTime.ToString("HH:mm")));
        
        // Проверяем, что кнопка START доступна (нажатие по таймеру)
        if (buttonSTART->Enabled) {
            // Включение элементов
            labelAutoStart->ForeColor = System::Drawing::Color::Blue;
            Label_Commands->Text = "[Ожидание] Выполняется команда остановки...";
            Label_Commands->ForeColor = System::Drawing::Color::Blue;
            GlobalLogger::LogMessage(Label_Commands->Text);
            
            CommandResponse startResp{};
            Command startCmd = CreateControlCommand(CmdProgControl::START);
            bool startOk = SendCommandAndWaitResponse(startCmd, startResp, "START");

            if (!startOk && startResp.status == CmdStatus::TIMEOUT) {
                GlobalLogger::LogMessage("Warning: Автозапуск: START не подтверждён (нет ответа от устройства)");
                Label_Commands->Text = "[Ожидание] START не подтверждён...";
                Label_Commands->ForeColor = System::Drawing::Color::Orange;
                GlobalLogger::LogMessage(Label_Commands->Text);
            }
            else if (!startOk) {
                GlobalLogger::LogMessage("Warning: Автозапуск: START не выполнен (ошибка связи/таймаут)");
                Label_Commands->Text = "[!] Ошибка: START не отправлен";
                Label_Commands->ForeColor = System::Drawing::Color::Orange;
                GlobalLogger::LogMessage(Label_Commands->Text);
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
            GlobalLogger::LogMessage(Label_Commands->Text);
            
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
        GlobalLogger::LogMessage("Error in RestoreAutoStartColor: " + ex->Message);
    }
}

//  ===========================================================================
// Отправка команды автоперезапуска по времени
//  ===========================================================================

System::Void ProjectServerW::DataForm::checkBoxAutoRestart_CheckedChanged(System::Object^ sender, System::EventArgs^ e) {
    if (checkBoxAutoRestart->Checked) {
        timerAutoRestart->Start();
        GlobalLogger::LogMessage(String::Format(
            "Information: Автоперезапуск установлен на {0}",
            dateTimePickerAutoRestart->Value.ToString("HH:mm")));

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
        GlobalLogger::LogMessage(gcnew String(L"Information: \u0410\u0432\u0442\u043E\u043F\u0435\u0440\u0435\u0437\u0430\u043F\u0443\u0441\u043A \u043E\u0442\u043A\u043B\u044E\u0447\u0451\u043D"));

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

    GlobalLogger::LogMessage(String::Format(
        "Information: Автоперезапуск сработал в {0} (целевое: {1})",
        now.ToString("HH:mm:ss"),
        targetTime.ToString("HH:mm")));

    // Если ни STOP ни START недоступны (устройство отключено), то "автоперезапуск" отменяется.
    if (!buttonSTOP->Enabled && !buttonSTART->Enabled) {
        GlobalLogger::LogMessage(gcnew String(L"Warning: \u0410\u0432\u0442\u043E\u043F\u0435\u0440\u0435\u0437\u0430\u043F\u0443\u0441\u043A: \u0443\u0441\u0442\u0440\u043E\u0439\u0441\u0442\u0432\u043E \u043E\u0442\u043A\u043B\u044E\u0447\u0435\u043D\u043E (\u043D\u0435\u0442 \u0441\u0432\u044F\u0437\u0438), \u043E\u0442\u043C\u0435\u043D\u0430..."));
        return;
    }

    labelAutoRestart->ForeColor = System::Drawing::Color::Blue;

    if (buttonSTOP->Enabled) {
        Label_Commands->Text = "[Выполняется] Отправлена команда STOP...";
        Label_Commands->ForeColor = System::Drawing::Color::Blue;
        GlobalLogger::LogMessage(Label_Commands->Text);

        CommandResponse stopResp{};
        Command stopCmd = CreateControlCommand(CmdProgControl::STOP);
        bool stopOk = SendCommandAndWaitResponse(stopCmd, stopResp, "STOP");

        if (stopOk) {
            autoRestartPending = true;
            autoRestartStopIssuedTime = DateTime::Now;
            Label_Commands->Text = "[Выполняется] STOP отправлена, ожидание экспорта данных...";
            Label_Commands->ForeColor = System::Drawing::Color::Blue;
            GlobalLogger::LogMessage(Label_Commands->Text);
        }
        else if (stopResp.status == CmdStatus::TIMEOUT) {
            autoRestartPending = true;
            autoRestartStopIssuedTime = DateTime::Now;
            GlobalLogger::LogMessage(gcnew String(L"Warning: \u0410\u0432\u0442\u043E\u043F\u0435\u0440\u0435\u0437\u0430\u043F\u0443\u0441\u043A: STOP \u043D\u0435 \u043F\u043E\u0434\u0442\u0432\u0435\u0440\u0436\u0434\u0451\u043D (\u043D\u0435\u0442 \u043E\u0442\u0432\u0435\u0442\u0430)"));
            Label_Commands->Text = "[Выполняется] STOP не подтверждён, ожидание...";
            Label_Commands->ForeColor = System::Drawing::Color::Orange;
            GlobalLogger::LogMessage(Label_Commands->Text);
        }
        else {
            autoRestartPending = false;
            autoRestartStopIssuedTime = DateTime::MinValue;
            GlobalLogger::LogMessage(gcnew String(L"Warning: \u0410\u0432\u0442\u043E\u043F\u0435\u0440\u0435\u0437\u0430\u043F\u0443\u0441\u043A: STOP \u043D\u0435 \u0432\u044B\u043F\u043E\u043B\u043D\u0435\u043D (\u043E\u0448\u0438\u0431\u043A\u0430 \u0441\u0432\u044F\u0437\u0438/\u0442\u0430\u0439\u043C\u0430\u0443\u0442)"));
            Label_Commands->Text = "[!] Ошибка выполнения: STOP не отправлен";
            Label_Commands->ForeColor = System::Drawing::Color::Orange;
            GlobalLogger::LogMessage(Label_Commands->Text);
        }

            autoRestartInternalUncheck = true;
        checkBoxAutoRestart->Checked = false; // сбрасываем, UX чекбокс не связан с AutoStart
    }
    else if (buttonSTART->Enabled) {
        // Устройство уже остановлено в этот момент — отправляем START.
        Label_Commands->Text = "[Выполняется] Ожидание отключения, отправка START...";
        Label_Commands->ForeColor = System::Drawing::Color::Blue;
        GlobalLogger::LogMessage(Label_Commands->Text);

        CommandResponse startResp{};
        Command startCmd = CreateControlCommand(CmdProgControl::START);
        bool startOk = SendCommandAndWaitResponse(startCmd, startResp, "START");

        if (startOk) {
            Label_Commands->Text = "[Выполняется] START отправлен";
            Label_Commands->ForeColor = System::Drawing::Color::Blue;
            GlobalLogger::LogMessage(Label_Commands->Text);
        }
        else if (startResp.status == CmdStatus::TIMEOUT) {
            GlobalLogger::LogMessage(gcnew String(L"Warning: \u0410\u0432\u0442\u043E\u043F\u0435\u0440\u0435\u0437\u0430\u043F\u0443\u0441\u043A: START \u043D\u0435 \u043F\u043E\u0434\u0442\u0432\u0435\u0440\u0436\u0434\u0451\u043D (\u043D\u0435\u0442 \u043E\u0442\u0432\u0435\u0442\u0430)"));
            Label_Commands->Text = "[Выполняется] START не подтверждён...";
            Label_Commands->ForeColor = System::Drawing::Color::Orange;
            GlobalLogger::LogMessage(Label_Commands->Text);
        }
        else {
            GlobalLogger::LogMessage(gcnew String(L"Warning: \u0410\u0432\u0442\u043E\u043F\u0435\u0440\u0435\u0437\u0430\u043F\u0443\u0441\u043A: START \u043D\u0435 \u0432\u044B\u043F\u043E\u043B\u043D\u0435\u043D (\u043E\u0448\u0438\u0431\u043A\u0430 \u0441\u0432\u044F\u0437\u0438/\u0442\u0430\u0439\u043C\u0430\u0443\u0442)"));
            Label_Commands->Text = "[!] Ошибка выполнения: START не отправлен";
            Label_Commands->ForeColor = System::Drawing::Color::Orange;
            GlobalLogger::LogMessage(Label_Commands->Text);
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
        GlobalLogger::LogMessage("Error in RestoreAutoRestartColor: " + ex->Message);
    }
}

void ProjectServerW::DataForm::ExecuteAutoRestartStart() {
    try {
        if (buttonSTART == nullptr || buttonSTART->IsDisposed) {
            return;
        }
        if (!buttonSTART->Enabled) {
            GlobalLogger::LogMessage(gcnew String(L"Warning: \u0410\u0432\u0442\u043E\u043F\u0435\u0440\u0435\u0437\u0430\u043F\u0443\u0441\u043A: START \u043D\u0435 \u0432\u044B\u043F\u043E\u043B\u043D\u0435\u043D (\u043A\u043D\u043E\u043F\u043A\u0430 \u043D\u0435\u0434\u043E\u0441\u0442\u0443\u043F\u043D\u0430)"));
            return;
        }

        Label_Commands->Text = "[Выполняется] Выполняется команда остановки...";
        Label_Commands->ForeColor = System::Drawing::Color::Blue;
        GlobalLogger::LogMessage(Label_Commands->Text);
        SendStartCommand();
    }
    catch (Exception^ ex) {
        GlobalLogger::LogMessage("Error: Exception in ExecuteAutoRestartStart: " + ex->ToString());
    }
}

// Отправка команды старта дефроста (реализация в CommandsDefroster.cpp)


