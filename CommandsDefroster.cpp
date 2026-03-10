#include "DataForm.h"
#include "Commands.h"
#include "PacketQueueProcessor.h"

using namespace System;
using namespace System::Windows::Forms;
using namespace ProjectServerW;

/*
********************** методы для работы с командами управления дефростером **********************
*/

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
            GlobalLogger::LogMessage("Error: Не могу отправить команду " + commandName +
                ", нет активного соединения с клиентом!");
            return false;
        }

        // Формируем буфер команды
        uint8_t buffer[MAX_COMMAND_SIZE];
        size_t commandLength = BuildCommandBuffer(cmd, buffer, sizeof(buffer));

        if (commandLength == 0) {
            String^ errorMsg = "Ошибка формирования команды " + commandName;
            MessageBox::Show(errorMsg);
            GlobalLogger::LogMessage("Error: " + errorMsg);
            return false;
        }

        int bytesSent = SOCKET_ERROR;
        // Полудуплекс: не отправлять команду, пока recv()-поток обрабатывает принятые от контроллера данные.
        // Мы получаем recvGate только после того, как recv() вернул данные (телеметрию). Сразу после этого
        // контроллер может быть ещё занят передачей или не готов принять команду — даём время на "разворот" линии.
        System::Object^ recvGate = PacketQueueProcessor::GetReceivingGate(clientSocket);
        System::Threading::Monitor::Enter(recvGate);
        try {
            System::Threading::Thread::Sleep(150);  // задержка разворота полудуплекса (мс)
            // Сериализуем send() на сокет, чтобы байтовые потоки не перемешивались с ACK телеметрии.
            System::Object^ sendGate = PacketQueueProcessor::GetSendGate(clientSocket);
            System::Threading::Monitor::Enter(sendGate);
            try {
                bytesSent = send(clientSocket, reinterpret_cast<const char*>(buffer),
                    static_cast<int>(commandLength), 0);
            }
            finally {
                System::Threading::Monitor::Exit(sendGate);
            }
        }
        finally {
            System::Threading::Monitor::Exit(recvGate);
        }

        if (bytesSent == SOCKET_ERROR) {
            int error = WSAGetLastError();
            String^ errorMsg = "Ошибка отправки команды " + commandName + ": " + error.ToString();
            MessageBox::Show(errorMsg);
            GlobalLogger::LogMessage("Error: " + errorMsg);
            return false;
        }
        else if (bytesSent == commandLength) {
            // Команда успешно отправлена
            Label_Commands->Text = "Команда " + commandName + " отправлена клиенту";
            GlobalLogger::LogMessage("Information: Команда " + commandName + " отправлена клиенту");
            return true;
        }
        else {
            // Отправлено меньше байт, чем ожидалось
            String^ errorMsg = "Отправлено только " + bytesSent.ToString() + " из " +
                commandLength.ToString() + " байт для команды " + commandName;
            MessageBox::Show(errorMsg);
            GlobalLogger::LogMessage("Error: Частичная отправка команды " +
                commandName + ": " + errorMsg);
            return false;
        }
    }
    catch (Exception^ ex) {
        String^ errorMsg = "Исключение при отправке команды " + commandName + ": " + ex->Message;
        MessageBox::Show(errorMsg);
        GlobalLogger::LogMessage("Error: " + errorMsg);
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
        // Команда успешно выполнена на контроллере; обновление кнопок обязательно в UI-потоке и до выхода
        if (this->InvokeRequired) {
            this->Invoke(gcnew System::Windows::Forms::MethodInvoker(this, &DataForm::buttonSTOPstate_TRUE));
        }
        else {
            buttonSTOPstate_TRUE();
        }
        Label_Commands->Text = "[OK] Программа запущена";
        Label_Commands->ForeColor = System::Drawing::Color::Green;
        GlobalLogger::LogMessage("Information: Команда START успешно выполнена контроллером");

        // Восстанавливаем цвет через 3 секунды с помощью таймера
        System::Windows::Forms::Timer^ colorTimer = gcnew System::Windows::Forms::Timer();
        colorTimer->Interval = 3000;
        colorTimer->Tick += gcnew EventHandler(this, &DataForm::RestoreLabelCommandsColor);
        colorTimer->Start();
    }
    else {
        // Ошибка выполнения команды - детали уже обработаны в ProcessResponse
        GlobalLogger::LogMessage(String::Format(
            "Error: Команда START не выполнена. Статус: 0x{0:X2} ({1})",
            response.status, gcnew String(GetStatusName(response.status))));

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
        // Команда успешно выполнена на контроллере; обновление кнопок обязательно в UI-потоке и до выхода
        if (this->InvokeRequired) {
            this->Invoke(gcnew System::Windows::Forms::MethodInvoker(this, &DataForm::buttonSTARTstate_TRUE));
        }
        else {
            buttonSTARTstate_TRUE();
        }
        lastStopSuccessTime = DateTime::Now; // чтобы запоздалый лог не перезаписал кнопки в течение 10 с
        Label_Commands->Text = "[OK] Программа остановлена";
        Label_Commands->ForeColor = System::Drawing::Color::Green;
        GlobalLogger::LogMessage("Information: Команда STOP успешно выполнена контроллером");

        // По успешному СТОП — записать таблицу данных в Excel
        StartExcelExportThread(false);

        // Восстанавливаем цвет через 3 секунды с помощью таймера
        System::Windows::Forms::Timer^ colorTimer = gcnew System::Windows::Forms::Timer();
        colorTimer->Interval = 3000;
        colorTimer->Tick += gcnew EventHandler(this, &DataForm::RestoreLabelCommandsColor);
        colorTimer->Start();
    }
    else {
        // Ошибка выполнения команды - детали уже обработаны в ProcessResponse
        GlobalLogger::LogMessage(String::Format(
            "Error: Команда STOP не выполнена. Статус: 0x{0:X2} ({1})",
            response.status, gcnew String(GetStatusName(response.status))));

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

        // Важно: после RESET контроллер перезагружается и некоторое время не читает UART4.
        // Поэтому планируем отложенную отправку GET_VERSION и SET_INTERVAL с повторами.
        SchedulePostResetInit();

        // Восстанавливаем цвет через 3 секунды с помощью таймера
        System::Windows::Forms::Timer^ colorTimer = gcnew System::Windows::Forms::Timer();
        colorTimer->Interval = 3000;
        colorTimer->Tick += gcnew EventHandler(this, &DataForm::RestoreLabelCommandsColor);
        colorTimer->Start();
    }
    else {
        // Ошибка выполнения команды - детали уже обработаны в ProcessResponse
        GlobalLogger::LogMessage(String::Format(
            "Error: Команда RESET не выполнена. Статус: 0x{0:X2} ({1})",
            response.status, gcnew String(GetStatusName(response.status))));

        // Дополнительная обработка специфичных ошибок для RESET
        switch (response.status) {
        case CmdStatus::EXECUTION_ERROR:
            // Контроллер не может выполнить сброс
            Label_Commands->Text = "[!] Невозможно выполнить сброс контроллера";
            Label_Commands->ForeColor = System::Drawing::Color::Orange;
            break;

        case CmdStatus::TIMEOUT:
            // После RESET контроллер перезагружается и не успевает ответить — ожидаем переподключение и запускаем отложенную инициализацию.
            if (Label_Commands != nullptr && !Label_Commands->IsDisposed) {
                Label_Commands->Text = "[i] Сброс отправлен, ожидание переподключения...";
                Label_Commands->ForeColor = System::Drawing::Color::Blue;
            }
            SchedulePostResetInit();
            break;

        default:
            // Другие ошибки уже отображены в ProcessResponse
            break;
        }
    }
}

// Применить запоздалый ответ GET_VERSION или SET_INTERVAL (при получении «чужого» ответа в цикле ожидания)
bool ProjectServerW::DataForm::ApplyLateStartupResponse(const CommandResponse& candidate) {
    if (candidate.status != CmdStatus::OK) return false;
    if (candidate.commandType == CmdType::REQUEST && candidate.commandCode == CmdRequest::GET_VERSION && candidate.dataLength > 0) {
        String^ version = gcnew String(reinterpret_cast<const char*>(candidate.data), 0, static_cast<int>(candidate.dataLength), System::Text::Encoding::ASCII);
        pendingVersion = version;
        if (label_Version != nullptr && !label_Version->IsDisposed) {
            if (label_Version->InvokeRequired) {
                label_Version->BeginInvoke(gcnew System::Windows::Forms::MethodInvoker(this, &DataForm::UpdateVersionLabelInternal));
            }
            else {
                label_Version->Text = version;
            }
        }
        GlobalLogger::LogMessage("Information: Версия прошивки получена (запоздалый ответ): " + version);
        return true;
    }
    if (candidate.commandType == CmdType::CONFIGURATION && candidate.commandCode == CmdConfig::SET_INTERVAL) {
        GlobalLogger::LogMessage("Information: Интервал измерений принят контроллером (запоздалый ответ)");
        return true;
    }
    return false;
}

// Метод для запроса версии прошивки контроллера
bool ProjectServerW::DataForm::SendVersionRequest() {
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
                }
                else {
                    label_Version->Text = version;
                }
            }

            Label_Commands->Text = "Версия прошивки получена: " + version;
            Label_Commands->ForeColor = System::Drawing::Color::Green;
            GlobalLogger::LogMessage("Information: Версия прошивки контроллера: " + version);
            return true;
        }
        else {
            Label_Commands->Text = "Версия получена, но данные пусты";
            Label_Commands->ForeColor = System::Drawing::Color::Orange;
            return false;
        }

        // Восстанавливаем цвет через 3 секунды в обратном таймере
        System::Windows::Forms::Timer^ colorTimer = gcnew System::Windows::Forms::Timer();
        colorTimer->Interval = 3000;
        colorTimer->Tick += gcnew EventHandler(this, &DataForm::RestoreLabelCommandsColor);
        colorTimer->Start();
    }
    else {
        // Ошибка выполнения команды
        GlobalLogger::LogMessage(String::Format(
            "Error: Команда GET_VERSION не выполнена. Статус: 0x{0:X2} ({1})",
            response.status, gcnew String(GetStatusName(response.status))));

        Label_Commands->Text = "[!] Не удалось получить версию прошивки";
        Label_Commands->ForeColor = System::Drawing::Color::Red;
        return false;
    }
}

bool ProjectServerW::DataForm::SendSetIntervalCommand(int intervalSeconds) {
    // Why: persist the chosen interval even when the controller is temporarily offline.
    SaveSettings();

    if (intervalSeconds <= 0) {
        Label_Commands->Text = "[!] Интервал должен быть больше 0";
        Label_Commands->ForeColor = System::Drawing::Color::Orange;
        return false;
    }

    // Why: firmware expects uint16 payload for SET_INTERVAL (2 bytes), not int32.
    if (intervalSeconds > 65535) {
        intervalSeconds = 65535;
    }
    const uint16_t intervalU16 = static_cast<uint16_t>(intervalSeconds);
    Command cmd = CreateConfigCommandU16(CmdConfig::SET_INTERVAL, intervalU16);
    CommandResponse response;

    if (SendCommandAndWaitResponse(cmd, response, "SET_INTERVAL")) {
        Label_Commands->Text = String::Format("[OK] Интервал измерений: {0} с", intervalSeconds);
        Label_Commands->ForeColor = System::Drawing::Color::Green;
        GlobalLogger::LogMessage(Label_Commands->Text);

        System::Windows::Forms::Timer^ colorTimer = gcnew System::Windows::Forms::Timer();
        colorTimer->Interval = 3000;
        colorTimer->Tick += gcnew EventHandler(this, &DataForm::RestoreLabelCommandsColor);
        colorTimer->Start();
        return true;
    }
    return false;
}

// Defrost params API (compatible with DefrostControl + CommandReceiver on controller)
bool ProjectServerW::DataForm::SetDefrostParam(uint8_t groupId, uint8_t paramId, const DefrostParamValue& value) {
    if (clientSocket == INVALID_SOCKET) return false;
    if (value.valueType != DefrostParamType::U8 && value.valueType != DefrostParamType::U16 && value.valueType != DefrostParamType::F32) return false;
    Command cmd = CreateConfigCommandDefrostSetParam(groupId, paramId, value);
    if (cmd.dataLength == 0) return false;
    CommandResponse response;
    return SendCommandAndWaitResponse(cmd, response, "SET_DEFROST_PARAM") && response.status == CmdStatus::OK;
}

bool ProjectServerW::DataForm::GetDefrostParam(uint8_t groupId, uint8_t paramId, DefrostParamValue* outValue) {
    if (clientSocket == INVALID_SOCKET || outValue == nullptr) return false;
    Command cmd = CreateRequestCommandDefrostGetParam(groupId, paramId);
    CommandResponse response;
    if (!SendCommandAndWaitResponse(cmd, response, "GET_DEFROST_PARAM") || response.status != CmdStatus::OK) return false;
    return ParseDefrostParamResponse(response, nullptr, nullptr, outValue);
}

bool ProjectServerW::DataForm::GetDefrostGroup(uint8_t groupId, uint8_t page, uint8_t* outData, uint8_t outCapacity, uint8_t* outLength) {
    if (clientSocket == INVALID_SOCKET || outData == nullptr || outLength == nullptr) return false;
    Command cmd = CreateRequestCommandDefrostGetGroup(groupId, page);
    CommandResponse response;
    if (!SendCommandAndWaitResponse(cmd, response, "GET_DEFROST_GROUP") || response.status != CmdStatus::OK) return false;
    if (response.dataLength < 2) return false;
    uint8_t payloadLen = (uint8_t)(response.dataLength - 2);
    if (payloadLen > outCapacity) return false;
    memcpy(outData, &response.data[2], payloadLen);
    *outLength = payloadLen;
    return true;
}

bool ProjectServerW::DataForm::SetDefrostGroup(uint8_t groupId, const uint8_t* payload, uint8_t payloadLen) {
    if (clientSocket == INVALID_SOCKET || payload == nullptr) return false;
    Command cmd = CreateConfigCommandSetDefrostGroup(groupId, payload, payloadLen);
    if (cmd.dataLength == 0) return false;
    CommandResponse response;
    return SendCommandAndWaitResponse(cmd, response, "SET_DEFROST_GROUP") && response.status == CmdStatus::OK;
}

void ProjectServerW::DataForm::SendCommandInfoRequest() {
    // Почему: это аудит обработки команд, а не "состояние устройства" (телеметрия).
    // Ожидается, что прошивка запоминает последнюю принятую команду и отдаёт её по этому запросу.
    Command cmd = CreateRequestCommand(CmdRequest::GET_CMD_INFO);
    CommandResponse response;

    const DateTime requestTime = DateTime::Now;
    bool received = false;

    // Делаем пару попыток: прошивка может быть занята сразу после переподключения/старта.
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
    GlobalLogger::LogMessage(String::Format(
        "Warning: GET_CMD_INFO: no response for {0:F1}s; device may be OFF and the connection can be breaking",
        waited.TotalSeconds));
    if (Label_Commands != nullptr && !Label_Commands->IsDisposed) {
        Label_Commands->Text = "[!] GET_CMD_INFO: нет ответа (возможно устройство выключено)";
        Label_Commands->ForeColor = System::Drawing::Color::Orange;
    }
}

void ProjectServerW::DataForm::ScheduleCommandInfoProbe(System::String^ reason) {
    // Почему: если ответ отброшен или истёк таймаут — спрашиваем устройство, что оно приняло/обработало последним.
    // Защищаемся от рекурсии и засорения лога.
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

        GlobalLogger::LogMessage("Information: Scheduling GET_CMD_INFO probe. Reason: " + (reason != nullptr ? reason : ""));

        if (this->InvokeRequired) {
            this->BeginInvoke(gcnew MethodInvoker(this, &DataForm::ExecuteCommandInfoProbe));
        }
        else {
            ExecuteCommandInfoProbe();
        }
    }
    catch (Exception^ ex) {
        GlobalLogger::LogMessage("Error: Exception in ScheduleCommandInfoProbe: " + ex->Message);
    }
}

void ProjectServerW::DataForm::ExecuteCommandInfoProbe() {
    // Почему: выполняем probe в UI-потоке; внутри синхронное ожидание ответа на команду.
    if (cmdInfoProbeInProgress) {
        return;
    }

    cmdInfoProbeInProgress = true;
    try {
        SendCommandInfoRequest();
    }
    catch (Exception^ ex) {
        GlobalLogger::LogMessage("Error: Exception in ExecuteCommandInfoProbe: " + ex->ToString());
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
        const wchar_t* statusDescription = GetStatusDescription(response.status);
        Command responseCmd{};
        responseCmd.commandType = response.commandType;
        responseCmd.commandCode = response.commandCode;
        const char* responseCmdName = GetCommandName(responseCmd);
        const char* responseCmdTypeName = GetCommandTypeName(response.commandType);

        // Формируем сообщение о результате выполнения команды
        String^ message;

        if (response.status == CmdStatus::OK) {
            message = String::Format(
                "[OK] Команда Type=0x{0:X2} ({1}), Code=0x{2:X2} ({3}) успешно выполнена"
                "\nСтатус: 0x{4:X2} ({5}) - {6}",
                response.commandType, gcnew String(responseCmdTypeName),
                response.commandCode, gcnew String(responseCmdName),
                response.status, gcnew String(statusName), gcnew String(statusDescription));

            // Если есть данные в ответе (для команд REQUEST)
            if (response.dataLength > 0) {
                // Обрабатываем данные в зависимости от типа команды
                if (response.commandType == CmdType::REQUEST) {
                    switch (response.commandCode) {
                    case CmdRequest::GET_CMD_INFO: {
                        // Ожидаемый payload (контракт прошивки):
                        // [0] lastCmdType (uint8)
                        // [1] lastCmdCode (uint8)
                        // [2..3] lastCmdDeviceTimeSeconds (uint16, те же единицы, что и telemetry MSGQUEUE_OBJ_t.Time)
                        // [4] ackSent (uint8, 0/1)
                        // [5] lastCmdStatus (uint8, CmdStatus)
                        if (response.dataLength >= 6) {
                            const uint8_t lastCmdType = response.data[0];
                            const uint8_t lastCmdCode = response.data[1];
                            uint16_t lastCmdSeconds = 0;
                            memcpy(&lastCmdSeconds, &response.data[2], 2);
                            const uint8_t ackSent = response.data[4];
                            const uint8_t lastCmdStatus = response.data[5];
                            Command lastCmd{};
                            lastCmd.commandType = lastCmdType;
                            lastCmd.commandCode = lastCmdCode;
                            const char* lastCmdName = GetCommandName(lastCmd);
                            const char* lastCmdTypeName = GetCommandTypeName(lastCmdType);
                            const char* lastStatusName = GetStatusName(lastCmdStatus);
                            const wchar_t* lastStatusDescription = GetStatusDescription(lastCmdStatus);

                            // Почему: время устройства относительное; привязываем к реальному времени по последней метке телеметрии.
                            // Это оценка; предполагается, что телеметрия и обработчик команд используют одну временную базу.
                            DateTime approxWallTime = DateTime::MinValue;
                            if (hasTelemetry && lastTelemetryTime != DateTime::MinValue) {
                                uint16_t cur = lastTelemetryDeviceSeconds;
                                uint16_t prev = lastCmdSeconds;
                                int delta = static_cast<int>(cur) - static_cast<int>(prev);
                                if (delta < 0) {
                                    // Обрабатываем переполнение uint16 (счётчик секунд устройства).
                                    delta += 65536;
                                }
                                approxWallTime = lastTelemetryTime.Subtract(TimeSpan::FromSeconds(delta));
                            }

                            message += String::Format(
                                "\nПоследняя команда (устройство): Type=0x{0:X2} ({1}), Code=0x{2:X2} ({3})"
                                "\nВремя устройства: {4} сек"
                                "\nACK отправлен: {5}"
                                "\nСтатус обработки: 0x{6:X2} ({7}) - {8}",
                                lastCmdType, gcnew String(lastCmdTypeName),
                                lastCmdCode, gcnew String(lastCmdName),
                                lastCmdSeconds,
                                (ackSent != 0 ? "ДА" : "НЕТ"),
                                lastCmdStatus, gcnew String(lastStatusName), gcnew String(lastStatusDescription));

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
                    }
                }
            }

            // Отображаем успешный результат
            Label_Commands->Text = message;
            GlobalLogger::LogMessage(message);
        }
        else {
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
            GlobalLogger::LogMessage(message);

            // Восстанавливаем цвет текста через 3 секунды с помощью таймера
            System::Windows::Forms::Timer^ colorTimer = gcnew System::Windows::Forms::Timer();
            colorTimer->Interval = 3000;
            colorTimer->Tick += gcnew EventHandler(this, &DataForm::RestoreLabelCommandsColor);
            colorTimer->Start();
        }
    }
    catch (Exception^ ex) {
        String^ errorMsg = "Исключение в ProcessResponse: " + ex->Message;
        MessageBox::Show(errorMsg, "Критическая ошибка",
            MessageBoxButtons::OK, MessageBoxIcon::Error);
        GlobalLogger::LogMessage(errorMsg);
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
        GlobalLogger::LogMessage("Error in RestoreLabelCommandsColor: " + ex->Message);
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
        }
        else {
            sendResult = SendCommand(cmd);
        }

        if (!sendResult) {
            GlobalLogger::LogMessage("Error: Failed to send command");
            return false;
        }

        // Ждём нужный ответ и допускаем "чужие" ответы в общей очереди.
        // Почему: recv()-цикл может поставить в очередь ответы от более ранних команд; отказ по первому несовпадению
        // ломает авто-сценарии и может ошибочно пометить валидную команду как неуспешную.
        const bool isCmdInfoRequest = (cmd.commandType == CmdType::REQUEST && cmd.commandCode == CmdRequest::GET_CMD_INFO);
        const bool isGetVersion = (cmd.commandType == CmdType::REQUEST && cmd.commandCode == CmdRequest::GET_VERSION);

        // Первая команда после подключения (GET_VERSION) и зонд GET_CMD_INFO — увеличенный таймаут:
        const int defaultTimeoutMs = 2000;
        const int totalTimeoutMs = (isGetVersion || isCmdInfoRequest) ? 6000 : defaultTimeoutMs;
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
                    // После RESET контроллер перезагружается и не отвечает; зонд GET_CMD_INFO бесполезен и долго ждёт на мёртвом сокете.
                    const bool isResetCmd = (cmd.commandType == CmdType::PROG_CONTROL && cmd.commandCode == CmdProgControl::RESET);
                    if (isResetCmd) {
                        SchedulePostResetInit();
                    }
                    else {
                        ScheduleCommandInfoProbe(String::Format(
                            "timeout waiting response Type=0x{0:X2}, Code=0x{1:X2}",
                            cmd.commandType, cmd.commandCode));
                    }
                }
                return false;
            }

            CommandResponse candidate{};
            cli::array<System::Byte>^ rawCandidate = nullptr;
            if (!ReceiveResponse(candidate, static_cast<int>(remaining.TotalMilliseconds), rawCandidate)) {
                continue;
            }

            if (candidate.commandType != cmd.commandType || candidate.commandCode != cmd.commandCode) {
                if (ApplyLateStartupResponse(candidate)) {
                    // Запоздалый ответ GET_VERSION или SET_INTERVAL применён, продолжаем ждать нужный ответ
                }
                else {
                    String^ hex = "";
                    if (rawCandidate != nullptr && rawCandidate->Length > 0) {
                        System::Text::StringBuilder^ sb = gcnew System::Text::StringBuilder(rawCandidate->Length * 3);
                        for (int i = 0; i < rawCandidate->Length; i++) {
                            if (i != 0) sb->Append(" ");
                            sb->Append(rawCandidate[i].ToString("X2"));
                        }
                        hex = sb->ToString();
                    }
                    GlobalLogger::LogMessage(String::Format(
                        "Warning: Discarding unrelated response. Expected Type=0x{0:X2}, Code=0x{1:X2}; Got Type=0x{2:X2}, Code=0x{3:X2}; Raw={4}",
                        cmd.commandType, cmd.commandCode, candidate.commandType, candidate.commandCode, hex));
                    if (!isCmdInfoRequest) {
                        const bool isResetCmd = (cmd.commandType == CmdType::PROG_CONTROL && cmd.commandCode == CmdProgControl::RESET);
                        if (!isResetCmd) {
                            ScheduleCommandInfoProbe(String::Format(
                                "discarded unrelated response while waiting Type=0x{0:X2}, Code=0x{1:X2}",
                                cmd.commandType, cmd.commandCode));
                        }
                    }
                }
                continue;
            }

            response = candidate;
            break;
        }

        // Обрабатываем полученный ответ
        ProcessResponse(response);

        return (response.status == CmdStatus::OK);

    }
    catch (Exception^ ex) {
        String^ errorMsg = "Exception in SendCommandAndWaitResponse: " + ex->Message;
        MessageBox::Show(errorMsg);
        GlobalLogger::LogMessage(errorMsg);
        return false;
    }
}

// Перегрузка SendCommandAndWaitResponse без указания имени команды
bool ProjectServerW::DataForm::SendCommandAndWaitResponse(
    const Command& cmd, CommandResponse& response) {
    return SendCommandAndWaitResponse(cmd, response, nullptr);
}

bool ProjectServerW::DataForm::TrySendControlCommandFireAndForget(uint8_t controlCode, System::String^ commandName) {
    // Почему: плановые операции не должны зависеть от синхронного ответа, потому что некоторые прошивки
    // могут выполнить START/STOP, но пропустить/потерять ACK. Реальное состояние подтверждаем по Work в телеметрии.
    try {
        Command cmd = CreateControlCommand(controlCode);
        return SendCommand(cmd, commandName);
    }
    catch (Exception^ ex) {
        GlobalLogger::LogMessage("Error: Exception in TrySendControlCommandFireAndForget: " + ex->Message);
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
    // Почему: сохраняем поведение протокола (ждём ACK и реагируем), но остаёмся устойчивыми при потере ACK.
    // Проверка состояния в fallback-варианте выполняется по телеметрии Work на более высоком уровне логики.
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

        // Часть ответов может относиться к другим командам; здесь повторяем поведение сопоставления как в SendCommandAndWaitResponse.
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
            GlobalLogger::LogMessage(String::Format(
                "Warning: Discarding unrelated response in SendControlCommandWithAck. Expected Type=0x{0:X2}, Code=0x{1:X2}; Got Type=0x{2:X2}, Code=0x{3:X2}; Raw={4}",
                cmd.commandType, cmd.commandCode, response.commandType, response.commandCode, hex));
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
        GlobalLogger::LogMessage("Error: Exception in button_CMDINFO_Click: " + ex->ToString());
    }
}


