#include "Commands.h"

// Таблица CRC16-CITT (та же, что используется для данных)
const static uint16_t crc16_table[] =
{
  0x0000, 0xc0c1, 0xc181, 0x0140, 0xc301, 0x03c0, 0x0280, 0xc241,0xc601, 0x06c0, 0x0780, 0xc741, 0x0500, 0xc5c1, 0xc481, 0x0440,
  0xcc01, 0x0cc0, 0x0d80, 0xcd41, 0x0f00, 0xcfc1, 0xce81, 0x0e40,0x0a00, 0xcac1, 0xcb81, 0x0b40, 0xc901, 0x09c0, 0x0880, 0xc841,
  0xd801, 0x18c0, 0x1980, 0xd941, 0x1b00, 0xdbc1, 0xda81, 0x1a40,0x1e00, 0xdec1, 0xdf81, 0x1f40, 0xdd01, 0x1dc0, 0x1c80, 0xdc41,
  0x1400, 0xd4c1, 0xd581, 0x1540, 0xd701, 0x17c0, 0x1680, 0xd641,0xd201, 0x12c0, 0x1380, 0xd341, 0x1100, 0xd1c1, 0xd081, 0x1040,
  0xf001, 0x30c0, 0x3180, 0xf141, 0x3300, 0xf3c1, 0xf281, 0x3240,0x3600, 0xf6c1, 0xf781, 0x3740, 0xf501, 0x35c0, 0x3480, 0xf441,
  0x3c00, 0xfcc1, 0xfd81, 0x3d40, 0xff01, 0x3fc0, 0x3e80, 0xfe41,0xfa01, 0x3ac0, 0x3b80, 0xfb41, 0x3900, 0xf9c1, 0xf881, 0x3840,
  0x2800, 0xe8c1, 0xe981, 0x2940, 0xeb01, 0x2bc0, 0x2a80, 0xea41,0xee01, 0x2ec0, 0x2f80, 0xef41, 0x2d00, 0xedc1, 0xec81, 0x2c40,
  0xe401, 0x24c0, 0x2580, 0xe541, 0x2700, 0xe7c1, 0xe681, 0x2640,0x2200, 0xe2c1, 0xe381, 0x2340, 0xe101, 0x21c0, 0x2080, 0xe041,
  0xa001, 0x60c0, 0x6180, 0xa141, 0x6300, 0xa3c1, 0xa281, 0x6240,0x6600, 0xa6c1, 0xa781, 0x6740, 0xa501, 0x65c0, 0x6480, 0xa441,
  0x6c00, 0xacc1, 0xad81, 0x6d40, 0xaf01, 0x6fc0, 0x6e80, 0xae41,0xaa01, 0x6ac0, 0x6b80, 0xab41, 0x6900, 0xa9c1, 0xa881, 0x6840,
  0x7800, 0xb8c1, 0xb981, 0x7940, 0xbb01, 0x7bc0, 0x7a80, 0xba41,0xbe01, 0x7ec0, 0x7f80, 0xbf41, 0x7d00, 0xbdc1, 0xbc81, 0x7c40,
  0xb401, 0x74c0, 0x7580, 0xb541, 0x7700, 0xb7c1, 0xb681, 0x7640,0x7200, 0xb2c1, 0xb381, 0x7340, 0xb101, 0x71c0, 0x7080, 0xb041,
  0x5000, 0x90c1, 0x9181, 0x5140, 0x9301, 0x53c0, 0x5280, 0x9241,0x9601, 0x56c0, 0x5780, 0x9741, 0x5500, 0x95c1, 0x9481, 0x5440,
  0x9c01, 0x5cc0, 0x5d80, 0x9d41, 0x5f00, 0x9fc1, 0x9e81, 0x5e40,0x5a00, 0x9ac1, 0x9b81, 0x5b40, 0x9901, 0x59c0, 0x5880, 0x9841,
  0x8801, 0x48c0, 0x4980, 0x8941, 0x4b00, 0x8bc1, 0x8a81, 0x4a40,0x4e00, 0x8ec1, 0x8f81, 0x4f40, 0x8d01, 0x4dc0, 0x4c80, 0x8c41,
  0x4400, 0x84c1, 0x8581, 0x4540, 0x8701, 0x47c0, 0x4680, 0x8641,0x8201, 0x42c0, 0x4380, 0x8341, 0x4100, 0x81c1, 0x8081, 0x4040,
};

// Вычисление CRC16 для буфера
uint16_t CalculateCommandCRC(const uint8_t* buffer, size_t length) {
    uint16_t crc_16 = 0xffff;
    for (size_t i = 0; i < length; i++) {
        crc_16 = (crc_16 >> 8) ^ crc16_table[(buffer[i] ^ crc_16) & 0xff];
    }
    return crc_16;
}

// Формирование буфера команды с CRC
size_t BuildCommandBuffer(const Command& cmd, uint8_t* buffer, size_t bufferSize) {
    // Проверяем размер буфера
    size_t requiredSize = 3 + static_cast<size_t>(cmd.dataLength) + 2; // тип + код + длина + данные + CRC
    if (bufferSize < requiredSize) {
        return 0; // Недостаточно места в буфере
    }

    // Заполняем буфер
    size_t offset = 0;
    buffer[offset++] = cmd.commandType;
    buffer[offset++] = cmd.commandCode;
    buffer[offset++] = cmd.dataLength;

    // Копируем данные параметров
    if (cmd.dataLength > 0) {
        memcpy(&buffer[offset], cmd.data, cmd.dataLength);
        offset += cmd.dataLength;
    }

    // Вычисляем CRC для всех данных кроме самого CRC
    uint16_t crc = CalculateCommandCRC(buffer, offset);

    // Добавляем CRC в конец
    memcpy(&buffer[offset], &crc, 2);
    offset += 2;

    return offset; // Возвращаем реальную длину команды
}

// Получение строкового имени команды
const char* GetCommandName(const Command& cmd) {
    // Команды подтверждения телеметрии
    if (cmd.commandType == 0x00) { // CmdType::TELEMETRY
        switch (cmd.commandCode) {
        case 0x01: return "DATA_OK";
        case 0x02: return "DATA_FALSE";
        default: return "TELEMETRY_UNKNOWN";
        }
    }
    // Команды управления программой
    else if (cmd.commandType == 0x01) { // CmdType::PROG_CONTROL
        switch (cmd.commandCode) {
        case 0x01: return "START";
        case 0x02: return "STOP";
        case 0x03: return "PAUSE";
        case 0x04: return "RESUME";
        case 0x05: return "RESET";
        default: return "PROG_CONTROL_UNKNOWN";
        }
    }
    // Команды конфигурации
    else if (cmd.commandType == 0x02) { // CmdType::CONFIGURATION
        switch (cmd.commandCode) {
        case 0x01: return "SET_TEMPERATURE";
        case 0x02: return "SET_INTERVAL";
        case 0x03: return "SET_MODE";
        case 0x04: return "SET_DEFROST_PARAM";
        case 0x05: return "SET_DEFROST_GROUP";
        default: return "CONFIG_UNKNOWN";
        }
    }
    // Команды запроса
    else if (cmd.commandType == 0x03) { // CmdType::REQUEST
        switch (cmd.commandCode) {
        case 0x02: return "GET_VERSION";
        case 0x03: return "GET_DATA";
        case 0x04: return "GET_CMD_INFO";
        case 0x06: return "GET_DEFROST_PARAM";
        case 0x07: return "GET_DEFROST_GROUP";
        case 0x08: return "SEND_STATE";
        default: return "REQUEST_UNKNOWN";
        }
    }
    // Команды управления устройствами
    else if (cmd.commandType == 0x04) { // CmdType::DEVICE_CONTROL
        return "DEVICE_CONTROL";
    }

    return "UNKNOWN_COMMAND";
}

const char* GetCommandTypeName(uint8_t commandType) {
    switch (commandType) {
    case CmdType::TELEMETRY: return "TELEMETRY";
    case CmdType::PROG_CONTROL: return "PROG_CONTROL";
    case CmdType::CONFIGURATION: return "CONFIGURATION";
    case CmdType::REQUEST: return "REQUEST";
    case CmdType::DEVICE_CONTROL: return "DEVICE_CONTROL";
    default: return "UNKNOWN_TYPE";
    }
}

// Defrost params API
Command CreateConfigCommandDefrostSetParam(uint8_t groupId, uint8_t paramId, const DefrostParamValue& value) {
    Command cmd;
    cmd.commandType = CmdType::CONFIGURATION;
    cmd.commandCode = CmdConfig::SET_DEFROST_PARAM;
    cmd.data[0] = groupId;
    cmd.data[1] = paramId;
    cmd.data[2] = value.valueType;
    if (value.valueType == DefrostParamType::U8) {
        cmd.dataLength = 4;
        cmd.data[3] = value.value.u8;
    }
    else if (value.valueType == DefrostParamType::U16) {
        cmd.dataLength = 5;
        memcpy(&cmd.data[3], &value.value.u16, sizeof(uint16_t));
    }
    else if (value.valueType == DefrostParamType::F32) {
        cmd.dataLength = 7;
        memcpy(&cmd.data[3], &value.value.f32, sizeof(float));
    }
    else {
        cmd.dataLength = 0;
    }
    return cmd;
}

Command CreateConfigCommandSetDefrostGroup(uint8_t groupId, const uint8_t* payload, uint8_t payloadLen) {
    Command cmd;
    cmd.commandType = CmdType::CONFIGURATION;
    cmd.commandCode = CmdConfig::SET_DEFROST_GROUP;
    const size_t maxPayload = sizeof(cmd.data) - 1;
    if (payload == nullptr || payloadLen > maxPayload) {
        cmd.dataLength = 0;
        return cmd;
    }
    cmd.dataLength = 1 + payloadLen;
    cmd.data[0] = groupId;
    memcpy(&cmd.data[1], payload, payloadLen);
    return cmd;
}

bool ParseDefrostParamResponse(const CommandResponse& response, uint8_t* outGroupId, uint8_t* outParamId, DefrostParamValue* outValue) {
    if (outValue == nullptr) return false;
    if (response.dataLength < 4) return false;
    if (outGroupId) *outGroupId = response.data[0];
    if (outParamId) *outParamId = response.data[1];
    outValue->valueType = response.data[2];
    if (outValue->valueType == DefrostParamType::U8) {
        if (response.dataLength != 4) return false;
        outValue->value.u8 = response.data[3];
        return true;
    }
    if (outValue->valueType == DefrostParamType::U16) {
        if (response.dataLength != 5) return false;
        memcpy(&outValue->value.u16, &response.data[3], sizeof(uint16_t));
        return true;
    }
    if (outValue->valueType == DefrostParamType::F32) {
        if (response.dataLength != 7) return false;
        memcpy(&outValue->value.f32, &response.data[3], sizeof(float));
        return true;
    }
    return false;
}

// ============================
// Функции для разбора ответа по протоколу
// ============================

// Проверка CRC ответа
bool ValidateResponseCRC(const uint8_t* buffer, size_t length) {
    if (length < 6) { // Минимальная длина ответа: Type + Code + Status + DataLen + CRC
        return false;
    }

    // Вычисляем CRC для всех байт кроме последних 2 байт (сам CRC)
    uint16_t calculatedCRC = CalculateCommandCRC(buffer, length - 2);

    // Сравниваем с полученным CRC (последние 2 байта)
    uint16_t receivedCRC;
    memcpy(&receivedCRC, &buffer[length - 2], 2);

    return (calculatedCRC == receivedCRC);
}

// Разбор буфера ответа в структуру CommandResponse
bool ParseResponseBuffer(const uint8_t* buffer, size_t bufferSize, CommandResponse& response) {
    // Поддерживаются 2 формата ответа:
    // 1) Legacy: [Type][Code][Status][DataLen][Data...][CRC]
    // 2) Новый (во втором байте после Type идёт Len): [Type][Len][Code][Status][DataLen][Data...][CRC]

    if (buffer == nullptr) {
        return false;
    }

    // Сначала пробуем Legacy, чтобы не было ложных срабатываний "new format"
    // (во втором байте в Legacy лежит Code, который может случайно совпасть с Len).
    if (bufferSize >= 6) { // Type + Code + Status + DataLen + CRC(2)
        size_t offset = 0;
        const uint8_t commandType = buffer[offset++];
        const uint8_t commandCode = buffer[offset++];
        const uint8_t status = buffer[offset++];
        const uint8_t dataLength = buffer[offset++];

        const size_t expectedSize = 4 + static_cast<size_t>(dataLength) + 2;
        if (bufferSize >= expectedSize) {
            if (dataLength > sizeof(response.data)) {
                return false;
            }

            response.commandType = commandType;
            response.commandCode = commandCode;
            response.status = status;
            response.dataLength = dataLength;

            if (dataLength > 0) {
                memcpy(response.data, &buffer[offset], dataLength);
            }

            memcpy(&response.crc, &buffer[expectedSize - 2], 2);

            if (ValidateResponseCRC(buffer, expectedSize)) {
                return true;
            }
        }
    }

    // Пробуем новый формат (если Legacy не подошёл).
    if (bufferSize >= 7) { // Type + Len + Code + Status + DataLen + CRC(2)
        const uint8_t len = buffer[1];
        const size_t expectedByLen = static_cast<size_t>(1 + 1 + len + 2);
        if (expectedByLen == bufferSize) {
            // Len включает в себя: Code + Status + DataLen + Data
            if (len < 3) {
                return false;
            }

            response.commandType = buffer[0];
            response.commandCode = buffer[2];
            response.status = buffer[3];
            response.dataLength = buffer[4];

            const size_t expectedLenField = static_cast<size_t>(3 + response.dataLength);
            if (expectedLenField != static_cast<size_t>(len)) {
                return false;
            }

            if (response.dataLength > sizeof(response.data)) {
                return false;
            }

            if (response.dataLength > 0) {
                memcpy(response.data, &buffer[5], response.dataLength);
            }

            memcpy(&response.crc, &buffer[bufferSize - 2], 2);

            if (!ValidateResponseCRC(buffer, bufferSize)) {
                return false;
            }

            return true;
        }
    }

    return false;
}

// Получение строкового имени статуса ответа (английский)
const char* GetStatusName(uint8_t status) {
    switch (status) {
    case 0x00: return "OK";
    case 0x01: return "CRC_ERROR";
    case 0x02: return "INVALID_TYPE";
    case 0x03: return "INVALID_CODE";
    case 0x04: return "INVALID_LENGTH";
    case 0x05: return "EXECUTION_ERROR";
    case 0x06: return "TIMEOUT";
    case 0xFF: return "UNKNOWN_ERROR";
    default: return "UNDEFINED_STATUS";
    }
}

// Получение текстового описания статуса ответа по коду статуса (L"..." для корректной записи в лог UTF-8)
const wchar_t* GetStatusDescription(uint8_t status) {
    switch (status) {
    case 0x00: return L"Команда успешно выполнена";
    case 0x01: return L"Ошибка контрольной суммы CRC. Данные были искажены при передаче";
    case 0x02: return L"Недопустимый тип команды. Устройство не поддерживает данный тип";
    case 0x03: return L"Недопустимый код команды. Код не распознан в данном типе";
    case 0x04: return L"Некорректная длина данных. Команда отклонена по протоколу";
    case 0x05: return L"Ошибка выполнения команды. Устройство не смогло выполнить действие";
    case 0x06: return L"Таймаут выполнения. Устройство не ответило в отведённое время";
    case 0xFF: return L"Неизвестная ошибка. Рекомендуется переподключиться к устройству";
    default: return L"Неопределённый статус ответа";
    }
}

// Команды, которые требуют ответа
bool CommandRequiresResponse(const Command& cmd) {
    // Для команд типа REQUEST ответ обязателен по определению
    if (cmd.commandType == CmdType::REQUEST) {
        return true;
    }

    // Остальные команды по умолчанию также ожидают ответ подтверждения
    // в зависимости от протокола и настроек
    return true; // По умолчанию считаем, что все команды требуют ответа
}
