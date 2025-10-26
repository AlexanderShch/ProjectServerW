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
    // Команды управления программой
    if (cmd.commandType == 0x01) { // CmdType::PROG_CONTROL
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
        default: return "CONFIG_UNKNOWN";
        }
    }
    // Команды запроса
    else if (cmd.commandType == 0x03) { // CmdType::REQUEST
        switch (cmd.commandCode) {
        case 0x01: return "GET_STATUS";
        case 0x02: return "GET_VERSION";
        case 0x03: return "GET_DATA";
        default: return "REQUEST_UNKNOWN";
        }
    }
    // Команды управления устройствами
    else if (cmd.commandType == 0x04) { // CmdType::DEVICE_CONTROL
        return "DEVICE_CONTROL";
    }

    return "UNKNOWN_COMMAND";
}

// ============================
// Функции для обработки ответов от контроллера
// ============================

// Проверка CRC ответа
bool ValidateResponseCRC(const uint8_t* buffer, size_t length) {
    if (length < 6) { // Минимальный размер ответа: Type + Code + Status + DataLen + CRC
        return false;
    }

    // Вычисляем CRC для всех данных кроме последних 2 байт (самого CRC)
    uint16_t calculatedCRC = CalculateCommandCRC(buffer, length - 2);

    // Извлекаем полученный CRC (последние 2 байта)
    uint16_t receivedCRC;
    memcpy(&receivedCRC, &buffer[length - 2], 2);

    return (calculatedCRC == receivedCRC);
}

// Разбор буфера ответа в структуру CommandResponse
bool ParseResponseBuffer(const uint8_t* buffer, size_t bufferSize, CommandResponse& response) {
    // Проверяем минимальный размер буфера
    if (bufferSize < 6) { // Type + Code + Status + DataLen + CRC(2)
        return false;
    }

    // Извлекаем поля заголовка
    size_t offset = 0;
    response.commandType = buffer[offset++];
    response.commandCode = buffer[offset++];
    response.status = buffer[offset++];
    response.dataLength = buffer[offset++];

    // Проверяем, соответствует ли длина данных размеру буфера
    size_t expectedSize = 4 + static_cast<size_t>(response.dataLength) + 2; // Type + Code + Status + DataLen + Data + CRC
    if (bufferSize < expectedSize) {
        return false;
    }

    // Копируем данные ответа
    if (response.dataLength > 0) {
        if (response.dataLength > sizeof(response.data)) {
            return false; // Данные не помещаются в структуру
        }
        memcpy(response.data, &buffer[offset], response.dataLength);
        offset += response.dataLength;
    }

    // Извлекаем CRC
    memcpy(&response.crc, &buffer[offset], 2);

    // Проверяем CRC
    if (!ValidateResponseCRC(buffer, bufferSize)) {
        return false;
    }

    return true;
}

// Получение строкового описания статуса (краткое)
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

// Получение детального описания ошибки на русском языке
const char* GetStatusDescription(uint8_t status) {
    switch (status) {
    case 0x00: return "Команда выполнена успешно";
    case 0x01: return "Ошибка контрольной суммы CRC. Данные повреждены при передаче";
    case 0x02: return "Неизвестный тип команды. Контроллер не поддерживает данный тип";
    case 0x03: return "Неизвестный код команды. Команда не реализована в контроллере";
    case 0x04: return "Неверная длина данных. Размер параметров не соответствует ожидаемому";
    case 0x05: return "Ошибка выполнения команды. Контроллер не смог выполнить операцию";
    case 0x06: return "Таймаут выполнения. Контроллер не успел выполнить операцию в отведенное время";
    case 0xFF: return "Неизвестная ошибка. Произошла непредвиденная ошибка в контроллере";
    default: return "Неопределенный статус ответа";
    }
}

// Проверка, требует ли команда ответа
bool CommandRequiresResponse(const Command& cmd) {
    // Все команды типа REQUEST всегда требуют ответа с данными
    if (cmd.commandType == CmdType::REQUEST) {
        return true;
    }

    // Остальные команды могут требовать только подтверждения
    // В зависимости от конфигурации системы
    return true; // По умолчанию считаем, что все команды требуют ответа
}