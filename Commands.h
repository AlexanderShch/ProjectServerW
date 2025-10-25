#pragma once
#include <cstdint>
#include <cstring>

// Определение типов команд
struct CmdType {
    static const uint8_t PROG_CONTROL = 0x01;      // Команды управления программой (СТАРТ, СТОП и т.д.)
    static const uint8_t CONFIGURATION = 0x02;     // Команды конфигурации
    static const uint8_t REQUEST = 0x03;           // Команды запроса данных
    static const uint8_t DEVICE_CONTROL = 0x04;    // Команды управления устройствами
};

// Коды команд управления (тип PROG_CONTROL)
struct CmdProgControl {
    static const uint8_t START = 0x01;         // Запуск работы
    static const uint8_t STOP = 0x02;          // Остановка работы
    static const uint8_t PAUSE = 0x03;         // Пауза
    static const uint8_t RESUME = 0x04;        // Возобновление после паузы
    static const uint8_t RESET = 0x05;         // Сброс устройства
};

// Коды команд конфигурации (тип CONFIGURATION)
struct CmdConfig {
    static const uint8_t SET_TEMPERATURE = 0x01;   // Установить температуру
    static const uint8_t SET_INTERVAL = 0x02;      // Установить интервал измерений
    static const uint8_t SET_MODE = 0x03;          // Установить режим работы
};

// Коды команд запроса (тип REQUEST)
struct CmdRequest {
    static const uint8_t GET_STATUS = 0x01;        // Запросить статус
    static const uint8_t GET_VERSION = 0x02;       // Запросить версию прошивки
    static const uint8_t GET_DATA = 0x03;          // Запросить данные
};

// Максимальный размер команды
const size_t MAX_COMMAND_SIZE = 64;

// Структура команды
struct Command {
    uint8_t commandType;           // Тип команды
    uint8_t commandCode;           // Код команды
    uint8_t dataLength;            // Длина данных параметров
    uint8_t data[MAX_COMMAND_SIZE - 5]; // Данные параметров (макс. 59 байт)
    uint16_t crc;                  // CRC16 для проверки целостности

    Command() : commandType(0), commandCode(0), dataLength(0), crc(0) {
        memset(data, 0, sizeof(data));
    }
};

// Функция для вычисления CRC16 команды
uint16_t CalculateCommandCRC(const uint8_t* buffer, size_t length);

// Функция для создания буфера команды с CRC
size_t BuildCommandBuffer(const Command& cmd, uint8_t* buffer, size_t bufferSize);

// Функция для получения строкового имени команды
const char* GetCommandName(const Command& cmd);

// Вспомогательные функции для создания команд

// Создать команду управления без параметров
inline Command CreateControlCommand(uint8_t commandCode) {
    Command cmd;
    cmd.commandType = CmdType::PROG_CONTROL;
    cmd.commandCode = commandCode;
    cmd.dataLength = 0;
    return cmd;
}

// Создать команду управления с одним байтом параметра
inline Command CreateControlCommandWithParam(uint8_t commandCode, uint8_t param) {
    Command cmd;
    cmd.commandType = CmdType::PROG_CONTROL;
    cmd.commandCode = commandCode;
    cmd.dataLength = 1;
    cmd.data[0] = param;
    return cmd;
}

// Создать команду конфигурации с целочисленным параметром (4 байта)
inline Command CreateConfigCommandInt(uint8_t commandCode, int32_t value) {
    Command cmd;
    cmd.commandType = CmdType::CONFIGURATION;
    cmd.commandCode = commandCode;
    cmd.dataLength = 4;
    memcpy(cmd.data, &value, 4);
    return cmd;
}

// Создать команду конфигурации с параметром типа float
inline Command CreateConfigCommandFloat(uint8_t commandCode, float value) {
    Command cmd;
    cmd.commandType = CmdType::CONFIGURATION;
    cmd.commandCode = commandCode;
    cmd.dataLength = 4;
    memcpy(cmd.data, &value, 4);
    return cmd;
}

// Создать команду запроса без параметров
inline Command CreateRequestCommand(uint8_t commandCode) {
    Command cmd;
    cmd.commandType = CmdType::REQUEST;
    cmd.commandCode = commandCode;
    cmd.dataLength = 0;
    return cmd;
}

// ============================
// Структуры и функции для обработки ответов от контроллера
// ============================

// Коды статуса выполнения команд
struct CmdStatus {
    static const uint8_t OK = 0x00;                   // Команда выполнена успешно
    static const uint8_t CRC_ERROR = 0x01;            // Ошибка контрольной суммы
    static const uint8_t INVALID_TYPE = 0x02;         // Неизвестный тип команды
    static const uint8_t INVALID_CODE = 0x03;         // Неизвестный код команды
    static const uint8_t INVALID_LENGTH = 0x04;       // Неверная длина данных
    static const uint8_t EXECUTION_ERROR = 0x05;      // Ошибка выполнения команды
    static const uint8_t TIMEOUT = 0x06;              // Таймаут выполнения
    static const uint8_t UNKNOWN_ERROR = 0xFF;        // Неизвестная ошибка
};

// Структура ответа от контроллера
struct CommandResponse {
    uint8_t commandType;           // Тип исходной команды
    uint8_t commandCode;           // Код исходной команды
    uint8_t status;                // Статус выполнения команды
    uint8_t dataLength;            // Длина данных ответа
    uint8_t data[MAX_COMMAND_SIZE - 6]; // Данные ответа (макс. 58 байт)
    uint16_t crc;                  // CRC16 для проверки целостности

    CommandResponse() : commandType(0), commandCode(0), status(0), dataLength(0), crc(0) {
        memset(data, 0, sizeof(data));
    }
};

// Функция для проверки CRC ответа
bool ValidateResponseCRC(const uint8_t* buffer, size_t length);

// Функция для разбора буфера ответа в структуру CommandResponse
bool ParseResponseBuffer(const uint8_t* buffer, size_t bufferSize, CommandResponse& response);

// Функция для получения строкового описания статуса
const char* GetStatusName(uint8_t status);

// Функция для проверки, требует ли команда ответа
bool CommandRequiresResponse(const Command& cmd);

