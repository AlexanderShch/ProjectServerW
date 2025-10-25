# Руководство по обработке ответов контроллера

**Проект:** ProjectServerW  
**Версия:** 1.0  
**Дата:** 25 октября 2025

---

## Содержание

1. [Обзор](#обзор)
2. [Новые структуры и константы](#новые-структуры-и-константы)
3. [API функций](#api-функций)
4. [Примеры использования](#примеры-использования)
5. [Обработка ошибок](#обработка-ошибок)

---

## Обзор

В проект добавлена полная поддержка приема и обработки ответов от контроллера на отправленные команды. Реализация включает:

- **Структуру ответа** (`CommandResponse`)
- **Коды статуса** выполнения команд (`CmdStatus`)
- **Функции разбора** и проверки CRC ответов
- **Методы приема** и обработки ответов в `DataForm`

---

## Новые структуры и константы

### Структура CommandResponse

```cpp
struct CommandResponse {
    uint8_t commandType;           // Тип исходной команды
    uint8_t commandCode;           // Код исходной команды
    uint8_t status;                // Статус выполнения команды
    uint8_t dataLength;            // Длина данных ответа
    uint8_t data[58];              // Данные ответа (макс. 58 байт)
    uint16_t crc;                  // CRC16 для проверки целостности
};
```

### Коды статуса (CmdStatus)

| Код  | Константа           | Описание                        |
|------|---------------------|---------------------------------|
| 0x00 | OK                  | Команда выполнена успешно       |
| 0x01 | CRC_ERROR           | Ошибка контрольной суммы        |
| 0x02 | INVALID_TYPE        | Неизвестный тип команды         |
| 0x03 | INVALID_CODE        | Неизвестный код команды         |
| 0x04 | INVALID_LENGTH      | Неверная длина данных           |
| 0x05 | EXECUTION_ERROR     | Ошибка выполнения команды       |
| 0x06 | TIMEOUT             | Таймаут выполнения              |
| 0xFF | UNKNOWN_ERROR       | Неизвестная ошибка              |

---

## API функций

### Commands.h / Commands.cpp

#### `bool ValidateResponseCRC(const uint8_t* buffer, size_t length)`

**Описание:** Проверяет CRC16 полученного ответа

**Параметры:**
- `buffer` - буфер с данными ответа
- `length` - длина буфера

**Возвращает:** `true` если CRC корректна, `false` если нет

**Пример:**
```cpp
uint8_t buffer[64];
size_t receivedBytes = recv(socket, buffer, sizeof(buffer), 0);

if (ValidateResponseCRC(buffer, receivedBytes)) {
    // CRC корректна, можно обрабатывать данные
}
```

---

#### `bool ParseResponseBuffer(const uint8_t* buffer, size_t bufferSize, CommandResponse& response)`

**Описание:** Разбирает буфер ответа в структуру `CommandResponse`

**Параметры:**
- `buffer` - буфер с данными ответа
- `bufferSize` - размер буфера
- `response` - структура для записи разобранного ответа

**Возвращает:** `true` при успешном разборе, `false` при ошибке

**Пример:**
```cpp
uint8_t buffer[64];
size_t receivedBytes = recv(socket, buffer, sizeof(buffer), 0);

CommandResponse response;
if (ParseResponseBuffer(buffer, receivedBytes, response)) {
    // Ответ успешно разобран
    if (response.status == CmdStatus::OK) {
        // Команда выполнена успешно
    }
}
```

---

#### `const char* GetStatusName(uint8_t status)`

**Описание:** Возвращает текстовое описание кода статуса

**Параметры:**
- `status` - код статуса

**Возвращает:** Строка с названием статуса

**Пример:**
```cpp
CommandResponse response;
// ... прием и разбор ответа ...

const char* statusName = GetStatusName(response.status);
printf("Status: %s\n", statusName); // Выведет: "Status: OK"
```

---

#### `bool CommandRequiresResponse(const Command& cmd)`

**Описание:** Определяет, требует ли команда ответа от контроллера

**Параметры:**
- `cmd` - команда для проверки

**Возвращает:** `true` если команда требует ответа

**Примечание:** Все команды типа `REQUEST` всегда требуют ответа с данными

---

### DataForm методы

#### `bool ReceiveResponse(CommandResponse& response, int timeoutMs = 1000)`

**Описание:** Принимает ответ от контроллера через сокет

**Параметры:**
- `response` - структура для записи полученного ответа
- `timeoutMs` - таймаут ожидания в миллисекундах (по умолчанию 1000 мс)

**Возвращает:** `true` при успешном приеме ответа, `false` при ошибке

**Пример:**
```cpp
CommandResponse response;
if (ReceiveResponse(response, 2000)) {
    // Ответ получен
    ProcessResponse(response);
}
```

---

#### `void ProcessResponse(const CommandResponse& response)`

**Описание:** Обрабатывает полученный ответ от контроллера

**Параметры:**
- `response` - структура с полученным ответом

**Функциональность:**
- Проверяет статус выполнения команды
- Извлекает данные для команд типа REQUEST
- Отображает информационные сообщения
- Логирует результаты

**Пример:**
```cpp
CommandResponse response;
if (ReceiveResponse(response)) {
    ProcessResponse(response);
    // Автоматически отобразится результат выполнения
}
```

---

#### `bool SendCommandAndWaitResponse(const Command& cmd, CommandResponse& response, String^ commandName = nullptr)`

**Описание:** Отправляет команду и ожидает ответ от контроллера

**Параметры:**
- `cmd` - команда для отправки
- `response` - структура для записи ответа
- `commandName` - имя команды для логирования (опционально)

**Возвращает:** `true` если команда успешно выполнена (статус OK), `false` при ошибке

**Пример:**
```cpp
Command cmd = CreateControlCommand(CmdProgControl::START);
CommandResponse response;

if (SendCommandAndWaitResponse(cmd, response)) {
    // Команда успешно выполнена
    buttonSTOPstate_TRUE();
}
```

---

## Примеры использования

### Пример 1: Отправка команды START с ожиданием ответа

```cpp
void ProjectServerW::DataForm::SendStartCommand() {
    // Создаем команду START
    Command cmd = CreateControlCommand(CmdProgControl::START);
    CommandResponse response;
    
    // Отправляем команду и ждем ответ
    if (SendCommandAndWaitResponse(cmd, response)) {
        // Команда успешно выполнена на контроллере
        buttonSTOPstate_TRUE();
        GlobalLogger::LogMessage("Information: Команда START успешно выполнена контроллером");
    } else {
        // Ошибка выполнения команды
        GlobalLogger::LogMessage("Error: Команда START не выполнена контроллером");
    }
}
```

---

### Пример 2: Запрос статуса устройства

```cpp
void ProjectServerW::DataForm::RequestDeviceStatus() {
    // Создаем команду запроса статуса
    Command cmd = CreateRequestCommand(CmdRequest::GET_STATUS);
    CommandResponse response;
    
    // Отправляем команду и ждем ответ
    if (SendCommandAndWaitResponse(cmd, response, "GET_STATUS")) {
        if (response.dataLength >= 2) {
            // Извлекаем статус из ответа
            uint16_t deviceStatus;
            memcpy(&deviceStatus, response.data, 2);
            
            // Отображаем статус
            String^ statusMsg = String::Format("Статус устройства: 0x{0:X4}", deviceStatus);
            MessageBox::Show(statusMsg);
        }
    }
}
```

---

### Пример 3: Запрос версии прошивки

```cpp
void ProjectServerW::DataForm::GetFirmwareVersion() {
    // Создаем команду запроса версии
    Command cmd = CreateRequestCommand(CmdRequest::GET_VERSION);
    CommandResponse response;
    
    // Отправляем команду и ждем ответ
    if (SendCommandAndWaitResponse(cmd, response)) {
        if (response.dataLength > 0) {
            // Извлекаем версию из ответа
            String^ version = gcnew String(
                reinterpret_cast<const char*>(response.data), 
                0, response.dataLength, System::Text::Encoding::ASCII);
            
            // Отображаем версию
            String^ msg = "Версия прошивки: " + version;
            MessageBox::Show(msg);
        }
    }
}
```

---

### Пример 4: Установка температуры с проверкой результата

```cpp
void ProjectServerW::DataForm::SetTargetTemperature(float temperature) {
    // Создаем команду установки температуры
    Command cmd = CreateConfigCommandFloat(CmdConfig::SET_TEMPERATURE, temperature);
    CommandResponse response;
    
    // Отправляем команду и ждем ответ
    if (SendCommandAndWaitResponse(cmd, response)) {
        String^ msg = String::Format(
            "Целевая температура установлена: {0:F1}°C", temperature);
        MessageBox::Show(msg);
    } else {
        // Обработка ошибки
        String^ errorMsg = String::Format(
            "Ошибка установки температуры!\nСтатус: {0}", 
            gcnew String(GetStatusName(response.status)));
        MessageBox::Show(errorMsg);
    }
}
```

---

### Пример 5: Ручная обработка ответа

```cpp
void ProjectServerW::DataForm::ManualResponseHandling() {
    // Создаем и отправляем команду
    Command cmd = CreateControlCommand(CmdProgControl::PAUSE);
    
    if (SendCommand(cmd)) {
        // Команда отправлена, ждем ответ вручную
        CommandResponse response;
        
        if (ReceiveResponse(response, 3000)) { // Таймаут 3 секунды
            // Ответ получен, проверяем статус
            if (response.status == CmdStatus::OK) {
                MessageBox::Show("Программа приостановлена");
            } else if (response.status == CmdStatus::EXECUTION_ERROR) {
                MessageBox::Show("Не удалось приостановить программу");
            } else {
                String^ msg = String::Format("Ошибка: {0}", 
                    gcnew String(GetStatusName(response.status)));
                MessageBox::Show(msg);
            }
        } else {
            MessageBox::Show("Таймаут ожидания ответа");
        }
    }
}
```

---

## Обработка ошибок

### Типы ошибок и их обработка

#### 1. Ошибка CRC (CmdStatus::CRC_ERROR)

Возникает при повреждении данных во время передачи.

**Обработка:**
```cpp
if (response.status == CmdStatus::CRC_ERROR) {
    // Повторная отправка команды
    SendCommandAndWaitResponse(cmd, response);
}
```

---

#### 2. Неверный тип/код команды (INVALID_TYPE, INVALID_CODE)

Контроллер не поддерживает данную команду.

**Обработка:**
```cpp
if (response.status == CmdStatus::INVALID_TYPE || 
    response.status == CmdStatus::INVALID_CODE) {
    MessageBox::Show("Команда не поддерживается контроллером");
    // Проверьте версию прошивки контроллера
}
```

---

#### 3. Таймаут приема ответа

Ответ от контроллера не получен в течение заданного времени.

**Обработка:**
```cpp
CommandResponse response;
if (!ReceiveResponse(response, 5000)) {
    // Увеличьте таймаут или проверьте соединение
    MessageBox::Show("Контроллер не отвечает. Проверьте соединение.");
}
```

---

#### 4. Несоответствие ответа отправленной команде

Ответ не соответствует отправленной команде.

**Обработка:**
```cpp
if (response.commandType != cmd.commandType || 
    response.commandCode != cmd.commandCode) {
    MessageBox::Show("Получен некорректный ответ от контроллера");
    // Возможно, в очереди находится ответ на предыдущую команду
}
```

---

## Рекомендации по использованию

### 1. Выбор метода отправки

- **Используйте `SendCommand()`** когда ответ не требуется или обрабатывается отдельно
- **Используйте `SendCommandAndWaitResponse()`** для команд, требующих подтверждения
- **Используйте `ReceiveResponse()` + `ProcessResponse()`** для ручной обработки

### 2. Установка таймаутов

- **Команды управления:** 1000-2000 мс
- **Команды запроса данных:** 2000-3000 мс
- **Команды конфигурации:** 1000-2000 мс

### 3. Логирование

Все методы автоматически логируют:
- Отправку команд
- Прием ответов
- Ошибки выполнения

Логи сохраняются в файл `log.txt` в директории приложения.

---

## Контактная информация

**Проект:** ProjectServerW  
**Дата создания:** 25 октября 2025  
**Версия документа:** 1.0

---

*Конец документа*

