# Краткая шпаргалка: Обработка ответов контроллера

## Быстрый старт

### 1. Отправка команды с ожиданием ответа (рекомендуется)

```cpp
// Создаем команду
Command cmd = CreateControlCommand(CmdProgControl::START);
CommandResponse response;

// Отправляем и ждем ответ
if (SendCommandAndWaitResponse(cmd, response)) {
    // Успех - команда выполнена (status == OK)
} else {
    // Ошибка - команда не выполнена
}
```

---

### 2. Отправка команды без ожидания

```cpp
// Старый способ (без проверки ответа)
Command cmd = CreateControlCommand(CmdProgControl::STOP);
SendCommand(cmd);
```

---

### 3. Ручная обработка ответа

```cpp
// Отправляем команду
Command cmd = CreateRequestCommand(CmdRequest::GET_STATUS);
SendCommand(cmd);

// Принимаем ответ
CommandResponse response;
if (ReceiveResponse(response, 2000)) {  // Таймаут 2 сек (явно указан)
    ProcessResponse(response);  // Автоматическая обработка
}

// Или с таймаутом по умолчанию (1000 мс)
if (ReceiveResponse(response)) {
    ProcessResponse(response);
}
```

---

## Создание команд

### Команды управления (PROG_CONTROL)

```cpp
// START, STOP, PAUSE, RESUME, RESET
Command cmd = CreateControlCommand(CmdProgControl::START);
```

### Команды конфигурации (CONFIGURATION)

```cpp
// Установка температуры (float)
Command cmd = CreateConfigCommandFloat(CmdConfig::SET_TEMPERATURE, 25.5f);

// Установка интервала (int)
Command cmd = CreateConfigCommandInt(CmdConfig::SET_INTERVAL, 1000);
```

### Команды запроса (REQUEST)

```cpp
// GET_STATUS, GET_VERSION, GET_DATA
Command cmd = CreateRequestCommand(CmdRequest::GET_STATUS);
```

---

## Проверка статуса ответа

```cpp
CommandResponse response;
// ... прием ответа ...

if (response.status == CmdStatus::OK) {
    // Успешно выполнено
} else if (response.status == CmdStatus::CRC_ERROR) {
    // Ошибка CRC
} else if (response.status == CmdStatus::TIMEOUT) {
    // Таймаут
} else {
    // Другая ошибка
    const char* statusName = GetStatusName(response.status);
}
```

---

## Извлечение данных из ответа

### Запрос статуса (2 байта)

```cpp
if (response.dataLength >= 2) {
    uint16_t status;
    memcpy(&status, response.data, 2);
    // Используем status
}
```

### Запрос версии (строка)

```cpp
if (response.dataLength > 0) {
    String^ version = gcnew String(
        reinterpret_cast<const char*>(response.data), 
        0, response.dataLength, System::Text::Encoding::ASCII);
}
```

### Запрос конфигурации (1 байт)

```cpp
if (response.dataLength >= 1) {
    uint8_t mode = response.data[0];
    // 0 = Автоматический, 1 = Ручной
}
```

---

## Типичные сценарии

### Сценарий 1: Запуск программы

```cpp
void DataForm::SendStartCommand() {
    Command cmd = CreateControlCommand(CmdProgControl::START);
    CommandResponse response;
    
    if (SendCommandAndWaitResponse(cmd, response)) {
        buttonSTOPstate_TRUE();
        MessageBox::Show("Программа запущена!");
    }
}
```

### Сценарий 2: Запрос и отображение статуса

```cpp
void DataForm::ShowDeviceStatus() {
    Command cmd = CreateRequestCommand(CmdRequest::GET_STATUS);
    CommandResponse response;
    
    if (SendCommandAndWaitResponse(cmd, response)) {
        uint16_t status;
        memcpy(&status, response.data, 2);
        String^ msg = String::Format("Статус: 0x{0:X4}", status);
        MessageBox::Show(msg);
    }
}
```

### Сценарий 3: Установка температуры

```cpp
void DataForm::SetTemperature(float temp) {
    Command cmd = CreateConfigCommandFloat(CmdConfig::SET_TEMPERATURE, temp);
    CommandResponse response;
    
    if (SendCommandAndWaitResponse(cmd, response)) {
        String^ msg = String::Format("Температура установлена: {0:F1}°C", temp);
        MessageBox::Show(msg);
    }
}
```

---

## Таймауты (рекомендуемые)

| Тип команды            | Таймаут    |
|------------------------|------------|
| Команды управления     | 1000-2000 мс |
| Команды конфигурации   | 1000-2000 мс |
| Команды запроса данных | 2000-3000 мс |

---

## Коды статуса

| Код  | Константа           | Значение                  |
|------|---------------------|---------------------------|
| 0x00 | CmdStatus::OK       | ? Успешно                 |
| 0x01 | CmdStatus::CRC_ERROR | ? Ошибка CRC             |
| 0x02 | CmdStatus::INVALID_TYPE | ? Неверный тип        |
| 0x03 | CmdStatus::INVALID_CODE | ? Неверный код        |
| 0x04 | CmdStatus::INVALID_LENGTH | ? Неверная длина    |
| 0x05 | CmdStatus::EXECUTION_ERROR | ? Ошибка выполнения |
| 0x06 | CmdStatus::TIMEOUT | ? Таймаут                  |
| 0xFF | CmdStatus::UNKNOWN_ERROR | ? Неизвестная ошибка |

---

## Отладка

### Проверка CRC

```cpp
uint8_t buffer[64];
size_t bytesReceived = recv(socket, buffer, sizeof(buffer), 0);

if (ValidateResponseCRC(buffer, bytesReceived)) {
    // CRC корректна
} else {
    // Ошибка CRC - проверьте линию связи
}
```

### Логирование

Все операции автоматически логируются в `log.txt`:
- Отправка команд
- Прием ответов
- Ошибки выполнения

---

## Важные замечания

?? **Внимание:**
- Все функции автоматически проверяют CRC
- Ошибки отображаются через `MessageBox` и логируются
- Метод `ProcessResponse()` автоматически обрабатывает все типы ответов
- Используйте `SendCommandAndWaitResponse()` для критичных команд

? **Рекомендации:**
- Всегда проверяйте возвращаемое значение функций
- Используйте адекватные таймауты
- Логируйте все действия для отладки
- Обрабатывайте все возможные статусы ответов

---

**Дата:** 25 октября 2025  
**Версия:** 1.0

