# Краткая шпаргалка: Обработка ответов контроллера

## Актуальный порядок работы

В рабочем режиме сервер инициирует обмен. Команда сервера имеет формат [Type][Code][DataLen][Data][CRC16], ответ контроллера — [AA 55][Type][Code][Status][DataLen][Data][CRC16]. CRC Modbus вычисляется без AA 55 и передаётся младшим байтом вперёд. Длина поля DataLen относится только к Data.

При запуске STM32 сервер запрашивает версию и группы параметров 5, 6 и 7, затем устанавливает интервал. Запрос группы 7 ограничен общим бюджетом 950 мс. После запуска сервер периодически отправляет SEND_STATE (03/08) с байтом 01. Ответ 03/08 содержит 60 байт базовой телеметрии и байт флагов. Бит 0 (смещение 60 в Data) означает, что есть события. Старый запрос без параметров получает 60 байт. Для получения флага отдельной команды нет.

Сервер отвечает на телеметрию DATA_OK или DATA_FALSE (00/01 или 00/02). В авторежиме после DATA_OK контроллер отправляет лог регулирования 00/01 с 113 байтами данных. Сервер также читает прежние варианты лога с 89, 105 и 109 байтами. Повтор неизменившейся телеметрии с ошибочной CRC сервер подтверждает DATA_OK и отбрасывает.

В RAM контроллера 16 мест общей очереди. Диагностика занимает до 12 мест, четыре места доступны кнопкам даже при заполнении диагностикой. Записываются обработанные физические СТАРТ и СТОП, первый отказ активного датчика в серии и восстановление. Переполнения считаются раздельно для кнопок и диагностики. Перезагрузка очищает очередь.

По флагу сервер читает головную запись командой GET_EVENT (03/0D) и записывает её в log.txt. После принудительной записи файла сервер отправляет ACK_EVENT (03/0E), и контроллер удаляет выданную запись. Повтор не создаёт дубликата строки благодаря идентификатору загрузки и номеру события. При наличии следующих событий сервер продолжает чтение. Флаг служебный и в WorkData не экспортируется.

GET_EVENT принимает 16-байтовый кандидат идентификатора сессии. Ответ содержит идентификатор (16 байт), потери кнопок (4), потери диагностики (4), количество записей (1). При непустой очереди идут номер (4), секунды от запуска (4) и тип (1). Типы 1 и 2 — СТАРТ и СТОП, типы 3 и 4 — сбой и восстановление. Длина данных ответа — 25, 34 или 61 байт. ACK_EVENT принимает идентификатор (16) и номер (4). Многобайтовые поля имеют порядок little-endian.

Дополнение для датчика занимает 27 байт. Первые восемь однобайтовых полей содержат индекс, адрес, функцию, причину, ожидаемую и принятую длину, код UART и длину снимка. Далее идут время ответа в мс (2 байта), число неудачных попыток и длительность серии в секундах (по 4), затем снимок кадра (9 байт, значащие байты указаны длиной снимка).

Коды причины 0–7 обозначают отсутствие ответа, поздний ответ, незавершённый кадр, ошибку UART, короткий кадр, ошибку CRC, ошибку формата и прочий сбой обмена. Коды 8–13 обозначают чужой адрес, неверную функцию, неверный ByteCount, лишние байты, исключение Modbus и неверное эхо записи. Код 255 соответствует успешному ответу при восстановлении.

Для прошивок без байта флагов сервер использует прежнее чтение кнопок GET_BUTTON_EVENTS (03/0B) с интервалом не менее 30 с и подтверждение ACK_BUTTON_EVENTS (03/0C). После INVALID_CODE опрос возобновляется при новом подключении или обнаружении перезапуска.


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
Command cmd = CreateRequestCommand(CmdRequest::GET_CMD_INFO);
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

// Установка интервала (uint16)
Command cmd = CreateConfigCommandU16(CmdConfig::SET_INTERVAL, 1000);

// Заводские параметры дефростации (без payload; прошивка >= 2.7.0)
Command cmd = CreateConfigCommand(CmdConfig::LOAD_DEFROST_DEFAULTS);
```

### Команды запроса (REQUEST)

```cpp
// GET_VERSION, GET_DATA, GET_CMD_INFO, GET_ALARM_FLAGS
Command cmd = CreateRequestCommand(CmdRequest::GET_CMD_INFO);
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

### Аудит последней команды (6 байт)

```cpp
// Payload зависит от команды; для GET_CMD_INFO ожидается 6 байт
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

### Запрос аварийных регистров (4 байта)

```cpp
// Для CmdRequest::GET_ALARM_FLAGS ожидается ровно 4 байта:
// [0..1] Device_AlarmFlags (uint16, little-endian)
// [2..3] Sensor_AlarmFlags (uint16, little-endian)
if (response.dataLength == 4) {
    uint16_t deviceFlags = 0;
    uint16_t sensorFlags = 0;
    memcpy(&deviceFlags, &response.data[0], sizeof(uint16_t));
    memcpy(&sensorFlags, &response.data[2], sizeof(uint16_t));
}
```

### Group 6 (GET_DEFROST_GROUP): новый параметр автостопа

```cpp
// После fishColdTarget_C в DefrostLogGlobalPayload_t добавлен uint8:
// debugDisableTargetTStop
//
// 0 -> автостоп по fishColdTarget_C включён
// 1 -> автостоп отключён (debug)
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

