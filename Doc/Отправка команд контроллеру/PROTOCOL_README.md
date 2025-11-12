# Протокол команд контроллера STM32F429

## Что добавлено в проект

В проект **ProjectServerW** добавлена полная поддержка протокола команд для управления контроллером STM32F429.

## Основные файлы

### Изменены:
- **SServer.h** - добавлены структуры протокола и класс `ControllerProtocol`
- **SServer.cpp** - реализация функций протокола команд
- **DataForm.h** - добавлены методы для отправки команд контроллеру
- **DataForm.cpp** - реализация методов отправки команд

### Новые файлы:
- **CONTROLLER_PROTOCOL_USAGE.md** - подробное руководство с примерами использования

## Ключевые возможности

### 1. Полнодуплексная связь
STM32F429 **может одновременно**:
- ? Слушать COM-порт и получать команды
- ? Передавать данные через COM-порт (отправлять ответы)

**НЕ нужно** останавливать прослушивание!

### 2. Типы команд

#### Управление программой (PROG_CONTROL)
```cpp
SendControllerCommand_StartProgram();   // Запустить программу
SendControllerCommand_StopProgram();    // Остановить программу
SendControllerCommand_PauseProgram();   // Приостановить
SendControllerCommand_ResumeProgram();  // Возобновить
```

#### Управление устройствами (DEVICE_CONTROL)
```cpp
SendControllerCommand_SetRelay(1, true);    // Включить реле №1
SendControllerCommand_SetDefrost(true);     // Включить дефрост
```

#### Запросы данных (REQUEST)
```cpp
ControllerResponse response;
SendControllerCommand_GetStatus(response);   // Получить статус
SendControllerCommand_GetSensors(response);  // Данные сенсоров
SendControllerCommand_GetVersion(response);  // Версия ПО
```

### 3. Автоматическая обработка
- ? Автоматический расчет CRC16 (ModBus)
- ? Проверка целостности данных
- ? Обработка ошибок и таймаутов
- ? Логирование всех операций

### 4. Безопасность
- Проверка валидности ответов
- Обработка всех возможных ошибок
- Защита от переполнения буферов
- Таймауты для предотвращения зависаний

## Быстрый старт

### Пример 1: Запуск программы
```cpp
// В обработчике кнопки
bool success = SendControllerCommand_StartProgram();
if (success) {
    MessageBox::Show("Программа запущена");
}
```

### Пример 2: Управление реле
```cpp
// Включить реле №1
SendControllerCommand_SetRelay(1, true);

// Выключить реле №2
SendControllerCommand_SetRelay(2, false);
```

### Пример 3: Получение данных сенсоров
```cpp
ControllerResponse response;
if (SendControllerCommand_GetSensors(response)) {
    // Обработка данных из response.data
    // Длина данных: response.dataLen
}
```

## Формат протокола

### Команда (сервер ? контроллер)
```
[Type][Code][DataLen][Data...][CRC16]
 1б    1б    1б       0-59б    2б
```

### Ответ (контроллер ? сервер)
```
[Type][Code][Status][DataLen][Data...][CRC16]
 1б    1б    1б      1б       0-59б    2б
```

## Статусы ответов

| Код  | Описание                    |
|------|-----------------------------|
| 0x00 | ? Команда выполнена успешно |
| 0x01 | ? Ошибка CRC                |
| 0x02 | ? Неизвестный тип команды   |
| 0x03 | ? Неизвестный код команды   |
| 0x04 | ? Неверная длина данных     |
| 0x05 | ? Ошибка выполнения         |
| 0x06 | ? Таймаут                   |
| 0xFF | ? Неизвестная ошибка        |

## Доступные методы в DataForm

### Команды управления программой
- `bool SendControllerCommand_StartProgram()`
- `bool SendControllerCommand_StopProgram()`
- `bool SendControllerCommand_PauseProgram()`
- `bool SendControllerCommand_ResumeProgram()`

### Команды управления устройствами
- `bool SendControllerCommand_SetRelay(uint8_t relayNum, bool state)`
- `bool SendControllerCommand_SetDefrost(bool enable)`

### Запросы данных
- `bool SendControllerCommand_GetStatus(ControllerResponse& response)`
- `bool SendControllerCommand_GetSensors(ControllerResponse& response)`
- `bool SendControllerCommand_GetVersion(ControllerResponse& response)`

### Универсальные методы
- `bool SendControllerCommandWithResponse(uint8_t cmdType, uint8_t cmdCode, const uint8_t* data, uint8_t dataLen, ControllerResponse& response)`
- `bool SendControllerCommandNoResponse(uint8_t cmdType, uint8_t cmdCode, const uint8_t* data, uint8_t dataLen)`

## Логирование

Все операции логируются в файл `log.txt`:
```
25.10.25 12:30:45 : Command created: Type=0x01, Code=0x01, DataLen=0, CRC=0x1234
25.10.25 12:30:45 : Command sent successfully: 5 bytes
25.10.25 12:30:45 : Response received: Type=0x01, Code=0x01, Status=0x00, DataLen=0
25.10.25 12:30:45 : Command completed successfully: OK: Команда выполнена успешно
```

## Обработка ошибок

```cpp
bool success = SendControllerCommand_StartProgram();

if (!success) {
    // Ошибка уже залогирована
    // Сообщение отображено в Label_Data
    MessageBox::Show("Не удалось выполнить команду", "Ошибка");
}
```

## Важные замечания

1. **Таймаут по умолчанию**: 5000 мс (5 секунд)
2. **Максимальная длина данных**: 59 байт
3. **Общая длина пакета**: 64 байта
4. **CRC16**: Автоматически вычисляется (ModBus)
5. **Потокобезопасность**: Методы безопасны для вызова из GUI потока

## Дополнительная документация

Подробное руководство с большим количеством примеров находится в файле:
**CONTROLLER_PROTOCOL_USAGE.md**

## Поддержка

При возникновении проблем проверьте:
1. Соединение с контроллером (`clientSocket != INVALID_SOCKET`)
2. Логи в файле `log.txt`
3. Сообщения в `Label_Data` на форме

---

**Дата создания**: 25 октября 2025  
**Версия**: 1.0.0  
**Проект**: ProjectServerW - Defrost Control System

