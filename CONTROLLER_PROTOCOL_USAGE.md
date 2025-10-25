# Руководство по использованию протокола команд контроллера

## Обзор

В проект ProjectServerW добавлена полная поддержка протокола команд для управления контроллером STM32F429. Протокол использует полнодуплексную связь через TCP/IP сокеты, что позволяет одновременно отправлять команды и получать данные.

## Важно: Полнодуплексная связь

**STM32F429 может одновременно:**
- ? Слушать COM-порт (получать новые команды)
- ? Передавать данные через COM-порт (отправлять ответы)

**НЕ нужно** останавливать прослушивание порта на время передачи данных!

## Структура протокола

### Команда (от сервера к контроллеру)
```
????????????????????????????????????????????????????????
?  Байт 0  ?  Байт 1  ?  Байт 2   ? Байт 3-N ? N+1, N+2?
????????????????????????????????????????????????????????
?   Type   ?   Code   ?  DataLen  ?   Data   ?  CRC16  ?
? (1 байт) ? (1 байт) ?  (1 байт) ?(0-59 б.) ?(2 байта)?
????????????????????????????????????????????????????????
```

### Ответ (от контроллера к серверу)
```
???????????????????????????????????????????????????????????????????
?  Байт 0  ?  Байт 1  ?  Байт 2  ?  Байт 3   ? Байт 4-N ? N+1, N+2?
???????????????????????????????????????????????????????????????????
?   Type   ?   Code   ?  Status  ?  DataLen  ?   Data   ?  CRC16  ?
? (1 байт) ? (1 байт) ? (1 байт) ?  (1 байт) ?(0-59 б.) ?(2 байта)?
???????????????????????????????????????????????????????????????????
```

## Типы команд

### 1. Команды управления программой (PROG_CONTROL = 0x01)
- `CMD_START_PROGRAM` (0x01) - Запустить программу
- `CMD_STOP_PROGRAM` (0x02) - Остановить программу
- `CMD_PAUSE_PROGRAM` (0x03) - Приостановить программу
- `CMD_RESUME_PROGRAM` (0x04) - Возобновить программу

### 2. Команды управления устройствами (DEVICE_CONTROL = 0x02)
- `CMD_SET_RELAY` (0x01) - Управление реле
- `CMD_SET_DEFROST` (0x02) - Управление дефростом

### 3. Команды конфигурации (CONFIGURATION = 0x03)
- `CMD_SET_PARAM` (0x01) - Установить параметр
- `CMD_GET_PARAM` (0x02) - Получить параметр

### 4. Запросы данных (REQUEST = 0x04)
- `CMD_GET_STATUS` (0x01) - Получить статус
- `CMD_GET_SENSORS` (0x02) - Получить данные сенсоров
- `CMD_GET_VERSION` (0x03) - Получить версию ПО

## Статусы выполнения

| Код  | Название | Описание |
|------|----------|----------|
| 0x00 | CMD_STATUS_OK | ? Команда выполнена успешно |
| 0x01 | CMD_STATUS_CRC_ERROR | ? Ошибка контрольной суммы |
| 0x02 | CMD_STATUS_INVALID_TYPE | ? Неизвестный тип команды |
| 0x03 | CMD_STATUS_INVALID_CODE | ? Неизвестный код команды |
| 0x04 | CMD_STATUS_INVALID_LENGTH | ? Неверная длина данных |
| 0x05 | CMD_STATUS_EXECUTION_ERROR | ? Ошибка выполнения команды |
| 0x06 | CMD_STATUS_TIMEOUT | ? Таймаут выполнения |
| 0xFF | CMD_STATUS_UNKNOWN_ERROR | ? Неизвестная ошибка |

## Примеры использования в DataForm

### Пример 1: Запуск программы контроллера

```cpp
// В обработчике кнопки START
private: System::Void buttonStartController_Click(System::Object^ sender, System::EventArgs^ e) 
{
    // Отправляем команду запуска программы
    bool success = SendControllerCommand_StartProgram();
    
    if (success) {
        // Команда выполнена успешно
        // Состояние кнопок уже изменено внутри функции
        MessageBox::Show("Программа контроллера запущена", "Успех", 
                        MessageBoxButtons::OK, MessageBoxIcon::Information);
    } else {
        // Произошла ошибка
        MessageBox::Show("Не удалось запустить программу контроллера", "Ошибка", 
                        MessageBoxButtons::OK, MessageBoxIcon::Error);
    }
}
```

### Пример 2: Остановка программы контроллера

```cpp
// В обработчике кнопки STOP
private: System::Void buttonStopController_Click(System::Object^ sender, System::EventArgs^ e) 
{
    // Отправляем команду остановки программы
    bool success = SendControllerCommand_StopProgram();
    
    if (success) {
        MessageBox::Show("Программа контроллера остановлена", "Успех", 
                        MessageBoxButtons::OK, MessageBoxIcon::Information);
    }
}
```

### Пример 3: Управление реле

```cpp
// Включение реле №1
private: System::Void buttonRelay1On_Click(System::Object^ sender, System::EventArgs^ e) 
{
    uint8_t relayNum = 1;
    bool state = true; // включить
    
    bool success = SendControllerCommand_SetRelay(relayNum, state);
    
    if (success) {
        Label_Data->Text = "Реле 1 включено";
    }
}

// Выключение реле №1
private: System::Void buttonRelay1Off_Click(System::Object^ sender, System::EventArgs^ e) 
{
    uint8_t relayNum = 1;
    bool state = false; // выключить
    
    bool success = SendControllerCommand_SetRelay(relayNum, state);
    
    if (success) {
        Label_Data->Text = "Реле 1 выключено";
    }
}
```

### Пример 4: Управление дефростом

```cpp
// Включение дефроста
private: System::Void buttonDefrostOn_Click(System::Object^ sender, System::EventArgs^ e) 
{
    bool success = SendControllerCommand_SetDefrost(true);
    
    if (success) {
        Label_Data->Text = "Дефрост включен";
        // Можно изменить цвет индикатора
        labelDefrostStatus->BackColor = System::Drawing::Color::Lime;
    }
}

// Выключение дефроста
private: System::Void buttonDefrostOff_Click(System::Object^ sender, System::EventArgs^ e) 
{
    bool success = SendControllerCommand_SetDefrost(false);
    
    if (success) {
        Label_Data->Text = "Дефрост выключен";
        labelDefrostStatus->BackColor = System::Drawing::Color::Gray;
    }
}
```

### Пример 5: Запрос статуса контроллера

```cpp
// Получение статуса контроллера
private: System::Void buttonGetStatus_Click(System::Object^ sender, System::EventArgs^ e) 
{
    ControllerResponse response;
    bool success = SendControllerCommand_GetStatus(response);
    
    if (success) {
        // Обрабатываем полученные данные статуса
        // Формат данных зависит от протокола контроллера
        
        if (response.dataLen > 0) {
            // Пример: предположим, что первый байт - это состояние программы
            uint8_t programState = response.data[0];
            
            System::String^ statusText = "Статус: ";
            switch (programState) {
                case 0: statusText += "Остановлена"; break;
                case 1: statusText += "Запущена"; break;
                case 2: statusText += "Приостановлена"; break;
                default: statusText += "Неизвестно"; break;
            }
            
            Label_Data->Text = statusText;
        }
    }
}
```

### Пример 6: Получение данных сенсоров

```cpp
// Получение данных сенсоров
private: System::Void buttonGetSensors_Click(System::Object^ sender, System::EventArgs^ e) 
{
    ControllerResponse response;
    bool success = SendControllerCommand_GetSensors(response);
    
    if (success && response.dataLen > 0) {
        // Пример обработки данных сенсоров
        // Предположим, что данные - это массив температур (float, 4 байта каждый)
        
        int numSensors = response.dataLen / sizeof(float);
        
        for (int i = 0; i < numSensors && i < 6; i++) {
            float temperature;
            memcpy(&temperature, &response.data[i * sizeof(float)], sizeof(float));
            
            // Обновляем соответствующие поля на форме
            switch (i) {
                case 0: T_def_left->Text = temperature.ToString("F1") + "°C"; break;
                case 1: T_def_right->Text = temperature.ToString("F1") + "°C"; break;
                case 2: T_def_center->Text = temperature.ToString("F1") + "°C"; break;
                case 3: T_product_left->Text = temperature.ToString("F1") + "°C"; break;
                case 4: T_product_right->Text = temperature.ToString("F1") + "°C"; break;
            }
        }
    }
}
```

### Пример 7: Получение версии ПО контроллера

```cpp
// Получение версии ПО
private: System::Void buttonGetVersion_Click(System::Object^ sender, System::EventArgs^ e) 
{
    ControllerResponse response;
    bool success = SendControllerCommand_GetVersion(response);
    
    if (success) {
        // Версия уже отображена в Label_Data внутри функции
        // Можно дополнительно обработать данные
        
        if (response.dataLen > 0) {
            System::String^ version = System::Text::Encoding::ASCII->GetString(
                response.data, 0, response.dataLen
            );
            
            // Показать в MessageBox
            MessageBox::Show("Версия ПО контроллера:\n" + version, 
                           "Информация о версии", 
                           MessageBoxButtons::OK, 
                           MessageBoxIcon::Information);
        }
    }
}
```

### Пример 8: Универсальная отправка команды

```cpp
// Отправка произвольной команды
private: System::Void SendCustomCommand() 
{
    // Создаем данные команды
    uint8_t data[4];
    data[0] = 0x01; // Параметр 1
    data[1] = 0x02; // Параметр 2
    data[2] = 0x03; // Параметр 3
    data[3] = 0x04; // Параметр 4
    
    ControllerResponse response;
    
    // Отправляем команду типа CONFIGURATION, код CMD_SET_PARAM
    bool success = SendControllerCommandWithResponse(
        CONFIGURATION,      // Тип команды
        CMD_SET_PARAM,      // Код команды
        data,               // Данные
        4,                  // Длина данных
        response            // Ответ
    );
    
    if (success) {
        Label_Data->Text = "Параметры установлены успешно";
        
        // Обработка ответа, если есть данные
        if (response.dataLen > 0) {
            // Обработка данных ответа...
        }
    } else {
        Label_Data->Text = "Ошибка установки параметров";
    }
}
```

### Пример 9: Периодический опрос контроллера

```cpp
// В классе DataForm добавить таймер для периодического опроса
private: System::Windows::Forms::Timer^ pollTimer;

// В конструкторе DataForm
DataForm(void)
{
    InitializeComponent();
    // ... другая инициализация ...
    
    // Инициализация таймера опроса
    pollTimer = gcnew System::Windows::Forms::Timer();
    pollTimer->Interval = 5000; // Опрос каждые 5 секунд
    pollTimer->Tick += gcnew EventHandler(this, &DataForm::PollController);
    pollTimer->Start();
}

// Обработчик таймера
private: System::Void PollController(Object^ sender, EventArgs^ e) 
{
    // Запрашиваем данные сенсоров
    ControllerResponse response;
    bool success = SendControllerCommand_GetSensors(response);
    
    if (success && response.dataLen > 0) {
        // Обновляем отображение данных на форме
        // ... обработка данных сенсоров ...
    } else {
        // Опрос не удался, можно показать предупреждение
        // но не блокируем работу приложения
    }
}
```

## Обработка ошибок

Все функции отправки команд возвращают `bool`:
- `true` - команда выполнена успешно
- `false` - произошла ошибка

При ошибке:
1. Сообщение об ошибке записывается в лог через `GlobalLogger::LogMessage()`
2. Сообщение об ошибке отображается в `Label_Data->Text`
3. Функция возвращает `false`

Пример обработки ошибок:

```cpp
bool success = SendControllerCommand_StartProgram();

if (!success) {
    // Проверяем возможные причины ошибки:
    
    // 1. Нет подключения к контроллеру
    if (clientSocket == INVALID_SOCKET) {
        MessageBox::Show("Нет подключения к контроллеру", "Ошибка");
        return;
    }
    
    // 2. Таймаут ответа
    MessageBox::Show("Контроллер не ответил вовремя.\n"
                    "Проверьте соединение.", "Таймаут");
    
    // 3. Ошибка на стороне контроллера
    // Текст ошибки уже отображен в Label_Data
}
```

## Логирование

Все операции с протоколом автоматически логируются:

```cpp
GlobalLogger::LogMessage("Command created: Type=0x01, Code=0x01, DataLen=0, CRC=0x1234");
GlobalLogger::LogMessage("Command sent successfully: 5 bytes");
GlobalLogger::LogMessage("Response received: Type=0x01, Code=0x01, Status=0x00, DataLen=0");
GlobalLogger::LogMessage("Command completed successfully: OK: Команда выполнена успешно");
```

Лог сохраняется в файл `log.txt` в директории приложения.

## Важные замечания

1. **Таймауты**: По умолчанию таймаут ожидания ответа составляет 5000 мс (5 секунд). Можно изменить при необходимости.

2. **CRC16**: Контрольная сумма вычисляется автоматически по алгоритму ModBus CRC16.

3. **Полнодуплексность**: Можно отправлять команды и получать данные одновременно без остановки прослушивания порта.

4. **Длина данных**: Максимальная длина данных в команде/ответе - 59 байт. Общая длина пакета - 64 байта.

5. **Безопасность потоков**: Все методы безопасны для вызова из GUI потока, так как работают с сокетом клиента формы.

## Добавление новых команд

Чтобы добавить новую команду:

1. Добавьте константу в `SServer.h`:
```cpp
enum MyNewCommandType : uint8_t {
    CMD_MY_NEW_COMMAND = 0x05
};
```

2. Добавьте метод в `DataForm.h`:
```cpp
bool SendControllerCommand_MyNewCommand(uint8_t param1, uint8_t param2);
```

3. Реализуйте метод в `DataForm.cpp`:
```cpp
bool ProjectServerW::DataForm::SendControllerCommand_MyNewCommand(uint8_t param1, uint8_t param2)
{
    uint8_t data[2];
    data[0] = param1;
    data[1] = param2;
    
    ControllerResponse response;
    bool result = SendControllerCommandWithResponse(
        PROG_CONTROL,           // Тип команды
        CMD_MY_NEW_COMMAND,     // Код команды
        data, 2,                // Данные
        response                // Ответ
    );
    
    if (result) {
        Label_Data->Text = "Моя команда выполнена";
    }
    
    return result;
}
```

## Заключение

Протокол команд полностью реализован и готов к использованию. Все примеры выше можно адаптировать под ваши конкретные задачи. Протокол поддерживает полнодуплексную связь, автоматическую проверку CRC, обработку ошибок и логирование всех операций.

