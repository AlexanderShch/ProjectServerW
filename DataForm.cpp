#include "DataForm.h"
#include "Chart.h"
#include "Commands.h"               // Для работы с командами управления
#include <objbase.h>                // Для CoCreateGuid - генерация уникального идентификатора
#include <string>
#include <vcclr.h>  // Для gcnew

// Добавьте эту строку для доступа к Process
using namespace System::Diagnostics;
using namespace ProjectServerW; // Добавлено пространство имен
using namespace Microsoft::Office::Interop::Excel;

std::map<std::wstring, gcroot<DataForm^>> formData_Map; // Определение переменной formData_Map

// Структура телеметрии от контроллера (упакованная, без выравнивания байтов)
// Контроллер использует __attribute__((packed)), здесь используем #pragma pack(1)
#pragma pack(push, 1)
typedef struct   // object data for Server type из STM32
{
    uint8_t DataType;           // Байт типа передаваемых данных (0х00 для телеметрии)
    uint16_t Time;				// Количество секунд с момента включения
    uint8_t SensorQuantity;		// Количество сенсоров
    uint8_t SensorType[SQ];		// Тип сенсора
    uint8_t Active[SQ];			// Активность сенсора
    short T[SQ];				// Значение 1 сенсора (температура)
    short H[SQ];				// Значение 2 сенсора (влажность)
    uint16_t CRC_SUM;			// Контрольное значение
} MSGQUEUE_OBJ_t;
#pragma pack(pop)

// Преобразование данных из буфера в структуру типа MSGQUEUE_OBJ_t из STM32
void ProjectServerW::DataForm::ParseBuffer(const char* buffer, size_t size) {

    MSGQUEUE_OBJ_t data;

    if (size < sizeof(data)) {
        return; // Принятых данных слишком мало
    }
    memcpy(&data, buffer, sizeof(data));
}

System::Void ProjectServerW::DataForm::выходToolStripMenuItem_Click(System::Object^ sender, System::EventArgs^ e)
{
    // Найдем GUID текущей формы в карте formData_Map
    std::wstring currentFormGuid;
    bool formFound = false;

    for (auto it = formData_Map.begin(); it != formData_Map.end(); ++it) {
        // Извлекаем управляемый указатель из gcroot
        ProjectServerW::DataForm^ formPtr = it->second;

        // Теперь сравниваем указатели
        if (formPtr == this) {
            currentFormGuid = it->first;
            formFound = true;
            break;
        }
    }

    try {
        if (formFound) {
            // Закрываем соединение с клиентом
            // Находим форму по GUID
            DataForm^ form = GetFormByGuid(currentFormGuid);

            if (form != nullptr && form->ClientSocket != INVALID_SOCKET) {
                // Закрываем сокет клиента
                closesocket(form->ClientSocket);
                form->ClientSocket = INVALID_SOCKET;
                GlobalLogger::LogMessage("Information: Закрыто соединение с клиентом по кнопке Выход");
            }
        }
    }
    catch (Exception^ ex) {
        MessageBox::Show("Ошибка при закрытии сокета: " + ex->Message);
        GlobalLogger::LogMessage(ConvertToStdString("Ошибка при закрытии сокета: " + ex->Message));
    }

	return System::Void();
}

System::Void ProjectServerW::DataForm::buttonEXCEL_Click(System::Object^ sender, System::EventArgs^ e)
{
    // Отключаем кнопку на время обработки
    buttonExcel->Enabled = false;
    // Создаем и запускаем поток
    excelThread = gcnew Thread(gcnew ThreadStart(this, &DataForm::AddDataToExcel));
    excelThread->SetApartmentState(ApartmentState::STA);  // Обязательно для Excel!
    // Важно! Установка как основного, не фонового потока. 
    // Даем потоку завершиться правильно, блокирует завершение программы до своего завершения
    excelThread->IsBackground = false; 
    excelThread->Start();
    GlobalLogger::LogMessage("Information: Запись в EXCEL по нажатию на кнопку");
}

System::Void ProjectServerW::DataForm::buttonBrowse_Click(System::Object^ sender, System::EventArgs^ e)
{
    // Создаем диалог выбора папки
    FolderBrowserDialog^ folderDialog = gcnew FolderBrowserDialog();

    // Настройка диалога
    folderDialog->Description = "Выберите папку для сохранения Excel файлов";
    folderDialog->ShowNewFolderButton = true;

    // Устанавливаем начальную директорию из текстового поля
    if (textBoxExcelDirectory->Text->Length > 0) {
        folderDialog->SelectedPath = textBoxExcelDirectory->Text;
    }

    // Показываем диалог и проверяем результат
    if (folderDialog->ShowDialog() == System::Windows::Forms::DialogResult::OK) {
        // Обновляем текстовое поле выбранным путем
        textBoxExcelDirectory->Text = folderDialog->SelectedPath;

        // Сохраняем выбранный путь
        excelSavePath = folderDialog->SelectedPath;
        // Сохранение настроек
        SaveSettings();

    }
}
       
void ProjectServerW::DataForm::CreateAndShowDataFormInThread(std::queue<std::wstring>& messageQueue,
                                                             std::mutex& mtx, 
                                                             std::condition_variable& cv) {
    DataForm^ form = gcnew DataForm();
    GUID guid;
    HRESULT result = CoCreateGuid(&guid);
    if (result == S_OK) {
        // Преобразование GUID в строку
        wchar_t guidString[40] = { 0 };
        int simb_N = StringFromGUID2(guid, guidString, 40);

        String^ formId = gcnew String(guidString);

        // Сохранение формы в карте
        formData_Map[guidString] = form;

        // Установка значения в очередь сообщений
        {
            std::lock_guard<std::mutex> lock(mtx);
            messageQueue.push(guidString);
        }
        cv.notify_one();
    }
    form->ShowDialog();

 }

// Метод закрытия формы, применяется при закрытии сокета
void ProjectServerW::DataForm::CloseForm(const std::wstring& guid) {
    // Находим форму
    ProjectServerW::DataForm^ form = ProjectServerW::DataForm::GetFormByGuid(guid);
    //MessageBox::Show("DataForm will be closed!");
    GlobalLogger::LogMessage("DataForm will be closed!");

    if (form != nullptr) {
        // ВАЖНО: Удаляем форму из карты СРАЗУ, до закрытия
        // Это позволяет новым соединениям от того же клиента создавать новые формы немедленно
        auto it = formData_Map.find(guid);
        if (it != formData_Map.end()) {
            formData_Map.erase(it);
            GlobalLogger::LogMessage("Information: Форма удалена из карты, разрешено новое соединение");
        }
        
        // Проверяем, нужен ли Invoke
        if (form->InvokeRequired) {
            // Закрываем форму через Invoke
            form->Invoke(gcnew System::Action(form, &ProjectServerW::DataForm::Close));
        }
        else {
            form->Close();
        }
    }
}
        
// Метод для получения формы по её GUID
DataForm^ ProjectServerW::DataForm::GetFormByGuid(const std::wstring& guid) {
    auto it = formData_Map.find(guid);
    if (it != formData_Map.end()) {
        DataForm^ form = it->second;

        // Проверяем, что форма существует и не закрыта
        if (form != nullptr && !form->IsDisposed && !form->Disposing) {
            return form;
        }
        else {
            // Форма закрыта или недействительна, удаляем её из карты
            formData_Map.erase(it);
            return nullptr;
        }

    }
    return nullptr;
}

// Метод для поиска формы по IP-адресу клиента
std::wstring ProjectServerW::DataForm::FindFormByClientIP(String^ clientIP) {
    // Проходим по всем формам в карте
    for (auto it = formData_Map.begin(); it != formData_Map.end(); ++it) {
        DataForm^ form = it->second;
        
        // Проверяем, что форма существует и не закрыта
        if (form != nullptr && !form->IsDisposed && !form->Disposing) {
            // Сравниваем IP-адреса
            if (form->ClientIP != nullptr && form->ClientIP->Equals(clientIP)) {
                // Нашли форму с таким же IP
                return it->first; // Возвращаем GUID
            }
        }
    }
    // Не нашли форму с таким IP
    return std::wstring();
}

// Записываем пару guid и поток в map
void ThreadStorage::StoreThread(const std::wstring& guid, std::thread& thread) {
    std::lock_guard<std::mutex> lock(GetMutex());
    GetThreadMap()[guid] = std::move(thread);
}

// Останавливаем поток по guid
void ThreadStorage::StopThread(const std::wstring& guid)
{
    std::lock_guard<std::mutex> lock(GetMutex());

    // Ищем поток
    auto it = GetThreadMap().find(guid);
    if (it != GetThreadMap().end() && it->second.joinable()) {
        it->second.join();  // Ждем завершения
        GetThreadMap().erase(it);  // Удаляем из карты
    }
}

// Функция для определения статической переменной map для потока
std::map<std::wstring, std::thread>& ThreadStorage::GetThreadMap() {
    static std::map<std::wstring, std::thread> threadMap;
    return threadMap;
}

// Функция для определения статической переменной Mutex для потока
std::mutex& ThreadStorage::GetMutex() {
    static std::mutex mtx;
    return mtx;
}

// Эта функция принимает буфер и его длину, затем преобразует каждый байт буфера в шестнадцатеричный формат и добавляет его в строку.
static String^ bufferToHex(const char* buffer, size_t length) {
    // Использование std::stringstream позволяет преобразовывать данные в строку и форматировать их.
    std::stringstream ss;
    ss << std::hex << std::setfill('0');
    for (int i = 0; i < length; ++i) {
        ss << std::setw(2) << static_cast<unsigned int>(static_cast<unsigned char>(buffer[i])) << " ";
    }
    return gcnew String(ss.str().c_str());
}

/*
*********** Создаём таблицу *******************
*/

// 1. Создаём таблицу данных
void ProjectServerW::DataForm::InitializeDataTable() {
    // Создаем таблицу данных
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

    // Обработка сигналов с устройства ввода-вывода
    dataTable->Columns->Add("Typ" + (SQ - 1), uint8_t::typeid);
    dataTable->Columns->Add("Act" + (SQ - 1), uint8_t::typeid);

    cli::array<cli::array<String^>^>^ bitNames = GetBitFieldNames();
    if (bitNames[0] != nullptr) {
        // Работа с массивом имен сигналов состояния дефростера
        for (uint8_t i = 0; i < 16; i++)
        {
            dataTable->Columns->Add(bitNames[0][i], uint8_t::typeid);
        }
        // Работа с массивом имен сигналов управления дефростером
        for (uint8_t i = 0; i < 16; i++)
        {
            dataTable->Columns->Add(bitNames[1][i], uint8_t::typeid);
        }
    }

    dataGridView->DataSource = dataTable;
    GlobalLogger::LogMessage("Information: Таблица данных создана");
}

// 2. Добавление данных
void ProjectServerW::DataForm::AddDataToTable(const char* buffer, size_t size, System::Data::DataTable^ table) {
    
    MSGQUEUE_OBJ_t data;
    DateTime now = DateTime::Now;   // Получение текущего времени
    String^ data_String{};

    if (size < sizeof(data)) {
        return; // Принятых данных слишком мало
    }
    memcpy(&data, buffer, sizeof(data));

    data_String = bufferToHex(buffer, size);
    // Выводим значение строки data_String в элемент Label_Data
    DataForm::SetData_TextValue("Принятые данные: " + data_String);

    // Создание новой строки
    DataRow^ row = table->NewRow();
    row["RealTime"] = now.ToString("HH:mm:ss");
    row["Time"] = data.Time;
    row["SQ"] = data.SensorQuantity;
    for (uint8_t i = 0; i < (SQ - 1); i++)
    {
        row["Typ" + i] = data.SensorType[i];
        row["Act" + i] = data.Active[i];
        row["T" + i] = data.T[i] / 10.0;
        row["H" + i] = data.H[i] / 10.0;
    }
    // Обновляем значения температуры в элементах интерфейса
    cli::array<double>^ temperatures = gcnew cli::array<double>(5);
    temperatures[0] = data.T[0] / 10.0;  // T_def_left
    temperatures[1] = data.T[1] / 10.0;  // T_def_right
    temperatures[2] = data.T[2] / 10.0;  // T_def_center
    temperatures[3] = data.T[3] / 10.0;  // T_product_left
    temperatures[4] = data.T[4] / 10.0;  // T_product_right
    UpdateAllTemperatureValues(temperatures);

    cli::array<cli::array<String^>^>^ bitNames = GetBitFieldNames();

    // Находим индекс бита "Work" в массиве bitNames[0]
    int workBitIndex = -1;
    // Найдем индекс бита "Work" в массиве имен
    for (int i = 0; i < bitNames[0]->Length; i++) {
        if (bitNames[0][i] == "Work") {
            workBitIndex = i;
            break;
        }
    }

    uint16_t bitField;
    bitField = data.T[SQ - 1];
    if (workBitIndex != -1) {
        // Проверяем текущее состояние бита "Work"
        bool currentWorkBitState = (bitField & (1 << workBitIndex)) != 0;

        // ===== УСТАНОВКА НАЧАЛЬНОГО СОСТОЯНИЯ КНОПОК ПРИ ПЕРВОМ ПРИЁМЕ ДАННЫХ =====
        if (!firstDataReceived) {
            // Это первый пакет данных - устанавливаем начальное состояние кнопок
            if (currentWorkBitState) {
                // Бит WORK активен - работа уже идёт
                buttonSTOPstate_TRUE();
                GlobalLogger::LogMessage("Information: Первый пакет данных - бит WORK активен, кнопка STOP активна");
            } else {
                // Бит WORK не активен - можно запустить
                buttonSTARTstate_TRUE();
                GlobalLogger::LogMessage("Information: Первый пакет данных - бит WORK не активен, кнопка START активна");
            }
            firstDataReceived = true;
        }

        // Если бит "Work" переходит из 0 в 1 (дефростер is ON)
        if (!workBitDetected && currentWorkBitState) {
            // Проверяем: это возврат после мигания или новый запуск?
            if (workBitZeroTimerActive) {
                // Это мигание лампы - бит вернулся в 1 во время отслеживания
                // НЕ логируем, только сбрасываем таймер отслеживания
                workBitZeroTimerActive = false;
                // НЕ сбрасываем workBitZeroLogged - чтобы не логировать повторно при следующем мигании
                if (delayedExcelTimer != nullptr && delayedExcelTimer->Enabled) {
                    delayedExcelTimer->Stop();
                }
            } else {
                // Это новый запуск работы (не мигание)
                // Сохраняем время начала сбора данных
                dataCollectionStartTime = now;
                // Создаем имя файла на основе времени начала (окончание добавится позже при записи в Excel)
                excelFileName = "WorkData_" + now.ToString("yyyy-MM-dd_HH-mm-ss") + "_Port" + clientPort.ToString();
                GlobalLogger::LogMessage(ConvertToStdString("Information: СТАРТ фиксации данных, создано имя файла " + excelFileName));
                // Сбрасываем флаги при новом запуске работы
                workBitZeroLogged = false;
                dataExportedToExcel = false;  // Новый цикл - данные еще не экспортированы
            }
            // Меняем состояние кнопок СТАРТ и СТОП
            buttonSTOPstate_TRUE();
        }

        // Если бит "Work" переходит из 1 в 0 (дефростер is OFF)
        if (workBitDetected && !currentWorkBitState) {
            // Начинаем отслеживать время, когда бит стал нулём
            workBitZeroStartTime = now;
            workBitZeroTimerActive = true;
            
            // Создаем и запускаем таймер для проверки состояния каждую секунду
            if (delayedExcelTimer == nullptr) {
                delayedExcelTimer = gcnew System::Windows::Forms::Timer();
                delayedExcelTimer->Interval = 1000; // Проверяем каждую секунду
                delayedExcelTimer->Tick += gcnew EventHandler(this, &DataForm::OnDelayedExcelTimerTick);
            }
            delayedExcelTimer->Start();
            
            // Логируем только первый переход в ноль
            if (!workBitZeroLogged) {
                GlobalLogger::LogMessage(ConvertToStdString("Information: Бит Work стал нулём, начато отслеживание времени..."));
                workBitZeroLogged = true;
            }
        }

        // Если бит "Work" остаётся в нуле и таймер активен, проверяем, прошла ли минута
        if (!currentWorkBitState && workBitZeroTimerActive) {
            TimeSpan elapsed = now.Subtract(workBitZeroStartTime);
            if (elapsed.TotalSeconds >= 60) {
                // Прошло не менее 1 минуты непрерывного нахождения в нуле
                workBitZeroTimerActive = false;
                workBitZeroLogged = false;  // Сбрасываем флаг логирования
                delayedExcelTimer->Stop();
                
                // Сохраняем время окончания сбора данных
                dataCollectionEndTime = now;
                
                GlobalLogger::LogMessage(ConvertToStdString("Information: СТОП фиксации данных, запись в файл " + excelFileName));
                // Меняем состояние кнопок СТАРТ и СТОП
                buttonSTARTstate_TRUE();
                
                // Устанавливаем флаг, что данные экспортированы (чтобы не дублировать при закрытии формы)
                dataExportedToExcel = true;
                
                // Выполняем запись в Excel
                this->BeginInvoke(gcnew MethodInvoker(this, &DataForm::TriggerExcelExport));
            }
        }

        // Обновляем состояние флага
        workBitDetected = currentWorkBitState;
    }

    // Извлекаем каждый бит и добавляем в таблицу сигналов состояния дефростера
    for (int bit = 0; bit < 16; bit++) {
        bool bitValue = (bitField & (1 << bit)) != 0;
        row[bitNames[0][bit]] = bitValue;
    }
    bitField = data.H[SQ - 1];
    // Извлекаем каждый бит и добавляем в таблицу сигналов управления дефростером
    for (int bit = 0; bit < 16; bit++) {
        bool bitValue = (bitField & (1 << bit)) != 0;
        row[bitNames[1][bit]] = bitValue;
    }

    // Добавление строки в таблицу только во время работы дефростера
    // ВРЕМЕННО ОТКЛЮЧЕНО ДЛЯ ОТЛАДКИ - данные записываются всегда
    if (workBitDetected)
    {
        table->Rows->Add(row);
    }
}


// 3. Сохраняем таблицу в EXCEL
void ProjectServerW::DataForm::AddDataToExcel() {
    // Создаём объект Excel в STA-потоке
    ExcelHelper^ excel = nullptr;
    excel = gcnew ExcelHelper();

    try {
        // Замеряем время выполнения для мониторинга производительности
        DateTime startTime = DateTime::Now;
        GlobalLogger::LogMessage("Information: Начало записи в Excel...");

        // DataTable не является потокобезопасной структурой, в Excel будем перегонять копию
        System::Data::DataTable^ copiedTable = dataTable->Copy();
        // Последняя строка может быть не заполнена полностью, её удалим
        if (copiedTable->Rows->Count > 0) {
            copiedTable->Rows->RemoveAt(copiedTable->Rows->Count - 1);
        }

        GlobalLogger::LogMessage(ConvertToStdString(String::Format(
            "Information: Количество строк для записи: {0}, столбцов: {1}", 
            copiedTable->Rows->Count, copiedTable->Columns->Count)));

        if (excel->CreateNewWorkbook()) {
            Microsoft::Office::Interop::Excel::Worksheet^ ws = excel->GetWorksheet();

            // ОПТИМИЗАЦИЯ: отключаем обновление экрана и автоматические вычисления
            Microsoft::Office::Interop::Excel::Application^ excelApp = safe_cast<Microsoft::Office::Interop::Excel::Application^>(ws->Application);
            
            try {
                excelApp->ScreenUpdating = false;
                excelApp->Calculation = Microsoft::Office::Interop::Excel::XlCalculation::xlCalculationManual;
                excelApp->EnableEvents = false;
                GlobalLogger::LogMessage("Information: Excel оптимизирован для быстрой записи");

            // Заголовки - ОПТИМИЗИРОВАННАЯ ВЕРСИЯ: записываем все заголовки сразу через массив
            cli::array<cli::array<String^>^>^ bitNames = GetBitFieldNames();
            
            // Определяем количество столбцов
            int colCount = 6 + 4 * (SQ - 1); // базовые столбцы
            if (bitNames[0] != nullptr) colCount += 16;
            if (bitNames[1] != nullptr) colCount += 16;

            // Создаем одномерный массив для заголовков
            cli::array<System::Object^>^ headerArray = gcnew cli::array<System::Object^>(colCount);
            
            int col = 0;
            headerArray[col++] = "RealTime";
            headerArray[col++] = "Time";
            headerArray[col++] = "SQ";
            
            for (uint8_t i = 0; i < (SQ - 1); i++)
            {
                headerArray[col++] = "Typ" + i;
                headerArray[col++] = "Act" + i;
                // Добавляем имя датчика в заголовок T-столбца, если оно задано
                if (sensorNames != nullptr && i < sensorNames->Length && sensorNames[i] != nullptr && sensorNames[i]->Length > 0) {
                    headerArray[col++] = "T" + i + " " + sensorNames[i];
                    headerArray[col++] = "H" + i + " " + sensorNames[i];
                }
                else {
                    headerArray[col++] = "T" + i;
                    headerArray[col++] = "H" + i;
                }
            }
            
            // Обработка сигналов с устройства ввода-вывода
            headerArray[col++] = "Typ" + (SQ - 1);
            headerArray[col++] = "Act" + (SQ - 1);

            // Работа с массивом имен сигналов состояния дефростера
            if (bitNames[0] != nullptr) {
                for (uint8_t i = 0; i < 16; i++)
                {
                    headerArray[col++] = bitNames[0][i];
                }
            }
            // Работа с массивом имен сигналов управления дефростером
            if (bitNames[1] != nullptr) {
                for (uint8_t i = 0; i < 16; i++)
                {
                    headerArray[col++] = bitNames[1][i];
                }
            }

            // КЛЮЧЕВАЯ ОПТИМИЗАЦИЯ: записываем весь ряд заголовков одной операцией
            Microsoft::Office::Interop::Excel::Range^ headerRange = ws->Range[ws->Cells[1, 1], ws->Cells[1, colCount]];
            headerRange->Value2 = headerArray;
            Marshal::ReleaseComObject(headerRange);

            // Данные - ОПТИМИЗИРОВАННАЯ ВЕРСИЯ: записываем все данные сразу через массив
            int rowCount = copiedTable->Rows->Count;
            if (rowCount > 0) {
                // Определяем количество столбцов
                int colCount = 6 + 4 * (SQ - 1); // базовые столбцы
                if (bitNames[0] != nullptr) colCount += 16;
                if (bitNames[1] != nullptr) colCount += 16;

                // Создаем двумерный массив для всех данных
                cli::array<System::Object^, 2>^ dataArray = gcnew cli::array<System::Object^, 2>(rowCount, colCount);

                // Заполняем массив данными из таблицы
                int arrayRow = 0;
                for each (DataRow ^ dr in copiedTable->Rows)
                {
                    int col = 0;
                    dataArray[arrayRow, col++] = dr["RealTime"]->ToString();
                    dataArray[arrayRow, col++] = Convert::ToInt32(dr["Time"]);
                    dataArray[arrayRow, col++] = Convert::ToUInt16(dr["SQ"]);
                    
                    for (uint8_t i = 0; i < (SQ - 1); i++)
                    {
                        dataArray[arrayRow, col++] = Convert::ToUInt16(dr["Typ" + i]);
                        dataArray[arrayRow, col++] = Convert::ToUInt16(dr["Act" + i]);
                        dataArray[arrayRow, col++] = dr["T" + i];
                        dataArray[arrayRow, col++] = dr["H" + i];
                    }

                    // Работа с массивом имен сигналов состояния дефростера
                    if (bitNames[0] != nullptr) {
                        for (uint8_t i = 0; i < 16; i++)
                        {
                            dataArray[arrayRow, col++] = Convert::ToUInt16(dr[bitNames[0][i]]);
                        }
                    }
                    // Работа с массивом имен сигналов управления дефростером
                    if (bitNames[1] != nullptr) {
                        for (uint8_t i = 0; i < 16; i++)
                        {
                            dataArray[arrayRow, col++] = Convert::ToUInt16(dr[bitNames[1][i]]);
                        }
                    }

                    arrayRow++;
                }

                // КЛЮЧЕВАЯ ОПТИМИЗАЦИЯ: записываем весь массив одной операцией
                Microsoft::Office::Interop::Excel::Range^ startCell = safe_cast<Microsoft::Office::Interop::Excel::Range^>(ws->Cells[2, 1]);
                Microsoft::Office::Interop::Excel::Range^ endCell = safe_cast<Microsoft::Office::Interop::Excel::Range^>(ws->Cells[rowCount + 1, colCount]);
                Microsoft::Office::Interop::Excel::Range^ dataRange = ws->Range[startCell, endCell];
                dataRange->Value2 = dataArray;

                // Освобождаем временные объекты
                Marshal::ReleaseComObject(startCell);
                Marshal::ReleaseComObject(endCell);
                Marshal::ReleaseComObject(dataRange);
            }
            int row = rowCount + 2; // Для дальнейшего использования в графике
            // Добавление листа с графиком температур
            try {
                int lastRow = row - 1;
                if (lastRow >= 2) {
                    // Получаем рабочую книгу из рабочего листа
                    Microsoft::Office::Interop::Excel::Workbook^ wb = safe_cast<Microsoft::Office::Interop::Excel::Workbook^>(ws->Parent);
                    System::Object^ missing = System::Type::Missing;

                    // Создаём новый лист для диаграммы
                    Microsoft::Office::Interop::Excel::Worksheet^ chartSheet = safe_cast<Microsoft::Office::Interop::Excel::Worksheet^>(wb->Worksheets->Add(missing, ws, 1, Microsoft::Office::Interop::Excel::XlSheetType::xlWorksheet));
                    chartSheet->Name = "Chart";

                    // Размещаем объект диаграммы
                    Microsoft::Office::Interop::Excel::ChartObjects^ chartObjects = safe_cast<Microsoft::Office::Interop::Excel::ChartObjects^>(chartSheet->ChartObjects(missing));
                    Microsoft::Office::Interop::Excel::ChartObject^ chartObject = chartObjects->Add(20, 20, 900, 500);
                    Microsoft::Office::Interop::Excel::Chart^ chart = chartObject->Chart;
                    chart->ChartType = Microsoft::Office::Interop::Excel::XlChartType::xlLine;

                    // Диапазон X (время) из первого листа: колонка 1 (RealTime)
                    Microsoft::Office::Interop::Excel::Range^ xRange = ws->Range[ws->Cells[2, 1], ws->Cells[lastRow, 1]];

                    // Добавляем серии T0..T(SQ-2)
                    Microsoft::Office::Interop::Excel::SeriesCollection^ seriesCollection = safe_cast<Microsoft::Office::Interop::Excel::SeriesCollection^>(chart->SeriesCollection(missing));
                    for (uint8_t i = 0; i < (SQ - 1); i++)
                    {
                        int seriesCol = 6 + 4 * i; // Колонка T{i}
                        Microsoft::Office::Interop::Excel::Range^ yRange = ws->Range[ws->Cells[2, seriesCol], ws->Cells[lastRow, seriesCol]];
                        Microsoft::Office::Interop::Excel::Series^ s = seriesCollection->NewSeries();
                        s->XValues = xRange;
                        s->Values = yRange;
                        if (sensorNames != nullptr && i < sensorNames->Length && sensorNames[i] != nullptr && sensorNames[i]->Length > 0) {
                            s->Name = "T" + Convert::ToString(i) + " " + sensorNames[i];
                        }
                        else {
                            s->Name = "T" + Convert::ToString(i);
                        }
                        // Применяем цвет к серии, если задан
                        if (sensorColors != nullptr && i < sensorColors->Length) {
                            System::Drawing::Color c = sensorColors[i];
                            int oleColor = System::Drawing::ColorTranslator::ToOle(c);
                            // Color у линии в Chart.Series.Format.Line.ForeColor / или Border.Color
                            try {
                                s->Border->Color = oleColor;
                            }
                            catch (...) {}
                            try {
                                s->Format->Line->ForeColor->RGB = oleColor;
                            }
                            catch (...) {}
                        }
                    }

                    // Заголовки осей
                    chart->HasTitle = true;
                    chart->ChartTitle->Text = "Температуры";
                    Microsoft::Office::Interop::Excel::Axis^ categoryAxis = safe_cast<Microsoft::Office::Interop::Excel::Axis^>(
                        chart->Axes(Microsoft::Office::Interop::Excel::XlAxisType::xlCategory, Microsoft::Office::Interop::Excel::XlAxisGroup::xlPrimary));
                    categoryAxis->HasTitle = true;
                    categoryAxis->AxisTitle->Text = "RealTime";
                    Microsoft::Office::Interop::Excel::Axis^ valueAxis = safe_cast<Microsoft::Office::Interop::Excel::Axis^>(
                        chart->Axes(Microsoft::Office::Interop::Excel::XlAxisType::xlValue, Microsoft::Office::Interop::Excel::XlAxisGroup::xlPrimary));
                    valueAxis->HasTitle = true;
                    valueAxis->AxisTitle->Text = "T, °C";
                }
            }
            catch (Exception^) {
                // Пропускаем создание графика при ошибке, чтобы не сорвать сохранение файла
            }

            } // конец try блока для работы с Excel
            finally {
                // ОПТИМИЗАЦИЯ: ВСЕГДА восстанавливаем настройки Excel перед сохранением
                try {
                    if (excelApp != nullptr) {
                        excelApp->Calculation = Microsoft::Office::Interop::Excel::XlCalculation::xlCalculationAutomatic;
                        excelApp->ScreenUpdating = true;
                        excelApp->EnableEvents = true;
                        Marshal::ReleaseComObject(excelApp);
                        GlobalLogger::LogMessage("Information: Настройки Excel восстановлены");
                    }
                }
                catch (...) {
                    GlobalLogger::LogMessage("Warning: Ошибка при восстановлении настроек Excel");
                }
            }

            // Сохранение файла
            // Формируем полный путь к файлу
            String^ filePath = excelSavePath;
            // Проверяем, существует ли указанный путь
            if (String::IsNullOrEmpty(filePath) || !System::IO::Directory::Exists(filePath)) {
                // Если путь не существует или пустой, создаем директорию в папке запуска программы
                String^ appPath = System::IO::Path::GetDirectoryName(System::Windows::Forms::Application::ExecutablePath);
                filePath = System::IO::Path::Combine(appPath, "SensorData");

                // Создаем директорию, если она не существует
                if (!System::IO::Directory::Exists(filePath)) {
                    System::IO::Directory::CreateDirectory(filePath);
                }

            // Обновляем переменную excelSavePath для будущих сохранений
            excelSavePath = filePath;

            // Обновляем текстовое поле, если оно существует и форма не закрыта
            if (textBoxExcelDirectory != nullptr && !textBoxExcelDirectory->IsDisposed && 
                !this->IsDisposed && this->IsHandleCreated) {
                try {
                    if (textBoxExcelDirectory->InvokeRequired) {
                        textBoxExcelDirectory->BeginInvoke(gcnew System::Action<String^>(this, &DataForm::UpdateDirectoryTextBox), filePath);
                    }
                    else {
                        textBoxExcelDirectory->Text = filePath;
                    }
                }
                catch (...) {
                    // Форма закрыта во время обновления - игнорируем
                }
            }
            // Сохраняем в настройках
            try {
                DataForm::SaveSettings();
            }
            catch (...) {
                // Игнорируем ошибки сохранения настроек при закрытии формы
            }
            }

            // Добавляем разделитель в конец пути, если его нет
            if (!filePath->EndsWith("\\")) {
                filePath += "\\";
            }

            // Добавляем имя файла с текущей датой и временем
            if (excelFileName != nullptr) {
                // Добавляем время окончания к имени файла
                String^ finalFileName = excelFileName;
                
                // Если есть время окончания, добавляем его к имени файла
                if (dataCollectionEndTime != DateTime::MinValue) {
                    finalFileName += "_End_" + dataCollectionEndTime.ToString("HH-mm-ss");
                }
                
                // Добавляем расширение .xlsx
                finalFileName += ".xlsx";
                
                // Используем имя файла с временем начала и окончания
                filePath += finalFileName;
            }
            else {
                // Используем стандартное имя с текущей датой и временем
                DateTime now = DateTime::Now;
                filePath += "SensorData_" + now.ToString("yyyy-MM-dd_HH-mm-ss") + "_Port" + clientPort.ToString() + ".xlsx";
            }

            // Сохранение по выбранному пути
            excel->SaveAs(filePath);
            
            // Вычисляем время выполнения
            DateTime endTime = DateTime::Now;
            TimeSpan elapsed = endTime.Subtract(startTime);
            GlobalLogger::LogMessage(ConvertToStdString(String::Format(
                "Information: Файл Excel успешно сохранен: {0}\nВремя записи: {1} секунд ({2} строк)", 
                excelFileName, elapsed.TotalSeconds.ToString("F2"), copiedTable->Rows->Count)));

            // Освободим память от copiedTable
            delete copiedTable;
            copiedTable = nullptr;
        }
    }
    catch (System::Runtime::InteropServices::COMException^ comEx) {
        // Специальная обработка COM-исключений
        String^ errorMsg = "COM error: " + comEx->Message + " (Code: " +
            comEx->ErrorCode.ToString("X8") + ")";
        //MessageBox::Show(errorMsg);
        GlobalLogger::LogMessage(ConvertToStdString(errorMsg));
    }
    catch (Exception^ ex) {
        // В случае ошибки
        String^ errorMsg = "Excel error: " + ex->Message;
        //MessageBox::Show(errorMsg);
        GlobalLogger::LogMessage(ConvertToStdString("Excel error: Не могу создать файл..." + ex->Message));
        // Необходимо освободить ресурсы даже при ошибке
        try {
            if (excel != nullptr) {
                excel->Close();
                delete excel;
                excel = nullptr;
            }
        }
        catch (...) { /* Игнорируем ошибки при очистке */ }
    }
    finally {
        try {
            // Важно! Сначала включить кнопку - до освобождения COM-объектов
            // Проверяем, не закрывается ли форма и существует ли она
    // Проверяем состояние формы
            bool canInvoke = false;
            try {
                canInvoke = !this->IsDisposed && this->IsHandleCreated && this->Visible;
            }
            catch (...) {
                canInvoke = false;
            }

            // Включаем кнопку только если можно безопасно вызвать Invoke
            if (canInvoke) {
                try {
                    this->BeginInvoke(gcnew MethodInvoker(this, &DataForm::EnableButton));
                }
                catch (...) {
                    // Игнорируем любые исключения при вызове BeginInvoke
                }
            }
            else {
                // Если форма закрывается, просто пропускаем включение кнопки
            }

            // Только после этого освобождаем Excel
            if (excel != nullptr) {
                excel->Close();
                delete excel;
                excel = nullptr;

                // Затем небольшая пауза для обработки
                Thread::Sleep(50);
                GlobalLogger::LogMessage("Information: Закрываю EXCEL...");
            }

            // Запускаем отложенную сборку мусора
            ThreadPool::QueueUserWorkItem(gcnew WaitCallback(DataForm::DelayedGarbageCollection));
            // Сигнализируем о завершении экспорта
            exportSuccessful = true;
            if (exportCompletedEvent != nullptr) {
                exportCompletedEvent->Set();
            }
        }
        catch (...) {}
    }    

}

// Отложенная сборка мусора
void DataForm::DelayedGarbageCollection(Object^ state) {
    try {
        // Пауза перед сборкой мусора
        //MessageBox::Show("Excel file was recorded successfully!");
        GlobalLogger::LogMessage(ConvertToStdString("Excel file was recorded successfully!"));
        Thread::Sleep(500);

        // Сборка мусора в отдельном потоке
        GC::Collect();
        GC::WaitForPendingFinalizers();
    }
    catch (...) {}
}

void DataForm::EnableButton()
{
    buttonExcel->Enabled = true;
}

void ProjectServerW::DataForm::AddDataToTableThreadSafe(cli::array<System::Byte>^ buffer, int size, int port) {
    // Этот метод выполняется в потоке формы благодаря Invoke
    // ВАЖНО! Методы, работающие с UI, должны выполняться в потоке, создавшем объект с UI!
    // Сохраняем порт клиента в форму DataForm
    this->clientPort = port;

    // Преобразуем управляемый массив байтов в неуправляемый буфер
    pin_ptr<Byte> pinnedBuffer = &buffer[0];
    char* rawBuffer = reinterpret_cast<char*>(pinnedBuffer);

    // ===== ОПТИМИЗАЦИЯ ПРОИЗВОДИТЕЛЬНОСТИ =====
    // Приостанавливаем обновление DataGridView для ускорения
    dataGridView->SuspendLayout();
    
    try {
        // Вызываем стандартный метод добавления данных
        AddDataToTable(rawBuffer, size, dataTable);
    }
    finally {
        // Возобновляем обновление DataGridView
        dataGridView->ResumeLayout(false);
    }

    // УБРАНО: Автоматическая прокрутка к последней строке
    // Причина: мешает пользователю просматривать данные в середине таблицы
    // Пользователь может самостоятельно прокрутить к последней строке при необходимости
    
    // УБРАНО: this->Refresh() - полное обновление формы
    // Причина: DataGridView обновляется автоматически при изменении DataSource
    // Полное обновление всей формы замедляет работу
}

/*
********************** методы для сохранения и загрузки настроек **********************
*/

void ProjectServerW::DataForm::SaveSettings() {
    try {
        // Создаем или открываем файл настроек
        // Получаем директорию, где находится исполняемый файл
        String^ appPath = System::IO::Path::GetDirectoryName(System::Windows::Forms::Application::ExecutablePath);
        String^ settingsPath = System::IO::Path::Combine(appPath, "ExcelSettings.txt");

        System::IO::StreamWriter^ writer = gcnew System::IO::StreamWriter(settingsPath);
        // Записываем путь
        writer->WriteLine(textBoxExcelDirectory->Text);

        // Закрываем файл
        writer->Close();
    }
    catch (Exception^ ex) {
        // Обработка ошибок
        MessageBox::Show("Не удалось сохранить настройки: " + ex->Message);
        GlobalLogger::LogMessage(ConvertToStdString("Не удалось сохранить настройки: " + ex->Message));
    }
}

void ProjectServerW::DataForm::LoadSettings() {
    try {
        // Проверяем существование файла настроек
        String^ appPath = System::IO::Path::GetDirectoryName(System::Windows::Forms::Application::ExecutablePath);
        String^ settingsPath = System::IO::Path::Combine(appPath, "ExcelSettings.txt");
        if (System::IO::File::Exists(settingsPath)) {
            // Открываем и читаем файл
            System::IO::StreamReader^ reader = gcnew System::IO::StreamReader("ExcelSettings.txt");

            // Читаем первую строку (путь)
            String^ path = reader->ReadLine();

            // Закрываем файл
            reader->Close();

            // Обновляем текстовое поле и переменную пути
            if (path != nullptr && path->Length > 0) {
                textBoxExcelDirectory->Text = path;
                excelSavePath = path;
            }
        }
    }
    catch (Exception^ ex) {
        // Обработка ошибок
        MessageBox::Show("Не удалось загрузить настройки: " + ex->Message);
        GlobalLogger::LogMessage(ConvertToStdString("Не удалось загрузить настройки: " + ex->Message));
    }
}

// Обновление текстового поля директории сохранения файлов EXCEL
void ProjectServerW::DataForm::UpdateDirectoryTextBox(String^ path) {
    if (textBoxExcelDirectory != nullptr && !textBoxExcelDirectory->IsDisposed) {
        textBoxExcelDirectory->Text = path;
    }
}

//*******************************************************************************************
// Реализация метода инициализации названий битовых полей
void ProjectServerW::DataForm::InitializeBitFieldNames(gcroot<cli::array<cli::array<String^>^>^>& namesRef) {
    namesRef = gcnew cli::array<cli::array<String^>^>(10);

    // Инициализация массива состояния устройств дефростера
    namesRef[0] = gcnew cli::array<String^>(16) {
        "Heat0", "Heat1", "Heat2", "Heat3", // Работа нагревателя  (ТЭНа) 1..4
        "Vent0", "Vent1", "Vent2", "Vent3", // Работа циркуляционного вентилятора 1..4
        "InjW",     // Работа водяных форсунок
        "UP",       // Сигнал «Двигать ворота вверх»
        "DOWN",     // Сигнал «Двигать ворота вниз»
        "STOP",     // Сигнал «Остановить ворота»
        "Clse",     // Сигнал «Ворота закрыты»
        "Open",     // Сигнал «Ворота открыты»
        "Work",     // Лампа РАБОТА 
        "Alrm",     // Лампа АВАРИЯ 
        
    };

    // Инициализация массива cигналов управления устройствами дефростера
    namesRef[1] = gcnew cli::array<String^>(16) {
        "_V0", "_V1", "_V2", "_V3", // Циркуляционный вентилятор 1..4 включить
        "_H0", "_H1", "_H2", "_H3", // Нагреватель (ТЭН) 1..4 включить
        "_Out",     // Вытяжной вентилятор включить
        "_Inj",     // Водяную форсунку включить
        "_Flp",     // Закрыть защитную заслонку вытяжного вентилятора 
        "_Opn",     // Ворота открыть
        "_Stp",     // Ворота остановить
        "_Cls",     // Ворота закрыть
        "_Wrk",     // Включить зелёную лампу РАБОТА
        "_Red"      // Включить красную лампу СТОП
    };

    // Добавьте другие типы сенсоров по мере необходимости
}

////******************** Обработка завершения формы ***************************************
System::Void ProjectServerW::DataForm::DataForm_FormClosing(System::Object^ sender, System::Windows::Forms::FormClosingEventArgs^ e) {
    try {
        // Проверяем, есть ли данные в таблице И не были ли они уже экспортированы
        if (dataTable != nullptr && dataTable->Rows->Count > 0) {
            
            // Проверяем флаг: были ли данные уже записаны в Excel
            if (dataExportedToExcel) {
                // Данные уже были экспортированы - не дублируем запись
                GlobalLogger::LogMessage("Information: Данные уже были экспортированы в Excel, повторная запись не требуется");
            }
            else {
                // Данные НЕ были экспортированы - это аварийное сохранение при потере связи
                // Создаем имя файла, если оно еще не создано
                if (excelFileName == nullptr) {
                    DateTime now = DateTime::Now;
                    excelFileName = "EmergencyData_" + now.ToString("yyyy-MM-dd_HH-mm-ss") + "_Port" + clientPort.ToString() + ".xlsx";
                }

                // Логируем начало сохранения
                GlobalLogger::LogMessage("АВАРИЙНОЕ сохранение данных в Excel (потеря связи)... " + ConvertToStdString(excelFileName));

                // Запускаем запись в Excel в отдельном потоке
                // Форма закроется сразу, не дожидаясь завершения записи
                DataForm::TriggerExcelExport();
                
                // НЕ ЖДЁМ завершения записи - форма закрывается сразу
                // Это позволяет клиенту переподключиться немедленно
                GlobalLogger::LogMessage("Information: Форма закрывается, запись в Excel продолжается в фоновом режиме");
            }
        }
        
        // Разрешаем закрытие формы немедленно
        e->Cancel = false;
    }
    catch (Exception^ ex) {
        GlobalLogger::LogMessage("Ошибка при попытке сохранения данных: " + ConvertToStdString(ex->Message));
        // В случае ошибки также разрешаем закрытие формы
        e->Cancel = false;
    }
}

// Обработчик события после закрытия формы
System::Void DataForm::DataForm_FormClosed(Object^ sender, FormClosedEventArgs^ e)
{
    // Удаляем форму из карты (если она ещё там есть)
    // Форма могла быть удалена ранее в CloseForm()
    bool formFoundInMap = false;
    for (auto it = formData_Map.begin(); it != formData_Map.end(); ++it) {
        // Извлекаем управляемый указатель из gcroot
        ProjectServerW::DataForm^ formPtr = it->second;

        // Теперь сравниваем указатели
        if (formPtr == this) {
            // Нашли текущую форму, удаляем её из карты
            formData_Map.erase(it);
            formFoundInMap = true;
            GlobalLogger::LogMessage("Information: Форма удалена из карты в FormClosed");
            break;
        }
    }
    
    if (!formFoundInMap) {
        GlobalLogger::LogMessage("Information: Форма уже была удалена из карты ранее");
    }
    
    // Закрываем сокет клиента
    closesocket(this->ClientSocket);
}

// Обработчик уничтожения дескриптора окна
System::Void DataForm::DataForm_HandleDestroyed(Object^ sender, EventArgs^ e)
{
    try {
        // Отписываемся от всех событий
        this->FormClosing -= gcnew FormClosingEventHandler(this, &DataForm::DataForm_FormClosing);
        this->FormClosed -= gcnew FormClosedEventHandler(this, &DataForm::DataForm_FormClosed);
        this->HandleDestroyed -= gcnew EventHandler(this, &DataForm::DataForm_HandleDestroyed);

        // Принудительная сборка мусора для освобождения COM-ресурсов
        GC::Collect();
        GC::WaitForPendingFinalizers();
        GC::Collect();
    }
    catch (...) {
        // Игнорируем исключения в деструкторе
    }

    // Очищаем таймер
    if (delayedExcelTimer != nullptr) {
        delayedExcelTimer->Stop();
        delete delayedExcelTimer;
        delayedExcelTimer = nullptr;
    }
}
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
            GlobalLogger::LogMessage("Error: Не могу отправить команду " + ConvertToStdString(commandName) + 
                                   ", нет активного соединения с клиентом!");
            return false;
        }

        // Формируем буфер команды
        uint8_t buffer[MAX_COMMAND_SIZE];
        size_t commandLength = BuildCommandBuffer(cmd, buffer, sizeof(buffer));

        if (commandLength == 0) {
            String^ errorMsg = "Ошибка формирования команды " + commandName;
            MessageBox::Show(errorMsg);
            GlobalLogger::LogMessage("Error: " + ConvertToStdString(errorMsg));
            return false;
        }

        // Отправляем команду клиенту
        const int bytesSent = send(clientSocket, reinterpret_cast<const char*>(buffer), 
                                  static_cast<int>(commandLength), 0);

        if (bytesSent == SOCKET_ERROR) {
            int error = WSAGetLastError();
            String^ errorMsg = "Ошибка отправки команды " + commandName + ": " + error.ToString();
            MessageBox::Show(errorMsg);
            GlobalLogger::LogMessage("Error: " + ConvertToStdString(errorMsg));
            return false;
        }
        else if (bytesSent == commandLength) {
            // Команда успешно отправлена
            Label_Commands->Text = "Команда " + commandName + " отправлена клиенту";
            GlobalLogger::LogMessage("Information: Команда " + ConvertToStdString(commandName) + 
                                   " отправлена клиенту");
            return true;
        }
        else {
            // Отправлено меньше байт, чем ожидалось
            String^ errorMsg = "Отправлено только " + bytesSent.ToString() + " из " + 
                             commandLength.ToString() + " байт для команды " + commandName;
            MessageBox::Show(errorMsg);
            GlobalLogger::LogMessage("Error: Частичная отправка команды " + 
                                   ConvertToStdString(commandName) + ": " + 
                                   ConvertToStdString(errorMsg));
            return false;
        }
    }
    catch (Exception^ ex) {
        String^ errorMsg = "Исключение при отправке команды " + commandName + ": " + ex->Message;
        MessageBox::Show(errorMsg);
        GlobalLogger::LogMessage("Error: " + ConvertToStdString(errorMsg));
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
        // Команда успешно выполнена на контроллере
        buttonSTOPstate_TRUE();
        Label_Commands->Text = "[OK] Программа запущена";
        Label_Commands->ForeColor = System::Drawing::Color::Green;
        GlobalLogger::LogMessage("Information: Команда START успешно выполнена контроллером");
        
        // Восстанавливаем цвет через 3 секунды с помощью таймера
        System::Windows::Forms::Timer^ colorTimer = gcnew System::Windows::Forms::Timer();
        colorTimer->Interval = 3000;
        colorTimer->Tick += gcnew EventHandler(this, &DataForm::RestoreLabelCommandsColor);
        colorTimer->Start();
    } else {
        // Ошибка выполнения команды - детали уже обработаны в ProcessResponse
        GlobalLogger::LogMessage(ConvertToStdString(String::Format(
            "Error: Команда START не выполнена. Статус: 0x{0:X2} ({1})",
            response.status, gcnew String(GetStatusName(response.status)))));
        
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
        // Команда успешно выполнена на контроллере
        buttonSTARTstate_TRUE();
        Label_Commands->Text = "[OK] Программа остановлена";
        Label_Commands->ForeColor = System::Drawing::Color::Green;
        GlobalLogger::LogMessage("Information: Команда STOP успешно выполнена контроллером");
        
        // Восстанавливаем цвет через 3 секунды с помощью таймера
        System::Windows::Forms::Timer^ colorTimer = gcnew System::Windows::Forms::Timer();
        colorTimer->Interval = 3000;
        colorTimer->Tick += gcnew EventHandler(this, &DataForm::RestoreLabelCommandsColor);
        colorTimer->Start();
    } else {
        // Ошибка выполнения команды - детали уже обработаны в ProcessResponse
        GlobalLogger::LogMessage(ConvertToStdString(String::Format(
            "Error: Команда STOP не выполнена. Статус: 0x{0:X2} ({1})",
            response.status, gcnew String(GetStatusName(response.status)))));
        
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
        
        // Восстанавливаем цвет через 3 секунды с помощью таймера
        System::Windows::Forms::Timer^ colorTimer = gcnew System::Windows::Forms::Timer();
        colorTimer->Interval = 3000;
        colorTimer->Tick += gcnew EventHandler(this, &DataForm::RestoreLabelCommandsColor);
        colorTimer->Start();
    } else {
        // Ошибка выполнения команды - детали уже обработаны в ProcessResponse
        GlobalLogger::LogMessage(ConvertToStdString(String::Format(
            "Error: Команда RESET не выполнена. Статус: 0x{0:X2} ({1})",
            response.status, gcnew String(GetStatusName(response.status)))));
        
        // Дополнительная обработка специфичных ошибок для RESET
        switch (response.status) {
            case CmdStatus::EXECUTION_ERROR:
                // Контроллер не может выполнить сброс
                Label_Commands->Text = "[!] Невозможно выполнить сброс контроллера";
                Label_Commands->ForeColor = System::Drawing::Color::Orange;
                break;
                
            case CmdStatus::TIMEOUT:
                // Контроллер не успел выполнить сброс
                Label_Commands->Text = "[!] Таймаут сброса контроллера";
                Label_Commands->ForeColor = System::Drawing::Color::Orange;
                break;
                
            default:
                // Другие ошибки уже отображены в ProcessResponse
                break;
        }
    }
}

// ============================
// Методы для обработки ответов от контроллера
// ============================

// Добавление ответа в очередь (вызывается из SServer)
void ProjectServerW::DataForm::EnqueueResponse(cli::array<System::Byte>^ response) {
    try {
        responseQueue->Enqueue(response);
        responseAvailable->Release(); // Сигнализируем о доступности ответа
        
        GlobalLogger::LogMessage(ConvertToStdString(String::Format(
            "Ответ поставлен в очередь на обработку: Type=0x{0:X2}, Size={1} bytes", 
            response[0], response->Length)));
    }
    catch (Exception^ ex) {
        GlobalLogger::LogMessage(ConvertToStdString("Ошибка при постановке ответа в очередь: " + ex->Message));
    }
}

// Прием ответа от контроллера
bool ProjectServerW::DataForm::ReceiveResponse(CommandResponse& response, int timeoutMs) {
    try {
        // Проверяем, что сокет клиента открыт
        if (clientSocket == INVALID_SOCKET) {
            GlobalLogger::LogMessage("Error: Клиентский сокет недопустим для получения ответа");
            return false;
        }

        // ===== ОЖИДАНИЕ ОТВЕТА ИЗ ОЧЕРЕДИ =====
        // Ждем появления ответа в очереди с таймаутом
        if (!responseAvailable->WaitOne(timeoutMs)) {
            String^ msg = "Timeout: Ответ от контроллера не получен";
            GlobalLogger::LogMessage(ConvertToStdString(msg));
            return false;
        }

        // Получаем ответ из очереди
        cli::array<System::Byte>^ responseBuffer;
        if (!responseQueue->TryDequeue(responseBuffer)) {
            String^ msg = "Error: Не удалось вывести ответ из очереди";
            GlobalLogger::LogMessage(ConvertToStdString(msg));
            return false;
        }

        // Копируем данные в неуправляемый буфер
        uint8_t buffer[64];
        pin_ptr<System::Byte> pinnedBuffer = &responseBuffer[0];
        memcpy(buffer, pinnedBuffer, responseBuffer->Length);

        // Разбираем полученный ответ
        if (!ParseResponseBuffer(buffer, responseBuffer->Length, response)) {
            String^ msg = "Error: Не удалось проанализировать ответ от контроллера";
            GlobalLogger::LogMessage(ConvertToStdString(msg));
            return false;
        }

        // Логируем успешное получение ответа
        String^ logMsg = String::Format(
            "Ответ обработан: Type=0x{0:X2}, Code=0x{1:X2}, Status={2}, DataLen={3}",
            response.commandType, response.commandCode, 
            gcnew String(GetStatusName(response.status)), response.dataLength);
        GlobalLogger::LogMessage(ConvertToStdString(logMsg));

        return true;

    } catch (Exception^ ex) {
        String^ errorMsg = "Исключение в полученном ответе: " + ex->Message;
        GlobalLogger::LogMessage(ConvertToStdString(errorMsg));
        return false;
    }
}

// Перегрузка ReceiveResponse с таймаутом по умолчанию
bool ProjectServerW::DataForm::ReceiveResponse(CommandResponse& response) {
    return ReceiveResponse(response, 1000); // Таймаут по умолчанию 1000 мс
}

// Обработка полученного ответа
void ProjectServerW::DataForm::ProcessResponse(const CommandResponse& response) {
    try {
        // Получаем имя команды и статус
        const char* statusName = GetStatusName(response.status);
        const char* statusDescription = GetStatusDescription(response.status);
        
        // Формируем сообщение о результате выполнения команды
        String^ message;
        
        if (response.status == CmdStatus::OK) {
            message = String::Format(
                "[OK] Команда Type=0x{0:X2}, Code=0x{1:X2} успешно выполнена",
                response.commandType, response.commandCode);

            // Если есть данные в ответе (для команд REQUEST)
            if (response.dataLength > 0) {
                // Обрабатываем данные в зависимости от типа команды
                if (response.commandType == CmdType::REQUEST) {
                    switch (response.commandCode) {
                        case CmdRequest::GET_STATUS: {
                            // Получаем статус (2 байта)
                            if (response.dataLength >= 2) {
                                uint16_t status;
                                memcpy(&status, response.data, 2);
                                message += String::Format("\nСтатус устройства: 0x{0:X4}", status);
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
                        case CmdRequest::GET_DATA: {
                            // Получаем данные конфигурации
                            if (response.dataLength >= 1) {
                                message += String::Format("\nРежим работы: {0}", 
                                    response.data[0] == 0 ? "Автоматический" : "Ручной");
                            }
                            break;
                        }
                    }
                }
            }
            
            // Отображаем успешный результат
            Label_Commands->Text = message;
            GlobalLogger::LogMessage(ConvertToStdString(message));
            
        } else {
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
            GlobalLogger::LogMessage(ConvertToStdString(message));
            
            // Восстанавливаем цвет текста через 3 секунды с помощью таймера
            System::Windows::Forms::Timer^ colorTimer = gcnew System::Windows::Forms::Timer();
            colorTimer->Interval = 3000;
            colorTimer->Tick += gcnew EventHandler(this, &DataForm::RestoreLabelCommandsColor);
            colorTimer->Start();
        }

    } catch (Exception^ ex) {
        String^ errorMsg = "Исключение в ProcessResponse: " + ex->Message;
        MessageBox::Show(errorMsg, "Критическая ошибка", 
                       MessageBoxButtons::OK, MessageBoxIcon::Error);
        GlobalLogger::LogMessage(ConvertToStdString(errorMsg));
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
        GlobalLogger::LogMessage(ConvertToStdString("Error in RestoreLabelCommandsColor: " + ex->Message));
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
        } else {
            sendResult = SendCommand(cmd);
        }

        if (!sendResult) {
            GlobalLogger::LogMessage("Error: Failed to send command");
            return false;
        }

        // Ждем ответ от контроллера
        if (!ReceiveResponse(response, 2000)) { // Таймаут 2 секунды
            GlobalLogger::LogMessage("Error: No response received from controller");
            return false;
        }

        // Проверяем, что ответ соответствует отправленной команде
        if (response.commandType != cmd.commandType || 
            response.commandCode != cmd.commandCode) {
            String^ errorMsg = String::Format(
                "Error: Response mismatch. Sent Type=0x{0:X2}, Code=0x{1:X2}; Received Type=0x{2:X2}, Code=0x{3:X2}",
                cmd.commandType, cmd.commandCode, response.commandType, response.commandCode);
            MessageBox::Show(errorMsg);
            GlobalLogger::LogMessage(ConvertToStdString(errorMsg));
            return false;
        }

        // Обрабатываем полученный ответ
        ProcessResponse(response);

        return (response.status == CmdStatus::OK);

    } catch (Exception^ ex) {
        String^ errorMsg = "Exception in SendCommandAndWaitResponse: " + ex->Message;
        MessageBox::Show(errorMsg);
        GlobalLogger::LogMessage(ConvertToStdString(errorMsg));
        return false;
    }
}

// Перегрузка SendCommandAndWaitResponse без указания имени команды
bool ProjectServerW::DataForm::SendCommandAndWaitResponse(
    const Command& cmd, CommandResponse& response) {
    return SendCommandAndWaitResponse(cmd, response, nullptr);
}

void ProjectServerW::DataForm::buttonSTARTstate_TRUE()
{
    buttonSTART->Enabled = true;
    labelSTART->BackColor = System::Drawing::Color::Snow;
    labelSTART->Text = "0";
    buttonSTOP->Enabled = false;
    labelSTOP->BackColor = System::Drawing::Color::OrangeRed;
    labelSTOP->Text = "1";
}

void ProjectServerW::DataForm::buttonSTOPstate_TRUE()
{
    buttonSTART->Enabled = false;
    labelSTART->BackColor = System::Drawing::Color::Lime;
    labelSTART->Text = "1";
    buttonSTOP->Enabled = true;
    labelSTOP->BackColor = System::Drawing::Color::Snow;
    labelSTOP->Text = "0";
}

// Обработчик события таймера для отложенной записи в Excel
void ProjectServerW::DataForm::OnDelayedExcelTimerTick(Object^ sender, EventArgs^ e) {
    // Проверяем, активен ли таймер отслеживания нуля
    if (!workBitZeroTimerActive) {
        delayedExcelTimer->Stop();
        return;
    }

    // Вычисляем прошедшее время с момента установки бита Work в ноль
    DateTime now = DateTime::Now;
    TimeSpan elapsed = now.Subtract(workBitZeroStartTime);
    
    if (elapsed.TotalSeconds >= 60) {
        // Прошло не менее 1 минуты непрерывного нахождения в нуле
        workBitZeroTimerActive = false;
        workBitZeroLogged = false;  // Сбрасываем флаг логирования
        delayedExcelTimer->Stop();
        
        // Сохраняем время окончания сбора данных
        dataCollectionEndTime = now;
        
        GlobalLogger::LogMessage(ConvertToStdString("Information: СТОП фиксации данных (по таймеру), запись в файл " + excelFileName));
        // Меняем состояние кнопок СТАРТ и СТОП
        buttonSTARTstate_TRUE();
        
        // Устанавливаем флаг, что данные экспортированы (чтобы не дублировать при закрытии формы)
        dataExportedToExcel = true;
        
        // Выполняем запись в Excel
        this->BeginInvoke(gcnew MethodInvoker(this, &DataForm::TriggerExcelExport));
    }
}

// ????????????????????????????????????????????????????????????????????????????
// Реализация методов записи в Excel (отложенный запуск)
// ????????????????????????????????????????????????????????????????????????????

// Триггер автоматического экспорта в Excel
void ProjectServerW::DataForm::TriggerExcelExport() {
    // Проверяем, вызывается ли метод из правильного потока
    if (this->InvokeRequired) {
        // Вызываем из потока UI
        try {
            this->Invoke(gcnew MethodInvoker(this, &DataForm::TriggerExcelExport));
        }
        catch (Exception^ ex) {
            GlobalLogger::LogMessage("Ошибка при вызове TriggerExcelExport: " + ConvertToStdString(ex->Message));
        }
        return;
    }
    
    // Теперь мы в потоке UI, можем безопасно обращаться к элементам управления
    try {
        // Проверка доступности кнопки Excel
        if (buttonExcel->Enabled) {
            // Автоматически запустить экспорт в Excel
            buttonEXCEL_Click(nullptr, nullptr);
        }
        else {
            // Кнопка недоступна, устанавливаем флаг ожидания и запускаем таймер
            pendingExcelExport = true;

            // Создаем таймер, если он еще не создан
            if (exportTimer == nullptr) {
                exportTimer = gcnew System::Windows::Forms::Timer();
                exportTimer->Interval = 500; // Проверка каждые 500 мс
                exportTimer->Tick += gcnew EventHandler(this, &DataForm::CheckExcelButtonStatus);
            }

            // Запускаем таймер
            exportTimer->Start();

            // Выводим сообщение для пользователя
            Label_Data->Text = "Ожидание возможности записи в Excel...";
        }
    }
    catch (Exception^ ex) {
        GlobalLogger::LogMessage("Ошибка в TriggerExcelExport: " + ConvertToStdString(ex->Message));
    }
}

// Обработчик события таймера проверки доступности кнопки экспорта в Excel
void ProjectServerW::DataForm::CheckExcelButtonStatus(Object^ sender, EventArgs^ e) {
    // Проверяем, доступна ли кнопка
    if (buttonExcel->Enabled && pendingExcelExport) {
        // Останавливаем таймер
        exportTimer->Stop();

        // Сбрасываем флаг
        pendingExcelExport = false;

        // Запускаем экспорт
        buttonEXCEL_Click(nullptr, nullptr);

        // Обновляем метку
        Label_Data->Text = "Данные записываются в Excel...";
    }
}

// ????????????????????????????????????????????????????????????????????????????
// Реализация методов автозапуска по времени
// ????????????????????????????????????????????????????????????????????????????

// Обработчик изменения состояния чекбокса автозапуска
System::Void ProjectServerW::DataForm::checkBoxAutoStart_CheckedChanged(System::Object^ sender, System::EventArgs^ e) {
    // Включение/выключение автозапуска
    if (checkBoxAutoStart->Checked) {
        // Включаем таймер
        timerAutoStart->Start();
        GlobalLogger::LogMessage(ConvertToStdString(String::Format(
            "Information: Автозапуск включен на {0}",
            dateTimePickerAutoStart->Value.ToString("HH:mm"))));
        
        // Визуальная индикация
        labelAutoStart->ForeColor = System::Drawing::Color::Green;
        dateTimePickerAutoStart->Enabled = true;
    }
    else {
        // Выключаем таймер
        timerAutoStart->Stop();
        GlobalLogger::LogMessage("Information: Автозапуск отключен");
        
        // Визуальная индикация
        labelAutoStart->ForeColor = System::Drawing::SystemColors::ControlText;
    }
}

// Обработчик тика таймера автозапуска
System::Void ProjectServerW::DataForm::timerAutoStart_Tick(System::Object^ sender, System::EventArgs^ e) {
    // Проверка времени для автозапуска
    if (!checkBoxAutoStart->Checked) {
        return; // Если автозапуск выключен, ничего не делаем
    }
    
    DateTime now = DateTime::Now;
    DateTime targetTime = dateTimePickerAutoStart->Value;
    
    // Сравниваем часы и минуты
    if (now.Hour == targetTime.Hour && now.Minute == targetTime.Minute) {
        // Время совпало!
        GlobalLogger::LogMessage(ConvertToStdString(String::Format(
            "Information: Автозапуск сработал в {0} (установлено: {1})",
            now.ToString("HH:mm:ss"),
            targetTime.ToString("HH:mm"))));
        
        // Проверяем, что кнопка START доступна (программа не запущена)
        if (buttonSTART->Enabled) {
            // Визуальная индикация
            labelAutoStart->ForeColor = System::Drawing::Color::Blue;
            Label_Commands->Text = "[АВТОЗАПУСК] Автоматический запуск программы...";
            Label_Commands->ForeColor = System::Drawing::Color::Blue;
            
            // Отправляем команду START
            SendStartCommand();
            
            // Отключаем автозапуск после выполнения
            checkBoxAutoStart->Checked = false;
            
            // Восстанавливаем цвет через 3 секунды
            System::Windows::Forms::Timer^ colorTimer = gcnew System::Windows::Forms::Timer();
            colorTimer->Interval = 3000;
            colorTimer->Tick += gcnew EventHandler(this, &DataForm::RestoreAutoStartColor);
            colorTimer->Start();
        }
        else {
            // Программа уже запущена
            GlobalLogger::LogMessage("Warning: Автозапуск не выполнен - программа уже запущена");
            Label_Commands->Text = "[!] Автозапуск пропущен - программа уже работает";
            Label_Commands->ForeColor = System::Drawing::Color::Orange;
            
            // Отключаем автозапуск
            checkBoxAutoStart->Checked = false;
        }
    }
}

// Обработчик восстановления цвета надписи автозапуска
System::Void ProjectServerW::DataForm::RestoreAutoStartColor(System::Object^ sender, System::EventArgs^ e) {
    try {
        // Останавливаем таймер
        System::Windows::Forms::Timer^ timer = safe_cast<System::Windows::Forms::Timer^>(sender);
        timer->Stop();
        timer->Tick -= gcnew EventHandler(this, &DataForm::RestoreAutoStartColor);
        
        // Восстанавливаем цвет
        labelAutoStart->ForeColor = System::Drawing::SystemColors::ControlText;
    }
    catch (Exception^ ex) {
        GlobalLogger::LogMessage(ConvertToStdString("Error in RestoreAutoStartColor: " + ex->Message));
    }
}
