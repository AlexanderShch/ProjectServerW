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
//
typedef struct   // object data for Server type из STM32
{
    uint16_t Time;				// Количество секунд с момента включения
    uint8_t SensorQuantity;		// Количество сенсоров
    uint8_t SensorType[SQ];		// Тип сенсора
    uint8_t Active[SQ];			// Активность сенсора
    short T[SQ];				// Значение 1 сенсора (температура)
    short H[SQ];				// Значение 2 сенсора (влажность)
    uint16_t CRC_SUM;			// Контрольное значение
} MSGQUEUE_OBJ_t;

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
String^ bufferToHex(const char* buffer, size_t length) {
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

        // Если бит "Work" переходит из 0 в 1 (дефростер is ON)
        if (!workBitDetected && currentWorkBitState) {
            // Сохраняем время начала сбора данных
            dataCollectionStartTime = now;
            // Создаем имя файла на основе времени начала (окончание добавится позже при записи в Excel)
            excelFileName = "WorkData_" + now.ToString("yyyy-MM-dd_HH-mm-ss") + "_Port" + clientPort.ToString();
            GlobalLogger::LogMessage(ConvertToStdString("Information: СТАРТ фиксации данных, создано имя файла " + excelFileName));
            // Меняем состояние кнопок СТАРТ и СТОП
            buttonSTOPstate_TRUE();
            // Сбрасываем таймер отслеживания нуля, если он был активен
            workBitZeroTimerActive = false;
            if (delayedExcelTimer != nullptr && delayedExcelTimer->Enabled) {
                delayedExcelTimer->Stop();
            }
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
            
            GlobalLogger::LogMessage(ConvertToStdString("Information: Бит Work стал нулём, начато отслеживание времени..."));
        }

        // Если бит "Work" остаётся в нуле и таймер активен, проверяем, прошла ли минута
        if (!currentWorkBitState && workBitZeroTimerActive) {
            TimeSpan elapsed = now.Subtract(workBitZeroStartTime);
            if (elapsed.TotalSeconds >= 60) {
                // Прошло не менее 1 минуты непрерывного нахождения в нуле
                workBitZeroTimerActive = false;
                delayedExcelTimer->Stop();
                
                // Сохраняем время окончания сбора данных
                dataCollectionEndTime = now;
                
                GlobalLogger::LogMessage(ConvertToStdString("Information: СТОП фиксации данных, запись в файл " + excelFileName));
                // Меняем состояние кнопок СТАРТ и СТОП
                buttonSTARTstate_TRUE();
                
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
        // DataTable не является потокобезопасной структурой, в Excel будем перегонять копию
        System::Data::DataTable^ copiedTable = dataTable->Copy();
        // Последняя строка может быть не заполнена полностью, её удалим
        if (copiedTable->Rows->Count > 0) {
            copiedTable->Rows->RemoveAt(copiedTable->Rows->Count - 1);
        }

        if (excel->CreateNewWorkbook()) {
            Microsoft::Office::Interop::Excel::Worksheet^ ws = excel->GetWorksheet();

            // Заголовки
            ws->Cells[1, 1] = "RealTime";
            ws->Cells[1, 2] = "Time";
            ws->Cells[1, 3] = "SQ";
            for (uint8_t i = 0; i < (SQ - 1); i++)
            {
                ws->Cells[1, 4 + 4*i] = "Typ" + i;
                ws->Cells[1, 5 + 4*i] = "Act" + i;
                // Добавляем имя датчика в заголовок T-столбца, если оно задано
                if (sensorNames != nullptr && i < sensorNames->Length && sensorNames[i] != nullptr && sensorNames[i]->Length > 0) {
                    ws->Cells[1, 6 + 4 * i] = "T" + i + " " + sensorNames[i];
                    ws->Cells[1, 7 + 4 * i] = "H" + i + " " + sensorNames[i];
                }
                else {
                    ws->Cells[1, 6 + 4 * i] = "T" + i;
                    ws->Cells[1, 7 + 4 * i] = "H" + i;
                }
            }
            // Обработка сигналов с устройства ввода-вывода
            ws->Cells[1, 4 + 4 * (SQ - 1)] = "Typ" + (SQ - 1);
            ws->Cells[1, 5 + 4 * (SQ - 1)] = "Act" + (SQ - 1);
            uint8_t ColomnNumber = 6 + 4 * (SQ - 1);

            cli::array<cli::array<String^>^>^ bitNames = GetBitFieldNames();
            // Работа с массивом имен сигналов состояния дефростера
            if (bitNames[0] != nullptr) {
                for (uint8_t i = 0; i < 16; i++)
                {
                    ws->Cells[1, ColomnNumber++] = bitNames[0][i];
                }
            }
            // Работа с массивом имен сигналов управления дефростером
            if (bitNames[1] != nullptr) {
                for (uint8_t i = 0; i < 16; i++)
                {
                    ws->Cells[1, ColomnNumber++] = bitNames[1][i];
                }
            }

            // Данные
            int row = 2;
            for each (DataRow ^ dr in copiedTable->Rows)
            {
                ws->Cells[row, 1] = dr["RealTime"]->ToString();
                ws->Cells[row, 2] = Convert::ToInt32(dr["Time"]);
                ws->Cells[row, 3] = Convert::ToUInt16(dr["SQ"]);
                for (uint8_t i = 0; i < (SQ - 1); i++)
                {
                    ws->Cells[row, 4 + 4*i] = Convert::ToUInt16(dr["Typ" + i]);
                    ws->Cells[row, 5 + 4*i] = Convert::ToUInt16(dr["Act" + i]);
                    ws->Cells[row, 6 + 4*i] = dr["T" + i];
                    ws->Cells[row, 7 + 4*i] = dr["H" + i];
                }

                ColomnNumber = 6 + 4 * (SQ - 1);
                // Работа с массивом имен сигналов состояния дефростера
                if (bitNames[0] != nullptr) {
                    for (uint8_t i = 0; i < 16; i++)
                    {
                        ws->Cells[row, ColomnNumber++] = Convert::ToUInt16(dr[bitNames[0][i]]);
                    }
                }
                // Работа с массивом имен сигналов управления дефростером
                if (bitNames[1] != nullptr) {
                    for (uint8_t i = 0; i < 16; i++)
                    {
                        ws->Cells[row, ColomnNumber++] = Convert::ToUInt16(dr[bitNames[1][i]]);
                    }
                }

                row++;
            }
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

                // Обновляем текстовое поле, если оно существует
                if (textBoxExcelDirectory != nullptr && !textBoxExcelDirectory->IsDisposed) {
                    if (textBoxExcelDirectory->InvokeRequired) {
                        textBoxExcelDirectory->BeginInvoke(gcnew System::Action<String^>(this, &DataForm::UpdateDirectoryTextBox), filePath);
                    }
                    else {
                        textBoxExcelDirectory->Text = filePath;
                    }
                }
                // Сохраняем в настройках
                DataForm::SaveSettings();
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
            GlobalLogger::LogMessage(ConvertToStdString("Information: EXCEL save to " + excelFileName));

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

    // Вызываем стандартный метод добавления данных
    AddDataToTable(rawBuffer, size, dataTable);

    // Автоматически обновляем форму после добавления данных
    if (dataGridView->RowCount > 0) {
        // Прокрутка к последней строке
        dataGridView->FirstDisplayedScrollingRowIndex = dataGridView->RowCount - 1;
    }

    // Обновляем UI (это безопасно, т.к. мы уже в потоке формы)
    this->Refresh();
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
        "_Snd",     // Звуковой сигнал включить
        "_Wrk"      // Включить зелёную лампу РАБОТА
    };

    // Добавьте другие типы сенсоров по мере необходимости
}

////******************** Обработка завершения формы ***************************************
System::Void ProjectServerW::DataForm::DataForm_FormClosing(System::Object^ sender, System::Windows::Forms::FormClosingEventArgs^ e) {
    try {
        // Проверка, что объект событий инициализирован
        if (exportCompletedEvent == nullptr) {
            exportCompletedEvent = gcnew System::Threading::ManualResetEvent(false);
        }
        // Сбрасываем событие перед началом экспорта
        exportCompletedEvent->Reset();
        exportSuccessful = false;

        // Проверяем, есть ли данные в таблице
        if (dataTable != nullptr && dataTable->Rows->Count > 0) {
            // Сначала отключаем обработчик события FormClosing, чтобы предотвратить повторный вызов
            this->FormClosing -= gcnew System::Windows::Forms::FormClosingEventHandler(this, &DataForm::DataForm_FormClosing);
            // Блокируем закрытие формы на время сохранения
            e->Cancel = true;

            // Создаем имя файла, если оно еще не создано
            if (excelFileName == nullptr) {
                DateTime now = DateTime::Now;
                excelFileName = "EmergencyData_" + now.ToString("yyyy-MM-dd_HH-mm-ss") + "_Port" + clientPort.ToString() + ".xlsx";
            }

            // Обновляем метку
            if (Label_Data != nullptr) {
                Label_Data->Text = "Сохранение данных в Excel...";
                GlobalLogger::LogMessage("Сохранение данных в Excel... " + ConvertToStdString(excelFileName));
                Label_Data->Refresh();
            }

            // Запускаем запись в Excel в отдельном потоке, текущий поток будет ожидать завершения записи
            DataForm::TriggerExcelExport();

            // Ожидаем завершения записи файла Excel с таймаутом
            if (exportCompletedEvent->WaitOne(60*60*1000)) { // 60 минут выделено на запись файла
                if (exportSuccessful) {
                    //MessageBox::Show("Экспорт успешно завершен!");
                    GlobalLogger::LogMessage("Information: Запись в EXCEL успешно завершена!");
                }
                else {
                    //MessageBox::Show("Ошибка при экспорте данных");
                    GlobalLogger::LogMessage("Error: Ошибка при записи данных в EXCEL");
                }
            }
            else {
                //MessageBox::Show("Превышено время ожидания экспорта");
                GlobalLogger::LogMessage("Error: Превышено время ожидания записи в EXCEL");
            }

            // Закрываем форму
            e->Cancel = false;

        }
    }
    catch (Exception^ ex) {
        //MessageBox::Show("Ошибка при попытке сохранения данных: " + ex->Message);
        GlobalLogger::LogMessage("Ошибка при попытке сохранения данных: " + ConvertToStdString(ex->Message));
        // В случае ошибки разрешаем закрытие формы
        e->Cancel = false;
    }
}

// Обработчик события после закрытия формы
System::Void DataForm::DataForm_FormClosed(Object^ sender, FormClosedEventArgs^ e)
{
    // Удаляем форму из карты
    for (auto it = formData_Map.begin(); it != formData_Map.end(); ++it) {
        // Извлекаем управляемый указатель из gcroot
        ProjectServerW::DataForm^ formPtr = it->second;

        // Теперь сравниваем указатели
        if (formPtr == this) {
            // Нашли текущую форму, удаляем её из карты
            formData_Map.erase(it);
            break;
        }
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
            Label_Data->Text = "Команда " + commandName + " отправлена клиенту";
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
    // Создаем и отправляем команду START (имя определится автоматически)
    Command cmd = CreateControlCommand(CmdProgControl::START);
    
    if (SendCommand(cmd)) {
        // Команда успешно отправлена, изменяем состояние кнопок
        buttonSTOPstate_TRUE();
    }
}

// Метод для отправки команды STOP клиенту
void ProjectServerW::DataForm::SendStopCommand() {
    // Создаем и отправляем команду STOP (имя определится автоматически)
    Command cmd = CreateControlCommand(CmdProgControl::STOP);
    
    if (SendCommand(cmd)) {
        // Команда успешно отправлена, изменяем состояние кнопок
        buttonSTARTstate_TRUE();
    }
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
        delayedExcelTimer->Stop();
        
        // Сохраняем время окончания сбора данных
        dataCollectionEndTime = now;
        
        GlobalLogger::LogMessage(ConvertToStdString("Information: СТОП фиксации данных (по таймеру), запись в файл " + excelFileName));
        // Меняем состояние кнопок СТАРТ и СТОП
        buttonSTARTstate_TRUE();
        
        // Выполняем запись в Excel
        this->BeginInvoke(gcnew MethodInvoker(this, &DataForm::TriggerExcelExport));
    }
}
