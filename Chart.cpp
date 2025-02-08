#include "Chart.h"
using namespace System;
using namespace System::Data;

typedef struct   // object data for Server type из STM32
{
    uint16_t Time;				// Количество секунд с момента включения
    uint8_t SensorQuantity;		// Количество сенсоров
    uint8_t SensorType[SQ];		// Тип сенсора
    uint8_t Active[SQ];			// Активность сенсора
    uint16_t T[SQ];				// Значение 1 сенсора (температура)
    uint16_t H[SQ];				// Значение 2 сенсора (влажность)
    uint16_t CRC_SUM;			// Контрольное значение
} MSGQUEUE_OBJ_t;

// Преобразование данных из буфера в структуру типа MSGQUEUE_OBJ_t из STM32
void ChartForm::ParseBuffer(const char* buffer, size_t size) {
    
    MSGQUEUE_OBJ_t data;

    if (size < sizeof(data)) {
        return; // Принятых данных слишком мало
    }
    memcpy(&data, buffer, sizeof(data));

    Sleep(200);
}

// 1. Создаём таблицу данных
void ChartForm::InitializeDataTable() {
    // Создаем таблицу данных
    dataTable = gcnew System::Data::DataTable();
    dataTable->Columns->Add("Time", uint16_t::typeid);
    dataTable->Columns->Add("SQ", uint8_t::typeid);

    for (uint8_t i = 0; i < SQ; i++)
    {
    dataTable->Columns->Add("Type" + i, uint8_t::typeid);
    dataTable->Columns->Add("Active" + i, uint8_t::typeid);
    dataTable->Columns->Add("T" + i, uint16_t::typeid);
    dataTable->Columns->Add("H" + i, uint16_t::typeid);
    }

    // Создаем BindingSource
    //bindingSource = gcnew BindingSource();
    bindingSource->DataSource = dataTable;

    // Привязываем к DataGridView
    dataGridView->DataSource = bindingSource;
}

// 2. Добавление данных
//void ChatForm::AddDataToTable(DataTable^ table) {
//    // Создание новой строки
//    DataRow^ row = table->NewRow();
//    //row["Time"] = data.Time;
//    row["SQ"] = 8080;
//
//    row["IP"] = "192.168.1.1";
//    row["Status"] = "Connected";
//    row["ConnectTime"] = DateTime::Now;
//
//    // Добавление строки в таблицу
//    table->Rows->Add(row);
//}

// 3. Сохраняем таблицу в EXCEL
// Установите NuGet пакет EPPlus
//using namespace OfficeOpenXml;
//
//void SaveToExcelEPPlus(DataTable^ table, String^ filePath) {
//    ExcelPackage^ package = gcnew ExcelPackage();
//    ExcelWorksheet^ worksheet = package->Workbook->Worksheets->Add("Sheet1");
//
//    // Записываем заголовки
//    for (int i = 0; i < table->Columns->Count; i++) {
//        worksheet->Cells[1, i + 1]->Value = table->Columns[i]->ColumnName;
//    }
//
//    // Записываем данные
//    for (int i = 0; i < table->Rows->Count; i++) {
//        for (int j = 0; j < table->Columns->Count; j++) {
//            worksheet->Cells[i + 2, j + 1]->Value =
//                table->Rows[i][j]->ToString();
//        }
//    }
//
//    package->SaveAs(gcnew System::IO::FileInfo(filePath));
//}