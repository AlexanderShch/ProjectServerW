#include "Chart.h"
using namespace System;
using namespace System::Data;

typedef struct   // object data for Server type �� STM32
{
    uint16_t Time;				// ���������� ������ � ������� ���������
    uint8_t SensorQuantity;		// ���������� ��������
    uint8_t SensorType[SQ];		// ��� �������
    uint8_t Active[SQ];			// ���������� �������
    uint16_t T[SQ];				// �������� 1 ������� (�����������)
    uint16_t H[SQ];				// �������� 2 ������� (���������)
    uint16_t CRC_SUM;			// ����������� ��������
} MSGQUEUE_OBJ_t;

// �������������� ������ �� ������ � ��������� ���� MSGQUEUE_OBJ_t �� STM32
void ChartForm::ParseBuffer(const char* buffer, size_t size) {
    
    MSGQUEUE_OBJ_t data;

    if (size < sizeof(data)) {
        return; // �������� ������ ������� ����
    }
    memcpy(&data, buffer, sizeof(data));

    Sleep(200);
}

// 1. ������ ������� ������
void ChartForm::InitializeDataTable() {
    // ������� ������� ������
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

    // ������� BindingSource
    //bindingSource = gcnew BindingSource();
    bindingSource->DataSource = dataTable;

    // ����������� � DataGridView
    dataGridView->DataSource = bindingSource;
}

// 2. ���������� ������
//void ChatForm::AddDataToTable(DataTable^ table) {
//    // �������� ����� ������
//    DataRow^ row = table->NewRow();
//    //row["Time"] = data.Time;
//    row["SQ"] = 8080;
//
//    row["IP"] = "192.168.1.1";
//    row["Status"] = "Connected";
//    row["ConnectTime"] = DateTime::Now;
//
//    // ���������� ������ � �������
//    table->Rows->Add(row);
//}

// 3. ��������� ������� � EXCEL
// ���������� NuGet ����� EPPlus
//using namespace OfficeOpenXml;
//
//void SaveToExcelEPPlus(DataTable^ table, String^ filePath) {
//    ExcelPackage^ package = gcnew ExcelPackage();
//    ExcelWorksheet^ worksheet = package->Workbook->Worksheets->Add("Sheet1");
//
//    // ���������� ���������
//    for (int i = 0; i < table->Columns->Count; i++) {
//        worksheet->Cells[1, i + 1]->Value = table->Columns[i]->ColumnName;
//    }
//
//    // ���������� ������
//    for (int i = 0; i < table->Rows->Count; i++) {
//        for (int j = 0; j < table->Columns->Count; j++) {
//            worksheet->Cells[i + 2, j + 1]->Value =
//                table->Rows[i][j]->ToString();
//        }
//    }
//
//    package->SaveAs(gcnew System::IO::FileInfo(filePath));
//}