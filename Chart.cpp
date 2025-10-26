#include "Chart.h"

using namespace System;
using namespace System::Windows::Forms;
using namespace System::Runtime::InteropServices;
using namespace Microsoft::Office::Interop::Excel;

// �������� ��� ��� �������� ��������
void CreateExcelFile() {
    ExcelHelper^ excel = gcnew ExcelHelper();
    try {
        // ������ � Excel
        excel->SaveAs("test.xlsx");
    }
    catch (Exception^ ex) {
        MessageBox::Show("Error: " + ex->Message);
    }
}

bool IsExcelInstalled() {
    try {
        Type^ type = Type::GetTypeFromProgID("Excel.Application");
        return (type != nullptr);
    }
    catch (...) {
        MessageBox::Show("Error: EXCEL �� ����������!");
        GlobalLogger::LogMessage("Error: EXCEL �� ����������!");
        return false;
    }
}