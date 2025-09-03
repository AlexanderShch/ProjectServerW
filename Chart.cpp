#include "Chart.h"
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
        return false;
    }
}