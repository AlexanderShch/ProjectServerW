#include "Chart.h"
// Тестовый код для проверки настроек
void CreateExcelFile() {
    ExcelHelper^ excel = gcnew ExcelHelper();
    try {
        // Работа с Excel
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
        MessageBox::Show("Error: EXCEL не установлен!");
        return false;
    }
}