#pragma once
#include "DataForm.h"

//#using <mscorlib.dll>
//#using <System.dll>
//#using <C:\Windows\assembly\GAC_MSIL\Microsoft.Office.Interop.Excel\15.0.0.0__71e9bce111e9429c\Microsoft.Office.Interop.Excel.dll>
//#using "Microsoft.Office.Interop.Excel.dll"

using namespace System;
using namespace System::Runtime::InteropServices;
using namespace Microsoft::Office::Interop::Excel;
using namespace System::Windows::Forms;

public ref class ExcelHelper {
private:
    Microsoft::Office::Interop::Excel::Application^ excel;
    Microsoft::Office::Interop::Excel::Workbook^ workbook;
    Microsoft::Office::Interop::Excel::Worksheet^ worksheet;
    bool isInitialized;

    void ReleaseComObject(Object^ obj) {
        try {
            if (obj != nullptr) {
                Marshal::ReleaseComObject(obj);
            }
        }
        catch (...) {}
    }

public:
    ExcelHelper() : isInitialized(false) {
        try {
            // Создание через COM
            Type^ excelType = Type::GetTypeFromProgID("Excel.Application");
            if (excelType != nullptr) {
                excel = safe_cast<Microsoft::Office::Interop::Excel::Application^>(
                    Activator::CreateInstance(excelType));
                excel->Visible = false;
                excel->DisplayAlerts = false;
                isInitialized = true;
            }
        }
        catch (Exception^ ex) {
            MessageBox::Show("Excel initialization error: " + ex->Message);
        }
    }

    bool IsInitialized() {
        return isInitialized;
    }

    bool CreateNewWorkbook() {
        if (!isInitialized) return false;

        try {
            workbook = excel->Workbooks->Add(Type::Missing);
            worksheet = safe_cast<Microsoft::Office::Interop::Excel::Worksheet^>(
                workbook->Sheets[1]);
            return true;
        }
        catch (Exception^ ex) {
            MessageBox::Show("Error creating workbook: " + ex->Message);
            return false;
        }
    }

    bool SaveAs(String^ filename) {
        if (!isInitialized || workbook == nullptr) return false;

        try {
            workbook->SaveAs(
                filename,
                XlFileFormat::xlOpenXMLWorkbook,
                Type::Missing,
                Type::Missing,
                Type::Missing,
                Type::Missing,
                XlSaveAsAccessMode::xlNoChange,
                Type::Missing,
                Type::Missing,
                Type::Missing,
                Type::Missing,
                Type::Missing
            );
            return true;
        }
        catch (Exception^ ex) {
            MessageBox::Show("Error saving file: " + ex->Message);
            return false;
        }
    }

    void Close() {
        if (workbook != nullptr) {
            // Закрываем с сохранением изменений
            workbook->Close(true, Type::Missing, Type::Missing);
        }
        if (excel != nullptr) {
            excel->Quit();
        }

        ReleaseComObject(worksheet);
        ReleaseComObject(workbook);
        ReleaseComObject(excel);

        worksheet = nullptr;
        workbook = nullptr;
        excel = nullptr;
        isInitialized = false;

        GC::Collect();
        GC::WaitForPendingFinalizers();
    }

    ~ExcelHelper() {
        Close();
    }
};