#pragma once
#include "DataForm.h"

#using <C:\Windows\assembly\GAC_MSIL\Microsoft.Office.Interop.Excel\15.0.0.0__71e9bce111e9429c\Microsoft.Office.Interop.Excel.dll>
#using <mscorlib.dll>
#using <System.dll>

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
    // ��������� ���� ������������� COM
    bool isComInitialized;

    // ��������� ����� ������������
    // ������������� ����� ��� ������������ COM-��������
    template<typename T>
    void ReleaseComObject(T% comObj) {
        if (comObj != nullptr) {
            try {
                while (Marshal::ReleaseComObject(comObj) > 0) {
                    // ���� ��� ������� ������������
                }
            }
            finally {
                comObj = nullptr;
            }
        }
    }

public:
    ExcelHelper() : isInitialized(false) {
        APTTYPE currentType;
        APTTYPEQUALIFIER currentQualifier;

        HRESULT hra = CoGetApartmentType(&currentType, &currentQualifier);

        if (SUCCEEDED(hra)) {
            // COM ��� ���������������
            DWORD desiredMode = COINIT_APARTMENTTHREADED;

            bool isSTA = (currentType == APTTYPE_STA);
            bool wantsSTA = (desiredMode == COINIT_APARTMENTTHREADED);

            if (isSTA != wantsSTA) {
                // ������ �� ���������
                throw gcnew InvalidOperationException(
                    "COM mode mismatch!"
                );
            }
        }
        // ����� ������������� COM � ���������� ������
        HRESULT hr = CoInitializeEx(
            NULL,
            COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE
        );

        // ������������ ��������� ���� ��������
        if (hr == RPC_E_CHANGED_MODE) {
             throw gcnew InvalidOperationException(
                "COM ��� ��������������� � ������ ������");
        }
        else if (FAILED(hr)) {
            // �������� ����� ������
            LPWSTR errorText = nullptr;
            FormatMessage(
                FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM,
                NULL,
                hr,
                MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT),
                (LPWSTR)&errorText,
                0,
                NULL);

            String^ message = gcnew String(errorText);
            LocalFree(errorText);

            throw gcnew Exception("COM initialization failed: " + message);
        }

        isComInitialized = true;


        try {
            // ������������� ������������
            HRESULT hr = CoInitializeSecurity(
                NULL,                        // Security descriptor
                -1,                         // COM authentication
                NULL,                       // Authentication services
                NULL,                       // Reserved
                RPC_C_AUTHN_LEVEL_CONNECT,  // Default authentication level
                RPC_C_IMP_LEVEL_IMPERSONATE,// Default impersonation level
                NULL,                       // Authentication info
                EOAC_NONE,                  // Additional capabilities
                NULL                        // Reserved
            );

            if (FAILED(hr) && (hr != RPC_E_TOO_LATE)) {  // RPC_E_TOO_LATE - �� ������
                switch (hr) {
                case RPC_E_TOO_LATE:
                    // ������������ ��� ����������������
                    break;
                case E_INVALIDARG:
                    // �������� ���������
                    break;
                case E_OUTOFMEMORY:
                    // ������������ ������
                    break;
                default:
                    // ������ ������
                    break;
                }
                throw gcnew COMException("Failed to initialize COM security", hr);
            }
            // �������� Excel
            // �������� ����� Interop
            excel = gcnew Microsoft::Office::Interop::Excel::ApplicationClass();
            excel->Visible = false;
            excel->DisplayAlerts = false;
            isInitialized = true;
        }
        catch (Exception^ ex) {
            CoUninitialize();
            MessageBox::Show("Excel initialization error: " + ex->Message);
        }
    }

    bool IsInitialized() {
        return isInitialized;
    }

    bool CreateNewWorkbook() {
        if (!isInitialized) return false;

        try {
            // ����������� ���������� �������
            if (workbook != nullptr) {
                ReleaseComObject(workbook);
                workbook = nullptr;
            }

            workbook = excel->Workbooks->Add(Type::Missing);
            worksheet = safe_cast<Worksheet^>(workbook->Worksheets->Item[1]);

            // ����������� ��������� �������
            Marshal::ReleaseComObject(workbook->Worksheets);
            return true;
        }
        catch (Exception^ ex) {
            MessageBox::Show("Error creating workbook: " + ex->Message);
            return false;
        }
    }

    // ����� ��� ��������� worksheet
    Microsoft::Office::Interop::Excel::Worksheet^ GetWorksheet() {
        return worksheet;
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
        if (!isInitialized) return;

        try {
            // ���������� ������� ��������
            if (worksheet != nullptr) {
                Marshal::ReleaseComObject(worksheet);
                worksheet = nullptr;
            }

            if (workbook != nullptr) {
                workbook->Close(
                    false,  // SaveChanges
                    Type::Missing,  // FileName
                    Type::Missing   // RouteWorkbook
                );
                Marshal::ReleaseComObject(workbook);
                workbook = nullptr;
            }

            if (excel != nullptr) {
                excel->Quit();
                Marshal::ReleaseComObject(excel);
                excel = nullptr;
            }
        }
        catch (Exception^ ex) {
            MessageBox::Show("Close error: " + ex->Message);
        }
        finally {
            isInitialized = false;
        }

    }

    ~ExcelHelper() {
        Close();
        if (isComInitialized) {
            CoUninitialize();
        }
    }
};