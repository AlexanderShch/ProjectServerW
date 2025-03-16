#pragma once
#pragma comment(lib, "ole32.lib")

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
    // Добавляем флаг инициализации COM
    bool isComInitialized;

    // Добавляем метод освобождения
    // Универсальный метод для освобождения COM-объектов
    template<typename T>
    void ReleaseComObject(T% comObj) {
        if (comObj != nullptr) {
            try {
                while (Marshal::ReleaseComObject(comObj) > 0) {
                    // Цикл для полного освобождения
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

        // Проверим тип апартмента текущего потока, должен быть STA для EXCEL
        if (SUCCEEDED(hra)) {
            DWORD desiredMode = COINIT_APARTMENTTHREADED;

            bool isSTA = (currentType == APTTYPE_STA);
            bool wantsSTA = (desiredMode == COINIT_APARTMENTTHREADED);

            if (isSTA != wantsSTA) {
                // Режимы не совпадают
                throw gcnew InvalidOperationException(
                    "COM mode mismatch!"
                );
            }
        }
        // Явная инициализация COM с обработкой ошибок
        HRESULT hr = CoInitializeEx(
            NULL,
            COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE
        );

        // Обрабатываем возможные коды возврата
        if (hr == RPC_E_CHANGED_MODE) {
             throw gcnew InvalidOperationException(
                "COM уже инициализирован в другом режиме");
        }
        else if (FAILED(hr)) {
            // Получаем текст ошибки
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

        // Все проверки пройдены, установим флаг, что COM инициализирован
        isComInitialized = true;

        try {
            // Инициализация безопасности
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

            if (FAILED(hr) && (hr != RPC_E_TOO_LATE)) {  // RPC_E_TOO_LATE - не ошибка
                switch (hr) {
                case RPC_E_TOO_LATE:
                    // Безопасность уже инициализирована
                    break;
                case E_INVALIDARG:
                    // Неверные параметры
                    break;
                case E_OUTOFMEMORY:
                    // Недостаточно памяти
                    break;
                default:
                    // Другие ошибки
                    break;
                }
                throw gcnew COMException("Failed to initialize COM security", hr);
            }
            // Создание Excel
            // Создание через Interop
            excel = gcnew Microsoft::Office::Interop::Excel::ApplicationClass();
            excel->Visible = false;
            excel->DisplayAlerts = false;
            isInitialized = true;
        }
        catch (Exception^ ex) {
            CoUninitialize();
            isComInitialized = false;
            MessageBox::Show("Excel initialization error: " + ex->Message);
        }
    }

    bool IsInitialized() {
        return isInitialized;
    }

    bool CreateNewWorkbook() {
        if (!isInitialized) return false;

        try {
            // Освобождаем предыдущие объекты
            if (workbook != nullptr) {
                ReleaseComObject(workbook);
                workbook = nullptr;
            }

            workbook = excel->Workbooks->Add(Type::Missing);
            worksheet = safe_cast<Worksheet^>(workbook->Worksheets->Item[1]);

            // Освобождаем временные объекты
            Marshal::ReleaseComObject(workbook->Worksheets);
            return true;
        }
        catch (Exception^ ex) {
            MessageBox::Show("Error creating workbook: " + ex->Message);
            return false;
        }
    }

    // Метод для получения worksheet
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
        try {
            // Сначала закрываем книгу, ЗАТЕМ освобождаем объекты
            if (workbook != nullptr) {
                try {
                    workbook->Close(false, Type::Missing, Type::Missing);
                }
                catch (...) {}
            }

            // Закрываем Excel
            if (excel != nullptr) {
                try {
                    excel->Quit();
                }
                catch (...) {}
            }

            // Теперь освобождаем объекты в обратном порядке их создания
            if (worksheet != nullptr) {
                try {
                    Marshal::ReleaseComObject(worksheet);
                }
                catch (...) {}
                worksheet = nullptr;
            }

            if (workbook != nullptr) {
                try {
                    Marshal::ReleaseComObject(workbook);
                }
                catch (...) {}
                workbook = nullptr;
            }
            // В последнюю очередь закрыть Excel приложение
            if (excel != nullptr) {
                try {
                    Marshal::ReleaseComObject(excel);
                }
                catch (...) {}
                excel = nullptr;
            }

        }
        catch (Exception^ ex) {
            // Логирование ошибки (можно заменить на свой метод)
#ifdef _DEBUG
            System::Diagnostics::Debug::WriteLine("Close error: " + ex->Message);
#endif
        }
        finally {
            isInitialized = false;  // Отмечаем, что объект закрыт
        }
    }

    // Финализатор
    !ExcelHelper() {
        try {
            if (worksheet != nullptr) {
                Marshal::ReleaseComObject(worksheet);
                worksheet = nullptr;
            }

            if (workbook != nullptr) {
                Marshal::ReleaseComObject(workbook);
                workbook = nullptr;
            }

            if (excel != nullptr) {
                excel->Quit();
                Marshal::ReleaseComObject(excel);
                excel = nullptr;
            }
        }
        catch (...) {}
    }


    ~ExcelHelper() {
        if (isComInitialized) {
            CoUninitialize();
            isComInitialized = false;
        }
        this->!ExcelHelper();
    }
};