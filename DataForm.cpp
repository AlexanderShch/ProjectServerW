#include "DataForm.h"
#include "Chart.h"
#include <objbase.h>                // ��� CoCreateGuid - ��������� ����������� ��������������
#include <string>
#include <vcclr.h>					// ��� ������������� gcroot

// �������� ��� ������ ��� ������� � Process
using namespace System::Diagnostics;

using namespace ProjectServerW; // ��������� ������������ ����
using namespace Microsoft::Office::Interop::Excel;

std::map<std::wstring, gcroot<DataForm^>> formData_Map; // ����������� ���������� formData_Map

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
void ProjectServerW::DataForm::ParseBuffer(const char* buffer, size_t size) {

    MSGQUEUE_OBJ_t data;

    if (size < sizeof(data)) {
        return; // �������� ������ ������� ����
    }
    memcpy(&data, buffer, sizeof(data));
}

System::Void ProjectServerW::DataForm::�����ToolStripMenuItem_Click(System::Object^ sender, System::EventArgs^ e)
{
    ProjectServerW::DataForm::Close();
    System::Windows::Forms::Application::Exit();
	return System::Void();
}

System::Void ProjectServerW::DataForm::buttonEXCEL_Click(System::Object^ sender, System::EventArgs^ e)
{
    // ��������� ������ �� ����� ���������
    buttonExcel->Enabled = false;
    // ������� � ��������� �����
    excelThread = gcnew Thread(gcnew ThreadStart(this, &DataForm::AddDataToExcel));
    excelThread->SetApartmentState(ApartmentState::STA);  // ����������� ��� Excel!
    // �����! ��������� ��� ���������, �� �������� ������. 
    // ���� ������ ����������� ���������, ��������� ���������� ��������� �� ������ ����������
    excelThread->IsBackground = false; 
    excelThread->Start();

}

void ProjectServerW::DataForm::CreateAndShowDataFormInThread(std::queue<std::wstring>& messageQueue,
                                                             std::mutex& mtx, 
                                                             std::condition_variable& cv) {
    DataForm^ form = gcnew DataForm();
    GUID guid;
    HRESULT result = CoCreateGuid(&guid);
    if (result == S_OK) {
        // �������������� GUID � ������
        wchar_t guidString[40] = { 0 };
        int simb_N = StringFromGUID2(guid, guidString, 40);

        String^ formId = gcnew String(guidString);

        // ���������� ����� � �����
        formData_Map[guidString] = form;

        // ��������� �������� � ������� ���������
        {
            std::lock_guard<std::mutex> lock(mtx);
            messageQueue.push(guidString);
        }
        cv.notify_one();
    }
    form->ShowDialog();
 }

// ����� �������� �����
void ProjectServerW::DataForm::CloseForm(const std::wstring& guid) {
    // ������� �����
    ProjectServerW::DataForm^ form = ProjectServerW::DataForm::GetFormByGuid(guid);
    MessageBox::Show("DataForm will be closed!");

    if (form != nullptr) {
        // ���������, ����� �� Invoke
        if (form->InvokeRequired) {
            // ��������� ����� ����� Invoke
            form->Invoke(gcnew System::Action(form, &ProjectServerW::DataForm::Close));
        }
        else {
            form->Close();
        }
    }
}
        
// ����� ��� ��������� ����� �� � GUID
DataForm^ ProjectServerW::DataForm::GetFormByGuid(const std::wstring& guid) {
    auto it = formData_Map.find(guid);
    if (it != formData_Map.end()) {
        return it->second;
    }
    return nullptr;
}

// ���������� ���� guid � ����� � map
void ThreadStorage::StoreThread(const std::wstring& guid, std::thread& thread) {
    std::lock_guard<std::mutex> lock(GetMutex());
    GetThreadMap()[guid] = std::move(thread);
}

// ������������� ����� �� guid
void ThreadStorage::StopThread(const std::wstring& guid)
{
    std::lock_guard<std::mutex> lock(GetMutex());

    // ���� �����
    auto it = GetThreadMap().find(guid);
    if (it != GetThreadMap().end() && it->second.joinable()) {
        it->second.join();  // ���� ����������
        GetThreadMap().erase(it);  // ������� �� �����
    }
}

// ������� ��� ����������� ����������� ���������� map ��� ������
std::map<std::wstring, std::thread>& ThreadStorage::GetThreadMap() {
    static std::map<std::wstring, std::thread> threadMap;
    return threadMap;
}

// ������� ��� ����������� ����������� ���������� Mutex ��� ������
std::mutex& ThreadStorage::GetMutex() {
    static std::mutex mtx;
    return mtx;
}

/*
*********** ������ ������� *******************
*/

// 1. ������ ������� ������
void ProjectServerW::DataForm::InitializeDataTable() {
    // ������� ������� ������
    dataTable = gcnew System::Data::DataTable("SensorData");
 
    dataTable->Columns->Add("RealTime", String::typeid);
    dataTable->Columns->Add("Time", uint16_t::typeid);
    dataTable->Columns->Add("SQ", uint8_t::typeid);

    for (uint8_t i = 0; i < SQ; i++)
    {
        dataTable->Columns->Add("Typ" + i, uint8_t::typeid);
        dataTable->Columns->Add("Act" + i, uint8_t::typeid);
        dataTable->Columns->Add("T" + i, uint16_t::typeid);
        dataTable->Columns->Add("H" + i, uint16_t::typeid);
    }
    dataGridView->DataSource = dataTable;
}

// 2. ���������� ������
void ProjectServerW::DataForm::AddDataToTable(const char* buffer, size_t size, System::Data::DataTable^ table) {
    
    MSGQUEUE_OBJ_t data;
    DateTime now = DateTime::Now;   // ��������� �������� �������

    if (size < sizeof(data)) {
        return; // �������� ������ ������� ����
    }
    memcpy(&data, buffer, sizeof(data));

    // �������� ����� ������
    DataRow^ row = table->NewRow();
    row["RealTime"] = now.ToString("HH:mm:ss");
    row["Time"] = data.Time;
    row["SQ"] = data.SensorQuantity;
    for (uint8_t i = 0; i < SQ; i++)
    {
        row["Typ" + i] = data.SensorType[i];
        row["Act" + i] = data.Active[i];
        row["T" + i] = data.T[i];
        row["H" + i] = data.H[i];
    }

    // ���������� ������ � �������
    table->Rows->Add(row);
}

// 3. ��������� ������� � EXCEL
void ProjectServerW::DataForm::AddDataToExcel() {
    // ������ ������ Excel � STA-������
    ExcelHelper^ excel = nullptr;
    excel = gcnew ExcelHelper();

    try {
        // DataTable �� �������� ���������������� ����������, � Excel ����� ���������� �����
        System::Data::DataTable^ copiedTable = dataTable->Copy();
        // ��������� ������ ����� ���� �� �������� ���������, � ������
        if (copiedTable->Rows->Count > 0) {
            copiedTable->Rows->RemoveAt(copiedTable->Rows->Count - 1);
        }

        if (excel->CreateNewWorkbook()) {
            Microsoft::Office::Interop::Excel::Worksheet^ ws = excel->GetWorksheet();

            // ���������
            ws->Cells[1, 1] = "RealTime";
            ws->Cells[1, 2] = "Time";
            ws->Cells[1, 3] = "SQ";
            for (uint8_t i = 0; i < SQ; i++)
            {
                ws->Cells[1, 4 + 4*i] = "Typ" + i;
                ws->Cells[1, 5 + 4*i] = "Act" + i;
                ws->Cells[1, 6 + 4*i] = "T" + i;
                ws->Cells[1, 7 + 4*i] = "H" + i;
            }

            // ������
            int row = 2;
            for each (DataRow ^ dr in copiedTable->Rows)
            {
                ws->Cells[row, 1] = dr["RealTime"]->ToString();
                ws->Cells[row, 2] = Convert::ToInt32(dr["Time"]);
                ws->Cells[row, 3] = Convert::ToUInt16(dr["SQ"]);
                for (uint8_t i = 0; i < SQ; i++)
                {
                    ws->Cells[row, 4 + 4*i] = Convert::ToUInt16(dr["Typ" + i]);
                    ws->Cells[row, 5 + 4*i] = Convert::ToUInt16(dr["Act" + i]);
                    ws->Cells[row, 6 + 4*i] = Convert::ToInt32(dr["T" + i]);
                    ws->Cells[row, 7 + 4*i] = Convert::ToInt32(dr["H" + i]);
                }
                row++;
            }

            // ����������
            excel->SaveAs("D:\\SensorData.xlsx");
            // �������� �� �������� �������� ����� �� �������� COM-��������
            this->BeginInvoke(gcnew MethodInvoker(this, &DataForm::ShowSuccess));

            // ��������� ������ �� copiedTable
            delete copiedTable;
            copiedTable = nullptr;
        }
    }
    catch (System::Runtime::InteropServices::COMException^ comEx) {
        // ����������� ��������� COM-����������
        String^ errorMsg = "COM error: " + comEx->Message + " (Code: " +
            comEx->ErrorCode.ToString("X8") + ")";
        this->BeginInvoke(gcnew System::Action<String^>(this, &DataForm::ShowError), errorMsg);
    }
    catch (Exception^ ex) {
        // � ������ ������
        String^ errorMsg = "Excel error: " + ex->Message;
        this->BeginInvoke(gcnew System::Action<String^>(this, &DataForm::ShowError), errorMsg);
        // ���������� ���������� ������� ���� ��� ������
        try {
            if (excel != nullptr) {
                excel->Close();
                delete excel;
                excel = nullptr;
            }
        }
        catch (...) { /* ���������� ������ ��� ������� */ }
    }
    finally {
        try {
            // �����! ������� �������� ������ - �� ������������ COM-��������
            this->BeginInvoke(gcnew MethodInvoker(this, &DataForm::EnableButton));

            // ����� ��������� ����� ��� ��������� BeginInvoke
            Thread::Sleep(50);

            // ������ ����� ����� ����������� Excel
            if (excel != nullptr) {
                excel->Close();
                delete excel;
                excel = nullptr;
            }

            // ������� COM-�������� � ����� �����
            GC::Collect();
            GC::WaitForPendingFinalizers();
        }
        catch (...) {}
    }    

}

void DataForm::EnableButton()
{
    buttonExcel->Enabled = true;
}

// ��������������� ������ ��� UI-��������
void DataForm::ShowSuccess() {
    MessageBox::Show("Excel file was recorded successfully!");
}

void DataForm::ShowError(String^ message) {
    MessageBox::Show(message);
}

void ProjectServerW::DataForm::AddDataToTableThreadSafe(cli::array<System::Byte>^ buffer, int size) {
    // ���� ����� ��� ����������� � ������ ����� ��������� Invoke

    // ����������� ����������� ������ ������ � ������������� �����
    pin_ptr<Byte> pinnedBuffer = &buffer[0];
    char* rawBuffer = reinterpret_cast<char*>(pinnedBuffer);

    // �������� ����������� ����� ���������� ������
    AddDataToTable(rawBuffer, size, dataTable);

    // ������������� ��������� ����� ����� ���������� ������
    if (dataGridView->RowCount > 0) {
        // ��������� � ��������� ������
        dataGridView->FirstDisplayedScrollingRowIndex = dataGridView->RowCount - 1;
    }

    // ��������� UI (��� ���������, �.�. �� ��� � ������ �����)
    this->Refresh();
}