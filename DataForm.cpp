#include "DataForm.h"
#include "Chart.h"
#include <objbase.h>                // ��� CoCreateGuid - ��������� ����������� ��������������
#include <string>

// �������� ��� ������ ��� ������� � Process
using namespace System::Diagnostics;
using namespace ProjectServerW; // ��������� ������������ ����
using namespace Microsoft::Office::Interop::Excel;

std::map<std::wstring, gcroot<DataForm^>> formData_Map; // ����������� ���������� formData_Map
//
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

System::Void ProjectServerW::DataForm::buttonBrowse_Click(System::Object^ sender, System::EventArgs^ e)
{
    // ������� ������ ������ �����
    FolderBrowserDialog^ folderDialog = gcnew FolderBrowserDialog();

    // ��������� �������
    folderDialog->Description = "�������� ����� ��� ���������� Excel ������";
    folderDialog->ShowNewFolderButton = true;

    // ������������� ��������� ���������� �� ���������� ����
    if (textBoxExcelDirectory->Text->Length > 0) {
        folderDialog->SelectedPath = textBoxExcelDirectory->Text;
    }

    // ���������� ������ � ��������� ���������
    if (folderDialog->ShowDialog() == System::Windows::Forms::DialogResult::OK) {
        // ��������� ��������� ���� ��������� �����
        textBoxExcelDirectory->Text = folderDialog->SelectedPath;

        // ��������� ��������� ����
        excelSavePath = folderDialog->SelectedPath;
        // ���������� ��������
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

// ��� ������� ��������� ����� � ��� �����, ����� ����������� ������ ���� ������ � ����������������� ������ � ��������� ��� � ������.
String^ bufferToHex(const char* buffer, size_t length) {
    // ������������� std::stringstream ��������� ��������������� ������ � ������ � ������������� ��.
    std::stringstream ss;
    ss << std::hex << std::setfill('0');
    for (int i = 0; i < length; ++i) {
        ss << std::setw(2) << static_cast<unsigned int>(static_cast<unsigned char>(buffer[i])) << " ";
    }
    return gcnew String(ss.str().c_str());
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

    for (uint8_t i = 0; i < (SQ - 1); i++)
    {
        dataTable->Columns->Add("Typ" + i, uint8_t::typeid);
        dataTable->Columns->Add("Act" + i, uint8_t::typeid);
        dataTable->Columns->Add("T" + i, uint16_t::typeid);
        dataTable->Columns->Add("H" + i, uint16_t::typeid);
    }

    // ��������� �������� � ���������� �����-������
    dataTable->Columns->Add("Typ" + (SQ - 1), uint8_t::typeid);
    dataTable->Columns->Add("Act" + (SQ - 1), uint8_t::typeid);

    cli::array<cli::array<String^>^>^ bitNames = GetBitFieldNames();
    if (bitNames[0] != nullptr) {
        // ������ � �������� ���� �������� ��������� ����������
        for (uint8_t i = 0; i < 16; i++)
        {
            dataTable->Columns->Add(bitNames[0][i], uint8_t::typeid);
        }
        // ������ � �������� ���� �������� ���������� �����������
        for (uint8_t i = 0; i < 16; i++)
        {
            dataTable->Columns->Add(bitNames[1][i], uint8_t::typeid);
        }
    }

    dataGridView->DataSource = dataTable;
}

// 2. ���������� ������
void ProjectServerW::DataForm::AddDataToTable(const char* buffer, size_t size, System::Data::DataTable^ table) {
    
    MSGQUEUE_OBJ_t data;
    DateTime now = DateTime::Now;   // ��������� �������� �������
    String^ data_String{};

    if (size < sizeof(data)) {
        return; // �������� ������ ������� ����
    }
    memcpy(&data, buffer, sizeof(data));

    data_String = bufferToHex(buffer, size);
    // ������� �������� ������ data_String � ������� Label_Data
    DataForm::SetData_TextValue("�������� ������: " + data_String);

    // �������� ����� ������
    DataRow^ row = table->NewRow();
    row["RealTime"] = now.ToString("HH:mm:ss");
    row["Time"] = data.Time;
    row["SQ"] = data.SensorQuantity;
    for (uint8_t i = 0; i < (SQ - 1); i++)
    {
        row["Typ" + i] = data.SensorType[i];
        row["Act" + i] = data.Active[i];
        row["T" + i] = data.T[i];
        row["H" + i] = data.H[i];
    }
 
    cli::array<cli::array<String^>^>^ bitNames = GetBitFieldNames();

    // ������� ������ ���� "Work" � ������� bitNames[0]
    int workBitIndex = -1;
    // ������ ������ ���� "Work" � ������� ����
    for (int i = 0; i < bitNames[0]->Length; i++) {
        if (bitNames[0][i] == "Work") {
            workBitIndex = i;
            break;
        }
    }

    uint16_t bitField;
    bitField = data.T[SQ - 1];
    if (workBitIndex != -1) {
        // ��������� ������� ��������� ���� "Work"
        bool currentWorkBitState = (bitField & (1 << workBitIndex)) != 0;

        // ���� ��� "Work" ��������� �� 0 � 1 (��������� is ON)
        if (!workBitDetected && currentWorkBitState) {
            // ������� ��� ����� �� ������ �������� �������
            excelFileName = "WorkData_" + now.ToString("yyyy-MM-dd_HH-mm-ss") + ".xlsx";
            //this->BeginInvoke(gcnew System::Action<String^>(this, &DataForm::UpdateFileNameLabel), excelFileName);
        }

        // ���� ��� "Work" ��������� �� 1 � 0 (��������� is OFF)
        if (workBitDetected && !currentWorkBitState) {
            // ��������� ������ � Excel � ��������� ������
            this->BeginInvoke(gcnew MethodInvoker(this, &DataForm::TriggerExcelExport));
        }

        // ��������� ��������� �����
        workBitDetected = currentWorkBitState;
    }

    // ��������� ������ ��� � ��������� � ������� �������� ��������� ����������
    for (int bit = 0; bit < 16; bit++) {
        bool bitValue = (bitField & (1 << bit)) != 0;
        row[bitNames[0][bit]] = bitValue;
    }
    bitField = data.H[SQ - 1];
    // ��������� ������ ��� � ��������� � ������� �������� ���������� �����������
    for (int bit = 0; bit < 16; bit++) {
        bool bitValue = (bitField & (1 << bit)) != 0;
        row[bitNames[1][bit]] = bitValue;
    }

    // ���������� ������ � ������� ������ �� ����� ������ ����������
    if (workBitDetected)
    {
        table->Rows->Add(row);
    }
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
            for (uint8_t i = 0; i < (SQ - 1); i++)
            {
                ws->Cells[1, 4 + 4*i] = "Typ" + i;
                ws->Cells[1, 5 + 4*i] = "Act" + i;
                ws->Cells[1, 6 + 4*i] = "T" + i;
                ws->Cells[1, 7 + 4*i] = "H" + i;
            }
            // ��������� �������� � ���������� �����-������
            ws->Cells[1, 4 + 4 * (SQ - 1)] = "Typ" + (SQ - 1);
            ws->Cells[1, 5 + 4 * (SQ - 1)] = "Act" + (SQ - 1);
            uint8_t ColomnNumber = 6 + 4 * (SQ - 1);

            cli::array<cli::array<String^>^>^ bitNames = GetBitFieldNames();
            // ������ � �������� ���� �������� ��������� ����������
            if (bitNames[0] != nullptr) {
                for (uint8_t i = 0; i < 16; i++)
                {
                    ws->Cells[1, ColomnNumber++] = bitNames[0][i];
                }
            }
            // ������ � �������� ���� �������� ���������� �����������
            if (bitNames[1] != nullptr) {
                for (uint8_t i = 0; i < 16; i++)
                {
                    ws->Cells[1, ColomnNumber++] = bitNames[1][i];
                }
            }

            // ������
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
                    ws->Cells[row, 6 + 4*i] = Math::Round(Convert::ToSingle(dr["T" + i]) / 10.0, 1);
                    ws->Cells[row, 7 + 4*i] = Math::Round(Convert::ToSingle(dr["H" + i]) / 10.0, 1);
                }

                ColomnNumber = 6 + 4 * (SQ - 1);
                // ������ � �������� ���� �������� ��������� ����������
                if (bitNames[0] != nullptr) {
                    for (uint8_t i = 0; i < 16; i++)
                    {
                        ws->Cells[row, ColomnNumber++] = Convert::ToUInt16(dr[bitNames[0][i]]);
                    }
                }
                // ������ � �������� ���� �������� ���������� �����������
                if (bitNames[1] != nullptr) {
                    for (uint8_t i = 0; i < 16; i++)
                    {
                        ws->Cells[row, ColomnNumber++] = Convert::ToUInt16(dr[bitNames[1][i]]);
                    }
                }

                row++;
            }

            // ���������� �����
            // ��������� ������ ���� � �����
            String^ filePath = excelSavePath;

            // ��������� ����������� � ����� ����, ���� ��� ���
            if (!filePath->EndsWith("\\")) {
                filePath += "\\";
            }

            // ��������� ��� ����� � ������� ����� � ��������
            if (excelFileName != nullptr) {
                // ���������� ��� �����, ��������� ��� ����������� ���� "Work"
                filePath += excelFileName;
            }
            else {
                // ���������� ����������� ��� � ������� ����� � ��������
                DateTime now = DateTime::Now;
                filePath += "SensorData_" + now.ToString("yyyy-MM-dd_HH-mm-ss") + ".xlsx";
            }

            // ���������� �� ���������� ����
            excel->SaveAs(filePath);

            // �������� �� �������� �������� ����� �� �������� COM-��������
            Invoke(gcnew MethodInvoker(this, &DataForm::ShowSuccess));

            // ��������� ������ �� copiedTable
            delete copiedTable;
            copiedTable = nullptr;
        }
    }
    catch (System::Runtime::InteropServices::COMException^ comEx) {
        // ����������� ��������� COM-����������
        String^ errorMsg = "COM error: " + comEx->Message + " (Code: " +
            comEx->ErrorCode.ToString("X8") + ")";
        Invoke(gcnew System::Action<String^>(this, &DataForm::ShowError), errorMsg);
    }
    catch (Exception^ ex) {
        // � ������ ������
        String^ errorMsg = "Excel error: " + ex->Message;
        Invoke(gcnew System::Action<String^>(this, &DataForm::ShowError), errorMsg);
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
            Invoke(gcnew MethodInvoker(this, &DataForm::EnableButton));

            // ����� ��������� ����� ��� ��������� BeginInvoke
            Thread::Sleep(50);

            // ������ ����� ����� ����������� Excel
            if (excel != nullptr) {
                excel->Close();
                delete excel;
                excel = nullptr;
            }

            // ��������� ���������� ������ ������
            ThreadPool::QueueUserWorkItem(gcnew WaitCallback(DataForm::DelayedGarbageCollection));
        }
        catch (...) {}
    }    

}

// ���������� ������ ������
void DataForm::DelayedGarbageCollection(Object^ state) {
    try {
        // ����� ����� ������� ������
        Thread::Sleep(500);

        // ������ ������ � ��������� ������
        GC::Collect();
        GC::WaitForPendingFinalizers();
    }
    catch (...) {}
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

/*
********************** ������ ��� ���������� � �������� �������� **********************
*/

void ProjectServerW::DataForm::SaveSettings() {
    try {
        // ������� ��� ��������� ���� ��������
        // �������� ����������, ��� ��������� ����������� ����
        String^ appPath = System::IO::Path::GetDirectoryName(System::Windows::Forms::Application::ExecutablePath);
        String^ settingsPath = System::IO::Path::Combine(appPath, "ExcelSettings.txt");

        System::IO::StreamWriter^ writer = gcnew System::IO::StreamWriter(settingsPath);
        // ���������� ����
        writer->WriteLine(textBoxExcelDirectory->Text);

        // ��������� ����
        writer->Close();
    }
    catch (Exception^ ex) {
        // ��������� ������
        MessageBox::Show("�� ������� ��������� ���������: " + ex->Message);
    }
}

void ProjectServerW::DataForm::LoadSettings() {
    try {
        // ��������� ������������� ����� ��������
        String^ appPath = System::IO::Path::GetDirectoryName(System::Windows::Forms::Application::ExecutablePath);
        String^ settingsPath = System::IO::Path::Combine(appPath, "ExcelSettings.txt");
        if (System::IO::File::Exists(settingsPath)) {
            // ��������� � ������ ����
            System::IO::StreamReader^ reader = gcnew System::IO::StreamReader("ExcelSettings.txt");

            // ������ ������ ������ (����)
            String^ path = reader->ReadLine();

            // ��������� ����
            reader->Close();

            // ��������� ��������� ���� � ���������� ����
            if (path != nullptr && path->Length > 0) {
                textBoxExcelDirectory->Text = path;
                excelSavePath = path;
            }
        }
    }
    catch (Exception^ ex) {
        // ��������� ������
        MessageBox::Show("�� ������� ��������� ���������: " + ex->Message);
    }
}

//*******************************************************************************************
// ���������� ������ ������������� �������� ������� �����
void ProjectServerW::DataForm::InitializeBitFieldNames(gcroot<cli::array<cli::array<String^>^>^>& namesRef) {
    namesRef = gcnew cli::array<cli::array<String^>^>(10);

    // ������������� ������� ��������� ��������� ����������
    namesRef[0] = gcnew cli::array<String^>(16) {
        "Vent0", "Vent1", "Vent2", "Vent3", // ������ ��������������� ����������� 1..4
        "Heat0", "Heat1", "Heat2", "Heat3", // ������ �����������  (����) 1..4
        "OutA",     // ������ ��������� �����������
        "InjW",     // ������ ������� ��������
        "Work",     // ����� ������ 
        "Alrm",     // ����� ������ 
        "Open",     // ������ ������� ��������
        "Clse",     // ������ ������� ��������
        "StUP",     // ������ �������������� ��������� ����� ��� �������� ������
        "StDN"      // ������ �������������� ��������� ����� ��� �������� ����
    };

    // ������������� ������� c������� ���������� ������������ ����������
    namesRef[1] = gcnew cli::array<String^>(16) {
        "_V0", "_V1", "_V2", "_V3", // �������������� ���������� 1..4 ��������
        "_H0", "_H1", "_H2", "_H3", // ����������� (���) 1..4 ��������
        "_Out",     // �������� ���������� ��������
        "_Inj",     // ������� �������� ��������
        "_Flp",     // ������� �������� �������� ��������� ����������� 
        "_Opn",     // ������ �������
        "_Stp",     // ������ ����������
        "_Cls",     // ������ �������
        "_Snd",     // �������� ������ ��������
        "_Wrk"      // �������� ������ ����� ������
    };

    // �������� ������ ���� �������� �� ���� �������������
}

////******************** ��������� ���������� ����� ***************************************
//System::Void ProjectServerW::DataForm::DataForm_FormClosing(System::Object^ sender, System::Windows::Forms::FormClosingEventArgs^ e) {
//    try {
//        // ���������, ���� �� ������ � �������
//        if (dataTable != nullptr && dataTable->Rows->Count > 0) {
//            // ������� ��������� ���������� ������� FormClosing, ����� ������������� ��������� �����
//            this->FormClosing -= gcnew System::Windows::Forms::FormClosingEventHandler(this, &DataForm::DataForm_FormClosing);
//            // ��������� �������� ����� �� ����� ����������
//            e->Cancel = true;
//
//            // ������� ��� �����, ���� ��� ��� �� �������
//            if (excelFileName == nullptr) {
//                DateTime now = DateTime::Now;
//                excelFileName = "CloseData_" + now.ToString("yyyy-MM-dd_HH-mm-ss") + ".xlsx";
//            }
//
//            // ��������� �����
//            if (Label_Data != nullptr) {
//                Label_Data->Text = "���������� ������ � Excel...";
//                Label_Data->Refresh();
//            }
//
//            // ������� � ��������� ����� � STA-������� ��� ������ � Excel
//            Thread^ closeExcelThread = gcnew Thread(gcnew ThreadStart(this, &DataForm::AddDataToExcel));
//            closeExcelThread->SetApartmentState(ApartmentState::STA);
//            closeExcelThread->IsBackground = false;
//
//            // ��������� �����, ������� ����� ������ ����� ���������� ������
//            this->BeginInvoke(gcnew MethodInvoker(this, &DataForm::CloseFormAfterExport));
//
//            closeExcelThread->Start();
//        }
//    }
//    catch (Exception^ ex) {
//        MessageBox::Show("������ ��� ������� ���������� ������: " + ex->Message);
//        // � ������ ������ ��������� �������� �����
//        e->Cancel = false;
//    }
//}
//
//// ����� ��� �������� � �������� �����
//void ProjectServerW::DataForm::ExportAndClose() {
//    try {
//        // �������� ����� �������� ��������, �� ����� ������
//        AddDataToExcel();
//
//        // ����� ���������� �������� ��������� ����� �� ������ UI
//        this->BeginInvoke(gcnew MethodInvoker(this, &DataForm::CloseFormAfterExport));
//    }
//    catch (Exception^ ex) {
//        this->BeginInvoke(gcnew System::Action<String^>(this, &DataForm::ShowError),
//            "������ ��� ��������: " + ex->Message);
//
//        // � ������ ������ ��� ����� ��������� �����
//        this->BeginInvoke(gcnew MethodInvoker(this, &DataForm::CloseFormAfterExport));
//    }
//}

//// ����� ��� �������� ����� ����� ��������
//void ProjectServerW::DataForm::CloseFormAfterExport() {
//    // ���������, ���������� �� ����� �������� � Excel
//    // �������� ��������� ������ ��� ���������� ��������
//    int timeoutSeconds = 5;
//    DateTime startTime = DateTime::Now;
//
//    while (DateTime::Now.Subtract(startTime).TotalSeconds < timeoutSeconds) {
//        if (buttonExcel->Enabled) {
//            // ������ Excel �������� �������, ������ ������� ��������
//            break;
//        }
//        Thread::Sleep(100);
//    }
//
//    // ��������� ���������� �������� �����, ����� �������� ���������� ������
//    this->FormClosing -= gcnew System::Windows::Forms::FormClosingEventHandler(this, &DataForm::DataForm_FormClosing);
//
//    // ��������� �����
//    this->Close();
//}
//// ����� ��� ����������� ������ BeginInvoke
//void ProjectServerW::DataForm::SafeBeginInvoke(MethodInvoker^ method) {
//    try {
//        // ���������, ��� ����� ������� � �������
//        if (!this->IsDisposed && this->IsHandleCreated && !this->Disposing) {
//            this->BeginInvoke(method);
//        }
//    }
//    catch (...) {
//        // ���������� ����� ������
//    }
//}

//// ���������� ��� ������� � ����������
//void ProjectServerW::DataForm::SafeBeginInvoke(System::Action<String^>^ method, String^ param) {
//    try {
//        // ���������, ��� ����� ������� � �������
//        if (!this->IsDisposed && this->IsHandleCreated && !this->Disposing) {
//            this->BeginInvoke(method, param);
//        }
//    }
//    catch (...) {
//        // ���������� ����� ������
//    }
//}
