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
    short T[SQ];				// �������� 1 ������� (�����������)
    short H[SQ];				// �������� 2 ������� (���������)
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
    // ������ GUID ������� ����� � ����� formData_Map
    std::wstring currentFormGuid;
    bool formFound = false;

    for (auto it = formData_Map.begin(); it != formData_Map.end(); ++it) {
        // ��������� ����������� ��������� �� gcroot
        ProjectServerW::DataForm^ formPtr = it->second;

        // ������ ���������� ���������
        if (formPtr == this) {
            currentFormGuid = it->first;
            formFound = true;
            break;
        }
    }

    try {
        if (formFound) {
            // ��������� ���������� � ��������
            // ������� ����� �� GUID
            DataForm^ form = GetFormByGuid(currentFormGuid);

            if (form != nullptr && form->ClientSocket != INVALID_SOCKET) {
                // ��������� ����� �������
                closesocket(form->ClientSocket);
                form->ClientSocket = INVALID_SOCKET;
            }
        }
    }
    catch (Exception^ ex) {
        MessageBox::Show("������ ��� �������� ������: " + ex->Message);
        GlobalLogger::LogMessage(ConvertToStdString("������ ��� �������� ������: " + ex->Message));
    }

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

// ����� �������� �����, ����������� ��� �������� ������
void ProjectServerW::DataForm::CloseForm(const std::wstring& guid) {
    // ������� �����
    ProjectServerW::DataForm^ form = ProjectServerW::DataForm::GetFormByGuid(guid);
    //MessageBox::Show("DataForm will be closed!");
    GlobalLogger::LogMessage(ConvertToStdString("DataForm will be closed!"));

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
        DataForm^ form = it->second;

        // ���������, ��� ����� ���������� � �� �������
        if (form != nullptr && !form->IsDisposed && !form->Disposing) {
            return form;
        }
        else {
            // ����� ������� ��� ���������������, ������� � �� �����
            formData_Map.erase(it);
            return nullptr;
        }

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
        dataTable->Columns->Add("T" + i, short::typeid);
        dataTable->Columns->Add("H" + i, short::typeid);
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
        row["T" + i] = data.T[i] / 10.0;
        row["H" + i] = data.H[i] / 10.0;
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
            excelFileName = "WorkData_" + now.ToString("yyyy-MM-dd_HH-mm-ss") + "_Port" + clientPort.ToString() + ".xlsx";
        }

        // ���� ��� "Work" ��������� �� 1 � 0 (��������� is OFF)
        if (workBitDetected && !currentWorkBitState) {
            // ��������� ������ � Excel � ��������� ����������� ������
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
        // ��������� ������ ����� ���� �� ��������� ���������, � ������
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
                    ws->Cells[row, 6 + 4*i] = dr["T" + i];
                    ws->Cells[row, 7 + 4*i] = dr["H" + i];
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
            // ���������, ���������� �� ��������� ����
            if (String::IsNullOrEmpty(filePath) || !System::IO::Directory::Exists(filePath)) {
                // ���� ���� �� ���������� ��� ������, ������� ���������� � ����� ������� ���������
                String^ appPath = System::IO::Path::GetDirectoryName(System::Windows::Forms::Application::ExecutablePath);
                filePath = System::IO::Path::Combine(appPath, "SensorData");

                // ������� ����������, ���� ��� �� ����������
                if (!System::IO::Directory::Exists(filePath)) {
                    System::IO::Directory::CreateDirectory(filePath);
                }

                // ��������� ���������� excelSavePath ��� ������� ����������
                excelSavePath = filePath;

                // ��������� ��������� ����, ���� ��� ����������
                if (textBoxExcelDirectory != nullptr && !textBoxExcelDirectory->IsDisposed) {
                    if (textBoxExcelDirectory->InvokeRequired) {
                        textBoxExcelDirectory->BeginInvoke(gcnew System::Action<String^>(this, &DataForm::UpdateDirectoryTextBox), filePath);
                    }
                    else {
                        textBoxExcelDirectory->Text = filePath;
                    }
                }
                // ��������� � ����������
                DataForm::SaveSettings();
            }

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
                filePath += "SensorData_" + now.ToString("yyyy-MM-dd_HH-mm-ss") + "_Port" + clientPort.ToString() + ".xlsx";
            }

            // ���������� �� ���������� ����
            excel->SaveAs(filePath);

            // ��������� ������ �� copiedTable
            delete copiedTable;
            copiedTable = nullptr;
        }
    }
    catch (System::Runtime::InteropServices::COMException^ comEx) {
        // ����������� ��������� COM-����������
        String^ errorMsg = "COM error: " + comEx->Message + " (Code: " +
            comEx->ErrorCode.ToString("X8") + ")";
        //MessageBox::Show(errorMsg);
        GlobalLogger::LogMessage(ConvertToStdString(errorMsg));
    }
    catch (Exception^ ex) {
        // � ������ ������
        String^ errorMsg = "Excel error: " + ex->Message;
        //MessageBox::Show(errorMsg);
        GlobalLogger::LogMessage(ConvertToStdString(errorMsg));
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
            // ���������, �� ����������� �� ����� � ���������� �� ���
    // ��������� ��������� �����
            bool canInvoke = false;
            try {
                canInvoke = !this->IsDisposed && this->IsHandleCreated && this->Visible;
            }
            catch (...) {
                canInvoke = false;
            }

            // �������� ������ ������ ���� ����� ��������� ������� Invoke
            if (canInvoke) {
                try {
                    this->BeginInvoke(gcnew MethodInvoker(this, &DataForm::EnableButton));
                }
                catch (...) {
                    // ���������� ����� ���������� ��� ������ BeginInvoke
                }
            }
            else {
                // ���� ����� �����������, ������ ���������� ��������� ������
            }

            // ������ ����� ����� ����������� Excel
            if (excel != nullptr) {
                excel->Close();
                delete excel;
                excel = nullptr;

                // ����� ��������� ����� ��� ���������
                Thread::Sleep(50);
            }

            // ��������� ���������� ������ ������
            ThreadPool::QueueUserWorkItem(gcnew WaitCallback(DataForm::DelayedGarbageCollection));
            // ������������� � ���������� ��������
            exportSuccessful = true;
            if (exportCompletedEvent != nullptr) {
                exportCompletedEvent->Set();
            }
        }
        catch (...) {}
    }    

}

// ���������� ������ ������
void DataForm::DelayedGarbageCollection(Object^ state) {
    try {
        // ����� ����� ������� ������
        //MessageBox::Show("Excel file was recorded successfully!");
        GlobalLogger::LogMessage(ConvertToStdString("Excel file was recorded successfully!"));
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

void ProjectServerW::DataForm::AddDataToTableThreadSafe(cli::array<System::Byte>^ buffer, int size, int port) {
    // ���� ����� ����������� � ������ ����� ��������� Invoke
    // �����! ������, ���������� � UI, ������ ����������� � ������, ��������� ������ � UI!
    // ��������� ���� ������� � ����� DataForm
    this->clientPort = port;

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
        GlobalLogger::LogMessage(ConvertToStdString("�� ������� ��������� ���������: " + ex->Message));
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
        GlobalLogger::LogMessage(ConvertToStdString("�� ������� ��������� ���������: " + ex->Message));
    }
}

// ���������� ���������� ���� ���������� ���������� ������ EXCEL
void ProjectServerW::DataForm::UpdateDirectoryTextBox(String^ path) {
    if (textBoxExcelDirectory != nullptr && !textBoxExcelDirectory->IsDisposed) {
        textBoxExcelDirectory->Text = path;
    }
}

//*******************************************************************************************
// ���������� ������ ������������� �������� ������� �����
void ProjectServerW::DataForm::InitializeBitFieldNames(gcroot<cli::array<cli::array<String^>^>^>& namesRef) {
    namesRef = gcnew cli::array<cli::array<String^>^>(10);

    // ������������� ������� ��������� ��������� ����������
    namesRef[0] = gcnew cli::array<String^>(16) {
        "Heat0", "Heat1", "Heat2", "Heat3", // ������ �����������  (����) 1..4
        "Vent0", "Vent1", "Vent2", "Vent3", // ������ ��������������� ����������� 1..4
        "InjW",     // ������ ������� ��������
        "UP",       // ������ �������� ������ ������
        "DOWN",     // ������ �������� ������ ����
        "STOP",     // ������ ����������� ������
        "Clse",     // ������ ������� ��������
        "Open",     // ������ ������� ��������
        "Work",     // ����� ������ 
        "Alrm",     // ����� ������ 
        
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
System::Void ProjectServerW::DataForm::DataForm_FormClosing(System::Object^ sender, System::Windows::Forms::FormClosingEventArgs^ e) {
    try {
        // ��������, ��� ������ ������� ���������������
        if (exportCompletedEvent == nullptr) {
            exportCompletedEvent = gcnew System::Threading::ManualResetEvent(false);
        }
        // ���������� ������� ����� ������� ��������
        exportCompletedEvent->Reset();
        exportSuccessful = false;

        // ���������, ���� �� ������ � �������
        if (dataTable != nullptr && dataTable->Rows->Count > 0) {
            // ������� ��������� ���������� ������� FormClosing, ����� ������������� ��������� �����
            this->FormClosing -= gcnew System::Windows::Forms::FormClosingEventHandler(this, &DataForm::DataForm_FormClosing);
            // ��������� �������� ����� �� ����� ����������
            e->Cancel = true;

            // ������� ��� �����, ���� ��� ��� �� �������
            if (excelFileName == nullptr) {
                DateTime now = DateTime::Now;
                excelFileName = "EmergencyData_" + now.ToString("yyyy-MM-dd_HH-mm-ss") + "_Port" + clientPort.ToString() + ".xlsx";
            }

            // ��������� �����
            if (Label_Data != nullptr) {
                Label_Data->Text = "���������� ������ � Excel...";
                GlobalLogger::LogMessage("���������� ������ � Excel... " + ConvertToStdString(excelFileName));
                Label_Data->Refresh();
            }

            // ��������� ������ � Excel � ��������� ������, ������� ����� ����� ������� ���������� ������
            DataForm::TriggerExcelExport();

            // ������� ���������� ������ ����� Excel � ���������
            if (exportCompletedEvent->WaitOne(5*60*1000)) { // 5 ����� �������
                if (exportSuccessful) {
                    //MessageBox::Show("������� ������� ��������!");
                    GlobalLogger::LogMessage("������� ������� ��������!");
                }
                else {
                    //MessageBox::Show("������ ��� �������� ������");
                    GlobalLogger::LogMessage("������ ��� �������� ������");
                }
            }
            else {
                //MessageBox::Show("��������� ����� �������� ��������");
                GlobalLogger::LogMessage("��������� ����� �������� ��������");
            }

            // ��������� �����
            e->Cancel = false;

        }
    }
    catch (Exception^ ex) {
        //MessageBox::Show("������ ��� ������� ���������� ������: " + ex->Message);
        GlobalLogger::LogMessage("������ ��� ������� ���������� ������: " + ConvertToStdString(ex->Message));
        // � ������ ������ ��������� �������� �����
        e->Cancel = false;
    }
}

// ���������� ������� ����� �������� �����
System::Void DataForm::DataForm_FormClosed(Object^ sender, FormClosedEventArgs^ e)
{
    // ������� ����� �� �����
    for (auto it = formData_Map.begin(); it != formData_Map.end(); ++it) {
        // ��������� ����������� ��������� �� gcroot
        ProjectServerW::DataForm^ formPtr = it->second;

        // ������ ���������� ���������
        if (formPtr == this) {
            // ����� ������� �����, ������� � �� �����
            formData_Map.erase(it);
            break;
        }
    }
    // ��������� ����� �������
    closesocket(this->ClientSocket);
}

// ���������� ����������� ����������� ����
System::Void DataForm::DataForm_HandleDestroyed(Object^ sender, EventArgs^ e)
{
    try {
        // ������������ �� ���� �������
        this->FormClosing -= gcnew FormClosingEventHandler(this, &DataForm::DataForm_FormClosing);
        this->FormClosed -= gcnew FormClosedEventHandler(this, &DataForm::DataForm_FormClosed);
        this->HandleDestroyed -= gcnew EventHandler(this, &DataForm::DataForm_HandleDestroyed);

        // �������������� ������ ������ ��� ������������ COM-��������
        GC::Collect();
        GC::WaitForPendingFinalizers();
        GC::Collect();
    }
    catch (...) {
        // ���������� ���������� � �����������
    }
}
