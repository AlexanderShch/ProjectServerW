#include "DataForm.h"
#include <objbase.h>                // ��� CoCreateGuid - ��������� ����������� ��������������
#include <string>
#include <vcclr.h>					// ��� ������������� gcroot

using namespace ProjectServerW; // ��������� ������������ ����

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
	Application::Exit();
	return System::Void();
}

void ProjectServerW::DataForm::CreateAndShowDataFormInThread(std::queue<std::wstring>& messageQueue,
                                                             std::mutex& mtx, 
                                                             std::condition_variable& cv) {
    DataForm^ form = gcnew DataForm();

    // ������������� ���������� COM
    HRESULT hr = CoInitialize(NULL);
    if (SUCCEEDED(hr)) {
        // ��������� ����������� �������������� �����
        GUID guid;
        HRESULT result = CoCreateGuid(&guid);
        if (result == S_OK) {
            // �������������� GUID � ������
            wchar_t guidString[40] = { 0 };
            int simb_N = StringFromGUID2(guid, guidString, 40);

            String^ formId = gcnew String(guidString);

            // ��������� ����������� �������������� � Label_ID
            form->SetData_FormID_value(formId);
			form->Refresh();

            // ���������� ����� � �����
            formData_Map[guidString] = form;

            // ��������� �������� � ������� ���������
            {
                std::lock_guard<std::mutex> lock(mtx);
                messageQueue.push(guidString);
            }
            cv.notify_one();
        }
        // ������������ ���������� COM
        CoUninitialize();
    }
    else {
         // ��������� ������ ������������� COM
         MessageBox::Show("������ ������������� COM ����������", "������", MessageBoxButtons::OK, MessageBoxIcon::Error);
    }

	form->ShowDialog();
}

// ����� �������� �����
void ProjectServerW::DataForm::CloseForm(const std::wstring& guid) {
    // ������� �����
    ProjectServerW::DataForm^ form = ProjectServerW::DataForm::GetFormByGuid(guid);

    if (form != nullptr) {
        // ���������, ����� �� Invoke
        if (form->InvokeRequired) {
            // ��������� ����� ����� Invoke
            form->Invoke(gcnew Action(form, &ProjectServerW::DataForm::Close));
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
    dataTable = gcnew DataTable("SensorData");
    dataTable->Columns->Add("Time", uint16_t::typeid);
    dataTable->Columns->Add("SQ", uint8_t::typeid);

    //DataGridViewCellStyle^ headerStyle = gcnew DataGridViewCellStyle();
    //headerStyle->Alignment = DataGridViewContentAlignment::MiddleLeft;
    //headerStyle->WrapMode = DataGridViewTriState::True;


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
void DataForm::AddDataToTable(DataTable^ table) {
    // �������� ����� ������
    DataRow^ row = table->NewRow();
    row["Time"] = 1010;// data.Time;
    row["SQ"] = 8080;

    // ���������� ������ � �������
    table->Rows->Add(row);
}

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