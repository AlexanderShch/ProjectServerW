#pragma once
#include "SServer.h"

#include <mutex>
#include <map>						// ��� ������������� std::map - ���������, ����������� ������������ ID � ������ �� �����
#include <condition_variable>
#include <queue>
#include <vcclr.h>					// ��� ������������� gcroot

#define SQ 6				// sensors quantity for measures (0-4) + sets of T (5, 6) + MB_IO

namespace ProjectServerW {

	using namespace System;
	using namespace System::ComponentModel;
	using namespace System::Collections;
	using namespace System::Windows::Forms;
	using namespace System::Data;
	using namespace System::Drawing;
	using namespace System::Threading;

	/// <summary>
	/// ������ ��� DataForm
	/// </summary>
	public ref class DataForm : public System::Windows::Forms::Form
	{
	private:
		SOCKET clientSocket;  // ����� �������
	public:
		property SOCKET ClientSocket {
			SOCKET get() { return clientSocket; }
			void set(SOCKET value) { clientSocket = value; }
		}

	private:
		System::Data::DataTable^ dataTable;  // ���������� ������� ��� ����� ������
		Thread^ excelThread;				 // ������� ������ ��� ������ � Excel � ��������� ������
		// ������ ������������ ������� ����� ��� ������� ���� �������
		// ������ ������� ������ - ��� �������, ������� ������ - ����� ����
		ManualResetEvent^ exportCompletedEvent;
		bool exportSuccessful;

	private: System::Windows::Forms::TabControl^ tabControl1;
	private: System::Windows::Forms::TabPage^ tabPage1;
	private: System::Windows::Forms::TabPage^ tabPage2;
	private: System::Windows::Forms::Button^ buttonBrowse;

	private: System::Windows::Forms::TextBox^ textBoxExcelDirectory;
	private: System::Windows::Forms::Label^ labelExcelDirectory;
	private: System::String^ excelSavePath;  // ��� �������� ����
		     System::String^ excelFileName;   // ��� �������� ����� �����, ���������� � ����� "Work"
		     bool workBitDetected;            // ���� ��� ������������ ��������� ���� "Work"
			 bool pendingExcelExport;		  // ���� �������� ������������ ������ ������ Excel
			 System::Windows::Forms::Timer^ exportTimer;	// ������ ��� �������� ������������ ������
	private: System::Windows::Forms::Label^ Label_Data;


	private: System::Windows::Forms::DataGridView^ dataGridView;

	public:
		DataForm(void)
		{
			InitializeComponent();
			GetBitFieldNames();	// ������������� ���� ����� (���� ��� �� ����������������)
			InitializeDataTable();

			// �������� �� ������� ��� �������� �����
			this->FormClosing += gcnew FormClosingEventHandler(this, &DataForm::DataForm_FormClosing);
			this->FormClosed += gcnew FormClosedEventHandler(this, &DataForm::DataForm_FormClosed);
			this->HandleDestroyed += gcnew EventHandler(this, &DataForm::DataForm_HandleDestroyed);
			// ������������� ������� ��� �������������
			exportCompletedEvent = gcnew System::Threading::ManualResetEvent(false);

			// �������������� ���� ���������� �� ���������� ����
			excelSavePath = textBoxExcelDirectory->Text;
			// ��� ����� ����� ������������� �����, ����� "Work" ������ ��������
			excelFileName = nullptr;
			workBitDetected = false;
			// ������������� ����� �������
			clientPort = 0;

			// ��������� ��������� ����� ����� ������������� �����������
			LoadSettings();

		}

	protected:
		/// <summary>
		/// ���������� ��� ������������ �������.
		/// </summary>
		~DataForm()
		{
			// ����� ������������ ����� ����������
			this->!DataForm();

			if (components)
			{
				delete components;
				components = nullptr;
			}
		}

		// �����������
		DataForm::!DataForm()
		{
			try {
				// ������� ��� ������������� �������
				if (clientSocket != INVALID_SOCKET) {
					closesocket(clientSocket);
					clientSocket = INVALID_SOCKET;
				}

				// ����������� COM-�������
				// ����������: ��� ������ ����������� � ������ STA
				if (Thread::CurrentThread->GetApartmentState() == ApartmentState::STA) {
					// ����������� COM-�������
				}
			}
			catch (...) {
				// ���������� ���������� � ������������
			}
		}

	private: System::Windows::Forms::MenuStrip^ menuStrip1;
	protected:
	private: System::Windows::Forms::ToolStripMenuItem^ �����ToolStripMenuItem;



	private: System::Windows::Forms::Button^ buttonExcel;

	private:
		/// <summary>
		/// ������������ ���������� ������������.
		/// </summary>
		System::ComponentModel::Container^ components;

#pragma region Windows Form Designer generated code
		/// <summary>
		/// ��������� ����� ��� ��������� ������������ � �� ��������� 
		/// ���������� ����� ������ � ������� ��������� ����.
		/// </summary>
		void InitializeComponent(void)
		{
			this->menuStrip1 = (gcnew System::Windows::Forms::MenuStrip());
			this->�����ToolStripMenuItem = (gcnew System::Windows::Forms::ToolStripMenuItem());
			this->buttonExcel = (gcnew System::Windows::Forms::Button());
			this->tabControl1 = (gcnew System::Windows::Forms::TabControl());
			this->tabPage1 = (gcnew System::Windows::Forms::TabPage());
			this->dataGridView = (gcnew System::Windows::Forms::DataGridView());
			this->tabPage2 = (gcnew System::Windows::Forms::TabPage());
			this->buttonBrowse = (gcnew System::Windows::Forms::Button());
			this->textBoxExcelDirectory = (gcnew System::Windows::Forms::TextBox());
			this->labelExcelDirectory = (gcnew System::Windows::Forms::Label());
			this->Label_Data = (gcnew System::Windows::Forms::Label());
			this->menuStrip1->SuspendLayout();
			this->tabControl1->SuspendLayout();
			this->tabPage1->SuspendLayout();
			(cli::safe_cast<System::ComponentModel::ISupportInitialize^>(this->dataGridView))->BeginInit();
			this->tabPage2->SuspendLayout();
			this->SuspendLayout();
			// 
			// menuStrip1
			// 
			this->menuStrip1->GripMargin = System::Windows::Forms::Padding(2, 2, 0, 2);
			this->menuStrip1->ImageScalingSize = System::Drawing::Size(24, 24);
			this->menuStrip1->Items->AddRange(gcnew cli::array< System::Windows::Forms::ToolStripItem^  >(1) { this->�����ToolStripMenuItem });
			this->menuStrip1->Location = System::Drawing::Point(0, 0);
			this->menuStrip1->Name = L"menuStrip1";
			this->menuStrip1->Size = System::Drawing::Size(2216, 33);
			this->menuStrip1->TabIndex = 0;
			this->menuStrip1->Text = L"menuStrip1";
			// 
			// �����ToolStripMenuItem
			// 
			this->�����ToolStripMenuItem->Name = L"�����ToolStripMenuItem";
			this->�����ToolStripMenuItem->Size = System::Drawing::Size(80, 29);
			this->�����ToolStripMenuItem->Text = L"�����";
			this->�����ToolStripMenuItem->Click += gcnew System::EventHandler(this, &DataForm::�����ToolStripMenuItem_Click);
			// 
			// buttonExcel
			// 
			this->buttonExcel->Location = System::Drawing::Point(1869, 16);
			this->buttonExcel->Name = L"buttonExcel";
			this->buttonExcel->Size = System::Drawing::Size(207, 42);
			this->buttonExcel->TabIndex = 5;
			this->buttonExcel->Text = L"������ � EXCEL";
			this->buttonExcel->UseVisualStyleBackColor = true;
			this->buttonExcel->Click += gcnew System::EventHandler(this, &DataForm::buttonEXCEL_Click);
			// 
			// tabControl1
			// 
			this->tabControl1->Controls->Add(this->tabPage1);
			this->tabControl1->Controls->Add(this->tabPage2);
			this->tabControl1->Location = System::Drawing::Point(0, 36);
			this->tabControl1->Name = L"tabControl1";
			this->tabControl1->SelectedIndex = 0;
			this->tabControl1->Size = System::Drawing::Size(2216, 527);
			this->tabControl1->TabIndex = 6;
			// 
			// tabPage1
			// 
			this->tabPage1->Controls->Add(this->Label_Data);
			this->tabPage1->Controls->Add(this->dataGridView);
			this->tabPage1->Controls->Add(this->buttonExcel);
			this->tabPage1->Location = System::Drawing::Point(4, 29);
			this->tabPage1->Name = L"tabPage1";
			this->tabPage1->Padding = System::Windows::Forms::Padding(3);
			this->tabPage1->Size = System::Drawing::Size(2208, 494);
			this->tabPage1->TabIndex = 0;
			this->tabPage1->Text = L"������";
			this->tabPage1->UseVisualStyleBackColor = true;
			// 
			// dataGridView
			// 
			this->dataGridView->AutoSizeColumnsMode = System::Windows::Forms::DataGridViewAutoSizeColumnsMode::ColumnHeader;
			this->dataGridView->AutoSizeRowsMode = System::Windows::Forms::DataGridViewAutoSizeRowsMode::AllCellsExceptHeaders;
			this->dataGridView->ColumnHeadersHeightSizeMode = System::Windows::Forms::DataGridViewColumnHeadersHeightSizeMode::AutoSize;
			this->dataGridView->Dock = System::Windows::Forms::DockStyle::Bottom;
			this->dataGridView->Location = System::Drawing::Point(3, 117);
			this->dataGridView->Name = L"dataGridView";
			this->dataGridView->RowHeadersWidth = 62;
			this->dataGridView->RowTemplate->Height = 28;
			this->dataGridView->Size = System::Drawing::Size(2202, 374);
			this->dataGridView->TabIndex = 6;
			// 
			// tabPage2
			// 
			this->tabPage2->Controls->Add(this->buttonBrowse);
			this->tabPage2->Controls->Add(this->textBoxExcelDirectory);
			this->tabPage2->Controls->Add(this->labelExcelDirectory);
			this->tabPage2->Location = System::Drawing::Point(4, 29);
			this->tabPage2->Name = L"tabPage2";
			this->tabPage2->Padding = System::Windows::Forms::Padding(3);
			this->tabPage2->Size = System::Drawing::Size(2208, 494);
			this->tabPage2->TabIndex = 1;
			this->tabPage2->Text = L"���������";
			this->tabPage2->UseVisualStyleBackColor = true;
			// 
			// buttonBrowse
			// 
			this->buttonBrowse->Location = System::Drawing::Point(734, 21);
			this->buttonBrowse->Name = L"buttonBrowse";
			this->buttonBrowse->Size = System::Drawing::Size(91, 36);
			this->buttonBrowse->TabIndex = 2;
			this->buttonBrowse->Text = L"�����...";
			this->buttonBrowse->UseVisualStyleBackColor = true;
			this->buttonBrowse->Click += gcnew System::EventHandler(this, &DataForm::buttonBrowse_Click);
			// 
			// textBoxExcelDirectory
			// 
			this->textBoxExcelDirectory->Location = System::Drawing::Point(268, 26);
			this->textBoxExcelDirectory->Name = L"textBoxExcelDirectory";
			this->textBoxExcelDirectory->Size = System::Drawing::Size(460, 26);
			this->textBoxExcelDirectory->TabIndex = 1;
			this->textBoxExcelDirectory->Text = L"D:\\SensorData\\";
			// 
			// labelExcelDirectory
			// 
			this->labelExcelDirectory->AutoSize = true;
			this->labelExcelDirectory->Location = System::Drawing::Point(40, 29);
			this->labelExcelDirectory->Name = L"labelExcelDirectory";
			this->labelExcelDirectory->Size = System::Drawing::Size(222, 20);
			this->labelExcelDirectory->TabIndex = 0;
			this->labelExcelDirectory->Text = L"���������� EXCEL �����:";
			// 
			// Label_Data
			// 
			this->Label_Data->AutoSize = true;
			this->Label_Data->Location = System::Drawing::Point(15, 72);
			this->Label_Data->Name = L"Label_Data";
			this->Label_Data->Size = System::Drawing::Size(169, 20);
			this->Label_Data->TabIndex = 7;
			this->Label_Data->Text = L"������ �� �������...";
			// 
			// DataForm
			// 
			this->AutoScaleDimensions = System::Drawing::SizeF(9, 20);
			this->AutoScaleMode = System::Windows::Forms::AutoScaleMode::Font;
			this->ClientSize = System::Drawing::Size(2216, 563);
			this->Controls->Add(this->tabControl1);
			this->Controls->Add(this->menuStrip1);
			this->MainMenuStrip = this->menuStrip1;
			this->Name = L"DataForm";
			this->Text = L"���� ������";
			this->menuStrip1->ResumeLayout(false);
			this->menuStrip1->PerformLayout();
			this->tabControl1->ResumeLayout(false);
			this->tabPage1->ResumeLayout(false);
			this->tabPage1->PerformLayout();
			(cli::safe_cast<System::ComponentModel::ISupportInitialize^>(this->dataGridView))->EndInit();
			this->tabPage2->ResumeLayout(false);
			this->tabPage2->PerformLayout();
			this->ResumeLayout(false);
			this->PerformLayout();

		}
#pragma endregion
	private:
		System::Void �����ToolStripMenuItem_Click(System::Object^ sender, System::EventArgs^ e);
		System::Void buttonEXCEL_Click(System::Object^ sender, System::EventArgs^ e);
		System::Void buttonBrowse_Click(System::Object^ sender, System::EventArgs^ e);
	private:
		int clientPort; // ���� �������

	public:
		property int ClientPort{
			int get() { return clientPort; }
			void set(int value) { clientPort = value; }
		}
	public:
		void SetData_TextValue(String^ text) {
			Label_Data->Text = text;
		};
		static void CreateAndShowDataFormInThread(std::queue<std::wstring>& messageQueue,
			std::mutex& mtx,
			std::condition_variable& cv);
		static void CloseForm(const std::wstring& guid);
		static DataForm^ GetFormByGuid(const std::wstring& guid);
		static void DelayedGarbageCollection(Object^ state);
		static void ParseBuffer(const char* buffer, size_t size);

		void InitializeDataTable();
		void ProjectServerW::DataForm::AddDataToTable(const char* buffer, size_t size) {	// ������ AddDataToTable ��� ����������, ������������ ���������� �������
			// ������� ���� ������ dataTable
			ProjectServerW::DataForm::AddDataToTable(buffer, size, this->dataTable);
		}
		void ProjectServerW::DataForm::AddDataToTable(const char* buffer, size_t size, System::Data::DataTable^ table);
		void ProjectServerW::DataForm::AddDataToTableThreadSafe(cli::array<System::Byte>^ buffer, int size, int port);
		void ProjectServerW::DataForm::AddDataToExcel();

		static cli::array<cli::array<String^>^>^ GetBitFieldNames() {
			static bool initialized = false;
			static gcroot<cli::array<cli::array<String^>^>^> names = nullptr;

			if (!initialized) {
				InitializeBitFieldNames(names);
				initialized = true;
			}

			return names;
		}

		void EnableButton();
		void ShowSuccess();
		void ShowError(String^ message);
		System::Void DataForm_FormClosed(Object^ sender, FormClosedEventArgs^ e);
		System::Void DataForm_HandleDestroyed(Object^ sender, EventArgs^ e);
	private: 
		void SaveSettings();
		void LoadSettings();
		void UpdateDirectoryTextBox(String^ path);
		// ���������� ������� �������� �����
		System::Void DataForm_FormClosing(System::Object^ sender, System::Windows::Forms::FormClosingEventArgs^ e);
		
		// ����� ��� ������������� ������������ �����
		static void InitializeBitFieldNames(gcroot<cli::array<cli::array<String^>^>^>& namesRef);

		void TriggerExcelExport() {
			// �������� ����������� ������ Excel
			if (buttonExcel->Enabled) {
				// ������������� ��������� ������� � Excel
				buttonEXCEL_Click(nullptr, nullptr);
			}
			else {
				// ������ ����������, ������������� ���� �������� � ��������� ������
				pendingExcelExport = true;

				// ������� ������, ���� �� ��� �� ������
				if (exportTimer == nullptr) {
					exportTimer = gcnew System::Windows::Forms::Timer();
					exportTimer->Interval = 500; // �������� ������ 500 ��
					exportTimer->Tick += gcnew EventHandler(this, &DataForm::CheckExcelButtonStatus);
				}

				// ��������� ������
				exportTimer->Start();

				// ������� ��������� ��� ������������
				Label_Data->Text = "�������� ����������� ������ � Excel...";
			}
		}

		// ���������� ������� �������
		void CheckExcelButtonStatus(Object^ sender, EventArgs^ e) {
			// ���������, �������� �� ������
			if (buttonExcel->Enabled && pendingExcelExport) {
				// ������������� ������
				exportTimer->Stop();

				// ���������� ����
				pendingExcelExport = false;

				// ��������� �������
				buttonEXCEL_Click(nullptr, nullptr);

				// ��������� �����
				Label_Data->Text = "������ ������������ � Excel...";
			}
		}
	};
}

// ������������� ����� ��� �������� �������
class ThreadStorage {
public:
	static void StoreThread(const std::wstring& guid, std::thread& thread);
	static void StopThread(const std::wstring& guid);
private:
	// ������� ��� ����������� ����������� ���������� Mutex ��� ������
	static std::mutex& GetMutex();
	// ������� ��� ����������� ����������� ���������� map ��� ������
	static std::map<std::wstring, std::thread>& GetThreadMap();
};

