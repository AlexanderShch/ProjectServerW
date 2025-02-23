#pragma once
#include "SServer.h"

#include <mutex>
#include <map>						// ��� ������������� std::map - ���������, ����������� ������������ ID � ������ �� �����
#include <condition_variable>
#include <queue>


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
		System::Data::DataTable^ dataTable;  // ���������� ������� ��� ����� ������
	private: System::Windows::Forms::TabControl^ tabControl1;
	private: System::Windows::Forms::TabPage^ tabPage1;
	private: System::Windows::Forms::TabPage^ tabPage2;
	private: System::Windows::Forms::DataGridView^ dataGridView;



		   Thread^ excelThread;				// ������� ������ ��� ������ � Excel � ��������� ������
	public:
		DataForm(void)
		{
			InitializeComponent();
			InitializeDataTable();
			//
			// TODO: �������� ��� ������������
			//
		}

	protected:
		/// <summary>
		/// ���������� ��� ������������ �������.
		/// </summary>
		~DataForm()
		{
			if (components)
			{
				delete components;
			}
		}
	private: System::Windows::Forms::MenuStrip^ menuStrip1;
	protected:
	private: System::Windows::Forms::ToolStripMenuItem^ �����ToolStripMenuItem;
	private: System::Windows::Forms::Label^ Label_Data;
	private: System::Windows::Forms::Label^ Label_ID;

	private: System::Windows::Forms::Label^ labelCRC;
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
			this->Label_Data = (gcnew System::Windows::Forms::Label());
			this->Label_ID = (gcnew System::Windows::Forms::Label());
			this->labelCRC = (gcnew System::Windows::Forms::Label());
			this->buttonExcel = (gcnew System::Windows::Forms::Button());
			this->tabControl1 = (gcnew System::Windows::Forms::TabControl());
			this->tabPage1 = (gcnew System::Windows::Forms::TabPage());
			this->dataGridView = (gcnew System::Windows::Forms::DataGridView());
			this->tabPage2 = (gcnew System::Windows::Forms::TabPage());
			this->menuStrip1->SuspendLayout();
			this->tabControl1->SuspendLayout();
			this->tabPage1->SuspendLayout();
			(cli::safe_cast<System::ComponentModel::ISupportInitialize^>(this->dataGridView))->BeginInit();
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
			// Label_Data
			// 
			this->Label_Data->AutoSize = true;
			this->Label_Data->Location = System::Drawing::Point(17, 81);
			this->Label_Data->Name = L"Label_Data";
			this->Label_Data->Size = System::Drawing::Size(157, 20);
			this->Label_Data->TabIndex = 1;
			this->Label_Data->Text = L"������ �� �������";
			// 
			// Label_ID
			// 
			this->Label_ID->AutoSize = true;
			this->Label_ID->Location = System::Drawing::Point(17, 38);
			this->Label_ID->Name = L"Label_ID";
			this->Label_ID->Size = System::Drawing::Size(67, 20);
			this->Label_ID->TabIndex = 2;
			this->Label_ID->Text = L"Form ID";
			// 
			// labelCRC
			// 
			this->labelCRC->AutoSize = true;
			this->labelCRC->Location = System::Drawing::Point(914, 27);
			this->labelCRC->Name = L"labelCRC";
			this->labelCRC->Size = System::Drawing::Size(43, 20);
			this->labelCRC->TabIndex = 4;
			this->labelCRC->Text = L"CRC";
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
			this->tabPage1->Controls->Add(this->dataGridView);
			this->tabPage1->Controls->Add(this->Label_ID);
			this->tabPage1->Controls->Add(this->buttonExcel);
			this->tabPage1->Controls->Add(this->Label_Data);
			this->tabPage1->Controls->Add(this->labelCRC);
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
			this->tabPage2->Location = System::Drawing::Point(4, 29);
			this->tabPage2->Name = L"tabPage2";
			this->tabPage2->Padding = System::Windows::Forms::Padding(3);
			this->tabPage2->Size = System::Drawing::Size(1541, 494);
			this->tabPage2->TabIndex = 1;
			this->tabPage2->Text = L"���������";
			this->tabPage2->UseVisualStyleBackColor = true;
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
			this->ResumeLayout(false);
			this->PerformLayout();

		}
#pragma endregion
	private:
		System::Void �����ToolStripMenuItem_Click(System::Object^ sender, System::EventArgs^ e);
		System::Void buttonEXCEL_Click(System::Object^ sender, System::EventArgs^ e);

	public:
		void SetData_CRC_Value(String^ text) {
			labelCRC->Text = text;
		};
		void SetData_TextValue(String^ text) {
			Label_Data->Text = text;
		};
		void SetData_FormID_value(String^ text) {
			Label_ID->Text = text;
		};
		static void CreateAndShowDataFormInThread(std::queue<std::wstring>& messageQueue,
			std::mutex& mtx,
			std::condition_variable& cv);
		static void CloseForm(const std::wstring& guid);
		static DataForm^ GetFormByGuid(const std::wstring& guid);
		
		static void ParseBuffer(const char* buffer, size_t size);

		void InitializeDataTable();
		void ProjectServerW::DataForm::AddDataToTable(const char* buffer, size_t size) {	// ������ AddDataToTable ��� ����������, ������������ ���������� �������
			// ������� ���� ������ dataTable
			ProjectServerW::DataForm::AddDataToTable(buffer, size, this->dataTable);
		}
		void ProjectServerW::DataForm::AddDataToTable(const char* buffer, size_t size, System::Data::DataTable^ table);
		void ProjectServerW::DataForm::AddDataToExcel();
	
		void EnableButton();

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

