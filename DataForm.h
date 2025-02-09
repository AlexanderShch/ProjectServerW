#pragma once
#include "SServer.h"

#include <mutex>
#include <map>						// ��� ������������� std::map - ���������, ����������� ������������ ID � ������ �� �����
#include <condition_variable>
#include <queue>

#include <cstdlib> // Для system()
#include <sstream>

#define SQ 6				// sensors quantity for measures (0-4) + sets of T (5, 6) + MB_IO

namespace ProjectServerW {

	using namespace System;
	using namespace System::ComponentModel;
	using namespace System::Collections;
	using namespace System::Windows::Forms;
	using namespace System::Data;
	using namespace System::Drawing;

	/// <summary>
	/// ������ ��� DataForm
	/// </summary>
	public ref class DataForm : public System::Windows::Forms::Form
	{
	private:
		DataTable^ dataTable;  // ���������� ������� ��� ����� ������
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
	private: System::Windows::Forms::DataGridView^ dataGridView;


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
			this->dataGridView = (gcnew System::Windows::Forms::DataGridView());
			this->menuStrip1->SuspendLayout();
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
			this->menuStrip1->Size = System::Drawing::Size(1628, 36);
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
			this->Label_Data->Location = System::Drawing::Point(86, 119);
			this->Label_Data->Name = L"Label_Data";
			this->Label_Data->Size = System::Drawing::Size(157, 20);
			this->Label_Data->TabIndex = 1;
			this->Label_Data->Text = L"������ �� �������";
			// 
			// Label_ID
			// 
			this->Label_ID->AutoSize = true;
			this->Label_ID->Location = System::Drawing::Point(90, 67);
			this->Label_ID->Name = L"Label_ID";
			this->Label_ID->Size = System::Drawing::Size(67, 20);
			this->Label_ID->TabIndex = 2;
			this->Label_ID->Text = L"Form ID";
			// 
			// dataGridView
			// 
			this->dataGridView->AutoSizeColumnsMode = System::Windows::Forms::DataGridViewAutoSizeColumnsMode::ColumnHeader;
			this->dataGridView->AutoSizeRowsMode = System::Windows::Forms::DataGridViewAutoSizeRowsMode::AllCellsExceptHeaders;
			this->dataGridView->ColumnHeadersHeightSizeMode = System::Windows::Forms::DataGridViewColumnHeadersHeightSizeMode::AutoSize;
			this->dataGridView->Dock = System::Windows::Forms::DockStyle::Bottom;
			this->dataGridView->Location = System::Drawing::Point(0, 229);
			this->dataGridView->Name = L"dataGridView";
			this->dataGridView->RowHeadersWidth = 62;
			this->dataGridView->RowTemplate->Height = 28;
			this->dataGridView->Size = System::Drawing::Size(1628, 334);
			this->dataGridView->TabIndex = 3;
			// 
			// DataForm
			// 
			this->AutoScaleDimensions = System::Drawing::SizeF(9, 20);
			this->AutoScaleMode = System::Windows::Forms::AutoScaleMode::Font;
			this->ClientSize = System::Drawing::Size(1628, 563);
			this->Controls->Add(this->dataGridView);
			this->Controls->Add(this->Label_ID);
			this->Controls->Add(this->Label_Data);
			this->Controls->Add(this->menuStrip1);
			this->MainMenuStrip = this->menuStrip1;
			this->Name = L"DataForm";
			this->Text = L"���� ������";
			this->menuStrip1->ResumeLayout(false);
			this->menuStrip1->PerformLayout();
			(cli::safe_cast<System::ComponentModel::ISupportInitialize^>(this->dataGridView))->EndInit();
			this->ResumeLayout(false);
			this->PerformLayout();

		}
#pragma endregion
	private:
		System::Void �����ToolStripMenuItem_Click(System::Object^ sender, System::EventArgs^ e);

	public:
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
		void AddDataToTable(const char* buffer, size_t size) {	// ������ AddDataToTable ��� ����������, ������������ ���������� �������
			// ������� ���� ������ dataTable
			AddDataToTable(buffer, size, this->dataTable);
		}
		void AddDataToTable(const char* buffer, size_t size, DataTable^ table);
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

