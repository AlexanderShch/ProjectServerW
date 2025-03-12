#pragma once
#include "SServer.h"

#include <mutex>
#include <map>						// Для использования std::map - структуры, сохраняющей соответствие ID и ссылки на форму
#include <condition_variable>
#include <queue>
#include <vcclr.h>					// Для использования gcroot

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
	/// Сводка для DataForm
	/// </summary>
	public ref class DataForm : public System::Windows::Forms::Form
	{
	private:
		System::Data::DataTable^ dataTable;  // Объявление таблицы как члена класса
		//static int excelOperationCount = 0;	 // Переменная-счетчик для очистки кэша COM
		Thread^ excelThread;				 // Объявим объект для работы с Excel в отдельном потоке
		// Массив наименований битовых полей для каждого типа сенсора
		// Индекс первого уровня - тип сенсора, второго уровня - номер бита

	private: System::Windows::Forms::TabControl^ tabControl1;
	private: System::Windows::Forms::TabPage^ tabPage1;
	private: System::Windows::Forms::TabPage^ tabPage2;
	private: System::Windows::Forms::Button^ buttonBrowse;
	private: System::Windows::Forms::Label^ Label_Data;

	private: System::Windows::Forms::TextBox^ textBoxExcelDirectory;
	private: System::Windows::Forms::Label^ labelExcelDirectory;
	private: System::String^ excelSavePath;  // Для хранения пути
		     System::String^ excelFileName;   // Для хранения имени файла, связанного с битом "Work"
		     bool workBitDetected;            // Флаг для отслеживания активации бита "Work"
			 bool pendingExcelExport;		  // Флаг ожидания освобождения кнопки записи Excel
			 System::Windows::Forms::Timer^ exportTimer;	// Таймер для проверки освобождения кнопки

	private: System::Windows::Forms::DataGridView^ dataGridView;

	public:
		DataForm(void)
		{
			InitializeComponent();
			GetBitFieldNames();	// Инициализация имен битов (если еще не инициализированы)
			InitializeDataTable();

			// Подписка на событие закрытия формы
			//this->FormClosing += gcnew System::Windows::Forms::FormClosingEventHandler(this, &DataForm::DataForm_FormClosing);

			// Инициализируем путь сохранения из текстового поля
			excelSavePath = textBoxExcelDirectory->Text;
			// Имя файла будет сгенерировано позже, когда "Work" станет активным
			excelFileName = nullptr;
			workBitDetected = false;

			// Загружаем настройки сразу после инициализации компонентов
			LoadSettings();

		}

	protected:
		/// <summary>
		/// Освободить все используемые ресурсы.
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
	private: System::Windows::Forms::ToolStripMenuItem^ выходToolStripMenuItem;



	private: System::Windows::Forms::Button^ buttonExcel;

	private:
		/// <summary>
		/// Обязательная переменная конструктора.
		/// </summary>
		System::ComponentModel::Container^ components;

#pragma region Windows Form Designer generated code
		/// <summary>
		/// Требуемый метод для поддержки конструктора — не изменяйте 
		/// содержимое этого метода с помощью редактора кода.
		/// </summary>
		void InitializeComponent(void)
		{
			this->menuStrip1 = (gcnew System::Windows::Forms::MenuStrip());
			this->выходToolStripMenuItem = (gcnew System::Windows::Forms::ToolStripMenuItem());
			this->buttonExcel = (gcnew System::Windows::Forms::Button());
			this->tabControl1 = (gcnew System::Windows::Forms::TabControl());
			this->tabPage1 = (gcnew System::Windows::Forms::TabPage());
			this->dataGridView = (gcnew System::Windows::Forms::DataGridView());
			this->tabPage2 = (gcnew System::Windows::Forms::TabPage());
			this->buttonBrowse = (gcnew System::Windows::Forms::Button());
			this->textBoxExcelDirectory = (gcnew System::Windows::Forms::TextBox());
			this->labelExcelDirectory = (gcnew System::Windows::Forms::Label());
			Label_Data = (gcnew System::Windows::Forms::Label());
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
			this->menuStrip1->Items->AddRange(gcnew cli::array< System::Windows::Forms::ToolStripItem^  >(1) { this->выходToolStripMenuItem });
			this->menuStrip1->Location = System::Drawing::Point(0, 0);
			this->menuStrip1->Name = L"menuStrip1";
			this->menuStrip1->Size = System::Drawing::Size(2216, 33);
			this->menuStrip1->TabIndex = 0;
			this->menuStrip1->Text = L"menuStrip1";
			// 
			// выходToolStripMenuItem
			// 
			this->выходToolStripMenuItem->Name = L"выходToolStripMenuItem";
			this->выходToolStripMenuItem->Size = System::Drawing::Size(80, 29);
			this->выходToolStripMenuItem->Text = L"Выход";
			this->выходToolStripMenuItem->Click += gcnew System::EventHandler(this, &DataForm::выходToolStripMenuItem_Click);
			// 
			// Label_Data
			// 
			Label_Data->AutoSize = true;
			Label_Data->Location = System::Drawing::Point(17, 81);
			Label_Data->Name = L"Label_Data";
			Label_Data->Size = System::Drawing::Size(157, 20);
			Label_Data->TabIndex = 1;
			Label_Data->Text = L"Данные от клиента";
			// 
			// buttonExcel
			// 
			this->buttonExcel->Location = System::Drawing::Point(1869, 16);
			this->buttonExcel->Name = L"buttonExcel";
			this->buttonExcel->Size = System::Drawing::Size(207, 42);
			this->buttonExcel->TabIndex = 5;
			this->buttonExcel->Text = L"Запись в EXCEL";
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
			this->tabPage1->Controls->Add(this->buttonExcel);
			this->tabPage1->Controls->Add(Label_Data);
			this->tabPage1->Location = System::Drawing::Point(4, 29);
			this->tabPage1->Name = L"tabPage1";
			this->tabPage1->Padding = System::Windows::Forms::Padding(3);
			this->tabPage1->Size = System::Drawing::Size(2208, 494);
			this->tabPage1->TabIndex = 0;
			this->tabPage1->Text = L"Данные";
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
			this->tabPage2->Text = L"Настройки";
			this->tabPage2->UseVisualStyleBackColor = true;
			// 
			// buttonBrowse
			// 
			this->buttonBrowse->Location = System::Drawing::Point(734, 21);
			this->buttonBrowse->Name = L"buttonBrowse";
			this->buttonBrowse->Size = System::Drawing::Size(91, 36);
			this->buttonBrowse->TabIndex = 2;
			this->buttonBrowse->Text = L"Обзор...";
			this->buttonBrowse->UseVisualStyleBackColor = true;
			this->buttonBrowse->Click += gcnew System::EventHandler(this, &DataForm::buttonBrowse_Click);
			// 
			// textBoxExcelDirectory
			// 
			this->textBoxExcelDirectory->Location = System::Drawing::Point(268, 26);
			this->textBoxExcelDirectory->Name = L"textBoxExcelDirectory";
			this->textBoxExcelDirectory->Size = System::Drawing::Size(460, 26);
			this->textBoxExcelDirectory->TabIndex = 1;
			this->textBoxExcelDirectory->Text = L"D:\\";
			// 
			// labelExcelDirectory
			// 
			this->labelExcelDirectory->AutoSize = true;
			this->labelExcelDirectory->Location = System::Drawing::Point(40, 29);
			this->labelExcelDirectory->Name = L"labelExcelDirectory";
			this->labelExcelDirectory->Size = System::Drawing::Size(222, 20);
			this->labelExcelDirectory->TabIndex = 0;
			this->labelExcelDirectory->Text = L"Размещение EXCEL файла:";
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
			this->Text = L"Приём данных";
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
		System::Void выходToolStripMenuItem_Click(System::Object^ sender, System::EventArgs^ e);
		System::Void buttonEXCEL_Click(System::Object^ sender, System::EventArgs^ e);
		System::Void buttonBrowse_Click(System::Object^ sender, System::EventArgs^ e);
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
		void ProjectServerW::DataForm::AddDataToTable(const char* buffer, size_t size) {	// запуск AddDataToTable без параметров, используется внутренняя таблица
			// Передаём член класса dataTable
			ProjectServerW::DataForm::AddDataToTable(buffer, size, this->dataTable);
		}

		static cli::array<cli::array<String^>^>^ GetBitFieldNames() {
			static bool initialized = false;
			static gcroot<cli::array<cli::array<String^>^>^> names = nullptr;

			if (!initialized) {
				InitializeBitFieldNames(names);
				initialized = true;
			}

			return names;
		}

		void ProjectServerW::DataForm::AddDataToTable(const char* buffer, size_t size, System::Data::DataTable^ table);
		void ProjectServerW::DataForm::AddDataToExcel();
		void ProjectServerW::DataForm::AddDataToTableThreadSafe(cli::array<System::Byte>^ buffer, int size);

		void EnableButton();
		void ShowSuccess();
		void ShowError(String^ message);
	private: 
		void SaveSettings();
		void LoadSettings();
		// Обработчик события закрытия формы
		//System::Void DataForm_FormClosing(System::Object^ sender, System::Windows::Forms::FormClosingEventArgs^ e);
		// Методы для безопасного вызова BeginInvoke
		//void SafeBeginInvoke(MethodInvoker^ method);
		//void SafeBeginInvoke(System::Action<String^>^ method, String^ param);		void CloseFormAfterExport();
		
		// Метод для инициализации наименований битов
		static void InitializeBitFieldNames(gcroot<cli::array<cli::array<String^>^>^>& namesRef);

		void TriggerExcelExport() {
			// Проверка доступности кнопки Excel
			if (buttonExcel->Enabled) {
				// Автоматически запустить экспорт в Excel
				buttonEXCEL_Click(nullptr, nullptr);
			}
			else {
				// Кнопка недоступна, устанавливаем флаг ожидания и запускаем таймер
				pendingExcelExport = true;

				// Создаем таймер, если он еще не создан
				if (exportTimer == nullptr) {
					exportTimer = gcnew System::Windows::Forms::Timer();
					exportTimer->Interval = 500; // Проверка каждые 500 мс
					exportTimer->Tick += gcnew EventHandler(this, &DataForm::CheckExcelButtonStatus);
				}

				// Запускаем таймер
				exportTimer->Start();

				// Выводим сообщение для пользователя
				Label_Data->Text = "Ожидание возможности записи в Excel...";
			}
		}

		// Обработчик события таймера
		void CheckExcelButtonStatus(Object^ sender, EventArgs^ e) {
			// Проверяем, доступна ли кнопка
			if (buttonExcel->Enabled && pendingExcelExport) {
				// Останавливаем таймер
				exportTimer->Stop();

				// Сбрасываем флаг
				pendingExcelExport = false;

				// Запускаем экспорт
				buttonEXCEL_Click(nullptr, nullptr);

				// Обновляем метку
				Label_Data->Text = "Данные записываются в Excel...";
			}
		}
	};
}

// Неуправляемый класс для хранения потоков
class ThreadStorage {
public:
	static void StoreThread(const std::wstring& guid, std::thread& thread);
	static void StopThread(const std::wstring& guid);
private:
	// Функция для определения статической переменной Mutex для потока
	static std::mutex& GetMutex();
	// Функция для определения статической переменной map для потока
	static std::map<std::wstring, std::thread>& GetThreadMap();
};

