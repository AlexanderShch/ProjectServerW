#pragma once
#include "SServer.h"

#include <mutex>
#include <map>						// Для использования std::map - структуры, сохраняющей соответствие ID и ссылки на форму
#include <condition_variable>
#include <queue>
#include <vcclr.h>					// Для использования gcroot

#define SQ 7				// датчики TH дефростера (0-2) + датчики продукта (3, 4) + T корпуса (5) + MB_IO (6)

// Forward declaration для избежания циклических зависимостей
struct Command;
struct CommandResponse;

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
			SOCKET clientSocket;  // Сокет клиента
			
		// Очередь для ответов от контроллера (статические для доступа из SServer)
		static System::Collections::Concurrent::ConcurrentQueue<cli::array<System::Byte>^>^ responseQueue;
		private: System::Windows::Forms::Label^ Label_Commands;
		private: System::Windows::Forms::Label^ label_Commands_Info;
		private: System::Windows::Forms::Button^ button_RESET;



			static System::Threading::Semaphore^ responseAvailable;
		
		// Статический конструктор для инициализации статических управляемых членов
		static DataForm() {
			responseQueue = gcnew System::Collections::Concurrent::ConcurrentQueue<cli::array<System::Byte>^>();
			responseAvailable = gcnew System::Threading::Semaphore(0, 100);
		}
		
		public:
			property SOCKET ClientSocket {
				SOCKET get() { return clientSocket; }
				void set(SOCKET value) { clientSocket = value; }
			}

		private:
			System::Data::DataTable^ dataTable;  // Объявление таблицы как члена класса
			Thread^ excelThread;				 // Объявим объект для работы с Excel в отдельном потоке
			// Массив наименований битовых полей для каждого типа сенсора
			// Индекс первого уровня - тип сенсора, второго уровня - номер бита
			ManualResetEvent^ exportCompletedEvent;
			bool exportSuccessful;
			cli::array<String^>^ sensorNames;					// Объявим массив с именами сенсоров
			cli::array<System::Drawing::Color>^ sensorColors;	// Объявим массив с цветом линии сенсора на графике

		private: System::Windows::Forms::TabControl^ tabControl1;
		private: System::Windows::Forms::TabPage^ tabPage1;
		private: System::Windows::Forms::TabPage^ tabPage2;
		private: System::Windows::Forms::Button^ buttonBrowse;

		private: System::Windows::Forms::TextBox^ textBoxExcelDirectory;
		private: System::Windows::Forms::Label^ labelExcelDirectory;
		private: System::String^ excelSavePath;  // Для хранения пути
				System::String^ excelFileName;   // Для хранения имени файла, связанного с битом "Work"
				DateTime dataCollectionStartTime; // Время начала сбора данных
				DateTime dataCollectionEndTime;   // Время окончания сбора данных
			bool workBitDetected;            // Флаг для отслеживания активации бита "Work"
			bool pendingExcelExport;		  // Флаг ожидания освобождения кнопки записи Excel
			System::Windows::Forms::Timer^ exportTimer;	// Таймер для проверки освобождения кнопки
		DateTime workBitZeroStartTime;   // время перехода бита Work в состояние ноль
		bool workBitZeroTimerActive;     // флаг "запущен таймер отслеживания активности таймера бита Work в нуле"
		bool firstDataReceived;          // флаг "получен первый пакет данных" (для установки начального состояния кнопок)
		private: System::Windows::Forms::Label^ Label_Data;
		private: System::Windows::Forms::Label^ LabelDefroster;
		private: System::Windows::Forms::Label^ T_def_left;
		private: System::Windows::Forms::Label^ T_def_right;
		private: System::Windows::Forms::Label^ T_def_center;
		private: System::Windows::Forms::Label^ T_product_right;
		private: System::Windows::Forms::Label^ T_product_left;
		private: System::Windows::Forms::Label^ LabelProduct;
		private: System::Windows::Forms::Button^ buttonSTOP;
		private: System::Windows::Forms::Button^ buttonSTART;
		private: System::Windows::Forms::Label^ labelSTOP;
		private: System::Windows::Forms::Label^ labelSTART;
		
		// Элементы автозапуска по времени
		private: System::Windows::Forms::CheckBox^ checkBoxAutoStart;
		private: System::Windows::Forms::DateTimePicker^ dateTimePickerAutoStart;
		private: System::Windows::Forms::Label^ labelAutoStart;
		private: System::Windows::Forms::Timer^ timerAutoStart;

		private: System::Windows::Forms::DataGridView^ dataGridView;

		public:
			DataForm(void)
			{
				InitializeComponent();
				GetBitFieldNames();	// Инициализация имен битов (если еще не инициализированы)
				InitializeDataTable();

				// Имена датчиков температуры для столбцов T0..T(SQ-2)
				sensorNames = gcnew cli::array<String^>(SQ - 1) {
					"дефр.Левый",
					"дефр.Правый",
					"дефр.Центр",
					"прод.Лев",
					"прод.Пр",
					"корпус"
				};

				// Цвета кривых для соответствующих датчиков
				sensorColors = gcnew cli::array<System::Drawing::Color>(SQ - 1) {
						System::Drawing::Color::Green,
						System::Drawing::Color::Red,
						System::Drawing::Color::Orange,
						System::Drawing::Color::DarkGreen,
						System::Drawing::Color::Purple,
						System::Drawing::Color::Brown
				};

				// Подписка на события при закрытии формы
				this->FormClosing += gcnew FormClosingEventHandler(this, &DataForm::DataForm_FormClosing);
				this->FormClosed += gcnew FormClosedEventHandler(this, &DataForm::DataForm_FormClosed);
				this->HandleDestroyed += gcnew EventHandler(this, &DataForm::DataForm_HandleDestroyed);
				// Инициализация объекта для синхронизации
				exportCompletedEvent = gcnew System::Threading::ManualResetEvent(false);

			// Инициализируем путь сохранения из текстового поля
			excelSavePath = textBoxExcelDirectory->Text;
			// Имя файла будет сгенерировано позже, когда "Work" станет активным
	excelFileName = nullptr;
	workBitDetected = false;
	workBitZeroTimerActive = false;
	firstDataReceived = false;      // Флаг первого приёма данных
	// Инициализация порта клиента
	clientPort = 0;

			// Загружаем настройки сразу после инициализации компонентов
			LoadSettings();

			// Состояние кнопок START/STOP будет установлено при получении первого пакета данных

		// ===== ОПТИМИЗАЦИЯ ПРОИЗВОДИТЕЛЬНОСТИ DataGridView =====
		// Отключаем автоматический пересчет размеров - это ОЧЕНЬ медленно при большом количестве строк
		dataGridView->AutoSizeColumnsMode = System::Windows::Forms::DataGridViewAutoSizeColumnsMode::None;
		dataGridView->AutoSizeRowsMode = System::Windows::Forms::DataGridViewAutoSizeRowsMode::None;
		
		// Устанавливаем ширину столбцов по умолчанию
		// Пользователь сможет изменить вручную, но не будет автоматического пересчета
		dataGridView->DefaultCellStyle->WrapMode = System::Windows::Forms::DataGridViewTriState::False;
		for each (System::Windows::Forms::DataGridViewColumn^ column in dataGridView->Columns) {
			if (column->Name == "RealTime") {
				column->Width = 60;  // Ширина для столбца времени (HH:mm:ss)
			} else {
				column->Width = 40;  // Ширина для остальных столбцов
			}
		}
		
		// Отключаем визуальные стили для увеличения скорости
		dataGridView->EnableHeadersVisualStyles = false;
		
		// Включаем двойную буферизацию через рефлексию (DoubleBuffered - защищённое свойство)
		System::Reflection::PropertyInfo^ pi = dataGridView->GetType()->GetProperty("DoubleBuffered",
			System::Reflection::BindingFlags::Instance | System::Reflection::BindingFlags::NonPublic);
		if (pi != nullptr) {
			pi->SetValue(dataGridView, true, nullptr);
		}
		
		// Устанавливаем виртуальный режим для больших таблиц (опционально)
		// dataGridView->VirtualMode = true;

		}

		protected:
			/// <summary>
			/// Освободить все используемые ресурсы.
			/// </summary>
			~DataForm()
			{
				// Вызов финализатора через деструктор
				this->!DataForm();

				if (components)
				{
					delete components;
					components = nullptr;
				}
			}

			// Финализатор
			DataForm::!DataForm()
			{
				try {
					// Очищаем все неуправляемые ресурсы
					if (clientSocket != INVALID_SOCKET) {
						closesocket(clientSocket);
						clientSocket = INVALID_SOCKET;
					}

					// Освобождаем COM-объекты
					// Примечание: это должно выполняться в потоке STA
					if (Thread::CurrentThread->GetApartmentState() == ApartmentState::STA) {
						// Освобождаем COM-объекты
					}
				}
				catch (...) {
					// Игнорируем исключения в финализаторе
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
				this->T_product_right = (gcnew System::Windows::Forms::Label());
				this->T_product_left = (gcnew System::Windows::Forms::Label());
				this->LabelProduct = (gcnew System::Windows::Forms::Label());
				this->T_def_right = (gcnew System::Windows::Forms::Label());
				this->T_def_center = (gcnew System::Windows::Forms::Label());
				this->T_def_left = (gcnew System::Windows::Forms::Label());
				this->LabelDefroster = (gcnew System::Windows::Forms::Label());
				this->Label_Data = (gcnew System::Windows::Forms::Label());
				this->dataGridView = (gcnew System::Windows::Forms::DataGridView());
				this->tabPage2 = (gcnew System::Windows::Forms::TabPage());
				this->Label_Commands = (gcnew System::Windows::Forms::Label());
				this->labelSTOP = (gcnew System::Windows::Forms::Label());
				this->labelSTART = (gcnew System::Windows::Forms::Label());
				this->buttonSTOP = (gcnew System::Windows::Forms::Button());
				this->buttonSTART = (gcnew System::Windows::Forms::Button());
				this->checkBoxAutoStart = (gcnew System::Windows::Forms::CheckBox());
				this->dateTimePickerAutoStart = (gcnew System::Windows::Forms::DateTimePicker());
				this->labelAutoStart = (gcnew System::Windows::Forms::Label());
				this->timerAutoStart = (gcnew System::Windows::Forms::Timer());
				this->buttonBrowse = (gcnew System::Windows::Forms::Button());
				this->textBoxExcelDirectory = (gcnew System::Windows::Forms::TextBox());
				this->labelExcelDirectory = (gcnew System::Windows::Forms::Label());
				this->label_Commands_Info = (gcnew System::Windows::Forms::Label());
				this->button_RESET = (gcnew System::Windows::Forms::Button());
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
				this->menuStrip1->Size = System::Drawing::Size(1494, 33);
				this->menuStrip1->TabIndex = 0;
				this->menuStrip1->Text = L"menuStrip1";
				this->menuStrip1->ItemClicked += gcnew System::Windows::Forms::ToolStripItemClickedEventHandler(this, &DataForm::menuStrip1_ItemClicked);
				// 
				// выходToolStripMenuItem
				// 
				this->выходToolStripMenuItem->Name = L"выходToolStripMenuItem";
				this->выходToolStripMenuItem->Size = System::Drawing::Size(80, 29);
				this->выходToolStripMenuItem->Text = L"Выход";
				this->выходToolStripMenuItem->Click += gcnew System::EventHandler(this, &DataForm::выходToolStripMenuItem_Click);
				// 
				// buttonExcel
				// 
				this->buttonExcel->Location = System::Drawing::Point(12, 6);
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
				this->tabControl1->Size = System::Drawing::Size(1494, 527);
				this->tabControl1->TabIndex = 6;
				// 
				// tabPage1
				// 
				this->tabPage1->AutoScroll = true;
				this->tabPage1->Controls->Add(this->T_product_right);
				this->tabPage1->Controls->Add(this->T_product_left);
				this->tabPage1->Controls->Add(this->LabelProduct);
				this->tabPage1->Controls->Add(this->T_def_right);
				this->tabPage1->Controls->Add(this->T_def_center);
				this->tabPage1->Controls->Add(this->T_def_left);
				this->tabPage1->Controls->Add(this->LabelDefroster);
				this->tabPage1->Controls->Add(this->Label_Data);
				this->tabPage1->Controls->Add(this->dataGridView);
				this->tabPage1->Controls->Add(this->buttonExcel);
				this->tabPage1->Location = System::Drawing::Point(4, 29);
				this->tabPage1->Name = L"tabPage1";
				this->tabPage1->Padding = System::Windows::Forms::Padding(3);
				this->tabPage1->Size = System::Drawing::Size(1486, 494);
				this->tabPage1->TabIndex = 0;
				this->tabPage1->Text = L"Данные";
				this->tabPage1->UseVisualStyleBackColor = true;
				// 
				// T_product_right
				// 
				this->T_product_right->AutoSize = true;
				this->T_product_right->Location = System::Drawing::Point(858, 17);
				this->T_product_right->Name = L"T_product_right";
				this->T_product_right->Size = System::Drawing::Size(23, 20);
				this->T_product_right->TabIndex = 13;
				this->T_product_right->Text = L"- -";
				this->T_product_right->TextAlign = System::Drawing::ContentAlignment::MiddleCenter;
				// 
				// T_product_left
				// 
				this->T_product_left->AutoSize = true;
				this->T_product_left->Location = System::Drawing::Point(787, 17);
				this->T_product_left->Name = L"T_product_left";
				this->T_product_left->Size = System::Drawing::Size(23, 20);
				this->T_product_left->TabIndex = 13;
				this->T_product_left->Text = L"- -";
				this->T_product_left->TextAlign = System::Drawing::ContentAlignment::MiddleCenter;
				// 
				// LabelProduct
				// 
				this->LabelProduct->AutoSize = true;
				this->LabelProduct->Location = System::Drawing::Point(667, 17);
				this->LabelProduct->Name = L"LabelProduct";
				this->LabelProduct->Size = System::Drawing::Size(93, 20);
				this->LabelProduct->TabIndex = 12;
				this->LabelProduct->Text = L"Т продукта";
				// 
				// T_def_right
				// 
				this->T_def_right->AutoSize = true;
				this->T_def_right->Location = System::Drawing::Point(551, 17);
				this->T_def_right->Name = L"T_def_right";
				this->T_def_right->Size = System::Drawing::Size(23, 20);
				this->T_def_right->TabIndex = 11;
				this->T_def_right->Text = L"- -";
				this->T_def_right->TextAlign = System::Drawing::ContentAlignment::MiddleCenter;
				// 
				// T_def_center
				// 
				this->T_def_center->AutoSize = true;
				this->T_def_center->Location = System::Drawing::Point(469, 17);
				this->T_def_center->Name = L"T_def_center";
				this->T_def_center->Size = System::Drawing::Size(23, 20);
				this->T_def_center->TabIndex = 10;
				this->T_def_center->Text = L"- -";
				this->T_def_center->TextAlign = System::Drawing::ContentAlignment::MiddleCenter;
				// 
				// T_def_left
				// 
				this->T_def_left->AutoSize = true;
				this->T_def_left->Location = System::Drawing::Point(389, 17);
				this->T_def_left->Name = L"T_def_left";
				this->T_def_left->Size = System::Drawing::Size(23, 20);
				this->T_def_left->TabIndex = 9;
				this->T_def_left->Text = L"- -";
				this->T_def_left->TextAlign = System::Drawing::ContentAlignment::MiddleCenter;
				// 
				// LabelDefroster
				// 
				this->LabelDefroster->AutoSize = true;
				this->LabelDefroster->Location = System::Drawing::Point(248, 17);
				this->LabelDefroster->Name = L"LabelDefroster";
				this->LabelDefroster->Size = System::Drawing::Size(119, 20);
				this->LabelDefroster->TabIndex = 8;
				this->LabelDefroster->Text = L"Т дефростера";
				// 
				// Label_Data
				// 
				this->Label_Data->AutoSize = true;
				this->Label_Data->Location = System::Drawing::Point(8, 51);
				this->Label_Data->Name = L"Label_Data";
				this->Label_Data->Size = System::Drawing::Size(169, 20);
				this->Label_Data->TabIndex = 7;
				this->Label_Data->Text = L"Данные от клиента...";
				// 
				// dataGridView
				// 
				this->dataGridView->AutoSizeColumnsMode = System::Windows::Forms::DataGridViewAutoSizeColumnsMode::AllCells;
				this->dataGridView->AutoSizeRowsMode = System::Windows::Forms::DataGridViewAutoSizeRowsMode::AllCells;
				this->dataGridView->ColumnHeadersHeightSizeMode = System::Windows::Forms::DataGridViewColumnHeadersHeightSizeMode::AutoSize;
				this->dataGridView->Location = System::Drawing::Point(12, 74);
				this->dataGridView->Name = L"dataGridView";
				this->dataGridView->RightToLeft = System::Windows::Forms::RightToLeft::No;
				this->dataGridView->RowHeadersWidthSizeMode = System::Windows::Forms::DataGridViewRowHeadersWidthSizeMode::AutoSizeToDisplayedHeaders;
			this->dataGridView->RowTemplate->Height = 20;  // Уменьшена высота строк с 28 до 20
			this->dataGridView->RowTemplate->Resizable = System::Windows::Forms::DataGridViewTriState::True;
			this->dataGridView->ScrollBars = System::Windows::Forms::ScrollBars::Both;  // Горизонтальный И вертикальный
				this->dataGridView->Size = System::Drawing::Size(1276, 400);
				this->dataGridView->TabIndex = 6;
				this->dataGridView->CellContentClick += gcnew System::Windows::Forms::DataGridViewCellEventHandler(this, &DataForm::dataGridView_CellContentClick);
				// 
				// tabPage2
				// 
				this->tabPage2->Controls->Add(this->button_RESET);
				this->tabPage2->Controls->Add(this->label_Commands_Info);
				this->tabPage2->Controls->Add(this->Label_Commands);
				this->tabPage2->Controls->Add(this->labelSTOP);
				this->tabPage2->Controls->Add(this->labelSTART);
				this->tabPage2->Controls->Add(this->buttonSTOP);
				this->tabPage2->Controls->Add(this->buttonSTART);
				this->tabPage2->Controls->Add(this->checkBoxAutoStart);
				this->tabPage2->Controls->Add(this->dateTimePickerAutoStart);
				this->tabPage2->Controls->Add(this->labelAutoStart);
				this->tabPage2->Controls->Add(this->buttonBrowse);
				this->tabPage2->Controls->Add(this->textBoxExcelDirectory);
				this->tabPage2->Controls->Add(this->labelExcelDirectory);
				this->tabPage2->Location = System::Drawing::Point(4, 29);
				this->tabPage2->Name = L"tabPage2";
				this->tabPage2->Padding = System::Windows::Forms::Padding(3);
				this->tabPage2->Size = System::Drawing::Size(1486, 494);
				this->tabPage2->TabIndex = 1;
				this->tabPage2->Text = L"Настройки";
				this->tabPage2->UseVisualStyleBackColor = true;
				// 
				// Label_Commands
				// 
				this->Label_Commands->AutoSize = true;
				this->Label_Commands->Location = System::Drawing::Point(246, 449);
				this->Label_Commands->Name = L"Label_Commands";
				this->Label_Commands->Size = System::Drawing::Size(138, 20);
				this->Label_Commands->TabIndex = 7;
				this->Label_Commands->Text = L"Команда не отправлялась";
				// 
				// labelSTOP
				// 
				this->labelSTOP->AutoSize = true;
				this->labelSTOP->BackColor = System::Drawing::Color::Snow;
				this->labelSTOP->Location = System::Drawing::Point(145, 135);
				this->labelSTOP->Name = L"labelSTOP";
				this->labelSTOP->Size = System::Drawing::Size(18, 20);
				this->labelSTOP->TabIndex = 6;
				this->labelSTOP->Text = L"0";
				// 
				// labelSTART
				// 
				this->labelSTART->AutoSize = true;
				this->labelSTART->BackColor = System::Drawing::Color::Snow;
				this->labelSTART->Location = System::Drawing::Point(145, 89);
				this->labelSTART->Name = L"labelSTART";
				this->labelSTART->Size = System::Drawing::Size(18, 20);
				this->labelSTART->TabIndex = 5;
				this->labelSTART->Text = L"0";
				// 
				// buttonSTOP
				// 
				this->buttonSTOP->Location = System::Drawing::Point(48, 129);
				this->buttonSTOP->Name = L"buttonSTOP";
				this->buttonSTOP->Size = System::Drawing::Size(91, 33);
				this->buttonSTOP->TabIndex = 4;
				this->buttonSTOP->Text = L"СТОП";
				this->buttonSTOP->UseVisualStyleBackColor = true;
				this->buttonSTOP->Click += gcnew System::EventHandler(this, &DataForm::buttonSTOP_Click);
				// 
				// buttonSTART
				// 
				this->buttonSTART->Location = System::Drawing::Point(48, 82);
				this->buttonSTART->Name = L"buttonSTART";
				this->buttonSTART->Size = System::Drawing::Size(91, 33);
				this->buttonSTART->TabIndex = 3;
				this->buttonSTART->Text = L"ПУСК";
				this->buttonSTART->UseVisualStyleBackColor = true;
				this->buttonSTART->Click += gcnew System::EventHandler(this, &DataForm::buttonSTART_Click);
				// 
				// labelAutoStart
				// 
				this->labelAutoStart->AutoSize = true;
				this->labelAutoStart->Location = System::Drawing::Point(48, 185);
				this->labelAutoStart->Name = L"labelAutoStart";
				this->labelAutoStart->Size = System::Drawing::Size(150, 20);
				this->labelAutoStart->TabIndex = 9;
				this->labelAutoStart->Text = L"Автозапуск в:";
				// 
				// checkBoxAutoStart
				// 
				this->checkBoxAutoStart->AutoSize = true;
				this->checkBoxAutoStart->Location = System::Drawing::Point(48, 215);
				this->checkBoxAutoStart->Name = L"checkBoxAutoStart";
				this->checkBoxAutoStart->Size = System::Drawing::Size(100, 24);
				this->checkBoxAutoStart->TabIndex = 10;
				this->checkBoxAutoStart->Text = L"Включен";
				this->checkBoxAutoStart->UseVisualStyleBackColor = true;
				this->checkBoxAutoStart->CheckedChanged += gcnew System::EventHandler(this, &DataForm::checkBoxAutoStart_CheckedChanged);
				// 
				// dateTimePickerAutoStart
				// 
				this->dateTimePickerAutoStart->Format = System::Windows::Forms::DateTimePickerFormat::Custom;
				this->dateTimePickerAutoStart->CustomFormat = L"HH:mm";
				this->dateTimePickerAutoStart->Location = System::Drawing::Point(160, 215);
				this->dateTimePickerAutoStart->Name = L"dateTimePickerAutoStart";
				this->dateTimePickerAutoStart->ShowUpDown = true;
				this->dateTimePickerAutoStart->Size = System::Drawing::Size(100, 26);
				this->dateTimePickerAutoStart->TabIndex = 11;
				this->dateTimePickerAutoStart->Value = System::DateTime::Now;
				// 
				// timerAutoStart
				// 
				this->timerAutoStart->Interval = 30000; // Проверка каждые 30 секунд
				this->timerAutoStart->Tick += gcnew System::EventHandler(this, &DataForm::timerAutoStart_Tick);	// Подписывает метод timerAutoStart_Tick на событие Tick таймера.
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
				this->textBoxExcelDirectory->Text = L"С:\\SensorData\\";
				this->textBoxExcelDirectory->TextChanged += gcnew System::EventHandler(this, &DataForm::textBoxExcelDirectory_TextChanged);
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
				// label_Commands_Info
				// 
				this->label_Commands_Info->AutoSize = true;
				this->label_Commands_Info->Location = System::Drawing::Point(44, 449);
				this->label_Commands_Info->Name = L"label_Commands_Info";
				this->label_Commands_Info->Size = System::Drawing::Size(196, 20);
				this->label_Commands_Info->TabIndex = 8;
				this->label_Commands_Info->Text = L"Информация о команде:";
				// 
				// button_RESET
				// 
				this->button_RESET->Location = System::Drawing::Point(48, 395);
				this->button_RESET->Name = L"button_RESET";
				this->button_RESET->RightToLeft = System::Windows::Forms::RightToLeft::Yes;
				this->button_RESET->Size = System::Drawing::Size(91, 33);
				this->button_RESET->TabIndex = 9;
				this->button_RESET->Text = L"СБРОС";
				this->button_RESET->UseVisualStyleBackColor = true;
				this->button_RESET->Click += gcnew System::EventHandler(this, &DataForm::button_RESET_Click);
				// 
				// DataForm
				// 
				this->AutoScaleDimensions = System::Drawing::SizeF(9, 20);
				this->AutoScaleMode = System::Windows::Forms::AutoScaleMode::Font;
				this->AutoScroll = true;
				this->AutoSizeMode = System::Windows::Forms::AutoSizeMode::GrowAndShrink;
				this->ClientSize = System::Drawing::Size(1317, 575);
				this->Controls->Add(this->tabControl1);
				this->Controls->Add(this->menuStrip1);
				this->MainMenuStrip = this->menuStrip1;
				this->Name = L"DataForm";
				this->RightToLeft = System::Windows::Forms::RightToLeft::No;
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
		private:
			int clientPort; // Порт клиента
			String^ clientIP;   // IP-адрес клиента
		private:			// Таймер для отложенной записи EXCEL
			System::Windows::Forms::Timer^ delayedExcelTimer;
			void OnDelayedExcelTimerTick(Object^ sender, EventArgs^ e);

		public:
			property int ClientPort{
				int get() { return clientPort; }
				void set(int value) { clientPort = value; }
			}
			property String^ ClientIP{
				String ^ get() { return clientIP; }
				void set(String ^ value) { clientIP = value; }
			}
		public:
			void SetData_TextValue(String^ text) {
				Label_Data->Text = text;
			};
			void SetT_def_left_Value(String^ text) {
				T_def_left->Text = text;
			};
			void SetT_def_right_Value(String^ text) {
				T_def_right->Text = text;
			};
			void SetT_def_center_Value(String^ text) {
				T_def_center->Text = text;
			};
			void SetT_product_left_Value(String^ text) {
				T_product_left->Text = text;
			};
			void SetT_product_right_Value(String^ text) {
				T_product_right->Text = text;
			};
			// Метод для обновления всех значений температуры
			void UpdateAllTemperatureValues(cli::array<double>^ temperatures) {
				if (temperatures->Length >= 5) {
					SetT_def_left_Value(temperatures[0].ToString("F1") + "°C");
					SetT_def_right_Value(temperatures[1].ToString("F1") + "°C");
					SetT_def_center_Value(temperatures[2].ToString("F1") + "°C");
					SetT_product_left_Value(temperatures[3].ToString("F1") + "°C");
					SetT_product_right_Value(temperatures[4].ToString("F1") + "°C");
				}
			};

			static void CreateAndShowDataFormInThread(std::queue<std::wstring>& messageQueue,
				std::mutex& mtx,
				std::condition_variable& cv);
		static void CloseForm(const std::wstring& guid);
		static DataForm^ GetFormByGuid(const std::wstring& guid);
		static std::wstring FindFormByClientIP(String^ clientIP); // найти форму по IP-адресу клиента
		static void DelayedGarbageCollection(Object^ state);
			static void ParseBuffer(const char* buffer, size_t size);

			void InitializeDataTable();
			void ProjectServerW::DataForm::AddDataToTable(const char* buffer, size_t size) {	// запуск AddDataToTable без параметров, используется внутренняя таблица
				// Передаём член класса dataTable
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
			bool SendCommand(const struct Command& cmd); // Универсальный метод отправки команды (имя определяется автоматически)
			bool SendCommand(const struct Command& cmd, System::String^ commandName); // Универсальный метод отправки команды с явным именем
			void SendStartCommand(); // Метод для отправки команды START клиенту
			void SendStopCommand(); // Метод для отправки команды STOP клиенту
			void SendResetCommand(); // Метод для отправки команды RESET клиенту
			
			// Методы для обработки ответов от контроллера
			static void EnqueueResponse(cli::array<System::Byte>^ response); // Добавление ответа в очередь (вызывается из SServer)
			bool ReceiveResponse(struct CommandResponse& response, int timeoutMs); // Прием ответа от контроллера с таймаутом
			bool ReceiveResponse(struct CommandResponse& response); // Прием ответа от контроллера (таймаут по умолчанию 1000 мс)
			void ProcessResponse(const struct CommandResponse& response); // Обработка полученного ответа
			void RestoreLabelCommandsColor(System::Object^ sender, System::EventArgs^ e); // Восстановление цвета Label_Commands
			bool SendCommandAndWaitResponse(const struct Command& cmd, struct CommandResponse& response, System::String^ commandName); // Отправка команды и ожидание ответа с именем
			bool SendCommandAndWaitResponse(const struct Command& cmd, struct CommandResponse& response); // Отправка команды и ожидание ответа (имя автоматически)
			
			void buttonSTARTstate_TRUE();	// Метод для изменения статуса кнопок Старт и Стоп
			void buttonSTOPstate_TRUE();	// Метод для изменения статуса кнопок Старт и Стоп
			System::Void DataForm_FormClosed(Object^ sender, FormClosedEventArgs^ e);
			System::Void DataForm_HandleDestroyed(Object^ sender, EventArgs^ e);
		private: 
			void SaveSettings();
			void LoadSettings();
			void UpdateDirectoryTextBox(String^ path);
			// Обработчик события закрытия формы
			System::Void DataForm_FormClosing(System::Object^ sender, System::Windows::Forms::FormClosingEventArgs^ e);
			
			// Метод для инициализации наименований битов
			static void InitializeBitFieldNames(gcroot<cli::array<cli::array<String^>^>^>& namesRef);

			void TriggerExcelExport();
			void CheckExcelButtonStatus(Object^ sender, EventArgs^ e);
		private: System::Void textBoxExcelDirectory_TextChanged(System::Object^ sender, System::EventArgs^ e) {
		}
	private: System::Void dataGridView_CellContentClick(System::Object^ sender, System::Windows::Forms::DataGridViewCellEventArgs^ e) {
	}
	private: System::Void menuStrip1_ItemClicked(System::Object^ sender, System::Windows::Forms::ToolStripItemClickedEventArgs^ e) {
	}
	private: System::Void записьВExcelToolStripMenuItem_Click(System::Object^ sender, System::EventArgs^ e) {
	}
	//private: System::Void label1_Click(System::Object^ sender, System::EventArgs^ e) {
	//}
	//private: System::Void label1_Click_1(System::Object^ sender, System::EventArgs^ e) {
	//}
	private: System::Void buttonSTART_Click(System::Object^ sender, System::EventArgs^ e) {
		// Формируем команду "START" для отправки клиенту
		SendStartCommand();
	}
	private: System::Void buttonSTOP_Click(System::Object^ sender, System::EventArgs^ e) {
		// Формируем команду "STOP" для отправки клиенту
		SendStopCommand();
	}
	private: System::Void button_RESET_Click(System::Object^ sender, System::EventArgs^ e) {
		// Формируем команду "RESET" для отправки клиенту
		SendResetCommand();
	}
	
	// ====================================================================
	// Обработчики автозапуска по времени (объявления)
	// ====================================================================
	
	private: System::Void checkBoxAutoStart_CheckedChanged(System::Object^ sender, System::EventArgs^ e);
	private: System::Void timerAutoStart_Tick(System::Object^ sender, System::EventArgs^ e);
	private: System::Void RestoreAutoStartColor(System::Object^ sender, System::EventArgs^ e);
	
};  // Конец класса DataForm

// Неуправляемый класс для хранения потоков (вне управляемого класса DataForm)
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

}  // Конец namespace ProjectServerW
