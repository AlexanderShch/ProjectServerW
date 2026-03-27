#pragma once
#include "Commands.h"   // первым: Command, CommandResponse, DefrostParamValue — полные определения до любых заголовков, использующих их в объявлениях
#include "SServer.h"

#include <mutex>
#include <map>						// Для использования std::map - структуры, сохраняющей соответствие ID и ссылки на форму
#include <condition_variable>
#include <queue>
#include <vcclr.h>					// Для использования gcroot

// Количество датчиков: TH дефростера (0-2) + датчики продукта (3, 4) + T корпуса (5) + MB_IO (6)
constexpr uint8_t SQ = 7;

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
			// Критично: нужно для корректной остановки/ожидания потока формы, когда форма закрывается сама (например, по таймауту).
			System::String^ formGuid;

			SOCKET clientSocket;  // Сокет клиента
			
		// Очередь/семафор ответов на команды — строго на уровне формы (на уровне устройства).
		// Причина: при нескольких устройствах общий static-буфер приводит к перемешиванию ответов между формами.
		System::Collections::Concurrent::ConcurrentQueue<cli::array<System::Byte>^>^ responseQueue;
		private: System::Windows::Forms::Label^ Label_Commands;
		private: System::Windows::Forms::Label^ label_Commands_Info;
		private: System::Windows::Forms::Button^ button_RESET;
	private: System::Windows::Forms::Label^ label_Version;

	private: System::Windows::Forms::Label^ labelVersion;
	private: System::Windows::Forms::TabPage^ tabPage3;
	private: System::Windows::Forms::DataGridView^ dataGridView1;
	private: System::Windows::Forms::DataGridViewTextBoxColumn^ Parameter;
	private: System::Windows::Forms::DataGridViewTextBoxColumn^ WarmUP;
	private: System::Windows::Forms::DataGridViewTextBoxColumn^ Plateau;
	private: System::Windows::Forms::DataGridViewTextBoxColumn^ Finish;
	private: System::Windows::Forms::DataGridView^ dataGridView2;
	private: System::Windows::Forms::DataGridViewTextBoxColumn^ Parameter2;
	private: System::Windows::Forms::DataGridViewTextBoxColumn^ Description;
	private: System::Windows::Forms::DataGridViewTextBoxColumn^ Value;
	private: System::Windows::Forms::Button^ buttonWriteParameters;
	private: System::Windows::Forms::Button^ buttonReadParameters;
	private: System::Windows::Forms::Label^ label1;


	private: System::Windows::Forms::Button^ buttonLoadFromFile;

	private: System::Windows::Forms::Label^ label2;





			System::Threading::Semaphore^ responseAvailable;
			// Previous tab for tabControl1 (used when leaving tabPage3 to ask save)
			System::Windows::Forms::TabPage^ tabControl1PrevTab;
			bool dataGridView1Dirty;  // true if dataGridView1 was edited and not saved
	private: System::Windows::Forms::Label^ label7;
	private: System::Windows::Forms::Label^ label6;
	private: System::Windows::Forms::Label^ label5;
	private: System::Windows::Forms::Label^ label4;
	private: System::Windows::Forms::Label^ label3;
		   bool dataGridView2Dirty;
	private: System::Windows::Forms::TabPage^ tabPage4;
	private: System::Windows::Forms::DataGridView^ dataGridEquipmentAlarm;
	private: System::Windows::Forms::Button^ buttonCheckAlarm;


		   // True only after both tables (group5->table1 and group6->table2) are loaded from controller.
			bool paramsLoadedFromDevice;
	private: System::Windows::Forms::Label^ label_Defroster_Info;

	private: System::Windows::Forms::Label^ labelDefrosterState;
		   bool startupSequenceCompleted;
		
		public:
			property System::String^ FormGuid {
				System::String^ get() { return formGuid; }
				void set(System::String^ value) { formGuid = value; }
			}

			// Вызывается при переподключении клиента (reuse формы по IP).
			// Важно: запускать только в UI-потоке формы (через BeginInvoke), т.к. внутри есть UI и синхронное ожидание ответа.
			void OnReconnectSendStartupCommands();
			bool ExecuteStartupCommandSequence();
			// Запуск отложенной инициализации после реконнекта (public wrapper for SServer).
			void ScheduleDeferredStartupOnReconnect();

			property SOCKET ClientSocket {
				SOCKET get() { return clientSocket; }
				void set(SOCKET value) {
					// Важно: setter вызывается из разных потоков (recv, telemetry worker, UI).
					// Поэтому здесь нельзя трогать UI (Timer/Close/Invoke). Только фиксируем состояние соединения.
					const SOCKET prev = clientSocket;
					clientSocket = value;
					if (value == INVALID_SOCKET || value != prev) {
						startupSequenceCompleted = false;
					}

					// Если сокет потерян — запоминаем момент, от которого считаем 30 минут ожидания "своего" клиента.
					// Если сокет снова валиден — сбрасываем ожидание.
					if (value == INVALID_SOCKET) {
						if (prev != INVALID_SOCKET) {
							System::Threading::Interlocked::Exchange(disconnectedSinceTicks, System::DateTime::Now.Ticks);
						}
					}
					else {
						System::Threading::Interlocked::Exchange(disconnectedSinceTicks, 0);
					}
				}
			}

		private:
			System::Data::DataTable^ dataTable;  // Объявление таблицы как члена класса
			// Critical: DataTable is not thread-safe; lock when copying for Excel while UI is appending rows.
			System::Object^ dataTableSync;
			// Строка, накопленная телеметрией; после прихода лога дополняется параметрами и при _Wrk=1 добавляется в таблицу.
			System::Data::DataRow^ pendingRow;
			bool pendingRowWrkBit;  // подтверждённое состояние авторежима для этой строки
			// Critical: explicit lock object for export state.
			System::Object^ excelExportSync;
			// Единый gate конвейера команд: в каждый момент времени только одна команда "send+wait".
			System::Object^ commandPipelineGate;
			// Critical: prevents starting multiple Excel export threads per form (they would only stack up on the global mutex).
			int excelExportInProgress;
			Thread^ excelThread;				 // Объявим объект для работы с Excel в отдельном потоке
			// Critical: tracks last telemetry receipt; used to decide whether to keep the form alive on reconnects.
			DateTime lastTelemetryTime;
			uint16_t lastTelemetryDeviceSeconds; // Why: map device time (seconds since boot) to wall-clock for command audit timestamps.
			bool hasTelemetry;
			bool inactivityCloseRequested;
			DateTime lastCmdInfoProbeTime; // Why: rate-limit automatic GET_CMD_INFO probes.
			bool cmdInfoProbeInProgress;  // Why: avoid recursive probing when GET_CMD_INFO itself times out.
			SOCKET lastTelemetrySocket;        // Why: detect reconnects (new TCP socket) and log data collection resume once per reconnect.
			bool reconnectFixationLogPending;  // Why: log "start of data collection after reconnect" on the first valid telemetry packet.
			System::Windows::Forms::Timer^ inactivityTimer;
			// Время, когда форма перешла в режим ожидания переподключения (после потери сокета).
			// Храним ticks для безопасного чтения/записи через Interlocked из разных потоков.
			long long disconnectedSinceTicks;
			// Защита от повторной отправки стартовых команд на одном и том же TCP-сокете.
			SOCKET lastStartupCommandsSocket;
			// Инициализация после RESET: контроллер перезагружается и некоторое время не читает UART4.
			// Поэтому отправляем GET_VERSION + SET_INTERVAL с задержкой и повторами, пока не начнёт отвечать.
			bool postResetInitPending;
			System::DateTime postResetInitDeadline;
			int postResetInitAttempt;
			System::Windows::Forms::Timer^ postResetInitTimer;
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
				System::String^ excelFileName;   // Имя файла для экспорта данных (формируется при первом приёме)
				DateTime dataCollectionStartTime; // Время начала сбора данных
				DateTime dataCollectionEndTime;   // Время окончания сбора данных
	bool firstDataReceived;          // флаг "получен первый пакет данных"
	bool dataExportedToExcel;        // флаг "данные уже были экспортированы в Excel" (для предотвращения дублирующей записи при закрытии формы)
	bool autoRestartPending;         // true: отправлен STOP по автоперезапуску, ждём экспорт и затем START
		DateTime autoRestartStopIssuedTime; // Время отправки STOP для таймаута автоперезапуска
	DateTime lastStopSuccessTime; // Время последнего успешного ответа на СТОП; лог не перезаписывает кнопки в течение 10 с
	DateTime lastStartSuccessTime; // Время последнего успешного START; защита от ложного "останов" сразу после запуска
	bool controllerAutoModeActive;   // true: в телеметрии бит _Wrk (DO) == 1 — контроллер в автоматическом режиме
	int wrkZeroConsecutiveCounts;    // Подряд идущие "нулевые" отсчёты _Wrk (по Time устройства)
	bool wrkLastSampleValid;         // true: есть предыдущий отсчёт Time для расчёта дельты
	uint16_t wrkLastSampleTime;      // Предыдущий Time устройства для расчёта дельты отсчётов
	DateTime lastControlLogTime;     // Время последней телеметрии с _Wrk=1; сброс флага при отсутствии такой телеметрии
	int controlLogAbsenceStrikeCount; // Счётчик подряд идущих "пропусков" _Wrk=1 для защиты от ложных срабатываний
	float lastFishCold_C;            // Последняя мин. Т рыбы из лога параметров (для записи в лог при остановке алгоритма)
	bool lastFishCold_C_Valid;       // true: lastFishCold_C получен из лога
	float lastActiveProductMinTemp_C; // Мин. Т продукта по активным датчикам из последней телеметрии
	bool lastActiveProductMinTemp_Valid; // true: в последней телеметрии был хотя бы один активный датчик продукта
	bool lastTelemetryAlrmBit;       // Предыдущее значение бита Alrm из телеметрии (для детекта перехода 0->1)
	System::Windows::Forms::Timer^ controlLogAbsenceTimer; // Таймер: при отсутствии телеметрии с _Wrk=1 сбрасывает controllerAutoModeActive
	System::Windows::Forms::Timer^ sendStateTimer;        // Таймер команды «Отправить состояние» по интервалу измерений
	bool autoRestartInternalUncheck; // Why: one-shot UX unchecks the box; we must not cancel the pending START.
	bool settingsLoading;            // Why: avoid side-effects (timers/log/save) while applying persisted settings.
	System::String^ pendingVersion;  // временное хранение версии для обновления UI из другого потока
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
		private: System::Windows::Forms::Button^ button_CMDINFO;
		private: System::Windows::Forms::Label^ labelSTOP;
		private: System::Windows::Forms::Label^ labelSTART;
		
		// Элементы автозапуска по времени
		private: System::Windows::Forms::CheckBox^ checkBoxAutoStart;
		private: System::Windows::Forms::DateTimePicker^ dateTimePickerAutoStart;
		private: System::Windows::Forms::Label^ labelAutoStart;
		private: System::Windows::Forms::Timer^ timerAutoStart;
		// Элементы автоперезапуска по времени
		private: System::Windows::Forms::CheckBox^ checkBoxAutoRestart;
		private: System::Windows::Forms::DateTimePicker^ dateTimePickerAutoRestart;
		private: System::Windows::Forms::Label^ labelAutoRestart;
		private: System::Windows::Forms::Timer^ timerAutoRestart;
		// Measurement interval (seconds)
		private: System::Windows::Forms::Label^ labelMeasurementInterval;
		private: System::Windows::Forms::NumericUpDown^ numericUpDownMeasurementInterval;
		private: System::Windows::Forms::Button^ buttonSaveMeasurementInterval;

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
				// Критично: используется для сигнализации завершения экспорта (неблокирующее закрытие).
				exportCompletedEvent = gcnew System::Threading::ManualResetEvent(false);
				// Критично: защищает DataTable от гонок между UI-обновлениями и фоновой копией в Excel.
				dataTableSync = gcnew System::Object();
				pendingRow = nullptr;
				pendingRowWrkBit = false;
				excelExportSync = gcnew System::Object();
				commandPipelineGate = gcnew System::Object();
				excelExportInProgress = 0;
				formGuid = nullptr;
				tabControl1PrevTab = nullptr;
				dataGridView1Dirty = false;
				dataGridView2Dirty = false;
				paramsLoadedFromDevice = false;
				startupSequenceCompleted = false;
				// Критично: recv()-поток не должен упираться в SemaphoreFullException при Release().
				responseQueue = gcnew System::Collections::Concurrent::ConcurrentQueue<cli::array<System::Byte>^>();
				responseAvailable = gcnew System::Threading::Semaphore(0, System::Int32::MaxValue);
				hasTelemetry = false;
				inactivityCloseRequested = false;
				lastTelemetryTime = DateTime::MinValue;
				lastTelemetryDeviceSeconds = 0;
				lastCmdInfoProbeTime = DateTime::MinValue;
				cmdInfoProbeInProgress = false;
				lastTelemetrySocket = INVALID_SOCKET;
				reconnectFixationLogPending = false;
				disconnectedSinceTicks = 0;
				lastStartupCommandsSocket = INVALID_SOCKET;
				postResetInitPending = false;
				postResetInitDeadline = DateTime::MinValue;
				postResetInitAttempt = 0;
				postResetInitTimer = nullptr;

				// Критично: устройство может быть выключено/включено; если переподключится в пределах 30 минут — продолжаем ту же таблицу.
				// Если телеметрии нет 30 минут — финализируем (экспорт) и закрываем форму.
				inactivityTimer = gcnew System::Windows::Forms::Timer();
				inactivityTimer->Interval = 10 * 1000; // 10 seconds
				inactivityTimer->Tick += gcnew EventHandler(this, &DataForm::OnInactivityTimerTick);
				inactivityTimer->Start();

				// Таймер отсутствия регулярного лога: сброс флага «контроллер в авторежиме» при отсутствии лога
				controlLogAbsenceTimer = gcnew System::Windows::Forms::Timer();
				controlLogAbsenceTimer->Interval = 2000; // 2 с
				controlLogAbsenceTimer->Tick += gcnew EventHandler(this, &DataForm::OnControlLogAbsenceTimerTick);
				controlLogAbsenceTimer->Start();

				sendStateTimer = gcnew System::Windows::Forms::Timer();
				{
					int intervalSec = 10;
					if (numericUpDownMeasurementInterval != nullptr && !numericUpDownMeasurementInterval->IsDisposed)
						intervalSec = System::Decimal::ToInt32(numericUpDownMeasurementInterval->Value);
					sendStateTimer->Interval = Math::Max(1000, intervalSec * 1000);
				}
				sendStateTimer->Tick += gcnew EventHandler(this, &DataForm::OnSendStateTimerTick);
				sendStateTimer->Stop();

			// Инициализируем путь сохранения из текстового поля
			excelSavePath = textBoxExcelDirectory->Text;
	excelFileName = nullptr;
	firstDataReceived = false;
	dataExportedToExcel = false;    // Флаг экспорта данных в Excel
	autoRestartPending = false;
	autoRestartStopIssuedTime = DateTime::MinValue;
		lastStopSuccessTime = DateTime::MinValue;
		lastStartSuccessTime = DateTime::MinValue;
		controllerAutoModeActive = false;
		wrkZeroConsecutiveCounts = 0;
		wrkLastSampleValid = false;
		wrkLastSampleTime = 0;
		lastControlLogTime = DateTime::MinValue;
		controlLogAbsenceStrikeCount = 0;
		lastFishCold_C_Valid = false;
		lastActiveProductMinTemp_Valid = false;
	lastTelemetryAlrmBit = false;
		controlLogAbsenceTimer = nullptr;
		autoRestartInternalUncheck = false;
	settingsLoading = false;
	pendingVersion = nullptr;       // Временная версия для обновления UI
	// Инициализация порта клиента
	clientPort = 0;

			// Загружаем настройки сразу после инициализации компонентов
			LoadSettings();
			// dataGridView1/2 для параметров алгоритма заполняются только при чтении с контроллера
			// (кнопка "Считать" — buttonReadParameters). ExcelSettings.txt для этих таблиц не используется.

			// Состояние кнопок START/STOP устанавливается при первом приёме данных

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
private: System::ComponentModel::IContainer^ components;

		private:
			/// <summary>
			/// Обязательная переменная конструктора.
			/// </summary>


	#pragma region Windows Form Designer generated code
			/// <summary>
			/// Требуемый метод для поддержки конструктора — не изменяйте 
			/// содержимое этого метода с помощью редактора кода.
			/// </summary>
			void InitializeComponent(void)
			{
				this->components = (gcnew System::ComponentModel::Container());
				this->menuStrip1 = (gcnew System::Windows::Forms::MenuStrip());
				this->выходToolStripMenuItem = (gcnew System::Windows::Forms::ToolStripMenuItem());
				this->buttonExcel = (gcnew System::Windows::Forms::Button());
				this->tabControl1 = (gcnew System::Windows::Forms::TabControl());
				this->tabPage1 = (gcnew System::Windows::Forms::TabPage());
				this->label7 = (gcnew System::Windows::Forms::Label());
				this->label6 = (gcnew System::Windows::Forms::Label());
				this->label5 = (gcnew System::Windows::Forms::Label());
				this->label4 = (gcnew System::Windows::Forms::Label());
				this->label3 = (gcnew System::Windows::Forms::Label());
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
				this->label_Version = (gcnew System::Windows::Forms::Label());
				this->labelVersion = (gcnew System::Windows::Forms::Label());
				this->button_RESET = (gcnew System::Windows::Forms::Button());
				this->button_CMDINFO = (gcnew System::Windows::Forms::Button());
				this->label_Commands_Info = (gcnew System::Windows::Forms::Label());
				this->Label_Commands = (gcnew System::Windows::Forms::Label());
				this->labelSTOP = (gcnew System::Windows::Forms::Label());
				this->labelSTART = (gcnew System::Windows::Forms::Label());
				this->buttonSTOP = (gcnew System::Windows::Forms::Button());
				this->buttonSTART = (gcnew System::Windows::Forms::Button());
				this->checkBoxAutoStart = (gcnew System::Windows::Forms::CheckBox());
				this->dateTimePickerAutoStart = (gcnew System::Windows::Forms::DateTimePicker());
				this->labelAutoStart = (gcnew System::Windows::Forms::Label());
				this->checkBoxAutoRestart = (gcnew System::Windows::Forms::CheckBox());
				this->dateTimePickerAutoRestart = (gcnew System::Windows::Forms::DateTimePicker());
				this->labelAutoRestart = (gcnew System::Windows::Forms::Label());
				this->buttonSaveMeasurementInterval = (gcnew System::Windows::Forms::Button());
				this->numericUpDownMeasurementInterval = (gcnew System::Windows::Forms::NumericUpDown());
				this->labelMeasurementInterval = (gcnew System::Windows::Forms::Label());
				this->buttonBrowse = (gcnew System::Windows::Forms::Button());
				this->textBoxExcelDirectory = (gcnew System::Windows::Forms::TextBox());
				this->labelExcelDirectory = (gcnew System::Windows::Forms::Label());
				this->tabPage3 = (gcnew System::Windows::Forms::TabPage());
				this->buttonLoadFromFile = (gcnew System::Windows::Forms::Button());
				this->label2 = (gcnew System::Windows::Forms::Label());
				this->buttonWriteParameters = (gcnew System::Windows::Forms::Button());
				this->buttonReadParameters = (gcnew System::Windows::Forms::Button());
				this->label1 = (gcnew System::Windows::Forms::Label());
				this->dataGridView2 = (gcnew System::Windows::Forms::DataGridView());
				this->Parameter2 = (gcnew System::Windows::Forms::DataGridViewTextBoxColumn());
				this->Description = (gcnew System::Windows::Forms::DataGridViewTextBoxColumn());
				this->Value = (gcnew System::Windows::Forms::DataGridViewTextBoxColumn());
				this->dataGridView1 = (gcnew System::Windows::Forms::DataGridView());
				this->Parameter = (gcnew System::Windows::Forms::DataGridViewTextBoxColumn());
				this->WarmUP = (gcnew System::Windows::Forms::DataGridViewTextBoxColumn());
				this->Plateau = (gcnew System::Windows::Forms::DataGridViewTextBoxColumn());
				this->Finish = (gcnew System::Windows::Forms::DataGridViewTextBoxColumn());
				this->tabPage4 = (gcnew System::Windows::Forms::TabPage());
				this->buttonCheckAlarm = (gcnew System::Windows::Forms::Button());
				this->dataGridEquipmentAlarm = (gcnew System::Windows::Forms::DataGridView());
				this->timerAutoStart = (gcnew System::Windows::Forms::Timer(this->components));
				this->timerAutoRestart = (gcnew System::Windows::Forms::Timer(this->components));
				this->label_Defroster_Info = (gcnew System::Windows::Forms::Label());
				this->labelDefrosterState = (gcnew System::Windows::Forms::Label());
				this->menuStrip1->SuspendLayout();
				this->tabControl1->SuspendLayout();
				this->tabPage1->SuspendLayout();
				(cli::safe_cast<System::ComponentModel::ISupportInitialize^>(this->dataGridView))->BeginInit();
				this->tabPage2->SuspendLayout();
				(cli::safe_cast<System::ComponentModel::ISupportInitialize^>(this->numericUpDownMeasurementInterval))->BeginInit();
				this->tabPage3->SuspendLayout();
				(cli::safe_cast<System::ComponentModel::ISupportInitialize^>(this->dataGridView2))->BeginInit();
				(cli::safe_cast<System::ComponentModel::ISupportInitialize^>(this->dataGridView1))->BeginInit();
				this->tabPage4->SuspendLayout();
				(cli::safe_cast<System::ComponentModel::ISupportInitialize^>(this->dataGridEquipmentAlarm))->BeginInit();
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
				this->tabControl1->Controls->Add(this->tabPage3);
				this->tabControl1->Controls->Add(this->tabPage4);
				this->tabControl1->Location = System::Drawing::Point(0, 36);
				this->tabControl1->Name = L"tabControl1";
				this->tabControl1->SelectedIndex = 0;
				this->tabControl1->Size = System::Drawing::Size(1494, 527);
				this->tabControl1->TabIndex = 6;
				this->tabControl1->SelectedIndexChanged += gcnew System::EventHandler(this, &DataForm::tabControl1_SelectedIndexChanged);
				// 
				// tabPage1
				// 
				this->tabPage1->AutoScroll = true;
				this->tabPage1->Controls->Add(this->label7);
				this->tabPage1->Controls->Add(this->label6);
				this->tabPage1->Controls->Add(this->label5);
				this->tabPage1->Controls->Add(this->label4);
				this->tabPage1->Controls->Add(this->label3);
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
				// label7
				// 
				this->label7->AutoSize = true;
				this->label7->Location = System::Drawing::Point(1052, 17);
				this->label7->Name = L"label7";
				this->label7->Size = System::Drawing::Size(43, 20);
				this->label7->TabIndex = 18;
				this->label7->Text = L"Прв:";
				// 
				// label6
				// 
				this->label6->AutoSize = true;
				this->label6->Location = System::Drawing::Point(924, 17);
				this->label6->Name = L"label6";
				this->label6->Size = System::Drawing::Size(43, 20);
				this->label6->TabIndex = 17;
				this->label6->Text = L"Лев:";
				// 
				// label5
				// 
				this->label5->AutoSize = true;
				this->label5->Location = System::Drawing::Point(649, 17);
				this->label5->Name = L"label5";
				this->label5->Size = System::Drawing::Size(43, 20);
				this->label5->TabIndex = 16;
				this->label5->Text = L"Прв:";
				// 
				// label4
				// 
				this->label4->AutoSize = true;
				this->label4->Location = System::Drawing::Point(527, 17);
				this->label4->Name = L"label4";
				this->label4->Size = System::Drawing::Size(25, 20);
				this->label4->TabIndex = 15;
				this->label4->Text = L"Ц:";
				// 
				// label3
				// 
				this->label3->AutoSize = true;
				this->label3->Location = System::Drawing::Point(391, 17);
				this->label3->Name = L"label3";
				this->label3->Size = System::Drawing::Size(43, 20);
				this->label3->TabIndex = 14;
				this->label3->Text = L"Лев:";
				// 
				// T_product_right
				// 
				this->T_product_right->AutoSize = true;
				this->T_product_right->Location = System::Drawing::Point(1101, 17);
				this->T_product_right->Name = L"T_product_right";
				this->T_product_right->Size = System::Drawing::Size(23, 20);
				this->T_product_right->TabIndex = 13;
				this->T_product_right->Text = L"- -";
				this->T_product_right->TextAlign = System::Drawing::ContentAlignment::MiddleCenter;
				// 
				// T_product_left
				// 
				this->T_product_left->AutoSize = true;
				this->T_product_left->Location = System::Drawing::Point(973, 17);
				this->T_product_left->Name = L"T_product_left";
				this->T_product_left->Size = System::Drawing::Size(23, 20);
				this->T_product_left->TabIndex = 13;
				this->T_product_left->Text = L"- -";
				this->T_product_left->TextAlign = System::Drawing::ContentAlignment::MiddleCenter;
				// 
				// LabelProduct
				// 
				this->LabelProduct->AutoSize = true;
				this->LabelProduct->Font = (gcnew System::Drawing::Font(L"Microsoft Sans Serif", 8, System::Drawing::FontStyle::Bold, System::Drawing::GraphicsUnit::Point,
					static_cast<System::Byte>(204)));
				this->LabelProduct->Location = System::Drawing::Point(799, 17);
				this->LabelProduct->Name = L"LabelProduct";
				this->LabelProduct->Size = System::Drawing::Size(103, 20);
				this->LabelProduct->TabIndex = 12;
				this->LabelProduct->Text = L"Т продукта";
				// 
				// T_def_right
				// 
				this->T_def_right->AutoSize = true;
				this->T_def_right->Location = System::Drawing::Point(698, 17);
				this->T_def_right->Name = L"T_def_right";
				this->T_def_right->Size = System::Drawing::Size(23, 20);
				this->T_def_right->TabIndex = 11;
				this->T_def_right->Text = L"- -";
				this->T_def_right->TextAlign = System::Drawing::ContentAlignment::MiddleCenter;
				// 
				// T_def_center
				// 
				this->T_def_center->AutoSize = true;
				this->T_def_center->Location = System::Drawing::Point(558, 17);
				this->T_def_center->Name = L"T_def_center";
				this->T_def_center->Size = System::Drawing::Size(23, 20);
				this->T_def_center->TabIndex = 10;
				this->T_def_center->Text = L"- -";
				this->T_def_center->TextAlign = System::Drawing::ContentAlignment::MiddleCenter;
				// 
				// T_def_left
				// 
				this->T_def_left->AutoSize = true;
				this->T_def_left->Location = System::Drawing::Point(440, 17);
				this->T_def_left->Name = L"T_def_left";
				this->T_def_left->Size = System::Drawing::Size(23, 20);
				this->T_def_left->TabIndex = 9;
				this->T_def_left->Text = L"- -";
				this->T_def_left->TextAlign = System::Drawing::ContentAlignment::MiddleCenter;
				// 
				// LabelDefroster
				// 
				this->LabelDefroster->AutoSize = true;
				this->LabelDefroster->Font = (gcnew System::Drawing::Font(L"Microsoft Sans Serif", 8, System::Drawing::FontStyle::Bold, System::Drawing::GraphicsUnit::Point,
					static_cast<System::Byte>(204)));
				this->LabelDefroster->Location = System::Drawing::Point(254, 17);
				this->LabelDefroster->Name = L"LabelDefroster";
				this->LabelDefroster->Size = System::Drawing::Size(131, 20);
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
				this->dataGridView->RowTemplate->Height = 20;
				this->dataGridView->RowTemplate->Resizable = System::Windows::Forms::DataGridViewTriState::True;
				this->dataGridView->Size = System::Drawing::Size(1276, 400);
				this->dataGridView->TabIndex = 6;
				this->dataGridView->CellContentClick += gcnew System::Windows::Forms::DataGridViewCellEventHandler(this, &DataForm::dataGridView_CellContentClick);
				// 
				// tabPage2
				// 
				this->tabPage2->Controls->Add(this->label_Defroster_Info);
				this->tabPage2->Controls->Add(this->labelDefrosterState);
				this->tabPage2->Controls->Add(this->label_Version);
				this->tabPage2->Controls->Add(this->labelVersion);
				this->tabPage2->Controls->Add(this->button_RESET);
				this->tabPage2->Controls->Add(this->button_CMDINFO);
				this->tabPage2->Controls->Add(this->label_Commands_Info);
				this->tabPage2->Controls->Add(this->Label_Commands);
				this->tabPage2->Controls->Add(this->labelSTOP);
				this->tabPage2->Controls->Add(this->labelSTART);
				this->tabPage2->Controls->Add(this->buttonSTOP);
				this->tabPage2->Controls->Add(this->buttonSTART);
				this->tabPage2->Controls->Add(this->checkBoxAutoStart);
				this->tabPage2->Controls->Add(this->dateTimePickerAutoStart);
				this->tabPage2->Controls->Add(this->labelAutoStart);
				this->tabPage2->Controls->Add(this->checkBoxAutoRestart);
				this->tabPage2->Controls->Add(this->dateTimePickerAutoRestart);
				this->tabPage2->Controls->Add(this->labelAutoRestart);
				this->tabPage2->Controls->Add(this->buttonSaveMeasurementInterval);
				this->tabPage2->Controls->Add(this->numericUpDownMeasurementInterval);
				this->tabPage2->Controls->Add(this->labelMeasurementInterval);
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
				// label_Version
				// 
				this->label_Version->AutoSize = true;
				this->label_Version->Location = System::Drawing::Point(234, 276);
				this->label_Version->Name = L"label_Version";
				this->label_Version->Size = System::Drawing::Size(96, 20);
				this->label_Version->TabIndex = 13;
				this->label_Version->Text = L"labelVersion";
				// 
				// labelVersion
				// 
				this->labelVersion->AutoSize = true;
				this->labelVersion->Location = System::Drawing::Point(48, 276);
				this->labelVersion->Name = L"labelVersion";
				this->labelVersion->Size = System::Drawing::Size(172, 20);
				this->labelVersion->TabIndex = 12;
				this->labelVersion->Text = L"Версия контроллера:";
				// 
				// button_RESET
				// 
				this->button_RESET->Location = System::Drawing::Point(48, 316);
				this->button_RESET->Name = L"button_RESET";
				this->button_RESET->RightToLeft = System::Windows::Forms::RightToLeft::Yes;
				this->button_RESET->Size = System::Drawing::Size(91, 33);
				this->button_RESET->TabIndex = 9;
				this->button_RESET->Text = L"СБРОС";
				this->button_RESET->UseVisualStyleBackColor = true;
				this->button_RESET->Click += gcnew System::EventHandler(this, &DataForm::button_RESET_Click);
				// 
				// button_CMDINFO
				// 
				this->button_CMDINFO->Location = System::Drawing::Point(160, 316);
				this->button_CMDINFO->Name = L"button_CMDINFO";
				this->button_CMDINFO->Size = System::Drawing::Size(130, 33);
				this->button_CMDINFO->TabIndex = 15;
				this->button_CMDINFO->Text = L"КОМАНДА";
				this->button_CMDINFO->UseVisualStyleBackColor = true;
				this->button_CMDINFO->Click += gcnew System::EventHandler(this, &DataForm::button_CMDINFO_Click);
				// 
				// label_Commands_Info
				// 
				this->label_Commands_Info->AutoSize = true;
				this->label_Commands_Info->Location = System::Drawing::Point(339, 322);
				this->label_Commands_Info->Name = L"label_Commands_Info";
				this->label_Commands_Info->Size = System::Drawing::Size(196, 20);
				this->label_Commands_Info->TabIndex = 8;
				this->label_Commands_Info->Text = L"Информация о команде:";
				// 
				// Label_Commands
				// 
				this->Label_Commands->AutoSize = true;
				this->Label_Commands->Location = System::Drawing::Point(541, 322);
				this->Label_Commands->Name = L"Label_Commands";
				this->Label_Commands->Size = System::Drawing::Size(212, 20);
				this->Label_Commands->TabIndex = 7;
				this->Label_Commands->Text = L"Команда не отправлялась";
				// 
				// labelSTOP
				// 
				this->labelSTOP->AutoSize = true;
				this->labelSTOP->BackColor = System::Drawing::Color::Snow;
				this->labelSTOP->Location = System::Drawing::Point(145, 122);
				this->labelSTOP->Name = L"labelSTOP";
				this->labelSTOP->Size = System::Drawing::Size(18, 20);
				this->labelSTOP->TabIndex = 6;
				this->labelSTOP->Text = L"0";
				// 
				// labelSTART
				// 
				this->labelSTART->AutoSize = true;
				this->labelSTART->BackColor = System::Drawing::Color::Snow;
				this->labelSTART->Location = System::Drawing::Point(145, 76);
				this->labelSTART->Name = L"labelSTART";
				this->labelSTART->Size = System::Drawing::Size(18, 20);
				this->labelSTART->TabIndex = 5;
				this->labelSTART->Text = L"0";
				// 
				// buttonSTOP
				// 
				this->buttonSTOP->Location = System::Drawing::Point(48, 116);
				this->buttonSTOP->Name = L"buttonSTOP";
				this->buttonSTOP->Size = System::Drawing::Size(91, 33);
				this->buttonSTOP->TabIndex = 4;
				this->buttonSTOP->Text = L"СТОП";
				this->buttonSTOP->UseVisualStyleBackColor = true;
				this->buttonSTOP->Click += gcnew System::EventHandler(this, &DataForm::buttonSTOP_Click);
				// 
				// buttonSTART
				// 
				this->buttonSTART->Location = System::Drawing::Point(48, 69);
				this->buttonSTART->Name = L"buttonSTART";
				this->buttonSTART->Size = System::Drawing::Size(91, 33);
				this->buttonSTART->TabIndex = 3;
				this->buttonSTART->Text = L"ПУСК";
				this->buttonSTART->UseVisualStyleBackColor = true;
				this->buttonSTART->Click += gcnew System::EventHandler(this, &DataForm::buttonSTART_Click);
				// 
				// checkBoxAutoStart
				// 
				this->checkBoxAutoStart->AutoSize = true;
				this->checkBoxAutoStart->Location = System::Drawing::Point(48, 215);
				this->checkBoxAutoStart->Name = L"checkBoxAutoStart";
				this->checkBoxAutoStart->Size = System::Drawing::Size(103, 24);
				this->checkBoxAutoStart->TabIndex = 10;
				this->checkBoxAutoStart->Text = L"Включен";
				this->checkBoxAutoStart->UseVisualStyleBackColor = true;
				this->checkBoxAutoStart->CheckedChanged += gcnew System::EventHandler(this, &DataForm::checkBoxAutoStart_CheckedChanged);
				// 
				// dateTimePickerAutoStart
				// 
				this->dateTimePickerAutoStart->CustomFormat = L"HH:mm";
				this->dateTimePickerAutoStart->Format = System::Windows::Forms::DateTimePickerFormat::Custom;
				this->dateTimePickerAutoStart->Location = System::Drawing::Point(160, 215);
				this->dateTimePickerAutoStart->Name = L"dateTimePickerAutoStart";
				this->dateTimePickerAutoStart->ShowUpDown = true;
				this->dateTimePickerAutoStart->Size = System::Drawing::Size(100, 26);
				this->dateTimePickerAutoStart->TabIndex = 11;
				this->dateTimePickerAutoStart->Value = System::DateTime(2025, 11, 12, 19, 29, 7, 631);
				this->dateTimePickerAutoStart->ValueChanged += gcnew System::EventHandler(this, &DataForm::dateTimePickerAutoStart_ValueChanged);
				// 
				// labelAutoStart
				// 
				this->labelAutoStart->AutoSize = true;
				this->labelAutoStart->Location = System::Drawing::Point(48, 185);
				this->labelAutoStart->Name = L"labelAutoStart";
				this->labelAutoStart->Size = System::Drawing::Size(113, 20);
				this->labelAutoStart->TabIndex = 9;
				this->labelAutoStart->Text = L"Автозапуск в:";
				// 
				// checkBoxAutoRestart
				// 
				this->checkBoxAutoRestart->AutoSize = true;
				this->checkBoxAutoRestart->Location = System::Drawing::Point(339, 215);
				this->checkBoxAutoRestart->Name = L"checkBoxAutoRestart";
				this->checkBoxAutoRestart->Size = System::Drawing::Size(103, 24);
				this->checkBoxAutoRestart->TabIndex = 12;
				this->checkBoxAutoRestart->Text = L"Включен";
				this->checkBoxAutoRestart->UseVisualStyleBackColor = true;
				this->checkBoxAutoRestart->CheckedChanged += gcnew System::EventHandler(this, &DataForm::checkBoxAutoRestart_CheckedChanged);
				// 
				// dateTimePickerAutoRestart
				// 
				this->dateTimePickerAutoRestart->CustomFormat = L"HH:mm";
				this->dateTimePickerAutoRestart->Format = System::Windows::Forms::DateTimePickerFormat::Custom;
				this->dateTimePickerAutoRestart->Location = System::Drawing::Point(451, 215);
				this->dateTimePickerAutoRestart->Name = L"dateTimePickerAutoRestart";
				this->dateTimePickerAutoRestart->ShowUpDown = true;
				this->dateTimePickerAutoRestart->Size = System::Drawing::Size(100, 26);
				this->dateTimePickerAutoRestart->TabIndex = 13;
				this->dateTimePickerAutoRestart->ValueChanged += gcnew System::EventHandler(this, &DataForm::dateTimePickerAutoRestart_ValueChanged);
				// 
				// labelAutoRestart
				// 
				this->labelAutoRestart->AutoSize = true;
				this->labelAutoRestart->Location = System::Drawing::Point(339, 185);
				this->labelAutoRestart->Name = L"labelAutoRestart";
				this->labelAutoRestart->Size = System::Drawing::Size(149, 20);
				this->labelAutoRestart->TabIndex = 11;
				this->labelAutoRestart->Text = L"Автоперезапуск в:";
				// 
				// buttonSaveMeasurementInterval
				// 
				this->buttonSaveMeasurementInterval->Location = System::Drawing::Point(720, 213);
				this->buttonSaveMeasurementInterval->Name = L"buttonSaveMeasurementInterval";
				this->buttonSaveMeasurementInterval->Size = System::Drawing::Size(110, 30);
				this->buttonSaveMeasurementInterval->TabIndex = 16;
				this->buttonSaveMeasurementInterval->Text = L"Сохранить";
				this->buttonSaveMeasurementInterval->UseVisualStyleBackColor = true;
				this->buttonSaveMeasurementInterval->Click += gcnew System::EventHandler(this, &DataForm::buttonSaveMeasurementInterval_Click);
				// 
				// numericUpDownMeasurementInterval
				// 
				this->numericUpDownMeasurementInterval->Location = System::Drawing::Point(620, 215);
				this->numericUpDownMeasurementInterval->Maximum = System::Decimal(gcnew cli::array< System::Int32 >(4) { 86400, 0, 0, 0 });
				this->numericUpDownMeasurementInterval->Minimum = System::Decimal(gcnew cli::array< System::Int32 >(4) { 1, 0, 0, 0 });
				this->numericUpDownMeasurementInterval->Name = L"numericUpDownMeasurementInterval";
				this->numericUpDownMeasurementInterval->Size = System::Drawing::Size(90, 26);
				this->numericUpDownMeasurementInterval->TabIndex = 15;
				this->numericUpDownMeasurementInterval->Value = System::Decimal(gcnew cli::array< System::Int32 >(4) { 10, 0, 0, 0 });
				// 
				// labelMeasurementInterval
				// 
				this->labelMeasurementInterval->AutoSize = true;
				this->labelMeasurementInterval->Location = System::Drawing::Point(620, 185);
				this->labelMeasurementInterval->Name = L"labelMeasurementInterval";
				this->labelMeasurementInterval->Size = System::Drawing::Size(190, 20);
				this->labelMeasurementInterval->TabIndex = 14;
				this->labelMeasurementInterval->Text = L"Интервал измерений, с:";
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
				// tabPage3
				// 
				this->tabPage3->Controls->Add(this->buttonLoadFromFile);
				this->tabPage3->Controls->Add(this->label2);
				this->tabPage3->Controls->Add(this->buttonWriteParameters);
				this->tabPage3->Controls->Add(this->buttonReadParameters);
				this->tabPage3->Controls->Add(this->label1);
				this->tabPage3->Controls->Add(this->dataGridView2);
				this->tabPage3->Controls->Add(this->dataGridView1);
				this->tabPage3->Location = System::Drawing::Point(4, 29);
				this->tabPage3->Name = L"tabPage3";
				this->tabPage3->Size = System::Drawing::Size(1486, 494);
				this->tabPage3->TabIndex = 2;
				this->tabPage3->Text = L"Параметры";
				this->tabPage3->UseVisualStyleBackColor = true;
				// 
				// buttonLoadFromFile
				// 
				this->buttonLoadFromFile->Location = System::Drawing::Point(1013, 8);
				this->buttonLoadFromFile->Name = L"buttonLoadFromFile";
				this->buttonLoadFromFile->Size = System::Drawing::Size(117, 36);
				this->buttonLoadFromFile->TabIndex = 7;
				this->buttonLoadFromFile->Text = L"Загрузить";
				this->buttonLoadFromFile->UseVisualStyleBackColor = true;
				this->buttonLoadFromFile->Click += gcnew System::EventHandler(this, &DataForm::buttonLoadFromFile_Click);
				// 
				// label2
				// 
				this->label2->AutoSize = true;
				this->label2->Location = System::Drawing::Point(858, 16);
				this->label2->Name = L"label2";
				this->label2->Size = System::Drawing::Size(149, 20);
				this->label2->TabIndex = 6;
				this->label2->Text = L"Данные из файла:";
				this->label2->TextAlign = System::Drawing::ContentAlignment::TopCenter;
				// 
				// buttonWriteParameters
				// 
				this->buttonWriteParameters->Location = System::Drawing::Point(1167, 60);
				this->buttonWriteParameters->Name = L"buttonWriteParameters";
				this->buttonWriteParameters->Size = System::Drawing::Size(117, 36);
				this->buttonWriteParameters->TabIndex = 5;
				this->buttonWriteParameters->Text = L"Записать";
				this->buttonWriteParameters->UseVisualStyleBackColor = true;
				this->buttonWriteParameters->Click += gcnew System::EventHandler(this, &DataForm::buttonWriteParameters_Click);
				// 
				// buttonReadParameters
				// 
				this->buttonReadParameters->Location = System::Drawing::Point(1013, 60);
				this->buttonReadParameters->Name = L"buttonReadParameters";
				this->buttonReadParameters->Size = System::Drawing::Size(117, 36);
				this->buttonReadParameters->TabIndex = 4;
				this->buttonReadParameters->Text = L"Считать";
				this->buttonReadParameters->UseVisualStyleBackColor = true;
				this->buttonReadParameters->Click += gcnew System::EventHandler(this, &DataForm::buttonReadParameters_Click);
				// 
				// label1
				// 
				this->label1->AutoSize = true;
				this->label1->Location = System::Drawing::Point(905, 68);
				this->label1->Name = L"label1";
				this->label1->Size = System::Drawing::Size(102, 20);
				this->label1->TabIndex = 3;
				this->label1->Text = L"Дефростер:";
				// 
				// dataGridView2
				// 
				this->dataGridView2->AutoSizeColumnsMode = System::Windows::Forms::DataGridViewAutoSizeColumnsMode::Fill;
				this->dataGridView2->ColumnHeadersHeightSizeMode = System::Windows::Forms::DataGridViewColumnHeadersHeightSizeMode::AutoSize;
				this->dataGridView2->Columns->AddRange(gcnew cli::array< System::Windows::Forms::DataGridViewColumn^  >(3) {
					this->Parameter2,
						this->Description, this->Value
				});
				this->dataGridView2->Location = System::Drawing::Point(23, 8);
				this->dataGridView2->Name = L"dataGridView2";
				this->dataGridView2->RowHeadersWidth = 24;
				this->dataGridView2->RowTemplate->Height = 16;
				this->dataGridView2->Size = System::Drawing::Size(564, 465);
				this->dataGridView2->TabIndex = 1;
				this->dataGridView2->CellValueChanged += gcnew System::Windows::Forms::DataGridViewCellEventHandler(this, &DataForm::dataGridView2_CellValueChanged);
				this->dataGridView2->UserAddedRow += gcnew System::Windows::Forms::DataGridViewRowEventHandler(this, &DataForm::dataGridView2_RowChanged);
				this->dataGridView2->UserDeletedRow += gcnew System::Windows::Forms::DataGridViewRowEventHandler(this, &DataForm::dataGridView2_RowChanged);
				// 
				// Parameter2
				// 
				this->Parameter2->HeaderText = L"Параметр";
				this->Parameter2->MinimumWidth = 60;
				this->Parameter2->Name = L"Parameter2";
				// 
				// Description
				// 
				this->Description->HeaderText = L"Описание параметра";
				this->Description->MinimumWidth = 80;
				this->Description->Name = L"Description";
				// 
				// Value
				// 
				this->Value->HeaderText = L"Величина";
				this->Value->MinimumWidth = 50;
				this->Value->Name = L"Value";
				// 
				// dataGridView1
				// 
				this->dataGridView1->AutoSizeColumnsMode = System::Windows::Forms::DataGridViewAutoSizeColumnsMode::Fill;
				this->dataGridView1->ColumnHeadersHeightSizeMode = System::Windows::Forms::DataGridViewColumnHeadersHeightSizeMode::AutoSize;
				this->dataGridView1->Columns->AddRange(gcnew cli::array< System::Windows::Forms::DataGridViewColumn^  >(4) {
					this->Parameter,
						this->WarmUP, this->Plateau, this->Finish
				});
				this->dataGridView1->Location = System::Drawing::Point(593, 115);
				this->dataGridView1->Name = L"dataGridView1";
				this->dataGridView1->RowHeadersWidth = 24;
				this->dataGridView1->RowTemplate->Height = 16;
				this->dataGridView1->Size = System::Drawing::Size(691, 358);
				this->dataGridView1->TabIndex = 0;
				this->dataGridView1->CellValueChanged += gcnew System::Windows::Forms::DataGridViewCellEventHandler(this, &DataForm::dataGridView1_CellValueChanged);
				this->dataGridView1->UserAddedRow += gcnew System::Windows::Forms::DataGridViewRowEventHandler(this, &DataForm::dataGridView1_RowChanged);
				this->dataGridView1->UserDeletedRow += gcnew System::Windows::Forms::DataGridViewRowEventHandler(this, &DataForm::dataGridView1_RowChanged);
				// 
				// Parameter
				// 
				this->Parameter->HeaderText = L"Параметр";
				this->Parameter->MinimumWidth = 80;
				this->Parameter->Name = L"Parameter";
				// 
				// WarmUP
				// 
				this->WarmUP->HeaderText = L"WarmUp";
				this->WarmUP->MinimumWidth = 50;
				this->WarmUP->Name = L"WarmUP";
				// 
				// Plateau
				// 
				this->Plateau->HeaderText = L"Plateau";
				this->Plateau->MinimumWidth = 50;
				this->Plateau->Name = L"Plateau";
				// 
				// Finish
				// 
				this->Finish->HeaderText = L"Finish";
				this->Finish->MinimumWidth = 50;
				this->Finish->Name = L"Finish";
				// 
				// tabPage4
				// 
				this->tabPage4->Controls->Add(this->buttonCheckAlarm);
				this->tabPage4->Controls->Add(this->dataGridEquipmentAlarm);
				this->tabPage4->Location = System::Drawing::Point(4, 29);
				this->tabPage4->Name = L"tabPage4";
				this->tabPage4->Size = System::Drawing::Size(1486, 494);
				this->tabPage4->TabIndex = 3;
				this->tabPage4->Text = L"АВАРИЯ";
				this->tabPage4->UseVisualStyleBackColor = true;
				// 
				// buttonCheckAlarm
				// 
				this->buttonCheckAlarm->Location = System::Drawing::Point(38, 18);
				this->buttonCheckAlarm->Name = L"buttonCheckAlarm";
				this->buttonCheckAlarm->Size = System::Drawing::Size(114, 36);
				this->buttonCheckAlarm->TabIndex = 1;
				this->buttonCheckAlarm->Text = L"Обновить";
				this->buttonCheckAlarm->UseVisualStyleBackColor = true;
				this->buttonCheckAlarm->Click += gcnew System::EventHandler(this, &DataForm::buttonCheckAlarm_Click);
				// 
				// dataGridEquipmentAlarm
				// 
				this->dataGridEquipmentAlarm->ColumnHeadersHeightSizeMode = System::Windows::Forms::DataGridViewColumnHeadersHeightSizeMode::AutoSize;
				this->dataGridEquipmentAlarm->Location = System::Drawing::Point(34, 69);
				this->dataGridEquipmentAlarm->Name = L"dataGridEquipmentAlarm";
				this->dataGridEquipmentAlarm->RowHeadersWidth = 24;
				this->dataGridEquipmentAlarm->RowTemplate->Height = 16;
				this->dataGridEquipmentAlarm->Size = System::Drawing::Size(1132, 388);
				this->dataGridEquipmentAlarm->TabIndex = 0;
				// 
				// timerAutoStart
				// 
				this->timerAutoStart->Interval = 30000;
				this->timerAutoStart->Tick += gcnew System::EventHandler(this, &DataForm::timerAutoStart_Tick);
				// 
				// timerAutoRestart
				// 
				this->timerAutoRestart->Interval = 30000;
				this->timerAutoRestart->Tick += gcnew System::EventHandler(this, &DataForm::timerAutoRestart_Tick);
				// 
				// label_Defroster_Info
				// 
				this->label_Defroster_Info->AutoSize = true;
				this->label_Defroster_Info->Location = System::Drawing::Point(339, 122);
				this->label_Defroster_Info->Name = L"label_Defroster_Info";
				this->label_Defroster_Info->Size = System::Drawing::Size(196, 20);
				this->label_Defroster_Info->TabIndex = 18;
				this->label_Defroster_Info->Text = L"Состояние дефростера:";
				// 
				// labelDefrosterState
				// 
				this->labelDefrosterState->AutoSize = true;
				this->labelDefrosterState->Location = System::Drawing::Point(541, 122);
				this->labelDefrosterState->Name = L"labelDefrosterState";
				this->labelDefrosterState->Size = System::Drawing::Size(212, 20);
				this->labelDefrosterState->TabIndex = 17;
				this->labelDefrosterState->Text = L"Команда не отправлялась";
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
				(cli::safe_cast<System::ComponentModel::ISupportInitialize^>(this->numericUpDownMeasurementInterval))->EndInit();
				this->tabPage3->ResumeLayout(false);
				this->tabPage3->PerformLayout();
				(cli::safe_cast<System::ComponentModel::ISupportInitialize^>(this->dataGridView2))->EndInit();
				(cli::safe_cast<System::ComponentModel::ISupportInitialize^>(this->dataGridView1))->EndInit();
				this->tabPage4->ResumeLayout(false);
				(cli::safe_cast<System::ComponentModel::ISupportInitialize^>(this->dataGridEquipmentAlarm))->EndInit();
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
			// Парсинг лога (Type 0x01): дополняет pendingRow параметрами; при _Wrk=1 добавляет строку в таблицу.
			void AppendControlLogToDataRow(cli::array<System::Byte>^ packet, int size);

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
			void OnExcelExportCompleted(bool enableButtonOnComplete);
			bool SendCommand(const ::Command& cmd); // Универсальный метод отправки команды (имя определяется автоматически)
			bool SendCommand(const ::Command& cmd, System::String^ commandName); // Универсальный метод отправки команды с явным именем
		void SendStartCommand(); // Метод для отправки команды START клиенту
		void SendStopCommand(); // Метод для отправки команды STOP клиенту
		void SendResetCommand(); // Метод для отправки команды RESET клиенту
			void SendCommandInfoRequest(); // Метод для запроса статуса обработки последней команды (device-side audit)
		bool SendVersionRequest(); // Метод для запроса версии прошивки контроллера
		bool SendSetIntervalCommand(int intervalSeconds); // Set measurement interval (seconds)
		// Defrost params API: exchange parameters with defroster (compatible with DefrostControl + CommandReceiver on controller)
		bool SetDefrostParam(uint8_t groupId, uint8_t paramId, const ::DefrostParamValue& value);
		bool GetDefrostParam(uint8_t groupId, uint8_t paramId, ::DefrostParamValue* outValue);
		bool GetDefrostGroup(uint8_t groupId, uint8_t page, uint8_t* outData, uint8_t outCapacity, uint8_t* outLength);
		bool GetAlarmFlags(::AlarmFlagsPayload* outFlags); // Запросить регистры аварий устройств и датчиков
		/** Отправить группу параметров (payload как в GET_DEFROST_GROUP). groupId 5 или 6. */
		bool SetDefrostGroup(uint8_t groupId, const uint8_t* payload, uint8_t payloadLen);
		void EnsureEquipmentAlarmGridColumns();
		void PopulateEquipmentAlarmGrid(uint16_t deviceFlags, uint16_t sensorFlags);
		void RefreshAlarmFlagsFromController();
		void UpdateVersionLabelInternal(); // Вспомогательный метод для обновления label_Version из UI потока
		
		// Методы для обработки ответов от контроллера
			void EnqueueResponse(cli::array<System::Byte>^ response); // Добавление ответа в очередь (вызывается из SServer)
			bool ReceiveResponse(::CommandResponse& response, int timeoutMs); // Прием ответа от контроллера с таймаутом
			bool ReceiveResponse(::CommandResponse& response, int timeoutMs, cli::array<System::Byte>^% rawResponse); // Receive response plus raw frame for diagnostics
			bool ReceiveResponse(::CommandResponse& response); // Прием ответа от контроллера (таймаут по умолчанию 1000 мс)
			void ProcessResponse(const ::CommandResponse& response); // Обработка полученного ответа
			void RestoreLabelCommandsColor(System::Object^ sender, System::EventArgs^ e); // Восстановление цвета Label_Commands
			bool SendCommandAndWaitResponse(const ::Command& cmd, ::CommandResponse& response, System::String^ commandName); // Отправка команды и ожидание ответа с именем
			bool SendCommandAndWaitResponse(const ::Command& cmd, ::CommandResponse& response); // Отправка команды и ожидание ответа (имя автоматически)
			
			void SetProgramStateUi(bool isRunning);	// Единый метод установки состояния кнопок/ламп ПУСК/СТОП
			/** Вызвать из обработчика лога: переключить в «программа запущена», только если ещё не в состоянии «остановлена» (чтобы запоздалый лог после СТОП не затирал кнопки). */
			void EnsureProgramRunningStateFromLog();
			void OnControlLogAbsenceTimerTick(System::Object^ sender, System::EventArgs^ e);
			void OnSendStateTimerTick(System::Object^ sender, System::EventArgs^ e);
			System::Void DataForm_FormClosed(Object^ sender, FormClosedEventArgs^ e);
			System::Void DataForm_HandleDestroyed(Object^ sender, EventArgs^ e);
		private:
			/** Применить запоздалый ответ GET_VERSION или SET_INTERVAL. Возвращает true, если ответ применён (не логировать как Discarding). */
			bool ApplyLateStartupResponse(const ::CommandResponse& candidate);
			void ScheduleCommandInfoProbe(System::String^ reason);
			void ExecuteCommandInfoProbe();
			void SchedulePostResetInit();
			void OnPostResetInitTimerTick(System::Object^ sender, System::EventArgs^ e);
			void SaveSettings();
			void LoadSettings();
			void UpdateDirectoryTextBox(String^ path);
			// Обработчик события закрытия формы
			System::Void DataForm_FormClosing(System::Object^ sender, System::Windows::Forms::FormClosingEventArgs^ e);
			
			// Метод для инициализации наименований битов
			static void InitializeBitFieldNames(gcroot<cli::array<cli::array<String^>^>^>& namesRef);

			void TriggerExcelExport();
			void ExecuteAutoRestartStart();
			bool TrySendControlCommandFireAndForget(uint8_t controlCode, System::String^ commandName);
			enum class CommandAckResult {
				Ok,
				ErrorResponse,
				NoResponse,
				SendFailed
			};
			CommandAckResult SendControlCommandWithAck(
				uint8_t controlCode,
				System::String^ commandName,
				int timeoutMs,
				int retries,
				::CommandResponse% lastResponse);
			bool StartExcelExportThread(bool isEmergency);
			void OnInactivityTimerTick(Object^ sender, EventArgs^ e);
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
private: System::Void buttonSaveMeasurementInterval_Click(System::Object^ sender, System::EventArgs^ e) {
	int intervalSeconds = System::Decimal::ToInt32(numericUpDownMeasurementInterval->Value);
	SendSetIntervalCommand(intervalSeconds);
	if (sendStateTimer != nullptr)
		sendStateTimer->Interval = Math::Max(1000, intervalSeconds * 1000);
}
private: System::Void button_CMDINFO_Click(System::Object^ sender, System::EventArgs^ e);

// ====================================================================
// Обработчики автозапуска по времени (объявления)
// ====================================================================
	
	private: System::Void checkBoxAutoStart_CheckedChanged(System::Object^ sender, System::EventArgs^ e);
	private: System::Void timerAutoStart_Tick(System::Object^ sender, System::EventArgs^ e);
	private: System::Void RestoreAutoStartColor(System::Object^ sender, System::EventArgs^ e);
	private: System::Void dateTimePickerAutoStart_ValueChanged(System::Object^ sender, System::EventArgs^ e);
	private: System::Void checkBoxAutoRestart_CheckedChanged(System::Object^ sender, System::EventArgs^ e);
	private: System::Void timerAutoRestart_Tick(System::Object^ sender, System::EventArgs^ e);
	private: System::Void RestoreAutoRestartColor(System::Object^ sender, System::EventArgs^ e);
	private: System::Void dateTimePickerAutoRestart_ValueChanged(System::Object^ sender, System::EventArgs^ e);
	private: System::Void tabControl1_SelectedIndexChanged(System::Object^ sender, System::EventArgs^ e);
	private: System::Void buttonLoadFromFile_Click(System::Object^ sender, System::EventArgs^ e);
	private: System::Void buttonSaveToFile_Click(System::Object^ sender, System::EventArgs^ e);
	private: System::Void buttonReadParameters_Click(System::Object^ sender, System::EventArgs^ e);
	private: System::Void AutoReadParametersFromController();
	private: System::Void buttonWriteParameters_Click(System::Object^ sender, System::EventArgs^ e);
	private: System::Void dataGridView1_CellValueChanged(System::Object^ sender, System::Windows::Forms::DataGridViewCellEventArgs^ e);
	private: System::Void dataGridView1_RowChanged(System::Object^ sender, System::Windows::Forms::DataGridViewRowEventArgs^ e);
	private: System::Void dataGridView2_CellValueChanged(System::Object^ sender, System::Windows::Forms::DataGridViewCellEventArgs^ e);
	private: System::Void dataGridView2_RowChanged(System::Object^ sender, System::Windows::Forms::DataGridViewRowEventArgs^ e);
	private: System::Void buttonCheckAlarm_Click(System::Object^ sender, System::EventArgs^ e);
	void LoadDataGridView1Defaults();
	void LoadDataGridView1FromFile();
	void SaveDataGridView1ToFile();
	void LoadDataGridView2Defaults();
	void LoadDataGridView2FromFile();
	void SaveDataGridView2ToFile();
	/** Загрузить параметры в таблицы 1 и 2 из листов Excel (Параметры по фазам, Параметры общие). */
	void LoadParamsFromExcelFile(System::String^ filePath);
	System::Data::DataTable^ BuildPhaseParamsDataTableForExcel();
	System::Data::DataTable^ BuildGlobalParamsDataTableForExcel();
	System::Data::DataTable^ BuildEquipmentAlarmDataTableForExcel();
	/** Заполнить dataGridView1 из payload ответа GET_DEFROST_GROUP(groupId=5). Параметры по фазам (WarmUP, Plateau, Finish). */
	void FillDataGridView1FromGroup5Payload(const uint8_t* payload, uint8_t payloadLen);
	/** Заполнить dataGridView2 из payload ответа GET_DEFROST_GROUP(groupId=6). Общие параметры. */
	void FillDataGridView2FromGroup6Payload(const uint8_t* payload, uint8_t payloadLen);
	
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


