#include "FormExcel.h"
#include "Chart.h"      // ExcelHelper
#include "DataForm.h"   // Для коллбека завершения экспорта

using namespace System;
using namespace System::Runtime::InteropServices;
using namespace System::Threading;
using namespace System::Drawing;

namespace ProjectServerW {

static System::String^ MapTemperatureColumnNameForExcel(System::String^ columnName)
{
	if (String::IsNullOrEmpty(columnName)) {
		return columnName;
	}

	if (columnName->Equals("T0")) return "T0 дефр.Левый";
	if (columnName->Equals("T1")) return "T1 дефр.Правый";
	if (columnName->Equals("T2")) return "T2 дефр.Центр";
	if (columnName->Equals("T3")) return "T3 прод.Лев";
	if (columnName->Equals("T4")) return "T4 прод.Пр";
	if (columnName->Equals("T5")) return "T5 корпус";

	return columnName;
}

static int GetExcelSeriesColorOle(System::String^ tColumnName)
{
	// Почему: Excel Interop ожидает OLE_COLOR. Используем ColorTranslator, чтобы не зависеть от порядка RGB/BGR.
	if (String::IsNullOrEmpty(tColumnName)) {
		return ColorTranslator::ToOle(Color::Black);
	}

	if (tColumnName->Equals("T0")) return ColorTranslator::ToOle(Color::FromArgb(0, 128, 0));       // зелёный
	if (tColumnName->Equals("T1")) return ColorTranslator::ToOle(Color::FromArgb(255, 0, 0));       // красный
	if (tColumnName->Equals("T2")) return ColorTranslator::ToOle(Color::FromArgb(255, 165, 0));     // оранжевый
	if (tColumnName->Equals("T3")) return ColorTranslator::ToOle(Color::FromArgb(0, 128, 0));       // зелёный
	if (tColumnName->Equals("T4")) return ColorTranslator::ToOle(Color::FromArgb(128, 0, 128));     // фиолетовый
	if (tColumnName->Equals("T5")) return ColorTranslator::ToOle(Color::FromArgb(165, 42, 42));     // коричневый

	if (tColumnName->Equals("_H0")) return ColorTranslator::ToOle(Color::FromArgb(255, 255, 153)); // light yellow
	if (tColumnName->Equals("_H1")) return ColorTranslator::ToOle(Color::FromArgb(210, 180, 140)); // light brown
	if (tColumnName->Equals("_H2")) return ColorTranslator::ToOle(Color::FromArgb(255, 255, 0));   // yellow
	if (tColumnName->Equals("_H3")) return ColorTranslator::ToOle(Color::FromArgb(165, 42, 42));   // brown
	if (tColumnName->Equals("H2")) return ColorTranslator::ToOle(Color::FromArgb(0, 0, 255));     // blue
	if (tColumnName->Equals("_Inj")) return ColorTranslator::ToOle(Color::FromArgb(0, 191, 255));  // sky blue
	if (tColumnName->Equals("T_filt_4")) return ColorTranslator::ToOle(Color::FromArgb(255, 182, 193)); // light pink
	return ColorTranslator::ToOle(Color::Black);
}

// Scaling belongs to the Excel export: raw telemetry, DI/DO flags and duty values stay unchanged.
static int GetDigitalOnLevelForExcel(String^ columnName)
{
	if (columnName->Equals("_H0") || columnName->Equals("Ten1_Left")) return 4;
	if (columnName->Equals("_H1") || columnName->Equals("Ten2_Left")) return 5;
	if (columnName->Equals("_H2") || columnName->Equals("Ten1_Right")) return 7;
	if (columnName->Equals("_H3") || columnName->Equals("Ten2_Right")) return 8;
	if (columnName->Equals("_Inj")) return 35;
	return 0;
}

static Object^ MapCellValueForExcel(String^ columnName, Object^ value)
{
	const int onLevel = GetDigitalOnLevelForExcel(columnName);
	if (onLevel == 0 || value == nullptr || Convert::IsDBNull(value)) return value;
	return Convert::ToDouble(value, System::Globalization::CultureInfo::InvariantCulture) == 1.0
		? safe_cast<Object^>(onLevel) : value;
}

static String^ GetExcelChartSeriesName(String^ columnName)
{
	if (columnName->Equals("_H0")) return "Нагреватель 1 левый (_H0)";
	if (columnName->Equals("_H1")) return "Нагреватель 2 левый (_H1)";
	if (columnName->Equals("_H2")) return "Нагреватель 1 правый (_H2)";
	if (columnName->Equals("_H3")) return "Нагреватель 2 правый (_H3)";
	if (columnName->Equals("H2")) return "Влажность центр (H2), %";
	if (columnName->Equals("_Inj")) return "Впрыск воды (_Inj)";
	if (columnName->Equals("T_filt_4")) return "T прод.Правый фильтр. (T_filt_4)";
	return MapTemperatureColumnNameForExcel(columnName);
}

static void AddExcelChartSeries(Microsoft::Office::Interop::Excel::SeriesCollection^ seriesCollection,
	Microsoft::Office::Interop::Excel::Worksheet^ ws, System::Data::DataTable^ table,
	Microsoft::Office::Interop::Excel::Range^ xRange, int lastRow, String^ columnName)
{
	const int columnIndex = table->Columns->IndexOf(columnName);
	if (columnIndex < 0) return; // Older sessions may not contain filtered temperatures.
	Microsoft::Office::Interop::Excel::Range^ yRange = nullptr;
	Microsoft::Office::Interop::Excel::Series^ series = nullptr;
	try {
		yRange = ws->Range[ws->Cells[2, columnIndex + 1], ws->Cells[lastRow, columnIndex + 1]];
		series = seriesCollection->NewSeries();
		series->XValues = xRange;
		series->Values = yRange;
		series->Name = GetExcelChartSeriesName(columnName);
		series->AxisGroup = columnName->Equals("H2")
			? Microsoft::Office::Interop::Excel::XlAxisGroup::xlSecondary
			: Microsoft::Office::Interop::Excel::XlAxisGroup::xlPrimary;
		const int oleColor = GetExcelSeriesColorOle(columnName);
		try { series->Border->Color = oleColor; } catch (...) {}
		try {
			series->Format->Line->ForeColor->RGB = oleColor;
			series->Format->Line->Weight = 2.0;
		} catch (...) {}
	}
	finally {
		if (series != nullptr) Marshal::ReleaseComObject(series);
		if (yRange != nullptr) Marshal::ReleaseComObject(yRange);
	}
}

FormExcel::FormExcel() {
	excelGlobalMutex = gcnew Mutex(false, "Global\\ProjectServerW_Excel_Mutex");
	excelExportQueue = gcnew System::Collections::Concurrent::ConcurrentQueue<ExcelExportJob^>();
	failedExportJobs = gcnew System::Collections::Concurrent::ConcurrentQueue<ExcelExportJob^>();
	excelExportQueueEvent = gcnew AutoResetEvent(false);
	excelExportWorkerSync = gcnew Object();
	excelExportWorkerThread = nullptr;
	excelExportWorkerStarted = false;
	excelAllExportsCompletedEvent = gcnew ManualResetEvent(true);
	excelActiveExportJobs = 0;
}

Mutex^ FormExcel::GetExcelGlobalMutex() {
	return excelGlobalMutex;
}

bool FormExcel::WaitForAllExports(int timeoutMs) {
	EnsureExcelExportWorker();
	return excelAllExportsCompletedEvent->WaitOne(timeoutMs);
}

void FormExcel::UpdateExcelIdleState() {
	try {
		if (excelActiveExportJobs == 0 && excelExportQueue != nullptr && excelExportQueue->IsEmpty) {
			excelAllExportsCompletedEvent->Set();
		}
		else {
			excelAllExportsCompletedEvent->Reset();
		}
	}
	catch (...) {}
}

void FormExcel::EnsureExcelExportWorker() {
	Monitor::Enter(excelExportWorkerSync);
	try {
		if (excelExportWorkerStarted && excelExportWorkerThread != nullptr) {
			return;
		}

		excelExportWorkerThread = gcnew Thread(gcnew ThreadStart(&ProjectServerW::FormExcel::ExcelExportWorkerLoop));
		excelExportWorkerThread->IsBackground = true;
		excelExportWorkerThread->SetApartmentState(ApartmentState::STA);
		excelExportWorkerThread->Start();
		excelExportWorkerStarted = true;
	}
	finally {
		Monitor::Exit(excelExportWorkerSync);
	}
}

void FormExcel::EnqueueExport(ExcelExportJob^ job) {
	if (job == nullptr) {
		return;
	}
	EnsureExcelExportWorker();
	try { excelAllExportsCompletedEvent->Reset(); } catch (...) {}
    // Failed immutable snapshots survive table clearing and are retried with the next export.
    ExcelExportJob^ failedJob = nullptr;
    while (failedExportJobs->TryDequeue(failedJob))
        excelExportQueue->Enqueue(failedJob);
	excelExportQueue->Enqueue(job);
	excelExportQueueEvent->Set();
}

void FormExcel::ExcelExportWorkerLoop() {
	while (true) {
		excelExportQueueEvent->WaitOne();

		ExcelExportJob^ job = nullptr;
		while (excelExportQueue->TryDequeue(job)) {
			ProcessExcelExportJob(job);
		}
		UpdateExcelIdleState();
	}
}

System::String^ FormExcel::ResolveExcelSaveDirectory(System::String^ preferredDirectory) {
	try {
		if (!String::IsNullOrEmpty(preferredDirectory) && System::IO::Directory::Exists(preferredDirectory)) {
			return preferredDirectory;
		}
	}
	catch (...) {}

	System::String^ appPath = System::IO::Path::GetDirectoryName(System::Windows::Forms::Application::ExecutablePath);
	System::String^ dir = System::IO::Path::Combine(appPath, "SensorData");
	if (!System::IO::Directory::Exists(dir)) {
		System::IO::Directory::CreateDirectory(dir);
	}
	return dir;
}

void FormExcel::ProcessExcelExportJob(ExcelExportJob^ job) {
	if (job == nullptr) {	// Если задача не существует, выходим
		return;
	}
	// Симметрия с Decrement в finally: иначе счётчик уходит в минус и WaitForAllExports никогда не получает сигнал.
	System::Threading::Interlocked::Increment(excelActiveExportJobs);
	bool exportSucceeded = false;
	bool exportRequeued = false;
	try {
	if (job->tableSnapshot == nullptr) {	// Если таблица не существует, выходим
		try {
			DataForm^ form = nullptr;
			if (job->formRef != nullptr && job->formRef->IsAlive) {	// Если ссылка на форму существует и жива, преобразуем её в объект DataForm
				form = dynamic_cast<DataForm^>(job->formRef->Target);	// Преобразуем ссылку на форму в объект DataForm
			}
			if (form != nullptr && !form->IsDisposed && !form->Disposing) {	// Если форма существует и не уничтожена, вызываем метод OnExcelExportCompleted
				if (job->completesCycle)
                    form->OnCompletedCycleSaved(job->sessionId, exportSucceeded);
                form->OnExcelExportCompleted(job->sessionId, job->enableButtonOnComplete, false, job->restartAfterExport);
			}
		}
		catch (...) {}	// Если произошла ошибка, выходим
		return;
	}
	bool mutexAcquired = false;	// Захвачен ли глобальный мьютекс Excel (нужен для finally)
	try {	// основной экспорт в Excel
		const int timeoutMs = 5 * 60 * 1000; // 5 минут
		try {
			mutexAcquired = excelGlobalMutex->WaitOne(timeoutMs);	// Ждём мьютекс на экспорт данных в Excel
		}
		catch (AbandonedMutexException^) {
			mutexAcquired = true;	// Если мьютекс занят другим процессом, устанавливаем флаг в true
			GlobalLogger::LogMessage("Предупреждение: Мьютекс Excel занят другим процессом; продолжаем экспорт.");	// Логируем предупреждение
		}

		if (!mutexAcquired) {      // Если мьютекс не занят, добавляем задачу в очередь и выходим
			// Критично: продолжаем ретраи через очередь — Excel может быть занят во время долгих экспортов.	
			exportRequeued = true;
			excelExportQueue->Enqueue(job);	// Добавляем задачу в очередь
			excelExportQueueEvent->Set();	// Устанавливаем событие очереди
			Thread::Sleep(1000);	// Ждём 1 секунду
			return;	// Выходим
		}

		DateTime exportStartTime = DateTime::Now;	// Начало экспорта после захвата мьютекса (для лога длительности)

		ExcelHelper^ excel = gcnew ExcelHelper();	// Создаём новый экземпляр ExcelHelper
		if (!excel->CreateNewWorkbook()) {	// Если не удалось создать новый workbook, выходим
			return;
		}

		Microsoft::Office::Interop::Excel::Worksheet^ ws = excel->GetWorksheet();
		Microsoft::Office::Interop::Excel::Application^ excelApp = safe_cast<Microsoft::Office::Interop::Excel::Application^>(ws->Application);

		try {
			excelApp->ScreenUpdating = false;	// Отключаем обновление экрана
			excelApp->Calculation = Microsoft::Office::Interop::Excel::XlCalculation::xlCalculationManual;	// Устанавливаем режим расчёта в ручной
			excelApp->EnableEvents = false;	// Отключаем события

			// Лист Info (метаданные)
			try {
				Microsoft::Office::Interop::Excel::Workbook^ wb = safe_cast<Microsoft::Office::Interop::Excel::Workbook^>(ws->Parent);	// Получаем workbook
				System::Object^ missing = System::Type::Missing;	// Получаем missing (отсутствие значения)
				// Type::Missing — это специальное значение «параметр не задан», Excel сам подставляет своё поведение по умолчанию.
				Microsoft::Office::Interop::Excel::Worksheet^ infoSheet =	// Создаём новый лист
					safe_cast<Microsoft::Office::Interop::Excel::Worksheet^>(wb->Worksheets->Add(missing, ws, 1, Microsoft::Office::Interop::Excel::XlSheetType::xlWorksheet));	// Создаём новый лист
				infoSheet->Name = "Info";	// Устанавливаем имя листа

				DateTime sessionStart = (job->sessionStart == DateTime::MinValue) ? DateTime::Now : job->sessionStart;
				DateTime sessionEnd = (job->sessionEnd == DateTime::MinValue) ? DateTime::Now : job->sessionEnd;

				infoSheet->Cells[1, 1] = "SessionStart";
				infoSheet->Cells[1, 2] = sessionStart.ToString("yyyy-MM-dd HH:mm:ss");
				infoSheet->Cells[2, 1] = "SessionEnd";
				infoSheet->Cells[2, 2] = sessionEnd.ToString("yyyy-MM-dd HH:mm:ss");
				infoSheet->Cells[3, 1] = "ClientPort";
				infoSheet->Cells[3, 2] = job->clientPort.ToString();
				infoSheet->Cells[4, 1] = "ClientIP";
				infoSheet->Cells[4, 2] = (job->clientIP != nullptr ? job->clientIP : "");
				infoSheet->Cells[5, 1] = "FormGuid";
				infoSheet->Cells[5, 2] = (job->formGuid != nullptr ? job->formGuid : "");

				Marshal::ReleaseComObject(infoSheet);
			}
			catch (...) {}

			int colCount = job->tableSnapshot->Columns->Count;
			int rowCount = job->tableSnapshot->Rows->Count;

			cli::array<System::Object^>^ headerArray = gcnew cli::array<System::Object^>(colCount);
			for (int c = 0; c < colCount; c++) {
				System::String^ colName = job->tableSnapshot->Columns[c]->ColumnName;
				headerArray[c] = MapTemperatureColumnNameForExcel(colName);
			}
			Microsoft::Office::Interop::Excel::Range^ headerRange = ws->Range[ws->Cells[1, 1], ws->Cells[1, colCount]];
			headerRange->Value2 = headerArray;
			Marshal::ReleaseComObject(headerRange);

			if (rowCount > 0) {
				cli::array<System::Object^, 2>^ dataArray = gcnew cli::array<System::Object^, 2>(rowCount, colCount);
				for (int r = 0; r < rowCount; r++) {
					System::Data::DataRow^ dr = job->tableSnapshot->Rows[r];
					for (int c = 0; c < colCount; c++) {
						dataArray[r, c] = MapCellValueForExcel(job->tableSnapshot->Columns[c]->ColumnName, dr[c]);
					}
				}

				Microsoft::Office::Interop::Excel::Range^ startCell = safe_cast<Microsoft::Office::Interop::Excel::Range^>(ws->Cells[2, 1]);
				Microsoft::Office::Interop::Excel::Range^ endCell = safe_cast<Microsoft::Office::Interop::Excel::Range^>(ws->Cells[rowCount + 1, colCount]);
				Microsoft::Office::Interop::Excel::Range^ dataRange = ws->Range[startCell, endCell];
				dataRange->Value2 = dataArray;
				Marshal::ReleaseComObject(startCell);
				Marshal::ReleaseComObject(endCell);
				Marshal::ReleaseComObject(dataRange);
			}

			// Добавление листа с графиком температур (Chart)
			try {
				if (rowCount > 0) {
					const int lastRow = rowCount + 1; // 1: заголовок
					if (lastRow >= 2) {
						Microsoft::Office::Interop::Excel::Workbook^ wb = safe_cast<Microsoft::Office::Interop::Excel::Workbook^>(ws->Parent);
						System::Object^ missing = System::Type::Missing;

						Microsoft::Office::Interop::Excel::Worksheet^ chartSheet =
							safe_cast<Microsoft::Office::Interop::Excel::Worksheet^>(wb->Worksheets->Add(missing, ws, 1, Microsoft::Office::Interop::Excel::XlSheetType::xlWorksheet));
						chartSheet->Name = "Chart";

						Microsoft::Office::Interop::Excel::ChartObjects^ chartObjects = nullptr;
						Microsoft::Office::Interop::Excel::ChartObject^ chartObject = nullptr;
						Microsoft::Office::Interop::Excel::Chart^ chart = nullptr;
						Microsoft::Office::Interop::Excel::Range^ xRange = nullptr;
						try {
							chartObjects = safe_cast<Microsoft::Office::Interop::Excel::ChartObjects^>(chartSheet->ChartObjects(missing));
							chartObject = chartObjects->Add(20, 20, 900, 500);
							chart = chartObject->Chart;
							chart->ChartType = Microsoft::Office::Interop::Excel::XlChartType::xlLine;
						}
						finally {
							if (chartObjects != nullptr) Marshal::ReleaseComObject(chartObjects);
						}

						// Ось X: RealTime или Time
						int xCol = -1;
						try {
							int idx = job->tableSnapshot->Columns->IndexOf("RealTime");
							if (idx >= 0) xCol = idx + 1;
							else {
								idx = job->tableSnapshot->Columns->IndexOf("Time");
								if (idx >= 0) xCol = idx + 1;
							}
						}
						catch (...) {}
						if (xCol < 1) xCol = 1;
						xRange = ws->Range[ws->Cells[2, xCol], ws->Cells[lastRow, xCol]];

						Microsoft::Office::Interop::Excel::SeriesCollection^ seriesCollection =
							safe_cast<Microsoft::Office::Interop::Excel::SeriesCollection^>(chart->SeriesCollection(missing));

						for (uint8_t i = 0; i < (SQ - 1); i++) {
							AddExcelChartSeries(seriesCollection, ws, job->tableSnapshot, xRange, lastRow,
								"T" + Convert::ToString(i));
						}
						cli::array<String^>^ extraColumns = gcnew cli::array<String^> {
							"_H0", "_H1", "_H2", "_H3", "H2", "_Inj", "T_filt_4"
						};
						for each (String^ columnName in extraColumns) {
							AddExcelChartSeries(seriesCollection, ws, job->tableSnapshot, xRange, lastRow, columnName);
						}
						Marshal::ReleaseComObject(seriesCollection);

						chart->HasTitle = true;
						chart->ChartTitle->Text = "Температуры, влажность и оборудование";

						Microsoft::Office::Interop::Excel::Axis^ categoryAxis = safe_cast<Microsoft::Office::Interop::Excel::Axis^>(
							chart->Axes(Microsoft::Office::Interop::Excel::XlAxisType::xlCategory, Microsoft::Office::Interop::Excel::XlAxisGroup::xlPrimary));
						categoryAxis->HasTitle = true;
						categoryAxis->AxisTitle->Text = "RealTime";

						Microsoft::Office::Interop::Excel::Axis^ valueAxis = safe_cast<Microsoft::Office::Interop::Excel::Axis^>(
							chart->Axes(Microsoft::Office::Interop::Excel::XlAxisType::xlValue, Microsoft::Office::Interop::Excel::XlAxisGroup::xlPrimary));
						valueAxis->HasTitle = true;
						valueAxis->AxisTitle->Text = "T, °C / уровни оборудования";
						if (job->tableSnapshot->Columns->Contains("H2")) {
							Microsoft::Office::Interop::Excel::Axis^ humidityAxis = safe_cast<Microsoft::Office::Interop::Excel::Axis^>(
								chart->Axes(Microsoft::Office::Interop::Excel::XlAxisType::xlValue,
									Microsoft::Office::Interop::Excel::XlAxisGroup::xlSecondary));
							humidityAxis->HasTitle = true;
							humidityAxis->AxisTitle->Text = "Влажность, %";
							humidityAxis->MinimumScale = 0.0;
							humidityAxis->MaximumScale = 100.0;
							Marshal::ReleaseComObject(humidityAxis);
						}

						if (xRange != nullptr) Marshal::ReleaseComObject(xRange);
					}
				}
			}
			catch (Exception^ ex) {
				try {
					GlobalLogger::LogMessage("Warning: Excel chart creation failed: " + ex->Message);
				}
				catch (...) {}
			}
			// Лист «Параметры по фазам»: таблица 1 (dataGridView1).
			try {
				if (job->paramsPhase != nullptr) {
					Microsoft::Office::Interop::Excel::Workbook^ wb = safe_cast<Microsoft::Office::Interop::Excel::Workbook^>(ws->Parent);
					System::Object^ missing = System::Type::Missing;

					// Найти существующий лист с таким именем или создать новый.
					Microsoft::Office::Interop::Excel::Worksheet^ sheet1 = nullptr;
					try {
						sheet1 = safe_cast<Microsoft::Office::Interop::Excel::Worksheet^>(
							wb->Worksheets->Item["Параметры по фазам"]);
					}
					catch (...) {
						sheet1 = nullptr;
					}
					if (sheet1 == nullptr) {
						sheet1 = safe_cast<Microsoft::Office::Interop::Excel::Worksheet^>(
							wb->Worksheets->Add(missing, ws, 1, Microsoft::Office::Interop::Excel::XlSheetType::xlWorksheet));
						sheet1->Name = "Параметры по фазам";
					}
					else {
						// Очищаем старые данные, если лист уже есть.
						try {
							Microsoft::Office::Interop::Excel::Range^ used1 = sheet1->UsedRange;
							used1->Clear();
							Marshal::ReleaseComObject(used1);
						}
						catch (...) {}
					}

					int colCount1 = job->paramsPhase->Columns->Count;
					int rowCount1 = job->paramsPhase->Rows->Count;
					cli::array<System::Object^>^ header1 = gcnew cli::array<System::Object^>(colCount1);
					for (int c = 0; c < colCount1; c++) {
						header1[c] = job->paramsPhase->Columns[c]->ColumnName;
					}
					Microsoft::Office::Interop::Excel::Range^ headerRange1 =
						sheet1->Range[sheet1->Cells[1, 1], sheet1->Cells[1, colCount1]];
					headerRange1->Value2 = header1;
					Marshal::ReleaseComObject(headerRange1);

					if (rowCount1 > 0) {
						cli::array<System::Object^, 2>^ data1 = gcnew cli::array<System::Object^, 2>(rowCount1, colCount1);
						for (int r = 0; r < rowCount1; r++) {
							System::Data::DataRow^ dr = job->paramsPhase->Rows[r];
							for (int c = 0; c < colCount1; c++) {
								data1[r, c] = dr[c];
							}
						}
						Microsoft::Office::Interop::Excel::Range^ start1 =
							safe_cast<Microsoft::Office::Interop::Excel::Range^>(sheet1->Cells[2, 1]);
						Microsoft::Office::Interop::Excel::Range^ end1 =
							safe_cast<Microsoft::Office::Interop::Excel::Range^>(sheet1->Cells[1 + rowCount1, colCount1]);
						Microsoft::Office::Interop::Excel::Range^ dataRange1 = sheet1->Range[start1, end1];
						dataRange1->Value2 = data1;
						Marshal::ReleaseComObject(start1);
						Marshal::ReleaseComObject(end1);
						Marshal::ReleaseComObject(dataRange1);
					}

					Marshal::ReleaseComObject(sheet1);
					
				}
			}
			catch (Exception^ ex) {
				try {
					GlobalLogger::LogMessage("Warning: Excel sheet 'Параметры по фазам' failed: " + ex->Message);
				}
				catch (...) {}
			}

			// Лист «Параметры общие»: таблица 2 (dataGridView2).
			try {
				if (job->paramsGlobal != nullptr) {
					Microsoft::Office::Interop::Excel::Workbook^ wb = safe_cast<Microsoft::Office::Interop::Excel::Workbook^>(ws->Parent);
					System::Object^ missing = System::Type::Missing;

					// Найти существующий лист с таким именем или создать новый.
					Microsoft::Office::Interop::Excel::Worksheet^ sheet2 = nullptr;
					try {
						sheet2 = safe_cast<Microsoft::Office::Interop::Excel::Worksheet^>(
							wb->Worksheets->Item["Параметры общие"]);
					}
					catch (...) {
						sheet2 = nullptr;
					}
					if (sheet2 == nullptr) {
						sheet2 = safe_cast<Microsoft::Office::Interop::Excel::Worksheet^>(
							wb->Worksheets->Add(missing, ws, 1, Microsoft::Office::Interop::Excel::XlSheetType::xlWorksheet));
						sheet2->Name = "Параметры общие";
					}
					else {
						// Очищаем старые данные, если лист уже есть.
						try {
							Microsoft::Office::Interop::Excel::Range^ used2 = sheet2->UsedRange;
							used2->Clear();
							Marshal::ReleaseComObject(used2);
						}
						catch (...) {}
					}

					int colCount2 = job->paramsGlobal->Columns->Count;
					int rowCount2 = job->paramsGlobal->Rows->Count;
					cli::array<System::Object^>^ header2 = gcnew cli::array<System::Object^>(colCount2);
					for (int c = 0; c < colCount2; c++) {
						header2[c] = job->paramsGlobal->Columns[c]->ColumnName;
					}
					Microsoft::Office::Interop::Excel::Range^ headerRange2 =
						sheet2->Range[sheet2->Cells[1, 1], sheet2->Cells[1, colCount2]];
					headerRange2->Value2 = header2;
					Marshal::ReleaseComObject(headerRange2);

					if (rowCount2 > 0) {
						cli::array<System::Object^, 2>^ data2 = gcnew cli::array<System::Object^, 2>(rowCount2, colCount2);
						for (int r = 0; r < rowCount2; r++) {
							System::Data::DataRow^ dr = job->paramsGlobal->Rows[r];
							for (int c = 0; c < colCount2; c++) {
								data2[r, c] = dr[c];
							}
						}
						Microsoft::Office::Interop::Excel::Range^ start2 =
							safe_cast<Microsoft::Office::Interop::Excel::Range^>(sheet2->Cells[2, 1]);
						Microsoft::Office::Interop::Excel::Range^ end2 =
							safe_cast<Microsoft::Office::Interop::Excel::Range^>(sheet2->Cells[1 + rowCount2, colCount2]);
						Microsoft::Office::Interop::Excel::Range^ dataRange2 = sheet2->Range[start2, end2];
						dataRange2->Value2 = data2;
						Marshal::ReleaseComObject(start2);
						Marshal::ReleaseComObject(end2);
						Marshal::ReleaseComObject(dataRange2);
					}

					Marshal::ReleaseComObject(sheet2);
					
				}
			}
			catch (Exception^ ex) {
				try {
					GlobalLogger::LogMessage("Warning: Excel sheet 'Параметры общие' failed: " + ex->Message);
				}
				catch (...) {}
			}

			// Лист «АВАРИЯ»: список аварийных устройств/датчиков на момент завершения цикла.
			try {
				if (job->equipmentAlarms != nullptr && job->equipmentAlarms->Rows->Count > 0) {
					Microsoft::Office::Interop::Excel::Workbook^ wb = safe_cast<Microsoft::Office::Interop::Excel::Workbook^>(ws->Parent);
					System::Object^ missing = System::Type::Missing;

					Microsoft::Office::Interop::Excel::Worksheet^ alarmSheet = nullptr;
					try {
						alarmSheet = safe_cast<Microsoft::Office::Interop::Excel::Worksheet^>(
							wb->Worksheets->Item["АВАРИЯ"]);
					}
					catch (...) {
						alarmSheet = nullptr;
					}
					if (alarmSheet == nullptr) {
						alarmSheet = safe_cast<Microsoft::Office::Interop::Excel::Worksheet^>(
							wb->Worksheets->Add(missing, ws, 1, Microsoft::Office::Interop::Excel::XlSheetType::xlWorksheet));
						alarmSheet->Name = "АВАРИЯ";
					}
					else {
						try {
							Microsoft::Office::Interop::Excel::Range^ used = alarmSheet->UsedRange;
							used->Clear();
							Marshal::ReleaseComObject(used);
						}
						catch (...) {}
					}

					int colCountA = job->equipmentAlarms->Columns->Count;
					int rowCountA = job->equipmentAlarms->Rows->Count;

					cli::array<System::Object^>^ headerA = gcnew cli::array<System::Object^>(colCountA);
					for (int c = 0; c < colCountA; c++) {
						headerA[c] = job->equipmentAlarms->Columns[c]->ColumnName;
					}
					Microsoft::Office::Interop::Excel::Range^ headerRangeA =
						alarmSheet->Range[alarmSheet->Cells[1, 1], alarmSheet->Cells[1, colCountA]];
					headerRangeA->Value2 = headerA;
					Marshal::ReleaseComObject(headerRangeA);

					cli::array<System::Object^, 2>^ dataA = gcnew cli::array<System::Object^, 2>(rowCountA, colCountA);
					for (int r = 0; r < rowCountA; r++) {
						System::Data::DataRow^ dr = job->equipmentAlarms->Rows[r];
						for (int c = 0; c < colCountA; c++) {
							dataA[r, c] = dr[c];
						}
					}
					Microsoft::Office::Interop::Excel::Range^ startA =
						safe_cast<Microsoft::Office::Interop::Excel::Range^>(alarmSheet->Cells[2, 1]);
					Microsoft::Office::Interop::Excel::Range^ endA =
						safe_cast<Microsoft::Office::Interop::Excel::Range^>(alarmSheet->Cells[1 + rowCountA, colCountA]);
					Microsoft::Office::Interop::Excel::Range^ dataRangeA = alarmSheet->Range[startA, endA];
					dataRangeA->Value2 = dataA;
					Marshal::ReleaseComObject(startA);
					Marshal::ReleaseComObject(endA);
					Marshal::ReleaseComObject(dataRangeA);
					Marshal::ReleaseComObject(alarmSheet);
				}
			}
			catch (Exception^ ex) {
				try {
					GlobalLogger::LogMessage("Warning: Excel sheet 'АВАРИЯ' failed: " + ex->Message);
				}
				catch (...) {}
			}
		}
		finally {
			try {
				excelApp->Calculation = Microsoft::Office::Interop::Excel::XlCalculation::xlCalculationAutomatic;
				excelApp->ScreenUpdating = true;
				excelApp->EnableEvents = true;
			}
			catch (...) {}
			Marshal::ReleaseComObject(excelApp);
		}

		System::String^ dir = ResolveExcelSaveDirectory(job->saveDirectory);
		if (!dir->EndsWith("\\")) {
			dir += "\\";
		}

		DateTime start = (job->sessionStart == DateTime::MinValue) ? DateTime::Now : job->sessionStart;
		DateTime end = (job->sessionEnd == DateTime::MinValue) ? DateTime::Now : job->sessionEnd;
		System::String^ finalFileName = String::Format(
			"{3}_Start_{0}_End_{1}_Port{2}.xlsx",
			start.ToString("yyyy-MM-dd_HH-mm-ss"),
			end.ToString("yyyy-MM-dd_HH-mm-ss"),
			job->clientPort.ToString(),
            String::IsNullOrEmpty(job->filePrefix) ? "WorkData" : job->filePrefix);

        if (System::IO::File::Exists(dir + finalFileName)) {
            finalFileName = System::IO::Path::GetFileNameWithoutExtension(finalFileName) +
                "_" + job->sessionId.ToString("N") + ".xlsx";
        }
		const bool saved = excel->SaveAs(dir + finalFileName);
		excel->Close();
		delete excel;
		excel = nullptr;
		if (!saved) {
			GlobalLogger::LogMessage("Error: Excel export job failed to save: " + finalFileName);
			return;
		}
		exportSucceeded = true;

		DateTime exportEndTime = DateTime::Now;
		TimeSpan elapsed = exportEndTime.Subtract(exportStartTime);
		int exportedRows = 0;
		try {
			exportedRows = job->tableSnapshot->Rows->Count;
		}
		catch (...) {}

		GlobalLogger::LogMessage(String::Format(
			"Information: Файл Excel успешно сохранен: {0}\nВремя записи: {1} секунд ({2} строк)",
			finalFileName,
			elapsed.TotalSeconds.ToString("F2"),
			exportedRows));

		// Запускаем отложенную сборку мусора, чтобы быстрее освобождать COM-обёртки после экспорта.
		try {
			ThreadPool::QueueUserWorkItem(gcnew WaitCallback(DataForm::DelayedGarbageCollection));
		}
		catch (...) {}
	}
	catch (Exception^ ex) {
		GlobalLogger::LogMessage("Error: Excel export job failed: " + ex->ToString());
	}
	finally {
		if (mutexAcquired) {
			try { excelGlobalMutex->ReleaseMutex(); }
			catch (...) {}
		}

		try {
			DataForm^ form = nullptr;
			if (job != nullptr && job->formRef != nullptr && job->formRef->IsAlive) {
				form = dynamic_cast<DataForm^>(job->formRef->Target);
			}

			if (!exportRequeued && form != nullptr && !form->IsDisposed && !form->Disposing) {
				if (job->completesCycle)
                    form->OnCompletedCycleSaved(job->sessionId, exportSucceeded);
                form->OnExcelExportCompleted(job->sessionId, job->enableButtonOnComplete, exportSucceeded, job->restartAfterExport);
			}
		}
		catch (...) {}
        if (!exportSucceeded && !exportRequeued && job->tableSnapshot != nullptr) {
            // Completion has already accounted for the original job. Retries are detached
            // from UI state so they cannot decrement a newer job or trigger an old START.
            job->formRef = nullptr;
            job->completesCycle = false;
            job->restartAfterExport = false;
            job->enableButtonOnComplete = false;
            failedExportJobs->Enqueue(job);
            GlobalLogger::LogMessage(L"Error: Снимок Excel сохранён в памяти для повторной записи при следующем экспорте");
        }
	}
	}
	finally {
		System::Threading::Interlocked::Decrement(excelActiveExportJobs);
		UpdateExcelIdleState();
	}
}

} // namespace ProjectServerW


