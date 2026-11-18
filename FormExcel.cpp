#include "FormExcel.h"
#include "PidPlotScale.h"
#include "Chart.h"      // ExcelHelper
#include "DataForm.h"   // Для коллбека завершения экспорта

using namespace System;
using namespace System::Runtime::InteropServices;
using namespace System::Threading;
using namespace System::Drawing;

namespace ProjectServerW {
namespace Excel = Microsoft::Office::Interop::Excel;

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
	if (tColumnName->Equals("H0")) return ColorTranslator::ToOle(Color::Teal);
	if (tColumnName->Equals("H1")) return ColorTranslator::ToOle(Color::DarkOrange);
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
	if (columnName->Equals("H0")) return L"Влажность слева (H0), %";
	if (columnName->Equals("H1")) return L"Влажность справа (H1), %";
	if (columnName->Equals("H2")) return "Влажность центр (H2), %";
	if (columnName->Equals("_Inj")) return "Впрыск воды (_Inj)";
	if (columnName->Equals("T_filt_4")) return "T прод.Правый фильтр. (T_filt_4)";
	return MapTemperatureColumnNameForExcel(columnName);
}

// Читаем значения снимка. Пустые отсчёты остаются пропусками и в диаграмме.
static double ChartValue(System::Data::DataRow^ row, String^ column)
{
	if (!row->Table->Columns->Contains(column) || row->IsNull(column)) return Double::NaN;
	const double value = Convert::ToDouble(row[column]);
	return Double::IsInfinity(value) ? Double::NaN : value;
}

static void AddExcelLine(Excel::SeriesCollection^ seriesCollection, Excel::Range^ xRange,
	Excel::Range^ yRange, String^ name, Color color, bool secondary, double weight)
{
	Excel::Series^ series = seriesCollection->NewSeries();
	try {
		series->XValues = xRange;
		series->Values = yRange;
		series->Name = name;
		series->AxisGroup = secondary ? Excel::XlAxisGroup::xlSecondary : Excel::XlAxisGroup::xlPrimary;
		series->Format->Line->ForeColor->RGB = ColorTranslator::ToOle(color);
		series->Format->Line->Weight = static_cast<float>(weight);
	}
	finally { Marshal::ReleaseComObject(series); }
}

static void AddExcelSourceLine(Excel::SeriesCollection^ lines, Excel::Range^ xRange,
	Excel::Worksheet^ source, System::Data::DataTable^ table, String^ column,
	String^ name, Color color, bool secondary, double weight)
{
	const int index = table->Columns->IndexOf(column);
	if (index < 0) return;
	Excel::Range^ range = source->Range[source->Cells[2, index + 1], source->Cells[table->Rows->Count + 1, index + 1]];
	try { AddExcelLine(lines, xRange, range, name, color, secondary, weight); }
	finally { Marshal::ReleaseComObject(range); }
}

// Преобразованные значения живут на листе графика. Исходные температуры,
// статусы, доли и фазы в основной таблице сохраняют свои единицы измерения.
static void AddExcelMappedLine(Excel::SeriesCollection^ lines, Excel::Range^ xRange,
	Excel::Worksheet^ sheet, int helperColumn, cli::array<Object^, 2>^ values,
	String^ name, Color color, bool secondary, double weight)
{
	int rows = values->GetLength(0);
	sheet->Cells[1, helperColumn] = name;
	Excel::Range^ range = sheet->Range[sheet->Cells[2, helperColumn], sheet->Cells[rows + 1, helperColumn]];
	try {
		range->Value2 = values;
		AddExcelLine(lines, xRange, range, name, color, secondary, weight);
	}
	finally { Marshal::ReleaseComObject(range); }
}

static Excel::Axis^ ExcelValueAxis(Excel::Chart^ chart, bool secondary)
{
	return safe_cast<Excel::Axis^>(chart->Axes(Excel::XlAxisType::xlValue,
		secondary ? Excel::XlAxisGroup::xlSecondary : Excel::XlAxisGroup::xlPrimary));
}

static void SetExcelAxisTitle(Excel::Axis^ axis, String^ title)
{
	axis->HasTitle = true;
	axis->AxisTitle->Text = title;
}

static void AddExcelPhase(Excel::Chart^ chart, Excel::SeriesCollection^ lines, Excel::Range^ xRange,
	Excel::Worksheet^ sheet, System::Data::DataTable^ table, int column, bool humidityChart)
{
	auto values = gcnew cli::array<Object^, 2>(table->Rows->Count, 1);
	bool hasPhase = false;
	for (int i = 0; i < table->Rows->Count; ++i) {
		double phase = ChartValue(table->Rows[i], "phase");
		if (phase == 0 || phase == 1 || phase == 2) {
			values[i,0] = humidityChart ? 110.0 + 10.0 * phase : phase;
			hasPhase = true;
		}
	}
	if (!hasPhase) return;

	// Отводим верхние 2/7 высоты под фазу на обоих листах. Правая шкала
	// Chart при этом оставляет рабочий диапазон влажности 0…100% внизу.
	Excel::Axis^ primary = ExcelValueAxis(chart, false);
	try {
		double lo = primary->MinimumScale, hi = primary->MaximumScale;
		primary->MinimumScale = lo;
		primary->MaximumScale = hi + .4 * (hi - lo);
	}
	finally { Marshal::ReleaseComObject(primary); }
	String^ label = humidityChart ? L"Фаза: 110 прогрев / 120 плато / 130 финиш" :
		L"Фаза: 0 прогрев / 1 плато / 2 финиш";
	AddExcelMappedLine(lines, xRange, sheet, column, values, label, Color::Indigo, true, 2.4);
	Excel::Axis^ right = ExcelValueAxis(chart, true);
	try {
		right->MinimumScale = humidityChart ? 0.0 : -11.0;
		right->MaximumScale = humidityChart ? 140.0 : 3.0;
		right->MajorUnit = humidityChart ? 10.0 : 1.0;
		if (!humidityChart) right->TickLabels->NumberFormat = L"[>=0]0;[<0]\"\"";
		SetExcelAxisTitle(right, humidityChart ? L"Влажность, % / фаза" : L"Фаза");
	}
	finally { Marshal::ReleaseComObject(right); }
}

static double AddExcelPidLines(Excel::Chart^ chart, Excel::SeriesCollection^ lines, Excel::Range^ xRange,
	Excel::Worksheet^ sheet, System::Data::DataTable^ table, int helperColumn)
{
	auto columns = gcnew cli::array<String^>{"uCommon_TEN","PID_P","PID_I","PID_D","PID_FF"};
	auto names = gcnew cli::array<String^>{"uCommon","P","I","D","FF"};
	auto colors = gcnew cli::array<Color>{Color::Black,Color::Brown,Color::Green,Color::Red,Color::Blue};
	double minimum = 0, maximum = 2;
	for each (System::Data::DataRow^ row in table->Rows)
		for each (String^ column in columns) {
			double value = ChartValue(row, column);
			if (!Double::IsNaN(value)) { minimum = Math::Min(minimum,value); maximum = Math::Max(maximum,value); }
		}
	// Одинаковый интервал между кривыми даёт отдельные полосы, включая
	// отрицательные P/I/D. В легенде видно, какое смещение вычесть.
	const double spacing = PidPlotSpacing(minimum, maximum);
	double plotMinimum=0, plotMaximum=4*spacing;
	for (int k = 0; k < columns->Length; ++k) {
		if (!table->Columns->Contains(columns[k])) continue;
		auto values = gcnew cli::array<Object^, 2>(table->Rows->Count, 1);
		const double offset = k * spacing;
		for (int i = 0; i < table->Rows->Count; ++i) {
			double value = ChartValue(table->Rows[i], columns[k]);
			if (!Double::IsNaN(value)) {
				const double shifted = value + offset;
				values[i,0] = shifted;
				plotMinimum = Math::Min(plotMinimum, shifted);
				plotMaximum = Math::Max(plotMaximum, shifted);
			}
		}
		AddExcelMappedLine(lines, xRange, sheet, helperColumn+k, values,
			String::Format(L"{0} (+{1:0})", names[k], offset), colors[k], false, 1.8);
	}
	if (lines->Count > 0) {
		Excel::Axis^ axis = ExcelValueAxis(chart,false);
		try {
			axis->MinimumScale = spacing * Math::Floor(plotMinimum/spacing);
			axis->MaximumScale = spacing * Math::Ceiling(plotMaximum/spacing);
			axis->MajorUnit = spacing;
			// Вспомогательные линии делят основную полосу на четыре равных интервала.
			axis->MinorUnit = spacing / PidGridSubdivisions;
			axis->HasMajorGridlines = true;
			axis->HasMinorGridlines = true;
			axis->MajorGridlines->Format->Line->ForeColor->RGB = ColorTranslator::ToOle(Color::FromArgb(145,156,173));
			axis->MajorGridlines->Format->Line->Weight = 1.2f;
			axis->MinorGridlines->Format->Line->ForeColor->RGB = ColorTranslator::ToOle(Color::FromArgb(228,233,241));
			axis->MinorGridlines->Format->Line->Weight = .65f;
		}
		finally { Marshal::ReleaseComObject(axis); }
	}
	return spacing;
}

static void BuildExcelPidPanel(int panel, Excel::Chart^ chart, Excel::SeriesCollection^ lines,
	Excel::Range^ xRange, Excel::Worksheet^ source, Excel::Worksheet^ sheet,
	System::Data::DataTable^ table, int% helperColumn)
{
	const int rowCount = table->Rows->Count;
	switch (panel) {
	case 0: {
		AddExcelSourceLine(lines,xRange,source,table,"T0",L"Сырая T слева",Color::DarkGray,false,1.2);
		AddExcelSourceLine(lines,xRange,source,table,"T1",L"Сырая T справа",Color::DarkGray,false,1.2);
		AddExcelSourceLine(lines,xRange,source,table,"T_filt_0",L"Фильтр 32 слева",Color::FromArgb(138,174,152),false,1.4);
		AddExcelSourceLine(lines,xRange,source,table,"T_filt_1",L"Фильтр 32 справа",Color::FromArgb(193,151,139),false,1.4);
		AddExcelSourceLine(lines,xRange,source,table,"T_sup_avg_C",L"Средняя T регулятора",Color::RoyalBlue,false,2.8);
		auto target = gcnew cli::array<Object^, 2>(rowCount,1);
		for (int i = 0; i < rowCount; ++i) {
			auto row = table->Rows[i];
			double average = ChartValue(row,"T_sup_avg_C");
			if (Double::IsNaN(average) && ChartValue(row,"Act0")!=0 && ChartValue(row,"Act1")!=0)
				average = (ChartValue(row,"T_filt_0")+ChartValue(row,"T_filt_1"))/2;
			double setpoint = average+ChartValue(row,"eT_common");
			if (!Double::IsNaN(setpoint)) target[i,0] = setpoint;
		}
		AddExcelMappedLine(lines,xRange,sheet,helperColumn++,target,L"Уставка",Color::Black,false,2.8);
		break;
	}
	case 1:
		AddExcelSourceLine(lines,xRange,source,table,"T3",L"Сырая T продукта слева",Color::ForestGreen,false,1.8);
		AddExcelSourceLine(lines,xRange,source,table,"T4",L"Сырая T продукта справа",Color::DarkViolet,false,1.8);
		break;
	case 2: {
		auto columns = gcnew cli::array<String^>{"Ten1_Left","Ten2_Left","Ten1_Right","Ten2_Right"};
		auto names = gcnew cli::array<String^>{L"Слева ТЭН1: 0/1",L"Слева ТЭН2: 1/2",L"Справа ТЭН1: 2/3",L"Справа ТЭН2: 3/4"};
		auto colors = gcnew cli::array<Color>{Color::DarkGreen,Color::MediumSeaGreen,Color::SaddleBrown,Color::Firebrick};
		for (int k = 0; k < columns->Length; ++k) {
			if (!table->Columns->Contains(columns[k])) continue;
			auto values = gcnew cli::array<Object^, 2>(rowCount,1);
			for (int i = 0; i < rowCount; ++i) {
				auto row = table->Rows[i];
				double feedback = ChartValue(row,columns[k]);
				double ioType = ChartValue(row,"Typ6");
				if (ChartValue(row,"Act6")!=0 && (Double::IsNaN(ioType) || ioType==4) && !Double::IsNaN(feedback))
					values[i,0] = k + (feedback>0 ? 1 : 0);
			}
			AddExcelMappedLine(lines,xRange,sheet,helperColumn++,values,names[k],colors[k],false,1.8);
		}
		if (lines->Count > 0) {
			Excel::Axis^ axis = ExcelValueAxis(chart,false);
			try { axis->MinimumScale=-1; axis->MaximumScale=5; axis->MajorUnit=1; }
			finally { Marshal::ReleaseComObject(axis); }

		}
		break;
	}
	case 3: {
		const double spacing = AddExcelPidLines(chart,lines,xRange,sheet,table,helperColumn);
		helperColumn += 5;
		AddExcelPhase(chart,lines,xRange,sheet,table,helperColumn++,false);
		if (lines->Count > 0) {
			// Верхняя полоса фазы тоже заканчивается на делении сетки ПИД.
			Excel::Axis^ axis = ExcelValueAxis(chart,false);
			try { axis->MaximumScale = spacing * Math::Ceiling(axis->MaximumScale/spacing); }
			finally { Marshal::ReleaseComObject(axis); }
		}
		break;
	}
	case 4: {
		AddExcelSourceLine(lines,xRange,source,table,"H2",L"Влажность возврата, %",Color::RoyalBlue,false,1.8);
		auto duty = gcnew cli::array<Object^, 2>(rowCount,1);
		auto water = gcnew cli::array<Object^, 2>(rowCount,1);
		for (int i = 0; i < rowCount; ++i) {
			double requested = ChartValue(table->Rows[i],"PID_waterDuty");
			double active = ChartValue(table->Rows[i],"_Inj");
			if (!Double::IsNaN(requested)) duty[i,0] = requested*100;
			if (!Double::IsNaN(active)) water[i,0] = active>0 ? 1 : 0;
		}
		AddExcelMappedLine(lines,xRange,sheet,helperColumn++,duty,L"Запрос воды, %",Color::SeaGreen,false,1.8);
		AddExcelMappedLine(lines,xRange,sheet,helperColumn++,water,L"Впрыск: 0/1 (правая шкала)",Color::DodgerBlue,true,1.8);
		Excel::Axis^ left = ExcelValueAxis(chart,false);
		Excel::Axis^ right = ExcelValueAxis(chart,true);
		try {
			left->MinimumScale=0; left->MaximumScale=100; left->MajorUnit=25;
			right->MinimumScale=0; right->MaximumScale=1; right->MajorUnit=1;
			SetExcelAxisTitle(right,L"Впрыск");
		}
		finally { Marshal::ReleaseComObject(left); Marshal::ReleaseComObject(right); }
		break;
	}
	case 5:
		AddExcelSourceLine(lines,xRange,source,table,"heatScale01",L"heatScale",Color::MediumPurple,false,1.8);
		break;
	case 6:
		AddExcelSourceLine(lines,xRange,source,table,"uLeft_TEN",L"Запрос слева",Color::SeaGreen,false,1.8);
		AddExcelSourceLine(lines,xRange,source,table,"uRight_TEN",L"Запрос справа",Color::Firebrick,false,1.8);
		break;
	}
}

static void BuildExcelCharts(Excel::Worksheet^ source, System::Data::DataTable^ table, bool pid)
{
	auto missing = Type::Missing;
	Excel::Workbook^ book = safe_cast<Excel::Workbook^>(source->Parent);
	Excel::Worksheet^ sheet = nullptr;
	Excel::ChartObjects^ objects = nullptr;
	Excel::Range^ xRange = nullptr;
	try {
		sheet = safe_cast<Excel::Worksheet^>(book->Worksheets->Add(missing,source,1,Excel::XlSheetType::xlWorksheet));
		sheet->Name = pid ? L"ПИД" : L"Chart";
		if (table->Rows->Count == 0) return;
		int xIndex = table->Columns->IndexOf("RealTime");
		if (xIndex < 0) xIndex = Math::Max(0,table->Columns->IndexOf("Time"));
		xRange = source->Range[source->Cells[2,xIndex+1],source->Cells[table->Rows->Count+1,xIndex+1]];
		objects = safe_cast<Excel::ChartObjects^>(sheet->ChartObjects(missing));
		auto titles = pid ? gcnew cli::array<String^>{
			L"Температура воздуха",L"Температура продукта",L"Подтверждённые ТЭНы",
			L"Общий запрос нагрева и составляющие ПИД",L"Влажность возврата, запрос воды и впрыск",
			L"Разрешённая доля нагрева",L"Запрос нагрева по сторонам"
		} : gcnew cli::array<String^>{L"Температуры, влажность, оборудование и фаза"};
		auto units = pid ? gcnew cli::array<String^>{
			L"°C",L"°C",L"Уровни ТЭНов",L"Экв. ТЭНа со смещением",L"%",L"Доля нагрева",L"Экв. ТЭНа"
		} : gcnew cli::array<String^>{L"T, °C / уровни оборудования"};
		int helperColumn = 24;
		double top = 20;
		for (int panel = 0; panel < titles->Length; ++panel) {
			// Общий график ПИД выше остальных. Следующий график начинается ниже его границы.
			const double height = pid ? (panel==3 ? 480 : 320) : 500;
			Excel::ChartObject^ object = objects->Add(20,top,900,height);
			top += height+18;
			Excel::Chart^ chart = object->Chart;
			Excel::SeriesCollection^ lines = nullptr;
			try {
				chart->ChartType = Excel::XlChartType::xlLine;
				chart->DisplayBlanksAs = Excel::XlDisplayBlanksAs::xlNotPlotted;
				lines = safe_cast<Excel::SeriesCollection^>(chart->SeriesCollection(missing));
				if (!pid) {
					auto columns = gcnew cli::array<String^>{"T0","T1","T2","T3","T4","T5","_H0","_H1","_H2","_H3","H0","H1","H2","_Inj","T_filt_4"};
					for each (String^ col in columns)
						AddExcelSourceLine(lines,xRange,source,table,col,GetExcelChartSeriesName(col),
							ColorTranslator::FromOle(GetExcelSeriesColorOle(col)),col=="H0"||col=="H1"||col=="H2",1.8);
					if (table->Columns->Contains("H0") || table->Columns->Contains("H1") || table->Columns->Contains("H2")) {
						Excel::Axis^ right = ExcelValueAxis(chart,true);
						try { right->MinimumScale=0; right->MaximumScale=100; SetExcelAxisTitle(right,L"Влажность, %"); }
						finally { Marshal::ReleaseComObject(right); }
					}
					AddExcelPhase(chart,lines,xRange,sheet,table,helperColumn++,true);
				} else {
					BuildExcelPidPanel(panel,chart,lines,xRange,source,sheet,table,helperColumn);
				}
				chart->HasTitle = true;
				chart->ChartTitle->Text = titles[panel];
				if (lines->Count != 0) {
					Excel::Axis^ axis = ExcelValueAxis(chart,false);
					try { SetExcelAxisTitle(axis,units[panel]); }
					finally { Marshal::ReleaseComObject(axis); }
					chart->HasLegend = true;
					chart->Legend->Position = Excel::XlLegendPosition::xlLegendPositionBottom;
					Excel::Axis^ timeAxis = safe_cast<Excel::Axis^>(
						chart->Axes(Excel::XlAxisType::xlCategory,Excel::XlAxisGroup::xlPrimary));
					try { SetExcelAxisTitle(timeAxis,table->Columns->Contains("RealTime") ? L"Время" : L"Отсчёт"); }
					finally { Marshal::ReleaseComObject(timeAxis); }
				}
			}
			finally {
				if (lines != nullptr) Marshal::ReleaseComObject(lines);
				Marshal::ReleaseComObject(chart);
				Marshal::ReleaseComObject(object);
			}
		}
	}
	finally {
		if (objects != nullptr) Marshal::ReleaseComObject(objects);
		if (xRange != nullptr) Marshal::ReleaseComObject(xRange);
		if (sheet != nullptr) Marshal::ReleaseComObject(sheet);
		Marshal::ReleaseComObject(book);
	}
}

// Сохраняем снимок WorkData числовыми ячейками, чтобы его можно было
// загрузить обратно независимо от региональных настроек Excel.
static void WriteExcelPidParameters(Excel::Worksheet^ source, System::Data::DataTable^ snapshot)
{
	Excel::Workbook^ book=safe_cast<Excel::Workbook^>(source->Parent);
	Excel::Worksheet^ sheet=nullptr;
	Excel::Range^ range=nullptr;
	try {
		sheet=safe_cast<Excel::Worksheet^>(book->Worksheets->Add(Type::Missing,source,1,Excel::XlSheetType::xlWorksheet));
		sheet->Name=L"Параметры ПИД";
		int rowCount=snapshot->Rows->Count;
		auto values=gcnew cli::array<Object^,2>(rowCount+1,3);
		for(int col=0;col<3;++col)values[0,col]=snapshot->Columns[col]->ColumnName;
		for(int row=0;row<rowCount;++row)
			for(int col=0;col<3;++col)values[row+1,col]=snapshot->Rows[row][col];
		range=sheet->Range[sheet->Cells[1,1],sheet->Cells[rowCount+1,3]];
		range->Value2=values;
		sheet->Range[L"A1",L"C1"]->Font->Bold=true;
		if(rowCount==0)
			sheet->Cells[2,1]=L"Параметры ПИД не получены от контроллера.";
		sheet->Columns->AutoFit();
	}
	finally {
		if(range!=nullptr)Marshal::ReleaseComObject(range);
		if(sheet!=nullptr)Marshal::ReleaseComObject(sheet);
		Marshal::ReleaseComObject(book);
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

			// Диаграммы обоих листов используют один и тот же снимок экспорта.
			for each (bool pid in gcnew cli::array<bool>{false,true}) {
				try { BuildExcelCharts(ws,job->tableSnapshot,pid); }
				catch (Exception^ ex) {
					GlobalLogger::LogMessage(String::Format(L"Не удалось построить лист {0}: {1}",pid ? L"ПИД" : L"Chart",ex->Message));
				}
			}
			if (job->paramsPid != nullptr) {
				try { WriteExcelPidParameters(ws,job->paramsPid); }
				catch (Exception^ ex) {
					GlobalLogger::LogMessage(L"Не удалось сохранить лист «Параметры ПИД»: "+ex->Message);
				}
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


