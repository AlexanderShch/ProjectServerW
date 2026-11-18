/* Included by CommandsDefroster.cpp; operations use the existing command/response pipeline. */
#include "PlateauParams.h"
#include "FormExcel.h"
#include <cmath>
#include <cstring>
#include "PidGraphForm.h"
namespace {
struct PlateauFieldSpec { const char* name; float initial, minimum, maximum; };
static const PlateauFieldSpec kPlateauFields[] = {
#define PLATEAU_SPEC(name, def, lo, hi) {#name, def, lo, hi},
 PLATEAU_FIELDS(PLATEAU_SPEC)
#undef PLATEAU_SPEC
};
static_assert(sizeof(PlateauParams)==80, "Group 7 wire size");
static const wchar_t* kPlateauDescriptions[] = {
 L"", // reserved wire slot 0: not displayed and never selects a controller
 L"Температура: Kp (Пропорциональный), экв./°C", L"Температура: Ki (Интегральный), экв./(°C·с)",
 L"Температура: Kd (Дифференциальный), экв.·с/°C", L"Окно скорости температуры, целые секунды",
 L"Фильтр скорости температуры, τ, с", L"Предел D (дифф.поправки), ± экв.",
 L"Максимальная компенсация впрыска, экв.", L"Затухающий хвост компенсации впрыска, с",
 L"Полоса разрешения компенсации впрыска ± °C от уставки",
 L"Влажность возврата: Kp (Пропорциональный), доля/(п.п. RH)", L"Влажность: Ki (Интегральный), доля/(п.п. RH·с)",
 L"Влажность: Kd (Дифференциальный), доля·с/(п.п. RH); 0 отключает D", L"Фильтр RH, τ, с",
 L"Максимальная длина импульса воды, целые с", L"Минимальная пауза между импульсами, целые с",
 L"Минимальная длина импульса воды, целые с", L"Максимальная доля времени подачи воды, 0…1",
 L"Зона нечувствительности ПИД влажности, п.п. RH", L"Превышение RH для прекращения впрыска, п.п."
};
}

// All controls are created in InitializeComponent so the tab is editable in the designer.
// Only runtime state rows belong here.
void ProjectServerW::DataForm::InitializePlateauTab() {
 plateauStateGrid->Rows->Add(L"Время строки основного лога",L"—",L"время сервера");
 plateauStateGrid->Rows->Add(L"Общий запрос после heatScale",L"—",L"экв. ТЭНа на сторону");
 // P берётся из того же такта основного лога, что и I/D/FF.
 plateauStateGrid->Rows->Add(L"Пропорциональная составляющая P",L"—",L"экв.");
 plateauStateGrid->Rows->Add(L"Интегратор I",L"—",L"экв.");
 plateauStateGrid->Rows->Add(L"Поправка D",L"—",L"экв.");
 plateauStateGrid->Rows->Add(L"Компенсация впрыска FF",L"—",L"экв.");
 plateauStateGrid->Rows->Add(L"Запрошенная доля воды",L"—",L"0…1");
 plateauStateGrid->Rows->Add(L"Форсунка",L"—",L"0 / 1");
 plateauStateGrid->Rows->Add(L"ТЭН 1 слева",L"—",L"подтверждение DI");
 plateauStateGrid->Rows->Add(L"ТЭН 2 слева",L"—",L"подтверждение DI");
 plateauStateGrid->Rows->Add(L"ТЭН 1 справа",L"—",L"подтверждение DI");
 plateauStateGrid->Rows->Add(L"ТЭН 2 справа",L"—",L"подтверждение DI");
}

void ProjectServerW::DataForm::ShowPlateauInfo(System::Object^, System::EventArgs^) {
 System::Windows::Forms::MessageBox::Show(this,
  L"Единый ПИД температуры и влажности работает на всех фазах: прогрев, плато и финиш.\r\nЗапись применяется на следующем такте без остановки процесса.\r\n"
  L"Параметры записываются в EEPROM контроллера и сохраняются после отключения питания.\r\n"
  L"Сначала считайте актуальные параметры. Изменённые значения отправляются только кнопкой «Записать».",
  L"Информация — ПИД",
  System::Windows::Forms::MessageBoxButtons::OK,
  System::Windows::Forms::MessageBoxIcon::Information);
}

namespace {
namespace Excel = Microsoft::Office::Interop::Excel;
using System::Runtime::InteropServices::Marshal;

// Источники используют одну проверку, чтобы импорт и кнопка «Записать»
// принимали один и тот же набор значений и ограничений группы 7.
static bool TryParsePlateauValue(unsigned index,Object^ cell,float% value) {
 auto inv=System::Globalization::CultureInfo::InvariantCulture;
 String^ text=Convert::ToString(cell,inv)->Trim()->Replace(L',',L'.');
 return Single::TryParse(text,System::Globalization::NumberStyles::Float,inv,value)&&
  !Single::IsNaN(value)&&!Single::IsInfinity(value)&&
  value>=kPlateauFields[index].minimum&&value<=kPlateauFields[index].maximum;
}

static bool ValidPlateauTiming(const PlateauParams& p) {
 return p.pulseMin_s<=p.pulseMax_s&&floorf(p.rateWindow_s)==p.rateWindow_s&&
  floorf(p.pulseMax_s)==p.pulseMax_s&&floorf(p.pulseMin_s)==p.pulseMin_s&&
  floorf(p.pauseMin_s)==p.pauseMin_s;
}

static PlateauParams PlateauFromValues(cli::array<float>^ values) {
 PlateauParams result={};
 pin_ptr<float> pinned=&values[0];
 memcpy(&result,pinned,sizeof(result));
 return result;
}
}

void ProjectServerW::DataForm::PopulatePlateauSettings(cli::array<float>^ values) {
 auto inv=System::Globalization::CultureInfo::InvariantCulture;
 plateauSettingsGrid->Rows->Clear();
 for(unsigned i=1;i<PLATEAU_PARAM_COUNT;++i) {
  float value=values[i];
  float lo=kPlateauFields[i].minimum,hi=kPlateauFields[i].maximum;
  plateauSettingsGrid->Rows->Add(gcnew String(kPlateauFields[i].name),gcnew String(kPlateauDescriptions[i]),
   value.ToString("R",inv),lo.ToString(inv)+L" … "+hi.ToString(inv));
 }
 buttonPlateauApply->Enabled=true;
}

void ProjectServerW::DataForm::ReadPlateauSettings(System::Object^, System::EventArgs^) {
 if(plateauSettingsTab==nullptr||plateauSettingsTab->IsDisposed)return;
 uint8_t raw[PLATEAU_PARAM_BYTES]={};uint8_t length=0;
 if(!GetDefrostGroup(7u,0u,raw,sizeof(raw),&length)||length!=sizeof(raw)) {
  plateauSettingsGrid->Rows->Clear();
  buttonPlateauApply->Enabled=false;
  plateauSettingsStatus->Text=L"Группа 7 недоступна: проверьте связь и версию прошивки.";
  GlobalLogger::LogMessage(L"Параметры ПИД: группа 7 недоступна (бюджет запроса до 1 с).");
  return;
 }
 auto values=gcnew cli::array<float>(PLATEAU_PARAM_COUNT);
 pin_ptr<float> pinned=&values[0];
 memcpy(pinned,raw,sizeof(raw));
 PopulatePlateauSettings(values);
 plateauSettingsStatus->Text=L"Прочитаны актуальные ПИД-параметры из контроллера.";
 GlobalLogger::LogMessage(L"Параметры ПИД: группа 7 считана.");
}

System::Data::DataTable^ ProjectServerW::DataForm::BuildPlateauParamsSnapshot() {
 auto snapshot=gcnew System::Data::DataTable();
 snapshot->Columns->Add(L"Параметр",String::typeid);
 snapshot->Columns->Add(L"Описание",String::typeid);
 snapshot->Columns->Add(L"Значение",double::typeid);
 // Снимаем подтверждённую конфигурацию в момент постановки WorkData
 // в очередь. Редактируемая таблица ПИД здесь может содержать черновик.
 if(isFormClosingNow||ClientSocket==INVALID_SOCKET||
  (radioButtonSiemens!=nullptr&&radioButtonSiemens->Checked))return snapshot;
 uint8_t raw[PLATEAU_PARAM_BYTES]={};uint8_t length=0;
 try {
  if(GetDefrostGroup(7u,0u,raw,sizeof(raw),&length)&&length==sizeof(raw)) {
   for(unsigned i=1;i<PLATEAU_PARAM_COUNT;++i) {
    float value;memcpy(&value,raw+4*i,4);
    snapshot->Rows->Add(gcnew String(kPlateauFields[i].name),gcnew String(kPlateauDescriptions[i]),
     static_cast<double>(value));
   }
   return snapshot;
  }
 } catch(Exception^ ex) {
  snapshot->Rows->Clear();
  GlobalLogger::LogMessage(L"Ошибка чтения параметров ПИД для WorkData: "+ex->Message);
 }
 GlobalLogger::LogMessage(L"Предупреждение: параметры ПИД для WorkData не получены от контроллера.");
 return snapshot;
}

void ProjectServerW::DataForm::LoadPlateauSettings(Object^,EventArgs^) {
 auto dialog=gcnew OpenFileDialog();
 try {
  dialog->Title=L"Выберите файл WorkData для загрузки параметров ПИД";
  dialog->Filter=L"Файлы WorkData (WorkData*.xlsx)|WorkData*.xlsx|Файлы Excel (*.xlsx)|*.xlsx";
  dialog->DefaultExt=L"xlsx";
  dialog->RestoreDirectory=true;
  if(!String::IsNullOrEmpty(excelSavePath))dialog->InitialDirectory=excelSavePath;
  if(dialog->ShowDialog(this)!=System::Windows::Forms::DialogResult::OK)return;
  try { LoadPlateauSettingsFromFile(dialog->FileName); }
  catch(Exception^ ex) {
   MessageBox::Show(this,ex->Message,L"Загрузка параметров ПИД",
    MessageBoxButtons::OK,MessageBoxIcon::Warning);
  }
 } finally { delete dialog; }
}

void ProjectServerW::DataForm::LoadPlateauSettingsFromFile(String^ filePath) {
 if(!String::Equals(System::IO::Path::GetExtension(filePath),L".xlsx",StringComparison::OrdinalIgnoreCase))
  throw gcnew System::IO::InvalidDataException(L"Выберите файл WorkData в формате XLSX.");
 auto mutex=FormExcel::GetExcelGlobalMutex();
 if(!mutex->WaitOne(15000))throw gcnew TimeoutException(L"Дождитесь окончания экспорта Excel и повторите загрузку.");
 Excel::Application^ app=nullptr;
 Excel::Workbook^ book=nullptr;
 Excel::Worksheet^ sheet=nullptr;
 Excel::Range^ range=nullptr;
 auto values=gcnew cli::array<float>(PLATEAU_PARAM_COUNT);
 values[0]=1.0f;
 try {
  app=gcnew Excel::ApplicationClass();
  app->Visible=false;
  app->DisplayAlerts=false;
  // Свойство Excel AutomationSecurity=3 запрещает выполнение макросов
  // при программном открытии. Позднее связывание не добавляет ссылку на Office.dll.
  Type::GetTypeFromProgID(L"Excel.Application")->InvokeMember(L"AutomationSecurity",
   System::Reflection::BindingFlags::SetProperty,nullptr,app,gcnew cli::array<Object^>{3});
  Object^ m=Type::Missing;
  book=app->Workbooks->Open(filePath,safe_cast<Object^>(0),true,m,m,m,true,m,m,false,false,m,false,true,m);
  try { sheet=safe_cast<Excel::Worksheet^>(book->Worksheets->Item[L"Параметры ПИД"]); }
  catch(System::Runtime::InteropServices::COMException^) {
   throw gcnew System::IO::InvalidDataException(L"В файле нет листа «Параметры ПИД».");
  }
  range=sheet->Range[L"A1",L"C20"];
  auto cells=safe_cast<Array^>(range->Value2);
  if(!Object::Equals(cells->GetValue(1,1),L"Параметр")||!Object::Equals(cells->GetValue(1,3),L"Значение"))
   throw gcnew System::IO::InvalidDataException(L"В файле нет сохранённых значений ПИД.");
  auto seen=gcnew cli::array<bool>(PLATEAU_PARAM_COUNT);
  for(int row=2;row<=20;++row) {
   String^ name=Convert::ToString(cells->GetValue(row,1))->Trim();
   unsigned index=1;
   while(index<PLATEAU_PARAM_COUNT&&!String::Equals(name,gcnew String(kPlateauFields[index].name),StringComparison::Ordinal))++index;
   if(index==PLATEAU_PARAM_COUNT||seen[index])
    throw gcnew System::IO::InvalidDataException(String::Format(L"Строка {0}: отсутствует, неизвестен или повторяется параметр ПИД.",row));
   float value=0;
   if(!TryParsePlateauValue(index,cells->GetValue(row,3),value))
    throw gcnew System::IO::InvalidDataException(L"Недопустимое значение ПИД: "+name);
   seen[index]=true;
   values[index]=value;
  }
  if(!ValidPlateauTiming(PlateauFromValues(values)))
   throw gcnew System::IO::InvalidDataException(L"Проверьте целые тайминги и минимум/максимум импульса.");
 } finally {
  if(range!=nullptr)Marshal::ReleaseComObject(range);
  if(sheet!=nullptr)Marshal::ReleaseComObject(sheet);
  if(book!=nullptr) {
   try { book->Close(false,Type::Missing,Type::Missing); } catch(Exception^) {}
   Marshal::ReleaseComObject(book);
  }
  if(app!=nullptr) {
   try { app->Quit(); } catch(Exception^) {}
   Marshal::ReleaseComObject(app);
  }
  mutex->ReleaseMutex();
 }
 // Меняем таблицу только после проверки всего файла.
 PopulatePlateauSettings(values);
 plateauSettingsStatus->Text=L"Загружено из файла. Для применения нажмите «Записать».";
}

void ProjectServerW::DataForm::ApplyPlateauSettings(System::Object^, System::EventArgs^) {
 if(plateauSettingsTab==nullptr||plateauSettingsTab->IsDisposed||plateauSettingsGrid->Rows->Count!=(PLATEAU_PARAM_COUNT-1u))return;
 plateauSettingsGrid->EndEdit();
 uint8_t raw[PLATEAU_PARAM_BYTES]={};
 const float reservedAlgorithm=1.0f;memcpy(raw,&reservedAlgorithm,sizeof(reservedAlgorithm));
 auto inv=System::Globalization::CultureInfo::InvariantCulture;
 for(unsigned i=1;i<PLATEAU_PARAM_COUNT;++i) {
  auto cell=plateauSettingsGrid->Rows[i-1u]->Cells[2]->Value;
  float value=0;
  if(!TryParsePlateauValue(i,cell,value)) {
   plateauSettingsStatus->Text=L"Неверное значение: "+gcnew String(kPlateauFields[i].name);return;
  }
  memcpy(raw+4*i,&value,4);
 }
 PlateauParams p;memcpy(&p,raw,sizeof(p));
 if(!ValidPlateauTiming(p)) {
  plateauSettingsStatus->Text=L"Проверьте целые тайминги и минимум/максимум импульса.";return;
 }
 if(!SetDefrostGroup(7u,raw,sizeof(raw))) {
  plateauSettingsStatus->Text=L"Контроллер не подтвердил запись. Считайте параметры повторно.";return;
 }
 uint8_t verified[PLATEAU_PARAM_BYTES]={};uint8_t length=0;
 if(!GetDefrostGroup(7u,0u,verified,sizeof(verified),&length)||length!=sizeof(verified)||memcmp(raw,verified,sizeof(raw))!=0) {
  plateauSettingsStatus->Text=L"Запись отправлена, но контрольное чтение не совпало. Считайте повторно.";return;
 }
 plateauSettingsStatus->Text=L"Записано в EEPROM контроллера, подтверждено чтением. Применение без перезапуска.";
 GlobalLogger::LogMessage(L"Information: Plateau group 7 committed to controller EEPROM and verified without process restart");
 for(unsigned i=1;i<PLATEAU_PARAM_COUNT;++i) {
  float value;memcpy(&value,raw+4*i,4);
  GlobalLogger::LogMessage(L"Plateau: "+gcnew String(kPlateauFields[i].name)+L"="+value.ToString("R",inv));
 }
}

void ProjectServerW::DataForm::UpdatePlateauStateFromLog(System::Data::DataRow^ row) {
 if(plateauStateGrid==nullptr||plateauStateGrid->IsDisposed||row==nullptr)return;
 plateauStateGrid->Rows[0]->Cells[1]->Value=row["RealTime"];
 auto names=gcnew cli::array<String^>{"uCommon_TEN","PID_P","PID_I","PID_D","PID_FF","PID_waterDuty","_Inj"};
 auto inv=System::Globalization::CultureInfo::InvariantCulture;
 for(int i=0;i<names->Length;++i) {
  Object^ value=row->Table->Columns->Contains(names[i])?row[names[i]]:DBNull::Value;
  plateauStateGrid->Rows[i+1]->Cells[1]->Value=(value==nullptr||value==DBNull::Value)
   ? L"нет данных" : (names[i]=="_Inj" ? (Convert::ToBoolean(value)?L"1":L"0") : Convert::ToDouble(value).ToString("F3",inv));
 }
 UpdatePlateauHeaterStates(row);
}

void ProjectServerW::DataForm::UpdatePlateauHeaterStates(System::Data::DataRow^ row) {
 if(plateauStateGrid==nullptr||plateauStateGrid->IsDisposed||row==nullptr)return;
 // Use feedback inputs, not the _H0.._H3 output commands. Invalid IO is unknown, not OFF.
 const bool inputsValid=row->Table->Columns->Contains("Typ6")&&row->Table->Columns->Contains("Act6")&&
  !row->IsNull("Typ6")&&!row->IsNull("Act6")&&Convert::ToInt32(row["Typ6"])==4&&Convert::ToInt32(row["Act6"])!=0;
 auto names=gcnew cli::array<String^>{"Ten1_Left","Ten2_Left","Ten1_Right","Ten2_Right"};
 for(int i=0;i<names->Length;++i) {
  const bool known=inputsValid&&row->Table->Columns->Contains(names[i])&&!row->IsNull(names[i]);
  plateauStateGrid->Rows[8+i]->Cells[1]->Value=known
   ? (Convert::ToBoolean(row[names[i]])?L"1":L"0") : L"нет данных";
 }
}

void ProjectServerW::DataForm::ReadPlateauDiagnostics(System::Object^, System::EventArgs^) {
 if(plateauSettingsTab==nullptr||plateauSettingsTab->IsDisposed)return;
 System::Threading::Monitor::Enter(dataTableSync);
 try {
  if(dataTable->Rows->Count>0)UpdatePlateauStateFromLog(dataTable->Rows[dataTable->Rows->Count-1]);
  else plateauSettingsStatus->Text=L"Основной лог пока не получен.";
 } finally {System::Threading::Monitor::Exit(dataTableSync);}
}

void ProjectServerW::DataForm::ShowPlateauGraph(System::Object^, System::EventArgs^) {
 System::Data::DataTable^ snapshot;
 System::Threading::Monitor::Enter(dataTableSync);
 try {
  snapshot=dataTable->Copy();
 } finally {System::Threading::Monitor::Exit(dataTableSync);}
 if(snapshot->Rows->Count<2) {plateauSettingsStatus->Text=L"Для графика нужны хотя бы две записи текущего цикла.";return;}
 auto graph=gcnew PidGraphForm(snapshot,dataCollectionStartTime);
 graph->Show(this);
}
