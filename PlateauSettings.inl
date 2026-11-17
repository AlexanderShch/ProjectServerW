/* Included by CommandsDefroster.cpp; operations use the existing command/response pipeline. */
#include "PlateauParams.h"
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

void ProjectServerW::DataForm::ReadPlateauSettings(System::Object^, System::EventArgs^) {
 if(plateauSettingsTab==nullptr||plateauSettingsTab->IsDisposed)return;
 uint8_t raw[PLATEAU_PARAM_BYTES]={};uint8_t length=0;
 const bool ok=GetDefrostGroup(7u,0u,raw,sizeof(raw),&length)&&length==sizeof(raw);
 if(!ok) {
  plateauSettingsGrid->Rows->Clear();
  plateauSettingsStatus->Text=L"Группа 7 недоступна: проверьте связь и версию прошивки.";
 } else {
  plateauSettingsGrid->Rows->Clear();
  auto inv=System::Globalization::CultureInfo::InvariantCulture;
  for(unsigned i=1;i<PLATEAU_PARAM_COUNT;++i) {
   float value;memcpy(&value,raw+4*i,4);
   float lo=kPlateauFields[i].minimum,hi=kPlateauFields[i].maximum;
   plateauSettingsGrid->Rows->Add(gcnew String(kPlateauFields[i].name),gcnew String(kPlateauDescriptions[i]),
    value.ToString("R",inv),lo.ToString(inv)+L" … "+hi.ToString(inv));
  }
  plateauSettingsStatus->Text=L"Прочитаны актуальные ПИД-параметры из контроллера.";
 }
 buttonPlateauApply->Enabled=ok;
}

void ProjectServerW::DataForm::ApplyPlateauSettings(System::Object^, System::EventArgs^) {
 if(plateauSettingsTab==nullptr||plateauSettingsTab->IsDisposed||plateauSettingsGrid->Rows->Count!=(PLATEAU_PARAM_COUNT-1u))return;
 plateauSettingsGrid->EndEdit();
 uint8_t raw[PLATEAU_PARAM_BYTES]={};
 const float reservedAlgorithm=1.0f;memcpy(raw,&reservedAlgorithm,sizeof(reservedAlgorithm));
 auto inv=System::Globalization::CultureInfo::InvariantCulture;
 for(unsigned i=1;i<PLATEAU_PARAM_COUNT;++i) {
  auto cell=plateauSettingsGrid->Rows[i-1u]->Cells[2]->Value;
  float value=0;String^ text=cell==nullptr?"":cell->ToString()->Trim()->Replace(L',',L'.');
  if(!Single::TryParse(text,System::Globalization::NumberStyles::Float,inv,value)||
   Single::IsNaN(value)||Single::IsInfinity(value)||value<kPlateauFields[i].minimum||value>kPlateauFields[i].maximum) {
   plateauSettingsStatus->Text=L"Неверное значение: "+gcnew String(kPlateauFields[i].name);return;
  }
  memcpy(raw+4*i,&value,4);
 }
 PlateauParams p;memcpy(&p,raw,sizeof(p));
 if(p.pulseMin_s>p.pulseMax_s||
  floorf(p.rateWindow_s)!=p.rateWindow_s||floorf(p.pulseMax_s)!=p.pulseMax_s||
  floorf(p.pulseMin_s)!=p.pulseMin_s||floorf(p.pauseMin_s)!=p.pauseMin_s) {
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
