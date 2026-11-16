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
 L"1 — ПИД Plateau включён; 0 — прежний алгоритм",
 L"Температура: Kp, экв./°C", L"Температура: Ki, экв./(°C·с)",
 L"Температура: Kd, экв.·с/°C", L"Окно скорости температуры, целые секунды",
 L"Фильтр скорости температуры, τ, с", L"Предел D, ± экв.",
 L"Максимальная компенсация впрыска, экв.", L"Затухающий хвост компенсации, с",
 L"Полоса разрешения компенсации ± °C от уставки",
 L"Влажность возврата: Kp, доля/(п.п. RH)", L"Влажность: Ki, доля/(п.п. RH·с)",
 L"Влажность: Kd, доля·с/(п.п. RH); 0 отключает D", L"Фильтр RH, τ, с",
 L"Максимальная длина импульса воды, целые с", L"Минимальная пауза между импульсами, целые с",
 L"Минимальная длина импульса воды, целые с", L"Максимальная доля времени подачи воды, 0…1",
 L"Зона нечувствительности ПИД влажности, п.п. RH", L"Превышение RH для прекращения впрыска, п.п."
};
}

void ProjectServerW::DataForm::InitializePlateauTab() {
 plateauSettingsTab = gcnew System::Windows::Forms::TabPage();
 plateauSettingsTab->Name = "tabPlateauPid";
 plateauSettingsTab->Text = L"ПИД Plateau";
 plateauSettingsTab->Padding = System::Windows::Forms::Padding(6);
 plateauSettingsTab->UseVisualStyleBackColor = true;
 auto info = gcnew System::Windows::Forms::Label();
 info->Dock = System::Windows::Forms::DockStyle::Top; info->Height=64;
 info->Text = L"Только Plateau. Запись применяется на следующем такте без остановки процесса.\r\n"
 L"Параметры записываются в EEPROM контроллера и сохраняются после отключения питания.\r\n"
 L"Сначала считайте актуальные параметры. Изменённые значения отправляются только кнопкой «Записать».";
 plateauSettingsGrid = gcnew System::Windows::Forms::DataGridView();
 plateauSettingsGrid->Dock=System::Windows::Forms::DockStyle::Fill;
 plateauSettingsGrid->AllowUserToAddRows=false; plateauSettingsGrid->AllowUserToDeleteRows=false;
 plateauSettingsGrid->AutoSizeColumnsMode=System::Windows::Forms::DataGridViewAutoSizeColumnsMode::Fill;
 plateauSettingsGrid->Columns->Add("Parameter","Параметр");
 plateauSettingsGrid->Columns->Add("Description","Назначение и единицы");
 plateauSettingsGrid->Columns->Add("Value","Значение");
 plateauSettingsGrid->Columns->Add("Bounds","Допустимый диапазон");
 for(int c=0;c<4;++c) {
  plateauSettingsGrid->Columns[c]->ReadOnly=(c!=2);
  plateauSettingsGrid->Columns[c]->SortMode=System::Windows::Forms::DataGridViewColumnSortMode::NotSortable;
 }
 plateauSettingsGrid->Columns[1]->FillWeight=230;
 auto footer=gcnew System::Windows::Forms::FlowLayoutPanel();
 footer->Dock=System::Windows::Forms::DockStyle::Bottom;footer->Height=95;
 auto read=gcnew System::Windows::Forms::Button();read->Text=L"Считать";read->Width=130;
 read->Click+=gcnew System::EventHandler(this,&DataForm::ReadPlateauSettings);
 auto apply=gcnew System::Windows::Forms::Button();apply->Text=L"Записать";apply->Width=130;
 apply->Enabled=false;apply->Name="ApplyPlateau";
 apply->Click+=gcnew System::EventHandler(this,&DataForm::ApplyPlateauSettings);
 plateauSettingsStatus=gcnew System::Windows::Forms::Label();plateauSettingsStatus->AutoSize=true;
 plateauSettingsStatus->Text=L"Не считано. Нужна прошивка с группой 7.";
 auto diag=gcnew System::Windows::Forms::Button();diag->Text=L"Обновить состояние";diag->Width=150;
 diag->Click+=gcnew System::EventHandler(this,&DataForm::ReadPlateauDiagnostics);
 auto graph=gcnew System::Windows::Forms::Button();graph->Text=L"График";graph->Width=110;
 graph->Click+=gcnew System::EventHandler(this,&DataForm::ShowPlateauGraph);
 footer->Controls->Add(read);footer->Controls->Add(apply);footer->Controls->Add(diag);
 footer->Controls->Add(graph);
 footer->SetFlowBreak(graph,true);
 plateauSettingsStatus->MaximumSize=System::Drawing::Size(1000,0);
 footer->Controls->Add(plateauSettingsStatus);
 auto split=gcnew System::Windows::Forms::SplitContainer();
 split->Dock=System::Windows::Forms::DockStyle::Fill;
 split->Orientation=System::Windows::Forms::Orientation::Horizontal;
 // Set a valid initial size before splitter distance; resized with the tab afterwards.
 split->Size=System::Drawing::Size(1050,600);split->SplitterDistance=360;
 split->Panel1->Controls->Add(plateauSettingsGrid);
 plateauStateGrid=gcnew System::Windows::Forms::DataGridView();
 plateauStateGrid->Dock=System::Windows::Forms::DockStyle::Fill;plateauStateGrid->ReadOnly=true;
 plateauStateGrid->AllowUserToAddRows=false;plateauStateGrid->AllowUserToDeleteRows=false;
 plateauStateGrid->AutoSizeColumnsMode=System::Windows::Forms::DataGridViewAutoSizeColumnsMode::Fill;
 plateauStateGrid->Columns->Add("Signal",L"Состояние ПИД");
 plateauStateGrid->Columns->Add("Current",L"Значение");plateauStateGrid->Columns->Add("Units",L"Единицы");
 for each(System::Windows::Forms::DataGridViewColumn^ col in plateauStateGrid->Columns)
  col->SortMode=System::Windows::Forms::DataGridViewColumnSortMode::NotSortable;
 plateauStateGrid->Rows->Add(L"Время строки основного лога",L"—",L"время сервера");
 plateauStateGrid->Rows->Add(L"Общий запрос после heatScale",L"—",L"экв. ТЭНа на сторону");
 plateauStateGrid->Rows->Add(L"Интегратор I",L"—",L"экв.");
 plateauStateGrid->Rows->Add(L"Поправка D",L"—",L"экв.");
 plateauStateGrid->Rows->Add(L"Компенсация впрыска FF",L"—",L"экв.");
 plateauStateGrid->Rows->Add(L"Запрошенная доля воды",L"—",L"0…1");
 plateauStateGrid->Rows->Add(L"Форсунка",L"—",L"0 / 1");
 split->Panel2->Controls->Add(plateauStateGrid);
 plateauSettingsTab->Controls->Add(split);
 plateauSettingsTab->Controls->Add(info);plateauSettingsTab->Controls->Add(footer);
 tabControl1->TabPages->Add(plateauSettingsTab);
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
  for(unsigned i=0;i<PLATEAU_PARAM_COUNT;++i) {
   float value;memcpy(&value,raw+4*i,4);
   float lo=kPlateauFields[i].minimum,hi=kPlateauFields[i].maximum;
   plateauSettingsGrid->Rows->Add(gcnew String(kPlateauFields[i].name),gcnew String(kPlateauDescriptions[i]),
    value.ToString("R",inv),lo.ToString(inv)+L" … "+hi.ToString(inv));
  }
  plateauSettingsStatus->Text=L"Прочитаны актуальные ПИД-параметры из контроллера.";
 }
 auto controls=plateauSettingsTab->Controls;
 for each(System::Windows::Forms::Control^ panel in controls)
  for each(System::Windows::Forms::Control^ item in panel->Controls)
   if(item->Name=="ApplyPlateau")item->Enabled=ok;
}

void ProjectServerW::DataForm::ApplyPlateauSettings(System::Object^, System::EventArgs^) {
 if(plateauSettingsTab==nullptr||plateauSettingsTab->IsDisposed||plateauSettingsGrid->Rows->Count!=PLATEAU_PARAM_COUNT)return;
 plateauSettingsGrid->EndEdit();
 uint8_t raw[PLATEAU_PARAM_BYTES]={};
 auto inv=System::Globalization::CultureInfo::InvariantCulture;
 for(unsigned i=0;i<PLATEAU_PARAM_COUNT;++i) {
  auto cell=plateauSettingsGrid->Rows[i]->Cells[2]->Value;
  float value=0;String^ text=cell==nullptr?"":cell->ToString()->Trim()->Replace(',','.');
  if(!Single::TryParse(text,System::Globalization::NumberStyles::Float,inv,value)||
   Single::IsNaN(value)||Single::IsInfinity(value)||value<kPlateauFields[i].minimum||value>kPlateauFields[i].maximum) {
   plateauSettingsStatus->Text=L"Неверное значение: "+gcnew String(kPlateauFields[i].name);return;
  }
  memcpy(raw+4*i,&value,4);
 }
 PlateauParams p;memcpy(&p,raw,sizeof(p));
 if((p.enabled!=0&&p.enabled!=1)||p.pulseMin_s>p.pulseMax_s||
  floorf(p.rateWindow_s)!=p.rateWindow_s||floorf(p.pulseMax_s)!=p.pulseMax_s||
  floorf(p.pulseMin_s)!=p.pulseMin_s||floorf(p.pauseMin_s)!=p.pauseMin_s) {
  plateauSettingsStatus->Text=L"Проверьте целые тайминги, 0/1 и минимум/максимум импульса.";return;
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
 for(unsigned i=0;i<PLATEAU_PARAM_COUNT;++i) {
  float value;memcpy(&value,raw+4*i,4);
  GlobalLogger::LogMessage(L"Plateau: "+gcnew String(kPlateauFields[i].name)+L"="+value.ToString("R",inv));
 }
}

void ProjectServerW::DataForm::UpdatePlateauStateFromLog(System::Data::DataRow^ row) {
 if(plateauStateGrid==nullptr||plateauStateGrid->IsDisposed||row==nullptr)return;
 plateauStateGrid->Rows[0]->Cells[1]->Value=row["RealTime"];
 auto names=gcnew cli::array<String^>{"uCommon_TEN","PID_I","PID_D","PID_FF","PID_waterDuty","_Inj"};
 auto inv=System::Globalization::CultureInfo::InvariantCulture;
 for(int i=0;i<names->Length;++i) {
  Object^ value=row[names[i]];
  plateauStateGrid->Rows[i+1]->Cells[1]->Value=(value==nullptr||value==DBNull::Value)
   ? L"нет данных" : Convert::ToDouble(value).ToString("F3",inv);
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
