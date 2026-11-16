#pragma once
#using <System.dll>
#using <System.Data.dll>
#using <System.Drawing.dll>
#using <System.Windows.Forms.dll>

namespace ProjectServerW {
using namespace System;
using namespace System::Data;
using namespace System::Drawing;
using namespace System::Windows::Forms;
using namespace System::Collections::Generic;

ref class PidPlotSeries {
public:
 int panel;String^ name;Color color;bool step;
 double offset,plotMin,plotMax;
 List<double>^ x;List<double>^ y;
 PidPlotSeries(int p,String^ n,Color c,bool s) {
  panel=p;name=n;color=c;step=s;offset=0;plotMin=0;plotMax=0;x=gcnew List<double>();y=gcnew List<double>();
 }
 void Add(DateTime at,double value) {x->Add(at.ToOADate());y->Add(value);}
};

ref class PidPlotCanvas : public Panel {
public:
 PidPlotCanvas() {DoubleBuffered=true;ResizeRedraw=true;BackColor=Color::White;}
};

/* Read-only snapshot window. No controller commands, Excel automation or reconstructed I. */
public ref class PidGraphForm : public Form {
 DataTable^ source;
 List<DateTime>^ rowTimes;List<DataRow^>^ sourceRows;
 List<PidPlotSeries^>^ series;
 DateTimePicker^ from;DateTimePicker^ until;ComboBox^ side;
 PidPlotCanvas^ canvas;Label^ message;
 double x0,x1,gapSeconds;int sampleCount;

 static double Number(DataRow^ r,String^ col) {
  if(!r->Table->Columns->Contains(col)||r->IsNull(col))return Double::NaN;
  try {double v=Convert::ToDouble(r[col],System::Globalization::CultureInfo::InvariantCulture);
   return Double::IsNaN(v)||Double::IsInfinity(v)?Double::NaN:v;
  } catch(Exception^) {return Double::NaN;}
 }
 static bool Finite(double v) {return !Double::IsNaN(v)&&!Double::IsInfinity(v);}
 PidPlotSeries^ AddSeries(int panel,String^ name,Color color,bool step) {
  auto s=gcnew PidPlotSeries(panel,name,color,step);series->Add(s);return s;
 }
 void BuildTimes(DateTime sessionStart) {
  rowTimes=gcnew List<DateTime>();sourceRows=gcnew List<DataRow^>();
  DateTime date=(sessionStart.Year>=1753?sessionStart:DateTime::Now).Date;
  TimeSpan previous=TimeSpan::Zero;bool have=false;
  for each(DataRow^ row in source->Rows) {
   if(row->IsNull("RealTime"))continue;
   TimeSpan tod;
   if(!TimeSpan::TryParse(row["RealTime"]->ToString(),tod)||tod.TotalSeconds<0||tod.TotalDays>=1)continue;
   if(have&&tod.TotalSeconds<previous.TotalSeconds-43200)date=date.AddDays(1);
   DateTime at=date.Add(tod);
   if(rowTimes->Count&&at<=rowTimes[rowTimes->Count-1])continue;
   rowTimes->Add(at);sourceRows->Add(row);previous=tod;have=true;
  }
  auto gaps=gcnew List<double>();
  for(int j=1;j<rowTimes->Count;++j) {
   double delta=rowTimes[j].Subtract(rowTimes[j-1]).TotalSeconds;
   if(delta>0&&delta<=60)gaps->Add(delta);
  }
  gaps->Sort();double median=gaps->Count?gaps[gaps->Count/2]:5;
  gapSeconds=Math::Max(10.,2*median);
 }

 void Replot(Object^,EventArgs^) {
  if(rowTimes->Count<2) {message->Text=L"Недостаточно записей с корректным временем.";return;}
  if(from->Value>=until->Value) {message->Text=L"Начало периода должно быть раньше конца.";return;}
  x0=from->Value.ToOADate();x1=until->Value.ToOADate();
  series=gcnew List<PidPlotSeries^>();sampleCount=0;
  int channel=side->SelectedIndex==0?0:1;String^ suffix=gcnew String(channel==0?"Left":"Right");
  String^ sideRequestColumn=gcnew String(channel==0?"uLeft_TEN":"uRight_TEN");
  String^ rawColumn=String::Format("T{0}",channel);
  String^ filterColumn=String::Format("T_filt_{0}",channel);
  String^ activeColumn=String::Format("Act{0}",channel);
  auto raw=AddSeries(0,L"Сырая T",Color::DarkGray,false);
  auto filt=AddSeries(0,L"Фильтр 32",Color::Firebrick,false);
  auto target=AddSeries(0,L"Уставка (общая)",Color::Black,false);
  auto productLeft=AddSeries(0,L"T продукта (левый)",Color::ForestGreen,false);
  auto productRight=AddSeries(0,L"T продукта (правый)",Color::DarkViolet,false);
  auto ten1=AddSeries(1,L"ТЭН1: 0/1",Color::Goldenrod,true);
  auto ten2=AddSeries(1,L"ТЭН2: 2/3",Color::SaddleBrown,true);
  auto water=AddSeries(1,L"Впрыск: 4/5",Color::DodgerBlue,true);
  auto rh=AddSeries(2,L"RH возврата, %",Color::RoyalBlue,false);
  auto duty=AddSeries(2,L"Запрос воды, %",Color::SeaGreen,false);
  auto scale=AddSeries(3,L"heatScale",Color::MediumPurple,false);
  auto common=AddSeries(4,L"uCommon",Color::Black,false);
  auto chosen=AddSeries(4,L"Запрос стороны",Color::DarkOrange,false);
  auto integral=AddSeries(4,L"I",Color::RoyalBlue,false);
  auto derivative=AddSeries(4,L"D",Color::MediumPurple,false);
  auto ff=AddSeries(4,L"FF",Color::SeaGreen,false);
  int diagnosticCount=0;
  for(int j=0;j<rowTimes->Count;++j) {
   DateTime at=rowTimes[j];if(at<from->Value||at>until->Value)continue;
   ++sampleCount;DataRow^ row=sourceRows[j];
   double traw=Number(row,rawColumn);double tf=Number(row,filterColumn);
   if(Number(row,activeColumn)==0) {traw=Double::NaN;tf=Double::NaN;}
   raw->Add(at,traw);filt->Add(at,tf);
   target->Add(at,(Number(row,"T_filt_0")+Number(row,"T_filt_1"))/2+Number(row,"eT_common"));
   productLeft->Add(at,Number(row,"Act3")==0?Double::NaN:Number(row,"T3"));
   productRight->Add(at,Number(row,"Act4")==0?Double::NaN:Number(row,"T4"));
   double h1=Number(row,"Ten1_"+suffix),h2=Number(row,"Ten2_"+suffix),wi=Number(row,"_Inj");
   ten1->Add(at,Finite(h1)?(h1>0?1:0):Double::NaN);
   ten2->Add(at,Finite(h2)?(h2>0?3:2):Double::NaN);
   water->Add(at,Finite(wi)?(wi>0?5:4):Double::NaN);
   rh->Add(at,Number(row,"Act2")==0?Double::NaN:Number(row,"H2"));
   scale->Add(at,Number(row,"heatScale01"));common->Add(at,Number(row,"uCommon_TEN"));
   chosen->Add(at,Number(row,sideRequestColumn));
   double I=Number(row,"PID_I");if(Finite(I))++diagnosticCount;
   integral->Add(at,I);derivative->Add(at,Number(row,"PID_D"));
   ff->Add(at,Number(row,"PID_FF"));duty->Add(at,100*Number(row,"PID_waterDuty"));
  }
  // Fixed offsets match the axis marks; keep full ranges without altering log values.
  int band=0;
  for each(PidPlotSeries^ s in series)if(s->panel==4) {
   double minimum=0,maximum=2;
   for each(double value in s->y)if(Finite(value)) {
    minimum=Math::Min(minimum,value);maximum=Math::Max(maximum,value);
   }
   s->offset=3.0*band++;s->plotMin=s->offset+minimum;
   s->plotMax=s->offset+maximum;
  }
  message->Text=String::Format(L"Снимок текущего цикла: {0} строк; {1} строк с I/D/FF.\r\n"
   L"I/D/FF и запрос воды взяты из тех же строк основного лога. В старом формате эти поля отсутствуют и не восстанавливаются.",sampleCount,diagnosticCount);
  canvas->Invalidate();
 }

 void DrawGraph(Object^,PaintEventArgs^ paintArgs) {
  Graphics^ g=paintArgs->Graphics;g->Clear(Color::White);
  if(series==nullptr||sampleCount<2||x1<=x0) {g->DrawString(L"Выберите период с данными и нажмите «Построить».",this->Font,Brushes::Black,30.f,30.f);return;}
  g->SmoothingMode=System::Drawing::Drawing2D::SmoothingMode::AntiAlias;
  float left=78.f,width=static_cast<float>(Math::Max(300,canvas->ClientSize.Width-105));
  float section=(canvas->ClientSize.Height-65.f)/5.f;
  auto titles=gcnew cli::array<String^>{L"Температура воздуха и продукта, °C",L"Подтверждённые ТЭНы и команда впрыска",L"Влажность возврата и запрос воды, %",L"Разрешённая доля нагрева",L"Запрос нагрева и ПИД, экв. ТЭНа на сторону (со смещениями)"};
  auto font=gcnew System::Drawing::Font(L"Segoe UI",9.f);
  auto bold=gcnew System::Drawing::Font(L"Segoe UI",10.f,System::Drawing::FontStyle::Bold);
  try {
   g->DrawString(String::Format(L"{0} • {1:dd.MM.yyyy HH:mm:ss} — {2:dd.MM.yyyy HH:mm:ss}",side->Text,from->Value,until->Value),bold,Brushes::Black,left,2.f);
   for(int panel=0;panel<5;++panel) {
    float top=panel*section+77,H=section-78;
    double lo=panel==0?Double::PositiveInfinity:0,hi=panel==0?Double::NegativeInfinity:(panel==1?5.0:(panel==2?100:(panel==3?1.0:2.0)));
    if(panel==0)for each(PidPlotSeries^ s in series)if(s->panel==panel)
     for each(double value in s->y)if(Finite(value)){lo=Math::Min(lo,value);hi=Math::Max(hi,value);}
    if(panel==4)for each(PidPlotSeries^ s in series)if(s->panel==panel){lo=Math::Min(lo,s->plotMin);hi=Math::Max(hi,s->plotMax);}
    if(!Finite(lo)||!Finite(hi)) {lo=0;hi=1;}
    if(panel==0) {double margin=Math::Max(.1,(hi-lo)*.08);lo-=margin;hi+=margin;}
    if(hi-lo<.01)hi=lo+1;
    g->DrawString(titles[panel],bold,Brushes::Black,left,panel*section+32);
    float legendX=left;
    for each(PidPlotSeries^ s in series)if(s->panel==panel) {
     auto brush=gcnew SolidBrush(s->color);
     try {bool hasValue=false;for each(double value in s->y)if(Finite(value)){hasValue=true;break;}
      String^ name=s->name;
      if(panel==4)name+=L" (+"+s->offset.ToString("0.##",System::Globalization::CultureInfo::InvariantCulture)+L")";
      if(!hasValue)name+=L" (нет данных)";
      g->DrawString(name,font,brush,legendX,panel*section+53);
      legendX+=g->MeasureString(name,font).Width+18;
     }finally{delete brush;}
    }
    auto grid=gcnew Pen(Color::FromArgb(225,230,237));
    try {
     int divisions=panel==1?5:4;
     auto ticks=gcnew List<double>();
     if(panel==4) {
      for(int tick=0;tick<5;++tick)ticks->Add(3.0*tick);
     } else for(int tick=0;tick<=divisions;++tick)ticks->Add(lo+(hi-lo)*tick/divisions);
     for each(double value in ticks) {
      float y=top+H-static_cast<float>((value-lo)/(hi-lo))*H;
      String^ label=(panel==1||panel==4)?value.ToString("F0"):panel==3
       ? (value==0?L"0":value.ToString("0.0#",System::Globalization::CultureInfo::InvariantCulture))
       : value.ToString("F2");
      g->DrawLine(grid,left,y,left+width,y);g->DrawString(label,font,Brushes::DimGray,3.f,y-8);
     }
     for(int tick=0;tick<7;++tick) {
      float x=left+width*tick/6;g->DrawLine(grid,x,top,x,top+H);
      String^ label=DateTime::FromOADate(x0+(x1-x0)*tick/6).ToString("HH:mm:ss");
      g->DrawString(label,font,Brushes::DimGray,x-25,top+H+3);
     }
    }finally{delete grid;}
    auto plotState=g->Save();
    try {
    g->SetClip(System::Drawing::RectangleF(left,top,width,H));
    for each(PidPlotSeries^ s in series)if(s->panel==panel) {
     auto pen=gcnew Pen(s->color,s->name==L"Фильтр 32"?2.f:1.5f);
     bool previous=false;float px=0,py=0;double pt=0;
     try {
      for(int q=0;q<s->x->Count;++q) {
       double value=s->y[q],time=s->x[q];if(!Finite(value)){previous=false;continue;}
       float x=left+(float)((time-x0)/(x1-x0))*width;
       float y=top+H-(float)((value+s->offset-lo)/(hi-lo))*H;
       double allowed=gapSeconds;
       if(previous&&(time-pt)*86400<=allowed) {
        if(s->step){g->DrawLine(pen,px,py,x,py);g->DrawLine(pen,x,py,x,y);}
        else g->DrawLine(pen,px,py,x,y);
       }
       previous=true;px=x;py=y;pt=time;
      }
     }finally{delete pen;}
    }
    }finally{g->Restore(plotState);}
   }
  } finally {delete font;delete bold;}
 }

 void SaveImage(Object^,EventArgs^) {
  auto dialog=gcnew SaveFileDialog();dialog->Filter=L"PNG (*.png)|*.png";dialog->FileName=L"PID_regulation.png";
  if(dialog->ShowDialog(this)!=System::Windows::Forms::DialogResult::OK)return;
  auto image=gcnew Bitmap(canvas->Width,canvas->Height);
  try {canvas->DrawToBitmap(image,System::Drawing::Rectangle(0,0,image->Width,image->Height));image->Save(dialog->FileName,System::Drawing::Imaging::ImageFormat::Png);}
  finally {delete image;}
 }
public:
 PidGraphForm(DataTable^ data,DateTime sessionStart) {
  source=data;Text=L"График регулирования — снимок текущего цикла";
  Size=System::Drawing::Size(1400,1000);MinimumSize=System::Drawing::Size(1100,780);
  BuildTimes(sessionStart);
  auto toolbar=gcnew FlowLayoutPanel();toolbar->Dock=DockStyle::Top;toolbar->Height=42;
  auto fl=gcnew Label();fl->Text=L"С:";fl->AutoSize=true;toolbar->Controls->Add(fl);
  from=gcnew DateTimePicker();from->Format=DateTimePickerFormat::Custom;from->CustomFormat="dd.MM.yyyy HH:mm:ss";from->Width=190;
  toolbar->Controls->Add(from);auto tl=gcnew Label();tl->Text=L"По:";tl->AutoSize=true;toolbar->Controls->Add(tl);
  until=gcnew DateTimePicker();until->Format=DateTimePickerFormat::Custom;until->CustomFormat="dd.MM.yyyy HH:mm:ss";until->Width=190;
  toolbar->Controls->Add(until);side=gcnew ComboBox();side->DropDownStyle=ComboBoxStyle::DropDownList;
  side->Items->Add(L"Левая сторона");side->Items->Add(L"Правая сторона");side->SelectedIndex=1;side->Width=155;toolbar->Controls->Add(side);
  auto build=gcnew Button();build->Text=L"Построить";build->Width=110;build->Click+=gcnew EventHandler(this,&PidGraphForm::Replot);toolbar->Controls->Add(build);
  auto save=gcnew Button();save->Text=L"Сохранить PNG";save->Width=140;save->Click+=gcnew EventHandler(this,&PidGraphForm::SaveImage);toolbar->Controls->Add(save);
  message=gcnew Label();message->Dock=DockStyle::Bottom;message->Height=55;
  canvas=gcnew PidPlotCanvas();canvas->Dock=DockStyle::Fill;canvas->Paint+=gcnew PaintEventHandler(this,&PidGraphForm::DrawGraph);
  Controls->Add(canvas);Controls->Add(toolbar);Controls->Add(message);
  if(rowTimes->Count>=2) {
   until->Value=rowTimes[rowTimes->Count-1];DateTime first=rowTimes[0];
   from->Value=until->Value.Subtract(first).TotalMinutes>30?until->Value.AddMinutes(-30):first;
  }
  Replot(nullptr,nullptr);
 }
};
}
