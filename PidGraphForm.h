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
 int panel;String^ name;Color color;bool step;bool rightAxis;float lineWidth;bool dashed;
 List<double>^ x;List<double>^ y;
 PidPlotSeries(int p,String^ n,Color c,bool s) {
  panel=p;name=n;color=c;step=s;rightAxis=false;lineWidth=1.5f;dashed=false;x=gcnew List<double>();y=gcnew List<double>();
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
 DateTimePicker^ from;DateTimePicker^ until;
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
  auto rawLeft=AddSeries(0,L"Сырая T слева",Color::DarkGray,false);
  auto rawRight=AddSeries(0,L"Сырая T справа",Color::DarkGray,false);
  rawRight->dashed=true;
  auto filterLeft=AddSeries(0,L"Фильтр 32 слева",Color::FromArgb(138,174,152),false);
  auto filterRight=AddSeries(0,L"Фильтр 32 справа",Color::FromArgb(193,151,139),false);
  auto average=AddSeries(0,L"Средняя T регулятора",Color::RoyalBlue,false);
  average->lineWidth=2.8f;
  auto target=AddSeries(0,L"Уставка",Color::Black,false);
  target->lineWidth=2.8f;
  auto productLeft=AddSeries(1,L"T продукта (левый, сырая)",Color::ForestGreen,false);
  auto productRight=AddSeries(1,L"T продукта (правый, сырая)",Color::DarkViolet,false);
  auto heaterNames=gcnew cli::array<String^>{"Ten1_Left","Ten2_Left","Ten1_Right","Ten2_Right"};
  auto heaters=gcnew cli::array<PidPlotSeries^>{
   AddSeries(2,L"Слева ТЭН1: 0/1",Color::DarkGreen,true),
   AddSeries(2,L"Слева ТЭН2: 1/2",Color::MediumSeaGreen,true),
   AddSeries(2,L"Справа ТЭН1: 2/3",Color::SaddleBrown,true),
   AddSeries(2,L"Справа ТЭН2: 3/4",Color::Firebrick,true)
  };
  // При совпадении уровней более толстая нижняя кривая видна по краям верхней.
  for(int k=0;k<heaters->Length;++k)heaters[k]->lineWidth=4.8f-k;
  auto water=AddSeries(3,L"Впрыск: 0/1 (правая шкала)",Color::DodgerBlue,true);
  water->rightAxis=true;
  auto rh=AddSeries(3,L"RH возврата, %",Color::RoyalBlue,false);
  auto duty=AddSeries(3,L"Запрос воды, %",Color::SeaGreen,false);
  auto scale=AddSeries(4,L"heatScale",Color::MediumPurple,false);
  auto common=AddSeries(5,L"uCommon",Color::Black,false);
  auto requestLeft=AddSeries(6,L"Слева",Color::SeaGreen,false);
  auto requestRight=AddSeries(6,L"Справа",Color::Firebrick,false);
  auto proportional=AddSeries(5,L"P",Color::DarkCyan,false);
  auto integral=AddSeries(5,L"I",Color::RoyalBlue,false);
  auto derivative=AddSeries(5,L"D",Color::MediumPurple,false);
  auto ff=AddSeries(5,L"FF",Color::SeaGreen,false);
  int diagnosticCount=0;
  for(int j=0;j<rowTimes->Count;++j) {
   DateTime at=rowTimes[j];if(at<from->Value||at>until->Value)continue;
   ++sampleCount;DataRow^ row=sourceRows[j];
   const bool leftValid=Number(row,"Act0")!=0;
   const bool rightValid=Number(row,"Act1")!=0;
   const double left=Number(row,"T_filt_0"),right=Number(row,"T_filt_1");
   rawLeft->Add(at,leftValid?Number(row,"T0"):Double::NaN);
   rawRight->Add(at,rightValid?Number(row,"T1"):Double::NaN);
   filterLeft->Add(at,leftValid?left:Double::NaN);
   filterRight->Add(at,rightValid?right:Double::NaN);
   // Новые записи содержат точный вход регулятора. Для старых сохраняем
   // прежнюю оценку уставки при наличии обеих температур подачи.
   const double supplyAverage=Number(row,"T_sup_avg_C");
   average->Add(at,supplyAverage);
   const double setpointBase=Finite(supplyAverage)?supplyAverage:
    (leftValid&&rightValid?(left+right)/2.0:Double::NaN);
   target->Add(at,setpointBase+Number(row,"eT_common"));
   productLeft->Add(at,Number(row,"Act3")==0?Double::NaN:Number(row,"T3"));
   productRight->Add(at,Number(row,"Act4")==0?Double::NaN:Number(row,"T4"));
   const double ioType=Number(row,"Typ6");
   const bool ioValid=Number(row,"Act6")!=0 && (!Finite(ioType)||ioType==4);
   for(int k=0;k<heaters->Length;++k) {
    double feedback=Number(row,heaterNames[k]);
    heaters[k]->Add(at,ioValid&&Finite(feedback)?k+(feedback>0?1:0):Double::NaN);
   }
   double wi=Number(row,"_Inj");
   water->Add(at,Finite(wi)?(wi>0?1:0):Double::NaN);
   rh->Add(at,Number(row,"Act2")==0?Double::NaN:Number(row,"H2"));
   scale->Add(at,Number(row,"heatScale01"));common->Add(at,Number(row,"uCommon_TEN"));
   requestLeft->Add(at,Number(row,"uLeft_TEN"));
   requestRight->Add(at,Number(row,"uRight_TEN"));
   double I=Number(row,"PID_I");if(Finite(I))++diagnosticCount;
   proportional->Add(at,Number(row,"PID_P"));
   integral->Add(at,I);derivative->Add(at,Number(row,"PID_D"));
   ff->Add(at,Number(row,"PID_FF"));duty->Add(at,100*Number(row,"PID_waterDuty"));
  }
  message->Text=String::Format(L"Снимок текущего цикла: {0} строк; {1} строк с I/D/FF.\r\n"
   L"I/D/FF и запрос воды взяты из тех же строк основного лога. В старом формате эти поля отсутствуют и не восстанавливаются.",sampleCount,diagnosticCount);
  canvas->Invalidate();
 }

 // Округляем шаг автоматически масштабируемых графиков к ряду 1 / 2 / 2,5 / 5.
 static double RoundedGridStep(double minimumStep) {
  const double magnitude=Math::Pow(10.0,Math::Floor(Math::Log10(minimumStep)));
  const double normalized=minimumStep/magnitude;
  if(normalized<=1.0)return magnitude;
  if(normalized<=2.0)return 2.0*magnitude;
  if(normalized<=2.5)return 2.5*magnitude;
  if(normalized<=5.0)return 5.0*magnitude;
  return 10.0*magnitude;
 }

 void DrawGraph(Object^,PaintEventArgs^ paintArgs) {
  Graphics^ g=paintArgs->Graphics;g->Clear(Color::White);
  if(series==nullptr||sampleCount<2||x1<=x0) {g->DrawString(L"Выберите период с данными и нажмите «Построить».",this->Font,Brushes::Black,30.f,30.f);return;}
  g->SmoothingMode=System::Drawing::Drawing2D::SmoothingMode::AntiAlias;
  float left=78.f,width=static_cast<float>(Math::Max(300,canvas->ClientSize.Width-155));
  auto titles=gcnew cli::array<String^>{L"Температура воздуха дефростера, °C",L"Температура продукта, °C",L"Подтверждённые ТЭНы",L"Влажность возврата, запрос воды и впрыск",L"Разрешённая доля нагрева",L"Общий запрос нагрева и составляющие ПИД, экв. ТЭНа",L"Запрос нагрева по сторонам, экв. ТЭНа"};
  float section=(canvas->ClientSize.Height-65.f)/titles->Length;
  auto font=gcnew System::Drawing::Font(L"Segoe UI",9.f);
  auto bold=gcnew System::Drawing::Font(L"Segoe UI",10.f,System::Drawing::FontStyle::Bold);
  try {
   g->DrawString(String::Format(L"{0:dd.MM.yyyy HH:mm:ss} — {1:dd.MM.yyyy HH:mm:ss}",from->Value,until->Value),bold,Brushes::Black,left,2.f);
   for(int panel=0;panel<titles->Length;++panel) {
    float top=panel*section+77,H=section-78;
    // Температуры воздуха и продукта имеют отдельные автоматические шкалы.
    bool temperaturePanel=panel<=1;
    double lo=temperaturePanel?Double::PositiveInfinity:0,hi=temperaturePanel?Double::NegativeInfinity:(panel==2?4.0:(panel==3?100:(panel==4?1.0:2.0)));
    if(temperaturePanel)for each(PidPlotSeries^ s in series)if(s->panel==panel)
     for each(double value in s->y)if(Finite(value)){lo=Math::Min(lo,value);hi=Math::Max(hi,value);}
    // Общая шкала включает штатный диапазон uCommon 0…2 и все значения P/I/D/FF.
    if(panel==5) {
     for each(PidPlotSeries^ s in series)if(s->panel==panel)
      for each(double value in s->y)if(Finite(value)){lo=Math::Min(lo,value);hi=Math::Max(hi,value);}
    }
    if(!Finite(lo)||!Finite(hi)) {lo=0;hi=1;}
    // Шкалы выходов, влажности, heatScale и запроса нагрева имеют штатные деления.
    double gridStep=(hi-lo)/4.0;
    if(temperaturePanel) {
     const double preferredStep=(hi-lo)/5.0;
     gridStep=preferredStep<=.1?.1:preferredStep<=.5?.5:Math::Ceiling(preferredStep);
    } else if(panel==5) {
     gridStep=RoundedGridStep(gridStep);
    }
    // Индексы отсчитываются от нуля, поэтому отрицательные температуры и
    // поправки ПИД попадают на ту же сетку. Округление устраняет погрешность
    // деления на дробный шаг, когда граница уже лежит на линии сетки.
    double firstTick=Math::Floor(Math::Round(lo/gridStep,10));
    double lastTick=Math::Ceiling(Math::Round(hi/gridStep,10));
    if(lastTick==firstTick) {--firstTick;++lastTick;}
    const double scaleLo=firstTick*gridStep,scaleHi=lastTick*gridStep;
    // Дополнительные поля оставляют крайние кривые внутри отсечения.
    const double edgeMargin=(scaleHi-scaleLo)*.04;
    lo=scaleLo-edgeMargin;hi=scaleHi+edgeMargin;
    g->DrawString(titles[panel],bold,Brushes::Black,left,panel*section+32);
    float legendX=left;
    for each(PidPlotSeries^ s in series)if(s->panel==panel) {
     auto brush=gcnew SolidBrush(s->color);
     try {bool hasValue=false;for each(double value in s->y)if(Finite(value)){hasValue=true;break;}
      String^ name=s->name;
      if(!hasValue)name+=L" (нет данных)";
      g->DrawString(name,font,brush,legendX,panel*section+53);
      legendX+=g->MeasureString(name,font).Width+18;
     }finally{delete brush;}
    }
    auto grid=gcnew Pen(Color::FromArgb(225,230,237));
    try {
     for(double tick=firstTick;tick<=lastTick;++tick) {
      const double value=tick*gridStep;
      float y=top+H-static_cast<float>((value-lo)/(hi-lo))*H;
      String^ label=temperaturePanel?value.ToString("0.#"):panel==2?value.ToString("F0"):panel==4
       ? (value==0?L"0":value.ToString("0.0#",System::Globalization::CultureInfo::InvariantCulture))
       : value.ToString("F2");
      g->DrawLine(grid,left,y,left+width,y);g->DrawString(label,font,Brushes::DimGray,3.f,y-8);
     }
     // Правая ось показывает команду форсунки целыми 0/1. Её границы
     // совпадают с 0/100% левой оси и сохраняют защитные поля от обрезания.
     if(panel==3) {
      auto axisPen=gcnew Pen(Color::DodgerBlue);
      auto axisBrush=gcnew SolidBrush(Color::DodgerBlue);
      try {
       float axisX=left+width;
       g->DrawLine(axisPen,axisX,top,axisX,top+H);
       for(int state=0;state<=1;++state) {
        float axisY=top+H-static_cast<float>((100.0*state-lo)/(hi-lo))*H;
        g->DrawLine(axisPen,axisX,axisY,axisX+5,axisY);
        g->DrawString(state.ToString(),font,axisBrush,axisX+8,axisY-8);
       }
      }finally{delete axisPen;delete axisBrush;}
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
     auto pen=gcnew Pen(s->color,s->lineWidth);
     if(s->dashed)pen->DashStyle=System::Drawing::Drawing2D::DashStyle::Dash;
     bool previous=false;float px=0,py=0;double pt=0;
     try {
      for(int q=0;q<s->x->Count;++q) {
       double value=s->y[q],time=s->x[q];if(!Finite(value)){previous=false;continue;}
       float x=left+(float)((time-x0)/(x1-x0))*width;
       // Впрыск хранится как 0/1; только координата переводится в диапазон
       // левой оси 0..100%, исходные значения лога не изменяются.
       double plotValue=s->rightAxis?100.0*value:value;
       float y=top+H-(float)((plotValue-lo)/(hi-lo))*H;
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
  toolbar->Controls->Add(until);
  auto build=gcnew Button();build->Text=L"Построить";build->Width=110;build->Click+=gcnew EventHandler(this,&PidGraphForm::Replot);toolbar->Controls->Add(build);
  auto save=gcnew Button();save->Text=L"Сохранить PNG";save->Width=140;save->Click+=gcnew EventHandler(this,&PidGraphForm::SaveImage);toolbar->Controls->Add(save);
  message=gcnew Label();message->Dock=DockStyle::Bottom;message->Height=55;
  // Прокрутка сохраняет высоту каждого из семи графиков при небольшом окне.
  auto graphViewport=gcnew Panel();graphViewport->Dock=DockStyle::Fill;graphViewport->AutoScroll=true;
  canvas=gcnew PidPlotCanvas();canvas->Dock=DockStyle::Top;canvas->Height=1465;canvas->Paint+=gcnew PaintEventHandler(this,&PidGraphForm::DrawGraph);
  graphViewport->Controls->Add(canvas);
  Controls->Add(graphViewport);Controls->Add(toolbar);Controls->Add(message);
  if(rowTimes->Count>=2) {
   until->Value=rowTimes[rowTimes->Count-1];DateTime first=rowTimes[0];
   from->Value=until->Value.Subtract(first).TotalMinutes>30?until->Value.AddMinutes(-30):first;
  }
  Replot(nullptr,nullptr);
 }
};
}
