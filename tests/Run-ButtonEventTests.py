"""Проверяет очередь прошивки, реальные обработчики кнопок и устойчивую запись сервера.
Запуск из среды vcvars64 с Python; оборудование и рабочий log.txt не используются.
"""
from pathlib import Path
import os
import subprocess
import tempfile
import sys

ROOT = Path(__file__).resolve().parents[1]
FIRMWARE = Path(os.environ.get("DEFROST_FIRMWARE", r"D:\ST\Defrost"))


def run():
    """Компилирует временные стенды и запускает проверки; ошибка прерывает выполнение."""
    with tempfile.TemporaryDirectory(prefix="defrost_buttons_") as folder:
        folder = Path(folder)
        control = (FIRMWARE / "STM32CubeIDE/ProjectCode/ModBus.cpp").read_text(encoding="utf-8")
        start = control.index("if (DefrostControl_UsesNewWrkAlrAlgorithm() != 0u)", control.index("const uint8_t butStopNow"))
        end = control.index("g_prevButStart = butStartNow;", start) + len("g_prevButStart = butStartNow;")
        handler = control[start:end]
        queue_test = r'''
#include <cassert>
#include <cstdio>
#include "ButtonEvents.hpp"
ButtonEventQueue g_buttonEvents;
namespace Model { uint8_t Gate_Alarm_Hardware=0; }
static uint8_t g_prevButStart=0;
static bool enabled=false;
static unsigned int TimeFromStart=100;
int DefrostControl_UsesNewWrkAlrAlgorithm(){return 1;}
void DetectGateAlarmAfterIoRead(){}
void HandleGateAlarmAfterIoRead(){enabled=false;}
int DefrostControl_IsEnabled(){return enabled;}
void DefrostControl_SetEnabled(int value){enabled=value!=0;}
void osKernelLock(){}
void osKernelUnlock(){}
void poll(uint8_t butStartNow,uint8_t butStopNow) {
''' + handler + r'''
}
uint32_t word(const uint8_t* p){uint32_t v;memcpy(&v,p,4);return v;}
int main(){
 uint8_t session[16]={1},other[16]={2},out[113]={},again[113]={};
 // Короткие, уже обработанные нажатия накапливаются до первого опроса сервера.
 poll(0,1); poll(1,1); poll(1,1); poll(0,1); poll(0,0); poll(0,0);
 assert(!enabled);assert(g_buttonEvents.Read(session,out)==39 && out[20]==2);
 assert(word(out+21)==1 && word(out+25)==100 && out[29]==1 && out[38]==2);
 // Потерянный ответ не удаляет очередь; удержание кнопки не создаёт события.
 assert(g_buttonEvents.Read(other,again)==39 && memcmp(out,again,39)==0);
 assert(!g_buttonEvents.Ack(other,2));assert(!g_buttonEvents.Ack(session,3));
 assert(g_buttonEvents.Ack(session,1));
 g_buttonEvents.Read(session,out);assert(out[20]==1 && word(out+21)==2);
 assert(g_buttonEvents.Ack(session,1));assert(g_buttonEvents.Ack(session,2));
 g_buttonEvents.Read(session,out);assert(out[20]==0);
 // Защита ворот имеет приоритет: необработанный СТАРТ не выдаётся за выполненный.
 Model::Gate_Alarm_Hardware=1;poll(1,1);g_buttonEvents.Read(session,out);assert(out[20]==0);
 Model::Gate_Alarm_Hardware=0;poll(0,1);poll(1,1);
 // Между чтением и ACK может прийти новая кнопка: она не удаляется старым ACK.
 g_buttonEvents.Read(session,out);auto offered=word(out+21);poll(0,0);
 assert(g_buttonEvents.Ack(session,offered));g_buttonEvents.Read(session,out);assert(out[20]==1 && out[29]==2);
 // Новая загрузка закрепляет новый кандидат, даже если TCP-соединение прежнее.
 ButtonEventQueue reboot;reboot.Push(1,1);reboot.Read(other,again);
 assert(memcmp(out,again,16)!=0);assert(!reboot.Ack(session,1));
 // Переполнение не вытесняет первые события; счётчик передаётся явно.
 ButtonEventQueue full;
 for(int i=0;i<515;++i)full.Push(i%2+1,i);
 full.Read(session,out);assert(out[20]==10 && word(out+16)==3 && word(out+21)==1);
 for(int batch=0;batch<52;++batch){
  auto size=full.Read(session,out);assert(size<=113);
  if(out[20])assert(full.Ack(session,word(out+21+9*(out[20]-1))));
 }
 full.Read(session,out);assert(out[20]==0);
 puts("PASS: handlers, held buttons, retries, partial/stale ACK, overflow and batch limits");
}
'''
        cpp = folder / "queue.cpp"
        cpp.write_text(queue_test,encoding="utf-8")
        subprocess.run(["cl", "/nologo", "/EHsc", "/std:c++14", "/utf-8", str(cpp),
                        "/I"+str(FIRMWARE/"STM32CubeIDE/ProjectCode"), "/Fe:"+str(folder/"queue.exe")],cwd=folder,check=True)
        subprocess.run([str(folder/"queue.exe")],check=True)
        # Используем именно рабочий класс журнала, исключая зависимости окон сервера.
        logger = (ROOT/"SServer.h").read_text(encoding="utf-8-sig")
        logger = logger[logger.index("ref class GlobalLogger"):logger.index("std::string ConvertToStdString")]
        logger_test = r'''
#using <System.dll>
#using <System.Core.dll>
#using <System.Windows.Forms.dll>
#include <string>
#include <climits>
using namespace System;
using namespace System::IO;
using namespace System::Threading;
using namespace System::Windows::Forms;
''' + logger + r'''
int main(array<String^>^ args){
 String^ key="[ButtonEvent:session:1]";
 if(!GlobalLogger::TryLogButtonEvent(key,L"Физическая кнопка СТАРТ")) return 1;
 if(!GlobalLogger::TryLogButtonEvent(key,L"Физическая кнопка СТАРТ")) return 2;
 if(args[0]=="retry" && !GlobalLogger::TryLogButtonEvent("[ButtonEvent:session:2]",L"Физическая кнопка СТОП")) return 3;
 GlobalLogger::Shutdown();
 auto lines=File::ReadAllLines(Path::Combine(Path::GetDirectoryName(Application::ExecutablePath),"log.txt"));
 if(lines->Length!=(args[0]=="retry"?2:1)) return 4;
 // Ошибка записи запрещает ACK: закрываем поток через отражение только в стенде.
 GlobalLogger::Initialize();
 auto field=GlobalLogger::typeid->GetField("writer",System::Reflection::BindingFlags::NonPublic|System::Reflection::BindingFlags::Static);
 safe_cast<StreamWriter^>(field->GetValue(nullptr))->BaseStream->Close();
 if(GlobalLogger::TryLogButtonEvent("[ButtonEvent:session:failure]","fail")) return 5;
 Console::WriteLine("PASS: durable log, duplicate response, process restart and disk failure");
 return 0;
}
'''
        cpp=folder/"logger.cpp";cpp.write_text(logger_test,encoding="utf-8")
        subprocess.run(["cl","/nologo","/clr","/EHa","/utf-8",str(cpp),"/Fe:"+str(folder/"logger.exe")],cwd=folder,check=True)
        subprocess.run([str(folder/"logger.exe"),"first"],check=True)
        subprocess.run([str(folder/"logger.exe"),"retry"],check=True)
        # Рабочий обработчик ответа сервера: CRC, тип кнопки, повторы и запрет ACK при ошибке файла.
        receiver = (ROOT/"CommandsDefroster.cpp").read_text(encoding="utf-8-sig")
        receiver = receiver[receiver.index("void ProjectServerW::DataForm::HandleButtonEventsResponse"):]
        prefix = r'''
#using <System.dll>
#using <System.Core.dll>
#using <System.Windows.Forms.dll>
#include <string>
#include <climits>
#include <vcclr.h>
#include "Commands.h"
using namespace System; using namespace System::IO; using namespace System::Threading;
using namespace System::Windows::Forms;
const int INVALID_SOCKET=-1;
''' + logger + r'''
namespace ProjectServerW {
ref class DataForm {
public:
 bool isFormClosingNow=false, IsDisposed=false, Disposing=false;
 bool buttonEventsPending=false,buttonEventsUnsupported=false;
 int clientSocket=1,acks=0; uint32_t acknowledged=0;
 String^ ClientIP="test"; Object^ commandPipelineGate=gcnew Object();
 bool SendCommand(const Command& c,String^) {++acks;uint32_t value;memcpy(&value,c.data+16,4);acknowledged=value;return true;}
 void HandleButtonEventsResponse(array<Byte>^);
};
}
''' + receiver + r'''
array<Byte>^ packet(int button=1) {
 auto p=gcnew array<Byte>(36);
 p[0]=CmdType::REQUEST;p[1]=CmdRequest::GET_BUTTON_EVENTS;p[3]=30;
 p[4]=127;p[24]=1;p[25]=1;p[29]=100;p[33]=button;
 pin_ptr<Byte> ptr=&p[0];uint16_t crc=CalculateCommandCRC(ptr,34);
 p[34]=crc&255;p[35]=crc>>8;return p;
}
int main(){
 auto r=gcnew ProjectServerW::DataForm();
 r->HandleButtonEventsResponse(packet());if(r->acks!=1 || r->acknowledged!=1)return 1;
 r->HandleButtonEventsResponse(packet());if(r->acks!=2)return 2;
 r->HandleButtonEventsResponse(packet(9));if(r->acks!=2)return 3;
 auto bad=packet();bad[29]^=1;r->HandleButtonEventsResponse(bad);if(r->acks!=2)return 4;
 auto field=GlobalLogger::typeid->GetField("writer",System::Reflection::BindingFlags::NonPublic|System::Reflection::BindingFlags::Static);
 safe_cast<StreamWriter^>(field->GetValue(nullptr))->BaseStream->Close();
 auto next=packet();next[25]=2;pin_ptr<Byte> ptr=&next[0];uint16_t crc=CalculateCommandCRC(ptr,34);next[34]=crc&255;next[35]=crc>>8;
 r->HandleButtonEventsResponse(next);if(r->acks!=2)return 5;
 Console::WriteLine("PASS: actual response handler, CRC, duplicate ACK, invalid event and failed-log ACK veto");
}
'''
        cpp=folder/"receiver.cpp";cpp.write_text(prefix,encoding="utf-8")
        subprocess.run(["cl","/nologo","/clr","/EHa","/utf-8","/I"+str(ROOT),str(cpp),str(ROOT/"Commands.cpp"),"/Fe:"+str(folder/"receiver.exe")],cwd=folder,check=True)
        subprocess.run([str(folder/"receiver.exe")],check=True)



if __name__ == "__main__":
    run()
