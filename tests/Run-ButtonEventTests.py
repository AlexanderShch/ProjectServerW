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
 for(int i=0;i<19;++i)full.Push(i%2+1,i);
 full.Read(session,out);assert(out[20]==10 && word(out+16)==3 && word(out+21)==1);
 for(int batch=0;batch<52;++batch){
  auto size=full.Read(session,out);assert(size<=113);
  if(out[20])assert(full.Ack(session,word(out+21+9*(out[20]-1))));
 }
 full.Read(session,out);assert(out[20]==0);
 ButtonEventQueue mixed;
 SensorEventDetail detail{};detail.received=9;
 for(int i=0;i<13;++i)mixed.PushDiagnostic(3,i,detail);
 for(int i=0;i<5;++i)mixed.Push(i%2+1,50+i);
 assert(mixed.Pending());
 for(unsigned i=1;i<=16;++i){
  auto len=mixed.ReadEvent(session,out);
  assert(out[24]==17-i && word(out+16)==1 && word(out+20)==1);
  assert(len==(i<=12?61:34) && word(out+25)==i);
  assert(mixed.ReadEvent(other,again)==len && memcmp(out,again,len)==0);
  assert(!mixed.AckEvent(other,i) && !mixed.AckEvent(session,i+1));
  assert(mixed.AckEvent(session,i) && mixed.AckEvent(session,i));
 }
 assert(!mixed.Pending() && mixed.ReadEvent(session,out)==25);
 puts("PASS: handlers, held buttons, retries, partial/stale ACK, overflow and batch limits");
}
'''
        # Проверяем рабочий валидатор UART5 на эталонном ответе Modbus из журнала.
        check = control[control.index("enum { MB_DETAIL_ADDRESS"):control.index("struct SensorFailureRun")]
        crc_start = control.index("uint16_t MB_GetCRC(volatile uint8_t* buf, uint16_t len)\n{")
        crc = control[crc_start:control.index("\n}", crc_start)+2]
        table = control.index("const uint16_t crc16_table[] =")
        crc = control[table:control.index("};", table)+2] + "\n" + crc
        status = (FIRMWARE/"STM32CubeIDE/ProjectCode/ModBus.hpp").read_text(encoding="utf-8")
        categories = status[status.index("MB_DIAG_NO_ANSWER = 0"):status.index("} MB_DiagCat_t")]
        validation = r'''
#include <cstdint>
#include <cstring>
#include <cassert>
#include <cstdio>
using MB_Command_t=int;
enum { MB_CMD_READ_REGS=3,MB_CMD_READ_COILS=1,MB_CMD_READ_INPUT=2,MB_CMD_WRITE_REG=6 };
int huart4,huart5,s_mbRxUartError=0;
namespace Model {int Type_of_sensor=2;}
struct MB_Active_t {int* UART=&huart5;uint8_t Tx_Buffer[128]{},Rx_Buffer[128]{};};
enum {''' + categories + "};\n" + crc + "\n" + check + r'''
int main(){
 MB_Active_t mb;auto f=mb.Rx_Buffer;mb.Tx_Buffer[0]=0x67;
 uint8_t sample[]={0x67,3,4,2,0xa2,0,0xea,0xdd,0xe0};memcpy(f,sample,9);
 auto test=[&](unsigned n){return MB_CheckSensorReply(&mb,3,MB_ExpectedLength(3,2),n);};
 assert(test(9)==MB_DETAIL_OK && test(0)==MB_DIAG_NO_ANSWER);
 assert(test(8)==MB_DIAG_SHORT && test(10)==MB_DETAIL_EXTRA && test(129)==MB_DETAIL_EXTRA);
 f[7]=0x1c;assert(test(9)==MB_DIAG_CRC);
 f[2]=255;assert(test(9)==MB_DETAIL_BYTE_COUNT);
 f[0]=0x66;assert(test(9)==MB_DETAIL_ADDRESS);
 f[0]=0x67;f[1]=4;assert(test(9)==MB_DETAIL_FUNCTION);
 f[1]=0x83;f[2]=2;auto sum=MB_GetCRC(f,3);f[3]=sum&255;f[4]=sum>>8;
 assert(test(5)==MB_DETAIL_EXCEPTION);f[3]^=1;assert(test(5)==MB_DIAG_CRC);
 s_mbRxUartError=1;assert(test(5)==MB_DIAG_UART_ERR);
 assert(MB_ExpectedLength(3,1)==7 && MB_ExpectedLength(2,16)==7 && MB_ExpectedLength(6,23)==8);
 puts("PASS: sensor reply length, CRC, address, function, byte count, exception and UART");
}
'''
        cpp = folder/"reply.cpp"
        cpp.write_text(validation,encoding="utf-8")
        subprocess.run(["cl","/nologo","/EHsc","/std:c++14","/utf-8",str(cpp),"/Fe:"+str(folder/"reply.exe")],cwd=folder,check=True)
        subprocess.run([str(folder/"reply.exe")],check=True)
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
        # Actual firmware ACK handler: a telemetry request owns exactly one log snapshot.
        firmware_receiver = (FIRMWARE/"STM32CubeIDE/ProjectCode/CommandReceiver.cpp").read_text(encoding="utf-8-sig")
        start = firmware_receiver.index("CommandStatus_t CommandReceiver_HandleTelemetry(")
        end = firmware_receiver.index("    return status;", start) + len("    return status;\n}")
        ack_test = r'''
#include <stdint.h>
#include <cstring>
#include <cassert>
#include <cstdio>
using CommandStatus_t=int;
constexpr int CMD_STATUS_OK=0,CMD_STATUS_INVALID_CODE=1,CMD_TYPE_TELEMETRY=0,TELEMETRY_DATA_OK=1,TELEMETRY_DATA_FALSE=2;
struct Command_t {uint8_t commandCode;};
struct CommandResponse_t {uint8_t commandType,commandCode,status,dataLength,data[122];};
struct ControlLogPayload_t {uint32_t marker;};
static bool s_requestTagged=true,s_stateTagged=true;
static uint64_t s_requestId=9,s_stateRequestId=9;
static bool s_stateAwaitingAck=false,s_stateNeedsLog=false;
static uint32_t s_stateRequestedAt=0,ticks=0;
static ControlLogPayload_t s_stateLog={42};
static unsigned logs=0,retries=0;
static int ServerResponseReceived_SemHandle=1;
uint32_t osKernelGetTickCount(){return ticks;}
uint32_t osKernelGetTickFreq(){return 1000;}
void osSemaphoreRelease(int){}
void ResendLastTelemetry(){++retries;}
void CommandReceiver_SendResponse(CommandResponse_t* r){uint32_t v;memcpy(&v,r->data,4);assert(v==42);++logs;}
''' + firmware_receiver[start:end] + r'''
int main(){
 Command_t ack{1},nack{2};
 CommandReceiver_HandleTelemetry(&ack);assert(logs==0);
 s_stateAwaitingAck=true;s_stateNeedsLog=true;
 s_requestId=8;CommandReceiver_HandleTelemetry(&ack);assert(s_stateAwaitingAck && logs==0);
 s_requestId=9;CommandReceiver_HandleTelemetry(&nack);assert(retries==1 && s_stateAwaitingAck);
 CommandReceiver_HandleTelemetry(&ack);assert(logs==0 && !s_stateAwaitingAck);
 CommandReceiver_HandleTelemetry(&ack);assert(logs==0);
 s_stateAwaitingAck=true;s_stateNeedsLog=false;
 CommandReceiver_HandleTelemetry(&ack);assert(logs==0);
 s_stateAwaitingAck=true;s_stateNeedsLog=true;ticks=1500;
 CommandReceiver_HandleTelemetry(&ack);assert(logs==0 && !s_stateAwaitingAck);
 puts("PASS: firmware telemetry ACK emits no log; stale ID, NACK, duplicate and timeout");
}
'''
        cpp=folder/"ack.cpp";cpp.write_text(ack_test,encoding="utf-8")
        subprocess.run(["cl","/nologo","/EHsc","/utf-8",str(cpp),"/Fe:"+str(folder/"ack.exe")],cwd=folder,check=True)
        subprocess.run([str(folder/"ack.exe")],check=True)
        # Рабочий обработчик ответа сервера: CRC, тип кнопки, повторы и запрет ACK при ошибке файла.
        receiver = (ROOT/"CommandsDefroster.cpp").read_text(encoding="utf-8-sig")
        receiver = receiver[receiver.index("void ProjectServerW::DataForm::ResetButtonEventPolling"): receiver.index("// Ожидание завершается recv-потоком")]
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
 bool unifiedEvents=false,eventsAvailable=false,moreEventsAfterAck=false;
 bool buttonEventsPending=false,buttonEventsUnsupported=false,controllerRestartVersionPending=false;
 System::Windows::Forms::Timer^ sendStateTimer=nullptr;
 DateTime telemetryRequestTime=DateTime::MinValue;
 bool WaitControllerExchange(int) {return true;} // Transport is covered by LogSessionIsolationTests.
 int versionRequests=0;
 bool SendVersionRequest() {++versionRequests;return true;}
 int clientSocket=1,buttonEventsSocket=1,acks=0,polls=0; uint32_t acknowledged=0;
 bool startupSequenceCompleted=true,sensorDiagInProgress=false,pendingRowWrkBit=false,hasTelemetry=true;
 DateTime lastTelemetryTime=DateTime::MinValue,buttonEventsRequestTime=DateTime::MinValue;
 String^ ClientIP="test"; Object^ commandPipelineGate=gcnew Object();
 bool SendCommand(const Command& c,String^) {if(c.commandCode!=CmdRequest::ACK_BUTTON_EVENTS && c.commandCode!=CmdRequest::ACK_EVENT){++polls;return true;}++acks;uint32_t value;memcpy(&value,c.data+16,4);acknowledged=value;return true;}
 void ResetButtonEventPolling();
 void PollButtonEvents(Object^,EventArgs^);
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
 // Старая прошивка не опрашивается повторно по времени. После рестарта
 // на том же сокете запрашивается версия, затем снова читается очередь.
 r->buttonEventsUnsupported=true;r->buttonEventsRequestTime=DateTime::Now;
 r->PollButtonEvents(nullptr,EventArgs::Empty);if(r->polls!=0)return 6;
 r->buttonEventsRequestTime=DateTime::Now.AddSeconds(-61);
 r->PollButtonEvents(nullptr,EventArgs::Empty);if(r->polls!=0)return 7;
 r->HandleButtonEventsResponse(packet());if(r->buttonEventsUnsupported)return 8;
 r->buttonEventsUnsupported=true;r->buttonEventsPending=true;r->ResetButtonEventPolling();
 r->controllerRestartVersionPending=true;
 if(r->buttonEventsUnsupported || r->buttonEventsPending)return 9;
 r->PollButtonEvents(nullptr,EventArgs::Empty);
 if(r->versionRequests!=1 || r->polls!=0)return 10;
 r->PollButtonEvents(nullptr,EventArgs::Empty);
 if(r->versionRequests!=1 || r->polls!=1)return 11;
 r->unifiedEvents=true;r->eventsAvailable=false;
 r->PollButtonEvents(nullptr,EventArgs::Empty);if(r->polls!=1)return 12;
 r->eventsAvailable=true;r->buttonEventsPending=false;
 r->PollButtonEvents(nullptr,EventArgs::Empty);if(r->polls!=2)return 13;
 auto diag=gcnew array<Byte>(67);
 diag[0]=3;diag[1]=CmdRequest::GET_EVENT;diag[3]=61;diag[4]=100;
 diag[28]=2;diag[29]=8;diag[37]=3;
 diag[38]=2;diag[39]=103;diag[40]=3;diag[41]=5;diag[42]=9;diag[43]=9;diag[45]=9;
 pin_ptr<Byte> d=&diag[0];auto sum=CalculateCommandCRC(d,65);diag[65]=sum&255;diag[66]=sum>>8;
 r->HandleButtonEventsResponse(diag);if(r->acks!=4 || r->acknowledged!=8)return 14;
 r->HandleButtonEventsResponse(diag);if(r->acks!=5)return 15;
 auto field=GlobalLogger::typeid->GetField("writer",System::Reflection::BindingFlags::NonPublic|System::Reflection::BindingFlags::Static);
 safe_cast<StreamWriter^>(field->GetValue(nullptr))->BaseStream->Close();
 auto next=packet();next[25]=2;pin_ptr<Byte> ptr=&next[0];uint16_t crc=CalculateCommandCRC(ptr,34);next[34]=crc&255;next[35]=crc>>8;
 r->HandleButtonEventsResponse(next);if(r->acks!=5)return 5;
 Console::WriteLine("PASS: response handler, CRC, duplicate ACK, disk failure and same-socket firmware upgrade recovery");
}
'''
        cpp=folder/"receiver.cpp";cpp.write_text(prefix,encoding="utf-8")
        subprocess.run(["cl","/nologo","/clr","/EHa","/utf-8","/I"+str(ROOT),str(cpp),str(ROOT/"Commands.cpp"),"/Fe:"+str(folder/"receiver.exe")],cwd=folder,check=True)
        subprocess.run([str(folder/"receiver.exe")],check=True)



if __name__ == "__main__":
    run()
