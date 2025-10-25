# ������� �������� ��������� recv()

## ?? ������������� �������: ���������� ������� �� ����

### �����������:

```
???????????????????????????????????????????
?         ClientHandler Thread            ?
?     (SServer.cpp - ����������� ����)   ?
???????????????????????????????????????????
             ? recv() �������� ��� ������
             ?
      ????????????????
      ?�������� ���� ?
      ?������� ����� ?
      ????????????????
             ?
      ?????????????????
      ?               ?
????????????    ???????????????
?����������?    ?����� �������?
?(0x00?)   ?    ?(0x01-0x04)  ?
????????????    ???????????????
     ?                 ?
     ?                 ?
AddDataToTable   ResponseQueue
                      ?
              ReceiveResponse()
              ���� � �������
```

---

## ��� ����������:

### ��� 1: �������� ������� ������� � DataForm.h

```cpp
// � DataForm.h
private:
    // ������� ��� �������� ������� �� �����������
    static System::Collections::Concurrent::ConcurrentQueue<cli::array<System::Byte>^>^ responseQueue;
    static System::Threading::Semaphore^ responseAvailable;
    
public:
    // ����� ��� ���������� ������ � ������� (���������� �� SServer)
    static void EnqueueResponse(cli::array<System::Byte>^ response);
```

### ��� 2: ������������� � DataForm.cpp

```cpp
// ����������� ����������
System::Collections::Concurrent::ConcurrentQueue<cli::array<System::Byte>^>^ ProjectServerW::DataForm::responseQueue = 
    gcnew System::Collections::Concurrent::ConcurrentQueue<cli::array<System::Byte>^>();

System::Threading::Semaphore^ ProjectServerW::DataForm::responseAvailable = 
    gcnew System::Threading::Semaphore(0, 100);

// ����� ��� ���������� ������ � �������
void ProjectServerW::DataForm::EnqueueResponse(cli::array<System::Byte>^ response) {
    responseQueue->Enqueue(response);
    responseAvailable->Release(); // ������������� � ����������� ������
}
```

### ��� 3: �������� ClientHandler � SServer.cpp

```cpp
// � SServer.cpp - ClientHandler
while (true) {
    bytesReceived = recv(clientSocket, buffer, sizeof(buffer), 0);
    
    if (bytesReceived == SOCKET_ERROR || bytesReceived == 0) {
        // ��������� ������
        break;
    }
    
    // ===== ���������� ���� ������ =====
    
    // ��������� ������ ���� ��� ����������� ���� ������
    uint8_t packetType = buffer[0];
    
    if (packetType >= 0x01 && packetType <= 0x04) {
        // ===== ��� ����� �� ������� =====
        // ������ � Type = 0x01-0x04 ��� �������/������
        
        // ������� ����������� ������ ��� ������
        cli::array<System::Byte>^ responseBuffer = gcnew cli::array<System::Byte>(bytesReceived);
        Marshal::Copy(IntPtr(buffer), responseBuffer, 0, bytesReceived);
        
        // ��������� � ������� �������
        DataForm::EnqueueResponse(responseBuffer);
        
        GlobalLogger::LogMessage(String::Format(
            "Response received: Type=0x{0:X2}, Size={1} bytes", 
            packetType, bytesReceived));
    }
    else {
        // ===== ��� ���������� =====
        // ����������� ��������� ���������� (48 ����)
        
        uint8_t LengthOfPackage = 48;
        uint16_t dataCRC = MB_GetCRC(buffer, LengthOfPackage - 2);
        uint16_t DatCRC;
        memcpy(&DatCRC, &buffer[LengthOfPackage - 2], 2);
        
        if (DatCRC == dataCRC) {
            // ���������� ���������� �������
            send(clientSocket, buffer, bytesReceived, 0);
            
            // ������������ ����������
            DataForm^ form2 = DataForm::GetFormByGuid(guid);
            if (form2 != nullptr && !form2->IsDisposed) {
                cli::array<System::Byte>^ dataBuffer = gcnew cli::array<System::Byte>(bytesReceived);
                Marshal::Copy(IntPtr(buffer), dataBuffer, 0, bytesReceived);
                
                form2->BeginInvoke(gcnew Action<cli::array<System::Byte>^, int, int>
                    (form2, &DataForm::AddDataToTableThreadSafe),
                    dataBuffer, bytesReceived, clientPort);
            }
        }
    }
}
```

### ��� 4: �������� ReceiveResponse � DataForm.cpp

```cpp
bool ProjectServerW::DataForm::ReceiveResponse(CommandResponse& response, int timeoutMs) {
    try {
        // ���������, ��� ����� ������� ������
        if (clientSocket == INVALID_SOCKET) {
            GlobalLogger::LogMessage("Error: Client socket is invalid for receiving response");
            return false;
        }

        // ===== �������� ������ �� ������� =====
        // ���� ��������� ������ � ������� � ���������
        if (!responseAvailable->WaitOne(timeoutMs)) {
            String^ msg = "Timeout: No response received from controller";
            GlobalLogger::LogMessage(ConvertToStdString(msg));
            return false;
        }

        // �������� ����� �� �������
        cli::array<System::Byte>^ responseBuffer;
        if (!responseQueue->TryDequeue(responseBuffer)) {
            String^ msg = "Error: Failed to dequeue response";
            GlobalLogger::LogMessage(ConvertToStdString(msg));
            return false;
        }

        // �������� ������ � ������������� �����
        uint8_t buffer[64];
        pin_ptr<System::Byte> pinnedBuffer = &responseBuffer[0];
        memcpy(buffer, pinnedBuffer, responseBuffer->Length);

        // ��������� ���������� �����
        if (!ParseResponseBuffer(buffer, responseBuffer->Length, response)) {
            String^ msg = "Error: Failed to parse response from controller";
            GlobalLogger::LogMessage(ConvertToStdString(msg));
            return false;
        }

        // �������� �������� ��������� ������
        String^ logMsg = String::Format(
            "Response processed: Type=0x{0:X2}, Code=0x{1:X2}, Status={2}",
            response.commandType, response.commandCode, 
            gcnew String(GetStatusName(response.status)));
        GlobalLogger::LogMessage(ConvertToStdString(logMsg));

        return true;

    } catch (Exception^ ex) {
        String^ errorMsg = "Exception in ReceiveResponse: " + ex->Message;
        GlobalLogger::LogMessage(ConvertToStdString(errorMsg));
        return false;
    }
}
```

---

## ?? ������������ �������:

? **����������** - ���������� �� ���� ������ (������ ����)  
? **����������������** - ����� �������� ����� ���� �������  
? **������������������** - ������������ ConcurrentQueue  
? **������������������** - ����������� ��������� �������  
? **�������������** - �� �������� ������������ ��������� ����������

---

## ?? ������� ����� �������:

| ������ ���� | ��� ������ | ���������� |
|-------------|------------|------------|
| 0x00 | ���������� | AddDataToTable |
| 0x01 | PROG_CONTROL response | ResponseQueue |
| 0x02 | CONFIGURATION response | ResponseQueue |
| 0x03 | REQUEST response | ResponseQueue |
| 0x04 | DEVICE_CONTROL response | ResponseQueue |

---

## ?? ������ ���������:

1. **������������� �������** ������ ���� �� ������� �������������
2. **������ �������** (100) ����� ��������� ��� �����
3. **������� ��������** ������������� �������� �������� ������
4. **������� �������** ��� �������� �����

---

## ?? �������������� �������:

### ������� 2: ��� ������
- ��������� ����� ��� ������
- ��������� ����� ��� ����������
- ? ������� ��������� �����������

### ������� 3: ������� �� recv()
- ���������� ������ ��� �������� �������
- ? ��������� ���������� �� ����� �������� ������

---

**������������:** ������������ **������� � ��������** (������� ����).

