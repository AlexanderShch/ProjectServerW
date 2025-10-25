# ������� �����������: �������� recv()

## ?? ��������
� SServer.cpp ����������� ���� `recv()` �������� **��� ������** �� ������, ������� ������ �� �������. DataForm.cpp �� ����� �������� ������.

## ? �������: ������� �������

### ��� 1: �������� � DataForm.h

```cpp
// � ������ private:
static System::Collections::Concurrent::ConcurrentQueue<cli::array<System::Byte>^>^ responseQueue;
static System::Threading::Semaphore^ responseAvailable;

// � ������ public:
static void EnqueueResponse(cli::array<System::Byte>^ response);
```

### ��� 2: ���������������� � DataForm.cpp (� ������ �����)

```cpp
// ����� using namespace...
System::Collections::Concurrent::ConcurrentQueue<cli::array<System::Byte>^>^ 
    ProjectServerW::DataForm::responseQueue = 
    gcnew System::Collections::Concurrent::ConcurrentQueue<cli::array<System::Byte>^>();

System::Threading::Semaphore^ ProjectServerW::DataForm::responseAvailable = 
    gcnew System::Threading::Semaphore(0, 100);

void ProjectServerW::DataForm::EnqueueResponse(cli::array<System::Byte>^ response) {
    responseQueue->Enqueue(response);
    responseAvailable->Release();
}
```

### ��� 3: �������� SServer.cpp ClientHandler (������ 268)

**����:**
```cpp
while (true) {
    bytesReceived = recv(clientSocket, buffer, sizeof(buffer), 0);
    // ... ������ ...
    
    // ��������� ������
    uint8_t LengthOfPackage = 48;
    // ... ��������� ���������� ...
}
```

**�����:**
```cpp
while (true) {
    bytesReceived = recv(clientSocket, buffer, sizeof(buffer), 0);
    // ... ������ ...
    
    // ���������� ���� ������
    uint8_t packetType = buffer[0];
    
    if (packetType >= 0x01 && packetType <= 0x04) {
        // ��� ����� �� �������
        cli::array<System::Byte>^ responseBuffer = 
            gcnew cli::array<System::Byte>(bytesReceived);
        Marshal::Copy(IntPtr(buffer), responseBuffer, 0, bytesReceived);
        DataForm::EnqueueResponse(responseBuffer);
        
        GlobalLogger::LogMessage(String::Format(
            "Response enqueued: Type=0x{0:X2}, Size={1}", 
            packetType, bytesReceived));
    }
    else {
        // ��� ���������� - ��������� ��� ������
        uint8_t LengthOfPackage = 48;
        // ... ��������� ���������� ...
    }
}
```

### ��� 4: �������� DataForm.cpp ReceiveResponse() (������ 1070)

**����:**
```cpp
bool ProjectServerW::DataForm::ReceiveResponse(CommandResponse& response, int timeoutMs) {
    // ... �������� ...
    
    // ������������� ������� ��� ������
    setsockopt(clientSocket, SOL_SOCKET, SO_RCVTIMEO, ...);
    
    uint8_t buffer[64];
    int bytesReceived = recv(clientSocket, ...); // ? �� ������� ������!
    
    // ... ��������� ...
}
```

**�����:**
```cpp
bool ProjectServerW::DataForm::ReceiveResponse(CommandResponse& response, int timeoutMs) {
    try {
        if (clientSocket == INVALID_SOCKET) {
            GlobalLogger::LogMessage("Error: Client socket is invalid");
            return false;
        }

        // ? ���� ����� �� �������
        if (!responseAvailable->WaitOne(timeoutMs)) {
            GlobalLogger::LogMessage("Timeout: No response from controller");
            return false;
        }

        // �������� ����� �� �������
        cli::array<System::Byte>^ responseBuffer;
        if (!responseQueue->TryDequeue(responseBuffer)) {
            GlobalLogger::LogMessage("Error: Failed to dequeue response");
            return false;
        }

        // �������� � ������������� �����
        uint8_t buffer[64];
        pin_ptr<System::Byte> pinnedBuffer = &responseBuffer[0];
        memcpy(buffer, pinnedBuffer, responseBuffer->Length);

        // ��������� �����
        if (!ParseResponseBuffer(buffer, responseBuffer->Length, response)) {
            GlobalLogger::LogMessage("Error: Failed to parse response");
            return false;
        }

        GlobalLogger::LogMessage(String::Format(
            "Response processed: Type=0x{0:X2}, Code=0x{1:X2}, Status={2}",
            response.commandType, response.commandCode, 
            gcnew String(GetStatusName(response.status))));

        return true;

    } catch (Exception^ ex) {
        GlobalLogger::LogMessage(ConvertToStdString("Exception: " + ex->Message));
        return false;
    }
}
```

---

## ?? ������������

����� �����������:

```cpp
// ���� 1: �������� START
Command cmd = CreateControlCommand(CmdProgControl::START);
CommandResponse response;

if (SendCommandAndWaitResponse(cmd, response)) {
    MessageBox::Show("? ������� ���������!");
} else {
    MessageBox::Show("? ������!");
}
```

**��������� ���������:** ? ������� ���������!

---

## ?? ������� �����������

- [ ] �������� responseQueue � responseAvailable � DataForm.h
- [ ] ���������������� ����������� ���������� � DataForm.cpp
- [ ] �������� ����� EnqueueResponse() � DataForm.cpp
- [ ] �������� ClientHandler � SServer.cpp (���������� �������)
- [ ] �������� ReceiveResponse() � DataForm.cpp (������ �� �������)
- [ ] �������������� ������
- [ ] �������������� �������� ������� START
- [ ] �������������� �������� ������� STOP
- [ ] ��������� ������ ����������

---

## ?? ����� �����������: ~30 �����

������, ����� � ���������� ��� ���������?

