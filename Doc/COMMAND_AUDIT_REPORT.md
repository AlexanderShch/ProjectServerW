# ����� �������� ������ �����������

**���� ��������:** 25 ������� 2025  
**��������� ������:** 6 (.cpp ������)

---

## ? ���������� ��������

### ?? ����� � ��������� ������:

#### **DataForm.cpp** ? ��� ������� ������� �����

| ����� | ������ | ������� | ����� �������� | ������ |
|-------|--------|---------|----------------|--------|
| `SendStartCommand()` | 1040 | START | `SendCommandAndWaitResponse()` | ? ������� ����� |
| `SendStopCommand()` | 1057 | STOP | `SendCommandAndWaitResponse()` | ? ������� ����� |

---

### ?? ����� ��� �������� ������:

| ���� | ��������� �������� |
|------|--------------------|
| **SServer.cpp** | ? ������� �� ������������ |
| **MyForm.cpp** | ? ������� �� ������������ |
| **Chart.cpp** | ? ������� �� ������������ |
| **Commands.cpp** | ? ������ ����������� ������� (�� ��������) |
| **MyForm_fixed.cpp** | ? ������� �� ������������ |

---

## ?? ��������� ������

### SendStartCommand() - ������ 1040

```cpp
void ProjectServerW::DataForm::SendStartCommand() {
    // ������� ������� START
    Command cmd = CreateControlCommand(CmdProgControl::START);
    CommandResponse response;
    
    // ���������� ������� � ���� ����� ?
    if (SendCommandAndWaitResponse(cmd, response)) {
        // ������� ������� ��������� �� �����������
        buttonSTOPstate_TRUE();
        GlobalLogger::LogMessage("Information: ������� START ������� ��������� ������������");
    } else {
        // ������ ���������� �������
        GlobalLogger::LogMessage("Error: ������� START �� ��������� ������������");
    }
}
```

**��������������:**
- ? ���������� `SendCommandAndWaitResponse()`
- ? ������� ����� �� ����������� (������� 2 ���)
- ? ��������� CRC ������
- ? ��������� ������ ����������
- ? UI ����������� ������ ��� ������
- ? �������� ��������� ����������

---

### SendStopCommand() - ������ 1057

```cpp
void ProjectServerW::DataForm::SendStopCommand() {
    // ������� ������� STOP
    Command cmd = CreateControlCommand(CmdProgControl::STOP);
    CommandResponse response;
    
    // ���������� ������� � ���� ����� ?
    if (SendCommandAndWaitResponse(cmd, response)) {
        // ������� ������� ��������� �� �����������
        buttonSTARTstate_TRUE();
        GlobalLogger::LogMessage("Information: ������� STOP ������� ��������� ������������");
    } else {
        // ������ ���������� �������
        GlobalLogger::LogMessage("Error: ������� STOP �� ��������� ������������");
    }
}
```

**��������������:**
- ? ���������� `SendCommandAndWaitResponse()`
- ? ������� ����� �� ����������� (������� 2 ���)
- ? ��������� CRC ������
- ? ��������� ������ ����������
- ? UI ����������� ������ ��� ������
- ? �������� ��������� ����������

---

## ?? ������� ��������� �������

```
������������
    ?
SendStartCommand() / SendStopCommand()
    ?
CreateControlCommand()
    ?
SendCommandAndWaitResponse()
    ?
?????????????????????????????????????
? SendCommand()                     ? ? �������� ����� �����
?????????????????????????????????????
                ?
?????????????????????????????????????
? ReceiveResponse(response, 2000)   ? ? �������� ������ 2 ���
?????????????????????????????????????
                ?
?????????????????????????????????????
? ParseResponseBuffer()             ? ? ������ ������
?????????????????????????????????????
                ?
?????????????????????????????????????
? ValidateResponseCRC()             ? ? �������� CRC16
?????????????????????????????????????
                ?
?????????????????????????????????????
? �������� ������������ �������     ? ? Type + Code
?????????????????????????????????????
                ?
?????????????????????????????????????
? ProcessResponse()                 ? ? ��������� ������
?????????????????????????????????????
                ?
        response.status == OK?
         ?              ?
        ��            ���
         ?              ?
    UI ��������    UI ��� ���������
         ?              ?
    ���: SUCCESS   ���: ERROR
```

---

## ? �������� ��������

### 1. �������� ����������� ������
- ? **CRC16 (ModBus)** ����������� ��� ������� ������
- ? **������ ������** ������������ ����� ����������

### 2. �������� ���������� �������
- ? **������ ����������** ����������� (OK/ERROR)
- ? **������������ �������** ����������� (Type + Code)

### 3. ��������� ������
- ? **�������** - 2000 �� ��� ���� ������
- ? **�����������** - ��� ���������� ������������ � log.txt
- ? **UI �������������** - ���������� ������ ��� ������

### 4. ������������������
- ? **��������� ��������** - ������������ ����� ��������� �������
- ? **��������������** - ������������ ����� ��������� � ����
- ? **��������� UI** - ������ ������������� ��������� ���������

---

## ?? ������������

### ? ������� ���������� ������������� �����������:

1. ? **��� ������� ������� �����** �� �����������
2. ? **��� ������ �����������** (CRC + ������)
3. ? **UI ���������������** � ���������� �����������
4. ? **�����������** ���� ��������
5. ? **��������� ������** �����������

### ?? �������������� ��������� (�����������):

#### 1. ���������� ���������� ��������

```cpp
void ProjectServerW::DataForm::SendStartCommand() {
    // ���������� ��������� ��������
    Cursor = Cursors::WaitCursor;
    
    Command cmd = CreateControlCommand(CmdProgControl::START);
    CommandResponse response;
    
    if (SendCommandAndWaitResponse(cmd, response)) {
        buttonSTOPstate_TRUE();
        GlobalLogger::LogMessage("Information: ������� START ������� ��������� ������������");
    } else {
        GlobalLogger::LogMessage("Error: ������� START �� ��������� ������������");
    }
    
    // ���������� ������� ������
    Cursor = Cursors::Default;
}
```

#### 2. ���������� �������� �������

```cpp
void ProjectServerW::DataForm::SendStartCommandWithRetry() {
    Command cmd = CreateControlCommand(CmdProgControl::START);
    CommandResponse response;
    
    const int maxRetries = 3;
    for (int attempt = 1; attempt <= maxRetries; attempt++) {
        if (SendCommandAndWaitResponse(cmd, response)) {
            buttonSTOPstate_TRUE();
            GlobalLogger::LogMessage("Information: ������� START ������� ��������� ������������");
            return;
        }
        
        if (attempt < maxRetries) {
            GlobalLogger::LogMessage(String::Format(
                "Warning: ������� {0} �� {1} �� �������, ������...", 
                attempt, maxRetries));
            Thread::Sleep(500); // ����� ����� ��������
        }
    }
    
    GlobalLogger::LogMessage("Error: ������� START �� ��������� ����� ���� �������");
}
```

#### 3. ���������� ���������� ������

```cpp
// � ������ DataForm �������� ��������:
private:
    int commandsSent = 0;
    int commandsSucceeded = 0;
    int commandsFailed = 0;

// � ������� ��������� ��������:
void ProjectServerW::DataForm::SendStartCommand() {
    commandsSent++;
    
    Command cmd = CreateControlCommand(CmdProgControl::START);
    CommandResponse response;
    
    if (SendCommandAndWaitResponse(cmd, response)) {
        commandsSucceeded++;
        buttonSTOPstate_TRUE();
    } else {
        commandsFailed++;
    }
    
    UpdateStatisticsLabel(); // �������� UI �� �����������
}
```

---

## ?? �������� �������

| �������� | ������ | ������ |
|----------|--------|--------|
| **���������� ������� ��������** | 2 | SendStartCommand, SendStopCommand |
| **������������� �������� ������** | 100% | ��� ������� ���������� SendCommandAndWaitResponse |
| **�������� CRC** | ? �� | ModBus CRC16 ��� ���� ������� |
| **�������� �������** | ? �� | ��� ������ ����������� �� OK |
| **�������** | 2000 �� | ��� ���� ������ |
| **�����������** | ? �� | ����� � ������ |
| **������������� UI** | ? �� | ������ ��� ������ |
| **��������� ������** | ? �� | ������ |

---

## ? ����������

**��� ������� � ������� ��������� ����������� � ������� ����� �� �����������.**

������� ����������:
- ? ����������� �������� �������� ������
- ? ������������ ������������� ���������
- ? ������������� ������ ���������� � �����������
- ? ������������� ��������� �����������

**������������:** ������� ���������� ������ � ������������� � ����������.

---

**���� ������:** 25 ������� 2025  
**������:** ? PASSED  
**��������� �������:** 2  
**������� �������:** 0

