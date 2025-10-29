# ����������: SendStartCommand � SendStopCommand

**����:** 25 ������� 2025  
**������:** 1.1

---

## ? ��� ���� ���������

### ��������:
������ `SendStartCommand()` � `SendStopCommand()` � DataForm.cpp **�� ������������ �������� ������ �� �����������**, ���� � ������������ RESPONSE_HANDLING_GUIDE.md ��� ������ ������ � ��������� ������.

### �������:
��������� ��� ������ ��� ������������� `SendCommandAndWaitResponse()` ������ �������� `SendCommand()`.

---

## ?? ��������� � DataForm.cpp

### SendStartCommand() - ��:

```cpp
void ProjectServerW::DataForm::SendStartCommand() {
    Command cmd = CreateControlCommand(CmdProgControl::START);
    
    if (SendCommand(cmd)) {
        // ������� ������� ����������
        buttonSTOPstate_TRUE();
    }
}
```

### SendStartCommand() - �����:

```cpp
void ProjectServerW::DataForm::SendStartCommand() {
    // ������� ������� START
    Command cmd = CreateControlCommand(CmdProgControl::START);
    CommandResponse response;
    
    // ���������� ������� � ���� �����
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

---

### SendStopCommand() - ��:

```cpp
void ProjectServerW::DataForm::SendStopCommand() {
    Command cmd = CreateControlCommand(CmdProgControl::STOP);
    
    if (SendCommand(cmd)) {
        // ������� ������� ����������
        buttonSTARTstate_TRUE();
    }
}
```

### SendStopCommand() - �����:

```cpp
void ProjectServerW::DataForm::SendStopCommand() {
    // ������� ������� STOP
    Command cmd = CreateControlCommand(CmdProgControl::STOP);
    CommandResponse response;
    
    // ���������� ������� � ���� �����
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

---

## ?? ������������ ����������

### 1. ��������������� ����������
**��:** ������ �������� ����� ����� �������� �������, ���� ���� ���������� � �� ��������.

**�����:** ������ �������� ������ ����� ������������� �� �����������, ��� ������� ���������.

### 2. ��������� ������
**��:** ��� ���������� �� ������� ���������� �� �����������.

**�����:** ������ ���������� � ���������� ���������� ������� � �����.

### 3. ������������� ���������
**��:** ��������� UI ����� �� ��������������� ��������� ��������� �����������.

**�����:** UI ������ �������� ����������� ��������� �����������.

---

## ?? ��������� ���������

### �������� 1: ������� ������ START ��� ���������� ������

| ���� | �� ���������� | ����� ���������� |
|------|---------------|------------------|
| 1. ������� ������ | �������� ������� | �������� ������� |
| 2. ��������� �������� | ������ �������� ���������� | �������� ������ (�� 2 ���) |
| 3. ����� ����������� | ������������ | ����������� ������ |
| 4. UI ���������� | ��� �������� | �������� ������ ��� OK |
| 5. ����������� | ������ ���� �������� | ��������� ���������� |

### �������� 2: ������� ������ START ��� ��������� �����

| ���� | �� ���������� | ����� ���������� |
|------|---------------|------------------|
| 1. ������� ������ | �������� ������� | �������� ������� |
| 2. ������� ����� | ������ ��� �������� | ������ �������� ��� ��������� |
| 3. UI ��������� | �� ������������� ����������� | ������������� ����������� |
| 4. ������������ | ������, ��� ������� ��������� | �����, ��� ������� �� ��������� |

---

## ?? �������� ������ (����� ����������)

```
������������ �������� ������ START
           ?
   SendStartCommand()
           ?
   �������� ������� START
           ?
   SendCommandAndWaitResponse()
           ?
   ???????????????????????????
   ? �������� ����� �����    ?
   ???????????????????????????
              ?
   ???????????????????????????
   ? �������� ������ 2 ���   ?
   ???????????????????????????
              ?
   ???????????????????????????
   ? ����� ������            ?
   ???????????????????????????
              ?
   ???????????????????????????
   ? �������� CRC            ?
   ???????????????????????????
              ?
   ???????????????????????????
   ? �������� �������        ?
   ???????????????????????????
              ?
      ������ == OK?
       ?           ?
      ��          ���
       ?           ?
  buttonSTOP   ������ ��
  state_TRUE   ��������
       ?           ?
  ���: OK     ���: ERROR
```

---

## ?? �������������� �����������

### ���� ����� ������ ������ (��� �������� ������):

����� ������������ �������� `SendCommand()`:

```cpp
void QuickStartWithoutWait() {
    Command cmd = CreateControlCommand(CmdProgControl::START);
    
    // ������� �������� ��� �������� ������
    if (SendCommand(cmd)) {
        buttonSTOPstate_TRUE();
    }
}
```

### ���� ����� ��������� �������:

```cpp
void StartCommandWithCustomTimeout() {
    Command cmd = CreateControlCommand(CmdProgControl::START);
    CommandResponse response;
    
    // ���������� �������
    if (SendCommand(cmd)) {
        // ���� ����� � ����������� ���������
        if (ReceiveResponse(response, 5000)) { // 5 ������
            ProcessResponse(response);
            if (response.status == CmdStatus::OK) {
                buttonSTOPstate_TRUE();
            }
        }
    }
}
```

---

## ? ����������� �����

1. **DataForm.cpp** - ��������� ������ SendStartCommand() � SendStopCommand()
2. **RESPONSE_HANDLING_GUIDE.md** - �������� ������ � ������������
3. **INTEGRATION_SUMMARY.md** - �������� ������ �������������

---

## ?? ������ � �������������!

������ ��� ������� ������ START/STOP:
- ? ������� ������������ �����������
- ? ��������� ������������� ����������
- ? UI ����������� ������ ��� ������
- ? ��� ���������� ����������
- ? ������������ ����� �������� ��������� �������

---

**���� ����������:** 25 ������� 2025  
**������:** 1.1

