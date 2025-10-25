# ������� ���������: ��������� ������� �����������

## ������� �����

### 1. �������� ������� � ��������� ������ (�������������)

```cpp
// ������� �������
Command cmd = CreateControlCommand(CmdProgControl::START);
CommandResponse response;

// ���������� � ���� �����
if (SendCommandAndWaitResponse(cmd, response)) {
    // ����� - ������� ��������� (status == OK)
} else {
    // ������ - ������� �� ���������
}
```

---

### 2. �������� ������� ��� ��������

```cpp
// ������ ������ (��� �������� ������)
Command cmd = CreateControlCommand(CmdProgControl::STOP);
SendCommand(cmd);
```

---

### 3. ������ ��������� ������

```cpp
// ���������� �������
Command cmd = CreateRequestCommand(CmdRequest::GET_STATUS);
SendCommand(cmd);

// ��������� �����
CommandResponse response;
if (ReceiveResponse(response, 2000)) {  // ������� 2 ��� (���� ������)
    ProcessResponse(response);  // �������������� ���������
}

// ��� � ��������� �� ��������� (1000 ��)
if (ReceiveResponse(response)) {
    ProcessResponse(response);
}
```

---

## �������� ������

### ������� ���������� (PROG_CONTROL)

```cpp
// START, STOP, PAUSE, RESUME, RESET
Command cmd = CreateControlCommand(CmdProgControl::START);
```

### ������� ������������ (CONFIGURATION)

```cpp
// ��������� ����������� (float)
Command cmd = CreateConfigCommandFloat(CmdConfig::SET_TEMPERATURE, 25.5f);

// ��������� ��������� (int)
Command cmd = CreateConfigCommandInt(CmdConfig::SET_INTERVAL, 1000);
```

### ������� ������� (REQUEST)

```cpp
// GET_STATUS, GET_VERSION, GET_DATA
Command cmd = CreateRequestCommand(CmdRequest::GET_STATUS);
```

---

## �������� ������� ������

```cpp
CommandResponse response;
// ... ����� ������ ...

if (response.status == CmdStatus::OK) {
    // ������� ���������
} else if (response.status == CmdStatus::CRC_ERROR) {
    // ������ CRC
} else if (response.status == CmdStatus::TIMEOUT) {
    // �������
} else {
    // ������ ������
    const char* statusName = GetStatusName(response.status);
}
```

---

## ���������� ������ �� ������

### ������ ������� (2 �����)

```cpp
if (response.dataLength >= 2) {
    uint16_t status;
    memcpy(&status, response.data, 2);
    // ���������� status
}
```

### ������ ������ (������)

```cpp
if (response.dataLength > 0) {
    String^ version = gcnew String(
        reinterpret_cast<const char*>(response.data), 
        0, response.dataLength, System::Text::Encoding::ASCII);
}
```

### ������ ������������ (1 ����)

```cpp
if (response.dataLength >= 1) {
    uint8_t mode = response.data[0];
    // 0 = ��������������, 1 = ������
}
```

---

## �������� ��������

### �������� 1: ������ ���������

```cpp
void DataForm::SendStartCommand() {
    Command cmd = CreateControlCommand(CmdProgControl::START);
    CommandResponse response;
    
    if (SendCommandAndWaitResponse(cmd, response)) {
        buttonSTOPstate_TRUE();
        MessageBox::Show("��������� ��������!");
    }
}
```

### �������� 2: ������ � ����������� �������

```cpp
void DataForm::ShowDeviceStatus() {
    Command cmd = CreateRequestCommand(CmdRequest::GET_STATUS);
    CommandResponse response;
    
    if (SendCommandAndWaitResponse(cmd, response)) {
        uint16_t status;
        memcpy(&status, response.data, 2);
        String^ msg = String::Format("������: 0x{0:X4}", status);
        MessageBox::Show(msg);
    }
}
```

### �������� 3: ��������� �����������

```cpp
void DataForm::SetTemperature(float temp) {
    Command cmd = CreateConfigCommandFloat(CmdConfig::SET_TEMPERATURE, temp);
    CommandResponse response;
    
    if (SendCommandAndWaitResponse(cmd, response)) {
        String^ msg = String::Format("����������� �����������: {0:F1}�C", temp);
        MessageBox::Show(msg);
    }
}
```

---

## �������� (�������������)

| ��� �������            | �������    |
|------------------------|------------|
| ������� ����������     | 1000-2000 �� |
| ������� ������������   | 1000-2000 �� |
| ������� ������� ������ | 2000-3000 �� |

---

## ���� �������

| ���  | ���������           | ��������                  |
|------|---------------------|---------------------------|
| 0x00 | CmdStatus::OK       | ? �������                 |
| 0x01 | CmdStatus::CRC_ERROR | ? ������ CRC             |
| 0x02 | CmdStatus::INVALID_TYPE | ? �������� ���        |
| 0x03 | CmdStatus::INVALID_CODE | ? �������� ���        |
| 0x04 | CmdStatus::INVALID_LENGTH | ? �������� �����    |
| 0x05 | CmdStatus::EXECUTION_ERROR | ? ������ ���������� |
| 0x06 | CmdStatus::TIMEOUT | ? �������                  |
| 0xFF | CmdStatus::UNKNOWN_ERROR | ? ����������� ������ |

---

## �������

### �������� CRC

```cpp
uint8_t buffer[64];
size_t bytesReceived = recv(socket, buffer, sizeof(buffer), 0);

if (ValidateResponseCRC(buffer, bytesReceived)) {
    // CRC ���������
} else {
    // ������ CRC - ��������� ����� �����
}
```

### �����������

��� �������� ������������� ���������� � `log.txt`:
- �������� ������
- ����� �������
- ������ ����������

---

## ������ ���������

?? **��������:**
- ��� ������� ������������� ��������� CRC
- ������ ������������ ����� `MessageBox` � ����������
- ����� `ProcessResponse()` ������������� ������������ ��� ���� �������
- ����������� `SendCommandAndWaitResponse()` ��� ��������� ������

? **������������:**
- ������ ���������� ������������ �������� �������
- ����������� ���������� ��������
- ��������� ��� �������� ��� �������
- ������������� ��� ��������� ������� �������

---

**����:** 25 ������� 2025  
**������:** 1.0

