# ����������� �� ��������� ������� �����������

**������:** ProjectServerW  
**������:** 1.0  
**����:** 25 ������� 2025

---

## ����������

1. [�����](#�����)
2. [����� ��������� � ���������](#�����-���������-�-���������)
3. [API �������](#api-�������)
4. [������� �������������](#�������-�������������)
5. [��������� ������](#���������-������)

---

## �����

� ������ ��������� ������ ��������� ������ � ��������� ������� �� ����������� �� ������������ �������. ���������� ��������:

- **��������� ������** (`CommandResponse`)
- **���� �������** ���������� ������ (`CmdStatus`)
- **������� �������** � �������� CRC �������
- **������ ������** � ��������� ������� � `DataForm`

---

## ����� ��������� � ���������

### ��������� CommandResponse

```cpp
struct CommandResponse {
    uint8_t commandType;           // ��� �������� �������
    uint8_t commandCode;           // ��� �������� �������
    uint8_t status;                // ������ ���������� �������
    uint8_t dataLength;            // ����� ������ ������
    uint8_t data[58];              // ������ ������ (����. 58 ����)
    uint16_t crc;                  // CRC16 ��� �������� �����������
};
```

### ���� ������� (CmdStatus)

| ���  | ���������           | ��������                        |
|------|---------------------|---------------------------------|
| 0x00 | OK                  | ������� ��������� �������       |
| 0x01 | CRC_ERROR           | ������ ����������� �����        |
| 0x02 | INVALID_TYPE        | ����������� ��� �������         |
| 0x03 | INVALID_CODE        | ����������� ��� �������         |
| 0x04 | INVALID_LENGTH      | �������� ����� ������           |
| 0x05 | EXECUTION_ERROR     | ������ ���������� �������       |
| 0x06 | TIMEOUT             | ������� ����������              |
| 0xFF | UNKNOWN_ERROR       | ����������� ������              |

---

## API �������

### Commands.h / Commands.cpp

#### `bool ValidateResponseCRC(const uint8_t* buffer, size_t length)`

**��������:** ��������� CRC16 ����������� ������

**���������:**
- `buffer` - ����� � ������� ������
- `length` - ����� ������

**����������:** `true` ���� CRC ���������, `false` ���� ���

**������:**
```cpp
uint8_t buffer[64];
size_t receivedBytes = recv(socket, buffer, sizeof(buffer), 0);

if (ValidateResponseCRC(buffer, receivedBytes)) {
    // CRC ���������, ����� ������������ ������
}
```

---

#### `bool ParseResponseBuffer(const uint8_t* buffer, size_t bufferSize, CommandResponse& response)`

**��������:** ��������� ����� ������ � ��������� `CommandResponse`

**���������:**
- `buffer` - ����� � ������� ������
- `bufferSize` - ������ ������
- `response` - ��������� ��� ������ ������������ ������

**����������:** `true` ��� �������� �������, `false` ��� ������

**������:**
```cpp
uint8_t buffer[64];
size_t receivedBytes = recv(socket, buffer, sizeof(buffer), 0);

CommandResponse response;
if (ParseResponseBuffer(buffer, receivedBytes, response)) {
    // ����� ������� ��������
    if (response.status == CmdStatus::OK) {
        // ������� ��������� �������
    }
}
```

---

#### `const char* GetStatusName(uint8_t status)`

**��������:** ���������� ��������� �������� ���� �������

**���������:**
- `status` - ��� �������

**����������:** ������ � ��������� �������

**������:**
```cpp
CommandResponse response;
// ... ����� � ������ ������ ...

const char* statusName = GetStatusName(response.status);
printf("Status: %s\n", statusName); // �������: "Status: OK"
```

---

#### `bool CommandRequiresResponse(const Command& cmd)`

**��������:** ����������, ������� �� ������� ������ �� �����������

**���������:**
- `cmd` - ������� ��� ��������

**����������:** `true` ���� ������� ������� ������

**����������:** ��� ������� ���� `REQUEST` ������ ������� ������ � �������

---

### DataForm ������

#### `bool ReceiveResponse(CommandResponse& response, int timeoutMs = 1000)`

**��������:** ��������� ����� �� ����������� ����� �����

**���������:**
- `response` - ��������� ��� ������ ����������� ������
- `timeoutMs` - ������� �������� � ������������� (�� ��������� 1000 ��)

**����������:** `true` ��� �������� ������ ������, `false` ��� ������

**������:**
```cpp
CommandResponse response;
if (ReceiveResponse(response, 2000)) {
    // ����� �������
    ProcessResponse(response);
}
```

---

#### `void ProcessResponse(const CommandResponse& response)`

**��������:** ������������ ���������� ����� �� �����������

**���������:**
- `response` - ��������� � ���������� �������

**����������������:**
- ��������� ������ ���������� �������
- ��������� ������ ��� ������ ���� REQUEST
- ���������� �������������� ���������
- �������� ����������

**������:**
```cpp
CommandResponse response;
if (ReceiveResponse(response)) {
    ProcessResponse(response);
    // ������������� ����������� ��������� ����������
}
```

---

#### `bool SendCommandAndWaitResponse(const Command& cmd, CommandResponse& response, String^ commandName = nullptr)`

**��������:** ���������� ������� � ������� ����� �� �����������

**���������:**
- `cmd` - ������� ��� ��������
- `response` - ��������� ��� ������ ������
- `commandName` - ��� ������� ��� ����������� (�����������)

**����������:** `true` ���� ������� ������� ��������� (������ OK), `false` ��� ������

**������:**
```cpp
Command cmd = CreateControlCommand(CmdProgControl::START);
CommandResponse response;

if (SendCommandAndWaitResponse(cmd, response)) {
    // ������� ������� ���������
    buttonSTOPstate_TRUE();
}
```

---

## ������� �������������

### ������ 1: �������� ������� START � ��������� ������

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

### ������ 2: ������ ������� ����������

```cpp
void ProjectServerW::DataForm::RequestDeviceStatus() {
    // ������� ������� ������� �������
    Command cmd = CreateRequestCommand(CmdRequest::GET_STATUS);
    CommandResponse response;
    
    // ���������� ������� � ���� �����
    if (SendCommandAndWaitResponse(cmd, response, "GET_STATUS")) {
        if (response.dataLength >= 2) {
            // ��������� ������ �� ������
            uint16_t deviceStatus;
            memcpy(&deviceStatus, response.data, 2);
            
            // ���������� ������
            String^ statusMsg = String::Format("������ ����������: 0x{0:X4}", deviceStatus);
            MessageBox::Show(statusMsg);
        }
    }
}
```

---

### ������ 3: ������ ������ ��������

```cpp
void ProjectServerW::DataForm::GetFirmwareVersion() {
    // ������� ������� ������� ������
    Command cmd = CreateRequestCommand(CmdRequest::GET_VERSION);
    CommandResponse response;
    
    // ���������� ������� � ���� �����
    if (SendCommandAndWaitResponse(cmd, response)) {
        if (response.dataLength > 0) {
            // ��������� ������ �� ������
            String^ version = gcnew String(
                reinterpret_cast<const char*>(response.data), 
                0, response.dataLength, System::Text::Encoding::ASCII);
            
            // ���������� ������
            String^ msg = "������ ��������: " + version;
            MessageBox::Show(msg);
        }
    }
}
```

---

### ������ 4: ��������� ����������� � ��������� ����������

```cpp
void ProjectServerW::DataForm::SetTargetTemperature(float temperature) {
    // ������� ������� ��������� �����������
    Command cmd = CreateConfigCommandFloat(CmdConfig::SET_TEMPERATURE, temperature);
    CommandResponse response;
    
    // ���������� ������� � ���� �����
    if (SendCommandAndWaitResponse(cmd, response)) {
        String^ msg = String::Format(
            "������� ����������� �����������: {0:F1}�C", temperature);
        MessageBox::Show(msg);
    } else {
        // ��������� ������
        String^ errorMsg = String::Format(
            "������ ��������� �����������!\n������: {0}", 
            gcnew String(GetStatusName(response.status)));
        MessageBox::Show(errorMsg);
    }
}
```

---

### ������ 5: ������ ��������� ������

```cpp
void ProjectServerW::DataForm::ManualResponseHandling() {
    // ������� � ���������� �������
    Command cmd = CreateControlCommand(CmdProgControl::PAUSE);
    
    if (SendCommand(cmd)) {
        // ������� ����������, ���� ����� �������
        CommandResponse response;
        
        if (ReceiveResponse(response, 3000)) { // ������� 3 �������
            // ����� �������, ��������� ������
            if (response.status == CmdStatus::OK) {
                MessageBox::Show("��������� ��������������");
            } else if (response.status == CmdStatus::EXECUTION_ERROR) {
                MessageBox::Show("�� ������� ������������� ���������");
            } else {
                String^ msg = String::Format("������: {0}", 
                    gcnew String(GetStatusName(response.status)));
                MessageBox::Show(msg);
            }
        } else {
            MessageBox::Show("������� �������� ������");
        }
    }
}
```

---

## ��������� ������

### ���� ������ � �� ���������

#### 1. ������ CRC (CmdStatus::CRC_ERROR)

��������� ��� ����������� ������ �� ����� ��������.

**���������:**
```cpp
if (response.status == CmdStatus::CRC_ERROR) {
    // ��������� �������� �������
    SendCommandAndWaitResponse(cmd, response);
}
```

---

#### 2. �������� ���/��� ������� (INVALID_TYPE, INVALID_CODE)

���������� �� ������������ ������ �������.

**���������:**
```cpp
if (response.status == CmdStatus::INVALID_TYPE || 
    response.status == CmdStatus::INVALID_CODE) {
    MessageBox::Show("������� �� �������������� ������������");
    // ��������� ������ �������� �����������
}
```

---

#### 3. ������� ������ ������

����� �� ����������� �� ������� � ������� ��������� �������.

**���������:**
```cpp
CommandResponse response;
if (!ReceiveResponse(response, 5000)) {
    // ��������� ������� ��� ��������� ����������
    MessageBox::Show("���������� �� ��������. ��������� ����������.");
}
```

---

#### 4. �������������� ������ ������������ �������

����� �� ������������� ������������ �������.

**���������:**
```cpp
if (response.commandType != cmd.commandType || 
    response.commandCode != cmd.commandCode) {
    MessageBox::Show("������� ������������ ����� �� �����������");
    // ��������, � ������� ��������� ����� �� ���������� �������
}
```

---

## ������������ �� �������������

### 1. ����� ������ ��������

- **����������� `SendCommand()`** ����� ����� �� ��������� ��� �������������� ��������
- **����������� `SendCommandAndWaitResponse()`** ��� ������, ��������� �������������
- **����������� `ReceiveResponse()` + `ProcessResponse()`** ��� ������ ���������

### 2. ��������� ���������

- **������� ����������:** 1000-2000 ��
- **������� ������� ������:** 2000-3000 ��
- **������� ������������:** 1000-2000 ��

### 3. �����������

��� ������ ������������� ��������:
- �������� ������
- ����� �������
- ������ ����������

���� ����������� � ���� `log.txt` � ���������� ����������.

---

## ���������� ����������

**������:** ProjectServerW  
**���� ��������:** 25 ������� 2025  
**������ ���������:** 1.0

---

*����� ���������*

