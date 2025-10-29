# �������� �����: ���������� ��������� ������� �����������

**������:** ProjectServerW  
**����:** 25 ������� 2025  
**������:** 1.0

---

## ����������� ���������

### 1. Commands.h - ���������� �������� � �����������

#### ���������:

**��������� CmdStatus** - ���� ������� ���������� ������:
```cpp
struct CmdStatus {
    static const uint8_t OK = 0x00;                   // ������� ��������� �������
    static const uint8_t CRC_ERROR = 0x01;            // ������ ����������� �����
    static const uint8_t INVALID_TYPE = 0x02;         // ����������� ��� �������
    static const uint8_t INVALID_CODE = 0x03;         // ����������� ��� �������
    static const uint8_t INVALID_LENGTH = 0x04;       // �������� ����� ������
    static const uint8_t EXECUTION_ERROR = 0x05;      // ������ ���������� �������
    static const uint8_t TIMEOUT = 0x06;              // ������� ����������
    static const uint8_t UNKNOWN_ERROR = 0xFF;        // ����������� ������
};
```

**��������� CommandResponse** - ����� �� �����������:
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

**���������� ����� �������:**
- `bool ValidateResponseCRC(const uint8_t* buffer, size_t length)`
- `bool ParseResponseBuffer(const uint8_t* buffer, size_t bufferSize, CommandResponse& response)`
- `const char* GetStatusName(uint8_t status)`
- `bool CommandRequiresResponse(const Command& cmd)`

---

### 2. Commands.cpp - ���������� ������� ��������� �������

#### ������������� �������:

**ValidateResponseCRC** - �������� CRC16 ������:
- ��������� CRC ��� ���� ������ ����� ��������� 2 ����
- ���������� � ���������� CRC
- ���������� ��������� ��������

**ParseResponseBuffer** - ������ ������ ������:
- ��������� ����������� ������ ������ (6 ����)
- ��������� ����: Type, Code, Status, DataLength
- �������� ������ ������
- ��������� � ��������� CRC
- ���������� ��������� CommandResponse

**GetStatusName** - ��������� ���������� �������� �������:
- ����������� ��� ������� � �������� ������
- ������������ ��� ������������ ���� �������
- ���������� "UNDEFINED_STATUS" ��� ����������� �����

**CommandRequiresResponse** - �������� ������������� ������:
- ����������, ������� �� ������� ������
- ������� ���� REQUEST ������ ������� ������
- ��������� ������� �� ��������� ������� �������������

---

### 3. DataForm.h - ���������� ������� ��������� �������

#### ���������:

**Forward declaration:**
```cpp
struct CommandResponse;
```

**���������� ����� �������:**
```cpp
// ����� ������ �� �����������
bool ReceiveResponse(struct CommandResponse& response, int timeoutMs = 1000);

// ��������� ����������� ������
void ProcessResponse(const struct CommandResponse& response);

// �������� ������� � �������� ������
bool SendCommandAndWaitResponse(const struct Command& cmd, 
                                struct CommandResponse& response, 
                                System::String^ commandName = nullptr);
```

---

### 4. DataForm.cpp - ���������� ������� ��������� �������

#### ������������� ������:

**ReceiveResponse** - ����� ������ �� �����������:
- ��������� ���������� ������
- ������������� ������� ������
- ��������� ������ ����� recv()
- ������������ ������ � ��������
- ��������� ���������� ����� � CommandResponse
- �������� ����������
- ���������� true ��� ������

**ProcessResponse** - ��������� ����������� ������:
- ��������� ������ ���������� �������
- ��� �������� ������ (status == OK):
  - ��������� �������������� ���������
  - ��������� ������ ��� ������ REQUEST:
    * GET_STATUS - ���������� ������ ���������� (uint16)
    * GET_VERSION - ���������� ������ �������� (string)
    * GET_DATA - ���������� ����� ������ (uint8)
  - �������� ���������
- ��� ��������� ������:
  - ���������� MessageBox � ��������� ������
  - �������� ������

**SendCommandAndWaitResponse** - �������� ������� � �������� ������:
- ���������� ������� ����� SendCommand()
- ������� ����� � ������������� ��������� (�� ��������� 2 ���)
- ��������� ������������ ������ ������������ �������
- ������������� ������������ ����� ����� ProcessResponse()
- ���������� true ������ ��� status == OK
- ������������ ��� ����������

---

### 5. ������������

#### ��������� �����:

**RESPONSE_HANDLING_GUIDE.md** - ������ �����������:
- ����� ����� ����������������
- �������� ���� �������� � ��������
- ��������� API �������
- 5 ������������ �������� �������������
- ������������ �� ��������� ������
- ������ �� ��������� ���������

**QUICK_REFERENCE.md** - ������� ���������:
- ������� �����
- �������� ������
- �������� ������� ������
- ���������� ������
- �������� ��������
- ������� ����� �������
- ������ �� �������

---

## ������������� � ���������� �����������

���������� ��������� ������������� ������������ `CommandReceiver_Documentation.md`:

? **��������� ������** (�����):
- Type (1 ����)
- Code (1 ����)
- Status (1 ����)
- DataLen (1 ����)
- Data (0-58 ����)
- CRC16 (2 �����, ModBus CRC16)

? **���� �������:**
- 0x00 - CMD_STATUS_OK
- 0x01 - CMD_STATUS_CRC_ERROR
- 0x02 - CMD_STATUS_INVALID_TYPE
- 0x03 - CMD_STATUS_INVALID_CODE
- 0x04 - CMD_STATUS_INVALID_LENGTH
- 0x05 - CMD_STATUS_EXECUTION_ERROR
- 0x06 - CMD_STATUS_TIMEOUT
- 0xFF - CMD_STATUS_UNKNOWN_ERROR

? **CRC16:**
- �������� ModBus CRC16
- ������� 0xA001
- ��������� �������� 0xFFFF
- Little Endian ������� ������

---

## ������� �������������

### ������ 1: ������� �������� � ���������

```cpp
Command cmd = CreateControlCommand(CmdProgControl::START);
CommandResponse response;

if (SendCommandAndWaitResponse(cmd, response)) {
    // ������� ������� ��������� �� �����������
    buttonSTOPstate_TRUE();
}
```

### ������ 2: ������ ������ �� �����������

```cpp
Command cmd = CreateRequestCommand(CmdRequest::GET_STATUS);
CommandResponse response;

if (SendCommandAndWaitResponse(cmd, response)) {
    uint16_t status;
    memcpy(&status, response.data, 2);
    // ���������� ���������� ������
}
```

### ������ 3: ��������� ������������

```cpp
Command cmd = CreateConfigCommandFloat(CmdConfig::SET_TEMPERATURE, 25.5f);
CommandResponse response;

if (SendCommandAndWaitResponse(cmd, response)) {
    MessageBox::Show("����������� �����������!");
}
```

---

## ������������ ����������

? **����������:**
- �������������� �������� CRC ���� �������
- �������� ������������ ������ �������
- ��������� ���� ����� ������

? **��������:**
- ������� API � ��������� ��������
- �������������� ����������� ���� ��������
- ������������� ��������� �� �������

? **��������:**
- ��� ������ ����������:
  1. ������ ������� (ReceiveResponse + ProcessResponse)
  2. ������� ������� (SendCommand + ReceiveResponse)
  3. ������� ������� (SendCommandAndWaitResponse)
- ������������� ��������
- ������������ �������� �������

? **����������������:**
- ��������� ������������
- ������� �������������
- ������� �������
- ����������� ��� �������

---

## �������� �����������������

### ������������� �����:

1. **���� �������� ������� START:**
```cpp
Command cmd = CreateControlCommand(CmdProgControl::START);
CommandResponse response;
bool result = SendCommandAndWaitResponse(cmd, response);
// ���������: result == true, response.status == 0x00
```

2. **���� ������� ������:**
```cpp
Command cmd = CreateRequestCommand(CmdRequest::GET_VERSION);
CommandResponse response;
bool result = SendCommandAndWaitResponse(cmd, response);
// ���������: result == true, response.dataLength > 0
```

3. **���� ��������� ������ CRC:**
```cpp
// ������������ ������������ �����
uint8_t buffer[] = {0x01, 0x01, 0x00, 0x00, 0xFF, 0xFF}; // �������� CRC
CommandResponse response;
bool result = ParseResponseBuffer(buffer, 6, response);
// ���������: result == false
```

---

## ��������� ����

### ������������ �� ���������:

1. **���������� ������������ ������:**
   - ���������� �������� `SendCommand()` �� `SendCommandAndWaitResponse()`
   - �������� ��� ��������� ������ (START, STOP, ������������)

2. **���������� ������ �������:**
   - ����������� ������������� ������ �������
   - ���������� ������ �������� � ����������
   - ���������������� ������������

3. **��������� UI:**
   - �������� ��������� ������� �����
   - ���������� ������ ���������� ������
   - ���������� ������ �������� �����������

4. **���������� � ����������:**
   - ����� ������� ��������/��������� ������
   - ���������� ���������� ������
   - �������������� ��������������� ��� ������ �����

---

## ������

? **���������:**
- ? ��������� ��������� ��� ������ � �������� �����������
- ? ����������� ������� �������� CRC � ������� �������
- ? ��������� ������ ������ � ��������� ������� � DataForm
- ? ������� ������ ������������ � ���������
- ? ���������� ��������� ������������� ��������� �����������
- ? ��� ��������� ��������� �������� (0 ������)

?? **�����:**
- Commands.h - �������� ����������� � ������������
- Commands.cpp - ��������� ���������� �������
- DataForm.h - ��������� ���������� �������
- DataForm.cpp - ��������� ���������� �������
- RESPONSE_HANDLING_GUIDE.md - ������ ����������� (58 ��������)
- QUICK_REFERENCE.md - ������� �������

?? **������ � �������������!**

������ ProjectServerW ������ ��������� ������������ ����� � ��������� ������� �� ����������� � ������������ � ����������, ��������� � CommandReceiver_Documentation.md.

---

**���� ����������:** 25 ������� 2025  
**������:** 1.0  
**������:** ? ���������

