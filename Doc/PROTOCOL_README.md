# �������� ������ ����������� STM32F429

## ��� ��������� � ������

� ������ **ProjectServerW** ��������� ������ ��������� ��������� ������ ��� ���������� ������������ STM32F429.

## �������� �����

### ��������:
- **SServer.h** - ��������� ��������� ��������� � ����� `ControllerProtocol`
- **SServer.cpp** - ���������� ������� ��������� ������
- **DataForm.h** - ��������� ������ ��� �������� ������ �����������
- **DataForm.cpp** - ���������� ������� �������� ������

### ����� �����:
- **CONTROLLER_PROTOCOL_USAGE.md** - ��������� ����������� � ��������� �������������

## �������� �����������

### 1. ��������������� �����
STM32F429 **����� ������������**:
- ? ������� COM-���� � �������� �������
- ? ���������� ������ ����� COM-���� (���������� ������)

**�� �����** ������������� �������������!

### 2. ���� ������

#### ���������� ���������� (PROG_CONTROL)
```cpp
SendControllerCommand_StartProgram();   // ��������� ���������
SendControllerCommand_StopProgram();    // ���������� ���������
SendControllerCommand_PauseProgram();   // �������������
SendControllerCommand_ResumeProgram();  // �����������
```

#### ���������� ������������ (DEVICE_CONTROL)
```cpp
SendControllerCommand_SetRelay(1, true);    // �������� ���� �1
SendControllerCommand_SetDefrost(true);     // �������� �������
```

#### ������� ������ (REQUEST)
```cpp
ControllerResponse response;
SendControllerCommand_GetStatus(response);   // �������� ������
SendControllerCommand_GetSensors(response);  // ������ ��������
SendControllerCommand_GetVersion(response);  // ������ ��
```

### 3. �������������� ���������
- ? �������������� ������ CRC16 (ModBus)
- ? �������� ����������� ������
- ? ��������� ������ � ���������
- ? ����������� ���� ��������

### 4. ������������
- �������� ���������� �������
- ��������� ���� ��������� ������
- ������ �� ������������ �������
- �������� ��� �������������� ���������

## ������� �����

### ������ 1: ������ ���������
```cpp
// � ����������� ������
bool success = SendControllerCommand_StartProgram();
if (success) {
    MessageBox::Show("��������� ��������");
}
```

### ������ 2: ���������� ����
```cpp
// �������� ���� �1
SendControllerCommand_SetRelay(1, true);

// ��������� ���� �2
SendControllerCommand_SetRelay(2, false);
```

### ������ 3: ��������� ������ ��������
```cpp
ControllerResponse response;
if (SendControllerCommand_GetSensors(response)) {
    // ��������� ������ �� response.data
    // ����� ������: response.dataLen
}
```

## ������ ���������

### ������� (������ ? ����������)
```
[Type][Code][DataLen][Data...][CRC16]
 1�    1�    1�       0-59�    2�
```

### ����� (���������� ? ������)
```
[Type][Code][Status][DataLen][Data...][CRC16]
 1�    1�    1�      1�       0-59�    2�
```

## ������� �������

| ��� | �������� |
|-----|----------|
| 0x00 | ? ������� ��������� ������� |
| 0x01 | ? ������ CRC |
| 0x02 | ? ����������� ��� ������� |
| 0x03 | ? ����������� ��� ������� |
| 0x04 | ? �������� ����� ������ |
| 0x05 | ? ������ ���������� |
| 0x06 | ? ������� |
| 0xFF | ? ����������� ������ |

## ��������� ������ � DataForm

### ������� ���������� ����������
- `bool SendControllerCommand_StartProgram()`
- `bool SendControllerCommand_StopProgram()`
- `bool SendControllerCommand_PauseProgram()`
- `bool SendControllerCommand_ResumeProgram()`

### ������� ���������� ������������
- `bool SendControllerCommand_SetRelay(uint8_t relayNum, bool state)`
- `bool SendControllerCommand_SetDefrost(bool enable)`

### ������� ������
- `bool SendControllerCommand_GetStatus(ControllerResponse& response)`
- `bool SendControllerCommand_GetSensors(ControllerResponse& response)`
- `bool SendControllerCommand_GetVersion(ControllerResponse& response)`

### ������������� ������
- `bool SendControllerCommandWithResponse(uint8_t cmdType, uint8_t cmdCode, const uint8_t* data, uint8_t dataLen, ControllerResponse& response)`
- `bool SendControllerCommandNoResponse(uint8_t cmdType, uint8_t cmdCode, const uint8_t* data, uint8_t dataLen)`

## �����������

��� �������� ���������� � ���� `log.txt`:
```
25.10.25 12:30:45 : Command created: Type=0x01, Code=0x01, DataLen=0, CRC=0x1234
25.10.25 12:30:45 : Command sent successfully: 5 bytes
25.10.25 12:30:45 : Response received: Type=0x01, Code=0x01, Status=0x00, DataLen=0
25.10.25 12:30:45 : Command completed successfully: OK: ������� ��������� �������
```

## ��������� ������

```cpp
bool success = SendControllerCommand_StartProgram();

if (!success) {
    // ������ ��� ������������
    // ��������� ���������� � Label_Data
    MessageBox::Show("�� ������� ��������� �������", "������");
}
```

## ������ ���������

1. **������� �� ���������**: 5000 �� (5 ������)
2. **������������ ����� ������**: 59 ����
3. **����� ����� ������**: 64 �����
4. **CRC16**: ������������� ����������� (ModBus)
5. **������������������**: ������ ��������� ��� ������ �� GUI ������

## �������������� ������������

��������� ����������� � ������� ����������� �������� ��������� � �����:
**CONTROLLER_PROTOCOL_USAGE.md**

## ���������

��� ������������� ������� ���������:
1. ���������� � ������������ (`clientSocket != INVALID_SOCKET`)
2. ���� � ����� `log.txt`
3. ��������� � `Label_Data` �� �����

---

**���� ��������**: 25 ������� 2025  
**������**: 1.0.0  
**������**: ProjectServerW - Defrost Control System

