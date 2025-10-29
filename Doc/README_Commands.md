# ������� ������ ���������� �����������

## �����

� ������ ��������� ����������������� ������� ������ ��� ���������� ������� �����������. ��� ������� ����� ������ ������ � ������������� �������� ����������� ������ CRC16.

## �������� �����

- **Commands.h** - ����������� ����� ������, ����� � ��������������� �������
- **Commands.cpp** - ���������� ������� ��� ������ � ���������
- **CommandsExamples.txt** - ��������� ������� �������������

## ������� �����

### �������� ������� �����
```cpp
void ProjectServerW::DataForm::SendStartCommand() {
    Command cmd = CreateControlCommand(ControlCommand::START);
    SendCommand(cmd, "START");
}
```

### �������� ������� ����
```cpp
void ProjectServerW::DataForm::SendStopCommand() {
    Command cmd = CreateControlCommand(ControlCommand::STOP);
    SendCommand(cmd, "STOP");
}
```

### ��������� �����������
```cpp
void SetTemperature(float temp) {
    Command cmd = CreateConfigCommandFloat(ConfigCommand::SET_TEMPERATURE, temp);
    SendCommand(cmd, "SET_TEMPERATURE");
}
```

## ��������� �������

```
[��� �������][��� �������][����� ������][���������...][CRC16]
    1 ����       1 ����        1 ����      0-59 ����    2 �����
```

## ���� ������

| ��� | ��� | �������� |
|-----|-----|----------|
| CONTROL | 0x01 | ������� ���������� (�����, ����, �����) |
| CONFIGURATION | 0x02 | ������� ������������ (��������� ����������) |
| REQUEST | 0x03 | ������� ������ �� ���������� |
| RESPONSE | 0x04 | ������ �� ���������� |

## ������� ���������� (CONTROL)

| ������� | ��� | �������� |
|---------|-----|----------|
| START | 0x01 | ������ ������ ���������� |
| STOP | 0x02 | ��������� ������ |
| PAUSE | 0x03 | ������������ |
| RESUME | 0x04 | ������������� |
| RESET | 0x05 | ����� ���������� |

## ������� ������������ (CONFIGURATION)

| ������� | ��� | �������� | �������� |
|---------|-----|----------|----------|
| SET_TEMPERATURE | 0x01 | ���������� ����������� | float (4 �����) |
| SET_INTERVAL | 0x02 | �������� ��������� | int (4 �����) |
| SET_MODE | 0x03 | ����� ������ | uint8_t (1 ����) |

## ������� ������� (REQUEST)

| ������� | ��� | �������� |
|---------|-----|----------|
| GET_STATUS | 0x01 | ��������� ������ |
| GET_VERSION | 0x02 | ��������� ������ �� |
| GET_CONFIG | 0x03 | ��������� ������������ |

## ��� �������� ����� �������

### 1. �������� ��������� � Commands.h
```cpp
namespace ControlCommand {
    // ... ������������ �������
    const uint8_t MY_NEW_COMMAND = 0x06;  // ����� �������
}
```

### 2. ������� ����� � DataForm.h
```cpp
void SendMyNewCommand(int parameter);
```

### 3. ����������� � DataForm.cpp
```cpp
void ProjectServerW::DataForm::SendMyNewCommand(int parameter) {
    Command cmd = CreateConfigCommandInt(ControlCommand::MY_NEW_COMMAND, parameter);
    SendCommand(cmd, "MY_NEW_COMMAND");
}
```

### 4. ������� �� ����������� ������
```cpp
private: System::Void button_Click(System::Object^ sender, System::EventArgs^ e) {
    SendMyNewCommand(100);
}
```

## ������������ ����� �������

? **������ ������** - ��� ������� ����� ���������� ���������  
? **������ ������** - �������������� ���������� CRC16  
? **����������������** - ������ ���������� ����� ������  
? **�������������** - ����� ��������� ����� �������  
? **�����������** - �������������� ����������� ���� ��������  
? **��������� ������** - ���������������� ��������� ������ ��������  

## �������������� ����������

��������� ������� � ���������� ��. � ����� **CommandsExamples.txt**

