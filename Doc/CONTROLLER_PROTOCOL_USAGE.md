# ����������� �� ������������� ��������� ������ �����������

## �����

� ������ ProjectServerW ��������� ������ ��������� ��������� ������ ��� ���������� ������������ STM32F429. �������� ���������� ��������������� ����� ����� TCP/IP ������, ��� ��������� ������������ ���������� ������� � �������� ������.

## �����: ��������������� �����

**STM32F429 ����� ������������:**
- ? ������� COM-���� (�������� ����� �������)
- ? ���������� ������ ����� COM-���� (���������� ������)

**�� �����** ������������� ������������� ����� �� ����� �������� ������!

## ��������� ���������

### ������� (�� ������� � �����������)
```
????????????????????????????????????????????????????????
?  ���� 0  ?  ���� 1  ?  ���� 2   ? ���� 3-N ? N+1, N+2?
????????????????????????????????????????????????????????
?   Type   ?   Code   ?  DataLen  ?   Data   ?  CRC16  ?
? (1 ����) ? (1 ����) ?  (1 ����) ?(0-59 �.) ?(2 �����)?
????????????????????????????????????????????????????????
```

### ����� (�� ����������� � �������)
```
???????????????????????????????????????????????????????????????????
?  ���� 0  ?  ���� 1  ?  ���� 2  ?  ���� 3   ? ���� 4-N ? N+1, N+2?
???????????????????????????????????????????????????????????????????
?   Type   ?   Code   ?  Status  ?  DataLen  ?   Data   ?  CRC16  ?
? (1 ����) ? (1 ����) ? (1 ����) ?  (1 ����) ?(0-59 �.) ?(2 �����)?
???????????????????????????????????????????????????????????????????
```

## ���� ������

### 1. ������� ���������� ���������� (PROG_CONTROL = 0x01)
- `CMD_START_PROGRAM` (0x01) - ��������� ���������
- `CMD_STOP_PROGRAM` (0x02) - ���������� ���������
- `CMD_PAUSE_PROGRAM` (0x03) - ������������� ���������
- `CMD_RESUME_PROGRAM` (0x04) - ����������� ���������

### 2. ������� ���������� ������������ (DEVICE_CONTROL = 0x02)
- `CMD_SET_RELAY` (0x01) - ���������� ����
- `CMD_SET_DEFROST` (0x02) - ���������� ���������

### 3. ������� ������������ (CONFIGURATION = 0x03)
- `CMD_SET_PARAM` (0x01) - ���������� ��������
- `CMD_GET_PARAM` (0x02) - �������� ��������

### 4. ������� ������ (REQUEST = 0x04)
- `CMD_GET_STATUS` (0x01) - �������� ������
- `CMD_GET_SENSORS` (0x02) - �������� ������ ��������
- `CMD_GET_VERSION` (0x03) - �������� ������ ��

## ������� ����������

| ���  | �������� | �������� |
|------|----------|----------|
| 0x00 | CMD_STATUS_OK | ? ������� ��������� ������� |
| 0x01 | CMD_STATUS_CRC_ERROR | ? ������ ����������� ����� |
| 0x02 | CMD_STATUS_INVALID_TYPE | ? ����������� ��� ������� |
| 0x03 | CMD_STATUS_INVALID_CODE | ? ����������� ��� ������� |
| 0x04 | CMD_STATUS_INVALID_LENGTH | ? �������� ����� ������ |
| 0x05 | CMD_STATUS_EXECUTION_ERROR | ? ������ ���������� ������� |
| 0x06 | CMD_STATUS_TIMEOUT | ? ������� ���������� |
| 0xFF | CMD_STATUS_UNKNOWN_ERROR | ? ����������� ������ |

## ������� ������������� � DataForm

### ������ 1: ������ ��������� �����������

```cpp
// � ����������� ������ START
private: System::Void buttonStartController_Click(System::Object^ sender, System::EventArgs^ e) 
{
    // ���������� ������� ������� ���������
    bool success = SendControllerCommand_StartProgram();
    
    if (success) {
        // ������� ��������� �������
        // ��������� ������ ��� �������� ������ �������
        MessageBox::Show("��������� ����������� ��������", "�����", 
                        MessageBoxButtons::OK, MessageBoxIcon::Information);
    } else {
        // ��������� ������
        MessageBox::Show("�� ������� ��������� ��������� �����������", "������", 
                        MessageBoxButtons::OK, MessageBoxIcon::Error);
    }
}
```

### ������ 2: ��������� ��������� �����������

```cpp
// � ����������� ������ STOP
private: System::Void buttonStopController_Click(System::Object^ sender, System::EventArgs^ e) 
{
    // ���������� ������� ��������� ���������
    bool success = SendControllerCommand_StopProgram();
    
    if (success) {
        MessageBox::Show("��������� ����������� �����������", "�����", 
                        MessageBoxButtons::OK, MessageBoxIcon::Information);
    }
}
```

### ������ 3: ���������� ����

```cpp
// ��������� ���� �1
private: System::Void buttonRelay1On_Click(System::Object^ sender, System::EventArgs^ e) 
{
    uint8_t relayNum = 1;
    bool state = true; // ��������
    
    bool success = SendControllerCommand_SetRelay(relayNum, state);
    
    if (success) {
        Label_Data->Text = "���� 1 ��������";
    }
}

// ���������� ���� �1
private: System::Void buttonRelay1Off_Click(System::Object^ sender, System::EventArgs^ e) 
{
    uint8_t relayNum = 1;
    bool state = false; // ���������
    
    bool success = SendControllerCommand_SetRelay(relayNum, state);
    
    if (success) {
        Label_Data->Text = "���� 1 ���������";
    }
}
```

### ������ 4: ���������� ���������

```cpp
// ��������� ��������
private: System::Void buttonDefrostOn_Click(System::Object^ sender, System::EventArgs^ e) 
{
    bool success = SendControllerCommand_SetDefrost(true);
    
    if (success) {
        Label_Data->Text = "������� �������";
        // ����� �������� ���� ����������
        labelDefrostStatus->BackColor = System::Drawing::Color::Lime;
    }
}

// ���������� ��������
private: System::Void buttonDefrostOff_Click(System::Object^ sender, System::EventArgs^ e) 
{
    bool success = SendControllerCommand_SetDefrost(false);
    
    if (success) {
        Label_Data->Text = "������� ��������";
        labelDefrostStatus->BackColor = System::Drawing::Color::Gray;
    }
}
```

### ������ 5: ������ ������� �����������

```cpp
// ��������� ������� �����������
private: System::Void buttonGetStatus_Click(System::Object^ sender, System::EventArgs^ e) 
{
    ControllerResponse response;
    bool success = SendControllerCommand_GetStatus(response);
    
    if (success) {
        // ������������ ���������� ������ �������
        // ������ ������ ������� �� ��������� �����������
        
        if (response.dataLen > 0) {
            // ������: �����������, ��� ������ ���� - ��� ��������� ���������
            uint8_t programState = response.data[0];
            
            System::String^ statusText = "������: ";
            switch (programState) {
                case 0: statusText += "�����������"; break;
                case 1: statusText += "��������"; break;
                case 2: statusText += "��������������"; break;
                default: statusText += "����������"; break;
            }
            
            Label_Data->Text = statusText;
        }
    }
}
```

### ������ 6: ��������� ������ ��������

```cpp
// ��������� ������ ��������
private: System::Void buttonGetSensors_Click(System::Object^ sender, System::EventArgs^ e) 
{
    ControllerResponse response;
    bool success = SendControllerCommand_GetSensors(response);
    
    if (success && response.dataLen > 0) {
        // ������ ��������� ������ ��������
        // �����������, ��� ������ - ��� ������ ���������� (float, 4 ����� ������)
        
        int numSensors = response.dataLen / sizeof(float);
        
        for (int i = 0; i < numSensors && i < 6; i++) {
            float temperature;
            memcpy(&temperature, &response.data[i * sizeof(float)], sizeof(float));
            
            // ��������� ��������������� ���� �� �����
            switch (i) {
                case 0: T_def_left->Text = temperature.ToString("F1") + "�C"; break;
                case 1: T_def_right->Text = temperature.ToString("F1") + "�C"; break;
                case 2: T_def_center->Text = temperature.ToString("F1") + "�C"; break;
                case 3: T_product_left->Text = temperature.ToString("F1") + "�C"; break;
                case 4: T_product_right->Text = temperature.ToString("F1") + "�C"; break;
            }
        }
    }
}
```

### ������ 7: ��������� ������ �� �����������

```cpp
// ��������� ������ ��
private: System::Void buttonGetVersion_Click(System::Object^ sender, System::EventArgs^ e) 
{
    ControllerResponse response;
    bool success = SendControllerCommand_GetVersion(response);
    
    if (success) {
        // ������ ��� ���������� � Label_Data ������ �������
        // ����� ������������� ���������� ������
        
        if (response.dataLen > 0) {
            System::String^ version = System::Text::Encoding::ASCII->GetString(
                response.data, 0, response.dataLen
            );
            
            // �������� � MessageBox
            MessageBox::Show("������ �� �����������:\n" + version, 
                           "���������� � ������", 
                           MessageBoxButtons::OK, 
                           MessageBoxIcon::Information);
        }
    }
}
```

### ������ 8: ������������� �������� �������

```cpp
// �������� ������������ �������
private: System::Void SendCustomCommand() 
{
    // ������� ������ �������
    uint8_t data[4];
    data[0] = 0x01; // �������� 1
    data[1] = 0x02; // �������� 2
    data[2] = 0x03; // �������� 3
    data[3] = 0x04; // �������� 4
    
    ControllerResponse response;
    
    // ���������� ������� ���� CONFIGURATION, ��� CMD_SET_PARAM
    bool success = SendControllerCommandWithResponse(
        CONFIGURATION,      // ��� �������
        CMD_SET_PARAM,      // ��� �������
        data,               // ������
        4,                  // ����� ������
        response            // �����
    );
    
    if (success) {
        Label_Data->Text = "��������� ����������� �������";
        
        // ��������� ������, ���� ���� ������
        if (response.dataLen > 0) {
            // ��������� ������ ������...
        }
    } else {
        Label_Data->Text = "������ ��������� ����������";
    }
}
```

### ������ 9: ������������� ����� �����������

```cpp
// � ������ DataForm �������� ������ ��� �������������� ������
private: System::Windows::Forms::Timer^ pollTimer;

// � ������������ DataForm
DataForm(void)
{
    InitializeComponent();
    // ... ������ ������������� ...
    
    // ������������� ������� ������
    pollTimer = gcnew System::Windows::Forms::Timer();
    pollTimer->Interval = 5000; // ����� ������ 5 ������
    pollTimer->Tick += gcnew EventHandler(this, &DataForm::PollController);
    pollTimer->Start();
}

// ���������� �������
private: System::Void PollController(Object^ sender, EventArgs^ e) 
{
    // ����������� ������ ��������
    ControllerResponse response;
    bool success = SendControllerCommand_GetSensors(response);
    
    if (success && response.dataLen > 0) {
        // ��������� ����������� ������ �� �����
        // ... ��������� ������ �������� ...
    } else {
        // ����� �� ������, ����� �������� ��������������
        // �� �� ��������� ������ ����������
    }
}
```

## ��������� ������

��� ������� �������� ������ ���������� `bool`:
- `true` - ������� ��������� �������
- `false` - ��������� ������

��� ������:
1. ��������� �� ������ ������������ � ��� ����� `GlobalLogger::LogMessage()`
2. ��������� �� ������ ������������ � `Label_Data->Text`
3. ������� ���������� `false`

������ ��������� ������:

```cpp
bool success = SendControllerCommand_StartProgram();

if (!success) {
    // ��������� ��������� ������� ������:
    
    // 1. ��� ����������� � �����������
    if (clientSocket == INVALID_SOCKET) {
        MessageBox::Show("��� ����������� � �����������", "������");
        return;
    }
    
    // 2. ������� ������
    MessageBox::Show("���������� �� ������� �������.\n"
                    "��������� ����������.", "�������");
    
    // 3. ������ �� ������� �����������
    // ����� ������ ��� ��������� � Label_Data
}
```

## �����������

��� �������� � ���������� ������������� ����������:

```cpp
GlobalLogger::LogMessage("Command created: Type=0x01, Code=0x01, DataLen=0, CRC=0x1234");
GlobalLogger::LogMessage("Command sent successfully: 5 bytes");
GlobalLogger::LogMessage("Response received: Type=0x01, Code=0x01, Status=0x00, DataLen=0");
GlobalLogger::LogMessage("Command completed successfully: OK: ������� ��������� �������");
```

��� ����������� � ���� `log.txt` � ���������� ����������.

## ������ ���������

1. **��������**: �� ��������� ������� �������� ������ ���������� 5000 �� (5 ������). ����� �������� ��� �������������.

2. **CRC16**: ����������� ����� ����������� ������������� �� ��������� ModBus CRC16.

3. **�����������������**: ����� ���������� ������� � �������� ������ ������������ ��� ��������� ������������� �����.

4. **����� ������**: ������������ ����� ������ � �������/������ - 59 ����. ����� ����� ������ - 64 �����.

5. **������������ �������**: ��� ������ ��������� ��� ������ �� GUI ������, ��� ��� �������� � ������� ������� �����.

## ���������� ����� ������

����� �������� ����� �������:

1. �������� ��������� � `SServer.h`:
```cpp
enum MyNewCommandType : uint8_t {
    CMD_MY_NEW_COMMAND = 0x05
};
```

2. �������� ����� � `DataForm.h`:
```cpp
bool SendControllerCommand_MyNewCommand(uint8_t param1, uint8_t param2);
```

3. ���������� ����� � `DataForm.cpp`:
```cpp
bool ProjectServerW::DataForm::SendControllerCommand_MyNewCommand(uint8_t param1, uint8_t param2)
{
    uint8_t data[2];
    data[0] = param1;
    data[1] = param2;
    
    ControllerResponse response;
    bool result = SendControllerCommandWithResponse(
        PROG_CONTROL,           // ��� �������
        CMD_MY_NEW_COMMAND,     // ��� �������
        data, 2,                // ������
        response                // �����
    );
    
    if (result) {
        Label_Data->Text = "��� ������� ���������";
    }
    
    return result;
}
```

## ����������

�������� ������ ��������� ���������� � ����� � �������������. ��� ������� ���� ����� ������������ ��� ���� ���������� ������. �������� ������������ ��������������� �����, �������������� �������� CRC, ��������� ������ � ����������� ���� ��������.

