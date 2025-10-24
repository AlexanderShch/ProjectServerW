#pragma once
#include <cstdint>
#include <cstring>

// ����������� ����� ������
struct CmdType {
    static const uint8_t PROG_CONTROL = 0x01;      // ������� ���������� ���������� (�����, ���� � �.�.)
    static const uint8_t CONFIGURATION = 0x02;     // ������� ������������
    static const uint8_t REQUEST = 0x03;           // ������� ������� ������
    static const uint8_t DEVICE_CONTROL = 0x04;    // ������� ���������� ������������
};

// ���� ������ ���������� (��� PROG_CONTROL)
struct CmdProgControl {
    static const uint8_t START = 0x01;         // ������ ������
    static const uint8_t STOP = 0x02;          // ��������� ������
    static const uint8_t PAUSE = 0x03;         // �����
    static const uint8_t RESUME = 0x04;        // ������������� ����� �����
    static const uint8_t RESET = 0x05;         // ����� ����������
};

// ���� ������ ������������ (��� CONFIGURATION)
struct CmdConfig {
    static const uint8_t SET_TEMPERATURE = 0x01;   // ���������� �����������
    static const uint8_t SET_INTERVAL = 0x02;      // ���������� �������� ���������
    static const uint8_t SET_MODE = 0x03;          // ���������� ����� ������
};

// ���� ������ ������� (��� REQUEST)
struct CmdRequest {
    static const uint8_t GET_STATUS = 0x01;        // ��������� ������
    static const uint8_t GET_VERSION = 0x02;       // ��������� ������ ��������
    static const uint8_t GET_DATA = 0x03;          // ��������� ������
};

// ������������ ������ �������
const size_t MAX_COMMAND_SIZE = 64;

// ��������� �������
struct Command {
    uint8_t commandType;           // ��� �������
    uint8_t commandCode;           // ��� �������
    uint8_t dataLength;            // ����� ������ ����������
    uint8_t data[MAX_COMMAND_SIZE - 5]; // ������ ���������� (����. 59 ����)
    uint16_t crc;                  // CRC16 ��� �������� �����������

    Command() : commandType(0), commandCode(0), dataLength(0), crc(0) {
        memset(data, 0, sizeof(data));
    }
};

// ������� ��� ���������� CRC16 �������
uint16_t CalculateCommandCRC(const uint8_t* buffer, size_t length);

// ������� ��� �������� ������ ������� � CRC
size_t BuildCommandBuffer(const Command& cmd, uint8_t* buffer, size_t bufferSize);

// ������� ��� ��������� ���������� ����� �������
const char* GetCommandName(const Command& cmd);

// ��������������� ������� ��� �������� ������

// ������� ������� ���������� ��� ����������
inline Command CreateControlCommand(uint8_t commandCode) {
    Command cmd;
    cmd.commandType = CmdType::PROG_CONTROL;
    cmd.commandCode = commandCode;
    cmd.dataLength = 0;
    return cmd;
}

// ������� ������� ���������� � ����� ������ ���������
inline Command CreateControlCommandWithParam(uint8_t commandCode, uint8_t param) {
    Command cmd;
    cmd.commandType = CmdType::PROG_CONTROL;
    cmd.commandCode = commandCode;
    cmd.dataLength = 1;
    cmd.data[0] = param;
    return cmd;
}

// ������� ������� ������������ � ������������� ���������� (4 �����)
inline Command CreateConfigCommandInt(uint8_t commandCode, int32_t value) {
    Command cmd;
    cmd.commandType = CmdType::CONFIGURATION;
    cmd.commandCode = commandCode;
    cmd.dataLength = 4;
    memcpy(cmd.data, &value, 4);
    return cmd;
}

// ������� ������� ������������ � ���������� ���� float
inline Command CreateConfigCommandFloat(uint8_t commandCode, float value) {
    Command cmd;
    cmd.commandType = CmdType::CONFIGURATION;
    cmd.commandCode = commandCode;
    cmd.dataLength = 4;
    memcpy(cmd.data, &value, 4);
    return cmd;
}

// ������� ������� ������� ��� ����������
inline Command CreateRequestCommand(uint8_t commandCode) {
    Command cmd;
    cmd.commandType = CmdType::REQUEST;
    cmd.commandCode = commandCode;
    cmd.dataLength = 0;
    return cmd;
}

