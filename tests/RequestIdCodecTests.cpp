#include "../Commands.h"
#include <cassert>
#include <cstdio>
int main() {
    const uint64_t id = 0xfedcba9876543210ULL;
    uint8_t wire[128] = {}, plain[128] = {}; uint64_t decoded = 0; size_t plainSize = 0;
    Command cmd = CreateRequestCommand(CmdRequest::GET_PROCESS_LOG); cmd.dataLength = 8;
    memcpy(cmd.data, &id, 8);
    assert(BuildRequestBuffer(cmd, id, wire, sizeof(wire)) == 21);
    assert(wire[0] == 0x83 && wire[1] == 15 && wire[2] == 16);
    assert(!memcmp(wire + 3, &id, 8) && !memcmp(wire + 11, &id, 8));
    assert(BuildRequestBuffer(cmd, 0, wire, sizeof(wire)) == 0);
    wire[0] = 0x83; wire[1] = 15; wire[2] = 0; wire[3] = 121;
    memcpy(wire + 4, &id, 8);
    for (unsigned i = 0; i < 113; ++i) wire[12+i] = static_cast<uint8_t>(i);
    auto crc = CalculateCommandCRC(wire, 125); memcpy(wire + 125, &crc, 2);
    assert(UnwrapResponseBuffer(wire, 127, decoded, plain, plainSize));
    assert(decoded == id && plainSize == 119 && plain[0] == 3 && plain[3] == 113);
    assert(!memcmp(plain + 4, wire + 12, 113));
    CommandResponse response{}; assert(ParseResponseBuffer(plain, plainSize, response));
    wire[4] ^= 1; assert(!UnwrapResponseBuffer(wire,127,decoded,plain,plainSize)); wire[4] ^= 1;
    assert(!UnwrapResponseBuffer(wire,126,decoded,plain,plainSize));
    wire[0] = 3; assert(!UnwrapResponseBuffer(wire,127,decoded,plain,plainSize));
    puts("PASS request ID codec: command, maximum process frame, CRC, truncation, mandatory envelope");
}
