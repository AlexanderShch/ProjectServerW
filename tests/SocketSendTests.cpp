#include "../SocketSend.h"
#include <iostream>
#include <string>
#include <vector>
#include <stdexcept>

static void Check(const std::vector<int>& chunks, bool expected, const std::string& expectedBytes) {
    const std::string frame = "AA55_ACK_CRC";
    std::string delivered;
    size_t call = 0;
    const bool result = SocketTransport::SendAllBytes(frame.data(), static_cast<int>(frame.size()),
        [&](const char* data, int remaining) {
            if (call >= chunks.size()) throw std::runtime_error("Unexpected extra send");
            const int sent = chunks[call++];
            if (sent > 0 && sent <= remaining) delivered.append(data, sent);
            return sent;
        });
    if (result != expected || delivered != expectedBytes)
        throw std::runtime_error("Frame completion or byte order mismatch");
}

int main() {
    try {
        Check({12}, true, "AA55_ACK_CRC");
        Check({2, 3, 1, 6}, true, "AA55_ACK_CRC");
        Check({1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1}, true, "AA55_ACK_CRC");
        Check({2, SOCKET_ERROR}, false, "AA");
        Check({2, 0}, false, "AA");
        Check({SOCKET_ERROR}, false, "");
        Check({13}, false, "");
        std::cout << "PASS 7 full-frame/partial-send/failure checks\n";
        return 0;
    }
    catch (const std::exception& ex) { std::cerr << ex.what() << '\n'; return 1; }
}
