#pragma once
#include <winsock2.h>

namespace SocketTransport {
    // send() may accept only a prefix. Keep the caller's send gate until the
    // entire frame has been accepted, or report failure instead of a false ACK.
    template<class Sender>
    bool SendAllBytes(const char* buffer, int length, Sender sender) {
        if (buffer == nullptr || length <= 0) return false;
        int offset = 0;
        while (offset < length) {
            const int remaining = length - offset;
            const int sent = sender(buffer + offset, remaining);
            if (sent <= 0 || sent > remaining) return false;
            offset += sent;
        }
        return true;
    }

    inline bool SendAll(SOCKET socket, const char* buffer, int length) {
        return SendAllBytes(buffer, length, [socket](const char* chunk, int size) {
            const int sent = send(socket, chunk, size, 0);
            if (sent == 0) WSASetLastError(WSAECONNRESET);
            return sent;
        });
    }
}
