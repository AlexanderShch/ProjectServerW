#pragma once

// Подключается после заголовка, объявляющего GlobalLogger.
namespace ProjectServerW {
    inline void LogReceivedPacket(System::String^ reason, System::String^ peer,
        const uint8_t* bytes, int count, bool hasResponseHeader, bool discarded = false)
    {
        using namespace System;
        String^ message = String::Format("Warning: {0}: {1}; peer={2}, bytes={3}",
            discarded ? "Dropping received data" : "Receive buffer", reason, peer, count);
        if (hasResponseHeader && count >= 4) {
            const bool tagged = (bytes[0] & 0x80u) != 0;
            message += String::Format("; Type=0x{0:X2}, Code=0x{1:X2}, Status=0x{2:X2}, DataLen={3}",
                static_cast<unsigned>(bytes[0] & 0x7fu), bytes[1], bytes[2], bytes[3]);
            if (tagged && bytes[3] >= 8 && count >= 12) {
                uint64_t requestId = 0;
                for (int i = 0; i < 8; ++i)
                    requestId |= static_cast<uint64_t>(bytes[4 + i]) << (8 * i);
                message += String::Format(", requestId={0:X16}", requestId);
            }
        } else if (count > 0) {
            auto prefix = gcnew cli::array<Byte>(Math::Min(count, 8));
            for (int i = 0; i < prefix->Length; ++i) prefix[i] = bytes[i];
            message += "; prefix=" + BitConverter::ToString(prefix);
        }
        GlobalLogger::LogMessage(message);
    }

    inline void LogDiscardedPacket(System::String^ reason, System::String^ peer,
        const uint8_t* bytes, int count, bool hasResponseHeader)
    {
        LogReceivedPacket(reason, peer, bytes, count, hasResponseHeader, true);
    }

    inline void LogDiscardedPacket(System::String^ reason, System::String^ peer,
        cli::array<System::Byte>^ packet)
    {
        if (packet == nullptr || packet->Length == 0) {
            LogDiscardedPacket(reason, peer, nullptr, 0, false);
            return;
        }
        pin_ptr<System::Byte> bytes = &packet[0];
        LogDiscardedPacket(reason, peer, bytes, packet->Length, true);
    }
}
