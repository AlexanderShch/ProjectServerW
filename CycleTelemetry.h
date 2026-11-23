#pragma once
#include <stdint.h>

enum class CycleTelemetry { Working, Shutdown, IdleGateOpen, Unconfirmed };

inline CycleTelemetry ClassifyCycleTelemetry(bool ioActive, uint16_t inputs, bool working, bool shutdown)
{
    if (shutdown) return CycleTelemetry::Shutdown;
    if (!ioActive) return CycleTelemetry::Unconfirmed;
    if (working) return CycleTelemetry::Working;
    const uint16_t lowerLimit = 1u << 12;
    const uint16_t upperLimit = 1u << 13;
    if ((inputs & (upperLimit | lowerLimit)) == upperLimit)
        return CycleTelemetry::IdleGateOpen;
    return CycleTelemetry::Unconfirmed;
}
