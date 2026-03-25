# Протокол обмена: сервер ProjectServerW <-> контроллер Defrost

Кодировка документа: UTF-8.

## Новая команда REQUEST: GET_ALARM_FLAGS

- Тип: `0x03` (REQUEST)
- Код: `0x09` (`GET_ALARM_FLAGS`)
- Запрос: без payload (`DataLen = 0`)
- Ответ: `DataLen = 4`
  - `data[0..1]` — `Device_AlarmFlags` (`uint16`, little-endian)
  - `data[2..3]` — `Sensor_AlarmFlags` (`uint16`, little-endian)

## Обновление groupId=6 (GET_DEFROST_GROUP)

- В структуру `DefrostLogGlobalPayload_t` добавлен параметр `debugDisableTargetTStop` (`uint8`).
- Порядок после обновления:
  - `fishColdTarget_C` (`float`)
  - `debugDisableTargetTStop` (`uint8`)
  - `sensorUseInDefrost[6]` (`uint8[6]`)
- Значения параметра:
  - `0` — автостоп по целевой температуре рыбы включён
  - `1` — автостоп отключён (режим отладки)

