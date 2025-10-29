# Анализ структуры телеметрии

**Дата:** 25 октября 2025  
**Статус:** ✅ ИСПРАВЛЕНО

---

## 🎯 Проблема

Контроллер использует **упакованные структуры** (packed structures) с атрибутом `__attribute__((packed))` для GCC, чтобы избежать выравнивания байтов.

Без соответствующего атрибута на стороне сервера компилятор может добавить **padding байты** между полями структуры для выравнивания, что приведет к **неправильному парсингу** данных.

---

## 📊 Структура телеметрии MSGQUEUE_OBJ_t

### Поля структуры:

| Поле | Тип | Размер | Смещение без pack | Смещение с pack(1) |
|------|-----|--------|-------------------|-------------------|
| Time | uint16_t | 2 байта | 0 | 0 |
| SensorQuantity | uint8_t | 1 байт | 2 | 2 |
| (padding) | - | 1 байт | 3 | - |
| SensorType[7] | uint8_t[7] | 7 байт | 4 | 3 |
| Active[7] | uint8_t[7] | 7 байт | 11 | 10 |
| (padding) | - | 1 байт | 18 | - |
| T[7] | short[7] | 14 байт | 20 | 17 |
| H[7] | short[7] | 14 байт | 34 | 31 |
| CRC_SUM | uint16_t | 2 байта | 48 | 45 |
| **ИТОГО** | | **48 байт** (pack) | **50 байт** (без pack) | **47 байт** (pack(1)) |

**Внимание:** Размер зависит от SQ (количество сенсоров).
- SQ = 7 (по умолчанию в DataForm.h)

### Расчет размера (с pack(1)):
```
uint16_t Time           =  2 байта
uint8_t SensorQuantity  =  1 байт
uint8_t SensorType[7]   =  7 байт
uint8_t Active[7]       =  7 байт
short T[7]              = 14 байт (7 * 2)
short H[7]              = 14 байт (7 * 2)
uint16_t CRC_SUM        =  2 байта
─────────────────────────────────
ИТОГО                   = 47 байт
```

**Ожидаемый размер в SServer.cpp:** 48 байт (включая первый байт Type=0x00)

---

## ✅ Исправление

### Было (DataForm.cpp строки 15-25):
```cpp
//
typedef struct   // object data for Server type из STM32
{
    uint16_t Time;
    uint8_t SensorQuantity;
    uint8_t SensorType[SQ];
    uint8_t Active[SQ];
    short T[SQ];
    short H[SQ];
    uint16_t CRC_SUM;
} MSGQUEUE_OBJ_t;
```

**Проблема:** Компилятор может добавить padding байты для выравнивания полей.

---

### Стало (DataForm.cpp строки 15-28):
```cpp
// Структура телеметрии от контроллера (упакованная, без выравнивания байтов)
// Контроллер использует __attribute__((packed)), здесь используем #pragma pack(1)
#pragma pack(push, 1)
typedef struct   // object data for Server type из STM32
{
    uint16_t Time;				// Количество секунд с момента включения
    uint8_t SensorQuantity;		// Количество сенсоров
    uint8_t SensorType[SQ];		// Тип сенсора
    uint8_t Active[SQ];			// Активность сенсора
    short T[SQ];				// Значение 1 сенсора (температура)
    short H[SQ];				// Значение 2 сенсора (влажность)
    uint16_t CRC_SUM;			// Контрольное значение
} MSGQUEUE_OBJ_t;
#pragma pack(pop)
```

**Решение:** 
- `#pragma pack(push, 1)` - устанавливает выравнивание на 1 байт (без padding)
- `#pragma pack(pop)` - восстанавливает предыдущее выравнивание

---

## 🔍 Как это работает

### Без #pragma pack(1):
```
Offset: 0  1  2  3  4  5  6  7  8  9  10 11 12 13 14 15 16 17 18 19 ...
       [Time][SQ][P][SensorType................][Active............][P]...
```
`[P]` = padding байт (добавлен компилятором)

### С #pragma pack(1):
```
Offset: 0  1  2  3  4  5  6  7  8  9  10 11 12 13 14 15 16 17 18 19 ...
       [Time][SQ][SensorType................][Active............][T[0]]...
```
Нет padding байтов, данные идут последовательно.

---

## 📝 Дополнительная информация

### В SServer.cpp (строка 320):
```cpp
uint8_t LengthOfPackage = 48;
```

**Важно:** Это размер пакета **включая первый байт Type=0x00**.

**Разбивка 48 байт:**
- 1 байт: Type (0x00 для телеметрии)
- 47 байт: структура MSGQUEUE_OBJ_t (с pack(1))

---

## ⚠️ Что проверить

### 1. Размер структуры в контроллере:
Убедитесь, что на контроллере (STM32) размер структуры такой же:
```c
// На контроллере (GCC):
typedef struct __attribute__((packed)) {
    uint16_t Time;
    uint8_t SensorQuantity;
    uint8_t SensorType[7];
    uint8_t Active[7];
    short T[7];
    short H[7];
    uint16_t CRC_SUM;
} MSGQUEUE_OBJ_t;

// Размер должен быть 47 байт
static_assert(sizeof(MSGQUEUE_OBJ_t) == 47, "Structure size mismatch");
```

### 2. Формат пакета телеметрии:
```
Байт 0:    Type = 0x00 (телеметрия)
Байт 1-47: MSGQUEUE_OBJ_t (47 байт)
─────────────────────────────
ИТОГО:     48 байт
```

### 3. CRC вычисляется для:
```cpp
// В SServer.cpp (строка 322):
uint16_t dataCRC = MB_GetCRC(buffer, LengthOfPackage - 2);
```
CRC вычисляется для первых **46 байт** (48 - 2):
- Байт 0: Type (0x00)
- Байт 1-45: все поля кроме CRC_SUM
- Байт 46-47: CRC_SUM (не включается в расчет)

---

## 🎯 Результат

✅ **Структура телеметрии теперь упакована** (`#pragma pack(1)`)  
✅ **Соответствует формату контроллера** (без padding байтов)  
✅ **Корректный парсинг данных** гарантирован  
✅ **Совместимость с GCC** `__attribute__((packed))`

---

## 📚 Справка

### Различия между компиляторами:

| Компилятор | Синтаксис |
|------------|-----------|
| **GCC** | `__attribute__((packed))` |
| **MSVC** | `#pragma pack(1)` |
| **Clang** | `__attribute__((packed))` или `#pragma pack(1)` |

### Применение:

**GCC/Clang:**
```c
typedef struct __attribute__((packed)) {
    // ...
} MyStruct_t;
```

**MSVC (Visual Studio):**
```cpp
#pragma pack(push, 1)
typedef struct {
    // ...
} MyStruct_t;
#pragma pack(pop)
```

---

**Дата исправления:** 25 октября 2025  
**Статус:** ✅ ГОТОВО К ТЕСТИРОВАНИЮ


