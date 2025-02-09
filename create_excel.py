import sys
import struct
from openpyxl import Workbook
from datetime import datetime
import pytz

def parse_buffer(buffer):
    struct_format = "<H B 6B 6B 6H 6H H"
    data = struct.unpack(struct_format, buffer)
    return {
        "Time": data[0],
        "SensorQuantity": data[1],
        "SensorType": data[2:8],
        "Active": data[8:14],
        "T": data[14:20],
        "H": data[20:26],
        "CRC_SUM": data[26]
    }

def create_excel(data, filename):
    wb = Workbook()
    ws = wb.active
    
    headers = ["Created (UTC+3)", "Time", "SensorQuantity"]
    for i in range(6):
        headers.extend([
            f"Sensor{i}_Type",
            f"Sensor{i}_Active",
            f"Sensor{i}_Temp",
            f"Sensor{i}_Humidity"
        ])
    ws.append(headers)

    # Московское время
    tz = pytz.timezone('Europe/Moscow')
    created_time = datetime.now(tz).strftime("%Y-%m-%d %H:%M:%S")
    
    # Формирование строки данных
    row = [created_time, data["Time"], data["SensorQuantity"]]
    for i in range(6):
        row += [
            data["SensorType"][i],
            data["Active"][i],
            data["T"][i],
            data["H"][i]
        ]
    
    ws.append(row)
    wb.save(filename)

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python create_excel.py <hex_buffer>")
        sys.exit(1)
    
    # Преобразование HEX в бинарные данные
    try:
        buffer = bytes.fromhex(sys.argv[1])
    except ValueError:
        print("Invalid HEX format")
        sys.exit(2)
    
    # Парсинг данных
    parsed_data = parse_buffer(buffer)
    
    # Генерация имени файла
    tz = pytz.timezone('Europe/Moscow')
    now = datetime.now(tz)
    filename = now.strftime("sensor_data_%y_%m_%d_%H-%M.xlsx")
    
    # Создание файла
    create_excel(parsed_data, filename)