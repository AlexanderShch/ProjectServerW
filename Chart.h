#pragma once
#include <cstdint>          // Для uint16_t
#include "SServer.h" // 

#define SQ 6				// sensors quantity for measures (0-4) + sets of T (5, 6) + MB_IO



// Объявление в классе формы
public ref class ChartForm : public System::Windows::Forms::Form {
private:
    System::Windows::Forms::DataGridView^ dataGridView;
    System::Windows::Forms::BindingSource^ bindingSource;
    System::Data::DataTable^ dataTable;
public:
    value struct MyData       // object data for Server type
    {
        uint16_t Time;				// Количество секунд с момента включения
        uint8_t SensorQuantity;		// Количество сенсоров
        cli::array<System::Byte>^ SensorType;     // Тип сенсора
        cli::array<System::Byte>^ Active;         // Активность сенсора
        cli::array<System::UInt16>^ T;            // Значение 1 сенсора (температура)
        cli::array<System::UInt16>^ H;            // Значение 2 сенсора (влажность)
        uint16_t CRC_SUM;			    // Контрольное значение

        // Конструктор
        static MyData Create() {
            MyData data;
            data.SensorType = gcnew cli::array<System::Byte>(SQ);
            data.Active = gcnew cli::array<System::Byte>(SQ);
            data.T = gcnew cli::array<System::UInt16>(SQ);
            data.H = gcnew cli::array<System::UInt16>(SQ);
            return data;
        }
    };

public:
    ChartForm(void) {
        //InitializeComponent();
        InitializeDataTable();
    }
    static void ParseBuffer(const char* buffer, size_t size);

private:
    void InitializeDataTable();
    //static void AddDataToTable(DataTable^ table);
};