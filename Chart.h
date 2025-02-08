#pragma once
#include <cstdint>          // ��� uint16_t
#include "SServer.h" // 

#define SQ 6				// sensors quantity for measures (0-4) + sets of T (5, 6) + MB_IO



// ���������� � ������ �����
public ref class ChartForm : public System::Windows::Forms::Form {
private:
    System::Windows::Forms::DataGridView^ dataGridView;
    System::Windows::Forms::BindingSource^ bindingSource;
    System::Data::DataTable^ dataTable;
public:
    value struct MyData       // object data for Server type
    {
        uint16_t Time;				// ���������� ������ � ������� ���������
        uint8_t SensorQuantity;		// ���������� ��������
        cli::array<System::Byte>^ SensorType;     // ��� �������
        cli::array<System::Byte>^ Active;         // ���������� �������
        cli::array<System::UInt16>^ T;            // �������� 1 ������� (�����������)
        cli::array<System::UInt16>^ H;            // �������� 2 ������� (���������)
        uint16_t CRC_SUM;			    // ����������� ��������

        // �����������
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