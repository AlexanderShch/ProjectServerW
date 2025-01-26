#pragma once

namespace ProjectServerW {

	using namespace System;
	using namespace System::ComponentModel;
	using namespace System::Collections;
	using namespace System::Windows::Forms;
	using namespace System::Data;
	using namespace System::Drawing;

	/// <summary>
	/// Сводка для MyForm
	/// </summary>
	public ref class MyForm : public System::Windows::Forms::Form
	{
	public:
		MyForm(void)
		{
			InitializeComponent();
			//
			//TODO: добавьте код конструктора
			//
		}

	protected:
		/// <summary>
		/// Освободить все используемые ресурсы.
		/// </summary>
		~MyForm()
		{
			if (components)
			{
				delete components;
			}
		}
	private: System::Windows::Forms::MenuStrip^ menuStrip1;
	protected:
	private: System::Windows::Forms::ToolStripMenuItem^ выходToolStripMenuItem;
	private: System::Windows::Forms::Button^ button_Listen;

	private: System::Windows::Forms::Label^ label_N_port;
	private: System::Windows::Forms::Label^ SocketState;
	private: System::Windows::Forms::Label^ SocketBind;
	private: System::Windows::Forms::Label^ WSAstartup;





	private:
		/// <summary>
		/// Обязательная переменная конструктора.
		/// </summary>
		System::ComponentModel::Container ^components;

#pragma region Windows Form Designer generated code
		/// <summary>
		/// Требуемый метод для поддержки конструктора — не изменяйте 
		/// содержимое этого метода с помощью редактора кода.
		/// </summary>
		void InitializeComponent(void)
		{
			this->menuStrip1 = (gcnew System::Windows::Forms::MenuStrip());
			this->выходToolStripMenuItem = (gcnew System::Windows::Forms::ToolStripMenuItem());
			this->button_Listen = (gcnew System::Windows::Forms::Button());
			this->label_N_port = (gcnew System::Windows::Forms::Label());
			this->SocketState = (gcnew System::Windows::Forms::Label());
			this->SocketBind = (gcnew System::Windows::Forms::Label());
			this->WSAstartup = (gcnew System::Windows::Forms::Label());
			this->menuStrip1->SuspendLayout();
			this->SuspendLayout();
			// 
			// menuStrip1
			// 
			this->menuStrip1->GripMargin = System::Windows::Forms::Padding(2, 2, 0, 2);
			this->menuStrip1->ImageScalingSize = System::Drawing::Size(24, 24);
			this->menuStrip1->Items->AddRange(gcnew cli::array< System::Windows::Forms::ToolStripItem^  >(1) { this->выходToolStripMenuItem });
			this->menuStrip1->Location = System::Drawing::Point(0, 0);
			this->menuStrip1->Name = L"menuStrip1";
			this->menuStrip1->Size = System::Drawing::Size(786, 33);
			this->menuStrip1->TabIndex = 0;
			this->menuStrip1->Text = L"menuStrip1";
			// 
			// выходToolStripMenuItem
			// 
			this->выходToolStripMenuItem->Name = L"выходToolStripMenuItem";
			this->выходToolStripMenuItem->Size = System::Drawing::Size(80, 32);
			this->выходToolStripMenuItem->Text = L"Выход";
			this->выходToolStripMenuItem->Click += gcnew System::EventHandler(this, &MyForm::выходToolStripMenuItem_Click);
			// 
			// button_Listen
			// 
			this->button_Listen->Font = (gcnew System::Drawing::Font(L"Microsoft Sans Serif", 12, System::Drawing::FontStyle::Regular, System::Drawing::GraphicsUnit::Point,
				static_cast<System::Byte>(204)));
			this->button_Listen->Location = System::Drawing::Point(31, 89);
			this->button_Listen->Name = L"button_Listen";
			this->button_Listen->Size = System::Drawing::Size(204, 53);
			this->button_Listen->TabIndex = 1;
			this->button_Listen->Text = L"Слушать";
			this->button_Listen->UseVisualStyleBackColor = true;
			this->button_Listen->Click += gcnew System::EventHandler(this, &MyForm::button_Listen_Click);
			// 
			// label_N_port
			// 
			this->label_N_port->AutoSize = true;
			this->label_N_port->Font = (gcnew System::Drawing::Font(L"Microsoft Sans Serif", 12, System::Drawing::FontStyle::Regular, System::Drawing::GraphicsUnit::Point,
				static_cast<System::Byte>(204)));
			this->label_N_port->Location = System::Drawing::Point(274, 254);
			this->label_N_port->Name = L"label_N_port";
			this->label_N_port->Size = System::Drawing::Size(110, 29);
			this->label_N_port->TabIndex = 3;
			this->label_N_port->Text = L"№ порта";
			// 
			// SocketState
			// 
			this->SocketState->AutoSize = true;
			this->SocketState->Font = (gcnew System::Drawing::Font(L"Microsoft Sans Serif", 12, System::Drawing::FontStyle::Regular, System::Drawing::GraphicsUnit::Point,
				static_cast<System::Byte>(204)));
			this->SocketState->Location = System::Drawing::Point(274, 148);
			this->SocketState->Name = L"SocketState";
			this->SocketState->Size = System::Drawing::Size(223, 29);
			this->SocketState->TabIndex = 4;
			this->SocketState->Text = L"Состояние сокета";
			// 
			// SocketBind
			// 
			this->SocketBind->AutoSize = true;
			this->SocketBind->Font = (gcnew System::Drawing::Font(L"Microsoft Sans Serif", 12, System::Drawing::FontStyle::Regular, System::Drawing::GraphicsUnit::Point,
				static_cast<System::Byte>(204)));
			this->SocketBind->Location = System::Drawing::Point(274, 200);
			this->SocketBind->Name = L"SocketBind";
			this->SocketBind->Size = System::Drawing::Size(207, 29);
			this->SocketBind->TabIndex = 5;
			this->SocketBind->Text = L"Привязка сокета";
			// 
			// WSAstartup
			// 
			this->WSAstartup->AutoSize = true;
			this->WSAstartup->Font = (gcnew System::Drawing::Font(L"Microsoft Sans Serif", 12, System::Drawing::FontStyle::Regular, System::Drawing::GraphicsUnit::Point,
				static_cast<System::Byte>(204)));
			this->WSAstartup->Location = System::Drawing::Point(274, 101);
			this->WSAstartup->Name = L"WSAstartup";
			this->WSAstartup->Size = System::Drawing::Size(144, 29);
			this->WSAstartup->TabIndex = 6;
			this->WSAstartup->Text = L"WSA startup";
			// 
			// MyForm
			// 
			this->AutoScaleDimensions = System::Drawing::SizeF(9, 20);
			this->AutoScaleMode = System::Windows::Forms::AutoScaleMode::Font;
			this->ClientSize = System::Drawing::Size(786, 392);
			this->Controls->Add(this->WSAstartup);
			this->Controls->Add(this->SocketBind);
			this->Controls->Add(this->SocketState);
			this->Controls->Add(this->label_N_port);
			this->Controls->Add(this->button_Listen);
			this->Controls->Add(this->menuStrip1);
			this->MainMenuStrip = this->menuStrip1;
			this->Name = L"MyForm";
			this->Text = L"Сервер по приёму данных от микроконтроллера";
			this->menuStrip1->ResumeLayout(false);
			this->menuStrip1->PerformLayout();
			this->ResumeLayout(false);
			this->PerformLayout();

		}
#pragma endregion
	private: System::Void выходToolStripMenuItem_Click(System::Object^ sender, System::EventArgs^ e);
	private: System::Void button_Listen_Click(System::Object^ sender, System::EventArgs^ e);
	public:	   
		void SetTextValue(String^ text) {
			label_N_port->Text = text;
		};
	public:
		void SetSocketState_TextValue(String^ text) {
			SocketState->Text = text;
		};
	public:
		void SetSocketBind_TextValue(String^ text) {
			SocketBind->Text = text;
		};
	public:
		void SetWSA_TextValue(String^ text) {
			WSAstartup->Text = text;
		};
	}; 
}
