#pragma once
#include "version.h"

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

			// Подписываемся на событие Load формы для автоматического запуска сервера
			this->Load += gcnew System::EventHandler(this, &MyForm::MyForm_Load);
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


	private: System::Windows::Forms::Label^ label_N_port;
	private: System::Windows::Forms::Label^ SocketState;
	private: System::Windows::Forms::Label^ SocketBind;
	private: System::Windows::Forms::Label^ WSAstartup;
	private: System::Windows::Forms::Label^ ClientAddr;
	private: System::Windows::Forms::Label^ labelMessage;
	private: System::Windows::Forms::Label^ labelVersion;





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
			this->label_N_port = (gcnew System::Windows::Forms::Label());
			this->SocketState = (gcnew System::Windows::Forms::Label());
			this->SocketBind = (gcnew System::Windows::Forms::Label());
			this->WSAstartup = (gcnew System::Windows::Forms::Label());
			this->ClientAddr = (gcnew System::Windows::Forms::Label());
			this->labelMessage = (gcnew System::Windows::Forms::Label());
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
			this->выходToolStripMenuItem->Size = System::Drawing::Size(80, 29);
			this->выходToolStripMenuItem->Text = L"Выход";
			this->выходToolStripMenuItem->Click += gcnew System::EventHandler(this, &MyForm::выходToolStripMenuItem_Click);
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
			// ClientAddr
			// 
			this->ClientAddr->AutoSize = true;
			this->ClientAddr->Font = (gcnew System::Drawing::Font(L"Microsoft Sans Serif", 12, System::Drawing::FontStyle::Regular, System::Drawing::GraphicsUnit::Point,
				static_cast<System::Byte>(204)));
			this->ClientAddr->Location = System::Drawing::Point(279, 308);
			this->ClientAddr->Name = L"ClientAddr";
			this->ClientAddr->Size = System::Drawing::Size(184, 29);
			this->ClientAddr->TabIndex = 7;
			this->ClientAddr->Text = L"Адрес клиента";
			// 
			// labelMessage
			// 
			this->labelMessage->AutoSize = true;
			this->labelMessage->Font = (gcnew System::Drawing::Font(L"Microsoft Sans Serif", 12, System::Drawing::FontStyle::Regular, System::Drawing::GraphicsUnit::Point,
				static_cast<System::Byte>(204)));
			this->labelMessage->Location = System::Drawing::Point(31, 371);
			this->labelMessage->Name = L"labelMessage";
			this->labelMessage->Size = System::Drawing::Size(150, 29);
			this->labelMessage->TabIndex = 8;
			this->labelMessage->Text = L"Сообщение";
			// 
			// MyForm
			// 
			this->AutoScaleDimensions = System::Drawing::SizeF(9, 20);
			this->AutoScaleMode = System::Windows::Forms::AutoScaleMode::Font;
			this->ClientSize = System::Drawing::Size(786, 442);
			// Инициализация метки версии
			this->labelVersion = (gcnew System::Windows::Forms::Label());
			this->labelVersion->AutoSize = true;
			this->labelVersion->Location = System::Drawing::Point(this->ClientSize.Width - 100, this->ClientSize.Height - 30);
			this->labelVersion->Name = L"labelVersion";
			this->labelVersion->Size = System::Drawing::Size(80, 20);
			this->labelVersion->Text = L"v" VERSION_STR;
			this->labelVersion->Anchor = static_cast<System::Windows::Forms::AnchorStyles>(System::Windows::Forms::AnchorStyles::Bottom | System::Windows::Forms::AnchorStyles::Right);

			this->Controls->Add(this->labelVersion);
			this->Controls->Add(this->labelMessage);
			this->Controls->Add(this->ClientAddr);
			this->Controls->Add(this->WSAstartup);
			this->Controls->Add(this->SocketBind);
			this->Controls->Add(this->SocketState);
			this->Controls->Add(this->label_N_port);
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
	private: System::Void MyForm_Load(System::Object^ sender, System::EventArgs^ e);
	public:
		void SetTextValue(String^ text) {
			if (this != nullptr && this->InvokeRequired) {
				this->Invoke(gcnew Action<String^>(this, &MyForm::SetTextValue), text);
			}
			else if (this != nullptr && label_N_port != nullptr) {
				label_N_port->Text = text;
			}
		};
		void SetSocketState_TextValue(String^ text) {
			if (this != nullptr && this->InvokeRequired) {
				this->Invoke(gcnew Action<String^>(this, &MyForm::SetSocketState_TextValue), text);
			}
			else if (this != nullptr && SocketState != nullptr) {
				SocketState->Text = text;
			}
		};
		void SetSocketBind_TextValue(String^ text) {
			if (this != nullptr && this->InvokeRequired) {
				this->Invoke(gcnew Action<String^>(this, &MyForm::SetSocketBind_TextValue), text);
			}
			else if (this != nullptr && SocketBind != nullptr) {
				SocketBind->Text = text;
			}
		};
		void SetWSA_TextValue(String^ text) {
			if (this != nullptr && this->InvokeRequired) {
				this->Invoke(gcnew Action<String^>(this, &MyForm::SetWSA_TextValue), text);
			}
			else if (this != nullptr && WSAstartup != nullptr) {
				WSAstartup->Text = text;
			}
		}
		void SetClientAddr_TextValue(String^ text) {
			if (this != nullptr && this->InvokeRequired) {
				this->Invoke(gcnew Action<String^>(this, &MyForm::SetClientAddr_TextValue), text);
			}
			else if (this != nullptr && WSAstartup != nullptr) {
				ClientAddr->Text = text;
			}
		}
		void SetMessage_TextValue(String^ text) {
			if (this != nullptr && this->InvokeRequired) {
				this->Invoke(gcnew Action<String^>(this, &MyForm::SetMessage_TextValue), text);
			}
			else if (this != nullptr && WSAstartup != nullptr) {
				labelMessage->Text = text;
			}
		}
	};
}
