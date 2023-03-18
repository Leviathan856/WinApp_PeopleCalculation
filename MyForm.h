﻿#include <Windows.h>
//#include <comdef.h>
#include <vector>
#include <ctime>
#include <vcclr.h>
#include <string>
#include <SetupAPI.h>
//#include <WinBase.h>

#pragma once

#pragma comment(lib, "user32.lib")
#pragma comment(lib, "setupapi.lib")
//#pragma comment(lib, "kernel32.lib")



// Set the global variables
HANDLE hSerial;
HWND comboBoxHandle;
std::vector<uint8_t> dataToWrite;
std::vector<uint8_t> timeToWrite;


// Define the CRC8 polynom and lookup table
#define CRC8_POLY 0x07
const uint8_t crc8_table[] = {
	0x00, 0x07, 0x0E, 0x09, 0x1C, 0x1B, 0x12, 0x15,
	0x38, 0x3F, 0x36, 0x31, 0x24, 0x23, 0x2A, 0x2D,
	0x70, 0x77, 0x7E, 0x79, 0x6C, 0x6B, 0x62, 0x65,
	0x48, 0x4F, 0x46, 0x41, 0x54, 0x53, 0x5A, 0x5D,
	0xE0, 0xE7, 0xEE, 0xE9, 0xFC, 0xFB, 0xF2, 0xF5,
	0xD8, 0xDF, 0xD6, 0xD1, 0xC4, 0xC3, 0xCA, 0xCD,
	0x90, 0x97, 0x9E, 0x99, 0x8C, 0x8B, 0x82, 0x85,
	0xA8, 0xAF, 0xA6, 0xA1, 0xB4, 0xB3, 0xBA, 0xBD,
};


namespace PeopleCalculation {

	using namespace System;
	using namespace System::ComponentModel;
	using namespace System::Collections;
	using namespace System::Windows::Forms;
	using namespace System::Data;
	using namespace System::Drawing;

	public ref class MyForm : public System::Windows::Forms::Form
	{
	public:
		MyForm(void)
		{
			InitializeComponent();
			IntPtr hwndComboBox = comboBox1->Handle;
			HWND comboBoxHandle = static_cast<HWND>(hwndComboBox.ToPointer());
			EnumerateDevices(comboBoxHandle);
			if (comboBox1->Items->Count == 0)
			{
				comboBox1->Visible = false;

				label1->Visible = true;
				label2->Visible = true;
				label3->Visible = true;
				label4->Visible = true;
				button1->Visible = true;
				button2->Visible = true;
				button3->Visible = true;
				button4->Visible = true;

				MessageBox::Show("No serial port devices are connected!");
			}
		}

	protected:
		~MyForm()
		{
			if (components)
			{
				delete components;
			}
		}
	private: System::Windows::Forms::Button^ button1;
	private: System::Windows::Forms::Button^ button3;
	private: System::Windows::Forms::Button^ button2;
	protected:

	private: System::Windows::Forms::Label^ label1;
	private: System::Windows::Forms::Label^ label2;
	private: System::Windows::Forms::Label^ label3;
	private: System::Windows::Forms::Label^ label4;
	private: System::Windows::Forms::Button^ button4;
	private: System::Windows::Forms::ComboBox^ comboBox1;

	private:
		System::ComponentModel::Container^ components;

#pragma region Windows Form Designer generated code
		void InitializeComponent(void)
		{
			this->button1 = (gcnew System::Windows::Forms::Button());
			this->button3 = (gcnew System::Windows::Forms::Button());
			this->button2 = (gcnew System::Windows::Forms::Button());
			this->label1 = (gcnew System::Windows::Forms::Label());
			this->label2 = (gcnew System::Windows::Forms::Label());
			this->label3 = (gcnew System::Windows::Forms::Label());
			this->label4 = (gcnew System::Windows::Forms::Label());
			this->button4 = (gcnew System::Windows::Forms::Button());
			this->comboBox1 = (gcnew System::Windows::Forms::ComboBox());
			this->SuspendLayout();
			// 
			// button1
			// 
			this->button1->Font = (gcnew System::Drawing::Font(L"Times New Roman", 11, System::Drawing::FontStyle::Bold, System::Drawing::GraphicsUnit::Point,
				static_cast<System::Byte>(0)));
			this->button1->Location = System::Drawing::Point(257, 520);
			this->button1->Name = L"button1";
			this->button1->Size = System::Drawing::Size(178, 52);
			this->button1->TabIndex = 0;
			this->button1->Text = L"Start";
			this->button1->UseVisualStyleBackColor = true;
			this->button1->Visible = false;
			this->button1->Click += gcnew System::EventHandler(this, &MyForm::button1_Click);
			// 
			// button3
			// 
			this->button3->Font = (gcnew System::Drawing::Font(L"Times New Roman", 11, System::Drawing::FontStyle::Bold, System::Drawing::GraphicsUnit::Point,
				static_cast<System::Byte>(0)));
			this->button3->Location = System::Drawing::Point(529, 520);
			this->button3->Name = L"button3";
			this->button3->Size = System::Drawing::Size(269, 52);
			this->button3->TabIndex = 1;
			this->button3->Text = L" Set time and date";
			this->button3->UseVisualStyleBackColor = true;
			this->button3->Visible = false;
			this->button3->Click += gcnew System::EventHandler(this, &MyForm::button3_Click);
			// 
			// button2
			// 
			this->button2->Font = (gcnew System::Drawing::Font(L"Times New Roman", 11, System::Drawing::FontStyle::Bold, System::Drawing::GraphicsUnit::Point,
				static_cast<System::Byte>(0)));
			this->button2->Location = System::Drawing::Point(874, 520);
			this->button2->Name = L"button2";
			this->button2->Size = System::Drawing::Size(199, 52);
			this->button2->TabIndex = 2;
			this->button2->Text = L"Delete data";
			this->button2->UseVisualStyleBackColor = true;
			this->button2->Visible = false;
			this->button2->Click += gcnew System::EventHandler(this, &MyForm::button2_Click);
			// 
			// label1
			// 
			this->label1->AutoSize = true;
			this->label1->Font = (gcnew System::Drawing::Font(L"Times New Roman", 10.875F, System::Drawing::FontStyle::Regular, System::Drawing::GraphicsUnit::Point,
				static_cast<System::Byte>(0)));
			this->label1->Location = System::Drawing::Point(499, 201);
			this->label1->Name = L"label1";
			this->label1->Size = System::Drawing::Size(193, 33);
			this->label1->TabIndex = 3;
			this->label1->Text = L"Total people in:";
			this->label1->Visible = false;
			// 
			// label2
			// 
			this->label2->AutoSize = true;
			this->label2->Font = (gcnew System::Drawing::Font(L"Times New Roman", 10.875F, System::Drawing::FontStyle::Regular, System::Drawing::GraphicsUnit::Point,
				static_cast<System::Byte>(0)));
			this->label2->Location = System::Drawing::Point(499, 292);
			this->label2->Name = L"label2";
			this->label2->Size = System::Drawing::Size(207, 33);
			this->label2->TabIndex = 4;
			this->label2->Text = L"Total people out:";
			this->label2->Visible = false;
			// 
			// label3
			// 
			this->label3->AutoSize = true;
			this->label3->Font = (gcnew System::Drawing::Font(L"Times New Roman", 10.875F, System::Drawing::FontStyle::Regular, System::Drawing::GraphicsUnit::Point,
				static_cast<System::Byte>(0)));
			this->label3->Location = System::Drawing::Point(747, 201);
			this->label3->Name = L"label3";
			this->label3->Size = System::Drawing::Size(36, 33);
			this->label3->TabIndex = 5;
			this->label3->Text = L"...";
			this->label3->Visible = false;
			// 
			// label4
			// 
			this->label4->AutoSize = true;
			this->label4->Font = (gcnew System::Drawing::Font(L"Times New Roman", 10.875F, System::Drawing::FontStyle::Regular, System::Drawing::GraphicsUnit::Point,
				static_cast<System::Byte>(0)));
			this->label4->Location = System::Drawing::Point(747, 292);
			this->label4->Name = L"label4";
			this->label4->Size = System::Drawing::Size(36, 33);
			this->label4->TabIndex = 6;
			this->label4->Text = L"...";
			this->label4->Visible = false;
			// 
			// button4
			// 
			this->button4->Font = (gcnew System::Drawing::Font(L"Times New Roman", 11, System::Drawing::FontStyle::Bold, System::Drawing::GraphicsUnit::Point,
				static_cast<System::Byte>(0)));
			this->button4->Location = System::Drawing::Point(554, 611);
			this->button4->Name = L"button4";
			this->button4->Size = System::Drawing::Size(217, 52);
			this->button4->TabIndex = 7;
			this->button4->Text = L"Show logs";
			this->button4->UseVisualStyleBackColor = true;
			this->button4->Visible = false;
			this->button4->Click += gcnew System::EventHandler(this, &MyForm::button4_Click);
			// 
			// comboBox1
			// 
			this->comboBox1->Font = (gcnew System::Drawing::Font(L"Times New Roman", 10.875F, System::Drawing::FontStyle::Regular, System::Drawing::GraphicsUnit::Point,
				static_cast<System::Byte>(0)));
			this->comboBox1->FormattingEnabled = true;
			this->comboBox1->Location = System::Drawing::Point(354, 102);
			this->comboBox1->Name = L"comboBox1";
			this->comboBox1->Size = System::Drawing::Size(605, 41);
			this->comboBox1->TabIndex = 8;
			this->comboBox1->Text = L"                            Choose device";
			this->comboBox1->SelectedIndexChanged += gcnew System::EventHandler(this, &MyForm::comboBox1_SelectedIndexChanged);
			// 
			// MyForm
			// 
			this->AutoScaleDimensions = System::Drawing::SizeF(12, 25);
			this->AutoScaleMode = System::Windows::Forms::AutoScaleMode::Font;
			this->AutoSize = true;
			this->BackgroundImageLayout = System::Windows::Forms::ImageLayout::None;
			this->ClientSize = System::Drawing::Size(1332, 758);
			this->Controls->Add(this->comboBox1);
			this->Controls->Add(this->button4);
			this->Controls->Add(this->label4);
			this->Controls->Add(this->label3);
			this->Controls->Add(this->label2);
			this->Controls->Add(this->label1);
			this->Controls->Add(this->button2);
			this->Controls->Add(this->button3);
			this->Controls->Add(this->button1);
			this->Name = L"MyForm";
			this->ShowIcon = false;
			this->Text = L"People Calculation";
			this->ResumeLayout(false);
			this->PerformLayout();

		}
#pragma endregion



	// Get checksun
	uint8_t crc8(const std::vector<uint8_t>& data)
	{
		uint8_t crc = 0x00;
		for (const auto& byte : data) {
			crc = crc8_table[crc ^ byte];
		}
		return crc;
	}


	// Split 4-byte number into bytes
	std::vector<uint8_t> splitToBytes(uint32_t number)
	{
		uint8_t byte1 = (number >> 24) & 0xFF;
		uint8_t byte2 = (number >> 16) & 0xFF;
		uint8_t byte3 = (number >> 8) & 0xFF;
		uint8_t byte4 = number & 0xFF;
		return { byte1, byte2, byte3, byte4 };
	}

	// Compose 4 bytes into 4-byte number
	uint32_t composeToNumber(std::vector<uint8_t> bytes, int startIndex)
	{
		return (static_cast<uint32_t>(bytes[startIndex]) << 24) |
			   (static_cast<uint32_t>(bytes[startIndex + 1]) << 16) |
			   (static_cast<uint32_t>(bytes[startIndex + 2]) << 8) |
			   static_cast<uint32_t>(bytes[startIndex + 3]);
	}


	// Write data to serial port
	void write(std::vector<uint8_t> instrNumber, bool showMessages)
	{
		// Sending an instruction to serial port
		DWORD bytesWritten;

		if (!WriteFile(hSerial, LP(&instrNumber), instrNumber.size(), &bytesWritten, NULL))
		{
			if (showMessages) MessageBox::Show("Error writing to serial port");
		}
		if (showMessages) MessageBox::Show("Bytes written: " + bytesWritten);		// Optional
	}
	
	// Read data from serial port
	std::vector<uint8_t> read(bool showMessages)
	{
		// Reading the data from serial port
		std::vector<uint8_t> dataToRead;
		DWORD bytesRead;

		if (!ReadFile(hSerial, LP(&dataToRead), dataToRead.size(), &bytesRead, NULL))
		{
			if (showMessages) MessageBox::Show("Error reading from serial port");
		}
		if (showMessages) MessageBox::Show("Bytes read: " + bytesRead);			// Optional
		return dataToRead;
	}


	// Send data and return the answer
	std::vector<uint8_t> sendData(std::vector<uint8_t> dataToWrite)
	{
		uint8_t checksum = crc8(dataToWrite);
		dataToWrite.push_back(checksum);
		write(dataToWrite, (dataToWrite[0]) > 2);
		std::vector<uint8_t> receivedData = read((dataToWrite[0]) > 2);
		if (receivedData.empty())
		{
			return { 0 };
		}
		else if (crc8(dataToWrite) != crc8(receivedData))
		{
			MessageBox::Show("Checksum error: Data is corrupted");
			return { 0 };
		}
		return receivedData;
	}


	// Check if device connected
	bool isConnected()
	{
		bool connection = false;
		if (sendData({ 1 })[0] == 0)
		{
			MessageBox::Show("Please connect a required device, and choose it in the list.");
		}
		else connection = true;
		return connection;
	}


	// Helper function to convert a WCHAR string to std::string
	std::string WideToMultiByte(const WCHAR* wideString) {
		int size = WideCharToMultiByte(CP_UTF8, 0, wideString, -1, nullptr, 0, nullptr, nullptr);
		std::string result(size, '\0');
		WideCharToMultiByte(CP_UTF8, 0, wideString, -1, (char*)(result).data(), size, nullptr, nullptr);
		return result;
	}

	// Enumerate connected devices and add their friendly names to the combobox
	void EnumerateDevices(HWND combobox) {
		// Get the device information set for all present devices
		HDEVINFO deviceInfoSet = SetupDiGetClassDevs(&GUID_DEVINTERFACE_SERENUM_BUS_ENUMERATOR, nullptr, nullptr, DIGCF_PRESENT);
		if (deviceInfoSet == INVALID_HANDLE_VALUE) {
			// Failed to get device information set
			return;
		}
		// Enumerate all devices in the set and add their friendly names to the combobox
		SP_DEVINFO_DATA deviceInfoData;
		deviceInfoData.cbSize = sizeof(SP_DEVINFO_DATA);
		DWORD index = 0;
		while (SetupDiEnumDeviceInfo(deviceInfoSet, index++, &deviceInfoData)) {
			// Get the friendly name of the device
			WCHAR friendlyName[MAX_PATH];
			if (SetupDiGetDeviceRegistryProperty(deviceInfoSet, &deviceInfoData, SPDRP_FRIENDLYNAME, nullptr,
				reinterpret_cast<PBYTE>(friendlyName), sizeof(friendlyName), nullptr)) {
				// Add the friendly name to the combobox
				std::string friendlyNameMB = WideToMultiByte(friendlyName);
				//SendMessage(combobox, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(friendlyNameMB.c_str()));
				comboBox1->Items->Add(gcnew String(friendlyNameMB.c_str()));
			}
		}
		// Clean up
		SetupDiDestroyDeviceInfoList(deviceInfoSet);
	}

	// Extract COM port number from device name
	int extractCOMPortNumber(String^ deviceName)
	{
		int COMPortNumber = -1;

		for (int i = 0; i < deviceName->Length; i++)
		{
			if (deviceName[i] == 'C')
				if (deviceName[i + 1] == 'O')
					if (deviceName[i + 2] == 'M')
						COMPortNumber = int(deviceName[i + 3] - '0');
		}
		return COMPortNumber;
	}

// Get data from device
private: System::Void button1_Click(System::Object^ sender, System::EventArgs^ e) {

	// Check if device connected
	if (isConnected())
	{

		// Check if data is available
		if (sendData({ 2 }).size() == 0)
		{
			MessageBox::Show("Error accessing data: There is no data available");
			return;
		}
		bool quitKey = true;
		while (quitKey)
		{
			// Number of people in
			dataToWrite = sendData({ 3 });
			label3->Text = Convert::ToString((dataToWrite.size() > 1) ? dataToWrite[0] : 0);

			// Number of people out
			dataToWrite = sendData({ 4 });
			label4->Text = Convert::ToString((dataToWrite.size() > 1) ? dataToWrite[0] : 0);

			// Set the timer
			clock_t timestamp = clock();
			while (quitKey && (clock() - timestamp) < 2000)
			{
				// Check if the user has pressed the "Q" key
				if (GetAsyncKeyState('Q') & 0x8000) {
					MessageBox::Show("Q was pressed");
					quitKey = false;
				}
			} 
		}
	}
}

// Reset device data
private: System::Void button2_Click(System::Object^ sender, System::EventArgs^ e) {

	// Check if device connected
	if (isConnected())
	{
		// Reseting data
		if (sendData({ 5 }).size() == 0)
		{
			MessageBox::Show("Failed to reset data");
		}
	}
}

// Set device time
private: System::Void button3_Click(System::Object^ sender, System::EventArgs^ e) {

	// Check if device connected
	if (isConnected())
	{
		// Set device time
		std::vector<uint8_t> unixTimeBitwise = splitToBytes(1676929636);
		if (sendData({ 6, unixTimeBitwise[0], unixTimeBitwise[1],
			unixTimeBitwise[2], unixTimeBitwise[3] }).size() == 0)
		{
			MessageBox::Show("Error setting time and date");
		}
	}
}

// Get device time
private: System::Void button4_Click(System::Object^ sender, System::EventArgs^ e) {

	// Check if device connected
	if (isConnected())
	{
		// Getting device time
		timeToWrite = sendData({ 7 });
		if (timeToWrite.size() == 0)
		{
			MessageBox::Show("Error getting time and date");
		}
		else
		{
			uint32_t unixTime = composeToNumber(timeToWrite, 0);
			MessageBox::Show("Unix time is " + (int)unixTime);
		}
	}
}

private: System::Void comboBox1_SelectedIndexChanged(System::Object^ sender, System::EventArgs^ e) {


	// COM port selection
	String^ COMPort = "COM";	// Must be in "COM#:" format!

	COMPort += extractCOMPortNumber(comboBox1->SelectedItem->ToString()) + ":";
	pin_ptr<const wchar_t> wch = PtrToStringChars(COMPort);
	const wchar_t* wcharCOMPort = wch;

	// Setting up a serial port
	DCB dcbSerialParams = { 0 };
	COMMTIMEOUTS timeouts = { 0 };

	hSerial = CreateFile(wcharCOMPort, GENERIC_READ | GENERIC_WRITE, 0,
		NULL, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL);

	if (hSerial == INVALID_HANDLE_VALUE)
	{
		MessageBox::Show("Error opening serial port.");
	}
	else
	{
		// Setting the serial port parameters
		dcbSerialParams.DCBlength = sizeof(dcbSerialParams);
		if (!GetCommState(hSerial, &dcbSerialParams))
		{
			MessageBox::Show("Error getting serial port state");
		}
		dcbSerialParams.BaudRate = CBR_115200;
		dcbSerialParams.ByteSize = 8;
		dcbSerialParams.Parity = NOPARITY;
		dcbSerialParams.StopBits = ONESTOPBIT;
		// Serial port parameters check
		if (!GetCommState(hSerial, &dcbSerialParams))
		{
			MessageBox::Show("Error getting serial port state");
		}

		// Setting a 2-second timeouts for reading and writing data
		timeouts.ReadTotalTimeoutConstant = 2000;
		timeouts.ReadIntervalTimeout = 2000;
		timeouts.ReadTotalTimeoutMultiplier = 10;
		timeouts.WriteTotalTimeoutConstant = 2000;
		timeouts.WriteTotalTimeoutMultiplier = 10;
		// Timeouts settings check
		if (!SetCommTimeouts(hSerial, &timeouts))
		{
			MessageBox::Show("Error setting timeouts");
		}
	}

	// Switch to the next screen
	comboBox1->Visible = false;

	label1->Visible = true;
	label2->Visible = true;
	label3->Visible = true;
	label4->Visible = true;

	button1->Visible = true;
	button2->Visible = true;
	button3->Visible = true;
	button4->Visible = true;

}

};
};

