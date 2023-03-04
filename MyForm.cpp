#include <tchar.h>
#include "MyForm.h"

using namespace PeopleCalculation;

[STAThread]

int WINAPI WinMain(HINSTANCE, HINSTANCE, LPSTR, int)
{
	Application::EnableVisualStyles();
	Application::SetCompatibleTextRenderingDefault(false);


	// Setting up a serial port
	DCB dcbSerialParams = { 0 };
	COMMTIMEOUTS timeouts = { 0 };

	hSerial = CreateFile(L"COM1:", GENERIC_READ | GENERIC_WRITE, 0,
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
//			CloseHandle(hSerial);
		}
		dcbSerialParams.BaudRate = CBR_115200;
		dcbSerialParams.ByteSize = 8;
		dcbSerialParams.Parity = NOPARITY;
		dcbSerialParams.StopBits = ONESTOPBIT;
		// Serial port parameters check
		if (!GetCommState(hSerial, &dcbSerialParams))
		{
			MessageBox::Show("Error getting serial port state");
//			CloseHandle(hSerial);
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
//			CloseHandle(hSerial);
		}
	}

	Application::Run(gcnew MyForm);
	CloseHandle(hSerial);
	return 0;
}

