#include <tchar.h>
#include "MyForm.h"

using namespace PeopleCalculation;

[STAThread]

int WINAPI WinMain(HINSTANCE, HINSTANCE, LPSTR, int)
{
	Application::EnableVisualStyles();
	Application::SetCompatibleTextRenderingDefault(false);
	Application::Run(gcnew MyForm);

	CloseHandle(hSerial);

	return 0;
}

