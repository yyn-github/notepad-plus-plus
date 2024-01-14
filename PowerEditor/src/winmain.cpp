// This file is part of Notepad++ project
// Copyright (C)2021 Don HO <don.h@free.fr>

// This program is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// at your option any later version.
//
// This program is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
// GNU General Public License for more details.
//
// You should have received a copy of the GNU General Public License
// along with this program.  If not, see <https://www.gnu.org/licenses/>.

#include "Notepad_plus_Window.h"
#include "Processus.h"
#include "Win32Exception.h"	//Win32 exception
#include "MiniDumper.h"			//Write dump files
#include "verifySignedfile.h"
#include "NppDarkMode.h"
#include <memory>

typedef std::vector<std::wstring> ParamVector;


namespace
{


void allowPrivilegeMessages(const Notepad_plus_Window& notepad_plus_plus, winVer winVer)
{
	#ifndef MSGFLT_ADD
	const DWORD MSGFLT_ADD = 1;
	#endif
	#ifndef MSGFLT_ALLOW
	const DWORD MSGFLT_ALLOW = 1;
	#endif
	// Tell UAC that lower integrity processes are allowed to send WM_COPYDATA (or other) messages to this process (or window)
	// This (WM_COPYDATA) allows opening new files to already opened elevated Notepad++ process via explorer context menu.
	if (winVer >= WV_VISTA || winVer == WV_UNKNOWN)
	{
		HMODULE hDll = GetModuleHandle(TEXT("user32.dll"));
		if (hDll)
		{
			// According to MSDN ChangeWindowMessageFilter may not be supported in future versions of Windows,
			// that is why we use ChangeWindowMessageFilterEx if it is available (windows version >= Win7).
			if (winVer == WV_VISTA)
			{
				typedef BOOL (WINAPI *MESSAGEFILTERFUNC)(UINT message,DWORD dwFlag);

				MESSAGEFILTERFUNC func = (MESSAGEFILTERFUNC)::GetProcAddress( hDll, "ChangeWindowMessageFilter" );

				if (func)
				{
					func(WM_COPYDATA, MSGFLT_ADD);
					func(NPPM_INTERNAL_RESTOREFROMTRAY, MSGFLT_ADD);
				}
			}
			else
			{
				typedef BOOL (WINAPI *MESSAGEFILTERFUNCEX)(HWND hWnd,UINT message,DWORD action,VOID* pChangeFilterStruct);

				MESSAGEFILTERFUNCEX funcEx = (MESSAGEFILTERFUNCEX)::GetProcAddress( hDll, "ChangeWindowMessageFilterEx" );

				if (funcEx)
				{
					funcEx(notepad_plus_plus.getHSelf(), WM_COPYDATA, MSGFLT_ALLOW, NULL);
					funcEx(notepad_plus_plus.getHSelf(), NPPM_INTERNAL_RESTOREFROMTRAY, MSGFLT_ALLOW, NULL);
				}
			}
		}
	}
}

// parseCommandLine() takes command line arguments part string, cuts arguments by using white space as separater.
// Only white space in double quotes will be kept, such as file path argument or '-settingsDir=' argument (ex.: -settingsDir="c:\my settings\my folder\")
// if '-z' is present, the 3rd argument after -z wont be cut - ie. all the space will also be kept
// ex.: '-notepadStyleCmdline -z "C:\WINDOWS\system32\NOTEPAD.EXE" C:\my folder\my file with whitespace.txt' will be separated to: 
// 1. "-notepadStyleCmdline"
// 2. "-z"
// 3. "C:\WINDOWS\system32\NOTEPAD.EXE"
// 4. "C:\my folder\my file with whitespace.txt" 
void parseCommandLine(const TCHAR* commandLine, ParamVector& paramVector)
{
	if (!commandLine)
		return;
	
	TCHAR* cmdLine = new TCHAR[lstrlen(commandLine) + 1];
	lstrcpy(cmdLine, commandLine);

	TCHAR* cmdLinePtr = cmdLine;

	bool isBetweenFileNameQuotes = false;
	bool isStringInArg = false;
	bool isInWhiteSpace = true;

	int zArg = 0; // for "-z" argument: Causes Notepad++ to ignore the next command line argument (a single word, or a phrase in quotes).
	              // The only intended and supported use for this option is for the Notepad Replacement syntax.

	bool shouldBeTerminated = false; // If "-z" argument has been found, zArg value will be increased from 0 to 1.
	                                 // then after processing next argument of "-z", zArg value will be increased from 1 to 2.
	                                 // when zArg == 2 shouldBeTerminated will be set to true - it will trigger the treatment which consider the rest as a argument, with or without white space(s).

	size_t commandLength = lstrlen(cmdLinePtr);
	std::vector<TCHAR *> args;
	for (size_t i = 0; i < commandLength && !shouldBeTerminated; ++i)
	{
		switch (cmdLinePtr[i])
		{
			case '\"': //quoted filename, ignore any following whitespace
			{
				if (!isStringInArg && !isBetweenFileNameQuotes && i > 0 && cmdLinePtr[i-1] == '=')
				{
					isStringInArg = true;
				}
				else if (isStringInArg)
				{
					isStringInArg = false;
				}
				else if (!isBetweenFileNameQuotes)	//" will always be treated as start or end of param, in case the user forgot to add an space
				{
					args.push_back(cmdLinePtr + i + 1);	//add next param(since zero terminated original, no overflow of +1)
					isBetweenFileNameQuotes = true;
					cmdLinePtr[i] = 0;

					if (zArg == 1)
					{
						++zArg; // zArg == 2
					}
				}
				else if (isBetweenFileNameQuotes)
				{
					isBetweenFileNameQuotes = false;
					//because we dont want to leave in any quotes in the filename, remove them now (with zero terminator)
					cmdLinePtr[i] = 0;
				}
				isInWhiteSpace = false;
			}
			break;

			case '\t': //also treat tab as whitespace
			case ' ':
			{
				isInWhiteSpace = true;
				if (!isBetweenFileNameQuotes && !isStringInArg)
				{
					cmdLinePtr[i] = 0;		//zap spaces into zero terminators, unless its part of a filename

					size_t argsLen = args.size();
					if (argsLen > 0 && lstrcmp(args[argsLen-1], L"-z") == 0)
						++zArg; // "-z" argument is found: change zArg value from 0 (initial) to 1
				}
			}
			break;

			default: //default TCHAR, if beginning of word, add it
			{
				if (!isBetweenFileNameQuotes && !isStringInArg && isInWhiteSpace)
				{
					args.push_back(cmdLinePtr + i);	//add next param
					if (zArg == 2)
					{
						shouldBeTerminated = true; // stop the processing, and keep the rest string as it in the vector
					}

					isInWhiteSpace = false;
				}
			}
		}
	}
	paramVector.assign(args.begin(), args.end());
	delete[] cmdLine;
}

// Converts /p or /P to -quickPrint if it exists as the first parameter
// This seems to mirror Notepad's behaviour
void convertParamsToNotepadStyle(ParamVector& params)
{
	for (auto it = params.begin(); it != params.end(); ++it)
	{
		if (lstrcmp(it->c_str(), TEXT("/p")) == 0 || lstrcmp(it->c_str(), TEXT("/P")) == 0)
		{
			it->assign(TEXT("-quickPrint"));
		}
	}
}

bool isInList(const TCHAR *token2Find, ParamVector& params, bool eraseArg = true)
{
	for (auto it = params.begin(); it != params.end(); ++it)
	{
		if (lstrcmp(token2Find, it->c_str()) == 0)
		{
			if (eraseArg) params.erase(it);
			return true;
		}
	}
	return false;
}

bool getParamVal(TCHAR c, ParamVector & params, std::wstring & value)
{
	value = TEXT("");
	size_t nbItems = params.size();

	for (size_t i = 0; i < nbItems; ++i)
	{
		const TCHAR * token = params.at(i).c_str();
		if (token[0] == '-' && lstrlen(token) >= 2 && token[1] == c) //dash, and enough chars
		{
			value = (token+2);
			params.erase(params.begin() + i);
			return true;
		}
	}
	return false;
}

bool getParamValFromString(const TCHAR *str, ParamVector & params, std::wstring & value)
{
	value = TEXT("");
	size_t nbItems = params.size();

	for (size_t i = 0; i < nbItems; ++i)
	{
		const TCHAR * token = params.at(i).c_str();
		std::wstring tokenStr = token;
		size_t pos = tokenStr.find(str);
		if (pos != std::wstring::npos && pos == 0)
		{
			value = (token + lstrlen(str));
			params.erase(params.begin() + i);
			return true;
		}
	}
	return false;
}

LangType getLangTypeFromParam(ParamVector & params)
{
	std::wstring langStr;
	if (!getParamVal('l', params, langStr))
		return L_EXTERNAL;
	return NppParameters::getLangIDFromStr(langStr.c_str());
}

std::wstring getLocalizationPathFromParam(ParamVector & params)
{
	std::wstring locStr;
	if (!getParamVal('L', params, locStr))
		return TEXT("");
	locStr = stringToLower(stringReplace(locStr, TEXT("_"), TEXT("-"))); // convert to lowercase format with "-" as separator
	return NppParameters::getLocPathFromStr(locStr.c_str());
}

intptr_t getNumberFromParam(char paramName, ParamVector & params, bool & isParamePresent)
{
	std::wstring numStr;
	if (!getParamVal(paramName, params, numStr))
	{
		isParamePresent = false;
		return -1;
	}
	isParamePresent = true;
	return static_cast<intptr_t>(_ttoi64(numStr.c_str()));
}

std::wstring getEasterEggNameFromParam(ParamVector & params, unsigned char & type)
{
	std::wstring EasterEggName;
	if (!getParamValFromString(TEXT("-qn="), params, EasterEggName))  // get internal easter egg
	{
		if (!getParamValFromString(TEXT("-qt="), params, EasterEggName)) // get user quote from cmdline argument
		{
			if (!getParamValFromString(TEXT("-qf="), params, EasterEggName)) // get user quote from a content of file
				return TEXT("");
			else
			{
				type = 2; // quote content in file
			}
		}
		else
			type = 1; // commandline quote
	}
	else
		type = 0; // easter egg

	if (EasterEggName.c_str()[0] == '"' && EasterEggName.c_str()[EasterEggName.length() - 1] == '"')
	{
		EasterEggName = EasterEggName.substr(1, EasterEggName.length() - 2);
	}

	if (type == 2)
		EasterEggName = relativeFilePathToFullFilePath(EasterEggName.c_str());

	return EasterEggName;
}

int getGhostTypingSpeedFromParam(ParamVector & params)
{
	std::wstring speedStr;
	if (!getParamValFromString(TEXT("-qSpeed"), params, speedStr))
		return -1;
	
	int speed = std::stoi(speedStr, 0);
	if (speed <= 0 || speed > 3)
		return -1;

	return speed;
}

const TCHAR FLAG_MULTI_INSTANCE[] = TEXT("-multiInst");
const TCHAR FLAG_NO_PLUGIN[] = TEXT("-noPlugin");
const TCHAR FLAG_READONLY[] = TEXT("-ro");
const TCHAR FLAG_NOSESSION[] = TEXT("-nosession");
const TCHAR FLAG_NOTABBAR[] = TEXT("-notabbar");
const TCHAR FLAG_SYSTRAY[] = TEXT("-systemtray");
const TCHAR FLAG_LOADINGTIME[] = TEXT("-loadingTime");
const TCHAR FLAG_HELP[] = TEXT("--help");
const TCHAR FLAG_ALWAYS_ON_TOP[] = TEXT("-alwaysOnTop");
const TCHAR FLAG_OPENSESSIONFILE[] = TEXT("-openSession");
const TCHAR FLAG_RECURSIVE[] = TEXT("-r");
const TCHAR FLAG_FUNCLSTEXPORT[] = TEXT("-export=functionList");
const TCHAR FLAG_PRINTANDQUIT[] = TEXT("-quickPrint");
const TCHAR FLAG_NOTEPAD_COMPATIBILITY[] = TEXT("-notepadStyleCmdline");
const TCHAR FLAG_OPEN_FOLDERS_AS_WORKSPACE[] = TEXT("-openFoldersAsWorkspace");
const TCHAR FLAG_SETTINGS_DIR[] = TEXT("-settingsDir=");
const TCHAR FLAG_TITLEBAR_ADD[] = TEXT("-titleAdd=");
const TCHAR FLAG_APPLY_UDL[] = TEXT("-udl=");
const TCHAR FLAG_PLUGIN_MESSAGE[] = TEXT("-pluginMessage=");
const TCHAR FLAG_MONITOR_FILES[] = TEXT("-monitor");

void doException(Notepad_plus_Window & notepad_plus_plus)
{
	Win32Exception::removeHandler();	//disable exception handler after excpetion, we dont want corrupt data structurs to crash the exception handler
	::MessageBox(Notepad_plus_Window::gNppHWND, TEXT("Notepad++ will attempt to save any unsaved data. However, dataloss is very likely."), TEXT("Recovery initiating"), MB_OK | MB_ICONINFORMATION);

	TCHAR tmpDir[1024];
	GetTempPath(1024, tmpDir);
	std::wstring emergencySavedDir = tmpDir;
	emergencySavedDir += TEXT("\\Notepad++ RECOV");

	bool res = notepad_plus_plus.emergency(emergencySavedDir);
	if (res)
	{
		std::wstring displayText = TEXT("Notepad++ was able to successfully recover some unsaved documents, or nothing to be saved could be found.\r\nYou can find the results at :\r\n");
		displayText += emergencySavedDir;
		::MessageBox(Notepad_plus_Window::gNppHWND, displayText.c_str(), TEXT("Recovery success"), MB_OK | MB_ICONINFORMATION);
	}
	else
		::MessageBox(Notepad_plus_Window::gNppHWND, TEXT("Unfortunatly, Notepad++ was not able to save your work. We are sorry for any lost data."), TEXT("Recovery failure"), MB_OK | MB_ICONERROR);
}

// Looks for -z arguments and strips command line arguments following those, if any
void stripIgnoredParams(ParamVector & params)
{
	for (auto it = params.begin(); it != params.end(); )
	{
		if (lstrcmp(it->c_str(), TEXT("-z")) == 0)
		{
			auto nextIt = std::next(it);
			if ( nextIt != params.end() )
			{
				params.erase(nextIt);
			}
			it = params.erase(it);
		}
		else
		{
			++it;
		}
	}
}

} // namespace


std::chrono::steady_clock::time_point g_nppStartTimePoint{};


int WINAPI wWinMain(_In_ HINSTANCE hInstance, _In_opt_ HINSTANCE /*hPrevInstance*/, _In_ PWSTR pCmdLine, _In_ int /*nShowCmd*/)
{
	g_nppStartTimePoint = std::chrono::steady_clock::now();

	bool TheFirstOne = true;
	::SetLastError(NO_ERROR);
	::CreateMutex(NULL, false, TEXT("nppInstance"));
	if (::GetLastError() == ERROR_ALREADY_EXISTS)
		TheFirstOne = false;

	CmdLineParams cmdLineParams;
	NppParameters& nppParameters = NppParameters::getInstance();
	bool isMultiInst = false;

	//Only after loading all the file paths set the working directory
	::SetCurrentDirectory(NppParameters::getInstance().getNppPath().c_str());	//force working directory to path of module, preventing lock

	if ((!isMultiInst) && (!TheFirstOne))
	{
		HWND hNotepad_plus = ::FindWindow(Notepad_plus_Window::getClassName(), NULL);
		for (int i = 0 ;!hNotepad_plus && i < 5 ; ++i)
		{
			Sleep(100);
			hNotepad_plus = ::FindWindow(Notepad_plus_Window::getClassName(), NULL);
		}

        if (hNotepad_plus)
        {
			// First of all, destroy static object NppParameters
			nppParameters.destroyInstance();

			// Restore the window, bring it to front, etc
			bool isInSystemTray = ::SendMessage(hNotepad_plus, NPPM_INTERNAL_RESTOREFROMTRAY, 0, 0);

			if (!isInSystemTray)
			{
				int sw = 0;

				if (::IsZoomed(hNotepad_plus))
					sw = SW_MAXIMIZE;
				else if (::IsIconic(hNotepad_plus))
					sw = SW_RESTORE;

				if (sw != 0)
					::ShowWindow(hNotepad_plus, sw);
			}
			::SetForegroundWindow(hNotepad_plus);
			return 0;
        }
	}

	auto upNotepadWindow = std::make_unique<Notepad_plus_Window>();
	Notepad_plus_Window & notepad_plus_plus = *upNotepadWindow.get();

	winVer ver = nppParameters.getWinVersion();
	std::wstring quotFileName;
	MSG msg{};
	msg.wParam = 0;
	Win32Exception::installHandler();
	MiniDumper mdump;	//for debugging purposes.
	try
	{
		notepad_plus_plus.init(hInstance, NULL, quotFileName.c_str(), &cmdLineParams);
		allowPrivilegeMessages(notepad_plus_plus, ver);
		bool going = true;
		while (going)
		{
			going = ::GetMessageW(&msg, NULL, 0, 0) != 0;
			if (going)
			{
				// if the message doesn't belong to the notepad_plus_plus's dialog
				if (!notepad_plus_plus.isDlgsMsg(&msg))
				{
					if (::TranslateAccelerator(notepad_plus_plus.getHSelf(), notepad_plus_plus.getAccTable(), &msg) == 0)
					{
						::TranslateMessage(&msg);
						::DispatchMessageW(&msg);
					}
				}
			}
		}
	}
	catch (int i)
	{
		TCHAR str[50] = TEXT("God Damned Exception : ");
		TCHAR code[10];
		wsprintf(code, TEXT("%d"), i);
		wcscat_s(str, code);
		::MessageBox(Notepad_plus_Window::gNppHWND, str, TEXT("Int Exception"), MB_OK);
		doException(notepad_plus_plus);
	}
	catch (std::runtime_error & ex)
	{
		::MessageBoxA(Notepad_plus_Window::gNppHWND, ex.what(), "Runtime Exception", MB_OK);
		doException(notepad_plus_plus);
	}
	catch (const Win32Exception & ex)
	{
		TCHAR message[1024];	//TODO: sane number
		wsprintf(message, TEXT("An exception occured. Notepad++ cannot recover and must be shut down.\r\nThe exception details are as follows:\r\n")
			TEXT("Code:\t0x%08X\r\nType:\t%S\r\nException address: 0x%p"), ex.code(), ex.what(), ex.where());
		::MessageBox(Notepad_plus_Window::gNppHWND, message, TEXT("Win32Exception"), MB_OK | MB_ICONERROR);
		mdump.writeDump(ex.info());
		doException(notepad_plus_plus);
	}
	catch (std::exception & ex)
	{
		::MessageBoxA(Notepad_plus_Window::gNppHWND, ex.what(), "General Exception", MB_OK);
		doException(notepad_plus_plus);
	}
	catch (...) // this shouldnt ever have to happen
	{
		::MessageBoxA(Notepad_plus_Window::gNppHWND, "An exception that we did not yet found its name is just caught", "Unknown Exception", MB_OK);
		doException(notepad_plus_plus);
	}

	return static_cast<int>(msg.wParam);
}
