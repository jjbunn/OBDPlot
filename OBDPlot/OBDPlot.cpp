// OBDPlot.cpp 
// (c) Julian Bunn, 2007-2013 
//
// v1.1a 29 July Pace time 100ms. Fix for 0.05 factor for millisecs instead of 5
// v1.1b 31 July Adjustments to scales. Correction for temperature calculation
// v1.1c 5 Aug Application properly terminates when user selects "Exit" from File menu
//				Increase pace time to 200ms
// v1.3  27 Sep 2008 Release version
// v1.3a 05 Jun 2012 Test version with factor 100 decrease in battery voltage constant
// v1.3b 06 Feb 2013 With baud rate 9600
//
// This source code is placed in the public domain.


#include "stdafx.h"
#include "OBDPlot.h"
#include "stdio.h"
#include "commdlg.h"
#include "commctrl.h"

#define MAX_LOADSTRING 200
#define PACE_TIME 0
//#define DCB_BAUDRATE 8800
#define DCB_BAUDRATE 9600

// Global Variables:
HINSTANCE hInst;								// current instance
TCHAR szTitle[MAX_LOADSTRING];					// The title bar text
TCHAR szWindowClass[MAX_LOADSTRING];			// the main window class name

// Forward declarations of functions included in this code module:
ATOM				MyRegisterClass(HINSTANCE hInstance);
BOOL				InitInstance(HINSTANCE, int);
LRESULT CALLBACK	WndProc(HWND, UINT, WPARAM, LPARAM);
INT_PTR CALLBACK	About(HWND, UINT, WPARAM, LPARAM);

BOOL InitInterface();
BOOL GetBlock(int);
void AppendText(LPCTSTR);
void SaveLog();
BOOL SendValueRequestBlock(int, char, char, char);
BOOL SendADCChannelRead(int,int);
BOOL SendAckBlock(int);
BOOL SendEndBlock(int);
BOOL SendBlockType(int,BYTE);

HWND				hwndParent = NULL;
HANDLE				hCom = NULL;
HWND				hwndDiagnostics = NULL;
HWND				hwndLog,hwndStart;
DCB					dcb;
HDC					hdcMem;
RECT				plotRect;
HANDLE				hLogFile;
static HPEN greenPen,redPen,pinkPen,whitePen,bluePen,orangePen,yellowPen,greyDashedPen,purplePen,greyPen;
static HBRUSH greenBrush,redBrush,pinkBrush,whiteBrush,blueBrush,orangeBrush,yellowBrush,purpleBrush,greyBrush,darkGreyBrush;
static DWORD LocaleID;
CRITICAL_SECTION criticalSection;


COLORREF colorGreen,colorRed,colorBlue,colorOrange,colorYellow,colorGrey,colorPurple,colorPink;

COMMTIMEOUTS		newCTO;
int					LenTimer = 500;
int					iWidth,iHeight;
BOOL				bInit = FALSE;
BOOL				bPaused = FALSE;
BOOL				bAnswerReceived = FALSE;
BOOL				bDebug = TRUE;
int					numRequests = 0;
int					border=250;
int					measuredBaudRate = 0;

int					blockNumber = 0;

#define MAXBLOCKS 256
BYTE				blockNum[MAXBLOCKS];
BYTE				blockType[MAXBLOCKS];
BYTE				blockData[MAXBLOCKS][MAX_LOADSTRING];
int					blockLength[MAXBLOCKS];

#define MAXBUF 512
char OutBuf[MAXBUF];
char InBuf[MAXBUF];
DWORD dLengthOfBlock, dWritten, dRead;

DWORD	ReaderThreadID = 0;
HANDLE	hReaderThread = NULL;
DWORD WINAPI ReaderLoop (LPVOID);

#define TEMPFILE _T("OBDPlot.log")

#define NPARAMETERS 8

static TCHAR szParameterName[NPARAMETERS][MAX_PATH] =
	{TEXT("Intake Air Temp"),		// ((n*115)/100)-26			deg F
	 TEXT("Cylinder Head Temp"),	// ((n*115)/100)-26			deg F
	 TEXT("AFM Voltage"),			// ((n*500)/255)			volts
	 TEXT("RPM"),					// n*40						rpm
	 TEXT("Injector Time"),			// n*.05					millisecs
	 TEXT("Ignition Advance"),		// (((n-0x68)*2075)/255)*-1	degrees
	 TEXT("MAF Sensor"),			// ((n*500)/255) volts
	 TEXT("Battery")};				// (n*682)/100 volts

static int iParameterType[NPARAMETERS] =
	{0,   // "Actual Value"
	 0,
	 0,
	 0,
	 0,
	 0,
	 1,   // ADC
	 1};

static short iParameterRequest[NPARAMETERS][3] =
	{0x01,0x00,0x37, // request bytes for parameters
	 0x01,0x00,0x38,
	 0x01,0x00,0x45,
	 0x01,0x00,0x3A,
	 0x01,0x00,0x42,
	 0x01,0x00,0x5D,
	 0x00,0x00,0x00, // first byte is ADC channel number, rest are unused
	 0x01,0x00,0x00};

static int dParameterID[NPARAMETERS] = {10101,10102,10103,10104,10105,10106,10107,10108};

BOOL bParameterSelected[NPARAMETERS] = {TRUE,TRUE,TRUE,TRUE,TRUE,TRUE,TRUE,TRUE};

static COLORREF cParameterColor[NPARAMETERS];
static HPEN hParameterPen[NPARAMETERS];
static HBRUSH hParameterBrush[NPARAMETERS];

#define MAX_RPM 9000.f
#define MAX_VOLTS 20.f
#define MAX_DEGREES 290.f
#define MAX_TEMP 450.f
#define MAX_TIME 90.f

#define MIN_RPM -1000.f
#define MIN_VOLTS -2.0f
#define MIN_DEGREES -10.f
#define MIN_TEMP -50.f
#define MIN_TIME -10.f

BOOL bLabelRPM,bLabelVolts,bLabelDegrees,bLabelTemp,bLabelTime;

#define MAX_SCREEN_WIDTH 1000
short sParameterValue[NPARAMETERS];
float fParameterHistory[NPARAMETERS][MAX_SCREEN_WIDTH];
long timeT[MAX_SCREEN_WIDTH];
static POINT ChartPoints[MAX_SCREEN_WIDTH];
TCHAR szTime[512];
int iParameter = 0;

float scaleFactorRPM,scaleFactorVolts,scaleFactorDegrees,scaleFactorTemp,scaleFactorTime;
int nHistory;
int penWidth = 10;

HWND hwndParameter[NPARAMETERS];

int APIENTRY _tWinMain(HINSTANCE hInstance,
                     HINSTANCE hPrevInstance,
                     LPTSTR    lpCmdLine,
                     int       nCmdShow)
{
	UNREFERENCED_PARAMETER(hPrevInstance);
	UNREFERENCED_PARAMETER(lpCmdLine);

 	// TODO: Place code here.
	MSG msg;
	HACCEL hAccelTable;

	// Initialize global strings
	LoadString(hInstance, IDS_APP_TITLE, szTitle, MAX_LOADSTRING);
	LoadString(hInstance, IDC_OBDPLOT, szWindowClass, MAX_LOADSTRING);
	MyRegisterClass(hInstance);

	// Perform application initialization:
	if (!InitInstance (hInstance, nCmdShow))
	{
		return FALSE;
	}

	hAccelTable = LoadAccelerators(hInstance, MAKEINTRESOURCE(IDC_OBDPLOT));

	// Main message loop:
	while (GetMessage(&msg, NULL, 0, 0))
	{
		if (!TranslateAccelerator(msg.hwnd, hAccelTable, &msg))
		{
			TranslateMessage(&msg);
			DispatchMessage(&msg);
		}
	}

	return (int) msg.wParam;
}



//
//  FUNCTION: MyRegisterClass()
//
//  PURPOSE: Registers the window class.
//
//  COMMENTS:
//
//    This function and its usage are only necessary if you want this code
//    to be compatible with Win32 systems prior to the 'RegisterClassEx'
//    function that was added to Windows 95. It is important to call this function
//    so that the application will get 'well formed' small icons associated
//    with it.
//
ATOM MyRegisterClass(HINSTANCE hInstance)
{
	WNDCLASSEX wcex;

		HICON		hIcon = LoadIcon(hInstance, MAKEINTRESOURCE(IDI_OBDPLOT));


	wcex.cbSize = sizeof(WNDCLASSEX);

	wcex.style			= CS_HREDRAW | CS_VREDRAW;
	wcex.lpfnWndProc	= WndProc;
	wcex.cbClsExtra		= 0;
	wcex.cbWndExtra		= 0;
	wcex.hInstance		= hInstance;
	wcex.hIcon			= LoadIcon(hInstance, MAKEINTRESOURCE(IDI_OBDPLOT));
	wcex.hCursor		= LoadCursor(NULL, IDC_ARROW);
	wcex.hbrBackground	= (HBRUSH)(COLOR_WINDOW+1);
	wcex.lpszMenuName	= MAKEINTRESOURCE(IDC_OBDPLOT);
	wcex.lpszClassName	= szWindowClass;
	wcex.hIconSm		= LoadIcon(wcex.hInstance, MAKEINTRESOURCE(IDI_SMALL));

	return RegisterClassEx(&wcex);
}

//
//   FUNCTION: InitInstance(HINSTANCE, int)
//
//   PURPOSE: Saves instance handle and creates main window
//
//   COMMENTS:
//
//        In this function, we save the instance handle in a global variable and
//        create and display the main program window.
//
BOOL InitInstance(HINSTANCE hInstance, int nCmdShow)
{
   HWND hWnd;
	TCHAR szPort[10];
	HANDLE hCfgFile;

   hInst = hInstance; // Store instance handle in our global variable

#define SCREENWIDTH 1000
#define SCREENHEIGHT 600

   hWnd = CreateWindow(szWindowClass, szTitle, WS_SYSMENU|WS_MINIMIZEBOX|WS_CAPTION,
      0, 0, SCREENWIDTH, SCREENHEIGHT, NULL, NULL, hInstance, NULL);

   if (!hWnd)
   {
      return FALSE;
   }

   ShowWindow(hWnd, nCmdShow);

	hwndParent = hWnd;

	DWORD langid = MAKELANGID(LANG_ENGLISH,SUBLANG_ENGLISH_US);
	LocaleID = MAKELCID(langid,SORT_DEFAULT);

	InitializeCriticalSection(&criticalSection);

	FillRect(hdcMem,&plotRect,(HBRUSH)GetStockObject(BLACK_BRUSH));
	SetTextColor(hdcMem,RGB(255,200,0));
	DrawText(hdcMem,_T("Starting"),8,&plotRect,DT_CENTER|DT_WORDBREAK);
	InvalidateRect(hwndParent,&plotRect,FALSE);

	UpdateWindow(hwndParent);

	hLogFile = CreateFile(TEMPFILE,
					GENERIC_WRITE,
					0,
					NULL,
					CREATE_ALWAYS,
					FILE_ATTRIBUTE_NORMAL,
					NULL);

	if(!hLogFile) {
		MessageBox(NULL,_T("Cannot open Log File"),_T("Error"),MB_OK);
		SendMessage (hWnd, WM_CLOSE, 0, 0);
		return FALSE;
	}

   AppendText(_T("Starting ...\r\n"));

	hCfgFile = CreateFile(_T("OBDPlot.cfg"),
					GENERIC_READ,
					0,
					NULL,
					OPEN_EXISTING,
					FILE_ATTRIBUTE_NORMAL,
					NULL);

	if(!hCfgFile) {
		AppendText(_T("Default port: "));
		wcscpy(szPort,_T("COM9:"));
	} else {
		DWORD dwRead;
		char port[MAX_LOADSTRING];
		// success ... read the key contained
		if(ReadFile(hCfgFile, (LPVOID) port, (DWORD) MAX_LOADSTRING, (LPDWORD) &dwRead, NULL)) {
			mbstowcs(szPort,port,dwRead);
			szPort[dwRead] = 0;
			AppendText(_T("Read Config file: "));
		}
		CloseHandle(hCfgFile);
	}

	AppendText(szPort);
	AppendText(_T("\r\n"));

	

   for(int i=0;i<MAXBLOCKS;i++) {blockType[i] = 0; blockLength[i] = 0;}
   for(int i=0;i<NPARAMETERS;i++) {
	   sParameterValue[i] = 0;
   }
   AppendText(_T("\r\n"));

	hCom = CreateFile(szPort,
		GENERIC_READ | GENERIC_WRITE,
		0,    // exclusive
		NULL, // default security attributes
		OPEN_EXISTING,
		//FILE_FLAG_OVERLAPPED,
		0, 
		NULL);

	if (hCom == INVALID_HANDLE_VALUE) {
		DWORD  nError= GetLastError();
		MessageBox(hWnd,_T("Cannot open COM port"),_T("Error"),MB_OK);
		AppendText( _T("\r\nCannot open COM port"));
		SendMessage (hWnd, WM_CLOSE, 0, 0);
		return FALSE;
	}
	AppendText( _T("COM Opened\r\n"));
	bInit = FALSE;

	SetTimer(hWnd,IDC_TIMER,LenTimer,NULL);

   return TRUE;
}

//
//  FUNCTION: WndProc(HWND, UINT, WPARAM, LPARAM)
//
//  PURPOSE:  Processes messages for the main window.
//
//  WM_COMMAND	- process the application menu
//  WM_PAINT	- Paint the main window
//  WM_DESTROY	- post a quit message and return
//
//
LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
{
	int wmId, wmEvent;
	PAINTSTRUCT ps;
	HDC hdc;
	DWORD dwStyle;
	INITCOMMONCONTROLSEX InitCtrls;
	long timestamp;
	TCHAR szText[512];
	char text[512];
	BYTE b;
	RECT r;
	HBITMAP hBitmap,holdBitmap;
	int h,step;
	POINT p[5];
	int plotWidth,labelcount;
	LPDRAWITEMSTRUCT lpdis;
	MENUITEMINFO mii;
	HMENU hMenu;
	DWORD itemdata;


	switch (message)
	{
	case WM_DRAWITEM:
		lpdis = (LPDRAWITEMSTRUCT) lParam; 
		{
			TCHAR wtext[100];
			HWND hwndC = lpdis->hwndItem;
			HDC hDC = lpdis->hDC;
			//HBRUSH hbr = (HBRUSH) SendMessage(hWnd,WM_CTLCOLORBTN, 0, (LPARAM)lpdis->hwndItem);
			HBRUSH hbr = (HBRUSH) GetStockObject(BLACK_BRUSH);
			RECT r = lpdis->rcItem;
			FillRect(hDC,&r,hbr);
			GetWindowText(hwndC,(LPTSTR) wtext, 100);

			int ib = 0;
			for(int i=0;i<NPARAMETERS;i++) {
				if(hwndC == hwndParameter[i]) {
					ib = i;
					if(!bParameterSelected[i]) {
						SelectObject(hDC,yellowPen);
						SelectObject(hDC,redBrush);
					} else {
						SelectObject(hDC,yellowPen);
						SelectObject(hDC,greenBrush);
					}
					Ellipse(hDC,r.left+2,r.top + 2,r.left+15,r.top+15);
					break;
				}
			}

			SetBkMode(hDC,TRANSPARENT);
			SetTextColor(hDC,cParameterColor[ib]);
			DrawText(hDC,(LPCTSTR)wtext,lstrlen((LPTSTR)wtext),&r,DT_CENTER|DT_SINGLELINE|DT_VCENTER);
			p[0].x = r.left;
			p[0].y = r.bottom-1;
			p[1].x = r.left;
			p[1].y = r.top;
			p[2].x = r.right-1;
			p[2].y = r.top;
			p[3].x = r.right-1;
			p[3].y = r.bottom-1;
			p[4].x = r.left;
			p[4].y = r.bottom-1;
			SelectObject(hDC,orangePen);
			Polyline(hDC,p,5);	

		}
		return TRUE;

	case WM_CTLCOLORBTN:
	case WM_CTLCOLOREDIT:
	case WM_CTLCOLORDLG:
	case WM_CTLCOLORSTATIC:
			
		//for(int i=0;i<NPARAMETERS;i++) {
		//	if((HWND)lParam == hwndParameter[i]) {
		//		SetTextColor((HDC)wParam,cParameterColor[i]);
		//		return (BOOL) GetStockObject(BLACK_BRUSH);
		//	}
		//}
		if((HWND)lParam == hwndStart) {
			SetTextColor((HDC)wParam,colorGreen);
			return (BOOL) GetStockObject(BLACK_BRUSH);
		}
		if((HWND)lParam == hwndDiagnostics) {
			SetTextColor((HDC)wParam,colorGrey);
			return (BOOL) GetStockObject(WHITE_BRUSH);
		}

		break;

	case WM_INITMENUPOPUP:

		hMenu = (HMENU) wParam;
	
		memset(&mii, 0,  sizeof(MENUITEMINFO));
		mii.cbSize = sizeof(MENUITEMINFO);
		mii.fMask = MIIM_ID | MIIM_SUBMENU;
		mii.cch = sizeof(itemdata);
		mii.dwItemData = (long)&itemdata;
		if(bDebug) {
			CheckMenuItem(hMenu,ID_FILE_DEBUG,MF_CHECKED);
		} else {
			CheckMenuItem(hMenu,ID_FILE_DEBUG,MF_UNCHECKED);
		}
		break;

	case WM_CREATE:

		colorGreen = RGB(0,255,0);
		colorRed = RGB(255,0,0);
		colorBlue = RGB(0,0,255);
		colorOrange = RGB(200,100,0);
		colorYellow = RGB(240,240,50);
		colorGrey = RGB(180,180,180);
		colorPurple = RGB(250,0,250);
		colorPink = RGB(255,200,200);

		greenPen = CreatePen(PS_SOLID,1,colorGreen);
		bluePen = CreatePen(PS_SOLID,1,colorBlue);
		greyPen = CreatePen(PS_SOLID,1,colorGrey);
		redPen = CreatePen(PS_SOLID,1,colorRed);
		whitePen = CreatePen(PS_SOLID,1,RGB(255,255,255));
		orangePen = CreatePen(PS_SOLID,1,colorOrange);
		yellowPen = CreatePen(PS_SOLID,1,colorYellow);
		purplePen = CreatePen(PS_SOLID,1,colorPurple);
		pinkPen = CreatePen(PS_SOLID,1,colorPink);
		greyDashedPen = CreatePen(PS_DASH,0,RGB(100,100,100));

		orangeBrush = CreateSolidBrush(colorOrange);
		greenBrush = CreateSolidBrush(colorGreen);
		redBrush = CreateSolidBrush(colorRed);
		yellowBrush = CreateSolidBrush(colorYellow);
		whiteBrush = CreateSolidBrush(RGB(255,255,255));
		blueBrush = CreateSolidBrush(colorBlue);
		greyBrush = CreateSolidBrush(colorGrey);
		purpleBrush = CreateSolidBrush(colorPurple);
		pinkBrush = CreateSolidBrush(colorPink);
		darkGreyBrush = CreateSolidBrush(RGB(40,40,40));

		hdcMem = CreateCompatibleDC(GetDC(hWnd));
		GetClientRect(hWnd,&r);
		iWidth = r.right - r.left;
		iHeight = r.bottom - r.top;
		plotRect = r;
		plotRect.top += 80;
		hBitmap = CreateCompatibleBitmap (GetDC(hWnd), iWidth, iHeight);
		holdBitmap = (HBITMAP) SelectObject (hdcMem, hBitmap);
		SetBkMode(hdcMem,TRANSPARENT);


		dwStyle = WS_CHILD | WS_VISIBLE | WS_BORDER | BS_OWNERDRAW;
		//hwndLog = CreateWindowEx(WS_EX_NOACTIVATE,_T("BUTTON"),_T("Save Log"),dwStyle,
		//		0,0,
		//		border,20,
		//		hWnd,(HMENU)IDC_SAVELOG,hInst,NULL);

		//ShowWindow(hwndLog,SW_SHOW);

		hwndStart = CreateWindowEx(WS_EX_NOACTIVATE,_T("BUTTON"),_T("Pause"),dwStyle,
				0,0,
				border,40,
				hWnd,(HMENU)IDC_REQUEST,hInst,NULL);

		ShowWindow(hwndStart,SW_SHOW);

		dwStyle = WS_CHILD | WS_VISIBLE | BS_OWNERDRAW;
		for(int i=0;i<NPARAMETERS;i++) {
			int row = i / 4;
			int col = i % 4;
			hwndParameter[i] = CreateWindowEx(WS_EX_NOACTIVATE,_T("BUTTON"),szParameterName[i],dwStyle,
				col*border,40+row*20,
				border,20,
				hWnd,(HMENU)dParameterID[i],hInst,NULL);
			ShowWindow(hwndParameter[i],SW_SHOW);
		}
		cParameterColor[0] = colorRed;
		hParameterPen[0] = redPen;
		hParameterBrush[0] = redBrush;
		cParameterColor[1] = colorOrange;
		hParameterPen[1] = orangePen;
		hParameterBrush[1] = orangeBrush;
		cParameterColor[2] = colorGreen;
		hParameterPen[2] = greenPen;
		hParameterBrush[2] = greenBrush;
		cParameterColor[3] = colorBlue;
		hParameterPen[3] = bluePen;
		hParameterBrush[3] = blueBrush;
		cParameterColor[4] = colorYellow;
		hParameterPen[4] = yellowPen;
		hParameterBrush[4] = yellowBrush;
		cParameterColor[5] = colorGrey;
		hParameterPen[5] = greyPen;
		hParameterBrush[5] = greyBrush;
		cParameterColor[6] = colorPurple;
		hParameterPen[6] = purplePen;
		hParameterBrush[6] = purpleBrush;
		cParameterColor[7] = colorPink;
		hParameterPen[7] = pinkPen;
		hParameterBrush[7] = pinkBrush;
		


		//dwStyle = WS_VISIBLE | WS_CHILD | WS_HSCROLL | WS_VSCROLL | WS_BORDER | ES_LEFT | ES_MULTILINE | ES_READONLY | ES_NOHIDESEL | ES_AUTOHSCROLL | ES_AUTOVSCROLL;
		dwStyle = WS_VISIBLE | WS_CHILD | ES_LEFT | ES_READONLY | ES_NOHIDESEL;

		h = 40 +NPARAMETERS*15;
		hwndDiagnostics = CreateWindowEx(WS_EX_NOACTIVATE, _T("edit"), NULL, dwStyle, 
			border, 0,
			iWidth-border, 40, 
			hWnd, NULL, hInst, NULL);
		ShowWindow(hwndDiagnostics,SW_SHOW);


		InitCtrls.dwSize = sizeof(INITCOMMONCONTROLSEX);
		InitCtrls.dwICC = ICC_UPDOWN_CLASS | ICC_WIN95_CLASSES;
		if(!InitCommonControlsEx(&InitCtrls)) {
		}


		break;
	case WM_TIMER:
		// Called every LenTimer millisecs

		if(!bInit) {

			if(!InitInterface()){
				SetWindowText(hwndDiagnostics,_T("Waiting to connect to ECU"));
				UpdateWindow(hwndDiagnostics);
				UpdateWindow(hwndParent);
				return TRUE;
			}
			//SetTimer(hWnd,IDC_TIMER,LenTimer,NULL);

			bInit = TRUE;

			int i = GetTimeFormat(LocaleID,LOCALE_NOUSEROVERRIDE, NULL, NULL, szTime, MAX_PATH);

			AppendText(szTime);
			for(int i=0;i<NPARAMETERS;i++) {
				AppendText(_T(","));
				AppendText(szParameterName[i]);
			}

		}
		timestamp = GetTickCount();

		if(bPaused) {
			wsprintf(szText,_T("Paused: SysTime %10dms Effective Baud Rate %d BlockNumber %3d History %5d"),timestamp,measuredBaudRate,blockNumber,nHistory);
		} else {
			wsprintf(szText,_T("Running: SysTime %10dms Effective Baud Rate %d BlockNumber %3d History %5d"),timestamp,measuredBaudRate,blockNumber,nHistory);
		}
		SetWindowText(hwndDiagnostics,szText);
		UpdateWindow(hwndDiagnostics);

		if(bPaused) break;

		FillRect(hdcMem,&plotRect,(HBRUSH)GetStockObject(BLACK_BRUSH));


		plotWidth = plotRect.right-plotRect.left+1;
		scaleFactorRPM = (float) (plotRect.bottom - plotRect.top) / (MAX_RPM-MIN_RPM);
		scaleFactorVolts = (float) (plotRect.bottom - plotRect.top) / (MAX_VOLTS-MIN_VOLTS);
		scaleFactorDegrees = (float) (plotRect.bottom - plotRect.top) / (MAX_DEGREES-MIN_DEGREES);
		scaleFactorTemp = (float) (plotRect.bottom - plotRect.top) / (MAX_TEMP-MIN_TEMP);
		scaleFactorTime = (float) (plotRect.bottom - plotRect.top) / (MAX_TIME-MIN_TIME);

		bLabelRPM = bLabelVolts = bLabelDegrees = bLabelTemp = bLabelTime = FALSE;


		nHistory++;
		if(nHistory >= plotWidth-1-penWidth) nHistory = plotWidth-1-penWidth;

		SelectObject(hdcMem,greyPen);
		//hOldFont = (HFONT) SelectObject(hdcMem,hButtonFont);

		step = (plotRect.bottom-plotRect.top)/10;
		for(int j=1;j<10;j++) {
			h = plotRect.bottom - step*j;
			p[0].x = plotRect.left;
			p[0].y = h;
			p[1].x = plotRect.right;
			p[1].y = h;
			Polyline(hdcMem,(POINT *)p,2);
			SelectObject(hdcMem,greyDashedPen);						
		}

		//i = GetDateFormat(LocaleID,DATE_LONGDATE, NULL, NULL, szDate, MAX_PATH);
		h = GetTimeFormat(LocaleID,LOCALE_NOUSEROVERRIDE, NULL, NULL, szTime, MAX_PATH);

		//wcscat(szDate,_T(" "));
		//wcscat(szDate,szTime);

		AppendText(szTime);
		AppendText(_T(" "));

		labelcount = 0;
		for(int j=0;j<NPARAMETERS;j++) {
			for(int i=0;i<plotWidth-1-penWidth;i++) {
				fParameterHistory[j][i] = fParameterHistory[j][i+1];
				if(j==0) timeT[i] = timeT[i+1];
			}
			if(!bParameterSelected[j]) continue;

			float value = (float)sParameterValue[j];
			float scale;
			float min;
			float max;
			HPEN pen = hParameterPen[j];
			HBRUSH brush = hParameterBrush[j];
			BOOL bAxislabel = FALSE;
			TCHAR label[20];
			if(j == 0 || j == 1) {
				value = value*1.15f-26.f; // degrees F
				scale = scaleFactorTemp;
				max = MAX_TEMP;
				min = MIN_TEMP;
				// this is daftness because wsprintf doesn't support %f!
				sprintf(text,"%5.2f F",value);
				if(!bLabelTemp) {
					bLabelTemp = TRUE;
					bAxislabel = TRUE;
					wcscpy(label,_T("Fahr"));
				}
			} else if (j == 2) {
				value = value*500.f/25500.f; // Volts
				scale = scaleFactorVolts;
				max = MAX_VOLTS;
				min = MIN_VOLTS;
				sprintf(text,"%5.2f Volts",value);
				bAxislabel = TRUE;
				bLabelVolts = TRUE;
				wcscpy(label,_T("Volts"));
			} else if (j == 3) {
				value *= 40.f;            // rpm
				scale = scaleFactorRPM;
				max = MAX_RPM;
				min = MIN_RPM;
				sprintf(text,"%d RPM",(int)value);
				bAxislabel = TRUE;
				bLabelRPM = TRUE;
				wcscpy(label,_T("RPM"));
			} else if (j == 4) {
				value *= 0.05f;             // millisecs
				scale = scaleFactorTime;
				max = MAX_TIME;
				min = MIN_TIME;
				sprintf(text,"%5.2f ms",value);
				bAxislabel = TRUE;
				bLabelTime = TRUE;
				wcscpy(label,_T("ms"));
			} else if (j == 5) {
				value = (104.-value)*207.5/255.; // degrees angle
				scale = scaleFactorDegrees;
				max = MAX_DEGREES;
				min = MIN_DEGREES;
				sprintf(text,"%5.2f deg",value);
				if(!bLabelDegrees) {
					bLabelDegrees = TRUE;
					bAxislabel = TRUE;
					wcscpy(label,_T("Angle"));
				}
			} else if(j == 6) {
				value = (value*500.f)/25500.f;
				scale = scaleFactorVolts;
				max = MAX_VOLTS;
				min = MIN_VOLTS;
				sprintf(text,"%5.2f Volts",value);
				if(!bLabelVolts) {
					bLabelVolts = TRUE;
					bAxislabel = TRUE;
					wcscpy(label,_T("Volts"));
				}
			} else if(j == 7) {
				value = (value*682.f)/1000000.f;
				scale = scaleFactorVolts;
				max = MAX_VOLTS;
				min = MIN_VOLTS;
				sprintf(text,"%5.2f Volts",value);
				if(!bLabelVolts) {
					bLabelVolts = TRUE;
					bAxislabel = TRUE;
					wcscpy(label,_T("Volts"));
				}
			}
			wsprintf(szText,_T("%s = %S"),szParameterName[j],text);
			SetWindowText(hwndParameter[j],szText);

			wsprintf(szText,_T("%S"),text);
			AppendText(szText);
			if(j != NPARAMETERS-1) AppendText(_T(","));

			fParameterHistory[j][plotWidth-1-penWidth] = value;
			timeT[plotWidth-1-penWidth] = timestamp;

			long lastVertical = timestamp;
			int nvertical = 0;
			int nc = 0;
			int i = plotWidth-1-penWidth;

			SelectObject(hdcMem,greyDashedPen);
			SetTextColor(hdcMem,colorGrey);
			
			while(i >= 0 && nc < nHistory) {
				ChartPoints[nc].x = plotRect.right - penWidth - nc;
				ChartPoints[nc].y = plotRect.bottom - (int) ((fParameterHistory[j][i]-min)*scale);
				if(ChartPoints[nc].y > plotRect.bottom) ChartPoints[nc].y = plotRect.bottom;
				if(ChartPoints[nc].y < plotRect.top) ChartPoints[nc].y = plotRect.top;
				if(lastVertical-timeT[i] > 30000) {
					lastVertical = timeT[i];
					nvertical++;
					p[0].x = ChartPoints[nc].x;
					p[0].y = plotRect.top;
					p[1].x = p[0].x;
					p[1].y = plotRect.bottom;
					Polyline(hdcMem,p,2);
					wsprintf((LPTSTR)text,_T("-%ds"),nvertical*30);
					ExtTextOut(hdcMem,p[0].x,p[1].y - 15,0,NULL,(LPCTSTR)text,lstrlen((LPTSTR)text),NULL);
				}
				i--;
				nc++;
			}

			SelectObject(hdcMem,pen);
			SelectObject(hdcMem,brush);
			Polyline(hdcMem,ChartPoints,nHistory);

			i = plotRect.bottom - (int) ((value-min)*scale);
			if(i > plotRect.bottom) i = plotRect.bottom;
			if(i < plotRect.top) i = plotRect.top;

			p[0].x = plotRect.right - penWidth;
			p[0].y = i;
			p[1].x = plotRect.right - 1;
			p[1].y = i - 5;
			if(p[1].y < plotRect.top) p[1].y = plotRect.top;
			p[2].x = plotRect.right - 1;
			p[2].y = i + 5;
			if(p[2].y > plotRect.bottom) p[2].y = plotRect.bottom;
			p[3]   = p[0];
			Polygon(hdcMem,(POINT *)p,4);


			if(bAxislabel) {
				SetTextColor(hdcMem,colorGrey);

				float vstep = (max-min)/10;
				for(int i=1;i<10;i++) {
					h = plotRect.bottom - step*i;
					float value = vstep*(float)i + min;							
					wsprintf((LPTSTR)text,_T("%d"),(int)value);
					ExtTextOut(hdcMem,plotRect.left+2+45*labelcount,h-15,0,NULL,(LPCTSTR)text,lstrlen((LPTSTR)text),NULL);
				}
				ExtTextOut(hdcMem,plotRect.left+2+45*labelcount,plotRect.top+15,0,NULL,(LPCTSTR)label,lstrlen((LPTSTR)label),NULL);
				labelcount++;
			}
		}

		AppendText(_T("\r\n"));


		InvalidateRect(hwndParent,&plotRect,FALSE);

		UpdateWindow(hwndParent);


		Sleep(50);

		break;
	case WM_COMMAND:
		wmId    = LOWORD(wParam);
		wmEvent = HIWORD(wParam);
		// Parse the menu selections:
		if(wmEvent == BN_CLICKED) {
			for(int i=0;i<NPARAMETERS;i++) {
				if(wmId == dParameterID[i]) {
					bParameterSelected[i] = !bParameterSelected[i];
					GetClientRect(hwndParameter[i],&r);
					InvalidateRect(hwndParameter[i],&r,FALSE);
					UpdateWindow(hwndParameter[i]);
					return TRUE;
				}
			}
		}

		switch (wmId)
		{
		case ID_FILE_DEBUG:
			bDebug = !bDebug;
			break;
		case IDC_SAVELOG:
			SaveLog();
			break;
		case IDC_REQUEST:
			if(bInit) {
				bPaused = !bPaused;
				if(bPaused) {
					SetWindowText(hwndStart,_T("Resume"));
					UpdateWindow(hwndStart);
				} else {
					SetWindowText(hwndStart,_T("Pause"));
					UpdateWindow(hwndStart);
				}
			} else {
				if(InitInterface()) {
					bInit = TRUE;
					SetWindowText(hwndStart,_T("Pause"));
					UpdateWindow(hwndStart);
					bPaused = FALSE;
				}
			}
			break;

		case IDM_ABOUT:
			DialogBox(hInst, MAKEINTRESOURCE(IDD_ABOUTBOX), hWnd, About);
			break;
		case IDM_EXIT:
			SendMessage(hWnd, WM_ACTIVATE, MAKEWPARAM(WA_INACTIVE, 0), (LPARAM)hWnd);
			SendMessage (hWnd, WM_CLOSE, 0, 0);
			DestroyWindow(hWnd);
			break;
		default:
			return DefWindowProc(hWnd, message, wParam, lParam);
		}
		break;
	case WM_PAINT:
		if(!hdcMem) break;
		hdc = BeginPaint(hWnd, &ps);
		BitBlt (hdc, 0, 0, iWidth, iHeight, hdcMem, 0, 0, SRCCOPY);
		EndPaint(hWnd, &ps);
		break;
	case WM_DESTROY:
		SendEndBlock(blockNumber);
		AppendText(_T("Session ended"));
		CloseHandle(hLogFile);
		//DeleteFile(TEMPFILE);
		CloseHandle(hCom);
		PostQuitMessage(0);
		break;
	default:
		return DefWindowProc(hWnd, message, wParam, lParam);
	}
	return 0;
}

void AppendText(LPCTSTR szText) {
	DWORD dwWritten;
	char text[MAX_LOADSTRING];
   //int ndx = GetWindowTextLength (hwndDiagnostics);
   //SetFocus (hwndDiagnostics);
   //SendMessage (hwndDiagnostics, EM_SETSEL, (WPARAM)ndx, (LPARAM)ndx);
	//SendMessage(hwndDiagnostics,EM_REPLACESEL,FALSE,(LPARAM) szText);
	//SendMessage(hwndDiagnostics,EM_SCROLLCARET,0,0L);

	//SetWindowText(hwndDiagnostics,szText);
	//UpdateWindow(hwndDiagnostics);

	wcstombs(text,szText,lstrlen(szText));
	text[lstrlen(szText)] = 0;

	EnterCriticalSection(&criticalSection);

	WriteFile(hLogFile, (LPCVOID) text, (DWORD) strlen(text), (LPDWORD) &dwWritten, NULL);

	LeaveCriticalSection(&criticalSection);

}

void SaveLog() {
	OPENFILENAME ofn;
	TCHAR szDesired[MAX_PATH];

	szDesired[0] = 0;
	memset( &(ofn), 0, sizeof(ofn));

	ofn.lStructSize   = sizeof(ofn);
	ofn.hwndOwner = hwndParent;
	ofn.lpstrFile = szDesired;
	ofn.nMaxFile = MAX_PATH;   
	ofn.lpstrInitialDir = NULL;
	ofn.lpstrFilter = TEXT("Log files (*.log)\0*.log\0");   
	ofn.lpstrTitle = TEXT("Save Log to File");
	ofn.Flags = OFN_HIDEREADONLY|OFN_PATHMUSTEXIST; 
	ofn.lpstrDefExt = TEXT("log");

	if(!GetSaveFileName(&ofn))  return;

	EnterCriticalSection(&criticalSection);

	CloseHandle(hLogFile);
	BOOL bFail = CopyFile(TEMPFILE,szDesired,FALSE);
	hLogFile = CreateFile(TEMPFILE,
					GENERIC_WRITE,
					0,
					NULL,
					OPEN_ALWAYS,
					FILE_ATTRIBUTE_TEMPORARY,
					NULL);
	SetFilePointer(hLogFile,0,0,FILE_END);

	LeaveCriticalSection(&criticalSection);

}

BOOL GetBlock(int iParameter) {
	// baud count in = 7 + 2*blocklen baud count out = 3 + blocklen
	TCHAR szText[500];
	int length = 0;
	BYTE type = 0;
	int i;
	if(!ReadFile(hCom,&InBuf,1,&dRead,NULL)){
		return FALSE;
	}
	length = (int)(InBuf[0]&0x00FF) - 3;
	OutBuf[0] = 0xFF - InBuf[0];
	WriteFile(hCom,&OutBuf,1,&dWritten,NULL);
	ReadFile(hCom,&InBuf,1,&dRead,NULL); // should be the same as OutBuf[0]
	ReadFile(hCom,&InBuf,1,&dRead,NULL); // blockNumber
	blockNum[blockNumber] = InBuf[0];
	OutBuf[0] = 0xFF - InBuf[0];
	WriteFile(hCom,&OutBuf,1,&dWritten,NULL);
	ReadFile(hCom,&InBuf,1,&dRead,NULL); // should be the same as OutBuf[0]
	ReadFile(hCom,&InBuf,1,&dRead,NULL); // blockType
	type = InBuf[0];
	OutBuf[0] = 0xFF - InBuf[0];
	WriteFile(hCom,&OutBuf,1,&dWritten,NULL);
	ReadFile(hCom,&InBuf,1,&dRead,NULL); // should be the same as OutBuf[0]

	if(type == 0x0A) {
		AppendText(_T("\r\nWrong title!"));
		return FALSE; // wrong title
	}

	blockType[blockNumber] = type;
	blockLength[blockNumber] = length;
	for(i=0;i<length;i++){
			ReadFile(hCom,&InBuf,1,&dRead,NULL);
			blockData[blockNumber][i] = InBuf[0];
			OutBuf[0] = 0xFF - InBuf[0];
			WriteFile(hCom,&OutBuf,1,&dWritten,NULL);
			ReadFile(hCom,&InBuf,1,&dRead,NULL); // should be the same as OutBuf[0]
	}
	blockData[blockNumber][length] = 0;

	sParameterValue[iParameter] = 0;
	if(type == 0xFE && length > 0) {
		// reply to Actual value request
		sParameterValue[iParameter] = blockData[blockNumber][0];
	} else if(type == 0xFB &&length > 1) {
		// reply to ADC channel request
		sParameterValue[iParameter] = blockData[blockNumber][1];
		sParameterValue[iParameter] |= blockData[blockNumber][0] << 8;
	}

	//wsprintf((LPTSTR)szText,TEXT("\r\nReceived Parameter = %d blockNumber %d block num read %d block len %d type %4X data %x"),
	//	iParameter,blockNumber,blockNum[blockNumber],length,type,sParameterValue[iParameter]);
	//AppendText( szText);

	ReadFile(hCom,&InBuf,1,&dRead,NULL); // should be end of block 0x03, not echoed

}

BOOL SendAckBlock(int blocknum) {
	// baud count in = 7 baud count out = 4
	OutBuf[0] = 0x03;
	WriteFile(hCom,&OutBuf,1,&dWritten,NULL);
	ReadFile(hCom,&InBuf,1,&dRead,NULL); // should be the same 
	ReadFile(hCom,&InBuf,1,&dRead,NULL); // should be 0xFC
	OutBuf[0] = (BYTE) (blocknum & 0xFF);
	WriteFile(hCom,&OutBuf,1,&dWritten,NULL);
	ReadFile(hCom,&InBuf,1,&dRead,NULL); // should be the same
	if(InBuf[0] != OutBuf[0]) {
		return FALSE;
	}
	ReadFile(hCom,&InBuf,1,&dRead,NULL); // should be 0xFF-blockNumber i.e. blockNumber complement
	OutBuf[0] = 0x09; // ACK Block identifier
	WriteFile(hCom,&OutBuf,1,&dWritten,NULL);
	ReadFile(hCom,&InBuf,1,&dRead,NULL); // should be the same
	ReadFile(hCom,&InBuf,1,&dRead,NULL); // should be 0xF6 i.e. 0xFF-0x09
	OutBuf[0] = 0x03; // end block
	WriteFile(hCom,&OutBuf,1,&dWritten,NULL);
	ReadFile(hCom,&InBuf,1,&dRead,NULL); // should be the same
}

BOOL SendEndBlock(int blocknum) {
	OutBuf[0] = 0x03;
	WriteFile(hCom,&OutBuf,1,&dWritten,NULL);
	ReadFile(hCom,&InBuf,1,&dRead,NULL); // should be the same 
	ReadFile(hCom,&InBuf,1,&dRead,NULL); // should be 0xFC
	OutBuf[0] = (BYTE) (blocknum & 0xFF);
	WriteFile(hCom,&OutBuf,1,&dWritten,NULL);
	ReadFile(hCom,&InBuf,1,&dRead,NULL); // should be the same
	if(InBuf[0] != OutBuf[0]) {
		return FALSE;
	}
	ReadFile(hCom,&InBuf,1,&dRead,NULL); // should be 0xFF-blockNumber i.e. blockNumber complement
	OutBuf[0] = 0x06; // END DIAGNOSIS Block identifier
	WriteFile(hCom,&OutBuf,1,&dWritten,NULL);
	ReadFile(hCom,&InBuf,1,&dRead,NULL); // should be the same
	ReadFile(hCom,&InBuf,1,&dRead,NULL); // should be 0xF6 i.e. 0xFF-0x09
	OutBuf[0] = 0x03; // end block
	WriteFile(hCom,&OutBuf,1,&dWritten,NULL);
	ReadFile(hCom,&InBuf,1,&dRead,NULL); // should be the same
}

BOOL SendBlockType(int blocknum,BYTE type) {
	// baud count in = 7 baud count out = 4
	OutBuf[0] = 0x03;
	WriteFile(hCom,&OutBuf,1,&dWritten,NULL);
	ReadFile(hCom,&InBuf,1,&dRead,NULL); // should be the same 
	ReadFile(hCom,&InBuf,1,&dRead,NULL); // should be 0xFC
	OutBuf[0] = (BYTE) (blocknum & 0xFF);
	WriteFile(hCom,&OutBuf,1,&dWritten,NULL);
	ReadFile(hCom,&InBuf,1,&dRead,NULL); // should be the same
	if(InBuf[0] != OutBuf[0]) {
		return FALSE;
	}
	ReadFile(hCom,&InBuf,1,&dRead,NULL); // should be 0xFF-blockNumber i.e. blockNumber complement
	OutBuf[0] = type; // Block identifier
	WriteFile(hCom,&OutBuf,1,&dWritten,NULL);
	ReadFile(hCom,&InBuf,1,&dRead,NULL); // should be the same
	ReadFile(hCom,&InBuf,1,&dRead,NULL); // should be 0xF6 i.e. 0xFF-0x09
	OutBuf[0] = 0x03; // end block
	WriteFile(hCom,&OutBuf,1,&dWritten,NULL);
	ReadFile(hCom,&InBuf,1,&dRead,NULL); // should be the same
}

BOOL SendValueRequestBlock(int blockNum, char value1, char value2, char value3) {
	// baud count in = 13 baud count out = 7
	OutBuf[0] = 0x06;
	WriteFile(hCom,&OutBuf,1,&dWritten,NULL);
	ReadFile(hCom,&InBuf,1,&dRead,NULL); // should be the same 
	ReadFile(hCom,&InBuf,1,&dRead,NULL); // should be 0xFC
	OutBuf[0] = (BYTE) (blockNum & 0xFF);
	WriteFile(hCom,&OutBuf,1,&dWritten,NULL);
	ReadFile(hCom,&InBuf,1,&dRead,NULL); // should be the same
	if(InBuf[0] != OutBuf[0]) return FALSE;
	ReadFile(hCom,&InBuf,1,&dRead,NULL); // should be 0xFF-blockNumber i.e. blockNumber complement
	OutBuf[0] = 0x01; // Value request Block identifier
	WriteFile(hCom,&OutBuf,1,&dWritten,NULL);
	ReadFile(hCom,&InBuf,1,&dRead,NULL); // should be the same
	ReadFile(hCom,&InBuf,1,&dRead,NULL);

	OutBuf[0] = value1;
	WriteFile(hCom,&OutBuf,1,&dWritten,NULL);
	ReadFile(hCom,&InBuf,1,&dRead,NULL); // should be the same
	ReadFile(hCom,&InBuf,1,&dRead,NULL);

	OutBuf[0] = value2;
	WriteFile(hCom,&OutBuf,1,&dWritten,NULL);
	ReadFile(hCom,&InBuf,1,&dRead,NULL); // should be the same
	ReadFile(hCom,&InBuf,1,&dRead,NULL);

	OutBuf[0] = value3;
	WriteFile(hCom,&OutBuf,1,&dWritten,NULL);
	ReadFile(hCom,&InBuf,1,&dRead,NULL); // should be the same
	ReadFile(hCom,&InBuf,1,&dRead,NULL);

	OutBuf[0] = 0x03; // end block
	WriteFile(hCom,&OutBuf,1,&dWritten,NULL);
	ReadFile(hCom,&InBuf,1,&dRead,NULL); // should be the same
}

BOOL SendADCChannelRead(int blockNum, int channel) {
	// baud count in = 9 baud count out = 5
	OutBuf[0] = 0x04; // length of block
	WriteFile(hCom,&OutBuf,1,&dWritten,NULL);
	ReadFile(hCom,&InBuf,1,&dRead,NULL); // should be the same 
	ReadFile(hCom,&InBuf,1,&dRead,NULL); // should be 0xFC
	OutBuf[0] = (BYTE) (blockNum & 0xFF);
	WriteFile(hCom,&OutBuf,1,&dWritten,NULL);
	ReadFile(hCom,&InBuf,1,&dRead,NULL); // should be the same
	if(InBuf[0] != OutBuf[0]) return FALSE;
	ReadFile(hCom,&InBuf,1,&dRead,NULL); // should be 0xFF-blockNumber i.e. blockNumber complement
	OutBuf[0] = 0x08; // ADC Channel read request Block identifier
	WriteFile(hCom,&OutBuf,1,&dWritten,NULL);
	ReadFile(hCom,&InBuf,1,&dRead,NULL); // should be the same
	ReadFile(hCom,&InBuf,1,&dRead,NULL);

	OutBuf[0] = (BYTE)(channel & 0xFF); // ADC channel number
	WriteFile(hCom,&OutBuf,1,&dWritten,NULL);
	ReadFile(hCom,&InBuf,1,&dRead,NULL); // should be the same
	ReadFile(hCom,&InBuf,1,&dRead,NULL);

	OutBuf[0] = 0x03; // end block
	WriteFile(hCom,&OutBuf,1,&dWritten,NULL);
	ReadFile(hCom,&InBuf,1,&dRead,NULL); // should be the same
}




// Message handler for about box.
INT_PTR CALLBACK About(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam)
{
	UNREFERENCED_PARAMETER(lParam);
	switch (message)
	{
	case WM_INITDIALOG:
		return (INT_PTR)TRUE;

	case WM_COMMAND:
		if (LOWORD(wParam) == IDOK || LOWORD(wParam) == IDCANCEL)
		{
			EndDialog(hDlg, LOWORD(wParam));
			return (INT_PTR)TRUE;
		}
		break;
	}
	return (INT_PTR)FALSE;
}

BOOL InitInterface() {
	TCHAR szText[500];

			nHistory = 0;

			GetCommTimeouts(hCom,&newCTO);

			newCTO.ReadIntervalTimeout=MAXDWORD;
			newCTO.ReadTotalTimeoutMultiplier=MAXDWORD;
			newCTO.ReadTotalTimeoutConstant=1500;
			newCTO.WriteTotalTimeoutMultiplier=0;
			newCTO.WriteTotalTimeoutConstant=0;

			if(!SetCommTimeouts(hCom,&newCTO)) {
				AppendText( _T("\r\nCannot Set Timeouts"));
				//MessageBox(hwndParent,_T("Cannot set comm timeouts"),_T("Error"),MB_OK);
				CloseHandle(hCom);
				return FALSE;
			}

			GetCommState(hCom,&dcb);

			dcb.DCBlength  = sizeof( dcb );
			dcb.BaudRate  = 9600;
			dcb.ByteSize  = 8;
			dcb.Parity   = FALSE;
			dcb.StopBits  = ONESTOPBIT;
			dcb.fOutxCtsFlow = 0;
			dcb.fOutxDsrFlow = 0;
			dcb.fDtrControl  = DTR_CONTROL_ENABLE;//DTR_CONTROL_DISABLE;
			dcb.fDsrSensitivity = 0;
			dcb.fOutX   = 0;
			dcb.fInX   = 0;
			dcb.fErrorChar  = 0;
			dcb.fRtsControl  = RTS_CONTROL_DISABLE;
			dcb.fNull   = 0;
			dcb.XonChar   = 0;
			dcb.XoffChar  = 0;
			dcb.ErrorChar  = 0;
			dcb.XonLim = 0;
			dcb.XoffLim = 0;

			if (SetCommState( hCom, &dcb) == 0) {
				//MessageBox(hwndParent,_T("Cannot set comm state"),_T("Error"),MB_OK);
				AppendText( _T("\r\nCannot Set State"));
				CloseHandle(hCom);
				return FALSE;
			}

			dcb.BaudRate  = DCB_BAUDRATE;
			
			if (SetCommState( hCom, &dcb) == 0) {
				//MessageBox(hwndParent,_T("Cannot set comm state"),_T("Error"),MB_OK);
				AppendText( _T("\r\nCannot Set State"));
				CloseHandle(hCom);
				return FALSE;
			}


			EscapeCommFunction(hCom,SETRTS);
			Sleep(200);
			EscapeCommFunction(hCom,SETRTS);
			Sleep(200);
			EscapeCommFunction(hCom,SETRTS);
			Sleep(200);
			EscapeCommFunction(hCom,SETRTS);
			Sleep(200);
			EscapeCommFunction(hCom,SETRTS);
			Sleep(200);
			EscapeCommFunction(hCom,CLRRTS);
			Sleep(200);
			EscapeCommFunction(hCom,SETRTS);
			Sleep(200);
			EscapeCommFunction(hCom,SETRTS);
			Sleep(200);
			EscapeCommFunction(hCom,SETRTS);
			Sleep(200);
			EscapeCommFunction(hCom,CLRRTS);
			Sleep(200);

			PurgeComm(hCom,PURGE_RXCLEAR|PURGE_TXCLEAR);

			if(!ReadFile(hCom,&InBuf,1,&dRead,NULL)){
				AppendText( _T("\r\nDevice not responding"));
				//MessageBox(NULL,TEXT("Device not responding..."),TEXT("ERROR"),NULL);
				return FALSE;

			}

			if (InBuf[0]!= 0x55) {
				wsprintf((LPTSTR)szText,TEXT("\r\nIncorrect baudrate = %2.2Xh"),InBuf[0]);
				AppendText( szText);
				return FALSE;
			}

			if(!ReadFile(hCom,&InBuf,1,&dRead,NULL)) {
				AppendText( _T("\r\nDevice not responding"));
				return FALSE;
			}

			if (InBuf[0]!= 0x0B){
				AppendText( _T("\r\nAnswer not 0x0B"));
				//MessageBox(NULL,TEXT("Incorrect answer..."),TEXT("ERROR"),NULL);
				return FALSE;

			}

			if(!ReadFile(hCom,&InBuf,1,&dRead,NULL)) {
				AppendText( _T("\r\nDevice not responding"));
				//MessageBox(NULL,TEXT("Device not responding..."),TEXT("ERROR"),NULL);
				return FALSE;
			}

			if (InBuf[0]!= 0x02){
				AppendText( _T("\r\nAnswer not 0x02"));
				//MessageBox(NULL,TEXT("Incorrect answer..."),TEXT("ERROR"),NULL);
				return FALSE;
			}

			//Start diagnostic session

			OutBuf[0]=0xFD;

			Sleep(10);//0);

			WriteFile(hCom,&OutBuf,1,&dWritten,NULL);

			ReadFile(hCom,&InBuf,1,&dRead,NULL);

			if (InBuf[0]!=OutBuf[0]){
				AppendText( _T("\r\nIncorrect echo"));
				//MessageBox(NULL,TEXT("Incorrect echo..."),TEXT("ERROR"),NULL);
				return FALSE;
			}

			blockNumber = 1;

			if(!GetBlock(0)) {
				MessageBox(NULL,_T("GetBlock failed"),_T("Error"),MB_OK);
				return FALSE;
			}
			wsprintf((LPTSTR)szText,TEXT("\r\nblockNumber %d Type = %2X Length = %d data %S"),blockNumber,blockType[blockNumber],blockLength[blockNumber],blockData[blockNumber]);
			AppendText( szText);

			blockNumber = 2;
			
			if(!SendAckBlock(blockNumber)) {
				MessageBox(NULL,_T("SendAckBlock failed"),_T("Error"),MB_OK);
				return FALSE;
			}

			blockNumber = 3;

			if(!GetBlock(0)) {
				MessageBox(NULL,_T("GetBlock failed"),_T("Error"),MB_OK);
				return FALSE;
			}
			wsprintf((LPTSTR)szText,TEXT("\r\nblockNumber %d Type = %2X Length = %d data %S"),blockNumber,blockType[blockNumber],blockLength[blockNumber],blockData[blockNumber]);
			AppendText( szText);

			blockNumber = 4;

			if(!SendAckBlock(blockNumber)) {
				MessageBox(NULL,_T("SendAckBlock failed"),_T("Error"),MB_OK);
				return FALSE;
			}

			blockNumber = 5;

			if(!GetBlock(0)) {
				MessageBox(NULL,_T("GetBlock failed"),_T("Error"),MB_OK);
				return FALSE;
			}
			wsprintf((LPTSTR)szText,TEXT("\r\nblockNumber %d Type = %2X Length = %d data %S"),blockNumber,blockType[blockNumber],blockLength[blockNumber],blockData[blockNumber]);
			AppendText( szText);

			blockNumber = 6;

			if(!SendAckBlock(blockNumber)) {
				MessageBox(NULL,_T("SendAckBlock failed"),_T("Error"),MB_OK);
				return FALSE;
			}

			blockNumber = 7;

			if(!GetBlock(0)) {
				MessageBox(NULL,_T("GetBlock failed"),_T("Error"),MB_OK);
				return FALSE;
			}
			wsprintf((LPTSTR)szText,TEXT("\r\nblockNumber %d Type = %2X Length = %d data %S"),blockNumber,blockType[blockNumber],blockLength[blockNumber],blockData[blockNumber]);
			AppendText( szText);

			blockNumber = 8;

			if(!SendAckBlock(blockNumber)) {
				MessageBox(NULL,_T("SendAckBlock failed"),_T("Error"),MB_OK);
				return FALSE;
			}

			blockNumber = 9;

			// should be type F7
			if(!GetBlock(0)) {
				MessageBox(NULL,_T("GetBlock failed"),_T("Error"),MB_OK);
				return FALSE;
			}
			wsprintf((LPTSTR)szText,TEXT("\r\nblockNumber %d Type = %2X Length = %d data %S"),blockNumber,blockType[blockNumber],blockLength[blockNumber],blockData[blockNumber]);
			AppendText( szText);

			blockNumber = 10;

			// Send error byte request
			if(!SendBlockType(blockNumber,0x07)) {
				MessageBox(NULL,_T("SendBlockType failed"),_T("Error"),MB_OK);
				return FALSE;
			}

			blockNumber = 11;

			// should be type FC
			if(!GetBlock(0)) {
				MessageBox(NULL,_T("GetBlock failed"),_T("Error"),MB_OK);
				return FALSE;
			}
			wsprintf((LPTSTR)szText,TEXT("\r\nblockNumber %d Type = %2X Length = %d data %S"),blockNumber,blockType[blockNumber],blockLength[blockNumber],blockData[blockNumber]);
			AppendText( szText);

			blockNumber = 12;

			if(!SendAckBlock(blockNumber)) {
				MessageBox(NULL,_T("SendAckBlock failed"),_T("Error"),MB_OK);
				return FALSE;
			}

			blockNumber = 13;

			// should be type F7
			if(!GetBlock(0)) {
				MessageBox(NULL,_T("GetBlock failed"),_T("Error"),MB_OK);
				return FALSE;
			}
			wsprintf((LPTSTR)szText,TEXT("\r\nblockNumber %d Type = %2X Length = %d data %S"),blockNumber,blockType[blockNumber],blockLength[blockNumber],blockData[blockNumber]);
			AppendText( szText);

			// OK
			if(!hReaderThread) {
				hReaderThread = CreateThread (NULL, 0, ReaderLoop, 0, CREATE_SUSPENDED, &ReaderThreadID);	
				if(hReaderThread) {
					SetThreadPriority(hReaderThread,THREAD_PRIORITY_NORMAL);
					ResumeThread(hReaderThread);
					AppendText(_T("\r\nReader Thread Started"));
				} else {
					MessageBox(hwndParent,TEXT("Error Creating Reader Thread"),_T("Error!"),MB_OK);
					SendMessage (hwndParent, WM_CLOSE, 0, 0);
					return FALSE;
				}
			}

			// InitInterface ends with a getblock from the ecu
			AppendText( _T("\r\nInitialised OK!\r\n"));
			return TRUE;
			
}

DWORD WINAPI ReaderLoop (LPVOID pvarg) {
	TCHAR szText[500];
	long timestart,timestart0,timenow;
	int baudCount;
	while(TRUE) {

		if(!bInit) break;
		timestart = timestart0 = GetTickCount();

		baudCount = 0;

		blockNumber++;
		if(blockNumber > 255) blockNumber = 0;

		if(bDebug) {
			wsprintf((LPTSTR)szText,TEXT("\r\nRequest Block %d Parameter = %d vals %x %x %x\r\n"),blockNumber,iParameter,
				iParameterRequest[iParameter][0],
				iParameterRequest[iParameter][1],
				iParameterRequest[iParameter][2]);
			AppendText( szText);

		}

		if(iParameterType[iParameter] == 0) { // Actual Value
			if(!SendValueRequestBlock(blockNumber,
				iParameterRequest[iParameter][0],
				iParameterRequest[iParameter][1],
				iParameterRequest[iParameter][2])) {
				MessageBox(hwndParent,_T("SendValueRequestBlock failed"),_T("Error"),MB_OK);
				bInit = FALSE;
				break;
			}
			baudCount += 20;
		} else if(iParameterType[iParameter] == 1) { // ADC Channel
			if(!SendADCChannelRead(blockNumber,iParameterRequest[iParameter][0])) {
				MessageBox(hwndParent,_T("SendADCChannelRead failed"),_T("Error"),MB_OK);
				bInit = FALSE;
				break;
			}
			baudCount += 14;
		}

		// we pace the conversation so that each block lasts around 200ms
		timenow = GetTickCount();
		if(timenow-timestart < PACE_TIME) {
			Sleep(PACE_TIME - timenow + timestart);
		}
		timestart = GetTickCount();
		blockNumber++;
		if(blockNumber > 255) blockNumber = 0;
		if(!GetBlock(iParameter)) {
			MessageBox(hwndParent,_T("GetBlock failed"),_T("Error"),MB_OK);
			bInit = FALSE;
			break;
		}
			// baud count in = 7 + 2*blocklen baud count out = 3 + blocklen

		baudCount += 10 + 3*blockLength[blockNumber];

		if(bDebug) {
			wsprintf((LPTSTR)szText,TEXT("\r\nGetBlock returned block %d blockNumber %d Type %2X\r\n"),blockNum[blockNumber],blockNumber,blockType[blockNumber]);
			AppendText( szText);
		}

		iParameter++;
		if(iParameter >= NPARAMETERS) iParameter = 0;

		// we pace the conversation so that each block lasts around 200ms
		timenow = GetTickCount();
		if(timenow-timestart < PACE_TIME) {
			Sleep(PACE_TIME - timenow + timestart);
		}

		measuredBaudRate = (int) (8000.f * (float)baudCount / (float) (GetTickCount()-timestart0));

		if(blockNum[blockNumber] != blockNumber) bInit = FALSE;
	}

	return 0;
}