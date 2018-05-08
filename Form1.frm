VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hosts Editor - PC-DOS Workshop"
   ClientHeight    =   7170
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   7140
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7170
   ScaleWidth      =   7140
   StartUpPosition =   3  '窗口缺省
   Begin VB.ListBox lstXPHOSTS 
      Height          =   3660
      ItemData        =   "Form1.frx":0ECA
      Left            =   1815
      List            =   "Form1.frx":0F07
      TabIndex        =   13
      Top             =   7005
      Visible         =   0   'False
      Width           =   6720
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2970
      Top             =   3345
   End
   Begin VB.Frame Frame1 
      Caption         =   "當前Hosts域名解析表"
      Height          =   5820
      Left            =   60
      TabIndex        =   9
      Top             =   1290
      Width           =   7005
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00FFFFFF&
         Height          =   2310
         Left            =   855
         ScaleHeight     =   2250
         ScaleWidth      =   5430
         TabIndex        =   11
         Top             =   1140
         Visible         =   0   'False
         Width           =   5490
         Begin VB.Image Image1 
            Height          =   720
            Left            =   120
            Picture         =   "Form1.frx":11E9
            Top             =   525
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "正在加載文件%File%請稍候..."
            Height          =   1380
            Left            =   1140
            TabIndex        =   12
            Top             =   540
            Visible         =   0   'False
            Width           =   4110
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00FF0000&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00FF0000&
            Height          =   330
            Left            =   0
            Top             =   0
            Visible         =   0   'False
            Width           =   5520
         End
      End
      Begin VB.ListBox List1 
         Height          =   5460
         ItemData        =   "Form1.frx":2D2B
         Left            =   120
         List            =   "Form1.frx":2D2D
         TabIndex        =   10
         Top             =   255
         Width           =   6750
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   1140
      Left            =   75
      ScaleHeight     =   1080
      ScaleWidth      =   6960
      TabIndex        =   0
      Top             =   90
      Width           =   7020
      Begin VB.CommandButton Command10 
         Caption         =   "編輯(&E)..."
         Height          =   1080
         Left            =   5880
         MaskColor       =   &H00000000&
         Picture         =   "Form1.frx":2D2F
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   1080
      End
      Begin VB.CommandButton Command8 
         Caption         =   "刪除(&D)..."
         Height          =   1080
         Left            =   4800
         MaskColor       =   &H00000000&
         Picture         =   "Form1.frx":5171
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   1080
      End
      Begin VB.CommandButton Command6 
         Caption         =   "插入(&I)..."
         Height          =   1080
         Left            =   3420
         MaskColor       =   &H00000000&
         Picture         =   "Form1.frx":75B3
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   1080
      End
      Begin VB.CommandButton Command5 
         Height          =   1080
         Left            =   4500
         MaskColor       =   &H00FFFFFF&
         Picture         =   "Form1.frx":99F5
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   300
      End
      Begin VB.CommandButton Command4 
         Caption         =   "保存(&S)..."
         Height          =   1080
         Left            =   1380
         Picture         =   "Form1.frx":9C3F
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   0
         Width           =   1080
      End
      Begin VB.CommandButton Command3 
         Height          =   1080
         Left            =   2460
         MaskColor       =   &H00FFFFFF&
         Picture         =   "Form1.frx":AB09
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   300
      End
      Begin VB.CommandButton Command2 
         Height          =   1080
         Left            =   1080
         MaskColor       =   &H00FFFFFF&
         Picture         =   "Form1.frx":AD53
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   300
      End
      Begin VB.CommandButton Command1 
         Caption         =   "打開(&O)..."
         Height          =   1080
         Left            =   0
         Picture         =   "Form1.frx":AF9D
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   0
         Width           =   1080
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "文件(&F)"
      Begin VB.Menu mnuDefault 
         Caption         =   "打開默認位置Hosts文件(&D)"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuCustom 
         Caption         =   "打開自定義位置Hosts文件(&C)..."
         Shortcut        =   ^O
      End
      Begin VB.Menu b1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "保存(&S)"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "保存為(&A)..."
      End
      Begin VB.Menu b9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAutoSave 
         Caption         =   "自動保存選項(&T)..."
      End
      Begin VB.Menu b2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSE 
         Caption         =   "保存並退出(&V)"
      End
      Begin VB.Menu mnuDE 
         Caption         =   "不保存退出(&O)"
      End
      Begin VB.Menu mnuRepair 
         Caption         =   "修復損壞的Hosts文檔(&R)"
      End
      Begin VB.Menu b5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "退出(&E)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "編輯(&E)"
      Begin VB.Menu mnuCopy 
         Caption         =   "複製選定條目(&C)"
         Shortcut        =   ^C
      End
      Begin VB.Menu b6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInsertDNS 
         Caption         =   "插入連接定向條目(&I)..."
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuDisallow 
         Caption         =   "插入禁止存取條目(&D)..."
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuCm 
         Caption         =   "插入註釋條目(&C)..."
         Shortcut        =   ^M
      End
      Begin VB.Menu b3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "刪除選定條目(&L)"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuDeleteAll 
         Caption         =   "刪除所有條目(&A)"
      End
      Begin VB.Menu b8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuES 
         Caption         =   "編輯選定條目(&E)..."
         Shortcut        =   ^E
      End
      Begin VB.Menu b7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewCurrent 
         Caption         =   "查看選定條目(&V)..."
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "幫助(&H)"
      Begin VB.Menu mnuAbout 
         Caption         =   "關於(&A)..."
      End
   End
   Begin VB.Menu mnuIOpen 
      Caption         =   "mnuOpen"
      Visible         =   0   'False
      Begin VB.Menu mnuPDefault 
         Caption         =   "打開默認位置Hosts文件(&D)..."
      End
      Begin VB.Menu mnuPCustom 
         Caption         =   "打開自定義位置Hosts文件(&C)..."
      End
   End
   Begin VB.Menu mnuISave 
      Caption         =   "mnuSave"
      Visible         =   0   'False
      Begin VB.Menu mnuPSave 
         Caption         =   "保存(&S)"
      End
      Begin VB.Menu mnuPSaveAs 
         Caption         =   "保存為(&A)..."
      End
   End
   Begin VB.Menu mnuIInsert 
      Caption         =   "mnuInsert"
      Visible         =   0   'False
      Begin VB.Menu mnuPInsertDNS 
         Caption         =   "插入連接定向條目(&I)..."
      End
      Begin VB.Menu mnuPDisallow 
         Caption         =   "插入禁止存取條目(&D)..."
      End
      Begin VB.Menu mnuPCM 
         Caption         =   "插入註釋條目(&C)..."
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lpAST As Long
Dim lpHosts As String
 Private Type PROCESSENTRY32
 dwSize As Long
 cntUsage As Long
 th32ProcessID As Long '进程ID
 th32DefaultHeapID As Long '堆栈ID
 th32ModuleID As Long '模块ID
 cntThreads As Long
 th32ParentProcessID As Long '父进程ID
 pcPriClassBase As Long
 dwFlags As Long
 szExeFile As String * 260
 End Type
 Private Type MEMORYSTATUS
 dwLength As Long
 dwMemoryLoad As Long
 dwTotalPhys As Long
 dwAvailPhys As Long
 dwTotalPageFile As Long
 dwAvailPageFile As Long
 dwTotalVirtual As Long
 dwAvailVirtual As Long
 End Type
 Private Declare Function NtQuerySystemInformation Lib "ntdll" (ByVal dwInfoType As Long, ByVal lpStructure As Long, ByVal dwSize As Long, ByVal dwReserved As Long) As Long
Private Const SYSTEM_BASICINFORMATION = 0&
Private Const SYSTEM_PERFORMANCEINFORMATION = 2&
Private Const SYSTEM_TIMEINFORMATION = 3&
Private Const NO_ERROR = 0
Private Type LARGE_INTEGER
    dwLow As Long
    dwHigh As Long
End Type

Private Type SYSTEM_PERFORMANCE_INFORMATION
    liIdleTime As LARGE_INTEGER
    dwSpare(0 To 75) As Long
End Type
Private Type SYSTEM_BASIC_INFORMATION
    dwUnknown1 As Long
    uKeMaximumIncrement As Long
    uPageSize As Long
    uMmNumberOfPhysicalPages As Long
    uMmLowestPhysicalPage As Long
    uMmHighestPhysicalPage As Long
    uAllocationGranularity As Long
    pLowestUserAddress As Long
    pMmHighestUserAddress As Long
    uKeActiveProcessors As Long
    bKeNumberProcessors As Byte
    bUnknown2 As Byte
    wUnknown3 As Integer
End Type
Private Type SYSTEM_TIME_INFORMATION
    liKeBootTime As LARGE_INTEGER
    liKeSystemTime As LARGE_INTEGER
    liExpTimeZoneBias As LARGE_INTEGER
    uCurrentTimeZoneId As Long
    dwReserved As Long
End Type

Private lidOldIdle As LARGE_INTEGER
Private liOldSystem As LARGE_INTEGER
 Private Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)
 Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long '获取首个进程
 Private Declare Function Process32First Lib "kernel32" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long '获取下个进程
 Private Declare Function Process32Next Lib "kernel32" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long '释放句柄
 Private Declare Sub CloseHandle Lib "kernel32" (ByVal hPass As Long)
 Private Const TH32CS_SNAPPROCESS = &H2&
 Private Const KEYEVENTF_EXTENDEDKEY = &H1
Private Const KEYEVENTF_KEYUP = &H2
Dim IsHideToTray As Boolean
Private Const FORMAT_MESSAGE_ALLOCATE_BUFFER = &H100
Private Const FORMAT_MESSAGE_ARGUMENT_ARRAY = &H2000
Private Const FORMAT_MESSAGE_FROM_HMODULE = &H800
Private Const FORMAT_MESSAGE_FROM_STRING = &H400
Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Private Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200
Private Const FORMAT_MESSAGE_MAX_WIDTH_MASK = &HFF
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
Private Const VK_LWIN = &H5B
Private Const WM_KEYUP = &H101
Private Const WM_KEYDOWN = &H100
Private Declare Function DebugActiveProcess Lib "kernel32" (ByVal dwProcessId As Long) As Long
Private Declare Sub DebugBreak Lib "kernel32" ()
Private Const SM_DEBUG = 22
Private Const DEBUG_ONLY_THIS_PROCESS = &H2
Private Const DEBUG_PROCESS = &H1
Private Type USER_DIALOG_CONFIG
lpTitle As String
lpIcon As Integer
lpMessage As String
End Type
Private Type USER_APP_RUN
lpAppPath As String
lpAppParam As String
lpRunMode As Integer
End Type
Private Type APP_TASK_PARAM
lpTimerType As Integer
lpDelay As Long
lpRunHour As Integer
lpRunMinute As Integer
lpRunSecond As Integer
lpCurrentHour As Integer
lpCurrentMinute As Integer
lpCurrentSecond As Integer
lpTaskEnum As Integer
lpTaskFriendlyDisplayName As String
lpRunning As Boolean
End Type
Dim lpDialogCfg As USER_DIALOG_CONFIG
Dim lpAppCfg As USER_APP_RUN
Dim lpTaskCfg As APP_TASK_PARAM
Const SC_SCREENSAVE = &HF140&
Dim IsCodeUse As Boolean
Private Declare Function WinExec Lib "kernel32" (ByVal lpCmdLine As String, ByVal nCmdShow As Long) As Long
Private Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Const GWL_WNDPROC = (-4)
Private Const WM_SYSCOMMAND = &H112
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function SetFileAttributes Lib "kernel32" Alias "SetFileAttributesA" (ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Long
Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Private Const FILE_ATTRIBUTE_ARCHIVE = &H20
Private Const FILE_ATTRIBUTE_COMPRESSED = &H800
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10
Private Const FILE_ATTRIBUTE_HIDDEN = &H2
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_ATTRIBUTE_READONLY = &H1
Private Const FILE_ATTRIBUTE_SYSTEM = &H4
Private Const FILE_ATTRIBUTE_TEMPORARY = &H100
Dim lpSize As Long
Dim bchk As Boolean
Dim lpFilePath As String
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1
Const REG_DWORD = 4
Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"
Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Const MAX_FILE_SIZE = 1.5 * (1024 ^ 3)
Private Const HWND_BOTTOM = 1
Private Const HWND_BROADCAST = &HFFFF&
Private Const HWND_DESKTOP = 0
Private Const HWND_NOTOPMOST = -2
Private Const WS_EX_TRANSPARENT = &H20&
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal y As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function BringWindowToTop Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Private Declare Function ClipCursor Lib "user32" (lpRect As Any) As Long
'很多朋友都见到过能在托盘图标上出现气球提示的软件，不说软件，就是在“磁盘空间不足”时Windows给出的提示就属于气球提示，那么怎样在自己的程序中添加这样的气球提示呢？
   
'其实并不难，关键就在添加托盘图标时所使用的NOTIFYICONDATA结构，源代码如下：
Private Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
   
Private Type NOTIFYICONDATA
cbSize   As Long     '   结构大小(字节)
hWnd   As Long     '   处理消息的窗口的句柄
uID   As Long     '   唯一的标识符
uFlags   As Long     '   Flags
uCallbackMessage   As Long     '   处理消息的窗口接收的消息
hIcon   As Long     '   托盘图标句柄
szTip   As String * 128         '   Tooltip   提示文本
dwState   As Long     '   托盘图标状态
dwStateMask   As Long     '   状态掩码
szInfo   As String * 256         '   气球提示文本
uTimeoutOrVersion   As Long     '   气球提示消失时间或版本
'   uTimeout   -   气球提示消失时间(单位:ms,   10000   --   30000)
'   uVersion   -   版本(0   for   V4,   3   for   V5)
szInfoTitle   As String * 64         '   气球提示标题
dwInfoFlags   As Long     '   气球提示图标
End Type
   
'   dwState   to   NOTIFYICONDATA   structure
Private Const NIS_HIDDEN = &H1           '   隐藏图标
Private Const NIS_SHAREDICON = &H2           '   共享图标
   
'   dwInfoFlags   to   NOTIFIICONDATA   structure
Private Const NIIF_NONE = &H0           '   无图标
Private Const NIIF_INFO = &H1           '   "消息"图标
Private Const NIIF_WARNING = &H2           '   "警告"图标
Private Const NIIF_ERROR = &H3           '   "错误"图标
   
'   uFlags   to   NOTIFYICONDATA   structure
Private Const NIF_ICON       As Long = &H2
Private Const NIF_INFO       As Long = &H10
Private Const NIF_MESSAGE       As Long = &H1
Private Const NIF_STATE       As Long = &H8
Private Const NIF_TIP       As Long = &H4
   
'   dwMessage   to   Shell_NotifyIcon
Private Const NIM_ADD       As Long = &H0
Private Const NIM_DELETE       As Long = &H2
Private Const NIM_MODIFY       As Long = &H1
Private Const NIM_SETFOCUS       As Long = &H3
Private Const NIM_SETVERSION       As Long = &H4
Private Type RECTL
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Dim cRect As RECT
Const LCR_UNLOCK = 0
Dim dwMouseFlag As Integer
Const ME_LBCLICK = 1
Const ME_LBDBLCLICK = 2
Const ME_RBCLICK = 3
Private Const MOUSEEVENTF_LEFTDOWN = &H2
Private Const MOUSEEVENTF_ABSOLUTE = &H8000
Private Const MOUSEEVENTF_LEFTUP = &H4
Private Const MOUSEEVENTF_MIDDLEDOWN = &H20
Private Const MOUSEEVENTF_MIDDLEUP = &H40
Private Const MOUSEEVENTF_MOVE = &H1
Private Const MOUSEEVENTF_RIGHTDOWN = &H8
Private Const MOUSEEVENTF_RIGHTUP = &H10
Private Const MOUSETRAILS = 39
Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Declare Function GetKeyboardState Lib "user32" (pbKeyState As Byte) As Long
Const SWP_NOACTIVATE = &H10
Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Dim lpszCaptionNew As String
Private Const SC_MINIMIZE = &HF020&
Private Const WS_MAXIMIZEBOX = &H10000
Dim HKStateCtrl As Integer
Dim HKStateFn As Integer
Private Const WS_MINIMIZEBOX = &H20000
Private Const WS_MAXIMIZE = &H1000000
Private Const WS_MINIMIZE = &H20000000
Private Const WS_ICONIC = WS_MINIMIZE
Const SC_ICON = SC_MINIMIZE
Const SC_TASKLIST = &HF130&
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Dim bCodeUse As Boolean
Private Const WS_CAPTION = &HC00000
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Const PROCESS_ALL_ACCESS = &H1F0FFF
Const SC_RESTORE = &HF120&
Const TH32CS_SNAPHEAPLIST = &H1
Const TH32CS_SNAPTHREAD = &H4
Const TH32CS_SNAPMODULE = &H8
Const TH32CS_SNAPALL = (TH32CS_SNAPHEAPLIST Or TH32CS_SNAPPROCESS Or TH32CS_SNAPTHREAD Or TH32CS_SNAPMODULE)
Const TH32CS_INHERIT = &H80000000
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Dim lMeWinStyle As Long
Const SWP_SHOWWINDOW = &H40
Const SWP_HIDEWINDOW = &H80
Const SWP_NOOWNERZORDER = &H200
Const SWP_NOREPOSITION = SWP_NOOWNERZORDER
Private Const SC_MOVE = &HF010&
Private Const SC_SIZE = &HF000&
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Const WS_EX_APPWINDOW = &H40000
Private Type WINDOWINFORMATION
hWindow As Long
hWindowDC As Long
hThreadProcess As Long
hThreadProcessID As Long
lpszCaption As String
lpszClassName As String
lpszThreadProcessName As String * 1024
lpszThreadProcessPath As String
lpszExe As String
lpszPath As String
End Type
Private Type WINDOWPARAM
bEnabled As Boolean
bHide As Boolean
bTrans As Boolean
bClosable As Boolean
bSizable As Boolean
bMinisizable As Boolean
bTop As Boolean
lpTransValue As Integer
End Type
Dim lpWindow As WINDOWINFORMATION
Dim lpWindowParam() As WINDOWPARAM
Dim lpCur As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Const WM_GETTEXT = &HD
Private Const WM_GETTEXTLENGTH = &HE
Dim lpRtn As Long
Dim hWindow As Long
Dim lpLength As Long
Dim lpArray() As Byte
Dim lpArray2() As Byte
Dim lpBuff As String
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Const WS_EX_LAYERED = &H80000
Private Const GWL_EXSTYLE = (-20)
Private Const LWA_ALPHA = &H2
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Const LWA_COLORKEY = &H1
Private Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal fEnable As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Private Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Const MF_BYPOSITION = &H400&
Private Const MF_REMOVE = &H1000&
Private Const WS_SYSMENU = &H80000
Private Const GWL_STYLE = (-16)
Private Const MF_BYCOMMAND = &H0
Private Const SC_CLOSE = &HF060&
Private Declare Function SetMenu Lib "user32" (ByVal hWnd As Long, ByVal hMenu As Long) As Long
Private Const MF_INSERT = &H0&
Private Const SC_MAXIMIZE = &HF030&
Private Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Private Type WINDOWINFOBOXDATA
lpszCaption As String
lpszClass As String
lpszThread As String
lpszHandle As String
lpszDC As String
End Type
Dim dwWinInfo As WINDOWINFOBOXDATA
Dim bError As Boolean
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Const HTCAPTION = 2
Private Const WM_NCLBUTTONDOWN = &HA1
Const WM_CLOSE = &H10
Private Const HWND_TOPMOST = -1
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOREDRAW = &H8
Private Const SWP_NOMOVE = &H2
Dim mov As Boolean
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Const ANYSIZE_ARRAY = 1
Const EWX_LOGOFF = 0
Const EWX_SHUTDOWN = 1
Const EWX_REBOOT = 2
Const EWX_FORCE = 4
Private Type LUID
UsedPart As Long
IgnoredForNowHigh32BitPart As Long
End Type
Private Type TOKEN_PRIVILEGES
PrivilegeCount As Long
TheLuid As LUID
Attributes As Long
End Type
Private Declare Function OpenProcessToken Lib "advapi32" (ByVal _
ProcessHandle As Long, _
ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Function LookupPrivilegeValue Lib "advapi32" _
Alias "LookupPrivilegeValueA" _
(ByVal lpSystemName As String, ByVal lpName As String, lpLuid _
As LUID) As Long
Private Declare Function AdjustTokenPrivileges Lib "advapi32" _
(ByVal TokenHandle As Long, _
ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES _
, ByVal BufferLength As Long, _
PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long
Private Type TestCounter
TimesLeft As Integer
ResetTime As Integer
End Type
Dim PassTest As TestCounter
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Type POINTAPI
X As Long
y As Long
End Type
Private Const VK_ADD = &H6B
Private Const VK_ATTN = &HF6
Private Const VK_BACK = &H8
Private Const VK_CANCEL = &H3
Private Const VK_CAPITAL = &H14
Private Const VK_CLEAR = &HC
Private Const VK_CONTROL = &H11
Private Const VK_CRSEL = &HF7
Private Const VK_DECIMAL = &H6E
Private Const VK_DELETE = &H2E
Private Const VK_DIVIDE = &H6F
Private Const VK_DOWN = &H28
Private Const VK_END = &H23
Private Const VK_EREOF = &HF9
Private Const VK_ESCAPE = &H1B
Private Const VK_EXECUTE = &H2B
Private Const VK_EXSEL = &HF8
Private Const VK_F1 = &H70
Private Const VK_F10 = &H79
Private Const VK_F11 = &H7A
Private Const VK_F12 = &H7B
Private Const VK_F13 = &H7C
Private Const VK_F14 = &H7D
Private Const VK_F15 = &H7E
Private Const VK_F16 = &H7F
Private Const VK_F17 = &H80
Private Const VK_F18 = &H81
Private Const VK_F19 = &H82
Private Const VK_F2 = &H71
Private Const VK_F20 = &H83
Private Const VK_F21 = &H84
Private Const VK_F22 = &H85
Private Const VK_F23 = &H86
Private Const VK_F24 = &H87
Private Const VK_F3 = &H72
Private Const VK_F4 = &H73
Private Const VK_F5 = &H74
Private Const VK_F6 = &H75
Private Const VK_F7 = &H76
Private Const VK_F8 = &H77
Private Const VK_F9 = &H78
Private Const VK_HELP = &H2F
Private Const VK_HOME = &H24
Private Const VK_INSERT = &H2D
Private Const VK_LBUTTON = &H1
Private Const VK_LCONTROL = &HA2
Private Const VK_LEFT = &H25
Private Const VK_LMENU = &HA4
Private Const VK_LSHIFT = &HA0
Private Const VK_MBUTTON = &H4
Private Const VK_MENU = &H12
Private Const VK_MULTIPLY = &H6A
Private Const VK_NEXT = &H22
Private Const VK_NONAME = &HFC
Private Const VK_NUMLOCK = &H90
Private Const VK_NUMPAD0 = &H60
Private Const VK_NUMPAD1 = &H61
Private Const VK_NUMPAD2 = &H62
Private Const VK_NUMPAD3 = &H63
Private Const VK_NUMPAD4 = &H64
Private Const VK_NUMPAD5 = &H65
Private Const VK_NUMPAD6 = &H66
Private Const VK_NUMPAD7 = &H67
Private Const VK_NUMPAD8 = &H68
Private Const VK_NUMPAD9 = &H69
Private Const VK_OEM_CLEAR = &HFE
Private Const VK_PA1 = &HFD
Private Const VK_PAUSE = &H13
Private Const VK_PLAY = &HFA
Private Const VK_PRINT = &H2A
Private Const VK_PRIOR = &H21
Private Const VK_PROCESSKEY = &HE5
Private Const VK_RBUTTON = &H2
Private Const VK_RCONTROL = &HA3
Private Const VK_RETURN = &HD
Private Const VK_RIGHT = &H27
Private Const VK_RMENU = &HA5
Private Const VK_RSHIFT = &HA1
Private Const VK_SCROLL = &H91
Private Const VK_SELECT = &H29
Private Const VK_SEPARATOR = &H6C
Private Const VK_SHIFT = &H10
Private Const VK_SNAPSHOT = &H2C
Private Const VK_SPACE = &H20
Private Const VK_SUBTRACT = &H6D
Private Const VK_TAB = &H9
Private Const VK_UP = &H26
Private Const VK_ZOOM = &HFB
Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal y As Long) As Long
Dim lpX As Long
Dim lpY As Long
Private Type FILEINFO
lpPath As String
lpDateLastChanged As Date
lpAttribList As Integer
lpSize As Long
lpHeader As String * 25
lpType As String
lpAttrib As String
End Type
Dim lpFile As FILEINFO
Public act As Boolean
Dim regsvrvrt
Dim unregsvrvrt
Dim regflag As Boolean
Dim unregflag  As Boolean
Dim ream
Private Type BROWSEINFO
hOwner As Long
pidlRoot As Long
pszDisplayName As String
lpszTitle As String
ulFlags As Long
lpfn As Long
lParam As Long
iImage As Long
End Type
Private Const BIF_RETURNONLYFSDIRS = &H1
Private Const BIF_NEWDIALOGSTYLE = &H40
Private Const BIF_NONEWFOLDERBUTTON = &H200
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" _
(ByVal pidl As Long, _
ByVal pszPath As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" _
(lpBrowseInfo As BROWSEINFO) As Long
Private Declare Function CloseScreenFun Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Const SC_MONITORPOWER = &HF170&
Private Declare Function GetLastError Lib "kernel32" () As Long
Private Function GetCPUUsage() As Long
    
    Dim sbSysBasicInfo As SYSTEM_BASIC_INFORMATION
    Dim spSysPerforfInfo As SYSTEM_PERFORMANCE_INFORMATION
    Dim stSysTimeInfo As SYSTEM_TIME_INFORMATION
    Dim curIdle As Currency
    Dim curSystem As Currency
    Dim lngResult As Long
    
    GetCPUUsage = -1
    
    lngResult = NtQuerySystemInformation(SYSTEM_BASICINFORMATION, VarPtr(sbSysBasicInfo), LenB(sbSysBasicInfo), 0&)
    If lngResult <> NO_ERROR Then Exit Function
    
    lngResult = NtQuerySystemInformation(SYSTEM_TIMEINFORMATION, VarPtr(stSysTimeInfo), LenB(stSysTimeInfo), 0&)
    If lngResult <> NO_ERROR Then Exit Function
    
    lngResult = NtQuerySystemInformation(SYSTEM_PERFORMANCEINFORMATION, VarPtr(spSysPerforfInfo), LenB(spSysPerforfInfo), ByVal 0&)
    If lngResult <> NO_ERROR Then Exit Function
    curIdle = ConvertLI(spSysPerforfInfo.liIdleTime) - ConvertLI(lidOldIdle)
    curSystem = ConvertLI(stSysTimeInfo.liKeSystemTime) - ConvertLI(liOldSystem)
    If curSystem <> 0 Then curIdle = curIdle / curSystem
    curIdle = 100 - curIdle * 100 / sbSysBasicInfo.bKeNumberProcessors + 0.5
    GetCPUUsage = Int(curIdle)
    
    lidOldIdle = spSysPerforfInfo.liIdleTime
    liOldSystem = stSysTimeInfo.liKeSystemTime
End Function

Private Function ConvertLI(liToConvert As LARGE_INTEGER) As Currency
    CopyMemory ConvertLI, liToConvert, LenB(liToConvert)
End Function
Private Function GetErrorDescription(ByVal lErr As Long) As String
    Dim sReturn As String
    sReturn = String$(256, 32)
    FormatMessage FORMAT_MESSAGE_FROM_SYSTEM Or _
        FORMAT_MESSAGE_IGNORE_INSERTS, 0&, lErr, _
        0&, sReturn, Len(sReturn), ByVal 0
    sReturn = Trim(sReturn)
    GetErrorDescription = sReturn
End Function
Private Function GetProcessID(lpszProcessName As String) As Long
'RETUREN VALUES
'VALUE=-25 : FUNCTION FAILED
'VALUE<>-25 : SUCCEED
Dim pid    As Long
Dim pname As String
Dim a As String
a = Trim(LCase(lpszProcessName))
Dim my    As PROCESSENTRY32
Dim L    As Long
Dim l1    As Long
Dim flag    As Boolean
Dim mName    As String
Dim I    As Integer
L = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
If L Then
    my.dwSize = 1060
End If
If (Process32First(L, my)) Then
    Do
        I = InStr(1, my.szExeFile, Chr(0))
        mName = LCase(Left(my.szExeFile, I - 1))
        If mName = a Then
            pid = my.th32ProcessID
            GetProcessID = pid
            Exit Function
        End If
Loop Until (Process32Next(L, my) < 1)
GetProcessID = -25
End If
End Function
Private Function GetProcessInfo(lpszProcessName As String, lpProcessInfo As PROCESSENTRY32) As Long
Dim pid    As Long
Dim pname As String
Dim a As String
a = Trim(LCase(lpszProcessName))
Dim my    As PROCESSENTRY32
Dim L    As Long
Dim l1    As Long
Dim flag    As Boolean
Dim mName    As String
Dim I    As Integer
L = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
If L Then
    my.dwSize = 1060
End If
If (Process32First(L, my)) Then
    Do
        I = InStr(1, my.szExeFile, Chr(0))
        mName = LCase(Left(my.szExeFile, I - 1))
        If mName = a Then
            pid = my.th32ProcessID
            lpProcessInfo = my
            GetProcessInfo = 245
            Exit Function
        End If
Loop Until (Process32Next(L, my) < 1)
GetProcessInfo = -245
End If
End Function
Private Sub CloseScreenA(ByVal sWitch As Boolean)
If sWitch = True Then
CloseScreenFun GetForegroundWindow, WM_SYSCOMMAND, SC_MONITORPOWER, 1&
Else
CloseScreenFun GetForegroundWindow, WM_SYSCOMMAND, SC_MONITORPOWER, -1&
End If
End Sub
Public Function GetFolderName(hWnd As Long, Text As String) As String
On Error Resume Next
Dim bi As BROWSEINFO
Dim pidl As Long
Dim path As String
With bi
.hOwner = hWnd
.pidlRoot = 0&
.lpszTitle = Text
.ulFlags = BIF_NONEWFOLDERBUTTON
End With
pidl = SHBrowseForFolder(bi)
path = Space$(512)
If SHGetPathFromIDList(ByVal pidl, ByVal path) Then
GetFolderName = Left(path, InStr(path, Chr(0)) - 1)
End If
End Function
Sub GetProcessName(ByVal processID As Long, szExeName As String, szPathName As String)
On Error Resume Next
Dim my As PROCESSENTRY32
Dim hProcessHandle As Long
Dim success As Long
Dim L As Long
L = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
If L Then
my.dwSize = 1060
If (Process32First(L, my)) Then
Do
If my.th32ProcessID = processID Then
CloseHandle L
szExeName = Left$(my.szExeFile, InStr(1, my.szExeFile, Chr$(0)) - 1)
For L = Len(szExeName) To 1 Step -1
If Mid$(szExeName, L, 1) = "\" Then
Exit For
End If
Next L
szPathName = Left$(szExeName, L)
Exit Sub
End If
Loop Until (Process32Next(L, my) < 1)
End If
CloseHandle L
End If
End Sub
Private Sub CreateFile(lpPath As String, lpSize As Long)
On Error Resume Next
End Sub
Private Sub DisableClose(hWnd As Long, Optional ByVal MDIChild As Boolean)
On Error Resume Next
Exit Sub
Dim hSysMenu As Long
Dim nCnt As Long
Dim cID As Long
hSysMenu = GetSystemMenu(hWnd, False)
If hSysMenu = 0 Then
Exit Sub
End If
nCnt = GetMenuItemCount(hSysMenu)
If MDIChild Then
cID = 3
Else
cID = 1
End If
If nCnt Then
RemoveMenu hSysMenu, nCnt - cID, MF_BYPOSITION Or MF_REMOVE
RemoveMenu hSysMenu, nCnt - cID - 1, MF_BYPOSITION Or MF_REMOVE
DrawMenuBar hWnd
End If
End Sub
Private Function GetPassword(hWnd As Long) As String
On Error Resume Next
lpLength = SendMessage(hWnd, WM_GETTEXTLENGTH, 0, 0)
If lpLength > 0 Then
ReDim lpArray(lpLength + 1) As Byte
ReDim lpArray2(lpLength - 1) As Byte
CopyMemory lpArray(0), lpLength, 2
SendMessage hWnd, WM_GETTEXT, lpLength + 1, lpArray(0)
CopyMemory lpArray2(0), lpArray(0), lpLength
GetPassword = StrConv(lpArray2, vbUnicode)
Else
GetPassword = ""
End If
End Function
Private Function GetWindowClassName(hWnd As Long) As String
On Error Resume Next
Dim lpszWindowClassName As String * 256
lpszWindowClassName = Space(256)
GetClassName hWnd, lpszWindowClassName, 256
lpszWindowClassName = Trim(lpszWindowClassName)
GetWindowClassName = lpszWindowClassName
End Function
Private Sub AdjustToken()
On Error Resume Next
Const TOKEN_ADJUST_PRIVILEGES = &H20
Const TOKEN_QUERY = &H8
Const SE_PRIVILEGE_ENABLED = &H2
Dim hdlProcessHandle As Long
Dim hdlTokenHandle As Long
Dim tmpLuid As LUID
Dim tkp As TOKEN_PRIVILEGES
Dim tkpNewButIgnored As TOKEN_PRIVILEGES
Dim lBufferNeeded As Long
hdlProcessHandle = GetCurrentProcess()
OpenProcessToken hdlProcessHandle, (TOKEN_ADJUST_PRIVILEGES Or _
TOKEN_QUERY), hdlTokenHandle
LookupPrivilegeValue "", "SeShutdownPrivilege", tmpLuid
tkp.PrivilegeCount = 1
tkp.TheLuid = tmpLuid
tkp.Attributes = SE_PRIVILEGE_ENABLED
AdjustTokenPrivileges hdlTokenHandle, False, _
tkp, Len(tkpNewButIgnored), tkpNewButIgnored, lBufferNeeded
End Sub
Private Function HexOpen(lpFilePath As String, bSafe As Boolean) As String
Dim strFileName As String
Dim arr() As Byte
strFileName = App.path & "\2.jpg"
Open lpFilePath For Binary As #1
ReDim arr(LOF(1))
Get #1, , arr()
Close #1
Dim T As String
Dim L As Integer
Dim te As String
Dim ASCII As String
L = 0
T = ""
te = ""
ASCII = ""
Dim I
For I = LBound(arr) To UBound(arr)
te = UCase(Hex$(arr(I)))
If arr(I) >= 32 And arr(I) <= 126 Then
ASCII = ASCII & Chr(arr(I))
Else
ASCII = ASCII & "."
End If
If Len(te) = 1 Then te = "0" & te
T = T & te & " "
L = L + 1
If L = 16 Then
T = T & " "
ASCII = ASCII & " "
End If
If L = 32 Then
't = t & " " & ASCII & vbCrLf
T = T
ASCII = ""
L = 0
End If
If bSafe = True Then
If Len(T) >= 72 Then
T = Left(T, 72)
Exit For
End If
End If
Next
HexOpen = T
End Function
Private Function OpenAsHexDocument(lpFile As String, lpHeadOnly As Boolean) As String
On Error Resume Next
Dim strFileName As String
Dim arr() As Byte
strFileName = lpFile
If 245 = 245 Then
Open strFileName For Binary As #1
ReDim arr(LOF(1))
Get #1, , arr()
Close #1
Dim T As String
Dim L As Integer
Dim te As String
Dim ASCII As String
L = 0
T = ""
te = ""
ASCII = ""
Dim I
For I = LBound(arr) To UBound(arr)
te = UCase(Hex$(arr(I)))
If arr(I) >= 32 And arr(I) <= 126 Then
ASCII = ASCII & Chr(arr(I))
Else
ASCII = ASCII & "."
End If
If Len(te) = 1 Then te = "0" & te
T = T & te & " "
If Len(T) >= 72 And lpHeadOnly = True Then
Exit For
End If
L = L + 1
If L = 16 Then
T = T & " "
ASCII = ASCII & " "
End If
If L = 32 Then
T = T
ASCII = ""
L = 0
End If
Next
End If
If lpHeadOnly = True Then
OpenAsHexDocument = Left(T, 72)
Else
OpenAsHexDocument = T
End If
End Function
Private Sub EnumProcess()
Dim SnapShot As Long
Dim NextProcess As Long
Dim PE As PROCESSENTRY32 '创建进程快照
SnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0) '如果队列不为空则搜索
If SnapShot <> -1 Then '设置进程结构长度
PE.dwSize = Len(PE) '获取首个进程
NextProcess = Process32First(SnapShot, PE)
Do While NextProcess '可对进程序做相应处理
'获取下一个
NextProcess = Process32Next(SnapShot, PE)
Loop '释放进程句柄 CloseHandle (SnapShot)
End If
End Sub
Private Sub Command1_Click()
On Error GoTo ep
Dim lpSyspath As String
lpSyspath = Environ("WinDir")
If Right(lpSyspath, 1) = "\" Then
lpSyspath = Left(lpSyspath, Len(lpSyspath) - 1)
End If
lpHosts = lpSyspath & "\System32\Drivers\Etc\Hosts"
Dim lpFileNum As Integer
lpFileNum = FreeFile
Open lpHosts For Input As #lpFileNum
Dim lpTmp As String
Picture2.Visible = True
Shape1.Visible = True
Image1.Visible = True
Label1.Visible = True
List1.Clear
Refresh
Do While Not EOF(lpFileNum)
Input #lpFileNum, lpTmp
List1.AddItem lpTmp
Loop
Close
Picture2.Visible = False
Shape1.Visible = False
Image1.Visible = False
Label1.Visible = False
Label1.Caption = "正在載入文件 " & lpHosts & " 請稍候..."
Refresh
Caption = "Hosts Editor - PC-DOS Workshop" & " [" & lpHosts & "]"
Command3.Enabled = True
Command4.Enabled = True
Command5.Enabled = True
Command6.Enabled = True
Command8.Enabled = True
Command10.Enabled = True
mnuViewCurrent.Enabled = True
Me.mnuCm.Enabled = True
Me.mnuCopy.Enabled = True
Me.mnuDE.Enabled = True
Me.mnuDelete.Enabled = True
Me.mnuDeleteAll.Enabled = True
Me.mnuDisallow.Enabled = True
Me.mnuES.Enabled = True
Me.mnuDE.Enabled = True
Me.mnuSave.Enabled = True
Me.mnuSaveAs.Enabled = True
Me.mnuSE.Enabled = True
Me.mnuInsertDNS.Enabled = True
Exit Sub
ep:
MsgBox "發生系統錯誤：" & Err.Description & vbCrLf & "請檢查Hosts文件是否存在于 " & lpSyspath & "\System32\Drivers\Etc\" & "目錄下且系統環境變量%WinDir%是否有誤", vbCritical, "Error"
Command3.Enabled = False
Command4.Enabled = False
Command5.Enabled = False
Command6.Enabled = False
Command8.Enabled = False
Command10.Enabled = False
List1.Clear
Picture2.Visible = False
Shape1.Visible = False
Image1.Visible = False
Label1.Visible = False
Me.mnuCm.Enabled = False
Me.mnuCopy.Enabled = False
Me.mnuDE.Enabled = False
Me.mnuDelete.Enabled = False
Me.mnuDeleteAll.Enabled = False
Me.mnuDisallow.Enabled = False
Me.mnuES.Enabled = False
Me.mnuDE.Enabled = False
Me.mnuSave.Enabled = False
Me.mnuSaveAs.Enabled = False
Me.mnuSE.Enabled = False
Me.mnuInsertDNS.Enabled = False
mnuViewCurrent.Enabled = False
End Sub
Private Sub Command10_Click()
On Error Resume Next
If List1.ListIndex >= 0 Then
frmEdit.Text3.Text = List1.List(List1.ListIndex)
frmEdit.Show 1
Else
MsgBox "請先選擇要編輯的項目", vbCritical, "Error"
End If
End Sub
Private Sub Command2_Click()
On Error Resume Next
PopupMenu Me.mnuIOpen, 1, Command2.Left + Picture1.Left, Picture1.Left + Picture1.Height, Me.mnuPDefault
End Sub
Private Sub Command3_Click()
On Error Resume Next
PopupMenu Me.mnuISave, 1, Command3.Left + Picture1.Left, Picture1.Left + Picture1.Height, Me.mnuPSave
End Sub
Private Sub Command4_Click()
On Error GoTo ep
Dim lpFreeFile As Integer
lpFreeFile = FreeFile
Open lpHosts For Output As #lpFreeFile
Dim I As Long
If List1.ListCount = 0 Then
Print #lpFreeFile, ""
Close
ElseIf List1.ListCount = 1 Then
Print #lpFreeFile, List1.List(0)
Close
Else
For I = 0 To List1.ListCount - 1
Print #lpFreeFile, List1.List(I)
Next
Close
End If
Close
Exit Sub
ep:
MsgBox "發生系統錯誤：" & Err.Description, vbCritical, "Error"
End Sub
Private Sub Command5_Click()
On Error Resume Next
PopupMenu Me.mnuIInsert, 1, Command5.Left + Picture1.Left, Picture1.Left + Picture1.Height, Me.mnuPInsertDNS
End Sub
Private Sub Command6_Click()
On Error Resume Next
frmIDNS.Show 1
End Sub
Private Sub Command8_Click()
On Error Resume Next
If List1.ListIndex >= 0 Then
Dim ans As Integer
ans = MsgBox("是否刪除選定條目 [" & List1.List(List1.ListIndex) & "] ?", vbQuestion + vbYesNo, "Ask")
If ans = vbYes Then
List1.RemoveItem (List1.ListIndex)
List1.Refresh
Else
Exit Sub
End If
Else
MsgBox "請選擇要刪除的條目", vbCritical, "Error"
Exit Sub
End If
End Sub
Private Sub Form_Load()
On Error Resume Next
lpAST = 0
lpHosts = ""
Caption = "Hosts Editor - PC-DOS Workshop [No file opened]"
Picture2.Visible = False
Shape1.Visible = False
Label1.Visible = False
Image1.Visible = False
Command3.Enabled = False
Command4.Enabled = False
Command5.Enabled = False
Command6.Enabled = False
Command8.Enabled = False
Command10.Enabled = False
Me.mnuCm.Enabled = False
Me.mnuCopy.Enabled = False
Me.mnuDE.Enabled = False
Me.mnuDelete.Enabled = False
Me.mnuDeleteAll.Enabled = False
Me.mnuDisallow.Enabled = False
Me.mnuES.Enabled = False
Me.mnuDE.Enabled = False
Me.mnuSave.Enabled = False
Me.mnuSaveAs.Enabled = False
Me.mnuSE.Enabled = False
Me.mnuInsertDNS.Enabled = False
mnuViewCurrent.Enabled = False
SetDebugToken.SetDebugToken
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
If lpHosts <> "" Then
Dim ans As Integer
ans = MsgBox("是否保存内容到 " & lpHosts & " ?", vbExclamation + vbYesNoCancel, "Ask")
If ans = vbYes Then
Dim lpFreeFile As Integer
lpFreeFile = FreeFile
Open lpHosts For Output As #lpFreeFile
Dim I As Long
If List1.ListCount = 0 Then
Print #lpFreeFile, ""
Close
ElseIf List1.ListCount = 1 Then
Print #lpFreeFile, List1.List(0)
Close
Else
For I = 0 To List1.ListCount - 1
Print #lpFreeFile, List1.List(I)
Next
Close
End If
Close
ElseIf ans = vbNo Then
Cancel = 0
End
Else
Cancel = 245
Exit Sub
End If
End If
End
End Sub
Private Sub List1_Click()
On Error Resume Next
End Sub
Private Sub List1_DblClick()
On Error Resume Next
If List1.ListIndex >= 0 Then
frmEdit.Text3.Text = List1.List(List1.ListIndex)
frmEdit.Show 1
End If
End Sub
Private Sub List1_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
If List1.ListIndex >= 0 Then
frmEdit.Text3.Text = List1.List(List1.ListIndex)
frmEdit.Show 1
End If
End If
End Sub
Private Sub mnuAbout_Click()
On Error Resume Next
frmAbout.Show 1
End Sub
Private Sub mnuAutoSave_Click()
On Error Resume Next
frmAutoSave.Show 1
End Sub
Private Sub mnuCm_Click()
On Error Resume Next
frmIC.Show 1
End Sub
Private Sub mnuCopy_Click()
On Error Resume Next
If List1.ListIndex >= 0 Then
Clipboard.SetText (List1.List(List1.ListIndex))
Else
MsgBox "請選擇要複製的條目", vbCritical, "Error"
End If
End Sub
Private Sub mnuCustom_Click()
On Error GoTo ep
Dim lpSyspath As String
lpSyspath = Environ("WinDir")
If Right(lpSyspath, 1) = "\" Then
lpSyspath = Left(lpSyspath, Len(lpSyspath) - 1)
End If
lpHosts = lpSyspath & "\System32\Drivers\Etc\Hosts"
Dim CommonDialogVar As New CCommonDialog
Dim IsCanceled As Boolean
With CommonDialogVar
.DialogTitle = "請選擇要打開的Hosts文件"
.Filter = "所有文件(*.*)|*.*"
.ShowModalWindow = True
.hWndCall = hWnd
IsCanceled = .ShowOpen
End With
If IsCanceled = False Then
Exit Sub
End If
lpHosts = CommonDialogVar.FileName
Dim lpFileNum As Integer
lpFileNum = FreeFile
Open lpHosts For Input As #lpFileNum
Dim lpTmp As String
Picture2.Visible = True
Shape1.Visible = True
Image1.Visible = True
Label1.Visible = True
List1.Clear
Refresh
Do While Not EOF(lpFileNum)
Input #lpFileNum, lpTmp
List1.AddItem lpTmp
Loop
Close
Picture2.Visible = False
Shape1.Visible = False
Image1.Visible = False
Label1.Visible = False
Label1.Caption = "正在載入文件 " & lpHosts & " 請稍候..."
Refresh
Caption = "Hosts Editor - PC-DOS Workshop" & " [" & lpHosts & "]"
Command3.Enabled = True
Command4.Enabled = True
Command5.Enabled = True
Command6.Enabled = True
Command8.Enabled = True
Command10.Enabled = True
Me.mnuCm.Enabled = True
Me.mnuCopy.Enabled = True
Me.mnuDE.Enabled = True
Me.mnuDelete.Enabled = True
Me.mnuDeleteAll.Enabled = True
Me.mnuDisallow.Enabled = True
Me.mnuES.Enabled = True
Me.mnuDE.Enabled = True
Me.mnuSave.Enabled = True
Me.mnuSaveAs.Enabled = True
Me.mnuSE.Enabled = True
mnuViewCurrent.Enabled = True
Me.mnuInsertDNS.Enabled = True
Exit Sub
ep:
MsgBox "發生系統錯誤：" & Err.Description, vbCritical, "Error"
Command3.Enabled = False
Command4.Enabled = False
Command5.Enabled = False
Command6.Enabled = False
Command8.Enabled = False
Command10.Enabled = False
List1.Clear
Picture2.Visible = False
Shape1.Visible = False
Image1.Visible = False
Label1.Visible = False
Me.mnuCm.Enabled = False
Me.mnuCopy.Enabled = False
Me.mnuDE.Enabled = False
Me.mnuDelete.Enabled = False
Me.mnuDeleteAll.Enabled = False
Me.mnuDisallow.Enabled = False
Me.mnuES.Enabled = False
Me.mnuDE.Enabled = False
Me.mnuSave.Enabled = False
Me.mnuSaveAs.Enabled = False
Me.mnuSE.Enabled = False
mnuViewCurrent.Enabled = False
Me.mnuInsertDNS.Enabled = False
End Sub
Private Sub mnuDE_Click()
On Error Resume Next
Close
End
End Sub
Private Sub mnuDefault_Click()
On Error GoTo ep
Dim lpSyspath As String
lpSyspath = Environ("WinDir")
If Right(lpSyspath, 1) = "\" Then
lpSyspath = Left(lpSyspath, Len(lpSyspath) - 1)
End If
lpHosts = lpSyspath & "\System32\Drivers\Etc\Hosts"
Dim lpFileNum As Integer
lpFileNum = FreeFile
Open lpHosts For Input As #lpFileNum
Dim lpTmp As String
Picture2.Visible = True
Shape1.Visible = True
Image1.Visible = True
Label1.Visible = True
List1.Clear
Refresh
Do While Not EOF(lpFileNum)
Input #lpFileNum, lpTmp
List1.AddItem lpTmp
Loop
Close
Picture2.Visible = False
Shape1.Visible = False
Image1.Visible = False
Label1.Visible = False
Label1.Caption = "正在載入文件 " & lpHosts & " 請稍候..."
Refresh
Caption = "Hosts Editor - PC-DOS Workshop" & " [" & lpHosts & "]"
Command3.Enabled = True
Command4.Enabled = True
Command5.Enabled = True
Command6.Enabled = True
Command8.Enabled = True
Command10.Enabled = True
Me.mnuCm.Enabled = True
Me.mnuCopy.Enabled = True
Me.mnuDE.Enabled = True
Me.mnuDelete.Enabled = True
Me.mnuDeleteAll.Enabled = True
Me.mnuDisallow.Enabled = True
Me.mnuES.Enabled = True
Me.mnuDE.Enabled = True
Me.mnuSave.Enabled = True
Me.mnuSaveAs.Enabled = True
Me.mnuSE.Enabled = True
mnuViewCurrent.Enabled = True
Me.mnuInsertDNS.Enabled = True
Exit Sub
ep:
MsgBox "發生系統錯誤：" & Err.Description & vbCrLf & "請檢查Hosts文件是否存在于 " & lpSyspath & "\System32\Drivers\Etc\" & "目錄下且系統環境變量%WinDir%是否有誤", vbCritical, "Error"
Command3.Enabled = False
Command4.Enabled = False
Command5.Enabled = False
Command6.Enabled = False
Command8.Enabled = False
Command10.Enabled = False
List1.Clear
Picture2.Visible = False
Shape1.Visible = False
Image1.Visible = False
Label1.Visible = False
Me.mnuCm.Enabled = False
Me.mnuCopy.Enabled = False
Me.mnuDE.Enabled = False
Me.mnuDelete.Enabled = False
Me.mnuDeleteAll.Enabled = False
Me.mnuDisallow.Enabled = False
Me.mnuES.Enabled = False
Me.mnuDE.Enabled = False
Me.mnuSave.Enabled = False
Me.mnuSaveAs.Enabled = False
Me.mnuSE.Enabled = False
mnuViewCurrent.Enabled = False
Me.mnuInsertDNS.Enabled = False
End Sub
Private Sub mnuDelete_Click()
On Error Resume Next
On Error Resume Next
If List1.ListIndex >= 0 Then
Dim ans As Integer
ans = MsgBox("是否刪除選定條目 [" & List1.List(List1.ListIndex) & "] ?", vbQuestion + vbYesNo, "Ask")
If ans = vbYes Then
List1.RemoveItem (List1.ListIndex)
List1.Refresh
Else
Exit Sub
End If
Else
MsgBox "請選擇要刪除的條目", vbCritical, "Error"
Exit Sub
End If
End Sub
Private Sub mnuDeleteAll_Click()
On Error Resume Next
Dim ans As Integer
ans = MsgBox("確定清除所有條目嗎?", vbQuestion + vbYesNo, "Ask")
If ans = vbYes Then
List1.Clear
Else
Exit Sub
End If
End Sub
Private Sub mnuDisallow_Click()
On Error Resume Next
frmIDisallow.Show 1
End Sub
Private Sub mnuES_Click()
On Error Resume Next
If List1.ListIndex >= 0 Then
frmEdit.Text3.Text = List1.List(List1.ListIndex)
frmEdit.Show 1
Else
MsgBox "請先選擇要編輯的項目", vbCritical, "Error"
End If
End Sub
Private Sub mnuExit_Click()
On Error Resume Next
If lpHosts <> "" Then
Dim ans As Integer
ans = MsgBox("是否保存内容到 " & lpHosts & " ?", vbExclamation + vbYesNo, "Ask")
If ans = vbYes Then
Dim lpFreeFile As Integer
lpFreeFile = FreeFile
Open lpHosts For Output As #lpFreeFile
Dim I As Long
If List1.ListCount = 0 Then
Print #lpFreeFile, ""
Close
ElseIf List1.ListCount = 1 Then
Print #lpFreeFile, List1.List(0)
Close
Else
For I = 0 To List1.ListCount - 1
Print #lpFreeFile, List1.List(I)
Next
Close
End If
Close
End If
End If
End
End Sub
Private Sub mnuInsertDNS_Click()
On Error Resume Next
frmIDNS.Show 1
End Sub
Private Sub mnuPCM_Click()
On Error Resume Next
frmIC.Show 1
End Sub
Private Sub mnuPCustom_Click()
On Error GoTo ep
Dim lpSyspath As String
lpSyspath = Environ("WinDir")
If Right(lpSyspath, 1) = "\" Then
lpSyspath = Left(lpSyspath, Len(lpSyspath) - 1)
End If
lpHosts = lpSyspath & "\System32\Drivers\Etc\Hosts"
Dim CommonDialogVar As New CCommonDialog
Dim IsCanceled As Boolean
With CommonDialogVar
.DialogTitle = "請選擇要打開的Hosts文件"
.Filter = "所有文件(*.*)|*.*"
.ShowModalWindow = True
.hWndCall = hWnd
IsCanceled = .ShowOpen
End With
If IsCanceled = False Then
Exit Sub
End If
lpHosts = CommonDialogVar.FileName
Dim lpFileNum As Integer
lpFileNum = FreeFile
Open lpHosts For Input As #lpFileNum
Dim lpTmp As String
Picture2.Visible = True
Shape1.Visible = True
Image1.Visible = True
Label1.Visible = True
List1.Clear
Refresh
Do While Not EOF(lpFileNum)
Input #lpFileNum, lpTmp
List1.AddItem lpTmp
Loop
Close
Picture2.Visible = False
Shape1.Visible = False
Image1.Visible = False
Label1.Visible = False
Label1.Caption = "正在載入文件 " & lpHosts & " 請稍候..."
Refresh
Caption = "Hosts Editor - PC-DOS Workshop" & " [" & lpHosts & "]"
Command3.Enabled = True
Command4.Enabled = True
Command5.Enabled = True
Command6.Enabled = True
Command8.Enabled = True
Command10.Enabled = True
Me.mnuCm.Enabled = True
Me.mnuCopy.Enabled = True
Me.mnuDE.Enabled = True
Me.mnuDelete.Enabled = True
Me.mnuDeleteAll.Enabled = True
Me.mnuDisallow.Enabled = True
Me.mnuES.Enabled = True
Me.mnuDE.Enabled = True
Me.mnuSave.Enabled = True
Me.mnuSaveAs.Enabled = True
Me.mnuSE.Enabled = True
mnuViewCurrent.Enabled = True
Me.mnuInsertDNS.Enabled = True
Exit Sub
ep:
MsgBox "發生系統錯誤：" & Err.Description, vbCritical, "Error"
Command3.Enabled = False
Command4.Enabled = False
Command5.Enabled = False
Command6.Enabled = False
Command8.Enabled = False
Command10.Enabled = False
List1.Clear
Picture2.Visible = False
Shape1.Visible = False
Image1.Visible = False
Label1.Visible = False
Me.mnuCm.Enabled = False
Me.mnuCopy.Enabled = False
Me.mnuDE.Enabled = False
Me.mnuDelete.Enabled = False
Me.mnuDeleteAll.Enabled = False
Me.mnuDisallow.Enabled = False
Me.mnuES.Enabled = False
Me.mnuDE.Enabled = False
Me.mnuSave.Enabled = False
Me.mnuSaveAs.Enabled = False
Me.mnuSE.Enabled = False
mnuViewCurrent.Enabled = False
Me.mnuInsertDNS.Enabled = False
End Sub
Private Sub mnuPDefault_Click()
On Error GoTo ep
Dim lpSyspath As String
lpSyspath = Environ("WinDir")
If Right(lpSyspath, 1) = "\" Then
lpSyspath = Left(lpSyspath, Len(lpSyspath) - 1)
End If
lpHosts = lpSyspath & "\System32\Drivers\Etc\Hosts"
Dim lpFileNum As Integer
lpFileNum = FreeFile
Open lpHosts For Input As #lpFileNum
Dim lpTmp As String
Picture2.Visible = True
Shape1.Visible = True
Image1.Visible = True
Label1.Visible = True
List1.Clear
Refresh
Do While Not EOF(lpFileNum)
Input #lpFileNum, lpTmp
List1.AddItem lpTmp
Loop
Close
Picture2.Visible = False
Shape1.Visible = False
Image1.Visible = False
Label1.Visible = False
Label1.Caption = "正在載入文件 " & lpHosts & " 請稍候..."
Refresh
Caption = "Hosts Editor - PC-DOS Workshop" & " [" & lpHosts & "]"
Command3.Enabled = True
Command4.Enabled = True
Command5.Enabled = True
Command6.Enabled = True
Command8.Enabled = True
Command10.Enabled = True
Me.mnuCm.Enabled = True
Me.mnuCopy.Enabled = True
Me.mnuDE.Enabled = True
Me.mnuDelete.Enabled = True
Me.mnuDeleteAll.Enabled = True
Me.mnuDisallow.Enabled = True
Me.mnuES.Enabled = True
Me.mnuDE.Enabled = True
Me.mnuSave.Enabled = True
Me.mnuSaveAs.Enabled = True
Me.mnuSE.Enabled = True
mnuViewCurrent.Enabled = True
Me.mnuInsertDNS.Enabled = True
Exit Sub
ep:
MsgBox "發生系統錯誤：" & Err.Description & vbCrLf & "請檢查Hosts文件是否存在于 " & lpSyspath & "\System32\Drivers\Etc\" & "目錄下且系統環境變量%WinDir%是否有誤", vbCritical, "Error"
Command3.Enabled = False
Command4.Enabled = False
Command5.Enabled = False
Command6.Enabled = False
Command8.Enabled = False
Command10.Enabled = False
List1.Clear
Picture2.Visible = False
Shape1.Visible = False
Image1.Visible = False
Label1.Visible = False
Me.mnuCm.Enabled = False
Me.mnuCopy.Enabled = False
Me.mnuDE.Enabled = False
Me.mnuDelete.Enabled = False
Me.mnuDeleteAll.Enabled = False
Me.mnuDisallow.Enabled = False
Me.mnuES.Enabled = False
Me.mnuDE.Enabled = False
Me.mnuSave.Enabled = False
Me.mnuSaveAs.Enabled = False
Me.mnuSE.Enabled = False
mnuViewCurrent.Enabled = False
Me.mnuInsertDNS.Enabled = False
End Sub
Private Sub mnuPDisallow_Click()
On Error Resume Next
frmIDisallow.Show 1
End Sub
Private Sub mnuPInsertDNS_Click()
frmIDNS.Show 1
End Sub
Private Sub mnuPSave_Click()
On Error GoTo ep
Dim lpFreeFile As Integer
lpFreeFile = FreeFile
Open lpHosts For Output As #lpFreeFile
Dim I As Long
If List1.ListCount = 0 Then
Print #lpFreeFile, ""
Close
ElseIf List1.ListCount = 1 Then
Print #lpFreeFile, List1.List(0)
Close
Else
For I = 0 To List1.ListCount - 1
Print #lpFreeFile, List1.List(I)
Next
Close
End If
Close
Exit Sub
ep:
MsgBox "發生系統錯誤：" & Err.Description, vbCritical, "Error"
End Sub
Private Sub mnuPSaveAs_Click()
On Error GoTo ep
Dim ans As Integer
Dim lpTraget As String
Dim CDV As New CCommonDialog
Dim IsCanceled As Boolean
Dim lpFreeFile As Integer
With CDV
.DialogTitle = "請選擇Hosts配置文檔保存位置"
.Filter = "所有文件(*.*)|*.*"
.ShowModalWindow = True
.hWndCall = hWnd
IsCanceled = .ShowSave
End With
If IsCanceled = False Then
Exit Sub
End If
lpTraget = CDV.FileName
If Dir(lpTraget) <> "" Then
ans = MsgBox("文件 " & lpTraget & " 已存在，是否替換?", vbExclamation + vbYesNo, "Ask")
If ans = vbYes Then
lpFreeFile = FreeFile
Open lpTraget For Output As #lpFreeFile
Dim I As Long
If List1.ListCount = 0 Then
Print #lpFreeFile, ""
Close
ElseIf List1.ListCount = 1 Then
Print #lpFreeFile, List1.List(0)
Close
Else
For I = 0 To List1.ListCount - 1
Print #lpFreeFile, List1.List(I)
Next
Close
End If
Close
Else
Exit Sub
End If
Else
lpFreeFile = FreeFile
Open lpTraget For Output As #lpFreeFile
If List1.ListCount = 0 Then
Print #lpFreeFile, ""
Close
ElseIf List1.ListCount = 1 Then
Print #lpFreeFile, List1.List(0)
Close
Else
For I = 0 To List1.ListCount - 1
Print #lpFreeFile, List1.List(I)
Next
Close
End If
Close
End If
Exit Sub
ep:
MsgBox "發生系統錯誤：" & Err.Description, vbCritical, "Error"
End Sub
Private Sub mnuRepair_Click()
On Error GoTo ep
Dim ansM As Integer
ansM = MsgBox("本功能用於修復被破壞的Hosts文檔以便恢復您對某些Internet或本地站點的存取,繼續嗎?", vbQuestion + vbYesNo, "Ask")
If ansM = vbYes Then
Dim ans As Integer
Dim lpTraget As String
Dim CDV As New CCommonDialog
Dim IsCanceled As Boolean
Dim lpFreeFile As Integer
With CDV
.DialogTitle = "請選擇Hosts配置文檔保存位置"
.Filter = "所有文件(*.*)|*.*"
.ShowModalWindow = True
.hWndCall = hWnd
IsCanceled = .ShowSave
End With
If IsCanceled = False Then
Close
Picture2.Visible = False
Shape1.Visible = False
Image1.Visible = False
Label1.Visible = False
Label1.Caption = "正在載入文件 " & lpHosts & " 請稍候..."
Refresh
Caption = "Hosts Editor - PC-DOS Workshop" & " [" & lpHosts & "]"
Command3.Enabled = True
Command4.Enabled = True
Command5.Enabled = True
Command6.Enabled = True
Command8.Enabled = True
Command10.Enabled = True
Me.mnuCm.Enabled = True
Me.mnuCopy.Enabled = True
Me.mnuDE.Enabled = True
Me.mnuDelete.Enabled = True
Me.mnuDeleteAll.Enabled = True
Me.mnuDisallow.Enabled = True
Me.mnuES.Enabled = True
Me.mnuDE.Enabled = True
Me.mnuSave.Enabled = True
Me.mnuSaveAs.Enabled = True
Me.mnuSE.Enabled = True
mnuViewCurrent.Enabled = True
Me.mnuInsertDNS.Enabled = True
Exit Sub
End If
lpTraget = CDV.FileName
If Dir(lpTraget) <> "" Then
ans = MsgBox("文件 " & lpTraget & " 已存在，是否替換?", vbExclamation + vbYesNo, "Ask")
If ans = vbYes Then
lpFreeFile = FreeFile
Open lpTraget For Output As #lpFreeFile
Dim I As Long
For I = 0 To lstXPHOSTS.ListCount - 1
Print #lpFreeFile, lstXPHOSTS.List(I)
Next
Close
Close
ans = MsgBox("文檔修復完成,您是否要載入並進入編輯模式?", vbExclamation + vbYesNo, "Ask")
If ans = vbYes Then
If 25 = 245 Then
Dim lpSyspath As String
lpSyspath = Environ("WinDir")
If Right(lpSyspath, 1) = "\" Then
lpSyspath = Left(lpSyspath, Len(lpSyspath) - 1)
End If
End If
lpHosts = lpTraget
Dim lpFileNum As Integer
lpFileNum = FreeFile
Open lpHosts For Input As #lpFileNum
Dim lpTmp As String
Picture2.Visible = True
Shape1.Visible = True
Image1.Visible = True
Label1.Visible = True
List1.Clear
Refresh
Do While Not EOF(lpFileNum)
Input #lpFileNum, lpTmp
List1.AddItem lpTmp
Loop
Close
Picture2.Visible = False
Shape1.Visible = False
Image1.Visible = False
Label1.Visible = False
Label1.Caption = "正在載入文件 " & lpHosts & " 請稍候..."
Refresh
Caption = "Hosts Editor - PC-DOS Workshop" & " [" & lpHosts & "]"
Command3.Enabled = True
Command4.Enabled = True
Command5.Enabled = True
Command6.Enabled = True
Command8.Enabled = True
Command10.Enabled = True
Me.mnuCm.Enabled = True
Me.mnuCopy.Enabled = True
Me.mnuDE.Enabled = True
Me.mnuDelete.Enabled = True
Me.mnuDeleteAll.Enabled = True
Me.mnuDisallow.Enabled = True
Me.mnuES.Enabled = True
Me.mnuDE.Enabled = True
Me.mnuSave.Enabled = True
Me.mnuSaveAs.Enabled = True
Me.mnuSE.Enabled = True
mnuViewCurrent.Enabled = True
Me.mnuInsertDNS.Enabled = True
Else
Close
Picture2.Visible = False
Shape1.Visible = False
Image1.Visible = False
Label1.Visible = False
Label1.Caption = "正在載入文件 " & lpHosts & " 請稍候..."
Refresh
Caption = "Hosts Editor - PC-DOS Workshop" & " [" & lpHosts & "]"
Command3.Enabled = True
Command4.Enabled = True
Command5.Enabled = True
Command6.Enabled = True
Command8.Enabled = True
Command10.Enabled = True
Me.mnuCm.Enabled = True
Me.mnuCopy.Enabled = True
Me.mnuDE.Enabled = True
Me.mnuDelete.Enabled = True
Me.mnuDeleteAll.Enabled = True
Me.mnuDisallow.Enabled = True
Me.mnuES.Enabled = True
Me.mnuDE.Enabled = True
Me.mnuSave.Enabled = True
Me.mnuSaveAs.Enabled = True
Me.mnuSE.Enabled = True
mnuViewCurrent.Enabled = True
Me.mnuInsertDNS.Enabled = True
Exit Sub
End If
Else
Close
Picture2.Visible = False
Shape1.Visible = False
Image1.Visible = False
Label1.Visible = False
Label1.Caption = "正在載入文件 " & lpHosts & " 請稍候..."
Refresh
Caption = "Hosts Editor - PC-DOS Workshop" & " [" & lpHosts & "]"
Command3.Enabled = True
Command4.Enabled = True
Command5.Enabled = True
Command6.Enabled = True
Command8.Enabled = True
Command10.Enabled = True
Me.mnuCm.Enabled = True
Me.mnuCopy.Enabled = True
Me.mnuDE.Enabled = True
Me.mnuDelete.Enabled = True
Me.mnuDeleteAll.Enabled = True
Me.mnuDisallow.Enabled = True
Me.mnuES.Enabled = True
Me.mnuDE.Enabled = True
Me.mnuSave.Enabled = True
Me.mnuSaveAs.Enabled = True
Me.mnuSE.Enabled = True
mnuViewCurrent.Enabled = True
Me.mnuInsertDNS.Enabled = True
Exit Sub
End If
Else
lpFreeFile = FreeFile
Open lpTraget For Output As #lpFreeFile
For I = 0 To lstXPHOSTS.ListCount - 1
Print #lpFreeFile, lstXPHOSTS.List(I)
Next
Close
ans = MsgBox("文檔修復完成,您是否要載入並進入編輯模式?", vbExclamation + vbYesNo, "Ask")
If ans = vbYes Then
If 25 = 245 Then
lpSyspath = Environ("WinDir")
If Right(lpSyspath, 1) = "\" Then
lpSyspath = Left(lpSyspath, Len(lpSyspath) - 1)
End If
End If
lpHosts = lpTraget
lpFileNum = FreeFile
Open lpHosts For Input As #lpFileNum
Picture2.Visible = True
Shape1.Visible = True
Image1.Visible = True
Label1.Visible = True
List1.Clear
Refresh
Do While Not EOF(lpFileNum)
Input #lpFileNum, lpTmp
List1.AddItem lpTmp
Loop
Close
Picture2.Visible = False
Shape1.Visible = False
Image1.Visible = False
Label1.Visible = False
Label1.Caption = "正在載入文件 " & lpHosts & " 請稍候..."
Refresh
Caption = "Hosts Editor - PC-DOS Workshop" & " [" & lpHosts & "]"
Command3.Enabled = True
Command4.Enabled = True
Command5.Enabled = True
Command6.Enabled = True
Command8.Enabled = True
Command10.Enabled = True
Me.mnuCm.Enabled = True
Me.mnuCopy.Enabled = True
Me.mnuDE.Enabled = True
Me.mnuDelete.Enabled = True
Me.mnuDeleteAll.Enabled = True
Me.mnuDisallow.Enabled = True
Me.mnuES.Enabled = True
Me.mnuDE.Enabled = True
Me.mnuSave.Enabled = True
Me.mnuSaveAs.Enabled = True
Me.mnuSE.Enabled = True
mnuViewCurrent.Enabled = True
Me.mnuInsertDNS.Enabled = True
Else
Close
Picture2.Visible = False
Shape1.Visible = False
Image1.Visible = False
Label1.Visible = False
Label1.Caption = "正在載入文件 " & lpHosts & " 請稍候..."
Refresh
Caption = "Hosts Editor - PC-DOS Workshop" & " [" & lpHosts & "]"
Command3.Enabled = True
Command4.Enabled = True
Command5.Enabled = True
Command6.Enabled = True
Command8.Enabled = True
Command10.Enabled = True
Me.mnuCm.Enabled = True
Me.mnuCopy.Enabled = True
Me.mnuDE.Enabled = True
Me.mnuDelete.Enabled = True
Me.mnuDeleteAll.Enabled = True
Me.mnuDisallow.Enabled = True
Me.mnuES.Enabled = True
Me.mnuDE.Enabled = True
Me.mnuSave.Enabled = True
Me.mnuSaveAs.Enabled = True
Me.mnuSE.Enabled = True
mnuViewCurrent.Enabled = True
Me.mnuInsertDNS.Enabled = True
Exit Sub
End If
End If
End If
Exit Sub
ep:
MsgBox "發生系統錯誤：" & Err.Description, vbCritical, "Error"
Close
Picture2.Visible = False
Shape1.Visible = False
Image1.Visible = False
Label1.Visible = False
Label1.Caption = "正在載入文件 " & lpHosts & " 請稍候..."
Refresh
Caption = "Hosts Editor - PC-DOS Workshop" & " [" & lpHosts & "]"
Command3.Enabled = True
Command4.Enabled = True
Command5.Enabled = True
Command6.Enabled = True
Command8.Enabled = True
Command10.Enabled = True
Me.mnuCm.Enabled = True
Me.mnuCopy.Enabled = True
Me.mnuDE.Enabled = True
Me.mnuDelete.Enabled = True
Me.mnuDeleteAll.Enabled = True
Me.mnuDisallow.Enabled = True
Me.mnuES.Enabled = True
Me.mnuDE.Enabled = True
Me.mnuSave.Enabled = True
Me.mnuSaveAs.Enabled = True
Me.mnuSE.Enabled = True
mnuViewCurrent.Enabled = True
Me.mnuInsertDNS.Enabled = True
End Sub
Private Sub mnuSaveAs_Click()
On Error GoTo ep
Dim ans As Integer
Dim lpTraget As String
Dim CDV As New CCommonDialog
Dim IsCanceled As Boolean
Dim lpFreeFile As Integer
With CDV
.DialogTitle = "請選擇Hosts配置文檔保存位置"
.Filter = "所有文件(*.*)|*.*"
.ShowModalWindow = True
.hWndCall = hWnd
IsCanceled = .ShowSave
End With
If IsCanceled = False Then
Exit Sub
End If
lpTraget = CDV.FileName
If Dir(lpTraget) <> "" Then
ans = MsgBox("文件 " & lpTraget & " 已存在，是否替換?", vbExclamation + vbYesNo, "Ask")
If ans = vbYes Then
lpFreeFile = FreeFile
Open lpTraget For Output As #lpFreeFile
Dim I As Long
If List1.ListCount = 0 Then
Print #lpFreeFile, ""
Close
ElseIf List1.ListCount = 1 Then
Print #lpFreeFile, List1.List(0)
Close
Else
For I = 0 To List1.ListCount - 1
Print #lpFreeFile, List1.List(I)
Next
Close
End If
Close
Else
Exit Sub
End If
Else
lpFreeFile = FreeFile
Open lpTraget For Output As #lpFreeFile
If List1.ListCount = 0 Then
Print #lpFreeFile, ""
Close
ElseIf List1.ListCount = 1 Then
Print #lpFreeFile, List1.List(0)
Close
Else
For I = 0 To List1.ListCount - 1
Print #lpFreeFile, List1.List(I)
Next
Close
End If
Close
End If
Exit Sub
ep:
MsgBox "發生系統錯誤：" & Err.Description, vbCritical, "Error"
End Sub
Private Sub mnuSE_Click()
On Error GoTo ep
Dim lpFreeFile As Integer
lpFreeFile = FreeFile
Open lpHosts For Output As #lpFreeFile
Dim I As Long
If List1.ListCount = 0 Then
Print #lpFreeFile, ""
Close
ElseIf List1.ListCount = 1 Then
Print #lpFreeFile, List1.List(0)
Close
Else
For I = 0 To List1.ListCount - 1
Print #lpFreeFile, List1.List(I)
Next
Close
End If
Close
End
Exit Sub
ep:
Dim ans As Integer
ans = MsgBox("發生系統錯誤：" & Err.Description & vbCrLf & "是否強制退出?", vbCritical + vbYesNo, "Error")
If ans = vbYes Then
End
Else
Exit Sub
End If
End Sub
Private Sub mnuViewCurrent_Click()
On Error Resume Next
If List1.ListIndex >= 0 Then
MsgBox "當前選定的條目：" & vbCrLf & List1.List(List1.ListIndex), vbInformation, "Info"
Else
MsgBox "尚未選擇條目", vbCritical, "Error"
End If
End Sub
Private Sub Timer1_Timer()
On Error Resume Next
lpAST = lpAST + 1
If CStr(lpAST) >= CStr(Me.Tag) Then
If lpHosts = "" Then
lpAST = 0
Exit Sub
End If
Dim lpFreeFile As Integer
lpFreeFile = FreeFile
Open lpHosts For Output As #lpFreeFile
Dim I As Long
If List1.ListCount = 0 Then
Print #lpFreeFile, ""
Close
ElseIf List1.ListCount = 1 Then
Print #lpFreeFile, List1.List(0)
Close
Else
For I = 0 To List1.ListCount - 1
Print #lpFreeFile, List1.List(I)
Next
Close
End If
Close
lpAST = 0
End If
End Sub
