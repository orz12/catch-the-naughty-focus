Attribute VB_Name = "Module1"

Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
(ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" _
(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal _
lParam As Long) As Long

Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" _
(ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal _
wParam As Long, ByVal lParam As Long) As Long

Public Const WM_NCACTIVATE = &H86
Public Const GWL_WNDPROC = (-4)
Public OldWndProc&

Public Declare Function GetForegroundWindow Lib "user32" () As Long
Public Declare Function GetWindowText Lib "user32.dll" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32.dll" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long

Public Declare Function GetWindowThreadProcessId Lib "user32 " (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Public Declare Function SetWindowPos& Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)

Public Declare Function NtQueryInformationProcess Lib "NTDLL.DLL " (ByVal ProcessHandle As Long, _
                                                                ByVal ProcessInformationClass As PROCESSINFOCLASS, _
                                                                ByVal ProcessInformation As Long, _
                                                                ByVal ProcessInformationLength As Long, _
                                                                ByRef ReturnLength As Long) As Long
 
Public Enum PROCESSINFOCLASS
        ProcessBasicInformation = 0
        ProcessQuotaLimits
        ProcessIoCounters
        ProcessVmCounters
        ProcessTimes
        ProcessBasePriority
        ProcessRaisePriority
        ProcessDebugPort
        ProcessExceptionPort
        ProcessAccessToken
        ProcessLdtInformation
        ProcessLdtSize
        ProcessDefaultHardErrorMode
        ProcessIoPortHandlers
        ProcessPooledUsageAndLimits
        ProcessWorkingSetWatch
        ProcessUserModeIOPL
        ProcessEnableAlignmentFaultFixup
        ProcessPriorityClass
        ProcessWx86Information
        ProcessHandleCount
        ProcessAffinityMask
        ProcessPriorityBoost
        ProcessDeviceMap
        ProcessSessionInformation
        ProcessForegroundInformation
        ProcessWow64Information
        ProcessImageFileName
        ProcessLUIDDeviceMapsEnabled
        ProcessBreakOnTermination
        ProcessDebugObjectHandle
        ProcessDebugFlags
        ProcessHandleTracing
        ProcessIoPriority
        ProcessExecuteFlags
        ProcessResourceManagement
        ProcessCookie
        ProcessImageInformation
        MaxProcessInfoClass
End Enum
Public Type PROCESS_BASIC_INFORMATION
        ExitStatus   As Long     'NTSTATUS
        PebBaseAddress   As Long     'PPEB
        AffinityMask   As Long     'ULONG_PTR
        BasePriority   As Long     'KPRIORITY
        UniqueProcessId   As Long     'ULONG_PTR
        InheritedFromUniqueProcessId   As Long     'ULONG_PTR
End Type
Public Type OBJECT_ATTRIBUTES
        Length   As Long
        RootDirectory   As Long
        ObjectName   As Long
        Attributes   As Long
        SecurityDescriptor   As Long
        SecurityQualityOfService   As Long
End Type
Public Type CLIENT_ID
        UniqueProcess   As Long
        UniqueThread     As Long
End Type
Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Public Declare Function NtOpenProcess Lib "ntdll.dll" (ByRef ProcessHandle As Long, _
                                                                ByVal AccessMask As Long, _
                                                                ByRef ObjectAttributes As OBJECT_ATTRIBUTES, _
                                                                ByRef ClientID As CLIENT_ID) As Long
'Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
'Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function NtClose Lib "ntdll.dll" (ByVal ObjectHandle As Long) As Long

Public Declare Function NtSuspendProcess Lib "ntdll.dll" (ByVal hProc As Long) As Long
Public Declare Function NtResumeProcess Lib "ntdll.dll" (ByVal hProc As Long) As Long

Public Const PROCESS_TERMINATE As Long = (&H1)
Public Const PROCESS_VM_READ = &H10
Public Const PROCESS_QUERY_INFORMATION As Long = (&H400)
Public Const PROCESS_SUSPEND_RESUME As Long = (&H800)

Public gblProcessId As Long
Public gblInheritedPID As Long

Public Function getInheritedPID(ByVal PID As Long) As Long
    Dim pclass As PROCESSINFOCLASS
    Dim baseInfo As PROCESS_BASIC_INFORMATION
    Dim objA As OBJECT_ATTRIBUTES
    Dim objCid As CLIENT_ID
    Dim lrtn As Long
    Dim hProcess As Long
    objA.Length = Len(objA)
    objCid.UniqueProcess = PID
    lrtn = NtOpenProcess(hProcess, PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, objA, objCid)
    ntstatus = NtQueryInformationProcess(hProcess, ProcessBasicInformation, VarPtr(baseInfo), Len(baseInfo), ByVal 0&)
    getInheritedPID = baseInfo.InheritedFromUniqueProcessId
    lrtn = NtClose(hProcess)
End Function

Public Function SuspendProcess(ByVal PID As Long) As Long
    Dim pclass As PROCESSINFOCLASS
    Dim baseInfo As PROCESS_BASIC_INFORMATION
    Dim objA As OBJECT_ATTRIBUTES
    Dim objCid As CLIENT_ID
    Dim lrtn As Long
    Dim hProcess As Long
    objA.Length = Len(objA)
    objCid.UniqueProcess = PID
    lrtn = NtOpenProcess(hProcess, PROCESS_SUSPEND_RESUME, objA, objCid)
    If hProcess Then SuspendProcess = NtSuspendProcess(hProcess)
    lrtn = NtClose(hProcess)
End Function

'根据窗口句柄得到该窗口的标题
Public Function getCaption(ByVal hWnd As Long)
    Dim hWndlength As Long, hWndTitle As String, A As Long
    hWndlength = GetWindowTextLength(hWnd)
    hWndTitle = String$(hWndlength, 0)
    A = GetWindowText(hWnd, hWndTitle, (hWndlength + 1))
    getCaption = hWndTitle
End Function

Public Function GetPID(ByVal hWnd As Long) As Long
    Dim lpdwProcessId As Long, lrtn As Long
    lrtn = GetWindowThreadProcessId(hWnd, lpdwProcessId)
    GetPID = lpdwProcessId
End Function

Public Function GetForegroundWindowInfo(Optional ByVal bSetGlobal As Boolean) As String

    Dim hWnd1 As Long, lpdwProcessId As Long, lInheritedPID As Long
    hWnd1 = GetForegroundWindow()
    lpdwProcessId = GetPID(hWnd1)
    lInheritedPID = getInheritedPID(lpdwProcessId)
    If bSetGlobal Then
        gblProcessId = lpdwProcessId
        gblInheritedPID = lInheritedPID
    End If
    GetForegroundWindowInfo = "窗体句柄：" & hWnd1 & vbCrLf & "窗体标题：" & getCaption(hWnd1) & vbCrLf & "窗体进程：" & lpdwProcessId & vbCrLf & "　父进程：" & lInheritedPID

End Function

Public Function Hook&(ByVal hWnd1&)
    OldWndProc = SetWindowLong(hWnd1, GWL_WNDPROC, AddressOf NewWndProc)
    Hook = OldWndProc
End Function

Public Sub UnHook(ByVal hWnd1&)
    SetWindowLong hWnd1, GWL_WNDPROC, OldWndProc
End Sub

Public Function NewWndProc&(ByVal hWnd1&, ByVal uMsg&, ByVal wParam&, ByVal lParam&)
    If uMsg = WM_NCACTIVATE Then
        If wParam = 0 Then '失去焦点
            Form1.Caption = "失去焦点"
            Form1.Command1.Enabled = True
            Form1.Command2.Enabled = True
            If Form1.Text1.BackColor = &H80000005 Then
                If Form1.Check2.Value Then Form1.Text1.BackColor = vbBlack
                Form1.Text1.Text = GetForegroundWindowInfo(True)
                If Form1.Check3.Value Then SuspendProcess gblProcessId
                If Form1.Check4.Value Then SuspendProcess gblInheritedPID
                
            End If
        Else
            Form1.Caption = "得到焦点"
            '在这里加入在得到焦点时想要执行的代码
        End If
    End If
    NewWndProc = CallWindowProc(OldWndProc, hWnd1, uMsg, wParam, lParam)
    
End Function

