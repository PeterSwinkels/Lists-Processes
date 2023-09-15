Attribute VB_Name = "ListProcessesModule"
'This module contains this program's core procedures.
Option Explicit

'Defines the Microsoft Windows API constants, functions, and structures used byt this program.
Private Type LUID
   LowPart As Long
   HighPart As Long
End Type

Private Type LUID_AND_ATTRIBUTES
   pLuid As LUID
   Attributes As Long
End Type

Private Type MODULEINFO
   lpBaseOfDLL As Long
   SizeOfImage As Long
   EntryPoint As Long
End Type

Private Type TOKEN_PRIVILEGES
   PrivilegeCount As Long
   Privileges(1) As LUID_AND_ATTRIBUTES
End Type

Private Type THREADENTRY32
   dwSize As Long
   cntUsage As Long
   th32ThreadID As Long
   th32OwnerProcessID As Long
   tpBasePri As Long
   tpDeltaPri As Long
   dwFlags As Long
End Type

Private Const ERROR_NO_MORE_FILES As Long = 18
Private Const ERROR_PARTIAL_COPY As Long = 299
Private Const ERROR_SUCCESS As Long = 0
Private Const FORMAT_MESSAGE_FROM_SYSTEM As Long = &H1000
Private Const FORMAT_MESSAGE_IGNORE_INSERTS As Long = &H200&
Private Const INVALID_HANDLE_VALUE As Long = -1
Private Const MAX_PATH As Long = 260
Private Const MAX_SHORT_STRING As Long = 255
Private Const PROCESS_ALL_ACCESS As Long = &H1F0FFF
Private Const PROCESS_QUERY_INFORMATION As Long = &H400&
Private Const SE_DEBUG_NAME As String = "SeDebugPrivilege"
Private Const SE_PRIVILEGE_DISABLED As Long = &H0&
Private Const SE_PRIVILEGE_ENABLED As Long = &H2&
Private Const TH32CS_SNAPTHREAD As Long = &H4&
Private Const TOKEN_ALL_ACCESS As Long = &HFF&

Private Declare Function AdjustTokenPrivileges Lib "Advapi32.dll" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long
Private Declare Function CloseHandle Lib "Kernel32.dll" (ByVal hObject As Long) As Long
Private Declare Function CreateToolhelp32Snapshot Lib "Kernel32.dll" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
Private Declare Function EnumProcessModules Lib "Psapi.dll" (ByVal hProcess As Long, lphModule As Long, ByVal cb As Long, lpcbNeeded As Long) As Long
Private Declare Function EnumProcesses Lib "Psapi.dll" (ByRef lpidProcess As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Private Declare Function EnumThreadWindows Lib "User32.dll" (ByVal dwThreadId As Long, ByVal lpfn As Long, ByVal lparam As Long) As Long
Private Declare Function FormatMessageA Lib "Kernel32.dll" (ByVal dwFlags As Long, lpSource As Long, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
Private Declare Function GetCurrentProcess Lib "Kernel32.dll" () As Long
Private Declare Function GetModuleFileNameExA Lib "Psapi.dll" (ByVal hProcess As Long, ByVal hModule As Long, ByVal ModuleName As String, ByVal nSize As Long) As Long
Private Declare Function GetModuleInformation Lib "Psapi.dll" (ByVal hProcess As Long, ByVal hModule As Long, lpmodinfo As MODULEINFO, ByVal cb As Long) As Long
Private Declare Function GetProcessImageFileNameW Lib "Psapi.dll" (ByVal hProcess As Long, ByVal lpImageFileName As Long, ByVal nSize As Long) As Long
Private Declare Function IsWow64Process Lib "Kernel32.dll" (ByVal hProcess As Long, ByRef Wow64Process As Long) As Long
Private Declare Function LookupPrivilegeValueA Lib "Advapi32.dll" (ByVal lpSystemName As String, ByVal lpName As String, lpLuid As LUID) As Long
Private Declare Function lstrcpy Lib "Kernel32.dll" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long
Private Declare Function OpenProcess Lib "Kernel32.dll" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function OpenProcessToken Lib "Advapi32.dll" (ByVal ProcessH As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Function SafeArrayGetDim Lib "Oleaut32.dll" (ByRef saArray() As Any) As Long
Private Declare Function Thread32First Lib "Kernel32.dll" (ByVal hObject As Long, p As THREADENTRY32) As Boolean
Private Declare Function Thread32Next Lib "Kernel32.dll" (ByVal hObject As Long, p As THREADENTRY32) As Boolean
Private Declare Function VDMEnumProcessWOW Lib "vdmdbg.dll" (ByVal fp As Long, lparam As Long) As Integer
Private Declare Function VDMEnumTaskWOWEx Lib "vdmdbg.dll " (ByVal dwProcessId As Long, ByVal fp As Long, lparam As Long) As Integer

'Defines the constants, structures, and variables used by this program.

Private Const MAX_STRING As Long = 65535   'Defines the maximum length allowed for a string buffer
Private Const NO_MODULEH As Long = 0       'Indicates "No module handle."
Private Const NO_PROCESSH As Long = 0      'Indicates "No process handle."
Private Const NO_PROCESSID As Long = 0     'Indicates "No process id."
Private Const NO_THREADID As Long = 0      'Indicates "No thread id."

'This structure defines a WOW process:
Private Type WOWProcessStr
   ModuleH() As Long      'Defines the list of module handles.
   ModulePath() As Long   'Defines the list of module paths.
   Path() As String       'Defines the path of the process.
   ProcessId As Long      'Defines the process id.
   ThreadId() As Long     'Defines the list of thread ids.
End Type

Private OutputFileH As Long               'Defines the handle of the output file.
Private ThreadWindowCount As Long         'Defines the number of windows attached to a thread.
Private WOWProcesses() As WOWProcessStr   'Contains the list WOW processes.

'This procedure checks whether an error has occurred during the most recent Windows API call.
Private Function CheckForError(ReturnValue As Long, Optional Ignored As Long = ERROR_SUCCESS) As Long
Dim Description As String
Dim ErrorCode As Long
Dim Length As Long
Dim Message As String

   ErrorCode = Err.LastDllError
   Err.Clear
   
   If Not (ErrorCode = ERROR_SUCCESS Or ErrorCode = Ignored) Then
      Description = String$(MAX_STRING, vbNullChar)
      Length = FormatMessageA(FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, CLng(0), ErrorCode, CLng(0), Description, Len(Description), CLng(0))
      If Length = 0 Then
         Description = "No description."
      ElseIf Length > 0 Then
         Description = Left$(Description, Length - 1)
      End If
   
      Message = Message & "API Error: " & CStr(ErrorCode) & vbCrLf
      Message = Message & Description & vbCrLf
      Message = Message & "Return value: " & CStr(ReturnValue) & vbCrLf
      On Error GoTo ErrorTrap
      Print #OutputFileH, String$(25, "=")
      Print #OutputFileH, Message
      Print #OutputFileH, String$(25, "=")
      On Error GoTo 0
   End If
   
EndRoutine:
   CheckForError = ReturnValue
   Exit Function
   
ErrorTrap:
   MsgBox Message, vbExclamation
   Resume EndRoutine
End Function


'This procedure returns a list of module handles for the specified process.
Private Function GetModuleHandles(ProcessId As Long) As Long()
Dim ModuleH() As Long
Dim ProcessH As Long
Dim Size As Long
Dim SizeUsed As Long

   ReDim ModuleH(0 To 0) As Long
   
   ProcessH = CheckForError(OpenProcess(PROCESS_ALL_ACCESS, CLng(False), ProcessId))
   If Not ProcessH = NO_PROCESSH Then
      Do
         Size = (UBound(ModuleH()) + 1) * Len(ModuleH(LBound(ModuleH())))
         CheckForError EnumProcessModules(ProcessH, ModuleH(0), Size, SizeUsed), ERROR_PARTIAL_COPY
         If Size > SizeUsed Then Exit Do
         ReDim Preserve ModuleH(LBound(ModuleH()) To UBound(ModuleH()) + 1) As Long
      Loop
      CheckForError CloseHandle(ProcessH), ERROR_PARTIAL_COPY
   End If
   
   GetModuleHandles = ModuleH()
End Function

'This procedure returns the memory information for the specified module.
Private Function GetModuleMemoryInformation(ProcessId As Long, ModuleH As Long) As MODULEINFO
Dim ModuleInformation As MODULEINFO
Dim ProcessH As Long

   ProcessH = CheckForError(OpenProcess(PROCESS_ALL_ACCESS, CLng(False), ProcessId))
   If Not ProcessH = NO_PROCESSH Then
      CheckForError GetModuleInformation(ProcessH, ModuleH, ModuleInformation, Len(ModuleInformation))
      CheckForError CloseHandle(ProcessH)
   End If
   
   GetModuleMemoryInformation = ModuleInformation
End Function


'This procedure returns the path of the specified module.
Private Function GetModulePath(ModuleH As Long, ProcessId As Long) As String
Dim Length As Long
Dim Path As String
Dim ProcessH As Long

   Path = vbNullString
   ProcessH = CheckForError(OpenProcess(PROCESS_ALL_ACCESS, CLng(False), ProcessId))
   If Not ProcessH = NO_PROCESSH Then
      Path = String$(MAX_PATH, vbNullChar)
      Length = CheckForError(GetModuleFileNameExA(ProcessH, ModuleH, Path, Len(Path)))
      Path = Left$(Path, Length)
      CheckForError CloseHandle(ProcessH)
   End If
   
   GetModulePath = Path
End Function

'This procedure returns a list of ids for all active processes.
Private Function GetProcessIds() As Long()
Dim ProcessIds() As Long
Dim Size As Long
Dim SizeUsed As Long

   ReDim ProcessIds(0 To 0) As Long
   Do
      Size = (UBound(ProcessIds()) + 1) * Len(ProcessIds(LBound(ProcessIds())))
      CheckForError EnumProcesses(ProcessIds(0), Size, SizeUsed), ERROR_PARTIAL_COPY
      If Size > SizeUsed Then Exit Do
      ReDim Preserve ProcessIds(LBound(ProcessIds()) To UBound(ProcessIds()) + 1) As Long
   Loop
   
   GetProcessIds = ProcessIds()
End Function

'This procedure returns the path of the specified process.
Private Function GetProcessPath(ProcessId As Long) As String
Dim Length As Long
Dim Path As String
Dim ProcessH As Long
   
   Path = vbNullString
   ProcessH = CheckForError(OpenProcess(PROCESS_QUERY_INFORMATION, CLng(False), ProcessId))
   If Not ProcessH = NO_PROCESSH Then
      Path = String$(MAX_PATH, vbNullChar)
      Length = CheckForError(GetProcessImageFileNameW(ProcessH, StrPtr(Path), Len(Path)))
      Path = Left$(Path, Length)
      CheckForError CloseHandle(ProcessH)
   End If
   
   GetProcessPath = Path
End Function

'This procedure returns a list of ids for all active threads.
Private Function GetThreads() As THREADENTRY32()
Dim ReturnValue As Long
Dim Threads() As THREADENTRY32
Dim ThreadsH As Long

   ThreadsH = CheckForError(CreateToolhelp32Snapshot(TH32CS_SNAPTHREAD, CLng(0)))
   If Not ThreadsH = INVALID_HANDLE_VALUE Then
      ReDim Threads(0 To 0) As THREADENTRY32
      Threads(UBound(Threads())).dwSize = Len(Threads(UBound(Threads())))
      ReturnValue = CheckForError(Thread32First(ThreadsH, Threads(UBound(Threads()))), ERROR_NO_MORE_FILES)
      Do Until ReturnValue = 0
         Threads(UBound(Threads())).dwSize = Len(Threads(UBound(Threads())))
         ReturnValue = CheckForError(Thread32Next(ThreadsH, Threads(UBound(Threads()))), ERROR_NO_MORE_FILES)
         ReDim Preserve Threads(LBound(Threads()) To UBound(Threads()) + 1) As THREADENTRY32
      Loop
      CheckForError CloseHandle(ThreadsH)
   End If
   
   GetThreads = Threads()
End Function

'This procedure handles the tasks attached to a WOW process.
Private Function HandleTasks(ByVal dwThreadId As Long, ByVal hMod16 As Long, ByVal hTask16 As Long, ByVal pszModName As Long, ByVal pszFileName As Long, lpUserDefined As Long) As Long
   With WOWProcesses(UBound(WOWProcesses()))
      .ModuleH(UBound(.ModuleH())) = hMod16
      .ModulePath(UBound(.ModulePath())) = PointerToString(pszModName)
      .Path(UBound(.Path())) = PointerToString(pszFileName)
      .ThreadId(UBound(.ThreadId())) = dwThreadId
      
      ReDim Preserve .ModuleH(LBound(.ModuleH()) To UBound(.ModuleH()) + 1) As Long
      ReDim Preserve .ModulePath(LBound(.ModulePath()) To UBound(.ModulePath()) + 1) As Long
      ReDim Preserve .Path(LBound(.Path()) To UBound(.Path()) + 1) As String
      ReDim Preserve .ThreadId(LBound(.ThreadId()) To UBound(.ThreadId()) + 1) As Long
   End With
   
   HandleTasks = CLng(False)
End Function
'This procedure handles a window attached to a thread.
Private Function HandleThreadWindows(ByVal lhWnd As Long, ByVal lparam As Long) As Long
   ThreadWindowCount = ThreadWindowCount + 1
   HandleThreadWindows = CLng(True)
End Function

'This procedure handles the WOW processes.
Private Function HandleWOWProcesses(ByVal dwProcessId As Long, ByVal dwAttributes As Long, lpUserDefined As Long) As Long
   With WOWProcesses(UBound(WOWProcesses()))
      ReDim .ModuleH(0 To 0) As Long
      ReDim .ModulePath(0 To 0) As Long
      ReDim .Path(0 To 0) As String
      ReDim .ThreadId(0 To 0) As Long
      .ProcessId = dwProcessId
   End With
   
   CheckForError VDMEnumTaskWOWEx(dwProcessId, AddressOf HandleTasks, CLng(0))
   
   ReDim Preserve WOWProcesses(LBound(WOWProcesses()) To UBound(WOWProcesses()) + 1) As WOWProcessStr
   
   HandleWOWProcesses = CLng(False)
End Function






'This procedure returns a value indicating whether the specified process is a 32 bit process.
Private Function Is32BitProcess(ProcessId As Long) As Boolean
Dim ProcessH As Long
Dim Result As Long

   Result = False
   ProcessH = CheckForError(OpenProcess(PROCESS_QUERY_INFORMATION, CLng(False), ProcessId))
   If Not ProcessH = NO_PROCESSH Then
      CheckForError IsWow64Process(ProcessH, Result)
      CheckForError CloseHandle(ProcessH)
   End If
   
   Is32BitProcess = CBool(Result)
End Function

'This procedure is executed when this program is started.
Private Sub Main()
On Error GoTo ErrorTrap
Dim ModuleIndex As Long
Dim ProcessIndex As Long
Dim ThreadIndex As Long
Dim ModuleH() As Long
Dim Path As String
Dim ProcessIds() As Long
Dim Threads() As THREADENTRY32
  
   ChDrive Left$(App.Path, InStr(App.Path, ":"))
   ChDir App.Path
   
   With App
      Path = InputBox$("Write process data to:", .Title & " v" & CStr(.Major) & "." & CStr(.Minor) & CStr(.Revision) & " - by: " & .CompanyName, ".\Processes.txt")
   End With
   
   If Not Path = vbNullString Then
      OutputFileH = FreeFile()
      Open Path For Output As OutputFileH
         SetPrivilege SE_DEBUG_NAME, SE_PRIVILEGE_ENABLED
      
         ProcessIds() = GetProcessIds()
         Threads() = GetThreads()
         
         For ProcessIndex = LBound(ProcessIds()) To UBound(ProcessIds())
            If Not ProcessIds(ProcessIndex) = NO_PROCESSID Then
               Print #OutputFileH, "Process: "; CStr(ProcessIds(ProcessIndex)); " "; GetProcessPath(ProcessIds(ProcessIndex))
               Print #OutputFileH, "64 bit: "; CStr(Not Is32BitProcess(ProcessIds(ProcessIndex)))
               
               Print #OutputFileH, "Threads:"
               For ThreadIndex = LBound(Threads()) To UBound(Threads())
                  If Threads(ThreadIndex).th32OwnerProcessID = ProcessIds(ProcessIndex) Then
                     ThreadWindowCount = 0
                     EnumThreadWindows Threads(ThreadIndex).th32ThreadID, AddressOf HandleThreadWindows, CLng(0)
                     Print #OutputFileH, " "; CStr(Threads(ThreadIndex).th32ThreadID);
                     If ThreadWindowCount > 0 Then Print #OutputFileH, " [Windows: "; CStr(ThreadWindowCount); "]";
                     Print #OutputFileH,
                  End If
               Next ThreadIndex
               
               ModuleH() = GetModuleHandles(ProcessIds(ProcessIndex))
               Print #OutputFileH, "Modules:"
               For ModuleIndex = LBound(ModuleH()) To UBound(ModuleH())
                  If Not ModuleH(ModuleIndex) = NO_MODULEH Then
                     Print #OutputFileH, " "; CStr(ModuleH(ModuleIndex)); " "; GetModulePath(ModuleH(ModuleIndex), ProcessIds(ProcessIndex))
                     With GetModuleMemoryInformation(ProcessIds(ProcessIndex), ModuleH(ModuleIndex))
                        Print #OutputFileH, "  Entry point: "; CStr(.EntryPoint)
                        Print #OutputFileH, "  DLL base: "; CStr(.lpBaseOfDLL)
                        Print #OutputFileH, "  Image size: "; CStr(.SizeOfImage)
                     End With
                  End If
               Next ModuleIndex
               Print #OutputFileH,
            End If
         Next ProcessIndex
         
         ReDim WOWProcesses(0 To 0) As WOWProcessStr
         VDMEnumProcessWOW AddressOf HandleWOWProcesses, CLng(0)
         For ProcessIndex = LBound(WOWProcesses()) To UBound(WOWProcesses())
            With WOWProcesses(ProcessIndex)
               If Not .ProcessId = NO_PROCESSID Then
                  Print #OutputFileH, "Process: "; CStr(WOWProcesses(ProcessIndex).ProcessId)
                  For ThreadIndex = LBound(.ThreadId()) To UBound(.ThreadId())
                     If Not .ThreadId(ThreadIndex) = NO_THREADID Then
                        Print #OutputFileH, "Thread: "; CStr(.ThreadId(ThreadIndex))
                        Print #OutputFileH, "Module: "; CStr(.ModuleH(ThreadIndex)); " "; .ModulePath(ThreadIndex)
                        Print #OutputFileH, "Path: "; .Path(ThreadIndex)
                     End If
                  Next ThreadIndex
               End If
            End With
         Next ProcessIndex
   End If
   
EndRoutine:
   SetPrivilege SE_DEBUG_NAME, SE_PRIVILEGE_DISABLED
   Close OutputFileH
   Exit Sub
   
ErrorTrap:
   If MsgBox(Err.Description & vbCr & "Error code: " & CStr(Err.Number), vbExclamation Or vbOKCancel) = vbCancel Then End
   Resume EndRoutine
End Sub


'This procedure returns the string indicated by the specified pointer.
Private Function PointerToString(Pointer As Long) As String
Dim Index As Long
Dim Text As String * MAX_SHORT_STRING

   CheckForError lstrcpy(Text, Pointer)
   Index = InStr(Text, vbNullChar)
   If Index > 0 Then Text = Left$(Text, Index - 1) Else Text = vbNullString
   
   PointerToString = Text
End Function
'This procedure disables/enables the specified privilege.
Private Sub SetPrivilege(PrivilegeName As String, Status As Long)
Dim Length As Long
Dim NewPrivileges As TOKEN_PRIVILEGES
Dim PreviousPrivileges As TOKEN_PRIVILEGES
Dim PrivilegeId As LUID
Dim ReturnValue As Long
Dim TokenH As Long

   ReturnValue = CheckForError(OpenProcessToken(GetCurrentProcess(), TOKEN_ALL_ACCESS, TokenH))
   If Not ReturnValue = 0 Then
      ReturnValue = CheckForError(LookupPrivilegeValueA(vbNullString, PrivilegeName, PrivilegeId))
      If Not ReturnValue = 0 Then
         NewPrivileges.Privileges(0).pLuid = PrivilegeId
         NewPrivileges.PrivilegeCount = CLng(1)
         NewPrivileges.Privileges(0).Attributes = Status
         
         CheckForError AdjustTokenPrivileges(TokenH, CLng(False), NewPrivileges, Len(NewPrivileges), PreviousPrivileges, Length)
      End If
      CheckForError CloseHandle(TokenH)
   End If
End Sub

