Attribute VB_Name = "SearchPID"
Public Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long

Public Declare Function Process32First Lib "kernel32" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long

Public Declare Function Process32Next Lib "kernel32" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long

Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal blnheritHandle As Long, ByVal dwAppProcessId As Long) As Long

Public Declare Function TerminateProcess Lib "kernel32" (ByVal ApphProcess As Long, ByVal uExitCode As Long) As Long

Public Type PROCESSENTRY32

dwSize As Long

cntUsage As Long

th32ProcessID As Long

th32DefaultHeapID As Long

th32ModuleID As Long

cntThreads As Long

th32ParentProcessID As Long

pcPriClassBase As Long

dwFlags As Long

szExeFile As String * 1024

End Type

Public Const TH32CS_SNAPHEAPLIST = &H1

Public Const TH32CS_SNAPPROCESS = &H2

Public Const TH32CS_SNAPTHREAD = &H4

Public Const TH32CS_SNAPMODULE = &H8

Public Const TH32CS_SNAPALL = (TH32CS_SNAPHEAPLIST Or TH32CS_SNAPPROCESS Or TH32CS_SNAPTHREAD Or TH32CS_SNAPMODULE)

Public Const TH32CS_INHERIT = &H80000000


Public Declare Function RegisterWindowMessage Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long
Private WM_TASKBARCREATED As Long

'**********************************************************************
'在查找函数
'**********************************************************************
Public Function FindPro(jinchenming As String) As Integer
Dim my As PROCESSENTRY32
Dim l As Long
Dim l1 As Long
Dim mName As String

Dim I As Integer
l = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
If l Then
my.dwSize = 1060

If (Process32First(l, my)) Then '遍历开始

Do
I = InStr(1, my.szExeFile, Chr(0))

mName = LCase(Left(my.szExeFile, I - 1))

If mName = jinchenming Then

Pid = my.th32ProcessID

pname = mName

'Dim mProcID As Long

'mProcID = OpenProcess(1&, -1&, Pid)

FindPro = Pid
'TerminateProcess mProcID, 0&


Exit Function
End If
Loop Until (Process32Next(l, my) < 1)
End If
l1 = CloseHandle(l)
End If
End Function


