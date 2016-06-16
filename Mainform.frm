VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form MainForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "学校机房破解"
   ClientHeight    =   2955
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3435
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   3435
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton UnlockGame 
      Caption         =   "解锁游戏"
      Height          =   615
      Left            =   720
      TabIndex        =   1
      Top             =   1560
      Width           =   1815
   End
   Begin VB.CommandButton KillStudentMain 
      Caption         =   "上锁学生端"
      Height          =   735
      Left            =   720
      TabIndex        =   0
      Top             =   360
      Width           =   1695
   End
   Begin VB.Frame fraTreeView 
      Height          =   6480
      Left            =   600
      TabIndex        =   2
      Top             =   3360
      Width           =   7800
      Begin VB.CommandButton cmdGetProcess 
         Caption         =   "Command2"
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   2400
         Visible         =   0   'False
         Width           =   615
      End
      Begin ComctlLib.TreeView treProcess 
         Height          =   5940
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   7485
         _ExtentX        =   13203
         _ExtentY        =   10478
         _Version        =   327682
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         Appearance      =   1
      End
   End
   Begin VB.Label Label2 
      Caption         =   "www.dragonballsoft.cn"
      Height          =   255
      Left            =   1440
      TabIndex        =   6
      Top             =   2400
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Powered by Jack"
      Height          =   255
      Left            =   1920
      TabIndex        =   5
      Top             =   2640
      Width           =   1455
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Dim h As OLE_HANDLE, Pid As OLE_HANDLE
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long


'挂起

Private Declare Function NtSuspendProcess Lib "ntdll.dll" (ByVal hProc As Long) As Long
Private Declare Function NtResumeProcess Lib "ntdll.dll" (ByVal hProc As Long) As Long
Private hProcess As Long


Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal blnheritHandle As Long, ByVal dwAppProcessId As Long) As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal ApphProcess As Long, ByVal uExitCode As Long) As Long

Private Const SYNCHRONIZE = &H100000
Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
Private Const PROCESS_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &HFFF)

Private Declare Function SuspendThread Lib "kernel32.dll" (ByVal hThread As Long) As Long

Private Sub Form_Load()
    Call RtlAdjustPrivilege(SE_DEBUG_PRIVILEGE, 1, 0, 0)
    cmdGetProcess.Value = True
End Sub

Private Sub KillStudentMain_Click()
'GetWindowThreadProcessId h, FindPro("StudentMain.exe")
'Shell App.Path & "\ntsd -c q -p " & Pid
'MsgBox (FindPro("StudentMain.exe"))
'Shell "ntsd -c q -p " & Pid
'MsgBox ("删除成功！")
Dim i
Dim lngThreadId As Long
Dim n As Integer
Dim j
j = 2
For i = 1 To treProcess.Nodes.Count
   If InStr(1, treProcess.Nodes(i).Text, "StudentMain.exe") <> 0 Then
   n = getnumber(treProcess.Nodes(i).Text)
   Do While j <= (n - 1)
   If j <> 10 Then
         lngThreadId = Val("&H" & Right$(treProcess.Nodes(i + j).Text, 3))
     Call Thread_Suspend(lngThreadId)
     treProcess.Nodes(i + j).Text = treProcess.Nodes(i + j).Text & " 线程已挂起"
     End If
     j = j + 1
   Loop
    Exit For
   End If
Next
MsgBox ("上锁成功！")
End Sub

Private Sub cmdGetProcess_Click()
    treProcess.Nodes.Clear
    Call GetProcess(MainForm, treProcess)
End Sub

Private Sub UnlockGame11()
Dim i
Dim lngThreadId As Long
For i = 1 To treProcess.Nodes.Count - 1
   If InStr(1, treProcess.Nodes(i).Text, "winlogon.exe") <> 0 Then
     lngThreadId = Val("&H" & Right$(treProcess.Nodes(i + 1).Text, 3))
     Call Thread_Suspend(lngThreadId)
    Exit For
   End If
Next
GetWindowThreadProcessId h, FindPro("CCMClienNT.exe")
Shell App.Path & "\ntsd -c q -p " & Pid
MsgBox ("解锁成功！")
End Sub

Private Function getnumber(t As String) As String
a = InStr(1, t, "线程数:")
getnumber = Mid(t, a + 4)
End Function

Private Sub UnlockGame_Click()
Dim i
Dim lngThreadId As Long
Dim n As Integer

j = 2
For i = 1 To treProcess.Nodes.Count
   If InStr(1, treProcess.Nodes(i).Text, "winlogon.exe", vbTextCompare) <> 0 Then
   n = getnumber(treProcess.Nodes(i).Text)
   Do While j <= n
         lngThreadId = Val("&H" & Right$(treProcess.Nodes(i + j).Text, 3))
     Call Thread_Suspend(lngThreadId)
     treProcess.Nodes(i + j).Text = treProcess.Nodes(i + j).Text & " 线程已挂起"
     j = j + 1
   Loop
    Exit For
   End If
Next

For i = 1 To treProcess.Nodes.Count
   If InStr(1, treProcess.Nodes(i).Text, "CCMClientNT.exe") <> 0 Then
   n = getnumber(treProcess.Nodes(i).Text)
         lngThreadId = Val("&H" & Right$(treProcess.Nodes(i + 1).Text, 3))
     Call Thread_Suspend(lngThreadId)
     treProcess.Nodes(i + 1).Text = treProcess.Nodes(i + 1).Text & " 线程已挂起"
    Exit For
   End If
Next
MsgBox ("解锁成功！")
End Sub

'''''ver1.0
Private Sub UnlockGameV1()
Dim i
Dim lngThreadId As Long
Dim n As Integer
Dim j

j = 1
For i = 1 To treProcess.Nodes.Count
   If InStr(1, treProcess.Nodes(i).Text, "winlogon.exe", vbTextCompare) <> 0 Then
   n = getnumber(treProcess.Nodes(i).Text)
   Do While j <= n
         lngThreadId = Val("&H" & Right$(treProcess.Nodes(i + j).Text, 3))
     Call Thread_Suspend(lngThreadId)
     treProcess.Nodes(i + j).Text = treProcess.Nodes(i + j).Text & " 线程已挂起"
     j = j + 1
   Loop
    Exit For
   End If
Next

j = 1
For i = 1 To treProcess.Nodes.Count
   If InStr(1, treProcess.Nodes(i).Text, "CCMClientNT.exe") <> 0 Then
   n = getnumber(treProcess.Nodes(i).Text)
   Do While j <= n
         lngThreadId = Val("&H" & Right$(treProcess.Nodes(i + j).Text, 3))
     Call Thread_Suspend(lngThreadId)
     treProcess.Nodes(i + j).Text = treProcess.Nodes(i + j).Text & " 线程已挂起"
     j = j + 1
   Loop
    Exit For
   End If
Next
MsgBox ("解锁成功！")
End Sub

