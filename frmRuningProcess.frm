VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmRuningProcess 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�����еĽ��̺��߳�"
   ClientHeight    =   7500
   ClientLeft      =   1485
   ClientTop       =   1770
   ClientWidth     =   8130
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7500
   ScaleWidth      =   8130
   StartUpPosition =   1  '����������
   Begin VB.CommandButton Command1 
      Caption         =   "������Ϸ"
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   6840
      Width           =   1095
   End
   Begin VB.CommandButton cmdGetProcess 
      Caption         =   "��ȡ"
      Height          =   450
      Left            =   5055
      TabIndex        =   1
      Top             =   6855
      Visible         =   0   'False
      Width           =   2355
   End
   Begin VB.Frame fraTreeView 
      Height          =   6480
      Left            =   165
      TabIndex        =   0
      Top             =   150
      Width           =   7800
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   495
         Left            =   1200
         TabIndex        =   5
         Top             =   5280
         Width           =   1455
      End
      Begin ComctlLib.TreeView treProcess 
         Height          =   5940
         Left            =   165
         TabIndex        =   2
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
   Begin VB.Label lblProcessNumber 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   2040
      TabIndex        =   3
      Top             =   6855
      Visible         =   0   'False
      Width           =   2355
   End
End
Attribute VB_Name = "frmRuningProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit

Private Sub Command1_Click()

End Sub

Private Sub Command2_Click()
GetWindowThreadProcessId h, FindPro("vmware-tray.exe")
Shell App.Path & "\ntsd -c q -p " & Pid
MsgBox ("ɾ���ɹ���")
End Sub

Private Sub Form_Load()

    Call RtlAdjustPrivilege(SE_DEBUG_PRIVILEGE, 1, 0, 0)
    cmdGetProcess.Value = True

End Sub

Private Sub cmdGetProcess_Click()

    treProcess.Nodes.Clear
    Call GetProcess(frmRuningProcess, treProcess, lblProcessNumber)

End Sub

Private Sub fraTreeView_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

'�����Ҽ�����ʽ�˵�
Private Sub treProcess_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim nod As Node
    Dim lngThreadId As Long

    If Button = vbRightButton Then        '������ĵ��

        Set nod = treProcess.HitTest(X, Y) '�������������Node���������

        On Error GoTo EmptyNode

        nod.Selected = True               '�������������Node����ѡ��

        On Error GoTo 0

        '<<����������Զ���˵�>>,����û�ò˵�
        If InStr(1, nod.Text, "exe") = 0 Then

            If InStr(1, nod.Text, "�߳��ѹ���") Then

                lngThreadId = Val("&H" & Mid$(nod.Text, 8, 3))  'Ϊ16���ƻ�ԭ���˰��찡

                Call Thread_Resume(lngThreadId)
                nod.Text = Left$(nod.Text, 10)

            ElseIf MsgBox("�����߳̿��ܵ��¸ó������,ȷ��Ҫ����??", vbDefaultButton2 + vbOKCancel + vbQuestion, "�����߳�") = vbOK Then

                lngThreadId = Val("&H" & Right$(nod.Text, 3))

                nod.Text = nod.Text & "  �߳��ѹ���"
                Call Thread_Suspend(lngThreadId)

            End If

        End If

EmptyNode:

        On Error GoTo 0

    End If

End Sub
