Attribute VB_Name = "mod_Process"
Option Explicit

Public Declare Function RtlAdjustPrivilege Lib "ntdll.dll" (ByVal Privilege As Long, ByVal Enable As Long, ByVal CurrentThread As Long, Enabled As Long) As Long
Public Const SE_DEBUG_PRIVILEGE = &H14

'************************************* ����ö�ٽ���*********************************
'CreateToolhelpSnapshotΪָ���Ľ��̡�����ʹ�õĶ�[HEAP]��ģ��[MODULE]���߳�[THREAD]������һ������[snapshot]��

'������
'dwFlags

'TH32CS_INHERIT -�������վ���ǿɼ̳е�
'TH32CS_SNAPall -�ڿ����а���ϵͳ�����еĽ��̺��߳�
'TH32CS_SNAPheaplist -�ڿ����а�����th32ProcessID��ָ���Ľ��̵����еĶ�
'TH32CS_SNAPmodule -�ڿ����а�����th32ProcessID��ָ���Ľ��̵����е�ģ��
'TH32CS_SNAPPROCESS -�ڿ����а���ϵͳ�����еĽ���
'TH32CS_SNAPthread -�ڿ����а���ϵͳ�����е��߳�

'th32ProcessID

'[����]ָ����Ҫ���յĽ���ID������ò���Ϊ0��ʾ���յ�ǰ���̡�
'�ò���ֻ����������TH32CS_SNAPHEAPLIST��TH32CS_SNAPMOUDLE�����Ч������������¸ò��������ԣ����еĽ��̶��ᱻ���ա�
Private Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long

'���ϵͳ�����еĵ�һ�����̵���Ϣ
Private Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long

'���ϵͳ�����е���һ�����̵���Ϣ
Private Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long

'�رվ��
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Type PROCESSENTRY32
    dwSize As Long                 '�ṹ��С
    cntUsage As Long               '�˽��̵����ü���
    th32ProcessID As Long          '����ID
    th32DefaultHeapID As Long      '����Ĭ�϶�ID
    th32ModuleID As Long           '����ģ��ID
    cntThreads As Long             '�˽��̿������̼߳���
    th32ParentProcessID As Long    '������ID
    pcPriClassBase As Long         '�߳�����Ȩ
    dwFlags As Long                '����
    szExeFile As String * 260      '����ȫ��
End Type

Private Const TH32CS_SNAPPROCESS = &H2     'TH32CS_SNAPPROCESS -�ڿ����а���ϵͳ�����еĽ���
Private Const TH32CS_SNAPmodule = &H8      '��ʾ����Ϊ��th32ProcessID����ָ���Ľ��̵��õ�����ģ��

'Thread32First����߳̿����еĵ�һ���̵߳���Ϣ
'������

'hSnapShot��Ȼ��CreateToolhelpSnapshot��ȡ���߳̿���
'lpte��THREADENTRY32�ṹ��Apiִ�й��̻��޸�THREADENTRY32�ṹ��Ա����ֵ��ע����ByRef lpte As THREADENTRY32
Private Declare Function Thread32First Lib "kernel32.dll" (ByVal hSnapShot As Long, ByRef lpte As THREADENTRY32) As Boolean
Private Declare Function Thread32Next Lib "kernel32.dll" (ByVal hSnapShot As Long, ByRef lpte As THREADENTRY32) As Boolean

Private Type THREADENTRY32
    dwSize As Long                 '�ṹ��С
    cntUsage As Long               '���̵߳����ü���
    th32ThreadID As Long           '�߳�ID
    th32OwnerProcessID As Long     '���߳�ID
    tpBasePri As Long              '��ʼ���̵߳����ȼ�
    tpDeltaPri As Long             '�޸��̵߳����ȼ�
    dwFlags As Long                '����ֵ��δʹ��
End Type

'OpenThread��һ���߳�

'������
'dwDesiredAccess  ��Ҫ��Ȩ��,����ֻ��ҪTHREAD_SUSPEND_RESUME����
'[in] Indicates whether the returned handle is to be inherited by a new process created by the current process. If this parameter is TRUE, the new process will inherit the handle.
'bInheritHandle   ��ǰ���̲����½����Ƿ���Լ̳д򿪵��߳̾��,����ѡfalse,Ӣ�ĶԳ���Ա��һ���鰡��
'dwThreadId       ����̵߳�ID��
Private Declare Function OpenThread Lib "kernel32.dll" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Boolean, ByVal dwThreadId As Long) As Long

'������ApiӦ�ú�������⣬����Ӣ�ĺ���֪���ǻָ��͹���һ���߳�,����hThread���߳̾��,��OpenThread��ȡ
Private Declare Function ResumeThread Lib "kernel32.dll" (ByVal hThread As Long) As Long
Private Declare Function SuspendThread Lib "kernel32.dll" (ByVal hThread As Long) As Long

'Enables the use of the thread handle in the SuspendThread or ResumeThread function to suspend and resume the thread.
'����Msdn2001 Oct�汾�Ľ��ͣ���˼ΪSuspendThread��ResumeThread����ʹ���߳̾����ͣ�ͻָ���ǰ�߳�,������ֻװMsdn6.0 for VB��ԶԶ�����ģ�����Api��Msdn2001�Ժ�������ϵġ�
Private Const THREAD_SUSPEND_RESUME As Long = &H2
Private Const TH32CS_SNAPthread = &H4        'TH32CS_SNAPthread -�ڿ����а���ϵͳ�����е��߳�

'*************************************************************************
'**�� �� ���� GetProcess
'**��    �룺 ByVal frmRuningProcess(Form)  - ֱ�Ӵ����������
'**     ��    ByVal treProcess(TreeView)    -
'**     ��    ByVal lblProcessNumber(Label) -
'**��    ���� ��
'**���������������������ṹ
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ� Mr.David
'**��    �ڣ� 2007-11-27 14:09:37
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ���� V1.0.0
'*************************************************************************
Public Sub GetProcess(ByVal frmRuningProcess As Form, ByVal treProcess As TreeView)

    Dim lngResult As Long
    Dim hSnapShot As Long           '���̿��վ��
    Dim hTSnapshot As Long          '�߳̿��վ��

    Dim strTreTxt As String
    Dim strTreKey As String
    Dim strProcName As String * 35 '����ǿ��ָ��35�ַ���Ϊ�˶�����

    Dim objTreNode As Node          '���ڵ�
    Dim lngProcCount As Long        '��������
    Dim lngPos As String            'Chr(0)λ��

    Dim blnAlreadyGetArray As Boolean '��һ�λ�ȡThread32First�Ϳ��Ա����߳��ˣ��Ժ����ˢ�¶������ٴ��˷�ʱ�����

    Dim TD As THREADENTRY32         '�ṹ����
    Dim PEE As PROCESSENTRY32

    Dim lngIndex As Long            '�߳�����
    Dim i As Long

    Dim astrThreadNum() As String   '�洢�߳���
    Dim lngThreadCount As Long      '�̼߳���

    On Error GoTo PROC_ERR

    PEE.dwSize = Len(PEE)
    TD.dwSize = Len(TD)

    hSnapShot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0)   '�������н���

    lngResult = ProcessFirst(hSnapShot, PEE)                    '��ȡ��һ����

    '��ѭ����ȡ������
    Do While lngResult <> 0

        lngProcCount = lngProcCount + 1                         '�ۼƽ�����

        hTSnapshot = CreateToolhelpSnapshot(TH32CS_SNAPthread, PEE.th32ProcessID)  '�����߳̿���

        strTreKey = PEE.szExeFile   '�������ؼ��ֲ��õ�PEE.szExeFile��ע������ͬһ������ʵ����ת��PROC_ERR����
        strTreTxt = PEE.szExeFile

        strProcName = Left$(strTreTxt, InStr(1, strTreTxt, Chr$(0)) - 1)
        Set objTreNode = treProcess.Nodes.Add(, , strTreKey, strProcName & "�߳���:" & PEE.cntThreads)    '���������ǽ�����&�߳���

        lngResult = Thread32First(hTSnapshot, TD)              '��ȡ��һ���߳�

        '����ѭ����ȡȫ���߳�
        Do While lngResult <> 0

            If blnAlreadyGetArray = True Then Exit Do          '����Ѿ�������

            lngThreadCount = lngThreadCount + 1                '�̼߳���
            ReDim Preserve astrThreadNum(lngThreadCount)       '���·����߳�����

            astrThreadNum(lngThreadCount) = "0x00000" & Hex$(TD.th32ThreadID)   '��ʽ�������߳�ID
            lngResult = Thread32Next(hTSnapshot, TD)           '��ȡ��һ�߳�

        Loop

        blnAlreadyGetArray = True
        Call CloseHandle(hTSnapshot)                           '�ر��߳̿��վ��

        '�Դ���������Ϊ�߳�����һ�α����������ǲ���ģ�����ȡ�߳�ʱ��ֻ��֪���߳��������ϴ�����λ�þͿ��ԣ�ͨ�������㷨�����������������
        For i = lngIndex + 1 To lngIndex + PEE.cntThreads

            Set objTreNode = treProcess.Nodes.Add(strTreKey, tvwChild, , astrThreadNum(i))   '�г��߳�����

        Next

        'objTreNode.EnsureVisible 'չ����֧,����ѡ�����
        lngResult = ProcessNext(hSnapShot, PEE)               '��ȡ��һ����
        lngIndex = i - 1

    Loop

    Erase astrThreadNum$()
    Call CloseHandle(hSnapShot)                               '�رս��̿��վ��

    

    Exit Sub

PROC_ERR:

    '������������еĹؼ��ֲ�Ψһ����ؼ���������������Ntϵͳ����ڶ��Svchost���̣��˹ؼ������ﲻ��Ҫ����㴦��һ��
    If Err.Number = 35602 Then

        strTreKey = strTreKey & "1"
        Resume

    Else

        Resume Next

    End If

End Sub

Public Function Thread_Suspend(T_ID As Long) As Long            '�����߳�

    Dim hThread As Long

    hThread = OpenThread(THREAD_SUSPEND_RESUME, False, T_ID)
    Thread_Suspend = SuspendThread(hThread)

    Call CloseHandle(hThread)

End Function

Public Function Thread_Resume(T_ID As Long) As Long             '�ָ��߳�

    Dim hThread As Long
    Dim lSuspendCount As Long

    hThread = OpenThread(THREAD_SUSPEND_RESUME, False, T_ID)
    Thread_Resume = ResumeThread(hThread)

    Call CloseHandle(hThread)

End Function
