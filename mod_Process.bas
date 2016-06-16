Attribute VB_Name = "mod_Process"
Option Explicit

Public Declare Function RtlAdjustPrivilege Lib "ntdll.dll" (ByVal Privilege As Long, ByVal Enable As Long, ByVal CurrentThread As Long, Enabled As Long) As Long
Public Const SE_DEBUG_PRIVILEGE = &H14

'************************************* 用于枚举进程*********************************
'CreateToolhelpSnapshot为指定的进程、进程使用的堆[HEAP]、模块[MODULE]、线程[THREAD]）建立一个快照[snapshot]。

'参数：
'dwFlags

'TH32CS_INHERIT -声明快照句柄是可继承的
'TH32CS_SNAPall -在快照中包含系统中所有的进程和线程
'TH32CS_SNAPheaplist -在快照中包含在th32ProcessID中指定的进程的所有的堆
'TH32CS_SNAPmodule -在快照中包含在th32ProcessID中指定的进程的所有的模块
'TH32CS_SNAPPROCESS -在快照中包含系统中所有的进程
'TH32CS_SNAPthread -在快照中包含系统中所有的线程

'th32ProcessID

'[输入]指定将要快照的进程ID。如果该参数为0表示快照当前进程。
'该参数只有在设置了TH32CS_SNAPHEAPLIST或TH32CS_SNAPMOUDLE后才有效，在其他情况下该参数被忽略，所有的进程都会被快照。
Private Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long

'获得系统快照中的第一个进程的信息
Private Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long

'获得系统快照中的下一个进程的信息
Private Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long

'关闭句柄
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Type PROCESSENTRY32
    dwSize As Long                 '结构大小
    cntUsage As Long               '此进程的引用计数
    th32ProcessID As Long          '进程ID
    th32DefaultHeapID As Long      '进程默认堆ID
    th32ModuleID As Long           '进程模块ID
    cntThreads As Long             '此进程开启的线程计数
    th32ParentProcessID As Long    '父进程ID
    pcPriClassBase As Long         '线程优先权
    dwFlags As Long                '保留
    szExeFile As String * 260      '进程全名
End Type

Private Const TH32CS_SNAPPROCESS = &H2     'TH32CS_SNAPPROCESS -在快照中包含系统中所有的进程
Private Const TH32CS_SNAPmodule = &H8      '表示对象为由th32ProcessID参数指定的进程调用的所有模块

'Thread32First获得线程快照中的第一个线程的信息
'参数表

'hSnapShot当然是CreateToolhelpSnapshot获取的线程快照
'lpte是THREADENTRY32结构，Api执行过程会修改THREADENTRY32结构成员变量值，注意是ByRef lpte As THREADENTRY32
Private Declare Function Thread32First Lib "kernel32.dll" (ByVal hSnapShot As Long, ByRef lpte As THREADENTRY32) As Boolean
Private Declare Function Thread32Next Lib "kernel32.dll" (ByVal hSnapShot As Long, ByRef lpte As THREADENTRY32) As Boolean

Private Type THREADENTRY32
    dwSize As Long                 '结构大小
    cntUsage As Long               '此线程的引用计数
    th32ThreadID As Long           '线程ID
    th32OwnerProcessID As Long     '父线程ID
    tpBasePri As Long              '初始化线程的优先级
    tpDeltaPri As Long             '修改线程的优先级
    dwFlags As Long                '保留值，未使用
End Type

'OpenThread打开一个线程

'参数表
'dwDesiredAccess  需要的权限,本文只需要THREAD_SUSPEND_RESUME即可
'[in] Indicates whether the returned handle is to be inherited by a new process created by the current process. If this parameter is TRUE, the new process will inherit the handle.
'bInheritHandle   当前进程产生新进程是否可以继承打开的线程句柄,这里选false,英文对程序员是一大考验啊。
'dwThreadId       需打开线程的ID号
Private Declare Function OpenThread Lib "kernel32.dll" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Boolean, ByVal dwThreadId As Long) As Long

'这两个Api应该很容易理解，看其英文含义知道是恢复和挂起一个线程,参数hThread是线程句柄,由OpenThread获取
Private Declare Function ResumeThread Lib "kernel32.dll" (ByVal hThread As Long) As Long
Private Declare Function SuspendThread Lib "kernel32.dll" (ByVal hThread As Long) As Long

'Enables the use of the thread handle in the SuspendThread or ResumeThread function to suspend and resume the thread.
'这是Msdn2001 Oct版本的解释，意思为SuspendThread和ResumeThread函数使用线程句柄暂停和恢复当前线程,所以你只装Msdn6.0 for VB是远远不够的，多数Api是Msdn2001以后才有资料的。
Private Const THREAD_SUSPEND_RESUME As Long = &H2
Private Const TH32CS_SNAPthread = &H4        'TH32CS_SNAPthread -在快照中包含系统中所有的线程

'*************************************************************************
'**函 数 名： GetProcess
'**输    入： ByVal frmRuningProcess(Form)  - 直接传入各对象名
'**     ：    ByVal treProcess(TreeView)    -
'**     ：    ByVal lblProcessNumber(Label) -
'**输    出： 无
'**功能描述：建立进程树结构
'**全局变量：
'**调用模块：
'**作    者： Mr.David
'**日    期： 2007-11-27 14:09:37
'**修 改 人：
'**日    期：
'**版    本： V1.0.0
'*************************************************************************
Public Sub GetProcess(ByVal frmRuningProcess As Form, ByVal treProcess As TreeView)

    Dim lngResult As Long
    Dim hSnapShot As Long           '进程快照句柄
    Dim hTSnapshot As Long          '线程快照句柄

    Dim strTreTxt As String
    Dim strTreKey As String
    Dim strProcName As String * 35 '这里强行指定35字符是为了对齐树

    Dim objTreNode As Node          '树节点
    Dim lngProcCount As Long        '进程总数
    Dim lngPos As String            'Chr(0)位置

    Dim blnAlreadyGetArray As Boolean '第一次获取Thread32First就可以遍历线程了，以后除非刷新都无需再次浪费时间遍历

    Dim TD As THREADENTRY32         '结构引用
    Dim PEE As PROCESSENTRY32

    Dim lngIndex As Long            '线程索引
    Dim i As Long

    Dim astrThreadNum() As String   '存储线程名
    Dim lngThreadCount As Long      '线程计数

    On Error GoTo PROC_ERR

    PEE.dwSize = Len(PEE)
    TD.dwSize = Len(TD)

    hSnapShot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0)   '快照所有进程

    lngResult = ProcessFirst(hSnapShot, PEE)                    '获取第一进程

    '外循环读取进程名
    Do While lngResult <> 0

        lngProcCount = lngProcCount + 1                         '累计进程数

        hTSnapshot = CreateToolhelpSnapshot(TH32CS_SNAPthread, PEE.th32ProcessID)  '建立线程快照

        strTreKey = PEE.szExeFile   '对树根关键字采用的PEE.szExeFile，注意运行同一程序多个实例会转入PROC_ERR处理
        strTreTxt = PEE.szExeFile

        strProcName = Left$(strTreTxt, InStr(1, strTreTxt, Chr$(0)) - 1)
        Set objTreNode = treProcess.Nodes.Add(, , strTreKey, strProcName & "线程数:" & PEE.cntThreads)    '进程树根是进程名&线程数

        lngResult = Thread32First(hTSnapshot, TD)              '获取第一个线程

        '初次循环获取全部线程
        Do While lngResult <> 0

            If blnAlreadyGetArray = True Then Exit Do          '如果已经遍历过

            lngThreadCount = lngThreadCount + 1                '线程计数
            ReDim Preserve astrThreadNum(lngThreadCount)       '重新分配线程数组

            astrThreadNum(lngThreadCount) = "0x00000" & Hex$(TD.th32ThreadID)   '格式化保存线程ID
            lngResult = Thread32Next(hTSnapshot, TD)           '获取下一线程

        Loop

        blnAlreadyGetArray = True
        Call CloseHandle(hTSnapshot)                           '关闭线程快照句柄

        '脑袋不大灵活，因为线程数组一次遍历后内容是不变的，进程取线程时，只需知道线程总数和上次数组位置就可以，通过索引算法，这里我是这样算的
        For i = lngIndex + 1 To lngIndex + PEE.cntThreads

            Set objTreNode = treProcess.Nodes.Add(strTreKey, tvwChild, , astrThreadNum(i))   '列出线程子树

        Next

        'objTreNode.EnsureVisible '展开分支,可以选用这句
        lngResult = ProcessNext(hSnapShot, PEE)               '获取下一进程
        lngIndex = i - 1

    Loop

    Erase astrThreadNum$()
    Call CloseHandle(hSnapShot)                               '关闭进程快照句柄

    

    Exit Sub

PROC_ERR:

    '如果发生集合中的关键字不唯一，则关键字重命名，比如Nt系统会存在多个Svchost进程，此关键字这里不重要，随便处理一下
    If Err.Number = 35602 Then

        strTreKey = strTreKey & "1"
        Resume

    Else

        Resume Next

    End If

End Sub

Public Function Thread_Suspend(T_ID As Long) As Long            '挂起线程

    Dim hThread As Long

    hThread = OpenThread(THREAD_SUSPEND_RESUME, False, T_ID)
    Thread_Suspend = SuspendThread(hThread)

    Call CloseHandle(hThread)

End Function

Public Function Thread_Resume(T_ID As Long) As Long             '恢复线程

    Dim hThread As Long
    Dim lSuspendCount As Long

    hThread = OpenThread(THREAD_SUSPEND_RESUME, False, T_ID)
    Thread_Resume = ResumeThread(hThread)

    Call CloseHandle(hThread)

End Function
