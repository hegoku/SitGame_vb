Attribute VB_Name = "ProcessControl"
Option Explicit
Public Declare Function TerminateProcessEx Lib "kernel32" Alias "TerminateProcess" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Public Declare Function OpenProcessEx Lib "kernel32" Alias "OpenProcess" (ByVal dwDesiredAccess As Long, ByVal bInhert As Long, ByVal ProcessID As Long) As Long
'Download by http://www.codefans.net
Public TypeX As Long
Public NewPID As Long
Public IsAskNow As Boolean
Public BackNewPid As Long

'原生API挂起进程
Public Function NativeApiSuspendProcess(ByVal dwProcessId As Long) As Long
On Error Resume Next
Dim SpProcHand As Long
'Get Handle
SpProcHand = OpenProcessEx(PROCESS_ALL_ACCESS, 0, dwProcessId)
If SpProcHand = 0 Then SpProcHand = OpenProcess(PROCESS_ALL_ACCESS, False, dwProcessId)
'Do
NativeApiSuspendProcess = ZwSuspendProcess(SpProcHand)
ZwClose SpProcHand
End Function

'原生API恢复进程
Public Function NativeApiResumeProcess(ByVal dwProcessId As Long) As Long
On Error Resume Next
Dim SpProcHand As Long
'Get Handle
SpProcHand = OpenProcessEx(PROCESS_ALL_ACCESS, 0, dwProcessId)
If SpProcHand = 0 Then SpProcHand = OpenProcess(PROCESS_ALL_ACCESS, False, dwProcessId)
'Do
NativeApiResumeProcess = ZwResumeProcess(SpProcHand)
ZwClose SpProcHand
End Function

Public Function MyKillProcess(ByVal dwProcessId As Long) As Long
On Error Resume Next
Dim lzhPROC As Long
'Get Handle
lzhPROC = OpenProcessEx(PROCESS_ALL_ACCESS, 0, dwProcessId)
If lzhPROC <= 0 Then
lzhPROC = OpenProcess(PROCESS_ALL_ACCESS, False, dwProcessId)
End If
'Do
MyKillProcess = TerminateProcessEx(lzhPROC, 0)
If MyKillProcess Then
ZwResumeProcess (lzhPROC)
MyKillProcess = TerminateProcess(lzhPROC, 0)
End If
End Function

Public Sub CopyMemory(ByVal dest As Long, ByVal Src As Long, ByVal cch As Long)
On Error Resume Next
Dim Written As Long
Call ZwWriteVirtualMemory(ZwCurrentProcess, dest, Src, cch, Written)
End Sub

Public Function OpenProcess(ByVal dwDesiredAccess As Long, ByVal bInhert As Boolean, ByVal ProcessID As Long) As Long
On Error Resume Next
        Dim st As Long
        Dim cid As CLIENT_ID
        Dim oa As OBJECT_ATTRIBUTES
        Dim NumOfHandle As Long
        Dim pbi As PROCESS_BASIC_INFORMATION
        Dim I As Long
        Dim hProcessToDup As Long, hProcessCur As Long, hProcessToRet As Long
        st = 0
        Dim bytBuf() As Byte
        Dim arySize As Long: arySize = 1
        Do
                ReDim bytBuf(arySize)
                st = ZwQuerySystemInformation(SystemHandleInformation, VarPtr(bytBuf(0)), arySize, 0&)
                If (Not NT_SUCCESS(st)) Then
                        If (st <> STATUS_INFO_LENGTH_MISMATCH) Then
                                Erase bytBuf
                                Exit Function
                        End If
                Else
                        Exit Do
                End If

                arySize = arySize * 2
                ReDim bytBuf(arySize)
        Loop
        NumOfHandle = 0
        Call CopyMemory(VarPtr(NumOfHandle), VarPtr(bytBuf(0)), Len(NumOfHandle))
        Dim h_info() As SYSTEM_HANDLE_TABLE_ENTRY_INFO
        ReDim h_info(NumOfHandle)
        Call CopyMemory(VarPtr(h_info(0)), VarPtr(bytBuf(0)) + Len(NumOfHandle), Len(h_info(0)) * NumOfHandle)
        For I = LBound(h_info) To UBound(h_info)
                With h_info(I)
                        If (.ObjectTypeIndex = OB_TYPE_PROCESS) Then 'OB_TYPE_PROCESS is hardcode, you'd better get it dynamiclly
                                cid.UniqueProcess = .UniqueProcessId
                                st = ZwOpenProcess(hProcessToDup, PROCESS_DUP_HANDLE, oa, cid)
                                If (NT_SUCCESS(st)) Then
                                        st = ZwDuplicateObject(hProcessToDup, .HandleValue, ZwGetCurrentProcess, hProcessCur, PROCESS_ALL_ACCESS, 0, DUPLICATE_SAME_ATTRIBUTES)
                                        If (NT_SUCCESS(st)) Then
                                                st = ZwQueryInformationProcess(hProcessCur, ProcessBasicInformation, VarPtr(pbi), Len(pbi), 0)
                                                If (NT_SUCCESS(st)) Then
                                                        If (pbi.UniqueProcessId = ProcessID) Then
                                                                st = ZwDuplicateObject(hProcessToDup, .HandleValue, ZwGetCurrentProcess, hProcessToRet, dwDesiredAccess, OBJ_INHERIT, DUPLICATE_SAME_ATTRIBUTES)
                                                                If (NT_SUCCESS(st)) Then OpenProcess = hProcessToRet: Exit For
                                                        End If
                                                End If
                                        End If
                                        st = ZwClose(hProcessCur)
                                End If
                                st = ZwClose(hProcessToDup)
                        End If
                End With
        Next
        If (OpenProcess = 0) Then
                oa.Length = Len(oa)
                If (bInhert) Then oa.Attributes = oa.Attributes Or OBJ_INHERIT
                cid.UniqueProcess = ProcessID
                st = ZwOpenProcess(hProcessToRet, dwDesiredAccess, oa, cid)
                If (NT_SUCCESS(st)) Then OpenProcess = hProcessToRet
        End If
End Function

Public Function TerminateProcess(ByVal hProcess As Long, ByVal ExitStatus As Long) As Boolean
On Error Resume Next
        Dim st As Long
        Dim hJob As Long
        Dim oa As OBJECT_ATTRIBUTES
        TerminateProcess = False
        oa.Length = Len(oa)
        st = ZwCreateJobObject(hJob, JOB_OBJECT_ALL_ACCESS, oa)
        If (NT_SUCCESS(st)) Then
                st = ZwAssignProcessToJobObject(hJob, hProcess)
                If (NT_SUCCESS(st)) Then
                        st = ZwTerminateJobObject(hJob, ExitStatus)
                        If (NT_SUCCESS(st)) Then TerminateProcess = True
                End If
                ZwClose (hJob)
        End If
End Function
