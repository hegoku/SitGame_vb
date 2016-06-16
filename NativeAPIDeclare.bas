Attribute VB_Name = "NativeAPIDeclare"
Option Explicit
'Download by http://www.codefans.net
'// ByRef AllocationSize As LARGE_INTEGER  optional
Public Declare Function ZwCreateFile _
                Lib "NTDLL.DLL" (ByRef FileHandle As Long, _
                                 ByVal DesiredAccess As Long, _
                                 ByRef ObjectAttributes As OBJECT_ATTRIBUTES, _
                                 ByRef IoStatusBlock As IO_STATUS_BLOCK, _
                                 ByVal AllocationSize As Long, _
                                 ByVal FileAttributes As Long, _
                                 ByVal ShareAccess As Long, _
                                 ByVal CreateDisposition As Long, _
                                 ByVal CreateOptions As Long, _
                                 ByVal EaBuffer As Long, _
                                 ByVal EaLength As Long) As Long
Public Type IO_STATUS_BLOCK
        Pointer As Long
        pInformation As Long
End Type
Public Declare Function ZwQueryInformationProcess _
               Lib "NTDLL.DLL" (ByVal ProcessHandle As Long, _
                                ByVal ProcessInformationClass As PROCESSINFOCLASS, _
                                ByVal ProcessInformation As Long, _
                                ByVal ProcessInformationLength As Long, _
                                ByRef ReturnLength As Long) As Long
Public Declare Function ZwSetInformationProcess _
               Lib "NTDLL.DLL" (ByVal ProcessHandle As Long, _
                                ByVal ProcessInformationClass As PROCESSINFOCLASS, _
                                ByVal ProcessInformation As Long, _
                                ByVal ProcessInformationLength As Long) As Long
Public Const FILE_OPEN As Long = 1
Public Type LARGE_INTEGER
        lowpart As Long
        highpart As Long
End Type
Public Enum PROCESSINFOCLASS
        ProcessBasicInformation
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
        ProcessIoPortHandlers           '// Note: this is kernel mode only
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
        MaxProcessInfoClass             '// MaxProcessInfoClass should always be the last enum
End Enum

Public Type PROCESS_BASIC_INFORMATION
        ExitStatus As Long 'NTSTATUS
        PebBaseAddress As Long 'PPEB
        AffinityMask As Long 'ULONG_PTR
        BasePriority As Long 'KPRIORITY
        UniqueProcessId As Long 'ULONG_PTR
        InheritedFromUniqueProcessId As Long 'ULONG_PTR
End Type

Public Type PROCESS_IMAGE_FILENAME '//ProcessImageFileName will return this :)
        Length As Long
        MaxLength As Long
        pBuffer As String * 256 'PWSTR Pointer to a null-terminated string of 16-bit
End Type

Public Type UNICODE_STRING '//ProcessImageFileName will return this :)
        Length As Integer
        MaxLength As Integer
        pBuffer As Long 'PWSTR Pointer to a null-terminated string of 16-bit
End Type

Public Type ANSI_STRING '//ProcessImageFileName will return this :)
        Length As Integer
        MaxLength As Integer
        pBuffer As Long
End Type

Public Declare Function ZwQuerySystemInformation _
               Lib "NTDLL.DLL" (ByVal SystemInformationClass As SYSTEM_INFORMATION_CLASS, _
                                ByVal pSystemInformation As Long, _
                                ByVal SystemInformationLength As Long, _
                                ByRef ReturnLength As Long) As Long
Public Declare Function ZwSetSystemInformation _
                Lib "NTDLL.DLL" (ByVal SystemInformationClass As SYSTEM_INFORMATION_CLASS, _
                                 ByVal pSystemInformation As Long, _
                                 ByVal SystemInformationLength As Long) As Long
Public Enum SYSTEM_INFORMATION_CLASS
        SystemBasicInformation
        SystemProcessorInformation             '// obsolete...delete
        SystemPerformanceInformation
        SystemTimeOfDayInformation
        SystemPathInformation
        SystemProcessInformation
        SystemCallCountInformation
        SystemDeviceInformation
        SystemProcessorPerformanceInformation
        SystemFlagsInformation
        SystemCallTimeInformation
        SystemModuleInformation
        SystemLocksInformation
        SystemStackTraceInformation
        SystemPagedPoolInformation
        SystemNonPagedPoolInformation
        SystemHandleInformation
        SystemObjectInformation
        SystemPagefileInformation
        SystemVdmInstemulInformation
        SystemVdmBopInformation
        SystemFileCacheInformation
        SystemPoolTagInformation
        SystemInterruptInformation
        SystemDpcBehaviorInformation
        SystemFullMemoryInformation
        SystemLoadGdiDriverInformation
        SystemUnloadGdiDriverInformation
        SystemTimeAdjustmentInformation
        SystemSummaryMemoryInformation
        SystemMirrorMemoryInformation
        SystemPerformanceTraceInformation
        SystemObsolete0
        SystemExceptionInformation
        SystemCrashDumpStateInformation
        SystemKernelDebuggerInformation
        SystemContextSwitchInformation
        SystemRegistryQuotaInformation
        SystemExtendServiceTableInformation
        SystemPrioritySeperation
        SystemVerifierAddDriverInformation
        SystemVerifierRemoveDriverInformation
        SystemProcessorIdleInformation
        SystemLegacyDriverInformation
        SystemCurrentTimeZoneInformation
        SystemLookasideInformation
        SystemTimeSlipNotification
        SystemSessionCreate
        SystemSessionDetach
        SystemSessionInformation
        SystemRangeStartInformation
        SystemVerifierInformation
        SystemVerifierThunkExtend
        SystemSessionProcessInformation
        SystemLoadGdiDriverInSystemSpace
        SystemNumaProcessorMap
        SystemPrefetcherInformation
        SystemExtendedProcessInformation
        SystemRecommendedSharedDataAlignment
        SystemComPlusPackage
        SystemNumaAvailableMemory
        SystemProcessorPowerInformation
        SystemEmulationBasicInformation
        SystemEmulationProcessorInformation
        SystemExtendedHandleInformation
        SystemLostDelayedWriteInformation
        SystemBigPoolInformation
        SystemSessionPoolTagInformation
        SystemSessionMappedViewInformation
        SystemHotpatchInformation
        SystemObjectSecurityMode
        SystemWatchdogTimerHandler
        SystemWatchdogTimerInformation
        SystemLogicalProcessorInformation
        SystemWow64SharedInformation
        SystemRegisterFirmwareTableInformationHandler
        SystemFirmwareTableInformation
        SystemModuleInformationEx
        SystemVerifierTriageInformation
        SystemSuperfetchInformation
        SystemMemoryListInformation
        SystemFileCacheInformationEx
        MaxSystemInfoClass  '// MaxSystemInfoClass should always be the last enum
End Enum
Public Type SYSTEM_HANDLE_TABLE_ENTRY_INFO
        UniqueProcessId As Integer
        CreatorBackTraceIndex As Integer
        ObjectTypeIndex As Byte
        HandleAttributes As Byte
        HandleValue As Integer
        pObject As Long
        GrantedAccess As Long
End Type
Public Type SYSTEM_HANDLE_INFORMATION
        NumberOfHandles As Long
        Handles(1 To 1) As SYSTEM_HANDLE_TABLE_ENTRY_INFO
End Type
Public Const STATUS_INFO_LENGTH_MISMATCH = &HC0000004
Public Enum SYSTEM_HANDLE_TYPE
        OB_TYPE_UNKNOWN = 0
        OB_TYPE_TYPE = 1
        OB_TYPE_DIRECTORY
        OB_TYPE_SYMBOLIC_LINK
        OB_TYPE_TOKEN
        OB_TYPE_PROCESS
        OB_TYPE_THREAD
        OB_TYPE_UNKNOWN_7
        OB_TYPE_EVENT
        OB_TYPE_EVENT_PAIR
        OB_TYPE_MUTANT
        OB_TYPE_UNKNOWN_11
        OB_TYPE_SEMAPHORE
        OB_TYPE_TIMER
        OB_TYPE_PROFILE
        OB_TYPE_WINDOW_STATION
        OB_TYPE_DESKTOP
        OB_TYPE_SECTION
        OB_TYPE_KEY
        OB_TYPE_PORT
        OB_TYPE_WAITABLE_PORT
        OB_TYPE_UNKNOWN_21
        OB_TYPE_UNKNOWN_22
        OB_TYPE_UNKNOWN_23
        OB_TYPE_UNKNOWN_24
        OB_TYPE_IO_COMPLETION
        OB_TYPE_FILE
End Enum
Public Type SYSTEM_PROCESS_INFORMATION
        NextEntryOffset As Long

        NumberOfThreads As Long
        SpareLi1 As LARGE_INTEGER
        SpareLi2 As LARGE_INTEGER
        SpareLi3 As LARGE_INTEGER
        CreateTime As LARGE_INTEGER
        UserTime As LARGE_INTEGER
        KernelTime As LARGE_INTEGER
        ImageName As UNICODE_STRING
        BasePriority As Long 'KPRIORITY
        UniqueProcessId As Long
        InheritedFromUniqueProcessId As Long
        HandleCount As Long
        SessionId As Long
        pPageDirectoryBase As Long '_PTR
        PeakVirtualSize As Long
        VirtualSize As Long
        PageFaultCount As Long
        PeakWorkingSetSize As Long
        WorkingSetSize As Long
        QuotaPeakPagedPoolUsage As Long
        QuotaPagedPoolUsage As Long
        QuotaPeakNonPagedPoolUsage As Long
        QuotaNonPagedPoolUsage As Long
        PagefileUsage As Long
        PeakPagefileUsage As Long
        publicPageCount As Long
        ReadOperationCount As LARGE_INTEGER
        WriteOperationCount As LARGE_INTEGER
        OtherOperationCount As LARGE_INTEGER
        ReadTransferCount As LARGE_INTEGER
        WriteTransferCount As LARGE_INTEGER
        OtherTransferCount As LARGE_INTEGER
End Type
Public Declare Function ZwDuplicateObject _
               Lib "NTDLL.DLL" (ByVal SourceProcessHandle As Long, _
                                ByVal SourceHandle As Long, _
                                ByVal TargetProcessHandle As Long, _
                                ByRef TargetHandle As Long, _
                                ByVal DesiredAccess As Long, _
                                ByVal HandleAttributes As Long, _
                                ByVal Options As Long) As Long
Public Const DUPLICATE_CLOSE_SOURCE = &H1            '// winnt
Public Const DUPLICATE_SAME_ACCESS = &H2                '// winnt
Public Const DUPLICATE_SAME_ATTRIBUTES = &H4
Public Declare Function ZwOpenProcess _
               Lib "NTDLL.DLL" (ByRef ProcessHandle As Long, _
                                ByVal AccessMask As Long, _
                                ByRef ObjectAttributes As OBJECT_ATTRIBUTES, _
                                ByRef ClientId As CLIENT_ID) As Long
Public Type OBJECT_ATTRIBUTES
        Length As Long
        RootDirectory As Long
        ObjectName As Long 'PUNICODE_STRING 的指针
        Attributes As Long
        SecurityDescriptor As Long
        SecurityQualityOfService As Long
End Type
Public Type CLIENT_ID
        UniqueProcess As Long
        UniqueThread  As Long
End Type
Public Type FLOATING_SAVE_AREA '// 0x70
     ControlWord As Long '// +0x0(0x4)
     StatusWord As Long '// +0x4(0x4)
     TagWord As Long '// +0x8(0x4)
     ErrorOffset As Long '// +0xc(0x4)
     ErrorSelector As Long '// +0x10(0x4)
     DataOffset As Long '// +0x14(0x4)
     DataSelector As Long '// +0x18(0x4)
     RegisterArea(1 To &H50) As Byte '// +0x1c(0x50)
     Cr0NpxState As Long '// +0x6c(0x4)
End Type
Public Type CONTEXT '// 0x2cc
     ContextFlags As Long  '// +0x0(0x4)
     Dr0 As Long  '// +0x4(0x4)
     Dr1 As Long '// +0x8(0x4)
     Dr2 As Long '// +0xc(0x4)
     Dr3 As Long '// +0x10(0x4)
     Dr6 As Long '// +0x14(0x4)
     Dr7 As Long '// +0x18(0x4)
     FloatSave As FLOATING_SAVE_AREA '// +0x1c(0x70)
     SegGs As Long '// +0x8c(0x4)
     SegFs As Long '// +0x90(0x4)
     SegEs As Long '// +0x94(0x4)
     SegDs As Long '// +0x98(0x4)
     Edi As Long '// +0x9c(0x4)
     Esi As Long '// +0xa0(0x4)
     Ebx As Long '// +0xa4(0x4)
     Edx As Long '// +0xa8(0x4)
     Ecx As Long '// +0xac(0x4)
     Eax As Long '// +0xb0(0x4)
     Ebp As Long '// +0xb4(0x4)
     Eip As Long '// +0xb8(0x4)
     SegCs As Long '// +0xbc(0x4)
     EFlags As Long '// +0xc0(0x4)
     Esp As Long '// +0xc4(0x4)
     SegSs As Long '// +0xc8(0x4)
     ExtendedRegisters(1 To &H200) As Byte '// +0xcc(0x200)
End Type
Public Const PROCESS_QUERY_INFORMATION As Long = (&H400)
Public Const STANDARD_RIGHTS_REQUIRED As Long = &HF0000
Public Const SYNCHRONIZE As Long = &H100000
Public Const PROCESS_ALL_ACCESS As Long = (STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &HFFF)
Public Const THREAD_ALL_ACCESS As Long = (STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &H3FF)
Public Const PROCESS_DUP_HANDLE As Long = (&H40)
Public Const CONTEXT_i386 As Long = &H10000     '// this assumes that i386 and
Public Const CONTEXT_i486 As Long = &H10000     '// i486 have identical context records
Public Const CONTEXT_CONTROL As Long = (CONTEXT_i386 Or &H1)        '// SS:SP, CS:IP, FLAGS, BP
Public Const CONTEXT_INTEGER As Long = (CONTEXT_i386 Or &H2)        '// AX, BX, CX, DX, SI, DI
Public Const CONTEXT_SEGMENTS As Long = (CONTEXT_i386 Or &H4)         '// DS, ES, FS, GS
Public Const CONTEXT_FLOATING_POINT As Long = (CONTEXT_i386 Or &H8)        '// 387 state
Public Const CONTEXT_DEBUG_REGISTERS As Long = (CONTEXT_i386 Or &H10)       '// DB 0-3,6,7
Public Const CONTEXT_EXTENDED_REGISTERS As Long = (CONTEXT_i386 Or &H20)       '// cpu specific extensions
Public Const CONTEXT_FULL As Long = (CONTEXT_CONTROL Or CONTEXT_INTEGER Or CONTEXT_SEGMENTS)
Public Const CONTEXT_ALL As Long = (CONTEXT_CONTROL Or CONTEXT_INTEGER Or CONTEXT_SEGMENTS Or CONTEXT_FLOATING_POINT Or CONTEXT_DEBUG_REGISTERS Or CONTEXT_EXTENDED_REGISTERS)

Public Declare Function ZwClose _
               Lib "NTDLL.DLL" (ByVal ObjectHandle As Long) As Long
Public Const ZwGetCurrentProcess As Long = -1 '//0xFFFFFFFF
Public Const ZwGetCurrentThread As Long = -2 '//0xFFFFFFFE
Public Const ZwCurrentProcess As Long = ZwGetCurrentProcess
Public Const ZwCurrentThread As Long = ZwGetCurrentThread
'Public Declare Sub CopyMemory _
 '              Lib "kernel32.dll" _
  '             Alias "RtlMoveMemory" (ByVal Destination As Long, _
   '                                   ByVal Source As Long, _
    '                                  ByVal Length As Long)
Public Declare Function ZwSuspendProcess _
               Lib "NTDLL.DLL" (ByVal ProcessHandle As Long) As Long
Public Declare Function ZwResumeProcess _
               Lib "NTDLL.DLL" (ByVal ProcessHandle As Long) As Long
Public Declare Function ZwSuspendThread _
               Lib "NTDLL.DLL" (ByVal ThreadHandle As Long, _
                                ByRef PreviousSuspendCount As Long) As Long
Public Declare Function ZwResumeThread _
               Lib "NTDLL.DLL" (ByVal ThreadHandle As Long, _
                                ByRef SuspendCount As Long) As Long
Public Declare Function ZwOpenThread _
               Lib "NTDLL.DLL" (ByRef ThreadHandle As Long, _
                                ByVal AccessMask As Long, _
                                ByRef ObjectAttributes As OBJECT_ATTRIBUTES, _
                                ByRef ClientId As CLIENT_ID) As Long
Public Declare Function ZwGetContextThread _
               Lib "NTDLL.DLL" (ByVal ThreadHandle As Long, _
                                ByRef ThreadContext As CONTEXT) As Long
Public Declare Function ZwSetContextThread _
               Lib "NTDLL.DLL" (ByVal ThreadHandle As Long, _
                                ByRef ThreadContext As CONTEXT) As Long
Public Declare Function ZwTerminateProcess _
               Lib "NTDLL.DLL" (ByVal ProcessHandle As Long, _
                                ByVal ExitStatus As Long) As Long
Public Declare Function DbgUiDebugActiveProcess _
               Lib "NTDLL.DLL" (ByVal ProcessHandle As Long) As Long
Public Declare Function DbgUiStopDebugging _
               Lib "NTDLL.DLL" (ByVal ProcessHandle As Long) As Long
Public Declare Function DbgUiConnectToDbg _
               Lib "NTDLL.DLL" () As Long
Public Const DEBUG_READ_EVENT = &H1
Public Const DEBUG_PROCESS_ASSIGN = &H2
Public Const DEBUG_SET_INFORMATION = &H4
Public Const DEBUG_QUERY_INFORMATION = &H8
Public Const DEBUG_ALL_ACCESS = STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or DEBUG_READ_EVENT Or DEBUG_PROCESS_ASSIGN Or DEBUG_SET_INFORMATION Or DEBUG_QUERY_INFORMATION
Public Const DEBUG_KILL_ON_CLOSE = &H1  '// Kill all debuggees on last handle close

Public Type LIST_ENTRY
        FLink As Long
        BLink As Long
End Type

Public Type LDR_MODULE
        InLoadOrderModuleList As LIST_ENTRY
        InMemoryOrderModuleList As LIST_ENTRY
        InInitializationOrderModuleList As LIST_ENTRY
        BaseAddress As Long
        EntryPoint As Long
        SizeOfImage As Long
        FullDllName As UNICODE_STRING
        BaseDllName As UNICODE_STRING
        Flags As Long
        LoadCount As Integer
        TlsIndex As Integer
        HashTableEntry As LIST_ENTRY
        TimeDateStamp As Long
End Type

Public Type PEB_LDR_DATA
        Length As Long ' : Uint4B
        Initialized As Byte
        SsHandle As Long ' : Ptr32 Void
        InLoadOrderModuleList As LIST_ENTRY  '//链表中的每一个指针都指向一个LDR_MODULE结构
        InMemoryOrderModuleList As LIST_ENTRY
        InInitializationOrderModuleList As LIST_ENTRY
        EntryInProgress As Long ' : Ptr32 Void
End Type

Public Type PEB
        InheritedAddressSpace As Byte
        ReadImageFileExecOptions As Byte
        BeingDebugged As Byte
        SpareBool As Byte
        Mutant As Long
        ImageBaseAddress As Long
        pLdr As Long ' _PEB_LDR_DATA
        ProcessParameters As Long ' _RTL_USER_PROCESS_PARAMETERS
        SubSystemData As Long
        ProcessHeap As Long
        FastPebLock As Long ' _RTL_CRITICAL_SECTION
        FastPebLockRoutine As Long
        FastPebUnlockRoutine As Long
        EnvironmentUpdateCount As Long
        KernelCallbackTable As Long
        SystemReserved(1 To 1) As Long
        AtlThunkSListPtr32 As Long
        FreeList As Long ' _PEB_FREE_BLOCK
        TlsExpansionCounter As Long
        TlsBitmap As Long
        TlsBitmapBits(1 To 2) As Long
        ReadOnlySharedMemoryBase As Long
        ReadOnlySharedMemoryHeap As Long
        ReadOnlyStaticServerData As Long
        AnsiCodePageData As Long
        OemCodePageData As Long
        UnicodeCaseTableData As Long
        NumberOfProcessors As Long
        NtGlobalFlag As Long
        CriticalSectionTimeout As LARGE_INTEGER
        HeapSegmentReserve As Long
        HeapSegmentCommit As Long
        HeapDeCommitTotalFreeThreshold As Long
        HeapDeCommitFreeBlockThreshold As Long
        NumberOfHeaps As Long
        MaximumNumberOfHeaps As Long
        ProcessHeaps As Long
        GdiSharedHandleTable As Long
        ProcessStarterHelper As Long
        GdiDCAttributeList As Long
        LoaderLock As Long
        OSMajorVersion As Long
        OSMinorVersion As Long
        OSBuildNumber As Integer
        OSCSDVersion As Integer
        OSPlatformId As Long
        ImageSubsystem As Long
        ImageSubsystemMajorVersion As Long
        ImageSubsystemMinorVersion As Long
        ImageProcessAffinityMask As Long
        GdiHandleBuffer(1 To 34) As Long
        PostProcessInitRoutine As Long
        TlsExpansionBitmap As Long
        TlsExpansionBitmapBits(1 To 32) As Long
        SessionId As Long
        AppCompatFlags As LARGE_INTEGER
        AppCompatFlagsUser As LARGE_INTEGER
        pShimData As Long
        AppCompatInfo As Long
        CSDVersion As UNICODE_STRING
        ActivationContextData As Long
        ProcessAssemblyStorageMap As Long
        SystemDefaultActivationContextData As Long
        SystemAssemblyStorageMap As Long
        MinimumStackCommit As Long
End Type
Public Declare Function ZwReadVirtualMemory _
               Lib "NTDLL.DLL" (ByVal ProcessHandle As Long, _
                                ByVal BaseAddress As Long, _
                                ByVal pBuffer As Long, _
                                ByVal NumberOfBytesToRead As Long, _
                                ByRef NumberOfBytesReaded As Long) As Long
Public Declare Function ZwWriteVirtualMemory _
               Lib "NTDLL.DLL" (ByVal ProcessHandle As Long, _
                                ByVal BaseAddress As Long, _
                                ByVal pBuffer As Long, _
                                ByVal NumberOfBytesToWrite As Long, _
                                ByRef NumberOfBytesWritten As Long) As Long
Public Type ACL
        AclRevision As Byte
        Sbz1 As Byte
        AclSize As Integer
        AceCount As Integer
        Sbz2 As Integer
End Type
Public Type SECURITY_DESCRIPTOR
        Revision As Byte
        Sbz1 As Byte
        Control As Long
        Owner As Long
        Group As Long
        Sacl As ACL
        Dacl As ACL
End Type
Public Declare Function RtlCreateUserThread _
        Lib "NTDLL.DLL" (ByVal ProcessHandle As Long, _
                         ByVal pSecurityDescriptor As Long, _
                         ByVal CreateSuspended As Long, _
                         ByVal StackZeroBits As Long, _
                         ByRef StackReserved As Long, _
                         ByRef StackCommit As Long, _
                         ByVal StartAddress As Long, _
                         ByVal StartParameter As Long, _
                         ByRef ThreadHandle As Long, _
                         ByRef ClientId As CLIENT_ID) As Long

Public Declare Function LdrGetDllHandle _
                        Lib "NTDLL.DLL" (ByVal pwPath As Long, _
                                         ByVal Unused As Long, _
                                         ByRef ModuleFileName As UNICODE_STRING, _
                                         ByRef pHModule As Long) As Long
Public Declare Function LdrGetProcedureAddress _
                        Lib "NTDLL.DLL" (ByVal ModuleHandle As Long, _
                                         ByRef FunctionName As ANSI_STRING, _
                                         ByVal Oridinal As Long, _
                                         ByRef FunctionAddress As Long) As Long
Public Declare Function RtlInitUnicodeString _
                        Lib "NTDLL.DLL" (ByRef DestinationString As UNICODE_STRING, _
                                         ByVal SourceString As Long) As Long
Public Declare Function RtlInitAnsiString _
                        Lib "NTDLL.DLL" (ByRef DestinationString As ANSI_STRING, _
                                         ByVal SourceString As Long) As Long
Public Declare Function RtlFreeUnicodeString _
                        Lib "NTDLL.DLL" (ByRef UnicodeString As UNICODE_STRING) As Long
Public Declare Function RtlFreeAnsiString _
                        Lib "NTDLL.DLL" (ByRef AnsiString As ANSI_STRING) As Long
Public Declare Function ZwWaitForSingleObject _
                        Lib "NTDLL.DLL" (ByVal hObject As Long, _
                                         ByVal Alertable As Long, _
                                         ByRef Timeout As LARGE_INTEGER) As Long
Public Declare Function ZwTerminateThread _
                        Lib "NTDLL.DLL" (ByVal ThreadHandle As Long, _
                                         ByVal ExitStatus As Long) As Long
Public Const IMAGE_DOS_SIGNATURE As Integer = &H5A4D ' MZ
Public Const IMAGE_OS2_SIGNATURE As Integer = &H454E ' NE
Public Const IMAGE_OS2_SIGNATURE_LE As Integer = &H454C 'LE
Public Const IMAGE_NT_SIGNATURE As Long = &H4550   ' PE00

Public Type DOS_MZ_HEADER
    e_magic As Integer
    e_cblp As Integer
    e_cp As Integer
    e_crlc As Integer
    e_cparhdr As Integer
    e_minalloc As Integer
    e_maxalloc As Integer
    e_ss As Integer
    e_sp As Integer
    e_csum As Integer
    e_ip As Integer
    e_cs As Integer
    e_lfarlc As Integer
    e_ovno As Integer
    e_res(0 To 3) As Integer
    e_oemid As Integer
    e_oeminfo As Integer
    e_res2(0 To 9)  As Integer
    e_lfanew As Long
End Type

Public Type IMAGE_FILE_HEADER
        Machine As Integer
        NumberOfSections As Integer
        TimeDateStamp As Long
        PointerToSymbolTable As Long
        NumberOfSymbols As Long
        SizeOfOptionalHeader As Integer
        Characteristics As Integer
End Type

Public Const IMAGE_NUMBEROF_DIRECTORY_ENTRIES = 16
'FLAGS
Public Const IMAGE_FILE_RELOCS_STRIPPED As Integer = &H1
Public Const IMAGE_FILE_EXECUTABLE_IMAGE As Integer = &H2
Public Const IMAGE_FILE_LINE_NUMS_STRIPPED As Integer = &H4
Public Const IMAGE_FILE_LOCAL_SYMS_STRIPPED As Integer = &H8
Public Const IMAGE_FILE_AGGRESIVE_WS_TRIM As Integer = &H10
Public Const IMAGE_FILE_BYTES_REVERSED_LO As Integer = &H80
Public Const IMAGE_FILE_32BIT_MACHINE As Integer = &H100
Public Const IMAGE_FILE_DEBUG_STRIPPED As Integer = &H200
Public Const IMAGE_FILE_REMOVABLE_RUN_FROM_SWAP As Integer = &H400
Public Const IMAGE_FILE_NET_RUN_FROM_SWAP As Integer = &H800
Public Const IMAGE_FILE_SYSTEM As Integer = &H1000
Public Const IMAGE_FILE_DLL As Integer = &H2000
Public Const IMAGE_FILE_UP_SYSTEM_ONLY As Integer = &H4000

'SUB_SYSTEM
Public Const SUB_SYS_UNKNOW As Integer = &H0
Public Const SUB_SYS_NATIVE As Integer = &H1
Public Const SUB_SYS_WINDOWS_GUI As Integer = &H2
Public Const SUB_SYS_WINDOWS_CHARACTER As Integer = &H3
Public Const SUB_SYS_OS2_CHARACTER As Integer = &H5
Public Const SUB_SYS_POSIX_CHARACTER As Integer = &H7

'DLL Flags
Public Const DLL_FLAGS_PRE_PROCESS_INIT As Integer = &H1
Public Const DLL_FLAGS_PRE_PROCESS_TER As Integer = &H2
Public Const DLL_FLAGS_PRE_THREAD_INIT As Integer = &H4
Public Const DLL_FLAGS_PRE_THREAD_TER As Integer = &H8

Public Const IMAGE_DIRECTORY_ENTRY_EXPORT = &H1

Public Type IMAGE_DATA_DIRECTORY
        VirtualAddress As Long
        Size As Long
End Type
 
Public Type IMAGE_OPTIONAL_HEADER
        Magic As Integer
        MajorLinkerVersion As Byte
        MinorLinkerVersion As Byte
        SizeOfCode As Long
        SizeOfInitializedData As Long
        SizeOfUninitializedData As Long
        AddressOfEntryPoint As Long
        BaseOfCode As Long
        BaseOfData As Long
        ' NT additional fields.24
        
        ImageBase As Long '28
        SectionAlignment As Long '32
        FileAlignment As Long '36
        MajorOperatingSystemVersion As Integer
        MinorOperatingSystemVersion As Integer '40
        MajorImageVersion As Integer
        MinorImageVersion As Integer '44
        MajorSubsystemVersion As Integer
        MinorSubsystemVersion As Integer '48
        Reserved1 As Long '56
        SizeOfImage As Long '60
        SizeOfHeaders As Long '64
        Checksum As Long '68
        Subsystem As Integer '70
        DllCharacteristics As Integer '72
        SizeOfStackReserve As Long '76
        SizeOfStackCommit As Long '80
        SizeOfHeapReserve As Long '84
        SizeOfHeapCommit As Long '88
        LoaderFlags As Long '92
        NumberOfRvaAndSizes As Long '96
        DataDirectory(1 To IMAGE_NUMBEROF_DIRECTORY_ENTRIES) As IMAGE_DATA_DIRECTORY
End Type

Public Type IMAGE_NT_HEADER
        Signature As Long
        FileHeader As IMAGE_FILE_HEADER
        OptionalHeader As IMAGE_OPTIONAL_HEADER
End Type

Public Type IMAGE_EXPORT_DIRECTORY
    Characteristics As Long
    TimeDateStamp As Long
    MajorVersion As Integer
    MinorVersion As Integer
    Name As Long
    Base As Long
    NumberOfNames As Long
    NumberOfFunctions As Long
    AddressOfFunctions As Long
    AddressOfNames As Long
    AddressOfNameOridinals As Long
End Type

Public Declare Function ZwFreeVirtualMemory _
               Lib "NTDLL.DLL" (ByVal ProcessHandle As Long, _
                                ByVal BaseAddress As Long, _
                                ByVal RegionSize As Long, _
                                ByVal FreeType As Long) As Long
Public Const MEM_DECOMMIT = &H4000             '// winnt ntddk wdm
Public Const MEM_RELEASE = &H8000              '// winnt ntddk wdm
Public Declare Function ZwProtectVirtualMemory _
               Lib "NTDLL.DLL" (ByVal ProcessHandle As Long, _
                                ByVal BaseAddress As Long, _
                                ByVal RegionSize As Long, _
                                ByVal NewProtect As Long, _
                                ByVal OldProtect As Long) As Long
Public Const PAGE_READWRITE As Long = &H4
Public Const PAGE_GUARD As Long = &H100
Public Const PAGE_EXECUTE As Long = &H10
Public Type ModuleInformation
        szFullDllName As String
        dwBaseAddress As Long
        dwImageSize As Long
End Type

Public Type ExportTableInformation
        szFunctionName As String
        dwFunctionAddr As Long
        dwOridinal As Long
End Type
Public Enum MEMORY_INFORMATION_CLASS
    MemoryBasicInformation
#If DEVL Then
    MemoryWorkingSetInformation
#End If
    MemoryMappedFilenameInformation
    MemoryRegionInformation
    MemoryWorkingSetExInformation
End Enum
Public Declare Function ZwUnmapViewOfSection _
                        Lib "NTDLL.DLL" (ByVal ProcessHandle As Long, _
                                         ByVal BaseAddress As Long) As Long
Public Declare Function ZwQueryVirtualMemory _
                        Lib "NTDLL.DLL" (ByVal ProcessHandle As Long, _
                                         ByVal BaseAddress As Long, _
                                         ByVal MemoryInformationClass As MEMORY_INFORMATION_CLASS, _
                                         ByVal MemoryInformation As Long, _
                                         ByVal MemoryInformationLength As Long, _
                                         ByRef ReturnLength As Long) As Long
Public Type JOB_SET_ARRAY
     JobHandle As Long '// Handle to job object to insert
     MemberLevel As Long '// Level of this job in the set. Must be > 0. Can be sparse.
     Flags As Long '// Unused. Must be zero
End Type
Public Declare Function ZwCreateJobObject _
               Lib "NTDLL.DLL" (ByRef JobHandle As Long, _
                                ByVal DesiredAccess As Long, _
                                ByRef ObjectAttributes As OBJECT_ATTRIBUTES) As Long
Public Declare Function ZwOpenJobObject _
               Lib "NTDLL.DLL" (ByRef JobHandle As Long, _
                                ByVal DesiredAccess As Long, _
                                ByRef ObjectAttributes As OBJECT_ATTRIBUTES) As Long
Public Declare Function ZwAssignProcessToJobObject _
               Lib "NTDLL.DLL" (ByVal JobHandle As Long, _
                                ByVal ProcessHandle As Long) As Long
Public Declare Function ZwTerminateJobObject _
               Lib "NTDLL.DLL" (ByVal JobHandle As Long, _
                                ByVal ExitStatus As Long) As Long
Public Declare Function ZwIsProcessInJob _
               Lib "NTDLL.DLL" (ByVal ProcessHandle As Long, _
                                ByVal JobHandle As Long) As Long
Public Declare Function ZwCreateJobSet _
               Lib "NTDLL.DLL" (ByVal NumJob As Long, _
                                ByRef UserJobSet As JOB_SET_ARRAY, _
                                ByVal Flags As Long) As Long
Public Const JOB_OBJECT_ASSIGN_PROCESS As Long = &H1
Public Const JOB_OBJECT_SET_ATTRIBUTES As Long = &H2
Public Const JOB_OBJECT_QUERY As Long = &H4
Public Const JOB_OBJECT_TERMINATE As Long = &H8
Public Const JOB_OBJECT_SET_SECURITY_ATTRIBUTES As Long = &H10
Public Const JOB_OBJECT_ALL_ACCESS As Long = STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &H1F
Public Const OBJ_INHERIT = &H2
Public Const STATUS_SUCCESS As Long = &H0

Public Function NT_SUCCESS(ByVal Status) As Boolean
        NT_SUCCESS = (Status >= 0)
End Function


