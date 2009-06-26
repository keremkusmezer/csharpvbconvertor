Imports System
Imports System.Collections.Generic
Imports System.EnterpriseServices
Imports System.IO
Imports System.Runtime.CompilerServices
Imports System.Runtime.ConstrainedExecution
Imports System.Runtime.InteropServices
Imports System.Security.Permissions
Imports System.Text
Imports System.Threading
Imports Microsoft.Win32.SafeHandles
Namespace VmcController.Services
    Friend Enum NT_STATUS
        STATUS_SUCCESS = &H0
        STATUS_BUFFER_OVERFLOW = DirectCast(&H80000005, Integer)
        STATUS_INFO_LENGTH_MISMATCH = DirectCast(&HC0000004, Integer)
    End Enum
    Friend Enum SYSTEM_INFORMATION_CLASS
        SystemBasicInformation = 0
        SystemPerformanceInformation = 2
        SystemTimeOfDayInformation = 3
        SystemProcessInformation = 5
        SystemProcessorPerformanceInformation = 8
        SystemHandleInformation = 16
        SystemInterruptInformation = 23
        SystemExceptionInformation = 33
        SystemRegistryQuotaInformation = 37
        SystemLookasideInformation = 45
    End Enum
    Friend Enum OBJECT_INFORMATION_CLASS
        ObjectBasicInformation = 0
        ObjectNameInformation = 1
        ObjectTypeInformation = 2
        ObjectAllTypesInformation = 3
        ObjectHandleInformation = 4
    End Enum
    <Flags()> _
    Friend Enum ProcessAccessRights
        PROCESS_DUP_HANDLE = &H40
        PROCESS_ALL_ACCESS = &HF0000 Or &H100000 Or &HFFF
    End Enum
    <Flags()> _
    Friend Enum DuplicateHandleOptions
        DUPLICATE_CLOSE_SOURCE = &H1
        DUPLICATE_SAME_ACCESS = &H2
    End Enum
    <SecurityPermission(SecurityAction.LinkDemand, UnmanagedCode:=True)> _
    Friend NotInheritable Class SafeObjectHandle
        Inherits SafeHandleZeroOrMinusOneIsInvalid
        Private Sub New()
            MyBase.New(True)
        End Sub
        Friend Sub New(ByVal preexistingHandle As IntPtr, ByVal ownsHandle As Boolean)
            MyBase.New(ownsHandle)
            MyBase.SetHandle(preexistingHandle)
        End Sub
        Protected Overloads Overrides Function ReleaseHandle() As Boolean
            Return NativeMethods.CloseHandle(MyBase.handle)
        End Function
    End Class
    <SecurityPermission(SecurityAction.LinkDemand, UnmanagedCode:=True)> _
    Friend NotInheritable Class SafeProcessHandle
        Inherits SafeHandleZeroOrMinusOneIsInvalid
        Private Sub New()
            MyBase.New(True)
        End Sub
        Friend Sub New(ByVal preexistingHandle As IntPtr, ByVal ownsHandle As Boolean)
            MyBase.New(ownsHandle)
            MyBase.SetHandle(preexistingHandle)
        End Sub
        Protected Overloads Overrides Function ReleaseHandle() As Boolean
            Return NativeMethods.CloseHandle(MyBase.handle)
        End Function
    End Class
    Friend Class NativeMethods
        <DllImport("ntdll.dll")> _
        Friend Shared Function NtQuerySystemInformation(<[In]()> ByVal SystemInformationClass As SYSTEM_INFORMATION_CLASS, <[In]()> ByVal SystemInformation As IntPtr, <[In]()> ByVal SystemInformationLength As Integer, <Out()> ByRef ReturnLength As Integer) As NT_STATUS
        End Function
        <DllImport("ntdll.dll")> _
        Friend Shared Function NtQueryObject(<[In]()> ByVal Handle As IntPtr, <[In]()> ByVal ObjectInformationClass As OBJECT_INFORMATION_CLASS, <[In]()> ByVal ObjectInformation As IntPtr, <[In]()> ByVal ObjectInformationLength As Integer, <Out()> ByRef ReturnLength As Integer) As NT_STATUS
        End Function
        <DllImport("kernel32.dll", SetLastError:=True)> _
        Friend Shared Function OpenProcess(<[In]()> ByVal dwDesiredAccess As ProcessAccessRights, <[In](), MarshalAs(UnmanagedType.Bool)> ByVal bInheritHandle As Boolean, <[In]()> ByVal dwProcessId As Integer) As SafeProcessHandle
        End Function
        <DllImport("kernel32.dll", SetLastError:=True)> _
        Friend Shared Function DuplicateHandle(<[In]()> ByVal hSourceProcessHandle As IntPtr, <[In]()> ByVal hSourceHandle As IntPtr, <[In]()> ByVal hTargetProcessHandle As IntPtr, <Out()> ByRef lpTargetHandle As SafeObjectHandle, <[In]()> ByVal dwDesiredAccess As Integer, <[In](), MarshalAs(UnmanagedType.Bool)> ByVal bInheritHandle As Boolean, _
         <[In]()> ByVal dwOptions As DuplicateHandleOptions) As <MarshalAs(UnmanagedType.Bool)> Boolean
        End Function
        <DllImport("kernel32.dll")> _
        Friend Shared Function GetCurrentProcess() As IntPtr
        End Function
        <DllImport("kernel32.dll", SetLastError:=True)> _
        Friend Shared Function GetProcessId(<[In]()> ByVal Process As IntPtr) As Integer
        End Function
        <ReliabilityContract(Consistency.WillNotCorruptState, Cer.Success)> _
        <DllImport("kernel32.dll", SetLastError:=True)> _
        Friend Shared Function CloseHandle(<[In]()> ByVal hObject As IntPtr) As <MarshalAs(UnmanagedType.Bool)> Boolean
        End Function
        <DllImport("kernel32.dll", SetLastError:=True)> _
        Friend Shared Function QueryDosDevice(<[In]()> ByVal lpDeviceName As String, <Out()> ByVal lpTargetPath As StringBuilder, <[In]()> ByVal ucchMax As Integer) As Integer
        End Function
        <DllImport("ntdll.dll")> _
        Friend Shared Sub NtClose(ByVal nProcess As IntPtr)
        End Sub
        <DllImport("ntdll.dll")> _
        Friend Shared Function NtDuplicateObject(ByVal hSourceProcess As IntPtr, ByVal hSourceHandle As IntPtr, ByVal hCopyProcess As IntPtr, ByVal CopyHandle As IntPtr, ByVal DesiredAccess As Long, ByVal Attributes As Long, ByVal Options As Long) As Long

        End Function
        'DuplicateHandle 
        '        BOOL WINAPI DuplicateHandle(  __in          HANDLE hSourceProcessHandle,
        ' __in          HANDLE hSourceHandle,  __in          HANDLE hTargetProcessHandle,  
        '__out         LPHANDLE lpTargetHandle,  __in          DWORD dwDesiredAccess,  
        '__in          BOOL bInheritHandle,  __in          DWORD dwOptions);
        'Private Sub CloseHandleEx(ByVal lPID As Long, ByVal lHandle As Long)    
        'Dim hProcess    hProcess = OpenProcess(PROCESS_ALL_ACCESS, 0, lPID)    
        'Call NtDuplicateObject(hProcess, lHandle, 0, ByVal 0, 0, 0, DUPLICATE_CLOSE_SOURCE)    
        'Call NtClose(hProcess)End Sub
        'DUPLICATE_CLOSE_SOURCE --> 1
    End Class
    <ComVisible(True), EventTrackingEnabled(True)> _
    Public Class DetectOpenFiles
        Inherits ServicedComponent
        Private Shared deviceMap As Dictionary(Of String, String)
        Private Const networkDevicePrefix As String = "\Device\LanmanRedirector\"
        Private Const MAX_PATH As Integer = 260
        Private Enum SystemHandleType
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
        Private Const handleTypeTokenCount As Integer = 27
        Private Shared ReadOnly handleTypeTokens As String() = New String() {"", "", "Directory", "SymbolicLink", "Token", "Process", _
         "Thread", "Unknown7", "Event", "EventPair", "Mutant", "Unknown11", _
         "Semaphore", "Timer", "Profile", "WindowStation", "Desktop", "Section", _
         "Key", "Port", "WaitablePort", "Unknown21", "Unknown22", "Unknown23", _
         "Unknown24", "IoCompletion", "File"}
        <StructLayout(LayoutKind.Sequential)> _
        Private Structure SYSTEM_HANDLE_ENTRY
            Public OwnerPid As Integer
            Public ObjectType As Byte
            Public HandleFlags As Byte
            Public HandleValue As Short
            Public ObjectPointer As Integer
            Public AccessMask As Integer
        End Structure
        'http://forum.sysinternals.com/forum_posts.asp?TID=14546&KW=open+handles
        'Public Shared Function CloseHandleEx(ByVal processId As Integer, ByVal fileName As String)
        '    Dim currentHandle As SafeProcessHandle = NativeMethods.OpenProcess(ProcessAccessRights.PROCESS_ALL_ACCESS, False, processId)
        '    Dim objectHandle As SafeObjectHandle = Nothing
        '    NativeMethods.DuplicateHandle(New System.IntPtr(processId), currentHandle.DangerousGetHandle, IntPtr.Zero, objectHandle, 0, False, DuplicateHandleOptions.DUPLICATE_CLOSE_SOURCE)

        '    'currentHandle.
        'End Function
        Public Shared Function GetOpenFilesEnumerator(ByVal processId As Integer) As IEnumerator(Of FileSystemInfo)
            Return New OpenFiles(processId).GetEnumerator2()
        End Function
        Private NotInheritable Class OpenFiles
            Implements IEnumerable(Of FileSystemInfo)
            Private ReadOnly processId As Integer
            Private resultCollection As New System.Collections.ObjectModel.Collection(Of FileSystemInfo)
            Friend Sub New(ByVal processId As Integer)
                Me.processId = processId
            End Sub
            Public Function GetEnumerator2() As IEnumerator(Of FileSystemInfo) Implements IEnumerable(Of System.IO.FileSystemInfo).GetEnumerator
                resultCollection.Clear()
                Dim ret As NT_STATUS
                Dim length As Integer = &H10000
                Do
                    Dim ptr As IntPtr = IntPtr.Zero
                    RuntimeHelpers.PrepareConstrainedRegions()
                    Try
                        RuntimeHelpers.PrepareConstrainedRegions()
                        Try
                        Finally
                            ptr = Marshal.AllocHGlobal(length)
                        End Try
                        Dim returnLength As Integer
                        ret = NativeMethods.NtQuerySystemInformation(SYSTEM_INFORMATION_CLASS.SystemHandleInformation, ptr, length, returnLength)
                        If ret = NT_STATUS.STATUS_INFO_LENGTH_MISMATCH Then
                            length = ((returnLength + &HFFFF) And Not &HFFFF)
                        ElseIf ret = NT_STATUS.STATUS_SUCCESS Then
                            Dim handleCount As Integer = Marshal.ReadInt32(ptr)
                            Dim offset As Integer = 4
                            Dim size As Integer = Marshal.SizeOf(GetType(SYSTEM_HANDLE_ENTRY))
                            Dim i As Integer = 0
                            While i < handleCount
                                Dim handleEntry As SYSTEM_HANDLE_ENTRY = DirectCast(Marshal.PtrToStructure(New IntPtr(ptr.ToInt32() + offset), GetType(SYSTEM_HANDLE_ENTRY)), SYSTEM_HANDLE_ENTRY)
                                If handleEntry.OwnerPid = processId Then
                                    Dim handle As IntPtr = New IntPtr(handleEntry.HandleValue) 'DirectCast(, IntPtr)
                                    Dim handleType As SystemHandleType
                                    If GetHandleType(handle, handleEntry.OwnerPid, handleType) AndAlso handleType = SystemHandleType.OB_TYPE_FILE Then
                                        Dim devicePath As String = String.Empty
                                        If GetFileNameFromHandle(handle, handleEntry.OwnerPid, devicePath) Then
                                            Dim dosPath As String = String.Empty
                                            If ConvertDevicePathToDosPath(devicePath, dosPath) Then
                                                If File.Exists(dosPath) Then
                                                    resultCollection.Add(New FileInfo(dosPath))
                                                ElseIf Directory.Exists(dosPath) Then
                                                    resultCollection.Add(New DirectoryInfo(dosPath))
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                                offset += size
                                System.Math.Max(System.Threading.Interlocked.Increment(i), i - 1)
                            End While
                        End If
                    Finally
                        Marshal.FreeHGlobal(ptr)
                    End Try
                Loop While ret = NT_STATUS.STATUS_INFO_LENGTH_MISMATCH
                Return resultCollection.GetEnumerator()
            End Function
            Function GetEnumerator() As System.Collections.IEnumerator Implements System.Collections.IEnumerable.GetEnumerator
                Return GetEnumerator2()
            End Function
        End Class
        Private Shared Function GetFileNameFromHandle(ByVal handle As IntPtr, ByVal processId As Integer, ByRef fileName As String) As Boolean
            Dim currentProcess As IntPtr = NativeMethods.GetCurrentProcess()
            Dim remote As Boolean = (processId <> NativeMethods.GetProcessId(currentProcess))
            Dim processHandle As SafeProcessHandle = Nothing
            Dim objectHandle As SafeObjectHandle = Nothing
            Try
                If remote Then
                    processHandle = NativeMethods.OpenProcess(ProcessAccessRights.PROCESS_DUP_HANDLE, True, processId)
                    If NativeMethods.DuplicateHandle(processHandle.DangerousGetHandle(), handle, currentProcess, objectHandle, 0, False, _
                     DuplicateHandleOptions.DUPLICATE_SAME_ACCESS) Then
                        handle = objectHandle.DangerousGetHandle()
                    End If
                End If
                Return GetFileNameFromHandle(handle, fileName, 200)
            Finally
                If remote Then
                    If processHandle IsNot Nothing Then
                        processHandle.Close()
                    End If
                    If objectHandle IsNot Nothing Then
                        objectHandle.Close()
                    End If
                End If
            End Try
        End Function
        Private Shared Function GetFileNameFromHandle(ByVal handle As IntPtr, ByRef fileName As String, ByVal wait As Integer) As Boolean
            Using f As New FileNameFromHandleState(handle)
                ThreadPool.QueueUserWorkItem(New WaitCallback(AddressOf GetFileNameFromHandle), f)
                If f.WaitOne(wait) Then
                    fileName = f.FileName
                    Return f.RetValue
                Else
                    fileName = String.Empty
                    Return False
                End If
            End Using
        End Function
        Private Class FileNameFromHandleState
            Implements IDisposable
            Private _mr As ManualResetEvent
            Private _handle As IntPtr
            Private _fileName As String
            Private _retValue As Boolean
            Public ReadOnly Property Handle() As IntPtr
                Get
                    Return _handle
                End Get
            End Property
            Public Property FileName() As String
                Get
                    Return _fileName
                End Get
                Set(ByVal value As String)
                    _fileName = value
                End Set
            End Property
            Public Property RetValue() As Boolean
                Get
                    Return _retValue
                End Get
                Set(ByVal value As Boolean)
                    _retValue = value
                End Set
            End Property
            Public Sub New(ByVal handle As IntPtr)
                _mr = New ManualResetEvent(False)
                Me._handle = handle
            End Sub
            Public Function WaitOne(ByVal wait As Integer) As Boolean
                Return _mr.WaitOne(wait, False)
            End Function
            Public Sub [Set]()
                _mr.[Set]()
            End Sub
            Public Sub Dispose() Implements IDisposable.Dispose
                If _mr IsNot Nothing Then
                    _mr.Close()
                End If
            End Sub
        End Class
        Private Shared Sub GetFileNameFromHandle(ByVal state As Object)
            Dim s As FileNameFromHandleState = DirectCast(state, FileNameFromHandleState)
            Dim fileName As String = String.Empty
            s.RetValue = GetFileNameFromHandle(s.Handle, fileName)
            s.FileName = fileName
            s.[Set]()
        End Sub
        Private Shared Function GetFileNameFromHandle(ByVal handle As IntPtr, ByRef fileName As String) As Boolean
            Dim ptr As IntPtr = IntPtr.Zero
            RuntimeHelpers.PrepareConstrainedRegions()
            Try
                Dim length As Integer = &H200
                RuntimeHelpers.PrepareConstrainedRegions()
                Try
                Finally
                    ptr = Marshal.AllocHGlobal(length)
                End Try
                Dim ret As NT_STATUS = NativeMethods.NtQueryObject(handle, OBJECT_INFORMATION_CLASS.ObjectNameInformation, ptr, length, length)
                If ret = NT_STATUS.STATUS_BUFFER_OVERFLOW Then
                    RuntimeHelpers.PrepareConstrainedRegions()
                    Try
                    Finally
                        Marshal.FreeHGlobal(ptr)
                        ptr = Marshal.AllocHGlobal(length)
                    End Try
                    ret = NativeMethods.NtQueryObject(handle, OBJECT_INFORMATION_CLASS.ObjectNameInformation, ptr, length, length)
                End If
                If ret = NT_STATUS.STATUS_SUCCESS Then
                    fileName = Marshal.PtrToStringUni(New IntPtr(ptr.ToInt32() + 8), ((length - 9) / 2))
                    Return fileName.Length <> 0
                End If
            Finally
                Marshal.FreeHGlobal(ptr)
            End Try
            fileName = String.Empty
            Return False
        End Function
        Private Shared Function GetHandleType(ByVal handle As IntPtr, ByVal processId As Integer, ByRef handleType As SystemHandleType) As Boolean
            Dim token As String = GetHandleTypeToken(handle, processId)
            Return GetHandleTypeFromToken(token, handleType)
        End Function
        Private Shared Function GetHandleType(ByVal handle As IntPtr, ByRef handleType As SystemHandleType) As Boolean
            Dim token As String = GetHandleTypeToken(handle)
            Return GetHandleTypeFromToken(token, handleType)
        End Function
        Private Shared Function GetHandleTypeFromToken(ByVal token As String, ByRef handleType As SystemHandleType) As Boolean
            Dim i As Integer = 1
            While i < handleTypeTokenCount
                If handleTypeTokens(i) = token Then
                    handleType = DirectCast(i, SystemHandleType)
                    Return True
                End If
                System.Math.Max(System.Threading.Interlocked.Increment(i), i - 1)
            End While
            handleType = SystemHandleType.OB_TYPE_UNKNOWN
            Return False
        End Function
        Private Shared Function GetHandleTypeToken(ByVal handle As IntPtr, ByVal processId As Integer) As String
            Dim currentProcess As IntPtr = NativeMethods.GetCurrentProcess()
            Dim remote As Boolean = (processId <> NativeMethods.GetProcessId(currentProcess))
            Dim processHandle As SafeProcessHandle = Nothing
            Dim objectHandle As SafeObjectHandle = Nothing
            Try
                If remote Then
                    processHandle = NativeMethods.OpenProcess(ProcessAccessRights.PROCESS_DUP_HANDLE, True, processId)
                    If NativeMethods.DuplicateHandle(processHandle.DangerousGetHandle(), handle, currentProcess, objectHandle, 0, False, _
                     DuplicateHandleOptions.DUPLICATE_SAME_ACCESS) Then
                        handle = objectHandle.DangerousGetHandle()
                    End If
                End If
                Return GetHandleTypeToken(handle)
            Finally
                If remote Then
                    If processHandle IsNot Nothing Then
                        processHandle.Close()
                    End If
                    If objectHandle IsNot Nothing Then
                        objectHandle.Close()
                    End If
                End If
            End Try
        End Function
        Private Shared Function GetHandleTypeToken(ByVal handle As IntPtr) As String
            Dim length As Integer
            NativeMethods.NtQueryObject(handle, OBJECT_INFORMATION_CLASS.ObjectTypeInformation, IntPtr.Zero, 0, length)
            Dim ptr As IntPtr = IntPtr.Zero
            RuntimeHelpers.PrepareConstrainedRegions()
            Try
                RuntimeHelpers.PrepareConstrainedRegions()
                Try
                Finally
                    ptr = Marshal.AllocHGlobal(length)
                End Try
                If NativeMethods.NtQueryObject(handle, OBJECT_INFORMATION_CLASS.ObjectTypeInformation, ptr, length, length) = NT_STATUS.STATUS_SUCCESS Then
                    Return Marshal.PtrToStringUni(New IntPtr(ptr.ToInt32() + &H60))
                End If
            Finally
                Marshal.FreeHGlobal(ptr)
            End Try
            Return String.Empty
        End Function
        Private Shared Function ConvertDevicePathToDosPath(ByVal devicePath As String, ByRef dosPath As String) As Boolean
            EnsureDeviceMap()
            Dim tempCollection As System.Collections.Generic.Dictionary(Of String, String).KeyCollection
            tempCollection = deviceMap.Keys
            For Each key As String In deviceMap.Keys
                If devicePath.StartsWith(key) Then
                    dosPath = devicePath.Replace(key, deviceMap(key))
                    Return dosPath.Length <> 0
                End If
            Next
            'Dim i As Integer = devicePath.Length
            'While i > 0 AndAlso (i = devicePath.LastIndexOf("\"c, i - 1)) <> -1
            '    Dim drive As String
            '    If deviceMap.TryGetValue(devicePath.Substring(0, i), drive) Then
            '        dosPath = String.Concat(drive, devicePath.Substring(i))
            '        Return dosPath.Length <> 0
            '    End If
            'End While
            dosPath = String.Empty
            Return False
        End Function
        Private Shared Sub EnsureDeviceMap()
            If deviceMap Is Nothing Then
                Dim localDeviceMap As Dictionary(Of String, String) = BuildDeviceMap()
                Interlocked.CompareExchange(Of Dictionary(Of String, String))(deviceMap, localDeviceMap, Nothing)
            End If
        End Sub
        Private Shared Function BuildDeviceMap() As Dictionary(Of String, String)
            Dim logicalDrives As String() = Environment.GetLogicalDrives()
            Dim localDeviceMap As New Dictionary(Of String, String)(logicalDrives.Length)
            Dim lpTargetPath As New StringBuilder(MAX_PATH)
            For Each drive As String In logicalDrives
                Dim lpDeviceName As String = drive.Substring(0, 2)
                NativeMethods.QueryDosDevice(lpDeviceName, lpTargetPath, MAX_PATH)
                localDeviceMap.Add(NormalizeDeviceName(lpTargetPath.ToString()), lpDeviceName)
            Next
            localDeviceMap.Add(networkDevicePrefix.Substring(0, networkDevicePrefix.Length - 1), "\")
            Return localDeviceMap
        End Function
        Private Shared Function NormalizeDeviceName(ByVal deviceName As String) As String
            If String.Compare(deviceName, 0, networkDevicePrefix, 0, networkDevicePrefix.Length, StringComparison.InvariantCulture) = 0 Then
                Dim shareName As String = deviceName.Substring(deviceName.IndexOf("\"c, networkDevicePrefix.Length) + 1)
                Return String.Concat(networkDevicePrefix, shareName)
            End If
            Return deviceName
        End Function

        Public Sub New()

        End Sub
    End Class
End Namespace