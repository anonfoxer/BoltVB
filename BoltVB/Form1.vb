Imports System.Runtime.InteropServices
Imports System.Text
Imports VB = Microsoft.VisualBasic

Public Class Form1
    Dim DLLDIR As String
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Integer
    Private Declare Function GetWindowThreadProcessId Lib "user32" Alias "GetWindowThreadProcessId" (ByVal hwnd As Integer, ByRef lpdwProcessId As Integer) As Integer
    Private Declare Function OpenProcess Lib "kernel32" Alias "OpenProcess" (ByVal dwDesiredAccess As Integer, ByVal bInheritHandle As Integer, ByVal dwProcessId As Integer) As Integer

    Public Function getPID(ByVal proc_name As String) As Integer
        Dim procs As Process()
        procs = Process.GetProcessesByName(proc_name)
        If procs.Length > 0 Then
            Return procs(0).Id
        Else
            Return 0
        End If
    End Function

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        MessageBox.Show("Bolt: An RVGL Modloader. Really, it's a simplified DLL injector. This went through so many damn iterations until I decided on VB. This was Way, way, wayyyyyyyy too much difficultly. So I feel proud. Go me. Version: 1.0.0. Updates at github.com/anonfoxer Injection module by CopyrightPhilly on UnknownCheats. Thank you forums from 2009.", "About", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Private Sub injectButt_Click(sender As Object, e As EventArgs) Handles injectButt.Click
        Dim game As String
        Dim myProcesses As Process()

        DLLDIR = modPathBox.Text
        game = "rvgl"
        myProcesses = Process.GetProcessesByName(game)

        If myProcesses.Length = 0 Then
            MessageBox.Show("RVGL was not Detected. Please launch RVGL, and load mods at the main menu.", "ERROR: PROCESS NOT FOUND", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Else
            Dim processID = getPID(game)
            inject(processID, DLLDIR)
        End If
    End Sub
End Class


Module ModInject
    ' MODULE MADE BY COPYRIGHTPHILLY ON UNKNOWNCHEATS https://www.unknowncheats.me/forum/members/134362.html
    Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Integer, ByVal bInheritHandle As Integer, ByVal dwProcessId As Integer) As Integer
    Private Declare Function VirtualAllocEx Lib "kernel32" (ByVal hProcess As Integer, ByVal lpAddress As Integer, ByVal dwSize As Integer, ByVal flAllocationType As Integer, ByVal flProtect As Integer) As Integer
    Private Declare Function WriteProcessMemory Lib "kernel32" (ByVal hProcess As Integer, ByVal lpBaseAddress As Integer, ByVal lpBuffer() As Byte, ByVal nSize As Integer, ByVal lpNumberOfBytesWritten As UInteger) As Boolean
    Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Integer, ByVal lpProcName As String) As Integer
    Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Integer
    Private Declare Function CreateRemoteThread Lib "kernel32" (ByVal hProcess As Integer, ByVal lpThreadAttributes As Integer, ByVal dwStackSize As Integer, ByVal lpStartAddress As Integer, ByVal lpParameter As Integer, ByVal dwCreationFlags As Integer, ByVal lpThreadId As Integer) As Integer
    Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Integer, ByVal dwMilliseconds As Integer) As Integer
    Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Integer) As Integer
    Public Function inject(ByVal ProcessID As Long, ByVal DLLPath As String) As Boolean
        On Error GoTo exiterror
        Dim DProc As Integer
        Dim DAdd As Integer
        Dim DWrote As UInteger
        Dim DAll As Integer
        Dim DThe As Integer
        Dim DMHD As Integer
        DProc = OpenProcess(&H1F0FFF, 1, ProcessID)
        DAdd = VirtualAllocEx(DProc, 0, DLLPath.Length, &H1000, &H4)
        If (DAdd > 0) Then
            Dim DByte() As Byte
            DByte = StrChar(DLLPath)
            WriteProcessMemory(DProc, DAdd, DByte, DLLPath.Length, DWrote)
            DMHD = GetModuleHandle("kernel32.dll")
            DAll = GetProcAddress(DMHD, "LoadLibraryA")
            DThe = CreateRemoteThread(DProc, 0, 0, DAll, DAdd, 0, 0)
            If (DThe > 0) Then
                WaitForSingleObject(DThe, &HFFFF)
                CloseHandle(DThe)
                Return True
            Else
                GoTo exiterror
            End If
        Else
            GoTo exiterror
        End If
        inject = True
        Exit Function
exiterror:
        inject = False
    End Function
    Private Function StrChar(ByRef strString As String) As Byte()
        Dim bytTemp() As Byte
        Dim i As Short
        ReDim bytTemp(0)
        For i = 1 To Len(strString)
            If bytTemp(UBound(bytTemp)) <> 0 Then ReDim Preserve bytTemp(UBound(bytTemp) + 1)
            bytTemp(UBound(bytTemp)) = Asc(Mid(strString, i, 1))
        Next i
        ReDim Preserve bytTemp(UBound(bytTemp) + 1)
        bytTemp(UBound(bytTemp)) = 0
        StrChar = bytTemp
    End Function
End Module