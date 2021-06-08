Attribute VB_Name = "ShellRoutines"
'@Folder "Common.Shared"
'@IgnoreModule ConstantNotUsed, AssignmentNotUsed, VariableNotUsed, UseMeaningfulName
Option Explicit

' The WaitForSingleObject function returns when one of the following occurs:
' - The specified object is in the signaled state.
' - The time-out interval elapses.
'
' The dwMilliseconds parameter specifies the time-out interval, in milliseconds.
' The function returns if the interval elapses, even if the object’s state is
' nonsignaled. If dwMilliseconds is zero, the function tests the object’s state
' and returns immediately. If dwMilliseconds is INFINITE, the function’s time-out
' interval never elapses.
'
' This example waits an INFINITE amount of time for the process to end. As a
' result this process will be frozen until the shelled process terminates. The
' down side is that if the shelled process hangs, so will this one.
'
' A better approach is to wait a specific amount of time. Once the time-out
' interval expires, test the return value. If it is WAIT_TIMEOUT, the process
' is still not signaled. Then you can either wait again or continue with your
' processing.
'
' DOS Applications:
' Waiting for a DOS application is tricky because the DOS window never goes
' away when the application is done. To get around this, prefix the app that
' you are shelling to with "command.com /c".
'
' For example: lPid = Shell("command.com /c " & txtApp.Text, vbNormalFocus)
'
' To get the return code from the DOS app, see the attached text file.
'

Const SYNCHRONIZE As Long = &H100000
'
' Wait forever
Const INFINITE As Long = &HFFFF
'
' The state of the specified object is signaled
Const WAIT_OBJECT_0 As Long = 0
'
' The time-out interval elapsed & the object’s state is not signaled
Const WAIT_TIMEOUT As Long = 1000


Public Function ReadLines(ByVal FilePath As String) As Variant
    Dim handle As Long: handle = FreeFile
    Open FilePath For Input As handle
    
    Dim Buffer As String: Buffer = Input$(LOF(handle), handle)
    If Right$(Buffer, Len(vbNewLine)) = vbNewLine Then
        Buffer = Left$(Buffer, Len(Buffer) - Len(vbNewLine))
    End If
    ReadLines = Split(Buffer, vbNewLine)
    Close handle
End Function


' Runs shell command, waits for completions, sends stdout to file and returns stdout as an array of strings
Public Function SyncRun(ByVal Command As String, Optional ByVal redirectStdout As Boolean = True) As Variant
    Dim cli As String
    If redirectStdout Then
        Dim GUID As String: GUID = Mid$(CreateObject("Scriptlet.TypeLib").GUID, 2, 36)
        Dim sys As IWshRuntimeLibrary.WshShell: Set sys = New IWshRuntimeLibrary.WshShell
        Dim TempFile As String: TempFile = sys.ExpandEnvironmentStrings("%temp%\stdout-") & GUID & ".txt"
        cli = Command & " >""" & TempFile & """"
    Else
        cli = Command
    End If

    Dim pid As Long: pid = Shell(cli, vbHide)

    If pid <> 0 Then
        'Get a handle to the shelled process.
        #If VBA7 Then
            Dim handle As LongPtr: handle = OpenProcess(SYNCHRONIZE, 0, pid)
        #Else
            Dim handle As Long: handle = OpenProcess(SYNCHRONIZE, 0, pid)
        #End If
        
        'If successful, wait for the application to end and close the handle.
        If handle <> 0 Then
            Dim Result As Long: Result = WaitForSingleObject(handle, WAIT_TIMEOUT)
            CloseHandle hObject:=handle
        End If
    End If
    If redirectStdout Then
        SyncRun = ReadLines(TempFile)
    End If
End Function


'@Ignore ProcedureNotUsed
Private Sub Test()
    Dim cmdline As String
    Dim output As Variant
    
    cmdline = "cmd /c dir /b c:\windows"
    output = SyncRun(cmdline)
    'ShellSync ("cmd /c dir /b c:\windows |clip")
End Sub
