Attribute VB_Name = "modLogging"
Option Explicit


Global Const scriptsLog = "logs/scripts.log"
Global Const printDebugLog = "logs/debug.log"
Global Const debugCrashLog = "logs/debugcrash.log"
Global Const errorHandlingLog = "logs/CRASHLOG.txt"

Sub CheckRollLog(LogName As String)
    Dim A As Long
    
    If (FileLen(LogName) > 10000000) Then
        If Dir(LogName & "5") <> "" Then Kill LogName & "5"
        For A = 4 To 1 Step -1
            If Dir(LogName & CStr(A)) <> "" Then Name LogName & CStr(A) As LogName & CStr(A + 1)
        Next A
        If Dir(LogName) <> "" Then Name LogName As LogName & "1"
    End If
End Sub

Sub LogServerStart()
    Dim Message As String
    Message = Now & " ----- Server Started --------------------------------------------------------------"
        
    If Dir("logs", vbDirectory) = "" Then
          MkDir "logs"
    End If
    
    Open errorHandlingLog For Append As #1
        Print #1, Message
    Close #1
    Open scriptsLog For Append As #1
        Print #1, Message
    Close #1
    Open printDebugLog For Append As #1
        Print #1, Message
    Close #1
    Open debugCrashLog For Append As #1
        Print #1, Message
    Close #1
End Sub

Sub LogScriptStart(Script As String)
    Dim Message As String
    Dim A As Long
        
    CheckRollLog scriptsLog
    
    Message = Now & " - Script " & Script & " started.  Parameters: "
    
    Message = Message & Parameter(0)
    For A = 1 To 9
        Message = Message & ", " & Parameter(A)
    Next A

    Open scriptsLog For Append As #1
    Print #1, Message
    Close #1
End Sub

Sub LogScriptEnd(Script As String)
    Dim Message As String
    Open scriptsLog For Append As #1
    Print #1, Now & "Script " & Script & " ended."
    Close #1
End Sub

Sub LogScriptNotExists(Script As String)
    Dim Message As String
    Open scriptsLog For Append As #1
    Print #1, Now & "Script " & Script & " was called but doesnt exist."
    Close #1
End Sub

Sub LogCrash(Message As String)
    Open errorHandlingLog For Append As #1
    Print #1, Now & " - " & Message
    Close #1

    PrintLog "ERROR HANDLED: " & Now & " - " & Message

    SendToGods Chr2(56) + Chr2(15) + "<GOD MESSAGE>ERROR HANDLED " & Message
   

End Sub

' old logging, ew
Sub PrintLog(St, Optional PrintToConsole As Boolean = True)
    If Dir("logs", vbDirectory) = "" Then
          MkDir "logs"
    End If
    
    CheckRollLog printDebugLog

    With frmMain.lstLog
        If PrintToConsole Then
            .AddItem Now & " - " & St
            If .ListCount > 200 Then .RemoveItem 0
            If .ListIndex = .ListCount - 2 Then .ListIndex = .ListCount - 1
        End If
        
        Open printDebugLog For Append As #1
        Print #1, Now & " - " & St
        Close #1
    End With
End Sub


Sub PrintDebug(St As String)
    If Dir("logs", vbDirectory) = "" Then
          MkDir "logs"
    End If
    
    CheckRollLog printDebugLog
    
    Open printDebugLog For Append As #1
    Print #1, Now & " - " & St
    Close #1
End Sub

Sub PrintCrashDebug(ByRef A As Long, ByRef B As Long)
    If Dir("logs", vbDirectory) = "" Then
          MkDir "logs"
    End If

    CheckRollLog debugCrashLog

    Open debugCrashLog For Append As #1
    Print #1, Now & " - " & (A) & ", " & (B)
    Close #1
    
End Sub
