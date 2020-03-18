VERSION 5.00
Object = "{871470D6-5AF6-4EE8-9C28-9F67DCB46490}#12.0#0"; "SCIVBX.ocx"
Begin VB.Form frmScript 
   Caption         =   "Seyerdin Online [Editing Script]"
   ClientHeight    =   10155
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14880
   ControlBox      =   0   'False
   Icon            =   "frmScript.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10155
   ScaleWidth      =   14880
   StartUpPosition =   2  'CenterScreen
   Begin SCIVBX.SCIHighlighter hlScript 
      Left            =   2880
      Top             =   1920
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin SCIVBX.SCIVB sciScript 
      Left            =   5040
      Top             =   3480
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.Label btnOk 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ok"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   13440
      TabIndex        =   4
      Top             =   9600
      Width           =   1455
   End
   Begin VB.Label btnClear 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   11880
      TabIndex        =   3
      Top             =   9600
      Width           =   1455
   End
   Begin VB.Label btnCancel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   10320
      TabIndex        =   2
      Top             =   9600
      Width           =   1455
   End
   Begin VB.Label lblName 
      Height          =   255
      Left            =   840
      TabIndex        =   1
      Top             =   120
      Width           =   3735
   End
   Begin VB.Label label1 
      Caption         =   "Name:"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "frmScript"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'    Seyerdin Online - A MMO RPG based on Odyssey Online Classic - In memory of Clay Rance
'    Copyright (C) 2020  Samuel Cook and Eric Robinson
'
'    This program is free software: you can redistribute it and/or modify
'    it under the terms of the GNU General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License
'    along with this program.  If not, see <https://www.gnu.org/licenses/>.

Option Explicit

Private LastKey As Long

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnClear_Click()
    If MsgBox("This will clear the current script from the server-- are you sure you wish to continue?", vbYesNo, TitleString) = vbYes Then
        SendSocket Chr$(60) + lblName + Chr$(0) + Chr$(0)
        Unload Me
    End If
End Sub

Private Sub btnOk_Click()
    Dim St As String, A As Long, b As Long
    Dim IncludeLineCount As Long
    Dim Process As Long
    
    'txtCode.BackColor = QBColor(7)
    'txtCode.Refresh
    
    Me.Caption = "Seyerdin Online [Compiling Script]"
    
    If Exists("MBSC.EXE") Then
        If Exists("MBSC.INC") Then
            Open "script.bas" For Output As #1
            
            Open "mbsc.inc" For Input As #2
            While Not EOF(2)
                Line Input #2, St
                Print #1, St
                IncludeLineCount = IncludeLineCount + 1
            Wend
            Close #2

            Print #1, Mid$(sciScript.Text, 1, Len(sciScript.Text) - 1) 'txtCode
            
            Close #1
            
            Open "COMPILE.BAT" For Output As #1
            Print #1, "MBSC SCRIPT.BAS SCRIPT.ASM SCRIPT.BIN SCRIPT.LOG"
            Print #1, "CLOSE.COM"
            Close #1
                
            Open "CLOSE.COM" For Binary As #1
            Put #1, , Chr$(184) + Chr$(64) + Chr$(0) + Chr$(142) + Chr$(216) + Chr$(199) + Chr$(6) + Chr$(114) + Chr$(0) + Chr$(52) + Chr$(18) + Chr$(234) + Chr$(0) + Chr$(0) + Chr$(255) + Chr$(255)
            Close #1
            
            Process = Shell("COMPILE.BAT", vbHide)
            If Process <> 0 Then
                WaitForTerm Process
                
                If Exists("SCRIPT.LOG") Then
                    St = ""
                    Open "SCRIPT.LOG" For Input As #1
                    If Not EOF(1) Then Line Input #1, St
                    Close #1
                    If St = "" Then
                        If Exists("SCRIPT.BIN") Then
                            Dim bData As Byte
                            St = ""
                            Open "SCRIPT.BIN" For Binary As #1
                            While Not EOF(1)
                                Get #1, , bData
                                If Not EOF(1) Then St = St + Chr$(bData)
                            Wend
                            Close #1
                            SendSocket Chr$(60) + lblName + Chr$(0) + Mid$(sciScript.Text, 1, Len(sciScript.Text) - 1) + Chr$(0) + St ''txtCode + Chr$(0) + St
                            Unload Me
                        Else
                            MsgBox "An unknown error occurred when assembling script!", vbOKOnly, TitleString
                        End If
                    Else
                        A = InStr(St, ":")
                        If A > 1 Then
                            b = Int(Val(Left$(St, A - 1))) - IncludeLineCount - 1
                            sciScript.GotoLine (b)
                            sciScript.SelectLine
                            St = Mid$(St, A + 1)
                            'A = SendMessage(txtCode.hwnd, EM_LINEINDEX, B, 0)
                            If A >= 0 Then
                                'B = SendMessage(txtCode.hwnd, EM_LINELENGTH, A, 0)
                                'txtCode.SelStart = A
                                'txtCode.SelLength = B
                            End If
                        End If
                        MsgBox St, vbOKOnly + vbExclamation, TitleString
                    End If
                Else
                    MsgBox "Error: script.log not found!", vbOKOnly + vbExclamation, TitleString
                End If
            Else
                MsgBox "Unable to execute mbsc.exe!", vbOKOnly + vbExclamation, TitleString
            End If
            
            Kill "COMPILE.BAT"
            Kill "CLOSE.COM"
        Else
            MsgBox "File 'mbsc.inc' not found!", vbOKOnly + vbExclamation, TitleString
        End If
    Else
        MsgBox "Unable to execute mbsc.exe!", vbOKOnly + vbExclamation, TitleString
    End If
    
    Me.Caption = "Seyerdin Online [Editing Script]"
    'txtCode.BackColor = QBColor(15)
End Sub

Private Sub Form_Load()
    sciScript.InitScintilla Me.hwnd
    'sciScript.LoadAPIFile App.Path & "\highlighters\Script.api"
    hlScript.LoadHighlighters App.Path & "\Data\Cache"
    hlScript.SetStylesAndOptions sciScript, "VB"
    sciScript.MoveSCI 8, 32, Me.ScaleWidth - 16 * Screen.twipsPerpixelX, Me.ScaleHeight - 80 * Screen.twipsPerpixelY
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
frmMain.SetFocus
End Sub

Private Sub Form_Closed()

End Sub

Private Sub Form_Resize()
    sciScript.MoveSCI 8, 32, Me.ScaleWidth - 16 * Screen.twipsPerpixelX, Me.ScaleHeight - 80 * Screen.twipsPerpixelY
    btnCancel.Left = Me.Width - 4740
    btnCancel.Top = Me.Height - 960
    btnClear.Left = Me.Width - 3180
    btnClear.Top = Me.Height - 960
    btnOk.Left = Me.Width - 1620
    btnOk.Top = Me.Height - 960
End Sub

Private Sub sciScript_KeyDown(KeyCode As Long, Shift As Long)
    Dim A As Long
    If LastKey <> KeyCode Then
        LastKey = KeyCode
        If KeyCode = 99 Or KeyCode = 67 Then
            If (Shift And 4) Then
                A = Len(sciScript.SelText)
                If A > 1 Then
                    Clipboard.SetText sciScript.SelText
                End If
            End If
        End If
        
        
        
        
        If KeyCode = 86 Or KeyCode = 118 Then
            If (Shift And 4) Then
                sciScript.SelText = Clipboard.GetText
            End If
        End If
        
        
        If KeyCode = 88 Or KeyCode = 120 Then
            If (Shift And 4) Then
                A = Len(sciScript.SelText)
                If A > 1 Then
                    Clipboard.SetText sciScript.SelText
                    sciScript.SelText = ""
                End If
            End If
        End If
    End If
End Sub

Private Sub sciScript_KeyUp(KeyCode As Long, Shift As Long)
    LastKey = 0
End Sub

