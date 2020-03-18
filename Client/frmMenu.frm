VERSION 5.00
Begin VB.Form frmMenu 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Seyerdin Online"
   ClientHeight    =   4890
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5790
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   11.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmMenu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   326
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   386
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtEmail 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0084D3E8&
      Height          =   315
      Left            =   1680
      MaxLength       =   100
      TabIndex        =   3
      Top             =   1920
      Visible         =   0   'False
      Width           =   3000
   End
   Begin VB.TextBox txtPass2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0084D3E8&
      Height          =   315
      Left            =   1680
      MaxLength       =   15
      TabIndex        =   2
      Top             =   3600
      Visible         =   0   'False
      Width           =   3000
   End
   Begin VB.TextBox txtPass 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0084D3E8&
      Height          =   315
      Left            =   1680
      MaxLength       =   15
      TabIndex        =   1
      Top             =   3240
      Visible         =   0   'False
      Width           =   3000
   End
   Begin VB.TextBox txtUser 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0084D3E8&
      Height          =   315
      Left            =   1680
      MaxLength       =   15
      TabIndex        =   0
      Top             =   2400
      Visible         =   0   'False
      Width           =   3000
   End
End
Attribute VB_Name = "frmMenu"
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

'VERY SIMPLE MENU SYSTEM THAT I AM NOT PROUD OF, BUT GETS THE JOB DONE
Private Type MenuWidgetData
    x As Long
    y As Long
    srcX As Long
    srcY As Long
    Width As Long
    Height As Long
    Highlighted As Boolean
    Clicked As Boolean
    Selected As Boolean
    Caption As String
    Type As Long
    Flags As Long
    Color As Long
End Type

Private Type MenuType
    Caption As String
    Enabled As Boolean
    Items() As MenuWidgetData
End Type

Public CurrentMenu As Long
Private CurrentClass As Long 'Only for New Character
Private HighlightedClass As Long 'Only for New Character

Private Const MBUTTON As Long = 1
Private Const MLABEL As Long = 2
Private Const MTOGGLE As Long = 3
Private Const MRADIO As Long = 4
Public clickedPlay As Boolean
'Buttons(Menu, Button)
Private Const MENU_MENU = 1
    Private Const BtnMenuPlay = 1
    Private Const BtnMenuCreate = 2
    Private Const BtnMenuCredits = 3
    Private Const BtnMenuOptions = 4
    Private Const BtnMenuMinimize = 5
    Private Const BtnMenuClose = 6
    Private Const chkMenuSavePassword = 7
    Private Const lblMenuStatus = 8
Private Const MENU_CREATEACCOUNT = 2
    Private Const BtnCreateAccountCancel = 1
    Private Const BtnCreateAccountCreate = 2
    Private Const BtnCreateAccountMinimize = 3
    Private Const BtnCreateAccountClose = 4
    Private Const LblCreateAccountStatus = 5
Private Const MENU_CHARACTER = 3
    Private Const BtnCharacterCreate = 1
    Private Const BtnCharacterChangePassword = 2
    Private Const BtnCharacterDelete = 3
    Private Const BtnCharacterCancel = 4
    Private Const BtnCharacterPlay = 5
    Private Const BtnCharacterMinimize = 6
    Private Const BtnCharacterClose = 7
    Private Const LblCharacterStatus = 8
Private Const MENU_CREDITS = 4
    Private Const BtnCreditsCancel = 1
    Private Const BtnCreditsMinimize = 2
    Private Const BtnCreditsClose = 3
Private Const MENU_NEWPASSWORD = 5
    Private Const BtnNewPasswordAccept = 1
    Private Const BtnNewPasswordCancel = 2
    Private Const BtnNewPasswordMinimize = 3
    Private Const BtnNewPasswordClose = 4
    Private Const LblNewPasswordStatus = 5
Private Const MENU_NEWCHARACTER = 6
    Private Const BtnNewCharacterCreate = 1
    Private Const BtnNewCharacterCancel = 2
    Private Const ChkNewCharacterMale = 3
    Private Const ChkNewCharacterFemale = 4
    Private Const BtnNewCharacterMinimize = 5
    Private Const BtnNewCharacterClose = 6
    Private Const LblNewCharacterStatus = 7

Private Menu(1 To 6) As MenuType
Private Items(1 To 7) As MenuWidgetData

Public Sub SetStatusText(StatusText As String, ColorCode As Byte)
    Dim Color As Long
    
    Select Case ColorCode
        Case 1 'Normal
            Color = vbWhite
        Case 2 'Error
            Color = &H4457FF
    End Select
    Select Case CurrentMenu
        Case MENU_MENU
            Menu(MENU_MENU).Items(lblMenuStatus).Caption = StatusText
            Menu(MENU_MENU).Items(lblMenuStatus).Color = Color
        Case MENU_CREATEACCOUNT
            Menu(MENU_CREATEACCOUNT).Items(LblCreateAccountStatus).Caption = StatusText
            Menu(MENU_CREATEACCOUNT).Items(LblCreateAccountStatus).Color = Color
        Case MENU_CHARACTER
            Menu(MENU_CHARACTER).Items(LblCharacterStatus).Caption = StatusText
            Menu(MENU_CHARACTER).Items(LblCharacterStatus).Color = Color
        Case MENU_NEWPASSWORD
            Menu(MENU_NEWPASSWORD).Items(LblNewPasswordStatus).Caption = StatusText
            Menu(MENU_NEWPASSWORD).Items(LblNewPasswordStatus).Color = Color
        Case MENU_NEWCHARACTER
            Menu(MENU_NEWCHARACTER).Items(LblNewCharacterStatus).Caption = StatusText
            Menu(MENU_NEWCHARACTER).Items(LblNewCharacterStatus).Color = Color
    End Select
    DrawMenu
End Sub

Public Sub SetMenu(MenuNum As Long)
    Dim pControl As Control
    For Each pControl In frmMenu.Controls
        pControl.Visible = False
    Next
    CurrentMenu = MenuNum
    Me.Caption = Menu(MenuNum).Caption
    Menu(CurrentMenu).Enabled = True
    If CurrentMenu <> MENU_MENU Then Menu(MENU_MENU).Items(lblMenuStatus).Caption = ""
    Select Case MenuNum
        Case MENU_MENU
            txtUser.Left = 153
            txtUser.Top = 197
            txtUser.Height = 21
            txtPass.Left = 153
            txtPass.Top = 230
            txtPass.Height = 21
            txtUser = User
            txtUser.SelStart = Len(txtUser)
            txtUser.SelLength = 0
            txtPass.PasswordChar = "*"
            txtPass = Pass
            txtUser.Visible = True
            txtPass.Visible = True
            Menu(MENU_MENU).Items(lblMenuStatus).Caption = ""
            Set Me.Picture = LoadPicture("Data/Graphics/Interface/Menu.rsc")
            If ReadInt("Login", "SavePassword") > 0 Then
                Menu(MENU_MENU).Items(chkMenuSavePassword).Selected = True
                Pass = ReadStr("Login", "Password")
                txtPass = Pass
            Else
                Menu(MENU_MENU).Items(chkMenuSavePassword).Selected = False
                Pass = ""
                txtPass = ""
            End If
            DrawMenu
        Case MENU_CREDITS
            Set Me.Picture = LoadPicture("Data/Graphics/Interface/Credits.rsc")
        Case MENU_CREATEACCOUNT
            txtUser = ""
            txtPass = ""
            txtUser.Top = 131
            txtUser.Left = 153
            txtPass.Top = 164
            txtPass.Left = 153
            txtPass2.Top = 197
            txtPass2.Left = 153
            txtEmail.Left = 153
            txtEmail.Top = 230
            txtPass.PasswordChar = ""
            txtUser.Visible = True
            txtPass.Visible = True
            txtPass2.Visible = True
            txtEmail.Visible = True
            Menu(MENU_CREATEACCOUNT).Items(LblCreateAccountStatus).Caption = ""
            Set Me.Picture = LoadPicture("Data/Graphics/Interface/Account.rsc")
        Case MENU_CHARACTER
            Menu(MENU_CHARACTER).Items(LblCharacterStatus).Caption = ""
            Set Me.Picture = LoadPicture("Data/Graphics/Interface/Character.rsc")
        Case MENU_NEWPASSWORD
            Menu(MENU_NEWPASSWORD).Items(LblNewPasswordStatus).Caption = ""
            Set Me.Picture = LoadPicture("Data/Graphics/Interface/NewPassword.rsc")
            txtPass.Top = 200
            txtPass.Left = 153
            txtPass2.Top = 233
            txtPass2.Left = 153
            txtPass.PasswordChar = ""
            txtPass2.PasswordChar = ""
            txtPass = ""
            txtPass2 = ""
            txtPass.Visible = True
            txtPass2.Visible = True
        Case MENU_NEWCHARACTER
            Set Me.Picture = LoadPicture("Data/Graphics/Interface/NewCharacter.rsc")
            txtUser.Top = 36
            txtUser.Left = 105
            txtEmail.Left = 153
            txtEmail.Top = 230
            txtUser.Visible = True
            Menu(MENU_NEWCHARACTER).Items(LblNewCharacterStatus).Caption = ""
            Menu(MENU_NEWCHARACTER).Items(ChkNewCharacterMale).Selected = True
            CurrentClass = 4
            Menu(MENU_NEWCHARACTER).Items(LblNewCharacterStatus).Caption = Class(4).Description
            HighlightedClass = 0
            DrawMenu
    End Select
End Sub

Private Sub MenuCreateButton(pWidget As MenuWidgetData, x As Long, y As Long, srcX As Long, srcY As Long, Width As Long, Height As Long)
    pWidget.x = x
    pWidget.y = y
    pWidget.srcX = srcX
    pWidget.srcY = srcY
    pWidget.Width = Width
    pWidget.Height = Height
    pWidget.Type = MBUTTON
End Sub

Private Sub MenuCreateToggle(pWidget As MenuWidgetData, x As Long, y As Long, srcX As Long, srcY As Long, Width As Long, Height As Long)
    pWidget.x = x
    pWidget.y = y
    pWidget.srcX = srcX
    pWidget.srcY = srcY
    pWidget.Width = Width
    pWidget.Height = Height
    pWidget.Type = MTOGGLE
End Sub

Private Sub MenuCreateRadio(pWidget As MenuWidgetData, x As Long, y As Long, srcX As Long, srcY As Long, Width As Long, Height As Long)
    pWidget.x = x
    pWidget.y = y
    pWidget.srcX = srcX
    pWidget.srcY = srcY
    pWidget.Width = Width
    pWidget.Height = Height
    pWidget.Type = MRADIO
End Sub

Private Sub MenuCreateLabel(pWidget As MenuWidgetData, x As Long, y As Long, Caption As String)
    pWidget.x = x
    pWidget.y = y
    pWidget.Caption = Caption
    pWidget.Type = MLABEL
    pWidget.Color = vbWhite
End Sub

Private Sub MenuCreateLabelEX(pWidget As MenuWidgetData, x As Long, y As Long, Width As Long, Height As Long, Flags As Long, Caption As String)
    pWidget.x = x
    pWidget.y = y
    pWidget.Width = Width
    pWidget.Height = Height
    pWidget.Caption = Caption
    pWidget.Flags = Flags
    pWidget.Type = MLABEL
End Sub

Private Sub Form_Load()
    Dim A As Long

    gHW = Me.hwnd
    Set Me.Picture = LoadPicture("Data/Graphics/Interface/Menu.rsc")
    SetTextColor Me.hdc, vbWhite
    For A = 1 To MAX_CLASS
        'cmbClass.AddItem Class(A).Name
    Next A
    
    With Menu(MENU_MENU)
        .Caption = "Seyerdin Online [Menu]"
        ReDim .Items(1 To 8)
        MenuCreateButton .Items(BtnMenuPlay), 31, 279, 145, 513, 78, 28
        MenuCreateButton .Items(BtnMenuCreate), 114, 279, 145, 541, 78, 28
        MenuCreateButton .Items(BtnMenuCredits), 196, 279, 145, 569, 78, 28
        MenuCreateButton .Items(BtnMenuOptions), 279, 279, 145, 597, 78, 28
        MenuCreateButton .Items(BtnMenuMinimize), 340, 7, 175, 625, 17, 17
        MenuCreateButton .Items(BtnMenuClose), 360, 7, 209, 625, 17, 17
        MenuCreateToggle .Items(chkMenuSavePassword), 358, 235, 145, 625, 15, 15
        MenuCreateLabel .Items(lblMenuStatus), 193, 266, ""
    End With
    With Menu(MENU_CREATEACCOUNT)
        .Caption = "Seyerdin Online [Create Account]"
        ReDim .Items(1 To 5)
        MenuCreateButton .Items(BtnCreateAccountCreate), 191, 279, 145, 541, 78, 28
        MenuCreateButton .Items(BtnCreateAccountCancel), 273, 279, 145, 485, 78, 28
        MenuCreateButton .Items(BtnCreateAccountMinimize), 340, 7, 175, 625, 17, 17
        MenuCreateButton .Items(BtnCreateAccountClose), 360, 7, 209, 625, 17, 17
        MenuCreateLabel .Items(LblCreateAccountStatus), 193, 266, ""
    End With
    With Menu(MENU_CHARACTER)
        .Caption = "Seyerdin Online [Character]"
        ReDim .Items(1 To 8)
        MenuCreateButton .Items(BtnCharacterCreate), 30, 286, 145, 541, 78, 28
        MenuCreateButton .Items(BtnCharacterChangePassword), 113, 286, 145, 429, 78, 28
        MenuCreateButton .Items(BtnCharacterDelete), 196, 286, 145, 457, 78, 28
        MenuCreateButton .Items(BtnCharacterCancel), 279, 286, 145, 485, 78, 28
        MenuCreateButton .Items(BtnCharacterPlay), 153, 249, 145, 513, 78, 28
        MenuCreateButton .Items(BtnCharacterMinimize), 340, 7, 175, 625, 17, 17
        MenuCreateButton .Items(BtnCharacterClose), 360, 7, 209, 625, 17, 17
        MenuCreateLabel .Items(LblCharacterStatus), 193, 235, ""
    End With
    With Menu(MENU_CREDITS)
        .Caption = "Seyerdin Online [Credits]"
        ReDim .Items(1 To 3)
        MenuCreateButton .Items(BtnCreditsCancel), 273, 279, 145, 485, 78, 28
        MenuCreateButton .Items(BtnCreditsMinimize), 340, 7, 175, 625, 17, 17
        MenuCreateButton .Items(BtnCreditsClose), 360, 7, 209, 625, 17, 17
    End With
    With Menu(MENU_NEWPASSWORD)
        .Caption = "Seyerdin Online [New Password]"
        ReDim .Items(1 To 5)
        MenuCreateButton .Items(BtnNewPasswordAccept), 191, 279, 301, 429, 78, 28
        MenuCreateButton .Items(BtnNewPasswordCancel), 273, 279, 145, 485, 78, 28
        MenuCreateButton .Items(BtnNewPasswordMinimize), 340, 7, 175, 625, 17, 17
        MenuCreateButton .Items(BtnNewPasswordClose), 360, 7, 209, 625, 17, 17
        MenuCreateLabel .Items(LblNewPasswordStatus), 193, 259, ""
    End With
    With Menu(MENU_NEWCHARACTER)
        .Caption = "Seyerdin Online [New Character]"
        ReDim .Items(1 To 7)
        MenuCreateButton .Items(BtnNewCharacterCreate), 113, 286, 145, 541, 78, 28
        MenuCreateButton .Items(BtnNewCharacterCancel), 196, 286, 145, 485, 78, 28
        MenuCreateButton .Items(BtnNewCharacterMinimize), 340, 7, 175, 625, 17, 17
        MenuCreateButton .Items(BtnNewCharacterClose), 360, 7, 209, 625, 17, 17
        MenuCreateRadio .Items(ChkNewCharacterMale), 126, 66, 145, 625, 15, 15
        MenuCreateRadio .Items(ChkNewCharacterFemale), 200, 66, 145, 625, 15, 15
        MenuCreateLabelEX .Items(LblNewCharacterStatus), 35, 218, 316, 55, DT_WORDBREAK Or DT_CENTER Or DT_VCENTER, ""
    End With

    User = ReadStr("Login", "User")


    'SetMenu MENU_MENU
    'DrawMenu
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim A As Long, Redraw As Boolean
    Redraw = False
    'If Menu(CurrentMenu).Enabled = False Then Exit Sub
    For A = 1 To UBound(Menu(CurrentMenu).Items)
        With Menu(CurrentMenu).Items(A)
            If x >= .x And x <= .x + .Width Then
                If y >= .y And y <= .y + .Height Then
                    If .Clicked = False Then
                        .Clicked = True
                        Redraw = True
                    End If
                Else
                    If .Clicked = True Then
                        .Clicked = False
                        Redraw = True
                    End If
                End If
            Else
                If .Clicked = True Then
                    .Clicked = False
                    Redraw = True
                End If
            End If
        End With
    Next A
    
    
    A = 0
    If Redraw Then
        DrawMenu
        A = 1
    End If
    If CurrentMenu = MENU_NEWCHARACTER Then
        If y >= 93 And y <= 124 Then
            If x >= 48 And x <= 335 Then
                A = 1
            End If
        End If
    End If
    If A = 0 Then
        Dim ReturnVal As Long
        ReleaseCapture
        ReturnVal = SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim A As Long, Redraw As Boolean
    Redraw = False
    'If Menu(CurrentMenu).Enabled = False Then Exit Sub
    For A = 1 To UBound(Menu(CurrentMenu).Items)
        With Menu(CurrentMenu).Items(A)
            If x >= .x And x <= .x + .Width Then
                If y >= .y And y <= .y + .Height Then
                    If .Highlighted = False Then
                        .Highlighted = True
                        Redraw = True
                    End If
                Else
                    If .Highlighted = True Then
                        .Highlighted = False
                        Redraw = True
                    End If
                End If
            Else
                If .Highlighted = True Then
                    .Highlighted = False
                    Redraw = True
                End If
            End If
        End With
    Next A
    
    If CurrentMenu = MENU_NEWCHARACTER Then
        If x >= 48 And y >= 93 And x <= 335 And y <= 124 Then
            A = ((x - 45) \ 32) + 1
            If A < 1 Then A = 1
            If A > MAX_CLASS Then A = MAX_CLASS
            If HighlightedClass <> A And A <> CurrentClass Then
                HighlightedClass = A
                Redraw = True
            End If
        Else
            HighlightedClass = 0
            Redraw = True
        End If
    End If
    
    If Redraw Then DrawMenu
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim A As Long, b As Long, Redraw As Boolean
    Redraw = False
    'If Menu(CurrentMenu).Enabled = False Then Exit Sub
    For A = 1 To UBound(Menu(CurrentMenu).Items)
        With Menu(CurrentMenu).Items(A)
            If x >= .x And x <= .x + .Width Then
                If y >= .y And y <= .y + .Height Then
                    Select Case .Type
                        Case MBUTTON, MTOGGLE
                            If .Clicked = True Then
                                .Clicked = False
                                If .Type = MTOGGLE Then .Selected = Not .Selected
                                Redraw = True
                                If DoItemEvent(CurrentMenu, A) Then Exit For
                            End If
                        Case MRADIO
                            If .Clicked = True Then
                                If .Selected = False Then
                                    .Selected = True
                                    For b = 1 To UBound(Menu(CurrentMenu).Items)
                                        If Menu(CurrentMenu).Items(b).Type = MRADIO And b <> A Then
                                            Menu(CurrentMenu).Items(b).Selected = False
                                        End If
                                    Next b
                                    Redraw = True
                                End If
                            End If
                    End Select
                Else
                    If .Clicked = True Then
                        .Clicked = False
                        Redraw = True
                    End If
                End If
            Else
                If .Clicked = True Then
                    .Clicked = False
                    Redraw = True
                End If
            End If
        End With
    Next A
    
    If CurrentMenu = MENU_NEWCHARACTER Then
        If x >= 48 And y >= 93 Then
            If x <= 335 And y <= 124 Then
                A = ((x - 48) \ 32) + 1
                If A < 1 Then A = 1
                If A > MAX_CLASS Then A = MAX_CLASS
                If CurrentClass <> A Then
                    If Class(A).Enabled Then
                        CurrentClass = A
                        Menu(MENU_NEWCHARACTER).Items(LblNewCharacterStatus).Caption = Class(A).Description
                        Redraw = True
                    End If
                End If
            End If
        End If
    End If
    
    If Redraw Then DrawMenu
End Sub

Private Function DoItemEvent(MenuNum As Long, ItemNum As Long) As Boolean
    Dim St As String, A As Long
    With Menu(MenuNum).Items(ItemNum)
        Select Case MenuNum
            Case MENU_MENU
                Select Case ItemNum
                    Case BtnMenuPlay
                        If Len(txtUser) >= 3 And Len(txtPass) >= 3 Then
                            clickedPlay = False
                            User = txtUser
                            Pass = txtPass
                            WriteString "Login", "User", User
                            If Menu(MenuNum).Items(chkMenuSavePassword).Selected Then
                                WriteString "Login", "Password", Pass
                            Else
                                DelItem "Login", "Password"
                            End If
                            NewAccount = False
                            SetStatusText "Connecting . . .", 1
                            DrawMenu
                            ConnectClient
                            DoItemEvent = True
                        Else
                            SetStatusText "Username/Password must be at least 3 characters.", 2
                        End If
                    Case BtnMenuCreate
                        SetMenu MENU_CREATEACCOUNT
                        DoItemEvent = True
                    Case BtnMenuCredits
                        DoItemEvent = True
                        SetMenu MENU_CREDITS
                    Case BtnMenuOptions
                        frmOptions.Show
                        Me.Hide
                    Case BtnMenuMinimize
                        Me.WindowState = vbMinimized
                    Case BtnMenuClose
                        blnEnd = True
                    Case chkMenuSavePassword
                        WriteString "Login", "SavePassword", IIf(Menu(MenuNum).Items(chkMenuSavePassword).Selected, 1, 0)
                End Select
            Case MENU_CREATEACCOUNT
                Select Case ItemNum
                    Case BtnCreateAccountCreate
                        txtUser = Trim$(txtUser)
                        txtPass = Trim$(txtPass)
                        txtPass2 = Trim$(txtPass2)
                        txtEmail = Trim$(txtEmail)
                        
                        If Len(txtUser) >= 3 Then
                            If Len(txtPass) >= 3 Then
                                A = Asc(Left$(txtUser, 1))
                                If (A >= 65 And A <= 90) Or (A >= 97 And A <= 122) Then
                                    If txtPass = txtPass2 Then
                                        If txtEmail Like "*@*.*" Then
                                            User = txtUser
                                            Pass = txtPass
                                            Email = txtEmail
                                            NewAccount = True

                                            WriteString "Login", "User", User
                                            SetStatusText "Connecting . . .", 1
                                            ClientSocket = ConnectSock(ServerIP, ServerPort, St, gHW, True)
                                        Else
                                            SetStatusText "Invalid Email format.", 2
                                        End If
                                    Else
                                        SetStatusText "Passwords do not match.", 2
                                    End If
                                Else
                                    SetStatusText "Username must start with letter.", 2
                                End If
                            Else
                                SetStatusText "Password must be at least 3 characters long.", 2
                            End If
                        Else
                            SetStatusText "Username must be at least 3 characters long.", 2
                        End If
                    Case BtnCreateAccountCancel, BtnCreateAccountClose
                        SetMenu MENU_MENU
                        DoItemEvent = True
                    Case BtnCreditsMinimize
                        Me.WindowState = vbMinimized
                End Select
            Case MENU_CHARACTER
                Select Case ItemNum
                    Case BtnCharacterPlay
                        
                        
                        If Character.Level > 0 And Character.Class > 0 Then
                            If Not clickedPlay Then
                                clickedPlay = True
                                CMap = 0
                                TargetMonster = 10
                                CMap2 = 5
                                cX = 0
                                cY = 0
                                CX2 = 5
                                CY2 = 5
                                CurFog = 0
                                For A = 1 To MAXUSERS
                                    With player(A)
                                        .Sprite = 0
                                        .map = 0
                                    End With
                                Next A
                                With Character
                                    For A = 1 To 20
                                        With .Inv(A)
                                            .Object = 0
                                            .Value = 0
                                        End With
                                    Next A
                                    For A = 1 To 5
                                        With .Equipped(A)
                                            .Object = 0
                                            .Value = 0
                                        End With
                                    Next A
                                End With
                                keyAlt = False
                                SetStatusText "Connecting . . .", 1
                                SendSocket Chr$(5) 'I wanna play
                                Menu(MENU_CHARACTER).Enabled = False
                            End If
                        Else
                            SetStatusText "You must create a character first.", 2
                        End If

                    Case BtnCharacterChangePassword
                        SetMenu MENU_NEWPASSWORD
                        DoItemEvent = True
                    Case BtnCharacterCreate
                        If Character.Level > 0 And Character.Class > 0 Then
                            If MsgBox("Creating a new character will overwrite your old character, are you sure you wish to continue?", vbYesNo + vbExclamation, TitleString) = vbYes Then
                                If MsgBox("Last chance to back out -- are you *sure* you wish to create a new character?", vbYesNo + vbExclamation, TitleString) = vbYes Then
                                    SetMenu MENU_NEWCHARACTER
                                    DoItemEvent = True
                                End If
                            End If
                        Else
                            SetMenu MENU_NEWCHARACTER
                            DoItemEvent = True
                        End If
                    Case BtnCharacterDelete
                    'MsgBox ("Deleting temporarily disabled")
                        'If MsgBox("This will permanently erase your character and account, are you *sure* you want to continue?", vbYesNo + vbExclamation, TitleString) = vbYes Then
                        '    If MsgBox("Last chance to back out -- are you *sure* you wish to delete you account?", vbYesNo + vbExclamation, TitleString) = vbYes Then
                        '        SendSocket Chr$(4)
                        '    End If
                        'End If
                    Case BtnCharacterClose, BtnCharacterCancel
                        SendSocket Chr$(30)
                        CloseClientSocket 0, True
                        SetMenu MENU_MENU
                        DoItemEvent = True
                    Case BtnCharacterMinimize
                        Me.WindowState = vbMinimized
                End Select
            Case MENU_CREDITS
                Select Case ItemNum
                    Case BtnCreditsCancel, BtnCreditsClose
                        SetMenu MENU_MENU
                        DoItemEvent = True
                    Case BtnCreditsMinimize
                        Me.WindowState = vbMinimized
                End Select
            Case MENU_NEWPASSWORD
                Select Case ItemNum
                    Case BtnNewPasswordCancel, BtnNewPasswordClose
                        SetMenu MENU_CHARACTER
                        DoItemEvent = True
                    Case BtnNewPasswordAccept
                        If Len(txtPass) >= 3 Then
                            If txtPass = txtPass2 Then
                                SendSocket Chr$(3) + SHA256(UCase(txtUser) & UCase(txtPass))
                                SetStatusText "Changing Password . . .", 1
                            Else
                                SetStatusText "Passwords do not match.", 2
                            End If
                        Else
                            SetStatusText "Minimum password length is 3 characters.", 2
                        End If
                    Case BtnNewPasswordMinimize
                        Me.WindowState = vbMinimized
                End Select
            Case MENU_NEWCHARACTER
                Select Case ItemNum
                    Case BtnNewCharacterCancel, BtnNewCharacterClose
                        SetMenu MENU_CHARACTER
                        DoItemEvent = True
                    Case BtnNewCharacterCreate
                        If Len(txtUser) >= 3 Then
                            A = Asc(Mid$(txtUser, 1, 1))
                            If (A >= 65 And A <= 90) Or (A >= 97 And A <= 122) Then
                                If InStr(1, txtUser, " ") = 0 Then
                                    If InStr(1, txtUser, "_") = 0 Then
                                        SetStatusText "Creating new character . . .", 1
                                        SendSocket Chr$(2) + Chr$(CurrentClass) + Chr$(1) + Chr$(IIf(Menu(MENU_NEWCHARACTER).Items(ChkNewCharacterMale).Selected, 0, 1)) + txtUser + Chr$(0) + ""
                                    Else
                                        SetStatusText "Names cannot end in underscores.", 2
                                    End If
                                Else
                                    SetStatusText "Names cannot contain spaces.", 2
                                End If
                            Else
                                SetStatusText "Name must start with a letter.", 2
                            End If
                        Else
                            SetStatusText "Name must be at least 3 characters long.", 2
                        End If
                    Case BtnNewCharacterMinimize
                        Me.WindowState = vbMinimized
                End Select
        End Select
    End With
End Function

Private Sub DrawMenu()
    On Error Resume Next
    Dim A As Long, R1 As RECT, r2 As RECT
    frmMenu.Cls
    Select Case CurrentMenu
        Case MENU_CHARACTER
            DrawCharacterData Character.Class > 0
        Case MENU_NEWCHARACTER
            DrawNewCharacterData
    End Select
    For A = 1 To UBound(Menu(CurrentMenu).Items)
        With Menu(CurrentMenu).Items(A)
            Select Case .Type
                Case MBUTTON, MTOGGLE, MRADIO
                    If (.Clicked = True And .Highlighted = True) Or (.Selected = True) Then
                        R1.Top = .srcY: R1.Bottom = R1.Top + .Height: R1.Left = .srcX + .Width: R1.Right = R1.Left + .Width
                        r2.Top = .y: r2.Bottom = r2.Top + .Height: r2.Left = .x: r2.Right = r2.Left + .Width
                        sfcInventory(0).Surface.BltToDC frmMenu.hdc, R1, r2
                    ElseIf .Highlighted = True Then
                        R1.Top = .srcY: R1.Bottom = R1.Top + .Height: R1.Left = .srcX: R1.Right = R1.Left + .Width
                        r2.Top = .y: r2.Bottom = r2.Top + .Height: r2.Left = .x: r2.Right = r2.Left + .Width
                        sfcInventory(0).Surface.BltToDC frmMenu.hdc, R1, r2
                    End If
                Case MLABEL
                    Me.ForeColor = .Color
                    If .Width Or .Height Then
                        R1.Left = .x: R1.Right = .x + .Width: R1.Top = .y: R1.Bottom = .y + .Height
                        DrawText frmMenu.hdc, .Caption, Len(.Caption), R1, .Flags
                    Else
                        TextOut frmMenu.hdc, .x - (Me.TextWidth(.Caption) \ 2), .y - (Me.textHeight(.Caption) \ 2), .Caption, Len(.Caption)
                    End If
            End Select
        End With
    Next A
    frmMenu.Refresh
End Sub

Public Sub DrawCharacterData(HasCharacter As Boolean)
    Dim St As String
    Dim R1 As RECT, r2 As RECT
    'frmMenu.Cls
    Me.ForeColor = vbWhite
    If HasCharacter Then
        R1.Left = ((Character.Sprite - 1) Mod 16) * 32: R1.Right = R1.Left + 32: R1.Top = ((Character.Sprite - 1) \ 16) * 32: R1.Bottom = R1.Top + 32
        r2.Left = 37: r2.Right = r2.Left + 32: r2.Top = 36: r2.Bottom = r2.Top + 32
        sfcSprites.Surface.BltToDC Me.hdc, R1, r2
        St = Character.Name
        TextOut Me.hdc, 193 - frmMenu.TextWidth(St) \ 2, 43, St, Len(St)
        St = "Level " & Character.Level & " " & Class(Character.Class).Name
        TextOut Me.hdc, 193 - frmMenu.TextWidth(St) \ 2, 59, St, Len(St)
        If Character.Guild > 0 Then
            St = Choose(Character.GuildRank + 1, "Initiate", "Member", "Officer", "Guildmaster") & " of " & Guild(Character.Guild).Name
            TextOut Me.hdc, 193 - frmMenu.TextWidth(St) \ 2, 76, St, Len(St)
        End If
        
        Dim Spacing As String
        Spacing = Space$(13)
        St = "HP: " & Character.MaxHP & Spacing & "Energy: " & Character.MaxEnergy & Spacing & "Mana: " & Character.MaxMana
        TextOut Me.hdc, 193 - frmMenu.TextWidth(St) \ 2, 110, St, Len(St)
        Spacing = Space$(9)
        St = "Strength: " & Character.strength & Spacing & "Agility: " & Character.Agility & Spacing & "Endurance: " & Character.Endurance
        TextOut Me.hdc, 193 - frmMenu.TextWidth(St) \ 2, 143, St, Len(St)
        Spacing = Space$(4)
        St = "Piety: " & Character.Wisdom & Spacing & "Intelligence: " & Character.Intelligence & Spacing & "Constitution: " & Character.Constitution
        TextOut Me.hdc, 193 - frmMenu.TextWidth(St) \ 2, 160, St, Len(St)
    Else
        St = "Click Create to create your character!"
        TextOut Me.hdc, 193 - frmMenu.TextWidth(St) \ 2, 160, St, Len(St)
    End If
    'Me.Refresh
End Sub
Public Sub DrawNewCharacterData()
    Dim St As String, A As Long, b As Long
    Dim R1 As RECT, r2 As RECT
    'frmMenu.Cls
    Me.ForeColor = vbWhite
    b = IIf(Menu(MENU_NEWCHARACTER).Items(ChkNewCharacterMale).Selected, 0, 1)
    For A = 1 To MAX_CLASS - 1
        If Class(A).Enabled = 1 Then
            If CurrentClass = A Then
                Draw sfcCursor, 0, 0, 32, 32, sfcInventory(0), 254, 292
            ElseIf HighlightedClass = A Then
                Draw sfcCursor, 0, 0, 32, 32, sfcInventory(0), 222, 292
            Else
                Draw sfcCursor, 0, 0, 32, 32, sfcInventory(0), 190, 292
            End If
            If b = 0 Then
                Draw sfcCursor, 0, 0, 32, 32, sfcSprites, ((Class(A).MaleSprite - 1) Mod 16) * 32, ((Class(A).MaleSprite - 1) \ 16) * 32, True
            Else
                Draw sfcCursor, 0, 0, 32, 32, sfcSprites, ((Class(A).FemaleSprite - 1) Mod 16) * 32, ((Class(A).FemaleSprite - 1) \ 16) * 32, True
            End If
            R1.Top = 0: R1.Bottom = R1.Top + 32: R1.Left = 0: R1.Right = R1.Left + 32
            r2.Top = 93: r2.Bottom = r2.Top + 32
            r2.Left = 48 + (A - 1) * 32: r2.Right = r2.Left + 32
            sfcCursor.Surface.BltToDC frmMenu.hdc, R1, r2
        End If
    Next A
    
    St = Class(CurrentClass).Name
    TextOut Me.hdc, 58, 133, St, Len(St)
    
    St = Class(CurrentClass).StartHP
    TextOut Me.hdc, 98, 153, St, Len(St)
    St = Class(CurrentClass).StartEnergy
    TextOut Me.hdc, 98, 169, St, Len(St)
    St = Class(CurrentClass).StartMana
    TextOut Me.hdc, 98, 185, St, Len(St)
    
    St = Class(CurrentClass).StartStrength
    TextOut Me.hdc, 308, 128, St, Len(St)
    St = Class(CurrentClass).StartAgility
    TextOut Me.hdc, 308, 143, St, Len(St)
    St = Class(CurrentClass).StartEndurance
    TextOut Me.hdc, 308, 158, St, Len(St)
    St = Class(CurrentClass).StartWisdom
    TextOut Me.hdc, 308, 173, St, Len(St)
    St = Class(CurrentClass).StartConstitution
    TextOut Me.hdc, 308, 188, St, Len(St)
    St = Class(CurrentClass).StartIntelligence
    TextOut Me.hdc, 308, 203, St, Len(St)
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set Me.Picture = Nothing
End Sub

Private Sub txtPass_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Len(txtUser.Text) > 3 And Len(txtPass.Text) > 3 Then DoItemEvent CurrentMenu, 1
End Sub

Private Sub txtUser_KeyPress(KeyAscii As Integer)
    If CurrentMenu = MENU_CREATEACCOUNT Then
        If KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Or (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Or KeyAscii = 95 Then
            'Valid Key
        Else
            KeyAscii = 0
            'Beep
        End If
    End If
End Sub
