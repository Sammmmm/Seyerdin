Attribute VB_Name = "modWidgets"
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

'Widget Constants
Global Widgets As New clsWidgets

Public Const DIALOG_TELL = 1
Public CurrentDialog As Long

'Remote Widget Stuff
    Public RemoteWidgetParent As String
'-------------------

    'Widget Types
        Public Const WIDGET_BUTTON = 1
        Public Const WIDGET_LABEL = 2
        Public Const WIDGET_FRAME = 3
        Public Const WIDGET_TEXTBOX = 4
        Public Const WIDGET_IMAGE = 5
    'Widget Style Flags
        Public Const STYLE_NOSTYLE = 0
        Public Const STYLE_SMALL = 1
        Public Const STYLE_MEDIUM = 2
        Public Const STYLE_LARGE = 4
        Public Const STYLE_DYNAMIC = 8
        Public Const STYLE_CENTERED = 16
        Public Const STYLE_SOLID = 32           'Sets back of label solid with black border
        Public Const STYLE_MULTILINE = 64
        Public Const STYLE_FADE = 128

Public Sub CreateDialogBox(DialogType As Long, x As Long, y As Long, Width As Long, Height As Long, Dialog As Long)

End Sub

Public Sub DrawWidgets(srcWidgets As clsWidgets)
    
    If Len(RemoteWidgetParent) > 0 Then
        Dim tmpWidget As clsWidget
    
        For Each tmpWidget In srcWidgets
            With tmpWidget
                If .Visible = True Then
                    Select Case .WidgetType
                        Case WIDGET_BUTTON
                            DrawButton tmpWidget
                        Case WIDGET_LABEL
                            DrawLabel tmpWidget
                        Case WIDGET_FRAME
                            DrawFrame tmpWidget
                        Case WIDGET_TEXTBOX
                            DrawTextBox tmpWidget
                        Case WIDGET_IMAGE
                            DrawWidgetImage tmpWidget
                    End Select
                    If .Children.Count > 0 Then DrawWidgets .Children
                End If
            End With
        Next
    End If
End Sub

Public Sub DrawFrame(wdgFrame As clsWidget)
    Dim WidthLeft As Long, WidthMul As Long
    Dim HeightLeft As Long, HeightMul As Long
    Dim A As Long, b As Long
    With wdgFrame
        If .Width > 48 And .Height > 48 Then
            'Draw Top Left Corner
            Draw3D .x, .y, 8, 8, 0, 0, TexControl1
            Draw3D .x + .Width - 8, .y, 8, 8, 42, 0, TexControl1
            Draw3D .x, .y + .Height - 8, 8, 8, 0, 42, TexControl1
            Draw3D .x + .Width - 8, .y + .Height - 8, 8, 8, 42, 42, TexControl1
            WidthLeft = .Width - 16 '48
            HeightLeft = .Height - 16

            WidthMul = Int(WidthLeft / 32)
            HeightMul = Int(HeightLeft / 32)

            For A = 0 To WidthMul
                For b = 0 To HeightMul
                    If A = 0 Then
                        If b < HeightMul Then
                            Draw3D .x, 8 + .y + (32 * b), 8, 32, 0, 9, TexControl1
                        Else
                            Draw3D .x, 8 + .y + (32 * b), 8, HeightLeft Mod 32, 0, 9, TexControl1
                        End If
                    End If
                    If b = 0 Then
                        If A < WidthMul Then
                            Draw3D 8 + .x + (32 * A), .y, 32, 8, 9, 0, TexControl1
                        Else
                            Draw3D 8 + .x + (32 * A), .y, WidthLeft Mod 32, 8, 9, 0, TexControl1
                        End If
                    End If

                    If b < HeightMul Then
                        If A < WidthMul Then
                            Draw3D 8 + .x + A * 32, 8 + .y + b * 32, 32, 32, 9, 9, TexControl1
                        Else
                            Draw3D 8 + .x + A * 32, 8 + .y + b * 32, WidthLeft Mod 32, 32, 9, 9, TexControl1
                        End If
                    Else
                        If A < WidthMul Then
                            Draw3D 8 + .x + A * 32, 8 + .y + b * 32, 32, HeightLeft Mod 32, 9, 9, TexControl1
                        Else
                            Draw3D 8 + .x + A * 32, 8 + .y + b * 32, WidthLeft Mod 32, HeightLeft Mod 32, 9, 9, TexControl1
                        End If
                    End If

                    If b = HeightMul Then
                        If A < WidthMul Then
                            Draw3D 8 + .x + (32 * A), .y + .Height - 8, 32, 8, 9, 42, TexControl1
                        Else
                            Draw3D 8 + .x + (32 * A), .y + .Height - 8, WidthLeft Mod 32, 8, 9, 42, TexControl1
                        End If
                    End If
                    If A = WidthMul Then
                        If b < HeightMul Then
                            Draw3D .x + .Width - 8, 8 + .y + (32 * b), 8, 32, 42, 9, TexControl1
                        Else
                            Draw3D .x + .Width - 8, 8 + .y + (32 * b), 8, HeightLeft Mod 32, 42, 9, TexControl1
                        End If
                    End If
                Next b
            Next A
        End If
    End With
End Sub



Public Sub DrawButton(wdgButton As clsWidget)
    Dim YSrcOffset As Long, TextColor As Long
    With wdgButton
        Select Case .Style
            Case STYLE_MEDIUM
                If .Selected Then
                    YSrcOffset = 48
                    TextColor = &HFF7F3F00
                ElseIf .Highlighted Then
                    YSrcOffset = 24
                    TextColor = &HFFF5E49C
                Else
                    YSrcOffset = 0
                    TextColor = &HFFFFFFFF
                End If
                    
                Draw3D .x, .y, 75, 24, 0, 123 + YSrcOffset, TexControl1
                DrawBmpString3D Create3DString(.Caption), .x + .Width \ 2, .y + 2, TextColor, True
        End Select
    End With
End Sub

Public Sub DrawLabel(wdgLabel As clsWidget)
    Dim St As String
    Dim A As Long
    Dim textHeight As Long

    With wdgLabel
        St = Replace(.Caption, "\n", vbCr)
        textHeight = 16
        If (.Style And STYLE_CENTERED) Then
            A = Len(.Caption) * 9
            DrawBmpString3D Create3DString(St), .x + .Width \ 2, .y + ((.Height - 14) / 2), &HFFFFFFFF, False
        Else
            DrawMultilineString3D St, .x, .y, .Width, .Style
        End If
    End With
End Sub

Public Sub DrawTextBox(wdgTextBox As clsWidget)
    Dim WidthLeft As Long, WidthMul As Long
    Dim HeightLeft As Long, HeightMul As Long
    Dim XSrcOffSet As Long
    Dim A As Long, b As Long
    With wdgTextBox
        If .Width > 48 And .Height >= 24 Then
            If .Selected Then
            '    XSrcOffSet = 80
            ElseIf .Highlighted Then
            '    XSrcOffSet = 40
            Else
            '    XSrcOffSet = 0
            End If
            XSrcOffSet = 0
            If (.Style And STYLE_MULTILINE) Then
                'Draw Top Left Corner
                Draw3D .x, .y, 2, 2, XSrcOffSet, 97, TexControl1
                Draw3D .x + .Width - 2, .y, 2, 2, XSrcOffSet + 36, 97, TexControl1
                Draw3D .x, .y + .Height - 2, 2, 2, XSrcOffSet, 121, TexControl1
                Draw3D .x + .Width - 2, .y + .Height - 2, 2, 2, XSrcOffSet + 36, 121, TexControl1
                
                WidthLeft = .Width - 4 '48
                HeightLeft = .Height - 4
    
                WidthMul = Int(WidthLeft / 32)
                HeightMul = Int(HeightLeft / 20)
                
                For A = 0 To WidthMul
                    For b = 0 To HeightMul
                        If A = 0 Then
                            If b < HeightMul Then
                                Draw3D .x, 2 + .y + (20 * b), 2, 20, XSrcOffSet, 100, TexControl1
                            Else
                                Draw3D .x, 2 + .y + (20 * b), 2, HeightLeft Mod 20, 0, 100, TexControl1
                            End If
                        End If
                        If b = 0 Then
                            If A < WidthMul Then
                                Draw3D 2 + .x + (32 * A), .y, 32, 2, XSrcOffSet + 3, 97, TexControl1
                            Else
                                Draw3D 2 + .x + (32 * A), .y, WidthLeft Mod 32, 2, XSrcOffSet + 3, 97, TexControl1
                            End If
                        End If
                        
                        If b < HeightMul Then
                            If A < WidthMul Then
                                Draw3D 2 + .x + A * 32, 2 + .y + b * 20, 32, 20, XSrcOffSet + 3, 100, TexControl1
                            Else
                                Draw3D 2 + .x + A * 32, 2 + .y + b * 20, WidthLeft Mod 32, 20, XSrcOffSet + 3, 100, TexControl1
                            End If
                        Else
                            If A < WidthMul Then
                                'If 2 + .X + A * 32 + 32 > 2 + .X + WidthMul * 32 Then
                                '    Draw3D 2 + .X + A * 32, 2 + .Y + b * 20, (2 + .X + A * 32 + 32) - (2 + .X + WidthMul * 32), HeightLeft Mod 20, XSrcOffSet + 3, 100, TexControl1
                                'Else
                                    Draw3D 2 + .x + A * 32, 2 + .y + b * 20, 32, HeightLeft Mod 20, XSrcOffSet + 3, 100, TexControl1
                                'End If
                            Else
                                Draw3D 2 + .x + A + 32, 2 + .y + b * 20, WidthLeft Mod 32, HeightLeft Mod 20, XSrcOffSet + 3, 100, TexControl1
                            End If
                        End If
                        
                        If b = HeightMul Then
                            If A < WidthMul Then
                                Draw3D 2 + .x + (32 * A), .y + .Height - 2, 32, 2, XSrcOffSet + 3, 121, TexControl1
                            Else
                                Draw3D 2 + .x + (32 * A), .y + .Height - 2, WidthLeft Mod 32, 2, XSrcOffSet + 3, 121, TexControl1
                            End If
                        End If
                        If A = WidthMul Then
                            If b < HeightMul Then
                                Draw3D .x + .Width - 2, 2 + .y + (20 * b), 2, 20, XSrcOffSet + 36, 100, TexControl1
                            Else
                                Draw3D .x + .Width - 2, 2 + .y + (20 * b), 2, HeightLeft Mod 20, XSrcOffSet + 36, 100, TexControl1
                            End If
                        End If
                    Next b
                Next A
                
                
            Else
                'Draw Top Left Corner
                Draw3D .x, .y, 2, 2, XSrcOffSet, 97, TexControl1
                Draw3D .x + .Width - 2, .y, 2, 2, XSrcOffSet + 36, 97, TexControl1
                Draw3D .x, .y + 22, 2, 2, XSrcOffSet, 121, TexControl1
                Draw3D .x + .Width - 2, .y + 22, 2, 2, XSrcOffSet + 36, 121, TexControl1
                
                WidthLeft = .Width - 4 '48
    
                WidthMul = Int(WidthLeft / 32)
                
                Draw3D .x, 2 + .y, 2, 20, XSrcOffSet, 100, TexControl1
                Draw3D .x + .Width - 2, 2 + .y, 2, 20, XSrcOffSet + 36, 100, TexControl1
                For A = 0 To WidthMul
                    If A < WidthMul Then
                        Draw3D 2 + .x + (32 * A), .y, 32, 2, XSrcOffSet + 3, 97, TexControl1
                        Draw3D 2 + .x + 32 * A, .y + 22, 32, 2, XSrcOffSet + 3, 121, TexControl1
                        Draw3D 2 + .x + A * 32, 2 + .y, 32, 20, XSrcOffSet + 3, 100, TexControl1
                    Else
                        Draw3D 2 + .x + (32 * A), .y, WidthLeft Mod 32, 2, XSrcOffSet + 3, 97, TexControl1
                        Draw3D 2 + .x + (32 * A), .y + 22, WidthLeft Mod 32, 2, XSrcOffSet + 3, 121, TexControl1
                        Draw3D 2 + .x + A * 32, 2 + .y, WidthLeft Mod 32, 20, XSrcOffSet + 3, 100, TexControl1
                    End If
                Next A
                
                'Draw Text
                Dim St1 As String, tWidth As Long, CursorWidth As Long
                CursorWidth = FontChar(2, Asc("|")).Width
                If Get3DFontWidth(.Caption) > Int((.Width - 8)) Then
                    For A = Len(.Caption) To 1 Step -1
                        tWidth = Get3DFontWidth(Mid$(.Caption, A, Len(.Caption) - A))
                        If tWidth >= (.Width - 8) Then
                            St1 = Mid$(.Caption, A + 1, Len(.Caption) - A - 1)
                            Exit For
                        End If
                    Next A
                    DrawBmpString3D Create3DString(St1), .x + 4, .y + 4, &HFFFFFFFF, False
                    If CurFrame And .Selected Then DrawBmpString3D Create3DString("|"), .x + tWidth + 4, .y + 4, &HFFFFFFFF 'TODO
                Else
                    tWidth = Get3DFontWidth(.Caption)
                    If Len(.Caption) > 0 Then
                        DrawBmpString3D Create3DString(.Caption), .x + 4, .y + 4, &HFFFFFFFF, False
                        If CurFrame And .Selected Then DrawBmpString3D Create3DString("|"), .x + tWidth + 4, .y + 4, &HFFFFFFFF
                    Else
                        If CurFrame And .Selected Then DrawBmpString3D Create3DString("|"), .x + 4, .y + 4, &HFFFFFFFF
                    End If
                End If
            End If
        End If
    End With
End Sub

Public Sub DrawWidgetImage(wdgImage As clsWidget)
    With wdgImage
        .DrawImage
    End With
End Sub



'------------Mouse/Keyboard Functions

Public Sub WidgetMouseMove(srcWidgets As clsWidgets, x As Long, y As Long)
    Dim tmpWidget As clsWidget
    For Each tmpWidget In srcWidgets
        With tmpWidget
            If .Visible = True Then
                If x >= .x * WindowScaleX And x < .x * WindowScaleX + .Width * WindowScaleX And y >= .y * WindowScaleY And y < .y * WindowScaleY + .Height * WindowScaleY Then
                    Select Case .WidgetType
                        Case WIDGET_BUTTON
                            .Highlighted = True
                        Case WIDGET_TEXTBOX
                            .Highlighted = True
                    End Select
                    If .Children.Count > 0 Then WidgetMouseMove .Children, x, y
                Else
                    .Highlighted = False
                End If
            End If
        End With
    Next
End Sub

Public Sub WidgetMouseDown(srcWidgets As clsWidgets, Button As Long, x As Long, y As Long)
    Dim tmpWidget As clsWidget
    For Each tmpWidget In srcWidgets
        With tmpWidget
            If .Visible = True Then
                If x >= .x * WindowScaleX And x < .x * WindowScaleX + .Width * WindowScaleX And y >= .y * WindowScaleY And y < .y * WindowScaleY + .Height * WindowScaleY Then
                    Select Case .WidgetType
                        Case WIDGET_BUTTON
                            If Button = 1 Then .Selected = True
                        Case WIDGET_TEXTBOX
                            If Button = 1 Then .Selected = True
                    End Select
                    If .Children.Count > 0 Then
                        WidgetMouseDown .Children, Button, x, y
                    End If
                Else
                    .Selected = False
                End If
            End If
            
        End With
    Next
End Sub

Public Sub WidgetMouseUp(srcWidgets As clsWidgets, Button As Long, x As Long, y As Long)
    Dim tmpWidget As clsWidget
    For Each tmpWidget In srcWidgets
        With tmpWidget
            If .Visible = True Then
                Select Case .WidgetType
                    Case WIDGET_BUTTON
                        If x >= .x * WindowScaleX And x < .x * WindowScaleX + .Width * WindowScaleX And y >= .y * WindowScaleY And y < .y * WindowScaleY + .Height * WindowScaleY Then
                            If .Remote And .Selected Then
                                .Selected = False
                                If RemoteWidgetParent <> "" Then
                                    Widgets.item(RemoteWidgetParent).Visible = False
                                    SendRemoteWidgetString (.Key)
                                    UnloadWidgets
                                End If
                            End If
                        Else
                            .Selected = False
                        End If
                    Case WIDGET_TEXTBOX
                End Select
                If .Children.Count > 0 Then WidgetMouseUp .Children, Button, x, y
            End If
            
        End With
    Next
End Sub

Public Function WidgetKeyDown(srcWidgets As clsWidgets, KeyCode As Integer) As Long
    Dim tmpWidget As clsWidget
    Dim RetVal As Long
    RetVal = 1
    For Each tmpWidget In srcWidgets
        With tmpWidget
            If .Visible = True Then
                If .Selected Then
                    Select Case .WidgetType
                        Case WIDGET_BUTTON
                        Case WIDGET_TEXTBOX
                            If KeyCode >= 32 And KeyCode <= 127 Then
                                .Caption = .Caption & Chr$(KeyCode)
                            ElseIf KeyCode = 8 Then
                                If Len(.Caption) > 0 Then
                                    .Caption = Left(.Caption, Len(.Caption) - 1)
                                End If
                            End If
                            RetVal = 0
                    End Select
                End If
                If .Children.Count > 0 Then RetVal = WidgetKeyDown(.Children, KeyCode)
            End If
        End With
    Next
    WidgetKeyDown = RetVal
End Function



'----------------Menu from String Generation

Public Function CreateMenuFromString(St As String)
    Dim tKey As String, CurSt As String
    Dim x As Long, y As Long, Width As Long, Height As Long, WidgetType As Long, Flags As Long
    Dim A As Long, b As Long
    'Is another menu open? If so, kill the fucker
    If RemoteWidgetParent <> "" Then
        UnloadWidgets
    End If
    
    'First make sure its a menu and then parse the data
    A = Asc(Mid$(St, 1, 1))
    CurSt = Mid$(St, 2, A)
    If Len(St) > A + 1 Then
        St = Mid$(St, A + 2)
    Else
        St = ""
    End If
    
    WidgetType = Asc(Mid$(CurSt, 1, 1))
    If WidgetType = 3 Then
        x = GetInt(Mid$(CurSt, 2, 2))
        y = GetInt(Mid$(CurSt, 4, 2))
        Width = GetInt(Mid$(CurSt, 6, 2))
        Height = GetInt(Mid$(CurSt, 8, 2))
        tKey = Mid$(CurSt, 10)
        Widgets.Add tKey, x, y, Width, Height, WIDGET_FRAME, STYLE_NOSTYLE
        RemoteWidgetParent = tKey
    Else
        PrintChat "There has been a WIDGET ERROR!", 15, Options.FontSize
        Exit Function
    End If
    
    Do While Len(St) <> 0
        With Widgets.item(RemoteWidgetParent).Children
            A = Asc(Mid$(St, 1, 1))
            CurSt = Mid$(St, 2, A)
            If Len(St) > A + 1 Then
                St = Mid$(St, A + 2)
            Else
                St = ""
            End If
    
            WidgetType = Asc(Mid$(CurSt, 1, 1))
            Select Case WidgetType
                Case WIDGET_BUTTON
                    x = GetInt(Mid$(CurSt, 2, 2))
                    y = GetInt(Mid$(CurSt, 4, 2))
                    Flags = GetLong(Mid$(CurSt, 6, 4))
                    If (Flags And STYLE_MEDIUM) Then
                        Width = 75
                        Height = 24
                    End If
                    A = Asc(Mid$(CurSt, 10, 1))
                    tKey = Mid$(CurSt, 11, A)
                    .Add tKey, x, y, Width, Height, WidgetType, STYLE_MEDIUM
                    .item(tKey).Remote = True
                    If Len(CurSt) > 11 + A Then
                        .item(tKey).Caption = Mid$(CurSt, 11 + A)
                    End If
                Case WIDGET_IMAGE
                    x = GetInt(Mid$(CurSt, 2, 2))
                    y = GetInt(Mid$(CurSt, 4, 2))
                    Width = GetInt(Mid$(CurSt, 6, 2))
                    Height = GetInt(Mid$(CurSt, 8, 2))
                    Flags = GetLong(Mid$(CurSt, 14, 4))
                    A = Asc(Mid$(CurSt, 18, 1))
                    tKey = Mid$(CurSt, 19, A)
                    .Add tKey, x, y, Width, Height, WidgetType, Flags
                    .item(tKey).Remote = True
                    b = Asc(Mid$(CurSt, 19 + A, 1))
                    .item(tKey).Data0 = GetInt(Mid$(CurSt, 10, 2))
                    .item(tKey).Data1 = GetInt(Mid$(CurSt, 12, 2))
                    If Len(CurSt) > 19 + A Then
                        .item(tKey).Caption = Mid$(CurSt, 20 + A, b)
                    End If
                    .item(tKey).InitTexture
                Case Else
                    x = GetInt(Mid$(CurSt, 2, 2))
                    y = GetInt(Mid$(CurSt, 4, 2))
                    Width = GetInt(Mid$(CurSt, 6, 2))
                    Height = GetInt(Mid$(CurSt, 8, 2))
                    Flags = GetLong(Mid$(CurSt, 10, 4))
                    A = Asc(Mid$(CurSt, 14, 1))
                    tKey = Mid$(CurSt, 15, A)
                    .Add tKey, x, y, Width, Height, WidgetType, Flags
                    .item(tKey).Remote = True
                    If Len(CurSt) > 15 + A Then
                        .item(tKey).Caption = Mid$(CurSt, 15 + A)
                    End If
            End Select
        End With
    Loop
End Function

Sub SendRemoteWidgetString(ActiveWidget As String)
Dim St As String, St1 As String
With Widgets.item(RemoteWidgetParent)
    If .Children.Count > 0 Then
        Dim tmpWidget As clsWidget
        For Each tmpWidget In .Children
            With tmpWidget
                'Loop through each item, and serialize it
                Select Case .WidgetType
                    Case WIDGET_TEXTBOX
                        St1 = Chr$(.WidgetType) + Chr$(Len(.Key)) + .Key + Chr$(Len(.Caption)) + .Caption
                        St = St + Chr$(Len(St1)) + St1
                    Case WIDGET_IMAGE
                    
                    Case Else
                   '     St1 = Chr$(.WidgetType) + Chr$(Len(.Key)) + .Key
                   '     St = St + Chr$(Len(St1)) + St1
                End Select
            End With
        Next
    End If
End With
    SendSocket Chr$(77) + St + ActiveWidget
End Sub

Sub UnloadWidgets()
    Widgets.Clear
    RemoteWidgetParent = ""
End Sub
