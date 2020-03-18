VERSION 5.00
Begin VB.Form frmCreateMap 
   Caption         =   "Map Generator"
   ClientHeight    =   8745
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8160
   LinkTopic       =   "Form1"
   ScaleHeight     =   583
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   544
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save Map"
      Height          =   495
      Left            =   5520
      TabIndex        =   9
      Top             =   8160
      Width           =   1215
   End
   Begin VB.PictureBox tMap 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5760
      Left            =   720
      ScaleHeight     =   384
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   384
      TabIndex        =   7
      Top             =   600
      Width           =   5760
   End
   Begin VB.PictureBox T 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   240
      ScaleHeight     =   495
      ScaleWidth      =   1215
      TabIndex        =   6
      Top             =   8160
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox Tm 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   1920
      ScaleHeight     =   495
      ScaleWidth      =   1215
      TabIndex        =   5
      Top             =   8160
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "Generate Map"
      Height          =   495
      Left            =   6840
      TabIndex        =   4
      Top             =   8160
      Width           =   1215
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   7695
      LargeChange     =   5
      Left            =   7800
      Max             =   33
      Min             =   1
      TabIndex        =   2
      Top             =   120
      Value           =   1
      Width           =   255
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      LargeChange     =   5
      Left            =   120
      Max             =   43
      Min             =   1
      TabIndex        =   1
      Top             =   7800
      Value           =   1
      Width           =   7695
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7680
      Left            =   120
      ScaleHeight     =   512
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   512
      TabIndex        =   0
      Top             =   120
      Width           =   7680
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   38400
         Left            =   0
         ScaleHeight     =   2560
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   3200
         TabIndex        =   3
         Top             =   0
         Width           =   48000
      End
   End
   Begin VB.Label lblStatus 
      Caption         =   "Doing Nothing..."
      Height          =   255
      Left            =   3240
      TabIndex        =   8
      Top             =   8160
      Width           =   3015
   End
End
Attribute VB_Name = "frmCreateMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Dim MapL(1 To 50, 1 To 40) As Integer

Private Sub cmdSave_Click()
Call SavePicture(Picture2.Image, App.Path & "/MAP.bmp")
End Sub

Private Sub cmdGenerate_Click()
Dim MapNum As Integer, tMapNum As Integer
lblStatus.Caption = "Creating Lookup Table..."
For Y = 1 To 40
    For X = 1 To 50
        MapNum = Y * 50 + X - 50
        MapL(X, Y) = MapNum
    Next X
Next Y

T.Picture = LoadPicture(App.Path & "\tiles.rsc")
Tm.Picture = LoadPicture(App.Path & "\tilesm.rsc")
tMap.Visible = True
Picture2.ForeColor = vbWhite
For X = 1 To 50
    For Y = 1 To 40
        MapNum = MapL(X, Y)
        MapRS.Seek "=", MapNum
        LoadMap MapRS!Data, MapNum
        tMap.Cls
        lblStatus.Caption = "Drawing Map " & MapNum & "."
        For mX = 0 To 11
            For mY = 0 To 11
                With Map(MapNum).Tile(mX, mY)
                    If .Ground > 0 Then
                        StretchBlt tMap.hdc, mX * 32, mY * 32, 32, 32, T.hdc, ((.Ground - 1) Mod 7) * 32, Int((.Ground - 1) / 7) * 32, 32, 32, vbSrcCopy
                    End If
                    If .Ground2 > 0 Then
                        StretchBlt tMap.hdc, mX * 32, mY * 32, 32, 32, Tm.hdc, ((.Ground2 - 1) Mod 7) * 32, Int((.Ground2 - 1) / 7) * 32, 32, 32, vbSrcAnd
                        StretchBlt tMap.hdc, mX * 32, mY * 32, 32, 32, T.hdc, ((.Ground2 - 1) Mod 7) * 32, Int((.Ground2 - 1) / 7) * 32, 32, 32, vbSrcPaint
                    End If
                    If .BGTile1 > 0 Then
                        StretchBlt tMap.hdc, mX * 32, mY * 32, 32, 32, Tm.hdc, ((.BGTile1 - 1) Mod 7) * 32, Int((.BGTile1 - 1) / 7) * 32, 32, 32, vbSrcAnd
                        StretchBlt tMap.hdc, mX * 32, mY * 32, 32, 32, T.hdc, ((.BGTile1 - 1) Mod 7) * 32, Int((.BGTile1 - 1) / 7) * 32, 32, 32, vbSrcPaint
                    End If
                    If .FGTile > 0 Then
                        StretchBlt tMap.hdc, mX * 32, mY * 32, 32, 32, Tm.hdc, ((.FGTile - 1) Mod 7) * 32, Int((.FGTile - 1) / 7) * 32, 32, 32, vbSrcAnd
                        StretchBlt tMap.hdc, mX * 32, mY * 32, 32, 32, T.hdc, ((.FGTile - 1) Mod 7) * 32, Int((.FGTile - 1) / 7) * 32, 32, 32, vbSrcPaint
                    End If
                End With
            Next mY
        Next mX
        tMap.Refresh
        DoEvents
        tMapNum = MapNum
        While tMapNum > 50
            tMapNum = tMapNum - 50
        Wend
        StretchBlt Picture2.hdc, (X - 1) * 64, (Y - 1) * 64, 64, 64, tMap.hdc, 0, 0, 384, 384, vbSrcCopy
        Dim StrMapNum As String
        StrMapNum = CStr(MapNum)
        TextOut Picture2.hdc, (X - 1) * 64 + 20, (Y - 1) * 64 + 28, StrMapNum, Len(StrMapNum)
    Next Y
Next X
lblStatus = "Drawing GridLines..."
Picture2.ForeColor = vbWhite
For X = 1 To 50
    Picture2.Line (X * 64, 0)-(X * 64, Picture2.Height)
Next X
For Y = 1 To 40
    Picture2.Line (0, Y * 64)-(Picture2.Width, Y * 64)
Next Y
Picture2.Refresh
DoEvents
lblStatus = "Drawing Map Numbers..."
'For x = 1 To 2000
'    'For y = 1 To 40
'        MapNum = x
'
'        DoEvents
'    'Next y
'Next x
tMap.Visible = False
End Sub

Sub LoadMap(MapData As String, MapNum As Integer)
    Dim A As Long, X As Long, Y As Long
    If Len(MapData) = 2219 Then
        With Map(MapNum)
            For Y = 0 To 11
                For X = 0 To 11
                    With .Tile(X, Y)
                        A = 60 + Y * 180 + X * 15
                        .Ground = Asc(Mid$(MapData, A, 1)) * 256 + Asc(Mid$(MapData, A + 1, 1))
                        .Ground2 = Asc(Mid$(MapData, A + 2, 1)) * 256 + Asc(Mid$(MapData, A + 3, 1))
                        .BGTile1 = Asc(Mid$(MapData, A + 4, 1)) * 256 + Asc(Mid$(MapData, A + 5, 1))
                        .BGTile2 = Asc(Mid$(MapData, A + 6, 1)) * 256 + Asc(Mid$(MapData, A + 7, 1))
                        .FGTile = Asc(Mid$(MapData, A + 8, 1)) * 256 + Asc(Mid$(MapData, A + 8, 1))
                    End With
                Next X
            Next Y
        End With
    End If
End Sub

Private Sub HScroll1_Change()
Picture2.Left = 0 - ((HScroll1.Value - 1) * 64)
End Sub

Private Sub HScroll1_Scroll()
Picture2.Left = 0 - ((HScroll1.Value - 1) * 64)
End Sub

Private Sub VScroll1_Change()
Picture2.Top = 0 - ((VScroll1.Value - 1) * 64)
End Sub

Private Sub VScroll1_Scroll()
Picture2.Top = 0 - ((VScroll1.Value - 1) * 64)
End Sub
