VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSound"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Filename As String

Public Key As String
Public stream As Long
Public Channel As Long

Public Sub Play(Filename As String)
    Dim tStream As Long
    Dim tStreamChannel As Long
    Dim tPtr As Long
    
    If Options.Wav Then
        If Exists(Filename) Then
            stream = FSOUND_Stream_Open(Filename, fsound_normal, 0, 0)
            Channel = FSOUND_Stream_Play(FSOUND_FREE, stream)
            FSOUND_SetVolume tStreamChannel, Options.SoundVolume
            'tPtr = GetValue(AddressOf EndCallBack)
            FSOUND_Stream_SetEndCallback stream, AddressOf EndCallBack, 0
        End If
    End If
End Sub



