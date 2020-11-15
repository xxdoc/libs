

Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByRef lpszName As Any, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Private Declare Sub vbaCopyBytes Lib "msvbvm60.dll" Alias "__vbaCopyBytes" (ByVal ByteCount As Long, ByRef Dest As Any, ByRef Src As Any)

Private Const SND_ASYNC As Long = &H1
Private Const SND_MEMORY As Long = &H4
Private Const SND_LOOP As Long = &H8

Private Type RiffHeader
    ID As Long
    Size As Long
    FileFormat As Long
End Type
Private Type FormatHeader
    ID As Long
    Size As Long
    AudioFormat As Integer
    NumChannels As Integer
    SR As Long
    ByteRate As Long
    BytesPerSample As Integer
    BitsPerChannel As Integer
End Type
Private Type DataHeader
    ID As Long
    Size As Long
End Type

Public Sub PlayWave(ByRef Wave() As Integer, Optional ByVal SR As Long = 48000, Optional ByVal LoopSound As Boolean, Optional ByVal WaitTillPlayFinished As Boolean)
    Dim Flags As Long
    Dim WavFile() As Byte
    Dim RH As RiffHeader
    Dim FH As FormatHeader
    Dim DH As DataHeader
    
    Flags = SND_MEMORY
    If WaitTillPlayFinished = False Then Flags = Flags Or SND_ASYNC
    If LoopSound Then Flags = Flags Or SND_LOOP
    
    
    DH.ID = &H61746164
    DH.Size = (UBound(Wave) + 1) * 2
    
    With FH
        .ID = &H20746D66
        .Size = LenB(FH) - 8
        .AudioFormat = 1
        .NumChannels = 1
        .SR = SR
        .BitsPerChannel = 16
        .BytesPerSample = (.BitsPerChannel \ 8) * .NumChannels
        .ByteRate = .SR * .BytesPerSample
    End With
    RH.ID = &H46464952
    RH.Size = DH.Size + 8 + FH.Size + 8 + 4
    RH.FileFormat = &H45564157

    ReDim WavFile(RH.Size + 8 - 1)
    vbaCopyBytes LenB(RH), WavFile(0), RH
    vbaCopyBytes LenB(FH), WavFile(LenB(RH)), FH
    vbaCopyBytes LenB(DH), WavFile(LenB(RH) + LenB(FH)), DH
    vbaCopyBytes DH.Size, WavFile(LenB(RH) + LenB(FH) + LenB(DH)), Wave(0)
    
    PlaySound WavFile(0), 0, Flags
End Sub

Public Sub StopWave()
    PlaySound ByVal 0&, 0, 0
End Sub



