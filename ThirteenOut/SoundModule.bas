Attribute VB_Name = "SoundModule"
'***This module is used for playing a *.Wav file***'
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" _
    (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Const SND_ASYNC = &H1

Public Sub PlaySound(WavName As String)
WavName = "/" & WavName & ".wav"
SoundFile$ = App.Path & WavName
wFlags% = SND_ASYNC
x% = sndPlaySound(SoundFile$, wFlags%)
End Sub

