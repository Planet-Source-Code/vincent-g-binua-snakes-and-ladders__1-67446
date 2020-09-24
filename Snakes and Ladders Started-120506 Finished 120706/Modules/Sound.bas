Attribute VB_Name = "Sound"

Public Declare Function PlaySound Lib _
"winmm.dll" Alias "sndPlaySoundA" (ByVal _
lpszSoundName As String, ByVal uFlags As Long) As Long

  
Public Const SND_ASYNC = 1 'play asynchronously, which means that the function returns while the sound is still playing.
Public Const SND_FILENAME = &H20000 'the first argument is a filename.

