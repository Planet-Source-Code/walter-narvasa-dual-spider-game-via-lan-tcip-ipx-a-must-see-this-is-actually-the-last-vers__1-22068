Attribute VB_Name = "modStartup"
'=============================================================================================================================
'
' Developed by Walter A. Narvasa
' jawoltze@edsamail.com.ph
'
' Walter A. Narvasa of
' WANCOM SYSTEMS
'
' READ THIS BEFORE USING THE CODE:
'
' You can study and view the source code for creating your
' own apps, but do not reproduce/release The Spider on the Stick fully
' or partially for any commercial and/or personal purposes. All
' rights of this product is related to it's author. Any violation
' of above conditions will be treated seriously and is punishable.
'
' I do not have full time to add complete explanation, read the help
' file (click Help->Contents) in The Spider on the Stick. Contact me for
' additional help/suggestions
'
'
' VISIT MY WEBSITE : http://jawoltze.gq.nu/
'
'=============================================================================================================================
' Sound Function
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Dim sound As String
Const SND_SYNC = &H0&
Const SND_ASYNC = &H1&
Const SND_NODEFAULT = &H2&
Const SND_LOOP = &H8&
Const SND_NOSTOP = &H10&

Global CurrentGameType As String
Global CurrentGameConnection As String
Global CurrentSpiderType As String
Global CurrentSpiderStage As String
Global StartMultiplayer As Boolean
Global aaa As String
Global aaaa As String

Function NonstopMuzik(Soundfile As String)
    wFlags% = SND_ASYNC Or SND_NODEFAULT Or SND_LOOP
    X% = sndPlaySound(Soundfile$, wFlags%)
End Function

Function StopMuzik()
    wFlags% = SND_ASYNC Or SND_NODEFAULT
    X% = sndPlaySound(Soundfile$, wFlags%)
End Function

Sub Main()
    On Error GoTo HandleErrors
    frmSplashScreen.Platform = pcstrAppPlatform
    frmSplashScreen.Show
    'Ensure the Splash form is refreshed prior to displaying the Main form.
    DoEvents
    '---------------------------------------------------------------------------------------------------------------------
    'Perform other start up tasks here...
    'For demo purposes we add a delay to simulate a typical applications initialisation.
    Call SplashDelay
  '---------------------------------------------------------------------------------------------------------------------
    frmFighter.Show
    DoEvents
    Unload frmSplashScreen
ExitHandleErrors:
  Exit Sub
HandleErrors:
  MsgBox Err.Description & " (" & Err.Number & ")", vbCritical, App.Title & " Error"
  Resume ExitHandleErrors
End Sub

Public Sub SplashDelay()
    On Error Resume Next
    Dim sngStartTime As Single
    sngStartTime = Timer
    Do Until (Timer - sngStartTime) > 4
          DoEvents
    Loop
End Sub
