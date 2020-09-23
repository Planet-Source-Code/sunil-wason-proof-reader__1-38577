VERSION 5.00
Object = "{F5BE8BC2-7DE6-11D0-91FE-00C04FD701A5}#2.0#0"; "AGENTCTL.DLL"
Object = "{2398E321-5C6E-11D1-8C65-0060081841DE}#1.0#0"; "Vtext.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmTxtToSpeech 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Proof Reader"
   ClientHeight    =   6585
   ClientLeft      =   1635
   ClientTop       =   1680
   ClientWidth     =   7635
   Icon            =   "Text2Speech.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   7635
   Begin VB.CommandButton cmdHelp 
      Height          =   615
      Left            =   6960
      Picture         =   "Text2Speech.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   51
      ToolTipText     =   " Credits "
      Top             =   2000
      Width           =   615
   End
   Begin VB.Frame Frame4 
      Height          =   615
      Left            =   120
      TabIndex        =   49
      Top             =   2000
      Width           =   6735
      Begin VB.CommandButton cmdClose 
         Caption         =   "Close"
         Height          =   375
         Left            =   120
         TabIndex        =   50
         ToolTipText     =   " Close Application "
         Top             =   160
         Width           =   6495
      End
   End
   Begin VB.CommandButton cmdDefault 
      Caption         =   "Default Settings"
      Height          =   315
      Left            =   3960
      TabIndex        =   24
      ToolTipText     =   " Sets mouth shape's default settings "
      Top             =   6000
      Width           =   1695
   End
   Begin VB.HScrollBar hsMouthWidth 
      Height          =   255
      LargeChange     =   25
      Left            =   3960
      Max             =   255
      TabIndex        =   23
      Top             =   5640
      Width           =   1695
   End
   Begin VB.HScrollBar hsMouthHeight 
      Height          =   255
      LargeChange     =   25
      Left            =   3960
      Max             =   255
      TabIndex        =   22
      Top             =   5040
      Width           =   1695
   End
   Begin VB.HScrollBar hsMouthUpturn 
      Height          =   255
      LargeChange     =   25
      Left            =   3960
      Max             =   255
      TabIndex        =   21
      Top             =   4440
      Width           =   1695
   End
   Begin VB.HScrollBar hsTonguePosn 
      Height          =   255
      LargeChange     =   25
      Left            =   3960
      Max             =   255
      TabIndex        =   20
      Top             =   3840
      Width           =   1695
   End
   Begin VB.HScrollBar hsLipTension 
      Height          =   255
      LargeChange     =   25
      Left            =   3960
      Max             =   255
      TabIndex        =   19
      Top             =   3240
      Width           =   1695
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Read Word Document"
      Height          =   375
      Left            =   120
      TabIndex        =   18
      ToolTipText     =   " Reads Selected Word Document "
      Top             =   1080
      Width           =   2535
   End
   Begin VB.CommandButton cmdReadText 
      Caption         =   "Read Text from Clipboard"
      Height          =   375
      Left            =   120
      TabIndex        =   17
      ToolTipText     =   " Reads copied text "
      Top             =   600
      Width           =   2535
   End
   Begin VB.CheckBox chkCloseDoc 
      Caption         =   "Close document after reading"
      Height          =   480
      Left            =   6000
      TabIndex        =   15
      ToolTipText     =   " Closes the word document after it has been read "
      Top             =   1410
      Width           =   1455
   End
   Begin VB.CheckBox chkSpkMin 
      Caption         =   "Minimise while reading word document"
      Height          =   615
      Left            =   6000
      TabIndex        =   14
      ToolTipText     =   " Minimises this window while reading the word document "
      Top             =   780
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3240
      Top             =   5400
   End
   Begin VB.CommandButton cmdRewind 
      Caption         =   "Rewind"
      Height          =   315
      Left            =   120
      TabIndex        =   13
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Height          =   315
      Left            =   2520
      TabIndex        =   12
      ToolTipText     =   " Stop speaking "
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton cmdPause 
      Caption         =   "Pause"
      Height          =   315
      Left            =   1320
      TabIndex        =   10
      ToolTipText     =   " Pause speaking "
      Top             =   3360
      Width           =   1095
   End
   Begin HTTSLibCtl.TextToSpeech TextToSpeech1 
      Height          =   2895
      Left            =   120
      OleObjectBlob   =   "Text2Speech.frx":074C
      TabIndex        =   9
      Top             =   3600
      Width           =   3615
   End
   Begin VB.ComboBox cboMode 
      Height          =   315
      Left            =   120
      TabIndex        =   8
      ToolTipText     =   " Select the voice "
      Top             =   3000
      Width           =   3495
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select Character"
      Height          =   1815
      Left            =   2760
      TabIndex        =   3
      Top             =   120
      Width           =   3015
      Begin VB.OptionButton optGenie 
         Caption         =   "Santa"
         Height          =   195
         Index           =   14
         Left            =   120
         TabIndex        =   48
         Top             =   1200
         Width           =   855
      End
      Begin VB.OptionButton optGenie 
         Caption         =   "James"
         Height          =   195
         Index           =   13
         Left            =   1200
         TabIndex        =   47
         Top             =   1200
         Width           =   855
      End
      Begin VB.CheckBox chkDispBaloon 
         Caption         =   "Display Baloon"
         Height          =   255
         Left            =   120
         TabIndex        =   46
         ToolTipText     =   " Displays / Hides the word baloon "
         Top             =   1440
         Width           =   1455
      End
      Begin VB.CheckBox chkSizeToText 
         Caption         =   "Size to Text"
         Height          =   255
         Left            =   1680
         TabIndex        =   45
         ToolTipText     =   " Check it to size the baloon so that all the text can be accomodated in it "
         Top             =   1440
         Width           =   1215
      End
      Begin VB.OptionButton optGenie 
         Caption         =   "Totem"
         Height          =   195
         Index           =   12
         Left            =   2040
         TabIndex        =   44
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton optGenie 
         Caption         =   "Electra"
         Height          =   195
         Index           =   11
         Left            =   2040
         TabIndex        =   43
         Top             =   480
         Width           =   855
      End
      Begin VB.OptionButton optGenie 
         Caption         =   "Gar"
         Height          =   195
         Index           =   10
         Left            =   2040
         TabIndex        =   42
         Top             =   720
         Width           =   855
      End
      Begin VB.OptionButton optGenie 
         Caption         =   "Max"
         Height          =   195
         Index           =   9
         Left            =   2040
         TabIndex        =   41
         Top             =   960
         Width           =   855
      End
      Begin VB.OptionButton optGenie 
         Caption         =   "Miku"
         Height          =   195
         Index           =   8
         Left            =   120
         TabIndex        =   40
         Top             =   720
         Width           =   855
      End
      Begin VB.OptionButton optGenie 
         Caption         =   "Oscar"
         Height          =   195
         Index           =   7
         Left            =   1200
         TabIndex        =   39
         Top             =   480
         Width           =   855
      End
      Begin VB.OptionButton optGenie 
         Caption         =   "Spaceman"
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   38
         Top             =   960
         Width           =   1095
      End
      Begin VB.OptionButton optGenie 
         Caption         =   "Joe"
         Height          =   255
         Index           =   5
         Left            =   1200
         TabIndex        =   32
         Top             =   720
         Width           =   855
      End
      Begin VB.OptionButton optGenie 
         Caption         =   "Al"
         Height          =   195
         Index           =   4
         Left            =   1200
         TabIndex        =   31
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton optGenie 
         Caption         =   "Mouth"
         Height          =   255
         Index           =   3
         Left            =   2040
         TabIndex        =   7
         Top             =   1200
         Width           =   855
      End
      Begin VB.OptionButton optGenie 
         Caption         =   "Robby"
         Height          =   255
         Index           =   2
         Left            =   1200
         TabIndex        =   6
         Top             =   960
         Width           =   855
      End
      Begin VB.OptionButton optGenie 
         Caption         =   "Merlin"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   855
      End
      Begin VB.OptionButton optGenie 
         Caption         =   "Genie"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   855
      End
   End
   Begin MSComDlg.CommonDialog cdlOpen 
      Left            =   3120
      Top             =   3840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CheckBox chkDontShow 
      Caption         =   "Hide word document"
      Height          =   375
      Left            =   6000
      TabIndex        =   2
      ToolTipText     =   " Hide the word document while the document is being read "
      Top             =   360
      Width           =   1455
   End
   Begin VB.CommandButton cmdInterrupt 
      Caption         =   "Hide Character"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   " Hides the genie "
      Top             =   1560
      Width           =   2535
   End
   Begin VB.CommandButton CharacterIntro 
      Caption         =   "Character Intro"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   " Selected character introduces itself "
      Top             =   120
      Width           =   2535
   End
   Begin VB.Frame Frame2 
      Caption         =   "Read Word Doc"
      Height          =   1815
      Left            =   5880
      TabIndex        =   16
      Top             =   105
      Width           =   1695
   End
   Begin VB.Frame Frame3 
      Caption         =   "Mouth Shape"
      Height          =   3735
      Left            =   3840
      TabIndex        =   25
      Top             =   2760
      Width           =   1935
      Begin VB.Label lblMouthWidth 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1320
         TabIndex        =   37
         Top             =   2600
         Width           =   495
      End
      Begin VB.Label lblMouthHeight 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1320
         TabIndex        =   36
         Top             =   2000
         Width           =   495
      End
      Begin VB.Label lblMouthUpturn 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1320
         TabIndex        =   35
         Top             =   1400
         Width           =   495
      End
      Begin VB.Label lblTonguePosn 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1320
         TabIndex        =   34
         Top             =   800
         Width           =   495
      End
      Begin VB.Label lblLipTension 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1320
         TabIndex        =   33
         Top             =   195
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Tongue Position"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Mouth Upturn"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "Mouth Width"
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   2640
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Mouth Height"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Lip Tension"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Width           =   1455
      End
   End
   Begin AgentObjectsCtl.Agent AgentX 
      Left            =   120
      Top             =   3840
      _cx             =   847
      _cy             =   847
   End
   Begin VB.Label Label1 
      Caption         =   "Select Speak Engine"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   2760
      Width           =   3495
   End
End
Attribute VB_Name = "frmTxtToSpeech"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit
' Name : Proof Reader
' Author : Sunil Wason (sunilwason@yahoo.com)
' Purpose : This application allows you to
'proof read any MS Office Word document
'(supported till MS Word 2000) or from the
'clipboard (by copying text contents of any
'word/text file). This application allows the
'user to choose an assistant/character from a
'library of 15 characters & including an
'animated mouth to read read the selected
'document/text. The program remembers the last
'MS Agent character used by the user and
'automatically loads the last used character
'when opening the application. It also
'remembers the shape of the mouth which can
'been configured by the user.
'This application also shows how to hide or
'display the balloon and/or size the text in it
'while the MS Agent reads the document/text.
'In addition to the use of MS Agents, this
'program also shows the use of VBA to
'communicate with MS Word documents.

'You are required to install L & H TruVoice
'Engine, Microsoft Agent 2.0 and
'Microsoft Text to Speech engines before using
'the Proof Reader application.
'These can be downloaded from the Microsoft
'site (http://msdn.microsoft.com/workshop/imedia/agent)

'If your application doesnot work go to
'Project --> References and check
'Microsoft Word x.x Object Library
'Microsoft Voice Text
'Voice Text x.x Type Library
'Microsoft Agent Control 2.0.

'This application supports the undermentioned characters which can be downloaded from
'http://www.msagentring.org/
'For the sake of convenience, I have included two charactes Joe & Miku
'The other characters supported are:-
'Genie, Al, Totem, Merlin, Oscar, Electra, Gar,
'Spaceman, Robby, Max, Santa & James and a mouth
'whose shape can changed according to your
'requirement.
'If you download these characters, copy them
'to App.Path

'Please Note:-
'This code is copyrighted.
'This software is provided 'as-is', without any
'express or implied warranty.  In no event
'will the author be held liable for any damages
'arising from the use of this software. You
'have been granted the ability to view and
'improve code and will be acknowledged in the
'credits as a tester.

'This software can be distributed with the
'following restriction:
'1. The origin of this software must not be
'misrepresented; you must not claim that you
'wrote the original software.
'2. Altered source versions must be plainly
'marked as such, and must not be misrepresented
'as being the original software. Notification
'must be sent to the author notifying us of
'such changes.
'3. You may not take this project on as your
'own and/or attempt to complete it.
'4. This notice may not be removed or altered
'from any source distribution.

'-------
'--Declare global variables
'-------
Const SizeToText = 2 'Text is sized to the baloon if <> 2
Const BalloonOn = 1 'or = 4 both Hide the baloon
Const Txt2SpchCfgFile = "C:\Program Files\Txt2Spch.cfg" 'Config file
Dim CharacReadDoc As Boolean
Dim engine
Dim Genie As IAgentCtlCharacterEx
Dim Status
Dim GeniePathName
Dim LoadRequest
Dim LoadShow
Dim LoadMove
Dim ClipboardText
Dim SpeakText1
Dim SpeakText2
Dim SpeakText3
Dim SpeakText4
Dim Speak1Request
Dim Speak2Request
Dim Speak3Request
Dim AppWord As Word.Application
Dim FilePath As String
Dim FileName As String
Dim ModeName As String
Dim Charac As String 'Name of the character
Dim CharacSelected As String 'Character selected by the user

Private Sub cboMode_Click()

engine = TextToSpeech1.Find("Mfg=Microsoft;Gender=1")
'Now Select the engine, SAPI style. This is synonymous with
'doing TextToSpeech1.CurrentMode = engine
'TextToSpeech1.Select engine
Rem Each time somebody selects a new voice/engine/mode from the combo box,
Rem select that voice as the active speaker.
    On Error GoTo ExitRoutine
ContinueSpeak:
    TextToSpeech1.CurrentMode = cboMode.ListIndex + 1
Rem Set the gender of the lips..Gender=1 means female.
    If (TextToSpeech1.Gender(TextToSpeech1.CurrentMode) = 1) Then
        TextToSpeech1.LipType = 0    'female full red lips
    Else
        TextToSpeech1.LipType = 1    'male thinner paler lips
    End If
'Rem speak the text in the text box when the button is pressed
'TextToSpeech1.Speak Text1.Text
Exit Sub
ExitRoutine:
    MsgBox "Your system doesnot support this speech engine." & vbCrLf & vbCrLf _
    & "Please select another character/voice."
    cboMode.ListIndex = 0
    GoTo ContinueSpeak
    
End Sub 'cboMode_Click()

Private Sub chkDispBaloon_Click()

CheckBaloonOn

End Sub 'chkDispBaloon_Click()

Private Sub chkSizeToText_Click()

CheckSizeToText

End Sub 'chkSizeToText_Click()

Private Sub cmdClose_Click()

Unload Me
End

End Sub 'cmdClose_Click()

Private Sub cmdDefault_Click()

hsTonguePosn.Value = 25
hsMouthUpturn.Value = 143
hsLipTension.Value = 164
hsMouthHeight.Value = 255
hsMouthWidth.Value = 255

End Sub 'cmdDefault_Click()

Private Sub cmdHelp_Click()

ShowHalfForm
frmAbout.Show 1

End Sub 'cmdHelp_Click()

Private Sub cmdOpen_Click()

Dim copytext As Variant
CheckSizeToText
CheckBaloonOn
AskUserFileName
On Error Resume Next
'If a Word document is open, then, close it
AppWord.ActiveDocument.Close
AppWord.Quit
If Charac <> "Mouth" Then
    CharacReadDoc = True
End If
If FilePath = "" Then
    Exit Sub
End If
Set AppWord = GetObject("Word.Application")
If AppWord Is Nothing Or AppWord Is Empty Then
    Set AppWord = CreateObject("Word.Application")
    If AppWord Is Nothing Then
        MsgBox "Could not open Word"
        Exit Sub
    End If
End If
AppWord.Visible = True
AppWord.Documents.Open (FilePath)
AppWord.ActiveDocument.Select
Clipboard.Clear
'copy the selected text into a variable
copytext = AppWord.Selection
Clipboard.SetText copytext
If chkDontShow.Value = 1 Then
    AppWord.ActiveDocument.Close
    AppWord.Quit
End If
If Charac <> "Mouth" Then
    If chkSpkMin.Value = 1 Then
        Me.WindowState = 1
    End If
End If
If optGenie(3).Value = True Then 'If mouth has been opted by the user
    TextToSpeech1.Speak copytext
    If chkSpkMin.Value = 1 Then
        Me.WindowState = 1
    End If
    Exit Sub
End If
CharacterRead

End Sub 'cmdOpen_Click()

Private Sub cmdPause_Click()

If cmdPause.Caption = "Pause" Then
    cmdPause.Caption = "Resume"
    TextToSpeech1.Pause
Else
    cmdPause.Caption = "Pause"
    TextToSpeech1.Resume
End If

End Sub 'cmdPause_Click()

Private Sub CharacterIntro_Click()

CheckBaloonOn
On Error Resume Next
Genie.MoveTo 10, 60
DoIntro
If CharacSelected <> "Mouth" Then
    HideCharacter
End If
    
End Sub 'CharacterIntro_Click()

Sub DoIntro1()

If Clipboard.GetFormat(vbCFText) Then
    ClipboardText = Clipboard.GetText
    Set Speak1Request = Genie.Speak(ClipboardText)
End If

End Sub 'DoIntro1()

Private Sub HideCharacter()

Genie.MoveTo 40, 60
Genie.Hide
    
End Sub 'HideCharacter()

Private Sub DoIntro()

Dim i As Integer

Select Case CharacSelected
    Case Is = "Genie"
        Charac = "Genie"
        SpeakIntro
    Case Is = "Merlin"
        Charac = "Merlin"
        SpeakIntro
    Case Is = "Robby"
        Charac = "Robby"
        SpeakIntro
    Case Is = "Al"
        Charac = "Al"
        SpeakIntro
    Case Is = "Joe"
        Charac = "Joe"
        SpeakIntro
    Case Is = "Miku"
        Charac = "Miku"
        SpeakIntro
    Case Is = "Spaceman"
        Charac = "Spaceman"
        SpeakIntro
    Case Is = "Max"
        Charac = "Max"
        SpeakIntro
    Case Is = "Gar"
        Charac = "Gar"
        SpeakIntro
    Case Is = "Totem"
        Charac = "Totem"
        SpeakIntro
    Case Is = "Electra"
        Charac = "Electra"
        SpeakIntro
    Case Is = "Oscar"
        Charac = "Oscar"
        SpeakIntro
    Case Is = "Santa"
        Charac = "Santa"
        SpeakIntro
    Case Is = "James"
        Charac = "James"
        SpeakIntro
    Case Is = "Mouth"
        Charac = "Mouth"
        MouthIntro
End Select
        
End Sub 'DoIntro()

Private Sub MouthIntro()

Dim IntroText As String

IntroText = "Hello, I am your Mr Mouth software agent." _
& "You can copy text and click the button Read Text from Clipboard." _
& "You may also select a Word document by clicking the 'Read Word Document' button and I shall read the entire document for you." _
& "Go ahead and try me out."
TextToSpeech1.Speak IntroText

End Sub 'MouthIntro()

Private Sub SpeakIntro()

SpeakText1 = "Hello, I am " & Charac _
& " , your software agent. "
SpeakText2 = "You can copy text and click the button 'Read Text from Clipboard' for me to read the copied text. "
SpeakText3 = "You may also select a Word document by clicking the 'Read Word Document' button and I shall read the entire document for you."
SpeakText4 = "Go ahead and try me out."
    
Set Speak1Request = Genie.Speak(SpeakText1 + SpeakText2 + SpeakText3 + SpeakText4)
    
End Sub 'SpeakIntro()

Private Sub cmdInterrupt_Click()

On Error GoTo HideGenie
Genie.StopAll "Move, Play, Speak"   ' button click stops character from speaking and flushes queue

    '----------
    '--If the character or initial animations aren't loaded yet
    '--then exit this subroutine
    '----------

    If LoadRequest.Status = 4 Or LoadShow.Status = 4 Then
        Status = "Still loading...please wait to click the button."
        Exit Sub

    ElseIf Speak1Request.Status = 4 Or Speak1Request.Status = 2 Then
        Genie.Speak "Hey, I am not done speaking yet.|Please wait until I have finished talking.|Don't click that button yet.|You interrupted me.|Please let me finish my intro.|I'll just have to start over."
    End If
    
    '----------
    '-- Move the character
    '----------
    
    Genie.MoveTo Int(Rnd * 600), Int(Rnd * 400)
    Set Speak2Request = Genie.Speak("I love to fly.|This is fun!|Aren't I amazing?")
    '----------
    '-- Hide the character
    '----------
HideGenie:
    Genie.Hide
    
End Sub 'cmdInterrupt_Click()

Private Sub CharacterRead()

    Genie.Show
    DoIntro1
    HideCharacter
    
End Sub 'CharacterRead()

Private Sub cmdReadText_Click()

CheckSizeToText
CheckBaloonOn
If optGenie(3).Value = True Then 'i.e. mouth only
    If Clipboard.GetFormat(vbCFText) Then
        ClipboardText = Clipboard.GetText
    Else
        Exit Sub
    End If
        TextToSpeech1.Speak ClipboardText
    Exit Sub
End If
DoIntro1 'for characters other than mouth

End Sub 'cmdReadText_Click()

Private Sub CheckSizeToText()

'If the user has checked the chk box for SizeToText,
'then, size the text in the baloon otherwise,
'place only four lines of text in the baloon.
On Error Resume Next
If chkSizeToText.Value = 0 Then
    Genie.Balloon.Style = Genie.Balloon.Style And (Not SizeToText)
ElseIf chkSizeToText.Value = 1 Then
    Genie.Balloon.Style = Genie.Balloon.Style Or SizeToText
End If
On Error GoTo 0

End Sub 'checkSizeToText()

Private Sub CheckBaloonOn()

'If the user has checked the chk box for SizeToText,
'then, size the text in the baloon otherwise,
'place only four lines of text in the baloon.
On Error Resume Next
If chkDispBaloon.Value = 0 Then
    Genie.Balloon.Style = Genie.Balloon.Style And (Not BalloonOn)
ElseIf chkDispBaloon.Value = 1 Then
    Genie.Balloon.Style = Genie.Balloon.Style Or BalloonOn
End If
On Error GoTo 0

End Sub 'CheckBaloonOn()

Private Sub cmdRewind_Click()

TextToSpeech1.Rewind

End Sub 'cmdRewind_Click()

Private Sub cmdStop_Click()

cmdPause.Caption = "Pause"
TextToSpeech1.StopSpeaking

End Sub 'cmdStop_Click()

Private Sub Form_Activate()

AlwaysOnTop frmTxtToSpeech, True

End Sub 'Form_Activate()

Private Sub Form_Load()

Dim i As Byte
Open Txt2SpchCfgFile For Random As #1 Len = Len(CfgFile)
NumRec = LOF(1) / Len(CfgFile)
Close #1
If NumRec = 0 Then
    'Create the config file
    FillInitValues
    optGenie(5).Value = True
End If
'Check from the config file which was the last
'charac opened by the user and open it now
'again.
On Error Resume Next
Open Txt2SpchCfgFile For Random As #1 Len = Len(CfgFile)
NumRec = LOF(1) / Len(CfgFile)
Get #1, NumRec, CfgFile
Close #1
CharacOpen = CfgFile.Charac
LipTension = CfgFile.LipTension
MouthHeight = CfgFile.MouthHeight
MouthUpturn = CfgFile.MouthUpturn
MouthWidth = CfgFile.MouthWidth
TonguePosn = CfgFile.TonguePosn
optGenie(CharacOpen).Value = True
For i = 1 To TextToSpeech1.CountEngines
    ModeName = TextToSpeech1.ModeName(i)
    cboMode.AddItem ModeName
Next i
chkDispBaloon.Value = 1
chkSizeToText.Value = 1
chkSizeToText.Value = 0

ShowHalfForm

End Sub 'Form_Load()

Private Sub FillInitValues()
Dim a As Byte
'Create the config file
Open Txt2SpchCfgFile For Random As #1 Len = Len(CfgFile)
NumRec = LOF(1) / Len(CfgFile)
With CfgFile
    .Charac = 0 'Character index selected
    .LipTension = 164
    .MouthHeight = 255
    .MouthUpturn = 143
    .TonguePosn = 25
    .MouthWidth = 255
    .RecNum = 1
End With
NumRec = NumRec + 1
Put #1, NumRec, CfgFile
Close #1

End Sub 'FillInitValues()

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

On Error Resume Next
AppWord.ActiveDocument.Close
AppWord.Quit
On Error GoTo 0
Open Txt2SpchCfgFile For Random As #1 Len = Len(CfgFile)
NumRec = LOF(1) / Len(CfgFile)
With CfgFile
    .Charac = CharacOpen
    .LipTension = LipTension
    .TonguePosn = TonguePosn
    .MouthUpturn = MouthUpturn
    .MouthHeight = MouthHeight
    .MouthWidth = MouthWidth
End With
Put #1, NumRec, CfgFile
Close #1

End Sub 'Form_QueryUnload

Private Sub AskUserFileName()
' This procedure takes control when the user clicks the
'Save As command, and gives the user the opportunity to
'create a new file.
 
Dim Answer As Integer

cdlOpen.Filter = "Word Files (*.doc)|*.doc"
cdlOpen.DialogTitle = "Open"
'If the user clicks Cancel on the Save dialog box, the
'ShowSave method generates an error. In this case, exit
'from this procedure.
On Error GoTo userCancel
  cdlOpen.ShowOpen
On Error GoTo 0
' Record the file name that the user has entered.
FilePath = cdlOpen.FileName
'Terminate the procedure if the user has clicked the
'Cancel button.
userCancel:

End Sub 'AskUserFileName

Private Sub hsLipTension_Change()

TextToSpeech1.LipTension = hsLipTension.Value
lblLipTension.Caption = hsLipTension.Value
LipTension = hsLipTension.Value

End Sub 'hsLipTension_Change()

Private Sub hsMouthHeight_Change()

TextToSpeech1.MouthHeight = hsMouthHeight.Value
lblMouthHeight.Caption = hsMouthHeight.Value
MouthHeight = hsMouthHeight.Value

End Sub 'hsMouthHeight_Change()

Private Sub hsMouthUpturn_Change()

TextToSpeech1.MouthUpturn = hsMouthUpturn.Value
lblMouthUpturn.Caption = hsMouthUpturn.Value
MouthUpturn = hsMouthUpturn.Value

End Sub 'hsMouthUpturn_Change()

Private Sub hsMouthWidth_Change()

TextToSpeech1.MouthWidth = hsMouthWidth.Value
lblMouthWidth.Caption = hsMouthWidth.Value
MouthWidth = hsMouthWidth.Value

End Sub 'hsMouthWidth_Change()

Private Sub hsTonguePosn_Change()

TextToSpeech1.TonguePosn = hsTonguePosn.Value
lblTonguePosn.Caption = hsTonguePosn.Value
TonguePosn = hsTonguePosn.Value

End Sub 'hsTonguePosn_Change()

Private Sub optGenie_Click(Index As Integer)

CharacSelected = optGenie(Index).Caption
If Index <> 3 Then 'If the mouth has not been selected then
    'Store the index No of the character which
    'has been opened / selected by the user.
    CharacOpen = Index
End If
On Error Resume Next
    'Stop all actions of the character
    Genie.StopAll "Move, Play, Speak"
    'Hide or show the character
    'Genie.Hide = Not Genie.Hide
    Genie.Hide
    
    Select Case CharacSelected
    Case Is = "Genie"
        TextToSpeech1.StopSpeaking
        ShowHalfForm
        cmdOpen.Enabled = True
        GeniePathName = App.Path & "\genie.acs"
        AgentX.Connected = True
        Charac = "Genie"
        Set LoadRequest = AgentX.Characters.Load("genie", GeniePathName)
        Set Genie = AgentX.Characters("genie")
        Set LoadShow = Genie.Get("state", "Showing, Speaking")
        Genie.MoveTo 600, 200
        Genie.Show
    Case Is = "Merlin"
        TextToSpeech1.StopSpeaking
        ShowHalfForm
        cmdOpen.Enabled = True
        GeniePathName = App.Path & "\merlin.acs"
        AgentX.Connected = True
        Charac = "Merlin"
        Set LoadRequest = AgentX.Characters.Load("merlin", GeniePathName)
        Set Genie = AgentX.Characters("merlin")
        Set LoadShow = Genie.Get("state", "Showing, Speaking")
        Genie.MoveTo 600, 200
        Genie.Show
    Case Is = "Robby"
        TextToSpeech1.StopSpeaking
        ShowHalfForm
        cmdOpen.Enabled = True
        GeniePathName = App.Path & "\robby.acs"
        AgentX.Connected = True
        Charac = "Robby"
        Set LoadRequest = AgentX.Characters.Load("robby", GeniePathName)
        Set Genie = AgentX.Characters("robby")
        Set LoadShow = Genie.Get("state", "Showing, Speaking")
        Genie.MoveTo 600, 200
        Genie.Show
    Case Is = "Al"
        TextToSpeech1.StopSpeaking
        ShowHalfForm
        cmdOpen.Enabled = True
        GeniePathName = App.Path & "\Al.acs"
        AgentX.Connected = True
        Charac = "Al"
        Set LoadRequest = AgentX.Characters.Load("Al", GeniePathName)
        Set Genie = AgentX.Characters("Al")
        Set LoadShow = Genie.Get("state", "Showing, Speaking")
        Genie.MoveTo 600, 200
        Genie.Show
    Case Is = "Joe"
        TextToSpeech1.StopSpeaking
        ShowHalfForm
        cmdOpen.Enabled = True
        GeniePathName = App.Path & "\Joe.acs"
        AgentX.Connected = True
        Charac = "Joe"
        Set LoadRequest = AgentX.Characters.Load("Joe", GeniePathName)
        Set Genie = AgentX.Characters("Joe")
        Set LoadShow = Genie.Get("state", "Showing, Speaking")
        Genie.MoveTo 600, 200
        Genie.Show
    Case Is = "Spaceman"
        TextToSpeech1.StopSpeaking
        ShowHalfForm
        cmdOpen.Enabled = True
        GeniePathName = App.Path & "\Spaceman.acs"
        AgentX.Connected = True
        Charac = "Spaceman"
        Set LoadRequest = AgentX.Characters.Load("Spaceman", GeniePathName)
        Set Genie = AgentX.Characters("Spaceman")
        Set LoadShow = Genie.Get("state", "Showing, Speaking")
        Genie.MoveTo 600, 200
        Genie.Show
    Case Is = "Oscar"
        TextToSpeech1.StopSpeaking
        ShowHalfForm
        cmdOpen.Enabled = True
        GeniePathName = App.Path & "\Oscar.acs"
        AgentX.Connected = True
        Charac = "Oscar"
        Set LoadRequest = AgentX.Characters.Load("Oscar", GeniePathName)
        Set Genie = AgentX.Characters("Oscar")
        Set LoadShow = Genie.Get("state", "Showing, Speaking")
        Genie.MoveTo 600, 200
        Genie.Show
    Case Is = "Miku"
        TextToSpeech1.StopSpeaking
        ShowHalfForm
        cmdOpen.Enabled = True
        GeniePathName = App.Path & "\Miku.acs"
        AgentX.Connected = True
        Charac = "Miku"
        Set LoadRequest = AgentX.Characters.Load("Miku", GeniePathName)
        Set Genie = AgentX.Characters("Miku")
        Set LoadShow = Genie.Get("state", "Showing, Speaking")
        Genie.MoveTo 600, 200
        Genie.Show
    Case Is = "Max"
        TextToSpeech1.StopSpeaking
        ShowHalfForm
        cmdOpen.Enabled = True
        GeniePathName = App.Path & "\Max.acs"
        AgentX.Connected = True
        Charac = "Max"
        Set LoadRequest = AgentX.Characters.Load("Max", GeniePathName)
        Set Genie = AgentX.Characters("Max")
        Set LoadShow = Genie.Get("state", "Showing, Speaking")
        Genie.MoveTo 600, 200
        Genie.Show
    Case Is = "Electra"
        TextToSpeech1.StopSpeaking
        ShowHalfForm
        cmdOpen.Enabled = True
        GeniePathName = App.Path & "\Electra.acs"
        AgentX.Connected = True
        Charac = "Electra"
        Set LoadRequest = AgentX.Characters.Load("Electra", GeniePathName)
        Set Genie = AgentX.Characters("Electra")
        Set LoadShow = Genie.Get("state", "Showing, Speaking")
        Genie.MoveTo 600, 200
        Genie.Show
    Case Is = "Gar"
        TextToSpeech1.StopSpeaking
        ShowHalfForm
        cmdOpen.Enabled = True
        GeniePathName = App.Path & "\Gar.acs"
        AgentX.Connected = True
        Charac = "Gar"
        Set LoadRequest = AgentX.Characters.Load("Gar", GeniePathName)
        Set Genie = AgentX.Characters("Gar")
        Set LoadShow = Genie.Get("state", "Showing, Speaking")
        Genie.MoveTo 600, 200
        Genie.Show
    Case Is = "Totem"
        TextToSpeech1.StopSpeaking
        ShowHalfForm
        cmdOpen.Enabled = True
        GeniePathName = App.Path & "\Totem.acs"
        AgentX.Connected = True
        Charac = "Totem"
        Set LoadRequest = AgentX.Characters.Load("Totem", GeniePathName)
        Set Genie = AgentX.Characters("Totem")
        Set LoadShow = Genie.Get("state", "Showing, Speaking")
        Genie.MoveTo 600, 200
        Genie.Show
    Case Is = "James"
        TextToSpeech1.StopSpeaking
        ShowHalfForm
        cmdOpen.Enabled = True
        GeniePathName = App.Path & "\James.acs"
        AgentX.Connected = True
        Charac = "Joe"
        Set LoadRequest = AgentX.Characters.Load("James", GeniePathName)
        Set Genie = AgentX.Characters("James")
        Set LoadShow = Genie.Get("state", "Showing, Speaking")
        Genie.MoveTo 600, 200
        Genie.Show
    Case Is = "Santa"
        TextToSpeech1.StopSpeaking
        ShowHalfForm
        cmdOpen.Enabled = True
        GeniePathName = App.Path & "\Santa.acs"
        AgentX.Connected = True
        Charac = "Santa"
        Set LoadRequest = AgentX.Characters.Load("Santa", GeniePathName)
        Set Genie = AgentX.Characters("Santa")
        Set LoadShow = Genie.Get("state", "Showing, Speaking")
        Genie.MoveTo 600, 200
        Genie.Show
    Case Is = "Mouth"
        ShowFullForm
        GetMouthShapeValuesFromConfigFile
        cmdOpen.Enabled = False
        Charac = "Mouth"
        cboMode.ListIndex = 0
End Select

End Sub 'optGenie_Click(Index As Integer)

Private Sub GetMouthShapeValuesFromConfigFile()

hsLipTension.Value = LipTension
hsTonguePosn.Value = TonguePosn
hsMouthUpturn.Value = MouthUpturn
hsMouthHeight.Value = MouthHeight
hsMouthWidth.Value = MouthWidth
If LipTension = 0 And TonguePosn = 0 _
And MouthUpturn = 0 And MouthHeight = 0 _
And MouthWidth = 0 Then
    cmdDefault_Click
End If

End Sub 'GetMouthShapeValuesFromConfigFile()

Private Sub ShowHalfForm()

Me.Width = 7770
Frame4.Width = 6735
cmdClose.Width = 6495
cmdHelp.Left = 6960
Me.Height = 3120

End Sub 'ShowHalfForm()

Private Sub ShowFullForm()

Me.Width = 5940
Frame4.Width = 4935
cmdClose.Width = 4695
cmdHelp.Left = 5160
Me.Height = 6960

End Sub 'ShowFullForm()

Private Sub TextToSpeech1_SpeakingDone()

'chkSpkMin.Value = 0
On Error Resume Next
If chkCloseDoc.Value = 1 Then
    AppWord.ActiveDocument.Close
    AppWord.Quit
End If
If Me.WindowState = 1 Then
    Me.WindowState = 0
End If
'Save the Mouth Shape values in the config file
With CfgFile
    .LipTension = LipTension
    .TonguePosn = TonguePosn
    .MouthUpturn = MouthUpturn
    .MouthHeight = MouthHeight
    .MouthWidth = MouthWidth
End With
Put #1, NumRec, CfgFile
Close #1

End Sub 'TextToSpeech1_SpeakingDone()

Private Sub Timer1_Timer()

If CharacReadDoc = True Then
    'When the character stops speaking then,
    'Speak1Request.Status = 0 and when it is
    'speaking then, Speak1Request.Status = 2
    On Error Resume Next
    If Speak1Request.Status = 0 Then
        'If the user has opted to close
        'the word document then close it.
        If chkCloseDoc.Value = 1 Then
            AppWord.ActiveDocument.Close
            AppWord.Quit
            CharacReadDoc = False
        Else
            CharacReadDoc = False
        End If
        'If the user had opted this form to be
        'minimised while reading the word document,
        'restore the doc after the reading is over
        If chkSpkMin.Value = 1 Then
            Me.WindowState = 0
        End If
    End If
End If

End Sub 'Timer1_Timer()

Private Function FileExists(FileName As String) As Boolean

'Checks to see if a file exists on disk. Returns True if the
'file is found otherwise False.

'Setup an error trap.
On Error GoTo nofile
Open FileName For Input As #1
Close #1
'Returns True if no error occurs.
FileExists = True
Exit Function
nofile:
'If the file cannot be found return False.
FileExists = False

End Function  'FileExists

