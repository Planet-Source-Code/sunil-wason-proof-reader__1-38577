Attribute VB_Name = "Module1"
Type Text2SpeechCfgFile 'Configuration file
  Charac As Byte
  LipTension As Byte
  TonguePosn As Byte
  MouthUpturn As Byte
  MouthHeight As Byte
  MouthWidth As Byte
  RecNum As Byte
End Type
Public CfgFile As Text2SpeechCfgFile

Public CurrRec As Byte  'Current Record of the database.
Public NumRec As Byte  'Total No of Records in the database.
Public CharacOpen As Byte 'Indicates the index No of the character which is opened.
Public LipTension As Byte
Public TonguePosn As Byte
Public MouthUpturn As Byte
Public MouthHeight As Byte
Public MouthWidth As Byte
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_SHOWWINDOW = &H40
  
Public Sub AlwaysOnTop(formname As Form, SetOnTop As Boolean)
Dim lFlag As Long
    If SetOnTop Then
        lFlag = HWND_TOPMOST
    Else
        lFlag = HWND_NOTOPMOST
    End If
    SetWindowPos formname.hwnd, lFlag, _
    formname.Left / Screen.TwipsPerPixelX, _
    formname.Top / Screen.TwipsPerPixelY, _
    formname.Width / Screen.TwipsPerPixelX, _
    formname.Height / Screen.TwipsPerPixelY, _
    SWP_NOACTIVATE Or SWP_SHOWWINDOW
End Sub


Sub Main()

    Screen.MousePointer = 11
    frmSplash.Show
    frmSplash.Refresh
    frmSplash.lblLoad.Caption = "Loading Characters. Please be patient..."
    Load frmTxtToSpeech
    Unload frmSplash
    frmTxtToSpeech.Show
    Screen.MousePointer = 0
    
End Sub 'Main()

