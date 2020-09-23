VERSION 5.00
Begin VB.Form frmAbout 
   Caption         =   "Credits"
   ClientHeight    =   2265
   ClientLeft      =   1830
   ClientTop       =   5490
   ClientWidth     =   6795
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2265
   ScaleWidth      =   6795
   Begin VB.PictureBox Picture4 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Height          =   585
      Left            =   5400
      Picture         =   "frmAbout.frx":0442
      ScaleHeight     =   585
      ScaleWidth      =   615
      TabIndex        =   13
      Top             =   720
      Width           =   615
      Begin VB.PictureBox Picture8 
         BackColor       =   &H00FF8080&
         BorderStyle     =   0  'None
         Height          =   585
         Left            =   1200
         Picture         =   "frmAbout.frx":0884
         ScaleHeight     =   585
         ScaleWidth      =   615
         TabIndex        =   17
         Top             =   0
         Width           =   615
      End
      Begin VB.PictureBox Picture7 
         BackColor       =   &H00FF8080&
         BorderStyle     =   0  'None
         Height          =   585
         Left            =   1800
         Picture         =   "frmAbout.frx":0CC6
         ScaleHeight     =   585
         ScaleWidth      =   615
         TabIndex        =   16
         Top             =   0
         Width           =   615
      End
      Begin VB.PictureBox Picture6 
         BackColor       =   &H00FF8080&
         BorderStyle     =   0  'None
         Height          =   585
         Left            =   0
         Picture         =   "frmAbout.frx":1108
         ScaleHeight     =   585
         ScaleWidth      =   615
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   0
         Width           =   615
      End
      Begin VB.PictureBox Picture5 
         BackColor       =   &H00FF8080&
         BorderStyle     =   0  'None
         Height          =   585
         Left            =   600
         Picture         =   "frmAbout.frx":1412
         ScaleHeight     =   585
         ScaleWidth      =   615
         TabIndex        =   14
         Top             =   0
         Width           =   615
      End
   End
   Begin VB.PictureBox Picture13 
      Height          =   1815
      Left            =   120
      Picture         =   "frmAbout.frx":1854
      ScaleHeight     =   1.672
      ScaleMode       =   0  'User
      ScaleWidth      =   0.969
      TabIndex        =   12
      Top             =   360
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Interval        =   150
      Left            =   3240
      Top             =   720
   End
   Begin VB.CommandButton cmdSysInfo 
      Caption         =   "&System Info..."
      Height          =   345
      Left            =   2880
      TabIndex        =   8
      Tag             =   "&System Info..."
      Top             =   1440
      Width           =   1725
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Default         =   -1  'True
      Height          =   345
      Left            =   4680
      MouseIcon       =   "frmAbout.frx":9D56
      TabIndex        =   7
      Tag             =   "OK"
      Top             =   1440
      Width           =   1860
   End
   Begin VB.PictureBox picLogo 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Height          =   585
      Left            =   4800
      Picture         =   "frmAbout.frx":A198
      ScaleHeight     =   585
      ScaleWidth      =   615
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   720
      Width           =   615
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Height          =   585
      Left            =   4200
      Picture         =   "frmAbout.frx":A4A2
      ScaleHeight     =   585
      ScaleWidth      =   615
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   720
      Width           =   615
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Height          =   585
      Left            =   6000
      Picture         =   "frmAbout.frx":A7AC
      ScaleHeight     =   585
      ScaleWidth      =   615
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   720
      Width           =   615
   End
   Begin VB.PictureBox Picture9 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Height          =   585
      Left            =   3600
      Picture         =   "frmAbout.frx":AAB6
      ScaleHeight     =   585
      ScaleWidth      =   615
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   720
      Width           =   615
   End
   Begin VB.PictureBox Picture10 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Height          =   585
      Left            =   2400
      Picture         =   "frmAbout.frx":ADC0
      ScaleHeight     =   585
      ScaleWidth      =   615
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   720
      Width           =   615
   End
   Begin VB.PictureBox Picture11 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Height          =   585
      Left            =   1800
      Picture         =   "frmAbout.frx":B0CA
      ScaleHeight     =   585
      ScaleWidth      =   615
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   720
      Width           =   615
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   1320
      ScaleHeight     =   255
      ScaleWidth      =   5295
      TabIndex        =   11
      Top             =   1920
      Width           =   5295
   End
   Begin VB.PictureBox Picture12 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Height          =   585
      Left            =   3000
      Picture         =   "frmAbout.frx":B3D4
      ScaleHeight     =   585
      ScaleWidth      =   615
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   720
      Width           =   615
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   1800
      X2              =   6600
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label Label2 
      Caption         =   "Version 1.0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1680
      TabIndex        =   9
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   1800
      X2              =   6615
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Proof Reader"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   1560
      TabIndex        =   10
      Top             =   120
      Width           =   5175
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   3000
      Top             =   720
      Width           =   615
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Reg Key Security Options...
Const KEY_ALL_ACCESS = &H2003F
                                          

' Reg Key ROOT Types...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' Unicode nul terminated string
Const REG_DWORD = 4                      ' 32-bit number


Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"


Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long

Private Sub Form_Load()
    'lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    'lblTitle.Caption = App.Title
End Sub

Private Sub cmdSysInfo_Click()

Call StartSysInfo

End Sub 'cmdSysInfo_Click

Private Sub cmdOK_Click()

Unload Me

End Sub 'cmdOK_Click


Public Sub StartSysInfo()
    On Error GoTo SysInfoErr


        Dim rc As Long
        Dim SysInfoPath As String
        

        ' Try To Get System Info Program Path\Name From Registry...
        If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
        ' Try To Get System Info Program Path Only From Registry...
        ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
                ' Validate Existance Of Known 32 Bit File Version
                If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
                        SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
                        

                ' Error - File Can Not Be Found...
                Else
                        GoTo SysInfoErr
                End If
        ' Error - Registry Entry Can Not Be Found...
        Else
                GoTo SysInfoErr
        End If
        

        Call Shell(SysInfoPath, vbNormalFocus)
        

        Exit Sub
SysInfoErr:
        MsgBox "System Information Is Unavailable At This Time", vbOKOnly
End Sub


Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
        Dim i As Long                                           ' Loop Counter
        Dim rc As Long                                          ' Return Code
        Dim hKey As Long                                        ' Handle To An Open Registry Key
        Dim hDepth As Long                                      '
        Dim KeyValType As Long                                  ' Data Type Of A Registry Key
        Dim tmpVal As String                                    ' Tempory Storage For A Registry Key Value
        Dim KeyValSize As Long                                  ' Size Of Registry Key Variable
        '------------------------------------------------------------
        ' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}
        '------------------------------------------------------------
        rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Open Registry Key
        

        If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Error...
        

        tmpVal = String$(1024, 0)                             ' Allocate Variable Space
        KeyValSize = 1024                                       ' Mark Variable Size
        

        '------------------------------------------------------------
        ' Retrieve Registry Key Value...
        '------------------------------------------------------------
        rc = RegQueryValueEx(hKey, SubKeyRef, 0, KeyValType, tmpVal, KeyValSize)    ' Get/Create Key Value
                                                

        If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Errors
        

        If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 Adds Null Terminated String...
                tmpVal = Left(tmpVal, KeyValSize - 1)               ' Null Found, Extract From String
        Else                                                    ' WinNT Does NOT Null Terminate String...
                tmpVal = Left(tmpVal, KeyValSize)                   ' Null Not Found, Extract String Only
        End If
        '------------------------------------------------------------
        ' Determine Key Value Type For Conversion...
        '------------------------------------------------------------
        Select Case KeyValType                                  ' Search Data Types...
        Case REG_SZ                                             ' String Registry Key Data Type
                KeyVal = tmpVal                                     ' Copy String Value
        Case REG_DWORD                                          ' Double Word Registry Key Data Type
                For i = Len(tmpVal) To 1 Step -1                    ' Convert Each Bit
                        KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Build Value Char. By Char.
                Next
                KeyVal = Format$("&h" + KeyVal)                     ' Convert Double Word To String
        End Select
        

        GetKeyValue = True                                      ' Return Success
        rc = RegCloseKey(hKey)                                  ' Close Registry Key
        Exit Function                                           ' Exit
        

GetKeyError:    ' Cleanup After An Error Has Occured...
        KeyVal = ""                                             ' Set Return Val To Empty String
        GetKeyValue = False                                     ' Return Failure
        rc = RegCloseKey(hKey)                                  ' Close Registry Key
End Function

Private Sub Timer1_Timer()

ShowMessage

End Sub 'Timer1_Timer()

Sub ShowMessage()

    Static MsgPtr As Integer
    Static MyText As String
    If Len(MyText) = 0 Then
      MsgPtr = 1
      MyText = "    This application has been developed by Sunil Wason. For any queries, incompatiblaties or problems, contact sunilwason@yahoo.com    "
    End If
    Picture1.Cls
    'Position the text to be displayed.
    Picture1.CurrentX = 0
    Picture1.CurrentY = 50
    Picture1.Print Mid$(MyText, MsgPtr); MyText; MyText; MyText;
    MsgPtr = MsgPtr + 1
    If MsgPtr > Len(MyText) Then
      MsgPtr = 1
    End If
    
End Sub 'ShowMessage()
