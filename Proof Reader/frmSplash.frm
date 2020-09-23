VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   0  'None
   ClientHeight    =   1935
   ClientLeft      =   2445
   ClientTop       =   1380
   ClientWidth     =   6660
   LinkTopic       =   "Form1"
   ScaleHeight     =   1935
   ScaleWidth      =   6660
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraMainFrame 
      BackColor       =   &H00FF8080&
      Height          =   2085
      Left            =   -240
      TabIndex        =   0
      Top             =   -120
      Width           =   6900
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00FF8080&
         BorderStyle     =   0  'None
         Height          =   585
         Left            =   600
         Picture         =   "frmSplash.frx":0000
         ScaleHeight     =   585
         ScaleWidth      =   615
         TabIndex        =   6
         Top             =   600
         Width           =   615
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FF8080&
         BorderStyle     =   0  'None
         Height          =   585
         Left            =   6000
         Picture         =   "frmSplash.frx":030A
         ScaleHeight     =   585
         ScaleWidth      =   615
         TabIndex        =   5
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Sunil Wason"
         Height          =   255
         Left            =   4440
         TabIndex        =   4
         Top             =   1800
         Width           =   2055
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Developed by"
         Height          =   255
         Left            =   4440
         TabIndex        =   3
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label lblLoad 
         Alignment       =   2  'Center
         BackColor       =   &H00FF8080&
         Height          =   735
         Left            =   2280
         TabIndex        =   2
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H00404040&
         Height          =   615
         Left            =   1200
         TabIndex        =   1
         Top             =   480
         Width           =   4695
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

lblLoad.Caption = "Initialising. Please wait..."

End Sub


