VERSION 5.00
Object = "{2398E321-5C6E-11D1-8C65-0060081841DE}#1.0#0"; "Vtext.dll"
Begin VB.Form frmMouth 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3240
   ClientLeft      =   2820
   ClientTop       =   5145
   ClientWidth     =   4290
   LinkTopic       =   "Form1"
   ScaleHeight     =   3240
   ScaleWidth      =   4290
   ShowInTaskbar   =   0   'False
   Begin HTTSLibCtl.TextToSpeech TextToSpeech1 
      Height          =   3135
      Left            =   0
      OleObjectBlob   =   "frmMouth.frx":0000
      TabIndex        =   0
      Top             =   0
      Width           =   4335
   End
End
Attribute VB_Name = "frmMouth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

