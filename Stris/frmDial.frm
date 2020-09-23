VERSION 5.00
Begin VB.Form frmDial 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Stris"
   ClientHeight    =   2865
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4410
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2865
   ScaleWidth      =   4410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton comDef 
      Caption         =   "Start"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   1545
      TabIndex        =   0
      Top             =   2280
      Width           =   2745
   End
   Begin VB.CommandButton comCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   450
      Left            =   105
      TabIndex        =   1
      Top             =   2280
      Width           =   1380
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2010
      Left            =   765
      TabIndex        =   2
      Top             =   135
      Width           =   3555
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   150
      Picture         =   "frmDial.frx":0000
      Top             =   165
      Width           =   480
   End
End
Attribute VB_Name = "frmDial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub comCancel_Click()
   OK = False: Unload Me
End Sub

Private Sub comDef_Click()
   OK = True: Unload Me
End Sub


Private Sub Form_Load()
   If DialogTitle = "Stris - Spel starten" Then
      comCancel.Caption = "Cancel"
      comDef.Caption = "Start"
      End If
   If DialogTitle = "Stris - Pauze" Then
      comCancel.Caption = "Stop"
      comDef.Caption = "Continue"
      End If
   If DialogTitle = "Stris - Game over" Then
      comCancel.Visible = False
      comDef.Caption = "OK"
      End If
   Caption = DialogTitle
   Label1.Caption = DialogText
End Sub


