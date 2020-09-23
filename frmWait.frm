VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmWait 
   Caption         =   "Connecting..."
   ClientHeight    =   1395
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmWait.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1395
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin MSComCtl2.Animation Ani1 
      Height          =   555
      Left            =   1845
      TabIndex        =   1
      Top             =   630
      Width           =   645
      _ExtentX        =   1138
      _ExtentY        =   979
      _Version        =   393216
      AutoPlay        =   -1  'True
      Center          =   -1  'True
      FullWidth       =   43
      FullHeight      =   37
   End
   Begin VB.Label lblConnecting 
      Caption         =   "Please wait while that information is gathered."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   180
      TabIndex        =   0
      Top             =   135
      Width           =   4335
   End
End
Attribute VB_Name = "frmWait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Ani1.Open App.Path & "\transferdata.avi"
    frmWait.Left = (Screen.Width - frmWait.Width) / 2
    frmWait.Top = (Screen.Height - frmWait.Height) / 2
End Sub
Private Sub Form_Activate()
    frmWait.Visible = True
End Sub

