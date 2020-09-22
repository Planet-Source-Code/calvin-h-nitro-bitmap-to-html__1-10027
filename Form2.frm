VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   0  'None
   ClientHeight    =   4320
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8640
   ControlBox      =   0   'False
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   8640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   4200
      Top             =   1920
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Splash Screen Removed To Reduce Filesize"
      Height          =   195
      Left            =   2760
      TabIndex        =   1
      Top             =   2160
      Width           =   3180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "http://www.computerpranks.org"
      Height          =   195
      Left            =   6000
      TabIndex        =   0
      Top             =   3840
      Width           =   2280
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Click()
  Form1.Show
  Unload Me
End Sub

Private Sub Timer1_Timer()
  Form1.Show
  Unload Me
End Sub
