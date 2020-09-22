VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Macro Creator"
   ClientHeight    =   4530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9885
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   9885
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7200
      TabIndex        =   12
      Text            =   "Nitro"
      Top             =   4125
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   9
      Text            =   "#000000"
      Top             =   4125
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3840
      TabIndex        =   7
      Text            =   "2"
      Top             =   4125
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Create Macro"
      Height          =   375
      Left            =   8520
      TabIndex        =   3
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Load Image"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   4080
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2520
      TabIndex        =   0
      Text            =   "X"
      Top             =   4120
      Width           =   735
   End
   Begin VB.PictureBox Picture2 
      Height          =   3615
      Left            =   0
      ScaleHeight     =   3555
      ScaleWidth      =   9795
      TabIndex        =   4
      Top             =   360
      Width           =   9855
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   1785
         Left            =   0
         Picture         =   "Form1.frx":1042
         ScaleHeight     =   1785
         ScaleWidth      =   2775
         TabIndex        =   5
         Top             =   0
         Width           =   2775
         Begin MSComDlg.CommonDialog CDB 
            Left            =   1560
            Top             =   480
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
            CancelError     =   -1  'True
         End
      End
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Title:"
      Height          =   195
      Left            =   6720
      TabIndex        =   11
      Top             =   4155
      Width           =   345
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Background Color:"
      Height          =   195
      Left            =   4320
      TabIndex        =   10
      Top             =   4155
      Width           =   1320
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Step:"
      Height          =   195
      Left            =   3360
      TabIndex        =   8
      Top             =   4155
      Width           =   375
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "HTML Macro creator by Nitro. http://www.computerpranks.org"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   60
      Width           =   4440
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pixel Character:"
      Height          =   195
      Left            =   1320
      TabIndex        =   1
      Top             =   4155
      Width           =   1110
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: HTML Macro Creator
'Author: Calvin H. (Nitro)
'Web Site: http://www.computerpranks.org
'E-mail: nitro@computerpranks.org
'Comments: This is a simple program to convert images to html.
'          feel free to modify this code, this program is open source.

Private Sub Command1_Click()
On Error GoTo NoLoad
  CDB.ShowOpen
  Picture1.Picture = LoadPicture(CDB.FileName)
NoLoad:
End Sub

Private Sub Command2_Click()
On Error GoTo NoSave
  CDB.FileName = ""
  CDB.Filter = "HTML Files|*.html"
  CDB.ShowSave
  CDB.Filter = ""
  c = 1
  Open CDB.FileName For Output As #1
    Print #1, "<html><head><title>" & Text4.Text & " Created With Macro Creator</title></head><body bgcolor=" & Text3.Text & Chr(34) & "><pre><font face=" & Chr(34) & "courier new" & Chr(34) & " size=1>"
    
    ph = Picture1.Height / Screen.TwipsPerPixelY
    pw = Picture1.Width / Screen.TwipsPerPixelX
    
    For y = 0 To Picture1.Height / Screen.TwipsPerPixelY - 1 Step 1.7 * Text2.Text
    
      For x = 0 To Picture1.Width / Screen.TwipsPerPixelX - 1 Step 1 * Text2.Text
      
        'get color of a pixel on image
        pix = Picture1.Point(Int(x / pw * Picture1.Width), Int(y / ph * Picture1.Height))
        '              % of X position * width of picture,  % of Y position * height of picture
        
        'check for change in color. if there is a change then we insert a new font color tag here
        If pix <> opix Then
          'get the hex value of the color
          hpx = Hex(pix)
          'add zeros to create a 6 digit hex value
          For z = 5 To Len(hpx) Step -1
            hpx = "0" & hpx
          Next z
          'switch red and blue values
          hpix = Mid(hpx, 5) & Mid(hpx, 3, 2) & Left(hpx, 2)
          'generate font color tag
          s = s & "<font color=" & Chr(34) & "#" & hpix & Chr(34) & ">"
          'set new color as old color
          opix = pix
        End If
        'insert macro character
        s = s & Mid(Text1.Text, c, 1)
        'goto next character
        c = c + 1
        If c > Len(Text1.Text) Then c = 1
        
      Next x
      
      DoEvents
      'show percentage done
      Me.Caption = "Macro Creator [" & Int(y / ph * 1000) / 10 & "%" & "]"
      'print to file
      Print #1, s
      s = ""
    Next y
    
    Me.Caption = "Macro Creator"
    
    Print #1, "</pre></body></html>"
    
    MsgBox "File Created. File Size: " & Int(LOF(1) / 1000) & "kb"
  Close #1
NoSave:
End Sub

Private Sub Form_Load()
  Picture1_Resize
End Sub

Private Sub Picture1_Resize()
  'centers picture
  Picture1.Top = Picture2.Height / 2 - Picture1.Height / 2
  Picture1.Left = Picture2.Width / 2 - Picture1.Width / 2
End Sub

Private Sub Text3_Click()
  On Error GoTo NoChoose
    CDB.ShowColor
    bgc = Hex(CDB.Color)
    For z = 5 To Len(bgc) Step -1
      bgc = "0" & bgc
    Next z
    Text3.Text = "#" & Mid(bgc, 5) & Mid(bgc, 3, 2) & Left(bgc, 2)
    Text3.BackColor = CDB.Color
NoChoose:
End Sub
