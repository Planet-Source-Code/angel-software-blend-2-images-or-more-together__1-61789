VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Blend 2 pictures together"
   ClientHeight    =   6420
   ClientLeft      =   36
   ClientTop       =   420
   ClientWidth     =   4476
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6420
   ScaleWidth      =   4476
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar pbar 
      Height          =   132
      Left            =   120
      TabIndex        =   8
      Top             =   5640
      Width           =   4212
      _ExtentX        =   7430
      _ExtentY        =   233
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComDlg.CommonDialog dlg 
      Left            =   3000
      Top             =   2880
      _ExtentX        =   677
      _ExtentY        =   677
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.PictureBox pic1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1092
      Left            =   600
      Picture         =   "form1.frx":0000
      ScaleHeight     =   1068
      ScaleWidth      =   1428
      TabIndex        =   7
      Top             =   240
      Width           =   1452
   End
   Begin VB.CommandButton changePic2 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   312
      Left            =   2520
      TabIndex        =   6
      Top             =   1440
      Width           =   372
   End
   Begin VB.CommandButton changePic1 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   312
      Left            =   960
      TabIndex        =   5
      Top             =   1440
      Width           =   372
   End
   Begin VB.CommandButton Command2 
      Caption         =   "cls"
      Height          =   312
      Left            =   3360
      TabIndex        =   4
      Top             =   5880
      Width           =   492
   End
   Begin VB.CommandButton go 
      Caption         =   "Go!"
      Default         =   -1  'True
      Height          =   252
      Left            =   2040
      TabIndex        =   3
      Top             =   2400
      Width           =   492
   End
   Begin VB.PictureBox picResult 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2172
      Left            =   120
      ScaleHeight     =   2148
      ScaleWidth      =   4188
      TabIndex        =   1
      Top             =   3360
      Width           =   4212
   End
   Begin VB.PictureBox pic2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1092
      Left            =   2280
      Picture         =   "form1.frx":08CC
      ScaleHeight     =   1068
      ScaleWidth      =   1428
      TabIndex        =   0
      Top             =   240
      Width           =   1452
   End
   Begin VB.Line Line2 
      X1              =   720
      X2              =   3720
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "V"
      Height          =   252
      Left            =   2160
      TabIndex        =   2
      Top             =   3000
      Width           =   252
   End
   Begin VB.Line Line1 
      X1              =   2280
      X2              =   2280
      Y1              =   2040
      Y2              =   3000
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   372
      Left            =   3960
      Shape           =   3  'Circle
      Top             =   5880
      Width           =   372
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Type RGB
    Red As Long
    Green As Long
    Blue As Long
End Type

Dim x As Integer
Dim y As Integer
Dim B As Long, G As Long, R As Long



'converts long values to RGB values
Private Function longToRGB(LongNumber) As RGB
Dim B As Long, G As Long, R As Long

 B = LongNumber \ 65536
 G = (LongNumber - B * 65536) \ 256
 R = LongNumber - B * 65536 - G * 256
 
longToRGB.Blue = B
longToRGB.Green = G
longToRGB.Red = R

End Function




Private Sub changePic1_Click()
On Error GoTo handler
dlg.ShowOpen
pic1.Picture = LoadPicture(dlg.FileName)
handler:
End Sub

Private Sub changePic2_Click()
On Error GoTo handler
dlg.ShowOpen
pic2.Picture = LoadPicture(dlg.FileName)
handler:
End Sub



Private Sub changePic3_Click()
On Error GoTo handler
dlg.ShowOpen
pic3.Picture = LoadPicture(dlg.FileName)
handler:
End Sub

Private Sub go_Click()
'busy...
Shape1.FillColor = vbRed
Me.MousePointer = 11
pbar.Max = pic1.Width - 3
DoEvents
For x = 0 To pic1.Width - 3
    For y = 0 To pic1.Height - 3
    
    Dim point1 As Long, point2 As Long
    Dim rgb1 As RGB, rgb2 As RGB
    
        'get color of point
        point1 = pic1.Point(x, y)
        point2 = pic2.Point(x, y)
        
        
        'I had to write this, if you erase them
        'this stops working I dont know why :S
        If point1 = -1 Then point1 = 5460297
        If point2 = -1 Then point2 = 5460297
        
        'get RGB of the first point
        rgb1 = longToRGB(point1)
        r1 = rgb1.Red
        g1 = rgb1.Green
        b1 = rgb1.Blue
        
        'get RGB of the second point
        rgb2 = longToRGB(point2)
        r2 = rgb2.Red
        g2 = rgb2.Green
        b2 = rgb2.Blue
        
        
        'average the vaule of both points
        r3 = (r1 + r2) / 2
        g3 = (g1 + g2) / 2
        b3 = (b1 + b2) / 2
        
        'Set the forecolor to the average
        picResult.ForeColor = RGB(r3, g3, b3)
        'Set the point :D
        picResult.PSet (x, y)
        pbar = x
    
    Next y
Next x
'ready
Shape1.FillColor = vbGreen
Me.MousePointer = 1
pbar = 0
End Sub

Private Sub Command2_Click()
picResult.Cls
End Sub

Private Sub Form_Load()


Shape1.FillColor = vbGreen

'set all the scalemodes to pixels
picResult.ScaleMode = vbPixels
picResult.ScaleMode = vbPixels
pic1.ScaleMode = vbPixels
pic2.ScaleMode = vbPixels
Me.ScaleMode = vbPixels
End Sub



Private Sub Label1_Click()

End Sub
