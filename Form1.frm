VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Texture generator 3 (made by JOHANNES BOHMAN)"
   ClientHeight    =   6240
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7560
   LinkTopic       =   "Form1"
   ScaleHeight     =   416
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   504
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7215
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   4320
         MaxLength       =   2
         TabIndex        =   16
         Text            =   "4"
         Top             =   840
         Width           =   255
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   4080
         MaxLength       =   2
         TabIndex        =   15
         Text            =   "2"
         Top             =   840
         Width           =   255
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   3840
         MaxLength       =   2
         TabIndex        =   14
         Text            =   "1"
         Top             =   840
         Width           =   255
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Auto refresh"
         Height          =   255
         Left            =   2280
         TabIndex        =   13
         Top             =   1080
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Random"
         Height          =   255
         Left            =   5280
         TabIndex        =   10
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   3480
         MaxLength       =   2
         TabIndex        =   9
         Text            =   "1"
         Top             =   480
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Color cycling (RGB)"
         Height          =   195
         Left            =   2400
         TabIndex        =   8
         Top             =   300
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1680
         MaxLength       =   2
         TabIndex        =   7
         Text            =   "1"
         Top             =   600
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Caption         =   "DRAW"
         Height          =   255
         Left            =   5280
         TabIndex        =   6
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   960
         MaxLength       =   2
         TabIndex        =   5
         Text            =   "1"
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   960
         MaxLength       =   4
         TabIndex        =   3
         Text            =   "1100"
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "Incrase color (RGB)"
         Height          =   255
         Left            =   2400
         TabIndex        =   17
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "B:"
         Height          =   255
         Left            =   1440
         TabIndex        =   12
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Speed:"
         Height          =   255
         Left            =   2880
         TabIndex        =   11
         Top             =   480
         Width           =   615
      End
      Begin VB.Shape Shape1 
         Height          =   975
         Left            =   2280
         Top             =   240
         Width           =   4695
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Step/        A:  Transparancy"
         Height          =   495
         Left            =   0
         TabIndex        =   4
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Loop to:"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   4455
      Left            =   0
      ScaleHeight     =   293
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   477
      TabIndex        =   0
      Top             =   1320
      Width           =   7215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim A, B As Integer
Dim col
Dim F1, F2, F3 As Integer
Dim F11, F22, F33 As Integer
Dim N As Integer
Private Sub Command1_Click()
On Error Resume Next
A = 0
B = 0
Picture1.Cls
Do

If Check1.Value = 1 Then Picture1.Refresh

If Option1.Value = True Then
N = N + 1
'FÄRG


If N >= Text4.Text Then

If F1 >= 255 Then F11 = 1
If F2 >= 255 Then F22 = 1
If F3 >= 255 Then F33 = 1

If F1 <= 0 Then F11 = 0
If F2 <= 0 Then F22 = 0
If F3 <= 0 Then F33 = 0

If F11 = 0 Then
F1 = F1 + Text5.Text
Else
F1 = F1 - Text5.Text
End If

If F22 = 0 Then
F2 = F2 + Text6.Text
Else
F2 = F2 - Text6.Text
End If

If F33 = 0 Then
F3 = F3 + Text7.Text
Else
F3 = F3 - Text7.Text
End If

col = RGB(F1, F2, F3)
N = 0

End If
End If

If Option2.Value = True Then
col = RGB(Rnd * 255, Rnd * 255, Rnd * 255)
End If





A = A + Val(Text2.Text)
B = B + Val(Text3.Text)




Picture1.Line (A, B)-(Picture1.ScaleWidth - A, Picture1.ScaleHeight - B), col, B
Loop Until A >= Val(Text1.Text)
End Sub

Private Sub Form_Load()
F2 = 100
F3 = 244
End Sub

Private Sub Form_Resize()
Picture1.Width = Form1.ScaleWidth
Picture1.Height = Form1.ScaleHeight - 80
End Sub


Private Sub Form_Unload(Cancel As Integer)
MsgBox "Please vote if you liked it!"
End Sub


