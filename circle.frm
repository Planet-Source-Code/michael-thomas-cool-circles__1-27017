VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   Caption         =   "Spiral Circle"
   ClientHeight    =   5310
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7440
   Icon            =   "circle.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   354
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   496
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtpb 
      Height          =   285
      Left            =   840
      TabIndex        =   10
      Text            =   "250"
      Top             =   3720
      Width           =   495
   End
   Begin VB.TextBox txtpg 
      Height          =   285
      Left            =   840
      TabIndex        =   9
      Text            =   "0"
      Top             =   3360
      Width           =   495
   End
   Begin VB.TextBox txtpr 
      Height          =   285
      Left            =   840
      TabIndex        =   8
      Text            =   "100"
      Top             =   3000
      Width           =   495
   End
   Begin VB.TextBox txtblu 
      Height          =   285
      Left            =   240
      TabIndex        =   7
      Text            =   ".7"
      Top             =   3720
      Width           =   495
   End
   Begin VB.TextBox txtgre 
      Height          =   285
      Left            =   240
      TabIndex        =   6
      Text            =   ".1"
      Top             =   3360
      Width           =   495
   End
   Begin VB.TextBox txtred 
      Height          =   285
      Left            =   240
      TabIndex        =   5
      Text            =   ".6"
      Top             =   3000
      Width           =   495
   End
   Begin VB.TextBox txtn 
      Height          =   285
      Left            =   240
      TabIndex        =   4
      Text            =   "50"
      Top             =   2040
      Width           =   855
   End
   Begin VB.TextBox txtr 
      Height          =   285
      Left            =   240
      TabIndex        =   3
      Text            =   "50"
      Top             =   1680
      Width           =   855
   End
   Begin VB.TextBox txty 
      Height          =   285
      Left            =   240
      TabIndex        =   2
      Text            =   "200"
      Top             =   960
      Width           =   855
   End
   Begin VB.TextBox txtx 
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Text            =   "300"
      Top             =   1320
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Make It!"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Blue"
      Height          =   255
      Index           =   2
      Left            =   1440
      TabIndex        =   18
      Top             =   3720
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "Green"
      Height          =   255
      Index           =   1
      Left            =   1440
      TabIndex        =   17
      Top             =   3360
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "Red"
      Height          =   255
      Index           =   0
      Left            =   1440
      TabIndex        =   16
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Colour"
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   2640
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Density"
      Height          =   255
      Index           =   3
      Left            =   1200
      TabIndex        =   14
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Size"
      Height          =   255
      Index           =   2
      Left            =   1200
      TabIndex        =   13
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "X"
      Height          =   255
      Index           =   1
      Left            =   1200
      TabIndex        =   12
      Top             =   960
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "Y"
      Height          =   255
      Index           =   0
      Left            =   1200
      TabIndex        =   11
      Top             =   1320
      Width           =   255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form1.Cls
r = Val(txtr.Text)
xc = Val(txtx.Text)
yc = Val(txty.Text)
For angle = 0 To 360 Step 360 / Val(txtn.Text)
x = r * Cos(1.74532925199433E-02 * angle) + xc
y = r * Sin(1.74532925199433E-02 * angle) + yc
Form1.Circle (x, y), r, RGB((Val(txtred) * angle) + Val(txtpr), (Val(txtgre) * angle) + Val(txtpg), (Val(txtgre) * angle) + Val(txtpb))
Next
End Sub

Private Sub Form_Load()
Form1.Cls
r = Val(txtr.Text)
xc = Val(txtx.Text)
yc = Val(txty.Text)
For angle = 0 To 360 Step 360 / Val(txtn.Text)
x = r * Cos(1.74532925199433E-02 * angle) + xc
y = r * Sin(1.74532925199433E-02 * angle) + yc
Form1.Circle (x, y), r, RGB((Val(txtred) * angle) + Val(txtpr), (Val(txtgre) * angle) + Val(txtpg), (Val(txtgre) * angle) + Val(txtpb))
Next
End Sub
