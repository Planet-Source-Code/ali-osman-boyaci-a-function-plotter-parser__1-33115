VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7500
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6255
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7500
   ScaleWidth      =   6255
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command6 
      Caption         =   "REDRAW"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2580
      TabIndex        =   12
      Top             =   6720
      Width           =   1695
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Square"
      Height          =   375
      Left            =   4260
      TabIndex        =   11
      Top             =   7080
      Width           =   1815
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2100
      TabIndex        =   10
      Text            =   "1"
      Top             =   7080
      Width           =   465
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2040
      TabIndex        =   7
      Text            =   "10"
      Top             =   6720
      Width           =   555
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   960
      TabIndex        =   6
      Text            =   "-10"
      Top             =   6720
      Width           =   555
   End
   Begin VB.CommandButton Command4 
      Caption         =   "CLEAR"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4260
      TabIndex        =   5
      Top             =   6720
      Width           =   1800
   End
   Begin VB.CommandButton Command3 
      Caption         =   "PLOT"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3300
      TabIndex        =   3
      Top             =   6360
      Width           =   2745
   End
   Begin VB.CommandButton Command2 
      Caption         =   "CIRCLE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2580
      TabIndex        =   2
      Top             =   7080
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   660
      TabIndex        =   1
      Text            =   "sin(x)*x*x"
      Top             =   6360
      Width           =   2610
   End
   Begin MSScriptControlCtl.ScriptControl engine 
      Left            =   2700
      Top             =   2520
      _ExtentX        =   1005
      _ExtentY        =   1005
      AllowUI         =   -1  'True
   End
   Begin VB.PictureBox picBoard 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6000
      Left            =   180
      MousePointer    =   99  'Custom
      ScaleHeight     =   5940
      ScaleWidth      =   5925
      TabIndex        =   0
      Top             =   135
      Width           =   5985
   End
   Begin VB.Label Label5 
      Caption         =   "Radius"
      Height          =   255
      Left            =   1320
      TabIndex        =   13
      Top             =   7140
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "  U="
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1560
      TabIndex        =   9
      Top             =   6720
      Width           =   465
   End
   Begin VB.Label Label2 
      Caption         =   "  L="
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   480
      TabIndex        =   8
      Top             =   6720
      Width           =   465
   End
   Begin VB.Label Label1 
      Caption         =   "  y(x)="
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   120
      TabIndex        =   4
      Top             =   6240
      Width           =   645
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim l As Integer
Dim u As Integer
Option Explicit
Private Sub Command1_Click()
picBoard.Cls
picBoard.Line (10, 10)-(10, 900)
picBoard.Line (10, 10)-(900, 10)

End Sub

Private Sub Command2_Click()
picBoard.Circle ((l + u) / 2, (l + u) / 2), Text4, vbRed
picBoard.CurrentX = (l + u) / 2
picBoard.CurrentY = (l + u) / 2
picBoard.Print "Center"
End Sub

Private Sub Command3_Click()
Dim temp As String, ex As String
Dim i As Double
picBoard.Scale (l, u)-(u, l)
picBoard.Line (l, 0)-(u, 0)
picBoard.Line (0, l)-(0, u)
temp = Text1.Text

For i = l To u Step 0.01
    ex = Replace$(temp, "x", i)
    If (engine.Eval(ex) > l) And (engine.Eval(ex) < u) Then
        picBoard.PSet (i, engine.Eval(ex)), vbGreen
    End If
Next i
End Sub




Private Sub Command4_Click()
picBoard.Cls

End Sub

Private Sub Command5_Click()
picBoard.Line (l / 4, u / 4)-(l / 2, u / 2), vbBlue, BF
End Sub

Private Sub Command6_Click()
picBoard.DrawMode = 1
l = Text2.Text
u = Text3.Text
Call Command3_Click
End Sub

Private Sub Form_Load()
l = -10
u = 10
End Sub

Private Sub picBoard_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
picBoard.PSet (X, Y)
End If
End Sub

Private Sub picBoard_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
picBoard.PSet (X, Y)
End If
End Sub


