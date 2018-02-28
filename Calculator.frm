VERSION 5.00
Begin VB.Form Calculator 
   Caption         =   "Calculator"
   ClientHeight    =   4170
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3720
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   3720
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdMod 
      Caption         =   "Mod"
      Height          =   375
      Left            =   3000
      TabIndex        =   28
      Top             =   1200
      Width           =   615
   End
   Begin VB.CommandButton cmdPOW 
      Caption         =   "x^2"
      Height          =   375
      Left            =   3000
      TabIndex        =   27
      Top             =   1680
      Width           =   615
   End
   Begin VB.CommandButton cmdLog 
      Caption         =   "Log"
      Height          =   375
      Left            =   2280
      TabIndex        =   26
      Top             =   1200
      Width           =   615
   End
   Begin VB.CommandButton cmdTan 
      Caption         =   "tan"
      Height          =   375
      Left            =   1560
      TabIndex        =   25
      Top             =   1200
      Width           =   615
   End
   Begin VB.CommandButton cmdCos 
      Caption         =   "Cos"
      Height          =   375
      Left            =   840
      TabIndex        =   24
      Top             =   1200
      Width           =   615
   End
   Begin VB.CommandButton cmdSin 
      Caption         =   "Sin"
      Height          =   375
      Left            =   120
      TabIndex        =   23
      Top             =   1200
      Width           =   615
   End
   Begin VB.CommandButton cmd2 
      Caption         =   "2"
      Height          =   375
      Left            =   840
      TabIndex        =   22
      Top             =   3120
      Width           =   615
   End
   Begin VB.CommandButton cmd1 
      Caption         =   "1"
      Height          =   375
      Left            =   120
      TabIndex        =   21
      Top             =   3120
      Width           =   615
   End
   Begin VB.CommandButton cmdADD 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   20
      Top             =   3120
      Width           =   615
   End
   Begin VB.CommandButton cmdDOT 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   19
      Top             =   3600
      Width           =   615
   End
   Begin VB.CommandButton cmd0 
      Caption         =   "0"
      Height          =   375
      Left            =   120
      TabIndex        =   18
      Top             =   3600
      Width           =   615
   End
   Begin VB.CommandButton cmdEQUALS 
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   17
      Top             =   3600
      Width           =   1335
   End
   Begin VB.CommandButton cmdSUBTRACT 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   16
      Top             =   2640
      Width           =   615
   End
   Begin VB.CommandButton cmd3 
      Caption         =   "3"
      Height          =   375
      Left            =   1560
      TabIndex        =   15
      Top             =   3120
      Width           =   615
   End
   Begin VB.CommandButton cmdinverse 
      Caption         =   "1/x"
      Height          =   375
      Left            =   3000
      TabIndex        =   14
      Top             =   3120
      Width           =   615
   End
   Begin VB.CommandButton cmdMULTIPLY 
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   13
      Top             =   2160
      Width           =   615
   End
   Begin VB.CommandButton cmd6 
      Caption         =   "6"
      Height          =   375
      Left            =   1560
      TabIndex        =   12
      Top             =   2640
      Width           =   615
   End
   Begin VB.CommandButton cmd5 
      Caption         =   "5"
      Height          =   375
      Left            =   840
      TabIndex        =   11
      Top             =   2640
      Width           =   615
   End
   Begin VB.CommandButton cmd4 
      Caption         =   "4"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   2640
      Width           =   615
   End
   Begin VB.CommandButton cmdPercent 
      Caption         =   "%"
      Height          =   375
      Left            =   3000
      TabIndex        =   9
      Top             =   2640
      Width           =   615
   End
   Begin VB.CommandButton cmdDIVIDE 
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   8
      Top             =   1680
      Width           =   615
   End
   Begin VB.CommandButton cmd9 
      Caption         =   "9"
      Height          =   375
      Left            =   1560
      TabIndex        =   7
      Top             =   2160
      Width           =   615
   End
   Begin VB.CommandButton cmd8 
      Caption         =   "8"
      Height          =   375
      Left            =   840
      TabIndex        =   6
      Top             =   2160
      Width           =   615
   End
   Begin VB.CommandButton cmd7 
      Caption         =   "7"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   2160
      Width           =   615
   End
   Begin VB.CommandButton CmdSqrt 
      Caption         =   "sqrt"
      Height          =   375
      Left            =   3000
      TabIndex        =   4
      Top             =   2160
      Width           =   615
   End
   Begin VB.CommandButton cmdSIGN 
      Caption         =   "+/-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   3
      Top             =   3600
      Width           =   615
   End
   Begin VB.CommandButton cmdCANCEL 
      Caption         =   "C"
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   1680
      Width           =   615
   End
   Begin VB.CommandButton cmdEXIT 
      Caption         =   "CE"
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   1680
      Width           =   615
   End
   Begin VB.CommandButton cmdBackspace 
      Caption         =   "<-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label Result 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000014&
      Height          =   495
      Left            =   120
      TabIndex        =   30
      Top             =   600
      Width           =   3495
   End
   Begin VB.Label txtNUMBER 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000014&
      Height          =   495
      Left            =   120
      TabIndex        =   29
      Top             =   120
      Width           =   3495
   End
End
Attribute VB_Name = "Calculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim mfirst As Double
Dim msecond As Double
Dim manswer As Double
Dim mbutton As Integer
Dim Signstate As Boolean
Const PI As Double = 3.14159265359
Private Sub cmd0_Click()
txtNUMBER = txtNUMBER + cmd0.Caption
Result = Result + "0"
End Sub
Private Sub cmd1_Click()
txtNUMBER = txtNUMBER + cmd1.Caption
Result = Result + "1"
End Sub
Private Sub cmd2_Click()
txtNUMBER = txtNUMBER + cmd2.Caption
Result = Result + "2"
End Sub
Private Sub cmd3_Click()
txtNUMBER = txtNUMBER + cmd3.Caption
Result = Result + "3"
End Sub
Private Sub cmd4_Click()
txtNUMBER = txtNUMBER + cmd4.Caption
Result = Result + "4"
End Sub
Private Sub cmd5_Click()
txtNUMBER = txtNUMBER + cmd5.Caption
Result = Result + "5"
End Sub
Private Sub cmd6_Click()
txtNUMBER = txtNUMBER + cmd6.Caption
Result = Result + "6"
End Sub
Private Sub cmd7_Click()
txtNUMBER = txtNUMBER + cmd7.Caption
Result = Result + "7"
End Sub
Private Sub cmd8_Click()
txtNUMBER = txtNUMBER + cmd8.Caption
Result = Result + "8"
End Sub
Private Sub cmd9_Click()
txtNUMBER = txtNUMBER + cmd9.Caption
Result = Result + "9"
End Sub
Private Sub cmdADD_Click()
mbutton = 1
mfirst = Val(Result)
txtNUMBER = txtNUMBER + cmdADD.Caption
Result = ""
End Sub
Private Sub cmdCos_Click()
 mfirst = Val(Result)
 msecond = Cos(mfirst * PI / 180)
 Result = msecond
End Sub
Private Sub cmdLog_Click()
 mfirst = Val(Result)
 msecond = Log(mfirst)
 Result = msecond
End Sub
Private Sub cmdMod_Click()
mbutton = 5
mfirst = Val(Result)
txtNUMBER = txtNUMBER + cmdMod.Caption
Result = ""
End Sub
Private Sub cmdPercent_Click()
Result.Caption = Result * (Val(Result.Caption) / 100)
End Sub
Private Sub cmdPOW_Click()
 mfirst = Val(Result)
 msecond = mfirst * mfirst
 Result = msecond
End Sub
Private Sub cmdSin_Click()
 mfirst = Val(Result)
 msecond = Sin(mfirst * PI / 180)
 Result = msecond
End Sub
Private Sub cmdSUBTRACT_Click()
mbutton = 2
mfirst = Val(Result)
txtNUMBER = txtNUMBER + cmdSUBTRACT.Caption
Result = ""
End Sub
Private Sub cmdMULTIPLY_Click()
mbutton = 3
mfirst = Val(Result)
txtNUMBER = txtNUMBER + cmdMULTIPLY.Caption
Result = ""
End Sub
Private Sub cmdDIVIDE_Click()
mbutton = 4
mfirst = Val(Result)
txtNUMBER = txtNUMBER + cmdDIVIDE.Caption
Result = ""
End Sub
Private Sub cmdEQUALS_Click()
msecond = Val(Result)
Select Case mbutton
Case Is = 1
manswer = mfirst + msecond
Case Is = 2
manswer = mfirst - msecond
Case Is = 3
manswer = mfirst * msecond
Case Is = 4
manswer = mfirst / msecond
Case Is = 5
manswer = mfirst Mod msecond
End Select
Result = manswer
End Sub
Private Sub cmdDOT_Click()
txtNUMBER = txtNUMBER + cmdDOT.Caption
Result = Result + "."
End Sub
Private Sub cmdSIGN_Click()
If txtNUMBER = "-" + txtNUMBER Then
MsgBox "error start again"
End If
If Signstate = False Then
txtNUMBER = "-" + txtNUMBER
Signstate = True
Else
minusvalue = Val(txtNUMBER)
minusvalue = Val("-1" * minusvalue)
txtNUMBER = minusvalue
Signstate = False
End If
End Sub
Private Sub cmdEXIT_Click()
Unload Calculator
End Sub
Private Sub cmdCANCEL_Click()
Result = ""
txtNUMBER = ""
End Sub
Private Sub cmdSqrt_Click()
Result = Sqr(Val(Result))
End Sub

Private Sub cmdinverse_Click()
Result = 1 / Val(Result)
End Sub
Private Sub cmdBackspace_Click()
If Len(txtNUMBER.Caption) = 0 Then Exit Sub
txtNUMBER.Caption = Mid(txtNUMBER.Caption, 1, Len(txtNUMBER.Caption) - 1)
End Sub
Private Sub cmdTan_Click()
 mfirst = Val(Result)
 msecond = Tan(mfirst * PI / 180)
 Result = msecond
End Sub

