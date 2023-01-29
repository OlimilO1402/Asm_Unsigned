VERSION 5.00
Begin VB.Form From1 
   Caption         =   "Demo für vb64Bit.dll Funktionen"
   ClientHeight    =   7710
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10515
   LinkTopic       =   "Form1"
   ScaleHeight     =   7710
   ScaleWidth      =   10515
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton cmdTest 
      Caption         =   "Test"
      Height          =   495
      Left            =   8160
      TabIndex        =   57
      Top             =   5880
      Width           =   1695
   End
   Begin VB.CommandButton cmdCallObjProc 
      Caption         =   "CallObjProc"
      Height          =   615
      Left            =   8160
      TabIndex        =   56
      Top             =   5040
      Width           =   1695
   End
   Begin VB.CommandButton cmdCallProc 
      Caption         =   "CallProc"
      Height          =   615
      Left            =   8160
      TabIndex        =   55
      Top             =   4200
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   8160
      TabIndex        =   54
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Frame Frame4 
      Caption         =   "Funktionen mit Operand 1"
      Height          =   1935
      Left            =   2640
      TabIndex        =   28
      Top             =   3720
      Width           =   5175
      Begin VB.CommandButton cmdSQR 
         Caption         =   "SQR"
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
         Left            =   240
         TabIndex        =   35
         ToolTipText     =   "Square Root"
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox txtBits 
         Alignment       =   1  'Rechts
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
         Left            =   2280
         TabIndex        =   33
         Text            =   "4"
         Top             =   960
         Width           =   855
      End
      Begin VB.CommandButton cmdNeg 
         Caption         =   "NEG"
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
         Left            =   240
         TabIndex        =   32
         ToolTipText     =   "negate, Zweierkomplement"
         Top             =   480
         Width           =   855
      End
      Begin VB.CommandButton cmdSHR 
         Caption         =   "SHR"
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
         Left            =   1320
         TabIndex        =   31
         ToolTipText     =   "Shift Right"
         Top             =   480
         Width           =   855
      End
      Begin VB.CommandButton cmdSHL 
         Caption         =   "SHL"
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
         Left            =   1320
         TabIndex        =   30
         ToolTipText     =   "Shift Left"
         Top             =   840
         Width           =   855
      End
      Begin VB.CommandButton cmdNOT 
         Caption         =   "NOT"
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
         Left            =   240
         TabIndex        =   29
         ToolTipText     =   "NOT, Einerkomplement"
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "Bits to Shift"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         TabIndex        =   34
         Top             =   600
         Width           =   975
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Ergebnis"
      Height          =   1695
      Left            =   120
      TabIndex        =   21
      Top             =   6000
      Width           =   7695
      Begin VB.TextBox txtBinOP3 
         Alignment       =   1  'Rechts
         Enabled         =   0   'False
         Height          =   285
         Left            =   720
         TabIndex        =   46
         Top             =   1200
         Width           =   6855
      End
      Begin VB.TextBox txtStrOP3 
         Alignment       =   1  'Rechts
         Enabled         =   0   'False
         Height          =   285
         Left            =   3720
         TabIndex        =   41
         Top             =   480
         Width           =   2295
      End
      Begin VB.TextBox txtResultHi 
         Alignment       =   1  'Rechts
         Enabled         =   0   'False
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
         Left            =   720
         TabIndex        =   25
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox txtResultLo 
         Alignment       =   1  'Rechts
         Enabled         =   0   'False
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
         Left            =   2160
         TabIndex        =   24
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox txtResultLo_HEX 
         Alignment       =   1  'Rechts
         Enabled         =   0   'False
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
         Left            =   2160
         TabIndex        =   23
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox txtResultHi_HEX 
         Alignment       =   1  'Rechts
         Enabled         =   0   'False
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
         Left            =   720
         TabIndex        =   22
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Rechts
         Caption         =   "String"
         Height          =   255
         Left            =   4800
         TabIndex        =   48
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Rechts
         Caption         =   "Lo"
         Height          =   255
         Left            =   3120
         TabIndex        =   47
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Rechts
         Caption         =   "bin"
         Height          =   255
         Left            =   240
         TabIndex        =   45
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Rechts
         Caption         =   "Hi"
         Height          =   255
         Left            =   1680
         TabIndex        =   44
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Rechts
         Caption         =   "Dez"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Rechts
         Caption         =   "HEX"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   840
         Width           =   495
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Funktionen mit 2 Operanden"
      Height          =   1935
      Left            =   120
      TabIndex        =   15
      Top             =   3720
      Width           =   2415
      Begin VB.CommandButton cmdsMod 
         Caption         =   "sMod"
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
         Left            =   840
         TabIndex        =   52
         ToolTipText     =   "signed Modulo"
         Top             =   1440
         Width           =   855
      End
      Begin VB.CommandButton cmduMod 
         Caption         =   "uMod"
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
         Left            =   840
         TabIndex        =   51
         ToolTipText     =   "unsigned Modulo"
         Top             =   1080
         Width           =   855
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "+"
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
         Left            =   240
         TabIndex        =   20
         ToolTipText     =   "Addition"
         Top             =   360
         Width           =   375
      End
      Begin VB.CommandButton cmdSub 
         Caption         =   "-"
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
         Left            =   240
         TabIndex        =   19
         ToolTipText     =   "Subtraction"
         Top             =   720
         Width           =   375
      End
      Begin VB.CommandButton cmdMul 
         Caption         =   "*"
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
         Left            =   240
         TabIndex        =   18
         ToolTipText     =   "Multiplication, signed/unsigned"
         Top             =   1080
         Width           =   375
      End
      Begin VB.CommandButton cmduDiv 
         Caption         =   "uDiv"
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
         Left            =   840
         TabIndex        =   17
         ToolTipText     =   "unsigned Division"
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton cmdsDiv 
         Caption         =   "sDiv"
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
         Left            =   840
         TabIndex        =   16
         ToolTipText     =   "signed Division"
         Top             =   720
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Input"
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7695
      Begin VB.TextBox txtBinOp2 
         Alignment       =   1  'Rechts
         Enabled         =   0   'False
         Height          =   285
         Left            =   720
         TabIndex        =   49
         Top             =   2880
         Width           =   6855
      End
      Begin VB.TextBox txtBinOP1 
         Alignment       =   1  'Rechts
         Height          =   285
         Left            =   720
         TabIndex        =   42
         Top             =   1320
         Width           =   6855
      End
      Begin VB.TextBox txtStrOP2 
         Alignment       =   1  'Rechts
         Height          =   285
         Left            =   3840
         TabIndex        =   40
         Top             =   2160
         Width           =   2175
      End
      Begin VB.TextBox txtStrOP1 
         Alignment       =   1  'Rechts
         Height          =   285
         Left            =   3840
         TabIndex        =   38
         Text            =   "9223372028281618431"
         Top             =   600
         Width           =   2175
      End
      Begin VB.TextBox txtOp2Hi 
         Alignment       =   1  'Rechts
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
         Left            =   720
         TabIndex        =   11
         Text            =   "0"
         Top             =   2160
         Width           =   1335
      End
      Begin VB.TextBox txtOp2Lo 
         Alignment       =   1  'Rechts
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
         Left            =   2160
         TabIndex        =   10
         Text            =   "3"
         Top             =   2160
         Width           =   1335
      End
      Begin VB.TextBox txtOp2Hi_HEX 
         Alignment       =   1  'Rechts
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
         Left            =   720
         TabIndex        =   9
         Top             =   2520
         Width           =   1335
      End
      Begin VB.TextBox txtOp2Lo_HEX 
         Alignment       =   1  'Rechts
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
         Left            =   2160
         TabIndex        =   8
         Top             =   2520
         Width           =   1335
      End
      Begin VB.TextBox txtOp1Hi 
         Alignment       =   1  'Rechts
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
         Left            =   720
         TabIndex        =   4
         Text            =   "0"
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox txtOp1Lo 
         Alignment       =   1  'Rechts
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
         Left            =   2160
         TabIndex        =   3
         Text            =   "2"
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox txtOp1Hi_HEX 
         Alignment       =   1  'Rechts
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
         Left            =   720
         TabIndex        =   2
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox txtOp1Lo_HEX 
         Alignment       =   1  'Rechts
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
         Left            =   2160
         TabIndex        =   1
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Rechts
         Caption         =   "Bin"
         Height          =   255
         Left            =   240
         TabIndex        =   50
         Top             =   2880
         Width           =   375
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Rechts
         Caption         =   "Bin"
         Height          =   255
         Left            =   240
         TabIndex        =   43
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Rechts
         Caption         =   "String"
         Height          =   255
         Left            =   4920
         TabIndex        =   39
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Rechts
         Caption         =   "Hi"
         Height          =   255
         Left            =   1680
         TabIndex        =   37
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Rechts
         Caption         =   "Lo"
         Height          =   255
         Left            =   3120
         TabIndex        =   36
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Rechts
         Caption         =   "Dez"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   2160
         Width           =   495
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Rechts
         Caption         =   "HEX"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   2520
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "Operand 2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   12
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Rechts
         Caption         =   "Dez"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Operand 1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Rechts
         Caption         =   "HEX"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   495
      End
   End
   Begin VB.Label Label9 
      Caption         =   "Die Umwandlung in Binärstring hat leider noch Fehler! So wird z.B. ein '-' Zeichen ausgegeben statt das Vorzeichenbit gesetzt."
      Height          =   2415
      Left            =   8040
      TabIndex        =   53
      Top             =   480
      Width           =   2055
   End
End
Attribute VB_Name = "From1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

' 64 Bit Operanden definieren
Dim OP1 As Large     ' Operand 1
Dim OP2 As Large     ' Operand 2
Dim OP3 As Large     ' Operand 3, Ergebnis-Operand

Dim BitsToShift As Long
Private Type ParametrListe
   Par1 As Long
   Par2 As Long
   Par3 As Long
End Type

Private Const BinFormat$ = "0000 0000 0000 0000 - 0000 0000 0000 0000 - 0000 0000 0000 0000 - 0000 0000 0000 0000"
Private Const sign& = 0

Private Declare Function Test Lib "C:\masm32\Project\vbExt\vbExt.dll" _
         (ByVal B1 As Byte, B2 As Byte) As Long

Private Sub cmdCallObjProc_Click()
   Dim CallObj As New cTestCall
   Dim ptrObj As Long
   Dim ParLst As ParametrListe
   Dim ret As Long
   
   With ParLst
      .Par1 = 10
      .Par2 = 20
      .Par3 = 30
   End With

   ptrObj = ObjPtr(CallObj)
   
   ret = CallObjProc(ParLst, 3&, ptrObj, 0&)
   Debug.Print "Ergebnis des Funktionsaufrufes über CallBack = "; ret
End Sub

Private Sub cmdCallProc_Click()
   Dim ParLst As ParametrListe
   Dim ptrParLst As Long
   Dim pProc As Long
   Dim ret As Long
   
   With ParLst
      .Par1 = 10
      .Par2 = 20
      .Par3 = 30
   End With
   
   ptrParLst = VarPtr(ParLst)
   
   pProc = VbStackParVal(AddressOf TestCall)
   Debug.Print "pProc : ", pProc
   ret = CallProc(ParLst, 3, pProc)
   Debug.Print "Result : "; ret
End Sub

Private Sub cmdTest_Click()
   Dim ts As TestStrukt
   Dim retStrukt As TestStrukt
   Dim ptrRetStrukt As Long
   Dim ret As Long
   Dim ptrProc As Long
   
   ts.p1 = 1
   ts.p2 = 2
   ts.p3 = 3
   ts.p4 = 4
   
   ptrRetStrukt = VarPtr(retStrukt)
   ptrProc = VbStackParVal(AddressOf CreateStrukt)
   ret = CallProc(ts, 4, ptrProc, ptrRetStrukt)
   Main
   With retStrukt
      Debug.Print "ReturnStrukt"
      Debug.Print "ret= "; ret, ptrRetStrukt
      Debug.Print "p1 = "; .p1
      Debug.Print "p2 = "; .p2
      Debug.Print "p3 = "; .p3
      Debug.Print "p4 = "; .p4
      Debug.Print
   End With
End Sub

Private Sub Command1_Click()
   Dim ret As Long
   ret = VbTest(1&, 2&)
   Debug.Print "Stacktest : "; ret
End Sub

Private Sub Test_IsArryDim()
   Dim ar() As Byte
   Dim ret As Boolean
   ret = IsArrayDim(ar())
   Debug.Print "dim : "; ret
   
   ReDim ar(10)
   ret = IsArrayDim(ar())
   Debug.Print "dim : "; ret
   
End Sub

Private Sub Form_Load()
   txtOp1Hi_Validate False
   txtOp1Lo_Validate False
   txtOp2Hi_Validate False
   txtOp2Lo_Validate False
   txtBits_Validate False
End Sub

' ===========================================================================
'  64Bit Arithmetikfunktionen aufrufen
' ===========================================================================

Private Sub cmdAdd_Click()
   OP3 = Add64(OP1, OP2)
   ShowResult
End Sub

Private Sub cmdSub_Click()
   OP3 = Sub64(OP1, OP2)
   ShowResult
End Sub

Private Sub cmdMul_Click()
   OP3 = Mul64(OP1, OP2)
   'TestProc OP1, OP2, OP3
   ShowResult
End Sub

Private Sub cmduDiv_Click()
   OP3 = Div64(OP1, OP2, False)
   ShowResult
End Sub

Private Sub cmdsDiv_Click()
   OP3 = Div64(OP1, OP2)
   ShowResult
End Sub

Private Sub cmduMod_Click()
   OP3 = Mod64(OP1, OP2, False)
   ShowResult
End Sub

Private Sub cmdsMod_Click()
   OP3 = Mod64(OP1, OP2)
   ShowResult
End Sub

Private Sub cmdSQR_Click()
   OP3 = Sqr64(OP1)
   ShowResult
End Sub

Private Sub cmdNeg_Click()
   OP3 = Neg64(OP1)
   ShowResult
End Sub

Private Sub cmdNOT_Click()
   OP3 = Not64(OP1)
   ShowResult
End Sub

Private Sub cmdSHR_Click()
   OP3 = Shr64(OP1, BitsToShift)
   ShowResult
End Sub

Private Sub cmdSHL_Click()
   OP3 = Shl64(OP1, BitsToShift)
   ShowResult
End Sub



' ===========================================================================
'  Werteübernahme bei Eingabe in Dezimalfeldern
' ===========================================================================

Private Sub txtOp1Hi_Validate(Cancel As Boolean)
   OP1.Hi = CLng(txtOp1Hi)
   txtOp1Hi_HEX = Hex$(OP1.Hi)
   txtStrOP1.Text = LargeToStr(OP1)
   txtBinOP1 = Format(LargeToStr(OP1, 2, sign), BinFormat)
End Sub

Private Sub txtOp1Lo_Validate(Cancel As Boolean)
   OP1.Lo = CLng(txtOp1Lo)
   txtOp1Lo_HEX = Hex$(OP1.Lo)
   txtStrOP1.Text = LargeToStr(OP1)
   txtBinOP1 = Format(LargeToStr(OP1, 2, sign), BinFormat)
End Sub

Private Sub txtOp2Hi_Validate(Cancel As Boolean)
   OP2.Hi = CLng(txtOp2Hi)
   txtOp2Hi_HEX = Hex$(OP2.Hi)
   txtStrOP2.Text = LargeToStr(OP2)
   txtBinOp2 = Format(LargeToStr(OP2, 2, sign), BinFormat)
End Sub

Private Sub txtOp2Lo_Validate(Cancel As Boolean)
   OP2.Lo = CLng(txtOp2Lo)
   txtOp2Lo_HEX = Hex$(OP2.Lo)
   txtStrOP2.Text = LargeToStr(OP2)
   txtBinOp2 = Format(LargeToStr(OP2, 2, sign), BinFormat)
End Sub

' ===========================================================================
'  Werteübernahme bei Eingabe in HEX-Feldern
' ===========================================================================

Private Sub txtOp1Hi_HEX_Validate(Cancel As Boolean)
   Dim H As Long
   OP1.Hi = CLng("&H" & txtOp1Hi_HEX)
   txtOp1Hi = OP1.Hi
   txtStrOP1.Text = LargeToStr(OP1)
   txtBinOP1 = Format(LargeToStr(OP1, 2, sign), BinFormat)
End Sub

Private Sub txtOp1Lo_HEX_Validate(Cancel As Boolean)
   Dim H As Long
   OP1.Lo = CLng("&H" & txtOp1Lo_HEX)
   txtOp1Lo = OP1.Lo
   txtStrOP1.Text = LargeToStr(OP1)
   txtBinOP1 = Format(LargeToStr(OP1, 2, sign), BinFormat)
End Sub

Private Sub txtOp2Hi_HEX_Validate(Cancel As Boolean)
   Dim H As Long
   OP2.Hi = CLng("&H" & txtOp2Hi_HEX)
   txtOp2Hi = OP2.Hi
   txtStrOP2.Text = LargeToStr(OP2)
   txtBinOp2 = Format(LargeToStr(OP2, 2, sign), BinFormat)
End Sub

Private Sub txtOp2Lo_HEX_Validate(Cancel As Boolean)
   Dim H As Long
   OP2.Lo = CLng("&H" & txtOp2Lo_HEX)
   txtOp2Lo = OP2.Lo
   txtStrOP2.Text = LargeToStr(OP2)
   txtBinOp2 = Format(LargeToStr(OP2, 2, sign), BinFormat)
End Sub

' ===========================================================================
'  Werteübernahme bei Eingabe in String-Feldern
' ===========================================================================

Private Sub txtStrOP1_Validate(Cancel As Boolean)
   OP1 = StrToLarge(txtStrOP1.Text)
   txtOp1Lo = OP1.Lo
   txtOp1Hi = OP1.Hi
   txtOp1Lo_HEX = Hex$(OP1.Lo)
   txtOp1Hi_HEX = Hex$(OP1.Hi)
   txtBinOP1 = Format(LargeToStr(OP1, 2, sign), BinFormat)
End Sub

Private Sub txtStrOP2_Validate(Cancel As Boolean)
   OP2 = StrToLarge(txtStrOP2.Text)
   txtOp2Lo = OP2.Lo
   txtOp2Hi = OP2.Hi
   txtOp2Lo_HEX = Hex$(OP2.Lo)
   txtOp2Hi_HEX = Hex$(OP2.Hi)
   txtBinOp2 = Format(LargeToStr(OP2, 2, sign), BinFormat)
End Sub

' ===========================================================================
'  BitsToShift
' ===========================================================================

Private Sub txtBits_Validate(Cancel As Boolean)
   BitsToShift = CLng(txtBits)
   Debug.Print "BitsToShift = "; BitsToShift
End Sub

' ===========================================================================
'  Ergebnis in Ausgabefeldern anzeigen
' ===========================================================================

Private Sub ShowResult()
   txtResultHi = CStr(OP3.Hi)
   txtResultLo = CStr(OP3.Lo)
   txtResultHi_HEX = Hex$(OP3.Hi)
   txtResultLo_HEX = Hex$(OP3.Lo)
   txtStrOP3 = LargeToStr(OP3)
   txtBinOP3 = Format(LargeToStr(OP3, 2, sign), BinFormat)
End Sub


