VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   2730
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5835
   LinkTopic       =   "Form2"
   ScaleHeight     =   2730
   ScaleWidth      =   5835
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   2520
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   1
      Top             =   120
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type TVar
    Value As Variant
End Type

'https://www.vbarchiv.net/workshop/workshop_84-der-gebrauch-des-datentyps-decimal.html
Private Type TDec
    vt      As Integer ' 2  0  &HE = 14
    nnk     As Byte    ' 1  2  Anzahl der Nachkommastellen
    sgn     As Byte    ' 1  3  Vorzeichen
    LngVal3 As Long    ' 4  4  hih part
    LngVal1 As Long    ' 4  8  low part
    LngVal2 As Long    ' 4 12  mid part
End Type

Private Function TDec_ToHex(Value As TDec) As String
    With Value
        TDec_ToHex = "vt : " & Hex4(.vt) & vbCrLf & _
                     "nnk: " & Hex2(.nnk) & vbCrLf & _
                     "sgn: " & Hex2(.sgn) & vbCrLf & _
                     "lv3: " & Hex8(.LngVal3) & vbCrLf & _
                     "lv1: " & Hex8(.LngVal1) & vbCrLf & _
                     "lv2: " & Hex8(.LngVal2)
    End With
End Function

Private Sub Command2_Click()
    Dim tvv As TVar
    Dim ti8 As TDec
    
    With ti8
        .vt = 14
        .nnk = 0
        .sgn = 0
        '
        '26021645401728484450406541978
        '541499BF CF60F160 5CA4AA9A
        
        .LngVal1 = &H5CA4AA9A 'CAFEBABE
        .LngVal2 = &HCF60F160 'CAFEBABE
        .LngVal3 = &H541499BF 'CAFEBABE
    End With
    
    RtlMoveMemory tvv, ti8, 16
    
    Dim s As String: s = CStr(tvv.Value)
    Text1.Text = s
    '&HCAFEBABECAFEBABECAFEBABE = 628239210217683 41419963890366
    '&HFFFFFFFFFFFFFFFFFFFFFFFF = 79228162514264337593543950335
End Sub

Private Sub Form_Load()
    '
End Sub

Private Sub Command1_Click()
    Dim tvv As TVar
    Dim ti8 As TDec
    
    'tvv.Value = CDec("1")
    'tvv.Value = CDec("-1")
    'tvv.Value = CDec("2147483647")
    'tvv.Value = CDec("&H90807060504030201")
    'tvv.Value = CDec("2147483648")
    'tvv.Value = CDec("4294967294")
    'tvv.Value = CDec("4294967295")
    tvv.Value = CDec("9223372036854775807")
    'tvv.Value = CDec("9223372036854775808")
    'tvv.Value = CDec("18446744073709551615")
    'tvv.Value = CDec("18446744073709551616")
    'tvv.Value = CDec("79228162514264337593543950335") '* CDec("1024") ' + CDec("1023")
    '541499BF CF60F160 5CA4AA9A
    
    RtlMoveMemory ti8, tvv, 16
    
    Dim s As String: s = TDec_ToHex(ti8)
    'MsgBox "Biggest 96-bit unsigned value: " & vbCrLf & TDec_ToHex(ti8) & vbCrLf & CStr(tvv.Value)
'    With ti8
'        s = s & "vbDecimal: " & CStr(.vt) & " = &H" & Hex(.vt) & vbCrLf
'        s = s & "n Nkst   : " & CStr(.nnk) & vbCrLf
'        s = s & "Vorz     : " & CStr(.sgn) & vbCrLf
'        s = s & "LngVal3  : " & CStr(.LngVal3) & vbCrLf
'        s = s & "LngVal1  : " & CStr(.LngVal1) & vbCrLf
'        s = s & "LngVal2  : " & CStr(.LngVal2) & vbCrLf
'    End With
    Text1.Text = s

End Sub
