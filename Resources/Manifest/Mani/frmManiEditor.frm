VERSION 5.00
Begin VB.Form frmManiEditor 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "PE-Manifestdatei-Editor"
   ClientHeight    =   4290
   ClientLeft      =   45
   ClientTop       =   450
   ClientWidth     =   6765
   Icon            =   "frmManiEditor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   6765
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton cmdNeu 
      Caption         =   "&Neu"
      Enabled         =   0   'False
      Height          =   465
      Left            =   1980
      TabIndex        =   13
      Top             =   3780
      Width           =   1185
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'Kein
      Height          =   1455
      Left            =   45
      ScaleHeight     =   1455
      ScaleWidth      =   6675
      TabIndex        =   2
      Top             =   2160
      Width           =   6675
      Begin VB.CheckBox chkStart 
         Caption         =   " Anwendung starten"
         Height          =   240
         Left            =   45
         TabIndex        =   14
         Top             =   945
         Visible         =   0   'False
         Width           =   1905
      End
      Begin VB.TextBox txtFile 
         Height          =   285
         Left            =   45
         TabIndex        =   12
         Top             =   578
         Width           =   6000
      End
      Begin VB.CommandButton cmdCommon 
         Caption         =   "..."
         Height          =   360
         Left            =   6165
         TabIndex        =   6
         Top             =   540
         Width           =   420
      End
      Begin VB.Label lblStatus 
         Caption         =   "Keine Datei ausgewählt"
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   675
         TabIndex        =   11
         Top             =   1215
         Width           =   1770
      End
      Begin VB.Label lblGrund 
         Caption         =   "Status:"
         Height          =   240
         Left            =   45
         TabIndex        =   10
         Top             =   1215
         Width           =   555
      End
      Begin VB.Label LblSchrit 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   45
         TabIndex        =   5
         Top             =   45
         Width           =   3885
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         Index           =   0
         X1              =   45
         X2              =   6570
         Y1              =   360
         Y2              =   360
      End
   End
   Begin VB.PictureBox picInfo 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'Kein
      Height          =   2040
      Left            =   22
      ScaleHeight     =   2040
      ScaleWidth      =   6720
      TabIndex        =   3
      Top             =   0
      Width           =   6720
      Begin VB.Image imgIcon 
         Height          =   555
         Left            =   6165
         Top             =   90
         Width           =   555
      End
      Begin VB.Label lblInfo 
         Appearance      =   0  '2D
         BackColor       =   &H0000FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Private Sub Form_Initialize()          Call InitCommonControls End Sub"
         ForeColor       =   &H00C00000&
         Height          =   645
         Index           =   3
         Left            =   135
         TabIndex        =   9
         Top             =   1350
         Width           =   2085
      End
      Begin VB.Label lblInfo 
         Appearance      =   0  '2D
         BackColor       =   &H0000FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Private Declare Sub InitCommonControls Lib ""comctl32"" ()"
         ForeColor       =   &H00C00000&
         Height          =   240
         Index           =   2
         Left            =   135
         TabIndex        =   8
         Top             =   1080
         Width           =   6135
      End
      Begin VB.Label lblInfo 
         Appearance      =   0  '2D
         BackColor       =   &H0000FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Folgende Code muss in Ihrem Projekt eingegeben werden:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   1
         Left            =   45
         TabIndex        =   7
         Top             =   735
         Width           =   6630
      End
      Begin VB.Label lblInfo 
         Appearance      =   0  '2D
         BackColor       =   &H0000FFFF&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   645
         Index           =   0
         Left            =   45
         TabIndex        =   4
         Top             =   45
         Width           =   6090
      End
   End
   Begin VB.CommandButton cbdCancel 
      Caption         =   "&Beenden"
      Height          =   465
      Left            =   5535
      TabIndex        =   1
      Top             =   3780
      Width           =   1185
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "&Manifest-Datei erstellen"
      Enabled         =   0   'False
      Height          =   465
      Left            =   3285
      TabIndex        =   0
      Top             =   3780
      Width           =   2175
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Index           =   1
      X1              =   45
      X2              =   6705
      Y1              =   3690
      Y2              =   3690
   End
End
Attribute VB_Name = "frmManiEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetFileVersionInfo Lib _
        "Version.dll" Alias "GetFileVersionInfoA" _
        (ByVal lptstrFilename As String, ByVal dwhandle _
        As Long, ByVal dwlen As Long, lpData As Any) _
        As Long

Private Declare Function GetFileVersionInfoSize Lib _
        "Version.dll" Alias "GetFileVersionInfoSizeA" _
        (ByVal lptstrFilename As String, lpdwHandle As _
        Long) As Long

Private Declare Function VerQueryValue Lib "Version.dll" _
        Alias "VerQueryValueA" (pBlock As Any, ByVal _
        lpSubBlock As String, lplpBuffer As Any, puLen _
        As Long) As Long

Private Declare Sub MoveMemory Lib "kernel32" Alias _
        "RtlMoveMemory" (dest As Any, ByVal Source As _
        Long, ByVal length As Long)
        
Private Declare Function GetOpenFileName Lib "comdlg32.dll" _
        Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
        
Private Declare Sub InitCommonControls Lib "comctl32.dll" ()

'Prüfen ob Datei existiert
Private Declare Function GetFileAttributes Lib "kernel32" Alias _
        "GetFileAttributesA" (ByVal lpFileName As String) As Long

Private Type OPENFILENAME
lStructSize As Long
hwndOwner As Long
hInstance As Long
lpstrFilter As String
lpstrCustomFilter As String
nMaxCustFilter As Long
nFilterIndex As Long
lpstrFile As String
nMaxFile As Long
lpstrFileTitle As String
nMaxFileTitle As Long
lpstrInitialDir As String
lpstrTitle As String
flags As Long
nFileOffset As Integer
nFileExtension As Integer
lpstrDefExt As String
lCustData As Long
lpfnHook As Long
lpTemplateName As String
End Type

Private Type VS_FIXEDFILEINFO
  dwSignature As Long
  dwStrucVersionl As Integer
  dwStrucVersionh As Integer
  dwFileVersionMSl As Integer
  dwFileVersionMSh As Integer
  dwFileVersionLSl As Integer
  dwFileVersionLSh As Integer
  dwProductVersionMSl As Integer
  dwProductVersionMSh As Integer
  dwProductVersionLSl As Integer
  dwProductVersionLSh As Integer
  dwFileFlagsMask As Long
  dwFileFlags As Long
  dwFileOS As Long
  dwFileType As Long
  dwFileSubtype As Long
  dwFileDateMS As Long
  dwFileDateLS As Long
End Type

Dim strFileName As String

Private Sub cmdNeu_Click()
txtFile.Text = ""
strFileName = ""
cmdNext.Enabled = False
lblStatus.ForeColor = vbRed
lblGrund.Caption = "Status:"
LblSchrit.Caption = " Erste Schritt: Ausführbare Datei auswählen "
lblStatus.Caption = " Keine Datei ausgewählt "
chkStart.Visible = False
cmdNeu.Enabled = False
End Sub

Private Sub cmdNext_Click()
Open txtFile.Text & ".manifest" For Output As #1
    Print #1, "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "UTF-8" & Chr(34) & " standalone=" & Chr(34) & "yes" & Chr(34) & "?>"
    Print #1, "<assembly xmlns=" & Chr(34) & "urn:schemas-microsoft-com:asm.v1" & Chr(34) & " manifestVersion=" & Chr(34) & "1.0" & Chr(34) & ">"
    Print #1, "     <assemblyIdentity"
    Print #1, "         name=" & Chr(34) & strFileName & Chr(34)
    Print #1, "         processorArchitecture=" & Chr(34) & "x86" & Chr(34)
    Print #1, "         version=" & Chr(34) & lblStatus.Caption & Chr(34)
    Print #1, "         type=" & Chr(34) & "win32" & Chr(34) & "/>"
    Print #1, "  <description>Created using PE-Manifestdatei Editor</description>"
    Print #1, "  <dependency>"
    Print #1, "     <dependentAssembly>"
    Print #1, "         <assemblyIdentity"
    Print #1, "         type=" & Chr(34) & "win32" & Chr(34)
    Print #1, "         name=" & Chr(34) & "Microsoft.Windows.Common-Controls" & Chr(34)
    Print #1, "         version=" & Chr(34) & "6.0.0.0" & Chr(34)
    Print #1, "         processorArchitecture=" & Chr(34) & "x86" & Chr(34)
    Print #1, "         publicKeyToken=" & Chr(34) & "6595b64144ccf1df" & Chr(34)
    Print #1, "         language=" & Chr(34) & "*" & Chr(34)
    Print #1, "      />"
    Print #1, "     </dependentAssembly>"
    Print #1, "  </dependency>"
    Print #1, "</assembly>"
Close #1
MsgBox "Manifestdatei für " & txtFile.Text & " ist erstellt!", _
    vbInformation, "XP Visual Style für Ihre Programm"
Call cmdNeu_Click
On Error Resume Next
If chkStart.Value = 1 Then
    Call Shell(txtFile.Text, vbNormalFocus)
End If
End Sub

Private Sub Form_Initialize()
  Call InitCommonControls
End Sub

Private Sub cbdCancel_Click()
Unload Me
End Sub

Private Sub cmdCommon_Click()
Dim Existens As Integer
Dim strDateiName As String
Dim Result As Boolean
Dim OFN As OPENFILENAME

OFN.lStructSize = Len(OFN)
OFN.hwndOwner = Me.hWnd
OFN.hInstance = App.hInstance
OFN.lpstrFilter = "Ausführbare Dateien (*.exe)" & Chr$(0) & "*.exe" & Chr$(0) & Chr$(0)
OFN.lpstrFile = Space$(254)
OFN.nMaxFile = 255
OFN.lpstrFileTitle = Space$(254)
OFN.nMaxFileTitle = 255
'OFN.lpstrInitialDir = CurDir
OFN.lpstrTitle = "Exe-Dateien suchen"
OFN.flags = 0

Result = GetOpenFileName(OFN)

If Result Then
    strDateiName = Left$(OFN.lpstrFile, InStr(OFN.lpstrFile, Chr$(0)) - 1)
    If fbFileExists(strDateiName) Then
        strFileName = OFN.lpstrFileTitle
        lblStatus.ForeColor = vbBlack
        LblSchrit.Caption = " Zweite Schritt: Manifest-Datei erstellen "
'        lblStatus.Caption = "Mit " & Chr(34) & " Weiter " & Chr(34) & " vortfahren"
        txtFile.Text = strDateiName
        Call DisplayVerInfo(strDateiName)
        lblGrund.Caption = "Version:"
        cmdNext.Enabled = True
        cmdNeu.Enabled = True
        chkStart.Visible = True
        cmdNext.SetFocus
    Else
        Existens = MsgBox("Datei " & strDateiName & " existiert nicht!" & _
            vbNewLine & "Möchten Sie eine andere Datei auswählen?", _
            vbYesNo + vbExclamation, "Datei existiert nicht")
        Select Case Existens
            Case vbYes
                Call cmdCommon_Click
        End Select
    End If
    Result = False
End If
End Sub

Private Sub DisplayVerInfo(ByVal FilePath$)
Dim FVer, Buff() As Byte
Dim l&, BuffL&, Pointer&, VERSION As VS_FIXEDFILEINFO
 
l = GetFileVersionInfoSize(FilePath, 0&)
If l < 1 Then
    MsgBox "Keine Versions-Info vorhanden!"
Else
    ReDim Buff(l)
    Call GetFileVersionInfo(FilePath, 0&, l, Buff(0))
    Call VerQueryValue(Buff(0), "\", Pointer, BuffL)
    Call MoveMemory(VERSION, Pointer, Len(VERSION))
      
    With VERSION
        FVer = Format$(.dwFileVersionMSh) & "." & _
                Format$(.dwFileVersionMSl) & "." & _
                Format$(.dwFileVersionLSh) & "." & _
                Format$(.dwFileVersionLSl)
    End With
End If
If Len(FVer) > 0 Then
    lblStatus.Caption = FVer
Else
    lblStatus.Caption = "1.0.0.0"
End If
       
End Sub
Private Function fbFileExists(spFilePath As String) As Boolean
    fbFileExists = (GetFileAttributes(spFilePath) <> -1)
End Function

Private Sub Form_Load()
imgIcon.Picture = Me.Icon
lblStatus.ForeColor = vbRed
LblSchrit.Caption = " Erste Schritt: Ausführbare Datei auswählen "
lblInfo(0).Caption = " Möchten Sie in Ihren, mit Visual Basic erstellten Programmen, Visual Styles von Windows XP benutzen, müssen Sie Anwendungsmanifest erstellen. Mit PE-Manifestdatei-Editor erstellen Sie diese Datei im Handumdrehen."
End Sub
