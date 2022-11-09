VERSION 5.00
Begin VB.Form Calculator 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Calculadora"
   ClientHeight    =   3270
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5475
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   5475
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Height          =   135
      Left            =   5520
      MaskColor       =   &H80000004&
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   480
      UseMaskColor    =   -1  'True
      Width           =   135
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   1
      Left            =   120
      MaskColor       =   &H00C00000&
      TabIndex        =   38
      Top             =   2235
      UseMaskColor    =   -1  'True
      Width           =   540
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   2
      Left            =   705
      MaskColor       =   &H00C00000&
      TabIndex        =   37
      Top             =   2235
      UseMaskColor    =   -1  'True
      Width           =   540
   End
   Begin VB.CommandButton cmdButton 
      Appearance      =   0  'Flat
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   3
      Left            =   1290
      MaskColor       =   &H00C00000&
      TabIndex        =   36
      Top             =   2235
      UseMaskColor    =   -1  'True
      Width           =   540
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   4
      Left            =   120
      MaskColor       =   &H00C00000&
      TabIndex        =   35
      Top             =   1755
      UseMaskColor    =   -1  'True
      Width           =   540
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   5
      Left            =   705
      MaskColor       =   &H00C00000&
      TabIndex        =   34
      Top             =   1755
      UseMaskColor    =   -1  'True
      Width           =   540
   End
   Begin VB.CommandButton cmdButton 
      Appearance      =   0  'Flat
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   6
      Left            =   1290
      MaskColor       =   &H00C00000&
      TabIndex        =   33
      Top             =   1755
      UseMaskColor    =   -1  'True
      Width           =   540
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   7
      Left            =   120
      MaskColor       =   &H00C00000&
      TabIndex        =   32
      Top             =   1275
      UseMaskColor    =   -1  'True
      Width           =   540
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   8
      Left            =   705
      MaskColor       =   &H00C00000&
      TabIndex        =   31
      Top             =   1275
      UseMaskColor    =   -1  'True
      Width           =   540
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   9
      Left            =   1290
      MaskColor       =   &H00C00000&
      TabIndex        =   30
      Top             =   1275
      UseMaskColor    =   -1  'True
      Width           =   540
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   0
      Left            =   120
      MaskColor       =   &H00C00000&
      TabIndex        =   29
      Top             =   2715
      UseMaskColor    =   -1  'True
      Width           =   540
   End
   Begin VB.CommandButton cmdChangeSign 
      Caption         =   "+/-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   705
      MaskColor       =   &H00C00000&
      TabIndex        =   28
      ToolTipText     =   "Change Sign"
      Top             =   2715
      UseMaskColor    =   -1  'True
      Width           =   540
   End
   Begin VB.CommandButton cmdDecimal 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1290
      MaskColor       =   &H00C00000&
      TabIndex        =   27
      Top             =   2715
      UseMaskColor    =   -1  'True
      Width           =   540
   End
   Begin VB.CommandButton cmdBasicFunctions 
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
      Height          =   435
      Index           =   1
      Left            =   1875
      TabIndex        =   26
      ToolTipText     =   "Add"
      Top             =   2715
      Width           =   540
   End
   Begin VB.CommandButton cmdBasicFunctions 
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
      Height          =   435
      Index           =   2
      Left            =   1875
      TabIndex        =   25
      ToolTipText     =   "Subtract"
      Top             =   2235
      Width           =   540
   End
   Begin VB.CommandButton cmdBasicFunctions 
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   3
      Left            =   1875
      TabIndex        =   24
      ToolTipText     =   "Multiplication"
      Top             =   1755
      Width           =   540
   End
   Begin VB.CommandButton cmdBasicFunctions 
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
      Height          =   435
      Index           =   4
      Left            =   1875
      TabIndex        =   23
      ToolTipText     =   "Division"
      Top             =   1275
      Width           =   540
   End
   Begin VB.CommandButton cmdBasicFunctions 
      Caption         =   "="
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   0
      Left            =   2460
      TabIndex        =   22
      Top             =   2715
      Width           =   540
   End
   Begin VB.CommandButton cmdSciFunctions 
      Caption         =   "x²"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   0
      Left            =   2460
      TabIndex        =   21
      ToolTipText     =   "Square Number"
      Top             =   2235
      Width           =   540
   End
   Begin VB.CommandButton cmdSciFunctions 
      Caption         =   "Sqr"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   1
      Left            =   2460
      TabIndex        =   20
      ToolTipText     =   "Square Route"
      Top             =   1755
      Width           =   540
   End
   Begin VB.CommandButton cmdSciFunctions 
      Caption         =   "x!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   2
      Left            =   2460
      TabIndex        =   19
      ToolTipText     =   "Factorial"
      Top             =   1275
      Width           =   540
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   120
      TabIndex        =   18
      Top             =   750
      Width           =   855
   End
   Begin VB.CommandButton cmdAc 
      Caption         =   "AC"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1065
      TabIndex        =   17
      Top             =   750
      Width           =   855
   End
   Begin VB.CommandButton cmdPi 
      Caption         =   " pi"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4215
      TabIndex        =   16
      ToolTipText     =   "3.141592654"
      Top             =   2715
      Width           =   540
   End
   Begin VB.CommandButton cmdMem 
      Caption         =   "STO"
      Height          =   435
      Index           =   0
      Left            =   3045
      TabIndex        =   15
      ToolTipText     =   "Store number into memory"
      Top             =   1275
      Width           =   540
   End
   Begin VB.CommandButton cmdMem 
      Caption         =   "RCL"
      Height          =   435
      Index           =   1
      Left            =   3630
      TabIndex        =   14
      ToolTipText     =   "Recall memory"
      Top             =   1275
      Width           =   540
   End
   Begin VB.CommandButton cmdMem 
      Caption         =   "MC"
      Height          =   435
      Index           =   2
      Left            =   4215
      TabIndex        =   13
      ToolTipText     =   "Clears Memory"
      Top             =   1275
      Width           =   540
   End
   Begin VB.CommandButton cmdTan 
      Caption         =   "tan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3045
      TabIndex        =   12
      Top             =   1755
      Width           =   540
   End
   Begin VB.CommandButton cmdSign 
      Caption         =   "sin"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3645
      TabIndex        =   11
      Top             =   1755
      Width           =   540
   End
   Begin VB.CommandButton cmdCos 
      Caption         =   "cos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4215
      TabIndex        =   10
      Top             =   1755
      Width           =   540
   End
   Begin VB.CommandButton cmdPercent 
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4215
      TabIndex        =   9
      Top             =   2235
      Width           =   540
   End
   Begin VB.CommandButton cmdnPr 
      Caption         =   "nPr"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3045
      TabIndex        =   8
      ToolTipText     =   "Permuntations"
      Top             =   2715
      Width           =   540
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   0
      Left            =   2010
      TabIndex        =   7
      Top             =   750
      Width           =   855
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "CE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   1
      Left            =   2955
      TabIndex        =   6
      Top             =   750
      Width           =   855
   End
   Begin VB.CommandButton cmdxy 
      Caption         =   "y^x"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3645
      TabIndex        =   5
      ToolTipText     =   "y^x"
      Top             =   2235
      Width           =   540
   End
   Begin VB.CommandButton cmdPCR 
      Caption         =   "nCr"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3630
      TabIndex        =   4
      ToolTipText     =   "Combination"
      Top             =   2715
      Width           =   540
   End
   Begin VB.CommandButton cmdX3 
      Caption         =   "x³"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3045
      TabIndex        =   3
      Top             =   2235
      Width           =   540
   End
   Begin VB.CommandButton cmdInverse 
      Caption         =   "1/x"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4800
      TabIndex        =   2
      ToolTipText     =   "Inverse"
      Top             =   1275
      Width           =   540
   End
   Begin VB.CommandButton cmdRandom 
      Caption         =   "Rnd"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4800
      TabIndex        =   1
      ToolTipText     =   "Random numbers between 0 and 1"
      Top             =   1755
      Width           =   540
   End
   Begin VB.CommandButton cmdMod 
      Caption         =   "Mod"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4800
      TabIndex        =   0
      ToolTipText     =   "Remainder "
      Top             =   2235
      Width           =   540
   End
   Begin VB.Label lblScreen 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   120
      TabIndex        =   42
      Top             =   120
      Width           =   3480
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Memoria:"
      Height          =   225
      Left            =   3690
      TabIndex        =   41
      Top             =   285
      Width           =   645
   End
   Begin VB.Label lblMem 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4425
      TabIndex        =   40
      Top             =   225
      Width           =   2790
   End
End
Attribute VB_Name = "Calculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim EraseNext As Boolean, Divide As Double, Times As Double, yX As Double
Const pi = 3.141592654
Dim Add As Double, codeMod As Double, nPr As Double
Dim Subtract As Double, nCr As Double

Private Sub cmdAc_Click()
On Error Resume Next
    lblScreen.Caption = ""
    Form_Load
    Add = 0
    Subtract = 0
    cmdBasicFunctions(0).SetFocus
End Sub

Private Sub cmdBasicFunctions_Click(Index As Integer)
    On Error Resume Next
    Call SubBasicFuntions(Index)
    
End Sub

Private Sub cmdBasicFunctions_KeyPress(Index As Integer, KeyAscii As Integer)
If (KeyAscii = 13 And FrmRegistroComprasList.Procedencia = Mbimponible) Then
    FrmRegistroComprasList.TxtValorcompra.Text = Format(Val(Me.lblScreen.Caption), "###0.00")
    Unload Me
End If
End Sub

Private Sub cmdButton_Click(Index As Integer)
On Error Resume Next
    Call PressKey(Index + 48)
    cmdBasicFunctions(0).SetFocus
End Sub

Private Sub cmdChangeSign_Click()
    On Error Resume Next
        If Len(lblScreen.Caption) = 0 Or lblScreen.Caption = "-" Then
        lblScreen.Caption = "-"
        Exit Sub
        End If
    lblScreen.Caption = lblScreen.Caption * -1
    cmdBasicFunctions(0).SetFocus
End Sub

Private Sub cmdClear_Click(Index As Integer)
    On Error Resume Next

    Select Case Index
    
    Case 0
    lblScreen.Caption = ""
    Times = 1
    Divide = 1
    
    Case 1
    lblScreen.Caption = ""
    End Select
    
    cmdBasicFunctions(0).SetFocus
End Sub

Private Sub cmdCos_Click()
    On Error Resume Next
    lblScreen.Caption = Round(Cos(Radians(Val(lblScreen.Caption))), 9)
    cmdBasicFunctions(0).SetFocus
End Sub

Private Sub cmdDecimal_Click()
    On Error Resume Next
    Dim i As Integer
        If EraseNext = True Then
        lblScreen.Caption = ""
        EraseNext = False
        End If
            For i = 1 To Len(lblScreen.Caption)
                If Mid(lblScreen.Caption, i, 1) = "." Then
                MsgBox ("ILLEGAL"), , "NYI"
                Exit Sub
                End If
            Next i
    lblScreen.Caption = lblScreen.Caption & "."
    cmdBasicFunctions(0).SetFocus
End Sub

Private Sub cmdInverse_Click()
    On Error Resume Next
    lblScreen.Caption = 1 / Val(lblScreen.Caption)
End Sub

Private Sub cmdMem_Click(Index As Integer)
    On Error Resume Next
    Select Case Index
    Case 0 'sto
    lblMem.Caption = lblScreen.Caption
    
    Case 1 'rcl
    lblScreen.Caption = lblMem.Caption
    lblMem.Caption = ""
    
    Case 2
    lblMem.Caption = "" 'memclear
    End Select
    cmdBasicFunctions(0).SetFocus
End Sub

Private Sub cmdMod_Click()
    EraseNext = 1
    codeMod = Val(lblScreen.Caption)
End Sub

Private Sub cmdnPr_Click()
    On Error Resume Next
    nPr = Val(lblScreen.Caption)
    EraseNext = True
End Sub

Private Sub cmdPCR_Click()
    On Error Resume Next
        If Val(lblScreen.Caption) = 0 Then
        lblScreen.Caption = 1
        Exit Sub
        End If
    nCr = Val(lblScreen.Caption)
    EraseNext = True
End Sub

Private Sub cmdPercent_Click()
On Error Resume Next
    lblScreen.Caption = (Val(lblScreen.Caption) * 0.01) + Val(lblScreen.Caption)
    cmdBasicFunctions(0).SetFocus
End Sub

Private Sub cmdPi_Click()
    On Error Resume Next
    lblScreen.Caption = FormatNumber(pi, 9)
    cmdBasicFunctions(0).SetFocus
End Sub

Private Sub cmdRandom_Click()
    On Error Resume Next
    lblScreen.Caption = Rnd
End Sub

Private Sub cmdSciFunctions_Click(Index As Integer)
    On Error Resume Next
    Dim Screen As Double, TheScreen As Double
    On Error Resume Next
        Select Case Index
        Case 0 'squared
            If Len(lblScreen.Caption) = 0 Then
            MsgBox ("ILLEGAL, 0 ²!!"), , "NYI"
            Exit Sub
            End If
        lblScreen.Caption = Val(lblScreen.Caption) ^ 2
        Case 1 'sqroute
            If Val(lblScreen.Caption) < 0 Then
            MsgBox ("ILLEGAL, sqr ( )!!"), , "NYI"
            Exit Sub
            End If
        lblScreen.Caption = Sqr(Val(lblScreen.Caption))
        Case 2 'factorial
            If Val(lblScreen.Caption) < 0 Then
            MsgBox ("ILLEGAL,FACT ( )!!"), , "NYI"
            Exit Sub
            Else
                If Val(lblScreen.Caption) = 0 Then
                lblScreen.Caption = 1
                Exit Sub
                End If
            End If
        Screen = Val(lblScreen.Caption)
        TheScreen = Val(lblScreen.Caption)
        lblScreen.Caption = Factorial(TheScreen, Screen)
        End Select
    cmdBasicFunctions(0).SetFocus
End Sub

Private Sub CmdBack_Click()
    On Error Resume Next
        If (lblScreen.Caption <> "") Then
        lblScreen.Caption = Mid(lblScreen.Caption, 1, Len(lblScreen.Caption) - 1)
        End If
cmdBasicFunctions(0).SetFocus
End Sub

Private Sub cmdSign_Click()
    On Error Resume Next
    lblScreen.Caption = Round(Sin(Radians(Val(lblScreen.Caption))), 9)
    cmdBasicFunctions(0).SetFocus
End Sub

Private Sub cmdTan_Click()
    On Error Resume Next
    lblScreen.Caption = Round(Tan(Radians(Val(lblScreen.Caption))), 9)
    cmdBasicFunctions(0).SetFocus
End Sub

Private Sub cmdX3_Click()
    On Error Resume Next
    lblScreen.Caption = Val(lblScreen.Caption) ^ 3
    cmdBasicFunctions(0).SetFocus
End Sub

Private Sub cmdxy_Click()
    Dim y As Integer
        If Len(lblScreen.Caption) = 0 Then
        MsgBox ("ILLEGAL, 0^x!!"), vbCritical, "NYI"
        Exit Sub
        End If
    yX = Val(lblScreen.Caption)
    cmdBasicFunctions(0).SetFocus
    EraseNext = True
End Sub

Private Sub CmdRnd_Click()
    On Error Resume Next
    lblScreen.Caption = Rnd
End Sub

Private Sub Form_Activate()
     On Error Resume Next
    cmdBasicFunctions(0).SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
     On Error Resume Next
    Call PressKey(KeyCode)
 
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    Unload Me
End If
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Randomize
    Times = 1
    Divide = 1
    Add = 0
    Subtract = 0
    CenterForm Me
End Sub

Private Function Radians(ByRef Degrees As Double)
    On Error Resume Next
'converts a number to radians for tan functions
    Radians = Degrees * pi / 180
End Function

Public Sub PressKey(ByVal Ind As Integer)
    On Error Resume Next
    If Ind = 13 Then
        Ind = 13
    End If
        If Ind >= 48 And Ind <= 57 Then
        Ind = Ind - 48
            If EraseNext = True Then
            EraseNext = False
            lblScreen.Caption = Ind
            Exit Sub
            Else
                lblScreen.Caption = lblScreen.Caption & Ind
                Exit Sub
            End If
        End If
    
             If Ind >= 96 And Ind < 106 Then
             Ind = Ind - 96
            ' EraseNext = False
                 If EraseNext = True Then
                 EraseNext = False
                 lblScreen.Caption = Ind
                 Exit Sub
                 Else
                     lblScreen.Caption = lblScreen.Caption & Ind
                     Exit Sub
             End If
                 End If
    
    Select Case Ind
    Case 107
    Ind = 1
    Call SubBasicFuntions(Ind)
    
    Case 109
    Ind = 2
    Call SubBasicFuntions(Ind)
    
    Case 106
    Ind = 3
    Call SubBasicFuntions(Ind)
    
    Case 111
    Ind = 4
    Call SubBasicFuntions(Ind)
    
    Case 110
    cmdDecimal_Click
    Case 8
    CmdBack_Click
    End Select
End Sub

Private Function Factorial(ByVal TheLabel As Double, Label As Double) As Double
'does the factorial of a number, used in combination and permutation
    On Error Resume Next
    Do Until Label = 1
    Label = Label - 1
    TheLabel = (Label) * TheLabel
    Loop
Factorial = TheLabel
End Function

Private Function Perm(ByRef x As Integer, N As Double, Temprary As Double, r As Double) As Double
'This does all the permutations, returns perm of 2 numbers
    On Error Resume Next
    Call PermComErrorCheck(N, r)
    x = 1
        Do Until x = r
        N = N * (Temprary - x)
        x = x + 1
        Loop
    EraseNext = True
    Perm = N
End Function

Private Sub lblScreen_Change()
Dim a As Integer
a = InStr(lblScreen, ",")
If a > 0 Then
lblScreen = Left(lblScreen, a - 1) & "." & Mid(lblScreen, a + 1, Len(lblScreen) - a - 1)
End If

    On Error Resume Next
        If Len(lblScreen.Caption) >= 20 Then
        CmdBack_Click
        Beep
        End If
End Sub

Private Sub Equal()
    On Error Resume Next
    Static NFirst As Double 'N in perm and comb
    Static Rsec As Double  'R in perm and comb
    Static temp As Double 'r in perm and comb
    Dim Counter As Integer
        If Divide <> 1 Then
            If Divide = 0 Or Val(lblScreen.Caption) = 0 Then
            MsgBox ("¡División por cero!"), vbCritical, "NYI"
            EraseNext = True
            Exit Sub
            End If
        lblScreen.Caption = Divide / Val(lblScreen.Caption)
        Divide = 1
            Else
            If Times <> 1 Then
            lblScreen.Caption = Val(lblScreen.Caption) * Times
            Times = 1
            Else
                If Add <> 0 Then
                lblScreen.Caption = Val(lblScreen.Caption) + Add
                Add = 0
                Else
                    If Subtract <> 0 Then
                    lblScreen.Caption = Subtract - Val(lblScreen.Caption)
                    Subtract = 0
                    Else
                        If codeMod <> 0 Then
                        lblScreen.Caption = codeMod Mod Val(lblScreen.Caption)
                        codeMod = 0
                        Else
                            If nPr <> 0 Then
                                If Val(lblScreen.Caption) = 0 Then
                                lblScreen.Caption = 1
                                Exit Sub
                                End If
                                    NFirst = nPr 'WORKING HERE!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
                                    Counter = 1
                                    Rsec = Val(lblScreen.Caption)
                                    EraseNext = False
                                    lblScreen.Caption = Perm(Counter, nPr, NFirst, Rsec) 'perm is a function
                                    nPr = 0
                                    Else
                                        If yX <> 0 Then
                                        Dim y As Double, Rsec1 As Double
                                        y = Val(lblScreen.Caption)
                                        lblScreen.Caption = yX ^ y
                                        yX = 0
                                        Else
                                            If nCr <> 0 Then 'gets perm
                                            NFirst = nCr
                                            Rsec = Val(lblScreen.Caption)
                                            Rsec1 = Rsec
                                            Counter = 1
                                            lblScreen.Caption = Perm(Counter, nCr, NFirst, Rsec) / Factorial(Rsec, Rsec1)
                                            nCr = 0
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
   ' EraseNext = True
    EraseNext = False
    cmdBasicFunctions(0).SetFocus
    If FrmRegistroComprasList.Procedencia = Mbimponible Then
        FrmRegistroComprasList.TxtValorcompra.Text = Format(Val(Me.lblScreen.Caption), "###0.00")
        Call Resalta(FrmRegistroComprasList.TxtValorcompra)
        FrmRegistroComprasList.Procedencia = Mneutro
        Unload Me
    End If
     If FrmRegistroComprasList.Procedencia = MbimponibleInafecta Then
        FrmRegistroComprasList.TxtValorCompraNoAfecta.Text = Format(Val(Me.lblScreen.Caption), "###0.00")
        Call Resalta(FrmRegistroComprasList.TxtValorCompraNoAfecta)
        FrmRegistroComprasList.Procedencia = Mneutro
        Unload Me
    End If
    If FrmRegistroComprasList.Procedencia = Migv Then
        FrmRegistroComprasList.txtTotal.Text = Format(Val(Me.lblScreen.Caption), "###0.00")
        FrmRegistroComprasList.CmdCentroCostos.SetFocus
        FrmRegistroComprasList.Procedencia = Mneutro
        Unload Me
    End If
    If FrmRegistroComprasList.Procedencia = MOtro Then
        FrmRegistroComprasList.TxtPercepcion = Format(Val(Me.lblScreen.Caption), "###0.00")
        Call Resalta(FrmRegistroComprasList.TxtPercepcion)
        FrmRegistroComprasList.Procedencia = Mneutro
        Unload Me
    End If
    
    If FrmRegistroComprasList.Procedencia = Misc Then
        FrmRegistroComprasList.txtisc = Format(Val(Me.lblScreen.Caption), "###0.00")
        Call Resalta(FrmRegistroComprasList.txtisc)
        FrmRegistroComprasList.Procedencia = Mneutro
        Unload Me
    End If
End Sub

Private Sub SubBasicFuntions(ByVal Index As Integer)
    On Error Resume Next
    Select Case Index
        Case 1
        Add = Val(lblScreen.Caption) + Add
        EraseNext = True
        lblScreen.Caption = Add
    
        Case 2
        Subtract = Val(lblScreen.Caption) - Subtract
        EraseNext = True
        lblScreen.Caption = Subtract
        
        Case 3
        Times = Val(lblScreen.Caption) * Times
        EraseNext = True
        lblScreen.Caption = Times
        
        Case 4
        Divide = Val(lblScreen.Caption) / Divide
        EraseNext = True
        lblScreen.Caption = Divide
        
        Case 0
        Call Equal
        End Select
End Sub

Private Sub PermComErrorCheck(ByVal First As Double, Second As Double)
On Error Resume Next
'number must be greater than or = to 0
    If (First < 0) Or (Second < 0) Or (Second > First) Then
    MsgBox ("ILLEGAL PERM ()!!"), , "NYI"
    End If
End Sub
Public Sub Resalta(ByVal texto As TextBox)
texto.SelStart = 0
texto.SelLength = Len(Trim(texto))
texto.Text = texto.SelText
texto.SetFocus
End Sub
