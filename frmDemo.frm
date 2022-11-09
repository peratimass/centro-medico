VERSION 5.00
Begin VB.Form frmDemo 
   Caption         =   "PDF417 Font Demo"
   ClientHeight    =   7530
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9135
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7530
   ScaleWidth      =   9135
   Begin VB.ComboBox cbxFontSize 
      Height          =   315
      ItemData        =   "frmDemo.frx":0000
      Left            =   1800
      List            =   "frmDemo.frx":001F
      TabIndex        =   8
      Top             =   3600
      Width           =   1935
   End
   Begin VB.ComboBox cbxFontName 
      Height          =   315
      ItemData        =   "frmDemo.frx":0043
      Left            =   1800
      List            =   "frmDemo.frx":0053
      TabIndex        =   7
      Top             =   3180
      Width           =   1935
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print PDF417"
      Height          =   615
      Left            =   6480
      TabIndex        =   10
      Top             =   1740
      Width           =   2295
   End
   Begin VB.CheckBox chkHandleTilde 
      Alignment       =   1  'Right Justify
      Caption         =   "Process Tilde"
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   2730
      Width           =   1815
   End
   Begin VB.CheckBox chkTrimSymbol 
      Alignment       =   1  'Right Justify
      Caption         =   "Truncate Symbol "
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   2235
      Width           =   1815
   End
   Begin VB.ComboBox cbxMode 
      Height          =   315
      ItemData        =   "frmDemo.frx":008F
      Left            =   1800
      List            =   "frmDemo.frx":009C
      TabIndex        =   4
      Top             =   1800
      Width           =   1815
   End
   Begin VB.TextBox txtColumns 
      Height          =   315
      Left            =   1800
      TabIndex        =   3
      Text            =   "3"
      Top             =   1380
      Width           =   1815
   End
   Begin VB.TextBox txtRows 
      Height          =   315
      Left            =   1800
      TabIndex        =   2
      Text            =   "3"
      Top             =   960
      Width           =   1815
   End
   Begin VB.ComboBox cbxSecurity 
      Height          =   315
      ItemData        =   "frmDemo.frx":00B4
      Left            =   1800
      List            =   "frmDemo.frx":00D3
      TabIndex        =   1
      Top             =   540
      Width           =   1815
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   615
      Left            =   6480
      TabIndex        =   11
      Top             =   3240
      Width           =   2295
   End
   Begin VB.CommandButton cmdDisplay 
      Caption         =   "Display PDF417"
      Height          =   615
      Left            =   6480
      TabIndex        =   9
      Top             =   240
      Width           =   2295
   End
   Begin VB.TextBox txtMessage 
      Height          =   315
      Left            =   1800
      TabIndex        =   0
      Text            =   "1234"
      Top             =   120
      Width           =   3615
   End
   Begin VB.TextBox txtPDF417 
      Alignment       =   2  'Center
      Height          =   3135
      Left            =   360
      MultiLine       =   -1  'True
      TabIndex        =   13
      Text            =   "frmDemo.frx":0128
      Top             =   4080
      Width           =   8295
   End
   Begin VB.Label lblFontSize 
      Caption         =   "Font Size"
      Height          =   255
      Left            =   240
      TabIndex        =   19
      Top             =   3600
      Width           =   975
   End
   Begin VB.Label lblFontName 
      Caption         =   "Font Name"
      Height          =   255
      Left            =   240
      TabIndex        =   18
      Top             =   3225
      Width           =   1095
   End
   Begin VB.Label lblMode 
      Caption         =   "Mode"
      Height          =   255
      Left            =   240
      TabIndex        =   17
      Top             =   1860
      Width           =   615
   End
   Begin VB.Label lblColumns 
      Caption         =   "Columns"
      Height          =   375
      Left            =   240
      TabIndex        =   16
      Top             =   1365
      Width           =   855
   End
   Begin VB.Label lblRows 
      Caption         =   "Rows"
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   990
      Width           =   615
   End
   Begin VB.Label lblSecurity 
      Caption         =   "Error Correction"
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   615
      Width           =   1455
   End
   Begin VB.Label lblMessage 
      Caption         =   "Message"
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "frmDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub PDF417Encode Lib "PDF417Font.dll" _
(ByVal Message As String, ByVal Mode As Integer, ByVal ECLevel As Integer, _
 ByVal Rows As Integer, ByVal Columns As Integer, ByVal TruncatedSymbol As Boolean, _
 ByVal HandleTilde As Boolean)

Private Declare Function PDF417GetRows Lib "PDF417Font.dll" () As Integer
Private Declare Function PDF417GetCols Lib "PDF417Font.dll" () As Integer
Private Declare Function PDF417GetCharAt Lib "PDF417Font.dll" (ByVal RowIndex As Integer, ByVal ColIndex As Integer) As Integer

Public Sub printer_barcode(ByVal in_contenido As String)
    Dim RowCount As Integer
    Dim ColCount As Integer
    Dim OneLine As String
    
   ' If (Not Check()) Then
   '     Exit Sub
   ' End If
    
    ' encode string using PDF417
    Call PDF417Encode(in_contenido, 2, 2, 3, 3, 0, 0)
    
    Printer.Font.name = "MW6 PDF417R3"
    Printer.Font.Size = 8
    Printer.CurrentX = 600
    Printer.CurrentY = 800
    
    ' how many rows?
    RowCount = PDF417GetRows
    
    ' how many characters in one row?
    ColCount = PDF417GetCols
    
    ' print all rows using PDF417 font
    For i = 1 To RowCount
        OneLine = ""
        For j = 1 To ColCount
            OneLine = OneLine & Chr(PDF417GetCharAt(i - 1, j - 1))
        Next j
        Printer.Print OneLine
        Printer.CurrentX = 600
    Next i
Printer.EndDoc
End Sub

Private Function Check() As Boolean
    Dim ErrMSg As String

    ' do error-check
    On Error GoTo HandleErr
    ErrMSg = ""
    If (Len(txtMessage.Text) = 0) Then
        ErrMSg = "Please enter the message to encode with PDF417."
        ErrMSg = ErrMSg & Chr(13) & Chr(10)
    End If
    
    If ((CInt(txtRows.Text) > 90) Or (CInt(txtRows.Text) < 3)) Then
        ErrMSg = ErrMSg & "Rows must be between 3 and 90."
        ErrMSg = ErrMSg & Chr(13) & Chr(10)
    End If
    
    If ((CInt(txtColumns.Text) > 30) Or (CInt(txtColumns.Text) < 3)) Then
        ErrMSg = ErrMSg & "Columns must be between 3 and 30."
        ErrMSg = ErrMSg & Chr(13) & Chr(10)
    End If
    
    If (Len(ErrMSg) > 0) Then
        MsgBox ErrMSg
        Check = False
        Exit Function
    End If
    
    Check = True
    Exit Function
    
HandleErr:
    Check = False
    
End Function

Private Sub cmdDataStream_Click()

End Sub

Private Sub cmdDisplay_Click()
    Dim RowCount As Integer
    Dim ColCount As Integer
    Dim OneLine As String
    Dim EncodedMsg As String
    
    If (Not Check()) Then
        Exit Sub
    End If
    
    txtPDF417.Text = ""
    Texto = Trim(cbxFontName.Text)
    txtPDF417.FontName = Texto
    txtPDF417.FontSize = CInt(cbxFontSize.Text)
    
    ' encode string using PDF417
    Call PDF417Encode(txtMessage.Text, CInt(cbxMode.ListIndex), _
                      CInt(cbxSecurity.ListIndex), CInt(txtRows.Text), _
                      CInt(txtColumns.Text), chkTrimSymbol.Value, chkHandleTilde.Value)
    
    ' how many rows?
    RowCount = PDF417GetRows
    
    ' how many characters in one row?
    ColCount = PDF417GetCols
    
    ' produce string for PDF417 font
    EncodedMsg = vbCrLf
    For i = 1 To RowCount
        For j = 1 To ColCount
            EncodedMsg = EncodedMsg & Chr(PDF417GetCharAt(i - 1, j - 1))
        Next j
        EncodedMsg = EncodedMsg & vbCrLf
    Next i
    txtPDF417.Text = EncodedMsg
    
End Sub

Private Sub cmdPrint_Click()
Call printer_barcode("Hola")
Exit Sub
End Sub

Private Sub Form_Load()
    cbxSecurity.ListIndex = 2
    cbxMode.ListIndex = 2
    cbxFontName.ListIndex = 0
    cbxFontSize.ListIndex = 2
End Sub

Private Sub cmdexit_Click()
    Unload Me
End Sub
