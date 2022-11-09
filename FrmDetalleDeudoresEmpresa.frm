VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmDetalleDeudoresEmpresa 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8505
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13230
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8505
   ScaleWidth      =   13230
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   1095
      Left            =   11400
      TabIndex        =   0
      Top             =   7320
      Width           =   1335
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HFTrabajadores 
      Height          =   3855
      Left            =   240
      TabIndex        =   1
      Top             =   3960
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   6800
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   8388608
      FixedCols       =   0
      ForeColorFixed  =   8388608
      GridColor       =   0
      FocusRect       =   0
      GridLines       =   2
      SelectionMode   =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfEmpresas 
      Height          =   2775
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   10920
      _ExtentX        =   19262
      _ExtentY        =   4895
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   8388608
      FixedCols       =   0
      ForeColorFixed  =   8388608
      GridColor       =   0
      FocusRect       =   0
      GridLines       =   2
      SelectionMode   =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label lblDeuda 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Deuda:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   8820
      TabIndex        =   11
      Top             =   8040
      Width           =   1605
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Empresas Afiliadas"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   270
      TabIndex        =   10
      Top             =   240
      Width           =   1845
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Deuda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   3435
      TabIndex        =   9
      Top             =   2040
      Width           =   645
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dirección:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   720
      TabIndex        =   8
      Top             =   1620
      Width           =   915
   End
   Begin VB.Label LblIdentificacion 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ruc:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   885
      TabIndex        =   7
      Top             =   1980
      Width           =   555
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Datos Persona:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   855
      TabIndex        =   6
      Top             =   960
      Width           =   1425
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Trabajadores Afiliados"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   300
      TabIndex        =   5
      Top             =   3600
      Width           =   2205
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Deuda:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   3000
      TabIndex        =   4
      Top             =   8040
      Width           =   5565
   End
   Begin VB.Label lbltotaldeuda 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   9945
      TabIndex        =   3
      Top             =   3480
      Width           =   75
   End
End
Attribute VB_Name = "FrmDetalleDeudoresEmpresa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 200
strCadena = "SELECT     cPersona as COD, NombrePersona as EMPRESA, sDireccionCliente1 AS DIRECCION, Per_Ruc AS RUC_DNI From Persona WHERE (credito = 'si') AND (Ruc_Empresa ='0')"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    Set Me.HfEmpresas.Recordset = rst
    Me.HfEmpresas.ColWidth(0) = 700
    Me.HfEmpresas.ColWidth(1) = 4250
    Me.HfEmpresas.ColWidth(2) = 4250
    Me.HfEmpresas.ColWidth(3) = 1400
End If
End Sub

Private Sub HfEmpresas_Click()
Dim rstD As New ADODB.Recordset
Dim total As Single
If Me.HfEmpresas.Rows > 0 Then
strCadena = "SELECT cPersona, NombrePersona, sDireccionCliente1, Int_Persona FROM Persona where credito='si' and ruc_empresa='" & Trim(Me.HfEmpresas.TextMatrix(Me.HfEmpresas.Row, 3)) & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
        Me.HFTrabajadores.Clear
       Me.HFTrabajadores.Rows = 0
   Exit Sub
 End If
        
       n = 1
       
       Me.HFTrabajadores.Clear
       Me.HFTrabajadores.Rows = 0
      ' Me.HfdGrilla.Rows = rst.RecordCount - 2
      ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
            ' -- almacena el ancho de los campos
            Me.HFTrabajadores.ColWidth(0) = 700
            Me.HFTrabajadores.ColWidth(1) = 4250
            Me.HFTrabajadores.ColWidth(2) = 4250
            Me.HFTrabajadores.ColWidth(3) = 1400
        Next
        
        ' Modificar
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
            strCadena = "SELECT sum(saldo) FROM DocumentoVenta WHERE cPersona='" & rst(0) & "' AND idformapago='" & KEY_CREDITO & "' AND saldo>0 "
            rstD.Open strCadena, CnBd, adOpenKeyset, adLockOptimistic
            If IsNull(rstD(0)) = True Then
                total = 0
            Else
                total = rstD(0)
             End If
            Set rstD = Nothing
            fila = fila & rst("cPersona") & vbTab & rst("NombrePersona") & vbTab & rst("sDireccionCliente1") & vbTab & Format(total, "###0.00")
            If (fila = "") Then
                x = 1
            End If
            Me.HFTrabajadores.AddItem fila
               
                                  
    
  
            fila = ""
            rst.MoveNext
             
        Next i

'---------
   
    strCadena = "SELECT    sum(DocumentoVenta.Saldo) AS Saldo FROM         Persona INNER JOIN " & _
    " DocumentoVenta ON Persona.cPersona = DocumentoVenta.cPersona " & _
    "WHERE     (Persona.credito = 'si') AND Persona.ruc_empresa='" & Trim(Me.HfEmpresas.TextMatrix(Me.HfEmpresas.Row, 3)) & "' AND DocumentoVenta.saldo>0 AND DocumentoVenta.anulado='F'  "
    Call ConfiguraRst(strCadena)
    If IsNull(rst(0)) = True Then
        Me.lblDeuda.Caption = "S/." & 0#
        Me.Label16.Caption = Me.HfEmpresas.TextMatrix(Me.HfEmpresas.Row, 1)
    Else
        Me.Label16.Caption = Me.HfEmpresas.TextMatrix(Me.HfEmpresas.Row, 1)
        Me.lblDeuda.Caption = "S/." & Format(rst(0), "###0.00")
    End If
    Set rst = Nothing
Else
    Me.lblDeuda.Caption = 0#
    Me.Label16.Caption = ""
End If
End Sub
