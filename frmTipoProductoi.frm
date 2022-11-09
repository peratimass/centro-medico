VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmTipoProductoi 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5055
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7935
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   7935
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmdetalle 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Visible         =   0   'False
      Width           =   6135
      Begin VB.TextBox txttipo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   1800
         TabIndex        =   13
         Top             =   2760
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CheckBox chkServicio 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "SERVICIO"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   300
         Left            =   1800
         TabIndex        =   12
         Top             =   1080
         Width           =   2895
      End
      Begin VB.TextBox txtid 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   6360
         TabIndex        =   2
         Top             =   3720
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox TxtDescripcion 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   1800
         TabIndex        =   1
         Top             =   600
         Width           =   2895
      End
      Begin VitekeySoft.ChameleonBtn cmdprocear 
         Height          =   615
         Left            =   1800
         TabIndex        =   3
         Top             =   1920
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1085
         BTYPE           =   5
         TX              =   "PROCESAR"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   12582912
         FCOLO           =   12582912
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmTipoProductoi.frx":0000
         PICN            =   "frmTipoProductoi.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdsalir 
         Height          =   615
         Left            =   3240
         TabIndex        =   4
         Top             =   1920
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1085
         BTYPE           =   5
         TX              =   "SALIR"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   12582912
         FCOLO           =   12582912
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmTipoProductoi.frx":3664
         PICN            =   "frmTipoProductoi.frx":3680
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DESCRIPCION :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   525
         TabIndex        =   5
         Top             =   645
         Width           =   990
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfdPersona 
      Height          =   4095
      Left            =   240
      TabIndex        =   6
      Top             =   480
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   7223
      _Version        =   393216
      ForeColor       =   8388608
      Rows            =   0
      FixedRows       =   0
      FixedCols       =   0
      ForeColorFixed  =   8388608
      BackColorBkg    =   16777215
      GridColor       =   8388608
      SelectionMode   =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
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
   Begin VitekeySoft.ChameleonBtn cmdexit 
      Height          =   855
      Left            =   6600
      TabIndex        =   7
      Top             =   2400
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1508
      BTYPE           =   5
      TX              =   "SALIR"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   12582912
      FCOLO           =   12582912
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmTipoProductoi.frx":3A70
      PICN            =   "frmTipoProductoi.frx":3A8C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdupdate 
      Height          =   855
      Left            =   6600
      TabIndex        =   8
      Top             =   1440
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1508
      BTYPE           =   5
      TX              =   "MODIFICAR"
      ENAB            =   0   'False
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   12582912
      FCOLO           =   12582912
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmTipoProductoi.frx":3E7C
      PICN            =   "frmTipoProductoi.frx":3E98
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdnuevo 
      Height          =   855
      Left            =   6600
      TabIndex        =   9
      Top             =   480
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1508
      BTYPE           =   5
      TX              =   "NUEVO"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   12582912
      FCOLO           =   12582912
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmTipoProductoi.frx":41B2
      PICN            =   "frmTipoProductoi.frx":41CE
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblAcoount 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   18480
      TabIndex        =   11
      Top             =   360
      Width           =   2295
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "TIPOS DE PRODUCTO"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   240
      TabIndex        =   10
      Top             =   120
      Width           =   1605
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   5055
      Left            =   0
      Top             =   0
      Width           =   7935
   End
End
Attribute VB_Name = "frmTipoProductoi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub llenarGrid_prod(ByVal Grilla As MSHFlexGrid)
Dim in_precio As String
On Error GoTo salir
strCadena = "SELECT * FROM tipo_producto WHERE ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 0
    
    Exit Sub
End If
  
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
            ' -- almacena el ancho de los campos
           Grilla.ColWidth(0) = 1100
           Grilla.ColWidth(1) = 3000
           Grilla.ColWidth(2) = 1200
           
    Next
         cabecera = "CODIGO" & vbTab & "DESCRIPCION" & vbTab & "TIPO"
         Grilla.AddItem cabecera
         For k = 0 To 2
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
                            Next k
                            
        rst.MoveFirst
        For i = 1 To rst.RecordCount
            If rst("servicio") = "si" Then
                in_tipo = "SERVICIO"
            Else
                in_tipo = "PRODUCTO"
            End If
            Fila = rst("id_tipoproducto") & vbTab & rst("descripcion") & vbTab & in_tipo
            Grilla.AddItem Fila
            rst.MoveNext
        Next i
        
Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub

Private Sub cmdexit_Click()
Unload Me
End Sub

Private Sub cmdnuevo_Click()
Me.frmdetalle.Visible = True
Me.txtdescripcion.Text = ""
Me.chkServicio.Value = 0
Me.txttipo.Text = 0
End Sub

Private Sub cmdprocear_Click()
Dim in_servicio As String
If Me.chkServicio.Value = 1 Then
   in_servicio = "si"
Else
   in_servicio = "no"
End If

If Trim(Me.txtdescripcion.Text) <> "" Then
    
    If Val(Me.txttipo.Text) < 1 Then
       strCadena = "SELECT id_tipoproducto FROM tipo_producto WHERE ruc='" & KEY_RUC & "' ORDER BY id_tipoproducto DESC LIMIT 1"
       Call ConfiguraRst(strCadena)
       If rst.RecordCount > 0 Then
           Me.txttipo.Text = Format(Val(rst("id_tipoproducto")) + 1, "00")
       Else
          Me.txttipo.Text = "01"
       End If
    End If
    strCadena = "put_tipo_producto('" & Trim(Me.txttipo.Text) & "','" & Trim(Me.txtdescripcion.Text) & "','" & in_servicio & "','" & KEY_RUC & "')"
    CnBd.Execute (strCadena)
    Me.frmdetalle.Visible = False
    Call Me.llenarGrid_prod(Me.HfdPersona)
End If
End Sub

Private Sub cmdsalir_Click()
Me.frmdetalle.Visible = False

End Sub

Private Sub cmdupdate_Click()
strCadena = "SELECT * FROM tipo_producto WHERE id_tipoproducto='" & Trim(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)) & "' and ruc='" & KEY_RUC & "' LIMIT 1"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    Me.txtdescripcion.Text = rst("descripcion")
    Me.txttipo.Text = rst("id_tipoproducto")
    If rst("servicio") = "si" Then
        Me.chkServicio.Value = 1
    Else
        Me.chkServicio.Value = 0
    End If
    Me.frmdetalle.Visible = True
End If
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 100
Call Me.llenarGrid_prod(Me.HfdPersona)
End Sub

Private Sub HfdPersona_Click()
If Me.HfdPersona.Rows > 0 Then
    If Val(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)) > 0 Then
        Me.cmdupdate.Enabled = True
      
    Else
        Me.cmdupdate.Enabled = False
      
    End If
End If

End Sub
