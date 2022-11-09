VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmPlanesServicio 
   BorderStyle     =   0  'None
   Caption         =   "Conceptos Gastos Internos"
   ClientHeight    =   6495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   14415
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "DETALLE PLAN"
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
      Height          =   3495
      Left            =   2760
      TabIndex        =   6
      Top             =   1200
      Visible         =   0   'False
      Width           =   7575
      Begin VB.TextBox txtid_producto 
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
         Left            =   2040
         TabIndex        =   21
         Top             =   2760
         Width           =   1455
      End
      Begin VB.TextBox txtid_plan 
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
         Left            =   6000
         TabIndex        =   19
         Top             =   600
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtMontoFacturado 
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
         Left            =   2040
         TabIndex        =   17
         Top             =   2280
         Width           =   1455
      End
      Begin VB.TextBox txtSucursales 
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
         Left            =   2040
         TabIndex        =   15
         Top             =   1800
         Width           =   1455
      End
      Begin VB.TextBox TxtUsuarios 
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
         Left            =   2040
         TabIndex        =   13
         Top             =   1320
         Width           =   1455
      End
      Begin VB.CheckBox chk_elimitados 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         Caption         =   "ILIMITADOS"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   3600
         TabIndex        =   11
         Top             =   840
         Width           =   1935
      End
      Begin VB.TextBox txtNumeroComprobantes 
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
         Left            =   2040
         TabIndex        =   10
         Top             =   840
         Width           =   1455
      End
      Begin VB.TextBox txtDescripcion 
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
         Left            =   2040
         TabIndex        =   8
         Top             =   360
         Width           =   3495
      End
      Begin VitekeySoft.ChameleonBtn cmdprocesar 
         Height          =   780
         Left            =   5760
         TabIndex        =   18
         Top             =   2280
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   1376
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
         FCOL            =   8388608
         FCOLO           =   8388608
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmGastos.frx":0000
         PICN            =   "FrmGastos.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lblproducto 
         BackColor       =   &H000080FF&
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
         Height          =   255
         Left            =   2040
         TabIndex        =   22
         Top             =   3135
         Width           =   5265
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SERVICIO VINCULADO :"
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
         Left            =   165
         TabIndex        =   20
         Top             =   3000
         Width           =   1545
      End
      Begin VB.Image cerrar 
         Height          =   240
         Left            =   7280
         Picture         =   "FrmGastos.frx":3664
         Top             =   120
         Width           =   240
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FACTURADO HASTA :"
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
         Left            =   285
         TabIndex        =   16
         Top             =   2520
         Width           =   1395
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N° SUCURSALES :"
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
         TabIndex        =   14
         Top             =   2040
         Width           =   1155
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N° USUARIOS :"
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
         Left            =   705
         TabIndex        =   12
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N° COMPROBANTES :"
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
         Left            =   255
         TabIndex        =   9
         Top             =   1080
         Width           =   1425
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
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
         Left            =   675
         TabIndex        =   7
         Top             =   480
         Width           =   1005
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfgLinea 
      Height          =   5895
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   12975
      _ExtentX        =   22886
      _ExtentY        =   10398
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
   Begin VitekeySoft.ChameleonBtn cmdnuevo 
      Height          =   780
      Left            =   13320
      TabIndex        =   2
      Top             =   480
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1376
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
      FCOL            =   8388608
      FCOLO           =   8388608
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmGastos.frx":6508
      PICN            =   "FrmGastos.frx":6524
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdmodificar 
      Height          =   780
      Left            =   13320
      TabIndex        =   3
      Top             =   1320
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1376
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
      FCOL            =   8388608
      FCOLO           =   8388608
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmGastos.frx":6976
      PICN            =   "FrmGastos.frx":6992
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdeliminar 
      Height          =   780
      Left            =   13320
      TabIndex        =   4
      Top             =   2160
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1376
      BTYPE           =   5
      TX              =   "ELIMINAR"
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
      FCOL            =   8388608
      FCOLO           =   8388608
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmGastos.frx":8FCB
      PICN            =   "FrmGastos.frx":8FE7
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdsalir 
      Height          =   780
      Left            =   13320
      TabIndex        =   5
      Top             =   3000
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1376
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
      FCOL            =   8388608
      FCOLO           =   8388608
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmGastos.frx":B431
      PICN            =   "FrmGastos.frx":B44D
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label LblEmpresa 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PLANES EMPRESARIALES"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   270
      TabIndex        =   1
      Top             =   120
      Width           =   1875
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   6495
      Left            =   0
      Top             =   0
      Width           =   14415
   End
   Begin VB.Shape ShpDatos 
      BackColor       =   &H00DFDFE0&
      FillColor       =   &H00DFDFE0&
      FillStyle       =   0  'Solid
      Height          =   675
      Left            =   240
      Top             =   5640
      Width           =   5895
   End
End
Attribute VB_Name = "frmPlanesServicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Procedencia As EnumProcede

Private Sub cerrar_Click()
Me.Frame1.Visible = False
End Sub

Private Sub cmdEliminar_Click()
Procedencia = Eliminar
frmsegurity.Show
Exit Sub
End Sub

Private Sub cmdmodificar_Click()
Call load_plan(Me.HfgLinea.TextMatrix(Me.HfgLinea.Row, 0))
End Sub

Private Sub cmdNuevo_Click()
Me.txtDescripcion.Text = ""
Me.txtNumeroComprobantes.Text = ""
Me.txtMontoFacturado.Text = ""
Me.txtSucursales.Text = ""
Me.txtid_plan.Text = ""
Me.TxtUsuarios.Text = ""
Me.Frame1.Visible = True
End Sub
Private Sub load_plan(ByVal in_plan As String)
strCadena = "SELECT * FROM plan_servicio WHERE id_plan='" & Val(in_plan) & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    Me.txtid_plan.Text = rst("id_plan")
    Me.txtDescripcion.Text = rst("descripcion")
    Me.txtNumeroComprobantes.Text = rst("numero_comprobante")
    Me.txtMontoFacturado.Text = rst("monto_facturado")
    Me.txtSucursales.Text = rst("numero_sucursal")
    Me.TxtUsuarios.Text = rst("numero_usuario")
    Me.txtid_plan.Text = rst("id_plan")
    If rst("comprobantes_elimitados") = "si" Then
       Me.chk_elimitados.Value = 1
    Else
       Me.chk_elimitados.Value = 0
    End If
    Me.txtid_producto.Text = rst("id_producto")
    Me.lblproducto.Caption = get_producto(rst("id_producto"))
    Me.Frame1.Visible = True
End If
End Sub
Private Sub cmdprocesar_Click()
Call Save
End Sub

Private Sub cmdsalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 200
Call actualizar(Me.HfgLinea)
End Sub
Public Sub Save()

If Trim(Me.txtDescripcion.Text) = "" Then
   MsgBox "INGRESE UNA DESCRIPCION", vbInformation, KEY_VENDEDOR
   Exit Sub
End If
in_ilimitado = "no"
If Me.chk_elimitados.Value = 0 Then
      If Val(Me.txtNumeroComprobantes.Text) < 1 Then
         MsgBox "INGRESE NUMERO DE COMPROBANTES", vbInformation, KEY_VENDEDOR
      End If
      
      in_ilimitado = "no"
      
      If Val(Me.txtMontoFacturado.Text) <= 0 Then
         MsgBox "INGRESE UN MONTO FACTURADO.", vbInformation, KEY_VENDEDOR
         Exit Sub
      End If
      
    
      
Else
    in_ilimitado = "si"
End If

If Val(Me.txtSucursales.Text) < 1 Then
     MsgBox "INGRESE NUMERO DE SUCURSALES", vbInformation, KEY_VENDEDOR
     Exit Sub
End If

If Val(Me.TxtUsuarios.Text) < 1 Then
     MsgBox "INGRESE NUMERO DE USUARIOS", vbInformation, KEY_VENDEDOR
     Exit Sub
End If

strCadena = "call put_plan('" & Val(Me.txtid_plan.Text) & "','" & Trim(Me.txtDescripcion.Text) & "','" & Val(Me.txtNumeroComprobantes.Text) & "','" & in_ilimitado & "','" & Val(Me.txtSucursales.Text) & "','" & Val(Me.txtMontoFacturado.Text) & "','" & Val(Me.TxtUsuarios.Text) & "','" & KEY_USUARIO & "','" & Trim(Me.txtid_producto.Text) & "','" & KEY_RUC & "')"
CnBd.Execute (strCadena)
Me.Frame1.Visible = False
Call actualizar(Me.HfgLinea)




End Sub



Public Sub actualizar(ByVal Grilla As MSHFlexGrid)

strCadena = "SELECT * FROM view_plan WHERE id_alm='" & KEY_ALM & "' and   ruc='" & KEY_RUC & "' ORDER BY id_plan"
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 0
    Exit Sub

End If
   
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
           Grilla.ColWidth(0) = 800
           Grilla.ColWidth(1) = 3200
           Grilla.ColWidth(2) = 1400
           Grilla.ColWidth(3) = 1100
           Grilla.ColWidth(4) = 1300
           Grilla.ColWidth(5) = 1400
           Grilla.ColWidth(6) = 1100
           Grilla.ColWidth(7) = 2000
       Next
        cabecera = "CODIGO" & vbTab & "DESCRIPCION" & vbTab & "N° COMPROBANTES" & vbTab & "SUCURSALES" & vbTab & "USUARIOS" & vbTab & "M.FACTURADO" & vbTab & "PRECIO" & vbTab & "OPERADOR"
        Grilla.AddItem cabecera
         For k = 0 To 7
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
                            Next k
                            
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
          
          If rst("comprobantes_elimitados") = "si" Then
             in_ilimitados = "ILIMITADOS"
          Else
             in_ilimitados = rst("numero_comprobante")
          End If
          Fila = Format(rst("id_plan"), "0000") & vbTab & rst("descripcion") & vbTab & in_ilimitados & vbTab & rst("numero_sucursal") & vbTab & rst("numero_usuario") & vbTab & Format(rst("monto_facturado"), "#,##0.00") & vbTab & Format(rst("precio_venta"), "#,##0.00") & vbTab & rst("nombre_completo")
          Grilla.AddItem Fila
          
          
          rst.MoveNext
      Next i
        
       
      
End Sub


Private Sub HfgLinea_SelChange()
If Val(Me.HfgLinea.TextMatrix(Me.HfgLinea.Row, 0)) > 0 Then
   Me.cmdmodificar.Enabled = True
   Me.cmdEliminar.Enabled = True
Else
   Me.cmdmodificar.Enabled = False
   Me.cmdEliminar.Enabled = False
End If
End Sub

Private Sub txtid_producto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Procedencia = Selecionar
    FrmProducto.Show
    Exit Sub
End If
End Sub
