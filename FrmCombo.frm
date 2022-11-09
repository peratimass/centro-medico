VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form frmCombo 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Detalle  Combo's"
   ClientHeight    =   7665
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14175
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7665
   ScaleWidth      =   14175
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtIdCompra 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8760
      TabIndex        =   20
      Top             =   6480
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtBuscar 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   7560
      MaxLength       =   80
      TabIndex        =   18
      Top             =   720
      Width           =   1455
   End
   Begin VitekeySoft.ChameleonBtn cmdConsultar 
      Height          =   320
      Left            =   12480
      TabIndex        =   16
      Top             =   720
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "CONSULTAR"
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmCombo.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdnuevo 
      Height          =   780
      Left            =   10080
      TabIndex        =   13
      Top             =   6720
      Width           =   855
      _ExtentX        =   1508
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
      MICON           =   "FrmCombo.frx":001C
      PICN            =   "FrmCombo.frx":0038
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox TxtDetalle 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1320
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Top             =   6600
      Width           =   3495
   End
   Begin VB.TextBox TxtCombo 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   11040
      TabIndex        =   8
      Top             =   360
      Width           =   1335
   End
   Begin VB.TextBox TxtCantidad 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   8160
      MaxLength       =   80
      TabIndex        =   0
      Top             =   7080
      Width           =   1455
   End
   Begin MSDataListLib.DataCombo DtcCombo 
      Height          =   315
      Left            =   2040
      TabIndex        =   3
      Top             =   720
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      ForeColor       =   8388608
      Text            =   "DataCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComCtl2.DTPicker DtpFecha 
      Height          =   315
      Left            =   11040
      TabIndex        =   4
      Top             =   720
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   172097537
      CurrentDate     =   39371
   End
   Begin MSDataListLib.DataCombo DtcAlmacen 
      Height          =   315
      Left            =   2070
      TabIndex        =   5
      Top             =   240
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      ForeColor       =   8388608
      Text            =   "DataCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfdDetalle 
      Height          =   5055
      Left            =   240
      TabIndex        =   11
      Top             =   1320
      Width           =   13335
      _ExtentX        =   23521
      _ExtentY        =   8916
      _Version        =   393216
      ForeColor       =   8388608
      FixedCols       =   0
      ForeColorFixed  =   8388608
      BackColorBkg    =   16777215
      SelectionMode   =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VitekeySoft.ChameleonBtn cmdprocesar 
      Height          =   780
      Left            =   10965
      TabIndex        =   14
      Top             =   6720
      Width           =   855
      _ExtentX        =   1508
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
      MICON           =   "FrmCombo.frx":048A
      PICN            =   "FrmCombo.frx":04A6
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
      Left            =   12765
      TabIndex        =   15
      Top             =   6720
      Width           =   855
      _ExtentX        =   1508
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
      MICON           =   "FrmCombo.frx":3AEE
      PICN            =   "FrmCombo.frx":3B0A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComctlLib.ProgressBar prog_indicador 
      Height          =   195
      Left            =   10080
      TabIndex        =   17
      Top             =   6480
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VitekeySoft.ChameleonBtn cmdImprimir 
      Height          =   780
      Left            =   11865
      TabIndex        =   19
      Top             =   6720
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1376
      BTYPE           =   5
      TX              =   "IMPRIMIR"
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
      MICON           =   "FrmCombo.frx":3EFA
      PICN            =   "FrmCombo.frx":3F16
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblanulado 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ANULADO"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   5040
      TabIndex        =   12
      Top             =   6360
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "OBSERVACION"
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
      Left            =   270
      TabIndex        =   10
      Top             =   6840
      Width           =   975
   End
   Begin VB.Shape Shape3 
      BorderStyle     =   3  'Dot
      Height          =   975
      Left            =   240
      Top             =   6480
      Width           =   4815
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   6  'Inside Solid
      Height          =   615
      Left            =   5400
      Top             =   6840
      Width           =   4335
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nº:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   345
      Left            =   10080
      TabIndex        =   7
      Top             =   360
      Width           =   465
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   6  'Inside Solid
      Height          =   1095
      Left            =   9600
      Top             =   120
      Width           =   3855
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ALMACEN :"
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
      Left            =   1020
      TabIndex        =   6
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "COMBO :"
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
      Left            =   1065
      TabIndex        =   2
      Top             =   720
      Width           =   705
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "N° PRODUCTOS TERMINADOS :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   165
      Left            =   5565
      TabIndex        =   1
      Top             =   7125
      Width           =   2385
   End
   Begin VB.Shape ShpDatos 
      BackColor       =   &H00DFDFE0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   7635
      Left            =   0
      Top             =   0
      Width           =   14175
   End
End
Attribute VB_Name = "frmCombo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strCombo As String
Attribute strCombo.VB_VarHelpID = -1
Dim rst2 As New ADODB.Recordset
Public Procedencia As EnumProcede
Private Sub cmdGenerar_Click()
Dim i As Integer
Dim cantidad As Single
Dim combos As Integer
Dim rst1 As New ADODB.Recordset
If Val(Me.txtCantidad.Text) > 0 Then
strCadena = "SELECT     Producto_sub.cProducto as Codigo, Producto.DescripcionProducto as Descripcion, Unidad.sAbreviatura as Unidad, Producto_sub.cantidad as Cantidad " & _
  "FROM Producto_sub INNER JOIN Producto ON Producto_sub.cProducto = Producto.cProducto INNER JOIN " & _
  "Unidad ON Producto.cUnidad = Unidad.cUnidad WHERE Producto_sub.cProductoPadre='" & Trim(Me.DtcCombo.BoundText) & "'"
   Call ConfiguraRst(strCadena)
   rst.MoveFirst
  For i = 0 To rst.RecordCount - 1
      cantidad = Val(Me.txtCantidad.Text) * rst(3)
      combos = Val(Me.txtCantidad.Text)
      rst1.Open "SELECT Stock FROM Almacen_Productos WHERE cProducto='" & Trim(rst(0)) & "' AND Alm_cod='" & Trim(Me.DtcAlmacen.BoundText) & "'", CnBd, adOpenKeyset, adLockOptimistic
      rst1(0) = rst1(0) - cantidad
      rst1.Update
      Set rst1 = Nothing
      
      rst.MoveNext
  Next i
 rst1.Open "SELECT Stock FROM Almacen_Productos WHERE cProducto='" & Trim(Me.DtcCombo.BoundText) & "' AND Alm_cod='" & Trim(Me.DtcAlmacen.BoundText) & "'", CnBd, adOpenKeyset, adLockOptimistic
      rst1(0) = rst1(0) + combos
      rst1.Update
      Set rst1 = Nothing
      Unload Me
    
Else
    MsgBox "Ingrese un Valor Valido", vbExclamation, "Mensaje para el Administrador"
    Me.txtCantidad.SetFocus
End If
End Sub

Private Sub cmdConsultar_Click()
strCadena = "SELECT * FROM view_combo_detalle WHERE id_productoc='" & Me.DtcCombo.BoundText & "' and   ruc='" & KEY_RUC & "' "
Call llenarGrid_prod(Me.HfdDetalle, 1)
End Sub

Private Sub cmdImprimir_Click()
'strCadena = "SELECT id_compra,comprobante,fecha_registro,fecha_emision,fecha_cancelacion,id_proveedor,nproveedor,id_producto,nombre_prod ,numero_procedimientos,cantidad,c_unitario,total,operador,ruc  FROM view_compra_vista WHERE id_compra='" & Val(Me.txtIdCompra.Text) & "'"
strCadena = "call ADM_reporte_standart('1','" & Val(Me.txtIdCompra.Text) & "','" & KEY_RUC & "')"
Call ConfiguraRst(strCadena)
Ans = ShowMultiReport(rst, "RptCompra", , App.Path + "\Reportes\")

End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub cmdNuevo_Click()
Call nuevo
End Sub

Private Sub cmdProcesar_Click()
Dim in_costo_acum As Double
strCadena = "SELECT * FROM view_combo_detalle WHERE id_productoc='" & Me.DtcCombo.BoundText & "' and   ruc='" & KEY_RUC & "' "
Call ConfiguraRstA(strCadena)
If rstA.RecordCount > 0 Then
   rstA.MoveFirst
   in_costo_acum = 0
   Me.prog_indicador.Min = 1
   Me.prog_indicador.Max = rstA.RecordCount + 1
   For i = 0 To rstA.RecordCount - 1
        
        strCadena = "SELECT numero,serie FROM movimiento_compra WHERE id_doc='0090' and id_alm='" & Me.DtcAlmacen.BoundText & "' and ruc='" & KEY_RUC & "' ORDER BY numero DESC LIMIT 1"
        Call ConfiguraRstK(strCadena)
       If rstK.RecordCount > 0 Then
              in_numero = Format(Val(rstK("numero")) + 1, "00000000")
              in_serie = rstK("serie")
       Else
             
             strCadena = "SELECT * FROM almacen_comprobante WHERE ruc='" & KEY_RUC & "' and id_alm='" & Me.DtcAlmacen.BoundText & "' and id_doc='0090' LIMIT 1"
             Call ConfiguraRstlocal(strCadena)
             If rstLocal.RecordCount > 0 Then
               in_serie = rstLocal("serie")
               in_numero = formato_item(rstLocal("numero"), 8)
             Else
                MsgBox "CREE EL COMPROBANTE SALIDA A ALMACEN" + Chr(13) + "PARA ESTA SUCURSAL", vbInformation
                Exit Sub
             End If
             
        End If
        
        in_cta_compra = KEY_CTA_COMPRA_SOLES
        in_costo = get_costo_ultimo(rstA("id_producto"), Me.Dtpfecha.Value)
        in_total = in_costo * rstA("cantidad")
        in_costo_acum = in_costo_acum + in_total
        
        strCadena = "call P_insert_compra_ultimate('0090','" & Me.DtcAlmacen.BoundText & "','" & Format(Me.Dtpfecha.Value, "YYYY-mm-dd") & "','" & Format(Me.Dtpfecha.Value, "YYYY-mm-dd") & "','02'," & _
        "'03','--','00001','" & formato_item(Month(Me.Dtpfecha.Value), 2) & "','" & Year(Me.Dtpfecha.Value) & "','" & in_serie & "'," & _
        "'" & in_numero & "','6','" & KEY_RUC & "','" & KEY_EMPRESA & "','" & KEY_CAMBIO_VENTA & "'," & _
        "'0','" & in_valor_venta & "','" & in_igv & "','0','0','0','0','0','0','" & in_total & "','0'," & _
        " '" & KEY_USUARIO & "','OBSERVACION','01','" & get_periodo_actual(Me.Dtpfecha.Value) & "','" & in_cta_compra & "','" & KEY_USUARIO & "','0','0','0','0','" & KEY_RUC & "')"
        Call ConfiguraRstP(strCadena)
        
        id_compra = rstP(0)
                
        strCadena = "UPDATE movimiento_compra SET fecha_kardex='" & Format(Me.Dtpfecha.Value, "YYYY-mm-dd") & "' WHERE id_compra='" & id_compra & "'"
        Call ConfiguraRstP(strCadena)
        
                in_total = in_costo * rstA("cantidad")
                If KEY_CON_IGV = "si" Then
                    in_valor_venta = in_total
                    in_igv = 0
                Else
                    in_valor_venta = in_total
                    in_igv = 0
                End If
                
                strCadena = "INSERT INTO movimiento_compra_detalle(id_compra,id_producto,cantidad,c_unitario,dsto_soles,dsto_procentaje,total_descuento,valor_neto,isc,igv,retencion,otros,percepcion, " & _
                "valor_venta,exonerado,total,p_venta,p_costo,id_alm,detalle,incremento_fs,ruc) VALUES ('" & id_compra & "','" & Trim(rstA("id_producto")) & "','" & rstA("cantidad") * Val(Me.txtCantidad.Text) & "','" & Val(in_costo) & "'," & _
                "'0','0','0','" & in_valor_venta & "','" & Val(in_igv) & "','0', " & _
                "'0','0','0','" & in_valor_venta & "','0','" & Val(in_costo) * rstA("cantidad") * Val(Me.txtCantidad.Text) & "','" & Val(in_costo) & "','" & Val(in_costo) & "','" & Me.DtcAlmacen.BoundText & "','" & rstA("nombre_prod") & "','0','" & KEY_RUC & "')"
                CnBd.Execute (strCadena)
                
                Call put_registrar_produccion(rstA("id_producto"), rstA("cantidad"), in_costo)
               
               strCadena = "call put_kardex_stock_vitekey_v1('02','" & Format(Me.Dtpfecha.Value, "YYYY-mm-dd") & "','" & Val(id_compra) & "','0090','" & in_serie & "','" & in_numero & "','" & KEY_RUC & "','" & Trim(rstA("id_producto")) & "','" & rstA("cantidad") * Val(Me.txtCantidad.Text) & "','" & Val(in_costo) & "','" & Me.DtcAlmacen.BoundText & "','" & KEY_USUARIO & "','no','" & KEY_RUC & "')"
                
                CnBd.Execute (strCadena)
                Call put_actualizar_kardex_update(rstA("id_producto"), Me.DtcAlmacen.BoundText)
                rstA.MoveNext
                DoEvents
                Me.prog_indicador.Value = i + 1
   Next i
   
   
            
            in_total = in_costo_acum
                If KEY_CON_IGV = "si" Then
                    in_valor_venta = in_total / (1 + KEY_IGV)
                    in_igv = in_total - in_valor_venta
                Else
                    in_valor_venta = in_total
                    in_igv = 0
                End If
            
            
            strCadena = "call P_insert_compra_ultimate('0089','" & Me.DtcAlmacen.BoundText & "','" & Format(Me.Dtpfecha.Value, "YYYY-mm-dd") & "','" & Format(Me.Dtpfecha.Value, "YYYY-mm-dd") & "','02'," & _
            "'03','--','00001','" & formato_item(Month(Me.Dtpfecha.Value), 2) & "','" & Year(Me.Dtpfecha.Value) & "','" & in_serie & "'," & _
            "'" & in_numero & "','6','" & KEY_RUC & "','" & KEY_EMPRESA & "','" & KEY_CAMBIO_VENTA & "'," & _
            "'0','" & Val(in_valor_venta) & "','" & Val(in_igv) & "','0','0','0','0','0','0','" & Val(in_total) & "','" & Val(in_total) & "'," & _
            " '" & KEY_USUARIO & "','OBSERVACION','01','" & get_periodo_actual(Me.Dtpfecha.Value) & "','" & in_cta_compra & "','" & KEY_USUARIO & "','0','0','0','0','" & KEY_RUC & "')"
            Call ConfiguraRstP(strCadena)
            id_compra = rstP(0)
            Me.txtIdCompra.Text = id_compra
           
           
            strCadena = "UPDATE movimiento_compra SET fecha_kardex='" & Format(Me.Dtpfecha.Value, "YYYY-mm-dd") & "' WHERE id_compra='" & id_compra & "'"
            Call ConfiguraRstP(strCadena)
        
        
           strCadena = "INSERT INTO movimiento_compra_detalle(id_compra,id_producto,cantidad,c_unitario,dsto_soles,dsto_procentaje,total_descuento,valor_neto,isc,igv,retencion,otros,percepcion, " & _
           "valor_venta,exonerado,total,p_venta,p_costo,id_alm,detalle,incremento_fs,ruc) VALUES ('" & id_compra & "','" & Trim(Me.DtcCombo.BoundText) & "','" & Val(Me.txtCantidad.Text) & "','" & Val(in_total) & "'," & _
           "'0','0','0','" & Val(in_total) & "','0','0', " & _
           "'0','0','0','" & Val(in_total) & "','0','" & Val(in_total) & "','" & Val(get_precio_venta_now(Me.DtcCombo.BoundText)) & "','" & Val(in_total) & "','" & Me.DtcAlmacen.BoundText & "','" & Me.DtcCombo.Text & "','0','" & KEY_RUC & "')"
           CnBd.Execute (strCadena)
           
           
           strCadena = "UPDATE combo_detalle_produccion SET id_compra='" & id_compra & "' WHERE dni_save='" & KEY_USUARIO & "' and  id_compra=0 and ruc='" & KEY_RUC & "'"
           CnBd.Execute (strCadena)
           
           
           strCadena = "call put_kardex_stock_vitekey_v1('02','" & Format(Me.Dtpfecha.Value, "YYYY-mm-dd") & "','" & Val(id_compra) & "','0089','" & in_serie & "','" & in_numero & "','" & KEY_RUC & "','" & Trim(Me.DtcCombo.BoundText) & "','" & Val(Me.txtCantidad.Text) & "','" & Val(in_total) & "','" & Me.DtcAlmacen.BoundText & "','" & KEY_USUARIO & "','no','" & KEY_RUC & "')"
           CnBd.Execute (strCadena)
           
           
           
           
           MsgBox "COMBO REALIZADO EXITOSAMENTE", vbInformation
   
   
End If




                
                
                
                
                
                
            


End Sub

Private Sub put_registrar_produccion(ByVal in_producto As String, ByVal in_cantidad As Double, ByVal in_precio As Double)

strCadena = "INSERT INTO combo_detalle_produccion(`fecha`,`id_producto_combo`,cantidad_combo,`id_producto`,`cantidad`,`precio`,`dni_save`,`ruc`)VALUES " & _
"('" & Format(Me.Dtpfecha.Value, "YYYY-mm-dd") & "','" & Me.DtcCombo.BoundText & "','" & Val(Me.txtCantidad.Text) & "','" & in_producto & "','" & in_cantidad & "','" & in_precio & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
CnBd.Execute (strCadena)




End Sub

Private Sub DtcCombo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call LLENA
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
    Unload Me
End If
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 500
Me.Dtpfecha.Value = KEY_FECHA


  strCadena = "SELECT id_producto as Codigo, CONCAT(id_producto,'-',nombre_prod) as Descripcion FROM producto WHERE id_combo='si' AND ruc='" & KEY_RUC & "' ORDER BY nombre_prod"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcCombo)
   strCadena = "SELECT id_alm as Codigo, descripcion as Descripcion FROM almacen WHERE ruc='" & KEY_RUC & "'  ORDER BY id_alm ASC"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcAlmacen)
  Me.DtcAlmacen.BoundText = KEY_ALM
  Call nuevo
  
End Sub
Private Sub nuevo()
  
 strCadena = "SELECT id_combo FROM combo WHERE ruc='" & KEY_RUC & "' ORDER BY id_combo DESC"
  Call ConfiguraRst(strCadena)
  strCombo = formato_item(ConsultaUltimoRegistro("combo", "id_combo", "ruc", KEY_RUC), 6)
  Me.TxtCombo.Text = strCombo
  Set rst = Nothing
  Me.txtCantidad.Text = 0
  Me.lblAnulado.Visible = False
  Me.txtCantidad.Locked = False
  
End Sub
Public Sub LLENA()

strCadena = "SELECT * FROM view_combo_detalle WHERE id_productoc='" & Me.DtcCombo.BoundText & "' and   ruc='" & KEY_RUC & "'"
Call llenarGrid_prod(Me.HfdDetalle, 0)



End Sub
Private Function get_costo_ultimo(ByVal in_producto As String, ByVal in_fecha As String) As Double

strCadena = "SELECT funct_costo_final('" & in_producto & "','" & Format(in_fecha, "YYYY-mm-dd") & "','" & Me.DtcAlmacen.BoundText & "','" & KEY_RUC & "')"
Call ConfiguraRstL(strCadena)
get_costo_ultimo = rstL(0)
End Function

Public Sub llenarGrid_prod(ByVal Grilla As MSHFlexGrid, ByVal in_flag As Integer)
On Error GoTo salir
Dim tTotal As Double
Dim in_costo As Double
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 0
    
    Exit Sub
End If
   
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
            ' -- almacena el ancho de los campos
           Grilla.ColWidth(0) = 1000
           Grilla.ColWidth(1) = 6000
           Grilla.ColWidth(2) = 1500
           Grilla.ColWidth(3) = 1200
           Grilla.ColWidth(4) = 1200
           Grilla.ColWidth(5) = 1500
           
         Next
        cabecera = "CODIGO" & vbTab & "DESCRIPCION ARTICULO" & vbTab & "UND" & vbTab & "CANT" & vbTab & "PRECIO COSTO" & vbTab & "TOTAL"
        Grilla.AddItem cabecera
         For k = 0 To 5
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        tTotal = 0
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
            If in_flag = 0 Then
                in_costo = 0 'get_costo_ultimo(rst("id_producto"), Me.DtpFecha.Value)
            Else
                in_costo = get_costo_ultimo(rst("id_producto"), Me.Dtpfecha.Value)
            End If
            
            Fila = rst("id_producto") & vbTab & rst("nombre_prod") & vbTab & rst("descripcion") & vbTab & Format(rst("cantidad"), "#,##0.00000000") & vbTab & Format(in_costo, "#,##0.0000000") & vbTab & Format(rst("cantidad") * in_costo, "#,##0.00000000")
            Grilla.AddItem Fila
            tTotal = tTotal + rst("cantidad") * in_costo
            For k = 3 To 5
                Grilla.col = k
                Grilla.Row = i + 1
                Grilla.CellBackColor = &HC0FFC0
            Next k
            Fila = ""
            rst.MoveNext
            DoEvents
    Next i
      Fila = "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & Format(tTotal, "#,##0.00")
    Grilla.AddItem Fila
     For k = 3 To 5
            Grilla.col = k
            Grilla.Row = i + 1
            Grilla.CellBackColor = &HC0FFFF
    Next k
  
  Grilla.Row = 1
  Grilla.col = 0
  Grilla.ColSel = 1
  Grilla.RowSel = 1
  
  
  Exit Sub
salir: MsgBox "Ocurrio un Error en la Conexion", vbInformation, "Mensaje para el usuario"
Set rst = Nothing

End Sub


Private Sub HfdDetalle_DblClick()
If Me.HfdDetalle.Rows > 0 Then
    FrmComboCantidad.Show
End If
End Sub


Public Function GeneraCodigo1(ByVal longitud As Integer) As String
Dim X As Integer
Dim Formato As String
  Formato = ""
  For X = 1 To longitud
    Formato = Formato + "0"
  Next X
   
  If (rst2.BOF And rst2.EOF) Then
    StrNumero = Format(str(Val(Formato) + 1), Formato)
  Else
    StrNumero = Format(Trim(str(Val(Right(rst2(0), longitud + 1)) + 1)), Formato)
  End If
  Set rst2 = Nothing
  GeneraCodigo1 = Gencodigo + StrNumero
  Gencodigo = ""

End Function



Private Sub txtBuscar_Change()
strCadena = "SELECT id_producto as Codigo, nombre_prod as Descripcion FROM producto WHERE id_combo='si' and nombre_prod LIKE '%" & Trim(Me.txtBuscar.Text) & "%' AND ruc='" & KEY_RUC & "' ORDER BY nombre_prod"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcCombo)
End Sub

Private Sub TxtBuscar_KeyPress(KeyAscii As Integer)
On Error GoTo salir
If KeyAscii = 13 Then
    Call Me.DtcCombo.SetFocus
End If
Exit Sub
salir:
End Sub
