VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form FrmMiscuentas 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9240
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   20145
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9240
   ScaleWidth      =   20145
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmajustebanco 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "AJUSTE TIPO CAMBIO"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   2055
      Left            =   13200
      TabIndex        =   8
      Top             =   2280
      Visible         =   0   'False
      Width           =   5415
      Begin VB.TextBox txtid_banco 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   3240
         TabIndex        =   14
         Top             =   960
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtTipocambio 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   1320
         TabIndex        =   12
         Top             =   960
         Width           =   1815
      End
      Begin MSDataListLib.DataCombo DtcPeriodo 
         Height          =   330
         Left            =   1320
         TabIndex        =   10
         Top             =   480
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         ForeColor       =   4194304
         Text            =   "DataCombo1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VitekeySoft.ChameleonBtn cmdProcesar 
         Height          =   435
         Left            =   1320
         TabIndex        =   13
         Top             =   1440
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   767
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
         BCOL            =   16777215
         BCOLO           =   16777215
         FCOL            =   8388608
         FCOLO           =   8388608
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmMiscuentas.frx":0000
         PICN            =   "FrmMiscuentas.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   4920
         Picture         =   "FrmMiscuentas.frx":2601
         Top             =   120
         Width           =   240
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00808080&
         Height          =   2055
         Left            =   0
         Top             =   0
         Width           =   5415
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "T.CAMBIO :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   195
         Left            =   360
         TabIndex        =   11
         Top             =   960
         Width           =   750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "PERIODO :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   195
         Left            =   360
         TabIndex        =   9
         Top             =   480
         Width           =   690
      End
   End
   Begin VitekeySoft.ChameleonBtn cmdnuevo 
      Height          =   855
      Left            =   18720
      TabIndex        =   2
      Top             =   600
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1508
      BTYPE           =   5
      TX              =   "NUEVA "
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bahnschrift SemiLight SemiConde"
         Size            =   8.25
         Charset         =   0
         Weight          =   350
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   12582912
      FCOLO           =   12582912
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmMiscuentas.frx":54A5
      PICN            =   "FrmMiscuentas.frx":54C1
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfgDetalle 
      Height          =   8295
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   18255
      _ExtentX        =   32200
      _ExtentY        =   14631
      _Version        =   393216
      ForeColor       =   8388608
      Rows            =   0
      FixedRows       =   0
      FixedCols       =   0
      ForeColorFixed  =   8388608
      BackColorBkg    =   16777215
      GridColor       =   12582912
      SelectionMode   =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bahnschrift SemiLight SemiConde"
         Size            =   9.75
         Charset         =   0
         Weight          =   350
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VitekeySoft.ChameleonBtn cmdEditar 
      Height          =   855
      Left            =   18720
      TabIndex        =   3
      Top             =   1560
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1508
      BTYPE           =   5
      TX              =   "EDITAR"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bahnschrift SemiLight SemiConde"
         Size            =   8.25
         Charset         =   0
         Weight          =   350
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   12582912
      FCOLO           =   12582912
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmMiscuentas.frx":5913
      PICN            =   "FrmMiscuentas.frx":592F
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmddetalle 
      Height          =   855
      Left            =   18720
      TabIndex        =   4
      Top             =   2520
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1508
      BTYPE           =   5
      TX              =   "DETALLE"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bahnschrift SemiLight SemiConde"
         Size            =   8.25
         Charset         =   0
         Weight          =   350
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   12582912
      FCOLO           =   12582912
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmMiscuentas.frx":8C05
      PICN            =   "FrmMiscuentas.frx":8C21
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdEliminar 
      Height          =   975
      Left            =   18720
      TabIndex        =   5
      Top             =   6720
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1720
      BTYPE           =   5
      TX              =   "ELIMINAR"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bahnschrift SemiLight SemiConde"
         Size            =   8.25
         Charset         =   0
         Weight          =   350
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   12582912
      FCOLO           =   12582912
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmMiscuentas.frx":AE7A
      PICN            =   "FrmMiscuentas.frx":AE96
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdCerrar 
      Height          =   975
      Left            =   18720
      TabIndex        =   6
      Top             =   7920
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1720
      BTYPE           =   5
      TX              =   "CERRAR"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bahnschrift SemiLight SemiConde"
         Size            =   8.25
         Charset         =   0
         Weight          =   350
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   12582912
      FCOLO           =   12582912
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmMiscuentas.frx":D2E0
      PICN            =   "FrmMiscuentas.frx":D2FC
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdAjuste 
      Height          =   915
      Left            =   18720
      TabIndex        =   7
      Top             =   3480
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1614
      BTYPE           =   5
      TX              =   "AJUSTE TC"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bahnschrift SemiLight SemiConde"
         Size            =   8.25
         Charset         =   0
         Weight          =   350
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   8388608
      FCOLO           =   8388608
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmMiscuentas.frx":10323
      PICN            =   "FrmMiscuentas.frx":1033F
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lbltc 
      BackStyle       =   0  'Transparent
      Caption         =   "TIPO CAMBIO :"
      BeginProperty Font 
         Name            =   "Bahnschrift SemiLight SemiConde"
         Size            =   9
         Charset         =   0
         Weight          =   350
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   375
      TabIndex        =   1
      Top             =   165
      Width           =   6480
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   6  'Inside Solid
      Height          =   375
      Left            =   240
      Top             =   120
      Width           =   12375
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   9240
      Left            =   0
      Top             =   0
      Width           =   20145
   End
End
Attribute VB_Name = "FrmMiscuentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Procedencia As EnumProcede
Public Sub actualizar()
Call llenarGrid(Me.HfgDetalle, Me)
End Sub

Private Sub HfgMarcas_Click()
If HfgMarcas.Row > 0 Then
    TlbAcciones.Buttons(KEY_UPDATE).Enabled = True
    TlbAcciones.Buttons(KEY_DELETE).Enabled = True
  End If
End Sub



Private Sub cmdAjuste_Click()
 
  
  strCadena = "SELECT id as Codigo,CONCAT(Nombre,'-',Ejercicio) as Descripcion FROM con_periodo order by Ejercicio DESC,mes DESC "
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcPeriodo)
  Me.DtcPeriodo.BoundText = get_periodo_actual(KEY_FECHA)
  
 

  Me.frmajustebanco.Visible = True
  
  
End Sub

Private Sub cmdCerrar_Click()
Unload Me
End Sub

Private Sub cmdDetalle_Click()
 
 frmCajaEgreso.Show
 frmCajaEgreso.TxtidCuenta.Text = Val(Me.HfgDetalle.TextMatrix(Me.HfgDetalle.Row, 0))
 frmCajaEgreso.lblcuenta.Caption = Me.HfgDetalle.TextMatrix(Me.HfgDetalle.Row, 2)
 
End Sub

Private Sub cmdEditar_Click()
 Procedencia = modificar
 FrmMisCuentasDet.Show
End Sub

Private Sub cmdEliminar_Click()
If MsgBox(MSGELIMINAR, vbQuestion + vbYesNo, MSGVERIFICACION) = vbYes Then
        
        strCadena = "SELECT * FROM mis_cuentas_det WHERE id_cuenta='" & Val(Me.HfgDetalle.TextMatrix(Me.HfgDetalle.Row, 0)) & "' AND ruc='" & KEY_RUC & "'"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount < 1 Then
            strCadena = "DELETE FROM mis_cuentas WHERE id_cuenta='" & Val(Me.HfgDetalle.TextMatrix(Me.HfgDetalle.Row, 0)) & "' AND ruc='" & KEY_RUC & "'"
            Call Execute_Sql(strCadena)
            Call actualizar
        Else
            MsgBox "Imposible Eliminar esta Cuenta" + Chr(13) + "Tiene Movimientos Relacionados", vbInformation, "Mensaje para el Usuario"
        End If
      End If
End Sub

Private Sub cmdnuevo_Click()
      Procedencia = Nuevos
      FrmMisCuentasDet.Show
End Sub

Private Sub cmdProcesar_Click()
    
    strCadena = "SELECT * FROM mis_cuentas WHERE id_cuenta='" & Val(Me.HfgDetalle.TextMatrix(Me.HfgDetalle.Row, 0)) & "' and ruc='" & KEY_RUC & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        strCadena = "CALL CON_InsertaAsiento_AjusteTC_Banco('" & KEY_RUC & "','" & Me.DtcPeriodo.BoundText & "','" & rst("cuenta_ctble") & "','" & Val(Me.TxtTipoCambio.Text) & "','" & KEY_USUARIO & "')"
        CnBd.Execute (strCadena)
    Else
        MsgBox "SELECCIONE UNA CUENTA", vbInformation
        Exit Sub
    End If
    
    MsgBox "AJUSTE DE TIPO CAMBIO REALIZADO", vbInformation, KEY_VENDEDOR
End Sub

Private Sub DtcPeriodo_Change()
 Dim in_fecha As Date
strCadena = "SELECT * FROM con_periodo WHERE id='" & Me.DtcPeriodo.BoundText & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    in_fecha = DateSerial(rst("Ejercicio"), rst("mes") + 1, 0)
    Me.TxtTipoCambio.Text = get_tipo_cambio_dia(in_fecha, "valor_venta")
End If


     
   '
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 100

Me.lbltc.Caption = "TIPO DE CAMBIO AL :" + Space(2) + KEY_FECHA + Space(3) + "***** " + str(KEY_CAMBIO_COMPRA) + " *****"
Call actualizar




End Sub

Private Sub TlbAcciones_ButtonClick(ByVal Button As MSComctlLib.Button)
  Select Case Button.key
    Case KEY_NEW
     
    Case KEY_UPDATE
      
       Exit Sub
    Case KEY_ACTUALIZAR
       
    Case KEY_DELETE
      
    Case KEY_EXIT
        Unload Me
  End Select
End Sub
Private Sub llenarGrid(ByVal Grilla As MSHFlexGrid, ByVal Formulario As Form)
On Error GoTo salir
Dim tTotalS As Double, tTotalD As Double
Dim totalD As Double, totalS As Double, simbolo As String
tTotalS = 0
tTotalD = 0
totalS = 0
totalD = 0

strCadena = "call ADM_auditoria_empresa('4','2','3','4','5','6','7','8','" & KEY_RUC & "')"
'strCadena = "SELECT * FROM  view_mis_cuentas WHERE Ejercicio='" & Year(KEY_FECHA) & "' and  ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
 
If rst.RecordCount < 1 Then
    Grilla.Rows = 0
    
    Exit Sub

End If
  
    Grilla.Rows = 0
      ' Me.HfdGrilla.Rows = rst.RecordCount - 2
       ReDim arrColWidth(1 To rst.Fields.Count)
       
            ' -- almacena el ancho de los campos
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 1000
           Grilla.ColWidth(2) = 4500
           Grilla.ColWidth(3) = 1400
           Grilla.ColWidth(4) = 4500
           Grilla.ColWidth(5) = 1400
           Grilla.ColWidth(6) = 2500
           Grilla.ColWidth(7) = 2500
       
        cabecera = "IDCUENTA" & vbTab & "ITEM" & vbTab & "CUENTA" & vbTab & "MONEDA" & vbTab & "Nº CUENTA BANCARIA" & vbTab & "T.CAMBIO" & vbTab & " SALDO. (US$)" & vbTab & "   SALDO. (S/.)"
        Grilla.AddItem cabecera
                            For k = 1 To 7
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
                            Next k
                            
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
         
          
            
                If rst("id_moneda") = "00001" Then
                    tTotalS = rst("saldo")
                    tTotalD = rst("saldo") / KEY_CAMBIO_COMPRA
                Else
                    tTotalS = rst("saldo") * KEY_CAMBIO_COMPRA
                    tTotalD = rst("saldo")
                End If
            
            
            totalS = totalS + tTotalS
            totalD = totalD + tTotalD
            Fila = rst("id_cuenta") & vbTab & Format(str(i + 1), "0000") & vbTab & rst("descripcion") & vbTab & rst("moneda") & vbTab & rst("numero_cuenta") & vbTab & KEY_CAMBIO_COMPRA & vbTab & Format(tTotalD, "#,##0.00") & vbTab & Format(tTotalS, "#,##0.00")
            Grilla.AddItem Fila
             For l = 5 To 7
                                Grilla.col = l
                                Grilla.Row = i + 1
                                Grilla.CellBackColor = &H80FF&
            Next l
            rst.MoveNext
    Next i
    cabecera = "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & " :: SALDOS ::" & vbTab & "(US$)" + Space(2) + Format(totalD, "#,##0.00") & vbTab & "(S/.)" + Space(2) + Format(totalS, "#,##0.00")
    Grilla.AddItem cabecera
    For l = 5 To 7
                                Grilla.col = l
                                Grilla.Row = i + 1
                                Grilla.CellBackColor = &H80FF&
                                
    Next l
    
    Me.HfgDetalle.ColAlignment(4) = 1
    
    Me.cmdEditar.Enabled = False
    Me.cmdEliminar.Enabled = False
    
Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub





Private Sub HfgDetalle_SelChange()
If Val(Me.HfgDetalle.TextMatrix(Me.HfgDetalle.Row, 0)) > 0 Then
    Me.cmdEditar.Enabled = True
    Me.cmdEliminar.Enabled = True
    Me.cmdDetalle.Enabled = True
    Me.cmdAjuste.Enabled = True
Else
    Me.cmdEditar.Enabled = False
    Me.cmdEliminar.Enabled = False
    Me.cmdDetalle.Enabled = False
    Me.cmdAjuste.Enabled = False
End If
End Sub

Private Sub Image1_Click()
Me.frmajustebanco.Visible = False
End Sub
