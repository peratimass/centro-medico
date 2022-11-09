VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form FrmPagoCuotaDeuda 
   BorderStyle     =   0  'None
   Caption         =   "PAGO A CUENTA"
   ClientHeight    =   7215
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10980
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7215
   ScaleWidth      =   10980
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Comprobantes"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   2655
      Left            =   240
      TabIndex        =   17
      Top             =   2040
      Width           =   10575
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfgSalidas 
         Height          =   2175
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   10335
         _ExtentX        =   18230
         _ExtentY        =   3836
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
   End
   Begin VB.TextBox TxtNumeroCredito 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   7515
      MaxLength       =   80
      TabIndex        =   14
      Text            =   "0000000000"
      Top             =   840
      Width           =   1455
   End
   Begin VB.TextBox TxtSerieCredito 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   6390
      MaxLength       =   80
      TabIndex        =   13
      Text            =   "0000"
      Top             =   840
      Width           =   975
   End
   Begin VB.TextBox TxtEntidad 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   6360
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   1440
      Width           =   4335
   End
   Begin VB.TextBox TxtMonto 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   4230
      TabIndex        =   0
      Top             =   5040
      Width           =   1515
   End
   Begin VB.TextBox TxtDeuda 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Left            =   7440
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   5040
      Width           =   1335
   End
   Begin VB.TextBox TxtCodEntidad 
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   5520
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   1440
      Width           =   735
   End
   Begin VB.TextBox TxtSaldo 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   4230
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   5490
      Width           =   1515
   End
   Begin MSComctlLib.ImageList ImgIconos 
      Left            =   2160
      Top             =   4080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPagoCuotaDeuda.frx":0000
            Key             =   "(Aceptar)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPagoCuotaDeuda.frx":031C
            Key             =   "(Imprimir)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPagoCuotaDeuda.frx":03A9
            Key             =   "(Eliminar)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPagoCuotaDeuda.frx":0809
            Key             =   "(Inicio)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPagoCuotaDeuda.frx":0C69
            Key             =   "(Modificar)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPagoCuotaDeuda.frx":0F85
            Key             =   "(Nuevo)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPagoCuotaDeuda.frx":13E5
            Key             =   "(Quitar)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPagoCuotaDeuda.frx":1701
            Key             =   "(Salir)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPagoCuotaDeuda.frx":1B61
            Key             =   "(Red)"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPagoCuotaDeuda.frx":1FC1
            Key             =   "(Grabar)"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPagoCuotaDeuda.frx":28A1
            Key             =   "(Agregar)"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPagoCuotaDeuda.frx":2BBD
            Key             =   "(Buscar)"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPagoCuotaDeuda.frx":2ED9
            Key             =   "(Cancelar)"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar ClbAcciones 
      Height          =   870
      Left            =   7500
      TabIndex        =   5
      Top             =   6120
      Width           =   3315
      _ExtentX        =   5847
      _ExtentY        =   1535
      BandCount       =   1
      ForeColor       =   -2147483635
      ImageList       =   "ImgIconos"
      FixedOrder      =   -1  'True
      VariantHeight   =   0   'False
      EmbossPicture   =   -1  'True
      _CBWidth        =   3315
      _CBHeight       =   870
      _Version        =   "6.0.8169"
      Child1          =   "TlbAcciones"
      MinHeight1      =   810
      Width1          =   3180
      FixedBackground1=   0   'False
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar TlbAcciones 
         Height          =   810
         Left            =   30
         TabIndex        =   6
         Top             =   30
         Width           =   3195
         _ExtentX        =   5636
         _ExtentY        =   1429
         ButtonWidth     =   1323
         ButtonHeight    =   1429
         Style           =   1
         ImageList       =   "ImgIconos"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   5
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Grabar"
               Key             =   "(Grabar)"
               Object.ToolTipText     =   "Grabar Ctrl+G"
               ImageKey        =   "(Grabar)"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Eliminar"
               Key             =   "(Eliminar)"
               ImageKey        =   "(Eliminar)"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Salir"
               Key             =   "(Salir)"
               ImageKey        =   "(Salir)"
            EndProperty
         EndProperty
         OLEDropMode     =   1
      End
   End
   Begin MSComCtl2.DTPicker DtpFechaPago 
      Height          =   360
      Left            =   1710
      TabIndex        =   7
      Top             =   1290
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   0
      CalendarForeColor=   255
      Format          =   121176067
      CurrentDate     =   39535
   End
   Begin MSDataListLib.DataCombo DtcAlmacen 
      Height          =   360
      Left            =   480
      TabIndex        =   15
      Top             =   360
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   635
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      BackColor       =   12648447
      ForeColor       =   8388608
      Text            =   "DataCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo DtcTipoDoc 
      Height          =   360
      Left            =   6360
      TabIndex        =   16
      Top             =   360
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   635
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      BackColor       =   12648447
      ForeColor       =   8388608
      Text            =   "DataCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   2
      Height          =   7215
      Left            =   0
      Top             =   0
      Width           =   10980
   End
   Begin VB.Label LblEntidad 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Entidad :"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   5640
      TabIndex        =   12
      Top             =   1200
      Width           =   630
   End
   Begin VB.Label LblDeuda 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Deuda:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   270
      Left            =   6600
      TabIndex        =   11
      Top             =   5040
      Width           =   705
   End
   Begin VB.Label LblFechaPago 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Fecha Pago:"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   480
      TabIndex        =   10
      Top             =   1320
      Width           =   915
   End
   Begin VB.Label LblMonto 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Monto a Pagar:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   270
      Left            =   2685
      TabIndex        =   9
      Top             =   5040
      Width           =   1515
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Saldo."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   3390
      TabIndex        =   8
      Top             =   5490
      Width           =   585
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      Height          =   1695
      Left            =   240
      Top             =   240
      Width           =   10575
   End
   Begin VB.Shape ShpDatos 
      BackColor       =   &H00A56E32&
      FillColor       =   &H00DFDFE0&
      FillStyle       =   0  'Solid
      Height          =   1140
      Left            =   240
      Top             =   4800
      Width           =   10575
   End
End
Attribute VB_Name = "FrmPagoCuotaDeuda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Public EnumFrmPago As EnumBuscarDocumento
Public Procedencia As EnumProcede
Dim DblTotal As Double
Dim doc_Tienda As String



Private Sub cmdEliminar_Click()
If MsgBox("Esta Seguro de Eliminar", vbQuestion + vbYesNo, KEY_EMPRESA) = vbYes Then
    'delete from DOCUME
End If
End Sub

Private Sub Command1_Click()
Dim Deuda As Double
Dim MontoPagar As Double
Dim Saldo As Double
Deuda = Me.TxtDeuda.Text
MontoPagar = Me.TxtMonto.Text
Saldo = Deuda - MontoPagar
Me.txtsaldo.Text = Format(Saldo, "#,##0.00")
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If Shift = 2 And KeyCode = Asc("G") Then
    Call Save
  End If
   
End Sub
Private Function Nuevo_pago() As String
strCadena = "SELECT cDocumentoVenta FROM DocumentoVenta WHERE (sSerie='" & Trim(Me.TxtSerieCredito.Text) & "' AND doc_cod='" & Trim(KEY_PAGOCRE) & "') ORDER BY intDocumentoVenta DESC"
    Call ConfiguraRst(strCadena)
    Nuevo_pago = GeneraCodigo(10)
    Set rst = Nothing
End Function
Private Sub Save()
Dim TotalDeuda As Single
Dim MontoCancelar As Single
Dim Saldo As Single
Dim rstP As New ADODB.Recordset

TotalDeuda = Val(Me.TxtDeuda.Text)
MontoCancelar = Val(Me.TxtMonto.Text)

strCadena = "SELECT DocumentoVenta.doc_cod,DocumentoVenta.sSerie," & _
    "DocumentoVenta.cDocumentoVenta ,  DocumentoVenta.Saldo AS SALDO FROM DocumentoVenta  WHERE " & _
    "DocumentoVenta.idFormaPago='" & Trim(KEY_CREDITO) & "' AND DocumentoVenta.anulado='F' AND saldo>0 AND DocumentoVenta.cPersona='" & Trim(Me.TxtCodEntidad.Text) & "'ORDER BY saldo ASC"
    rstP.Open strCadena, CnBd, adOpenKeyset, adLockOptimistic
     rstP.MoveFirst
     
    
     For i = 0 To rstP.RecordCount - 1
         If (MontoCancelar >= rstP("saldo")) Then
            MontoCancelar = MontoCancelar - rstP("saldo")
            strCadena = "INSERT INTO DocumentoVenta(id_documentoventa,cDocumentoVenta,doc_cod,Alm_cod,sSerie,cPersona,Persona,idFormaPago," & _
            "dEmisionVenta,nTotalVenta,FechaProceso,intDocumentoVenta,Anulado,id_usuario)" & _
            "VALUES ('" & Trim(Nuevo_pago) & "','" & Trim(Nuevo_pago) & "','" & KEY_PAGOCRE & "'," & _
            "'" & Trim(Me.DtcAlmacen.BoundText) & "','" & Trim(Me.TxtSerieCredito.Text) & "'," & _
            "'" & Trim(Me.TxtCodEntidad.Text) & "','" & Trim(Me.TxtEntidad.Text) & "','" & Trim(KEY_CONTADO) & "'," & _
            "'" & CVDate(KEY_FECHA) & "','" & rstP("saldo") & "','" & CVDate(Date) & "','" & Val(Me.TxtNumeroCredito.Text) & "','" & KEY_ANULF & "','" & KEY_USUARIO & "')"
            CnBd.Execute (strCadena)
             
            
           strCadena = "UPDATE DocumentoVenta SET Saldo='0',estado='Cancelado' WHERE cPersona='" & Trim(Me.TxtCodEntidad.Text) & "' AND " & _
           "sSerie='" & Trim(rstP("sSerie")) & "' AND cDocumentoVenta='" & Trim(rstP("cDocumentoVenta")) & "' AND doc_cod='" & Trim(rstP("doc_cod")) & "' AND Alm_cod='" & Trim(Me.DtcAlmacen.BoundText) & "'"
           CnBd.Execute (strCadena)
            
            
            monto_letras = UCase(EnLetras(rstP("saldo")))
            strCadena = "INSERT INTO DetallePagoCreditos(alm_cod,doc_cod,Serie,Numero,cPersona,FechaPago,Monto,TipoDocVenta,SerieVenta,NUmeroVenta,letras,id_usuario)" & _
            "VALUES ('" & Trim(Me.DtcAlmacen.BoundText) & "','" & Trim(Me.DtcTipoDoc.BoundText) & "','" & Trim(Me.TxtSerieCredito.Text) & "'," & _
            "'" & Trim(Me.TxtNumeroCredito.Text) & "','" & Trim(Me.TxtCodEntidad.Text) & "','" & CVDate(KEY_FECHA) & "','" & rstP("saldo") & "','" & Trim(rstP("doc_cod")) & "','" & Trim(rstP("sSerie")) & "','" & Trim(rstP("cDocumentoVenta")) & "','" & Trim(monto_letras) & "','" & Trim(KEY_USUARIO) & "')"
            CnBd.Execute (strCadena)
             
            Me.TxtNumeroCredito.Text = Nuevo_pago
        Else
            If MontoCancelar > 0 Then
                
            
            Saldo = rstP("saldo") - MontoCancelar
            strCadena = "INSERT INTO DocumentoVenta(id_documentoventa,cDocumentoVenta,doc_cod,Alm_cod,sSerie,cPersona,Persona,idFormaPago," & _
            "dEmisionVenta,nTotalVenta,FechaProceso,intDocumentoVenta,Anulado,id_usuario)" & _
            "VALUES ('" & Trim(Nuevo_pago) & "','" & Trim(Nuevo_pago) & "','" & KEY_PAGOCRE & "'," & _
            "'" & Trim(Me.DtcAlmacen.BoundText) & "','" & Trim(Me.TxtSerieCredito.Text) & "'," & _
            "'" & Trim(Me.TxtCodEntidad.Text) & "','" & Trim(Me.TxtEntidad.Text) & "','" & Trim(KEY_CONTADO) & "'," & _
            "'" & CVDate(KEY_FECHA) & "','" & MontoCancelar & "','" & CVDate(Date) & "','" & Val(Me.TxtNumeroCredito.Text) & "','" & KEY_ANULF & "','" & KEY_USUARIO & "')"
            CnBd.Execute (strCadena)
             
            
           strCadena = "UPDATE DocumentoVenta SET Saldo='" & Saldo & "',estado='Pendiente' WHERE cPersona='" & Trim(Me.TxtCodEntidad.Text) & "' AND " & _
           "sSerie='" & Trim(rstP("sSerie")) & "' AND cDocumentoVenta='" & Trim(rstP("cDocumentoVenta")) & "' AND doc_cod='" & Trim(rstP("doc_cod")) & "' AND Alm_cod='" & Trim(Me.DtcAlmacen.BoundText) & "'"
           CnBd.Execute (strCadena)
            
            
            monto_letras = UCase(EnLetras(MontoCancelar))
            strCadena = "INSERT INTO DetallePagoCreditos(alm_cod,doc_cod,Serie,Numero,cPersona,FechaPago,Monto,TipoDocVenta,SerieVenta,NUmeroVenta,letras,id_usuario)" & _
            "VALUES ('" & Trim(Me.DtcAlmacen.BoundText) & "','" & Trim(Me.DtcTipoDoc.BoundText) & "','" & Trim(Me.TxtSerieCredito.Text) & "'," & _
            "'" & Trim(Me.TxtNumeroCredito.Text) & "','" & Trim(Me.TxtCodEntidad.Text) & "','" & CVDate(KEY_FECHA) & "','" & MontoCancelar & "','" & Trim(rstP("doc_cod")) & "','" & Trim(rstP("sSerie")) & "','" & Trim(rstP("cDocumentoVenta")) & "','" & Trim(monto_letras) & "','" & Trim(KEY_USUARIO) & "')"
            CnBd.Execute (strCadena)
             
            MontoCancelar = 0
             Me.TxtNumeroCredito.Text = Nuevo_pago
            End If
             
        End If
        rstP.MoveNext
     Next i
 Set rstP = Nothing
    
     
        
     Me.TlbAcciones.Buttons(KEY_SAVE).Enabled = False

    Exit Sub
End Sub
Private Sub Form_Load()

CenterForm Me
Me.Top = 200
doc_Tienda = "V"
 strCadena = "SELECT Alm_cod as Codigo, Alm_des as Descripcion FROM Almacen " & _
  " ORDER BY Alm_des"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcAlmacen)
   Set rst = Nothing
  
 
   
  Set rst = Nothing
  
  strCadena = "SELECT doc_cod as Codigo, doc_abrev as Descripcion FROM Comprobantes " & _
  " WHERE doc_tienda='" & doc_Tienda & "' ORDER BY doc_abrev"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcTipoDoc)
  Me.DtcTipoDoc.BoundText = "0010"
  
  
  Set rst = Nothing
  Me.DtpFechaPago.Value = KEY_FECHA
Call LLENA
End Sub
Private Sub LLENA()
Dim TipoDoc As String
        Dim serie As String
        Dim Numero As String
        Dim Persona As String
        Dim idalmacen As String
If FrmPagoCredito.Procedencia = nuevo Then
        FrmPagoCredito.HfgFacturas.col = 2
        TipoDoc = Trim(FrmPagoCredito.HfgFacturas.Text)
        FrmPagoCredito.HfgFacturas.col = 4
        serie = Mid(FrmPagoCredito.HfgFacturas.Text, 1, 4)
        Numero = Mid(FrmPagoCredito.HfgFacturas.Text, 6, 10)
        FrmPagoCredito.HfgFacturas.col = 0
        Persona = Trim(FrmPagoCredito.HfgFacturas.Text)
        strCadena = "SELECT sSerie,cDocumentoVenta,dEmisionVenta,dVencimiento,cPersona,Persona,Saldo FROM DocumentoVenta WHERE (sSerie='" & serie & "' AND cDocumentoVenta='" & Numero & "' AND cPersona='" & Persona & "' AND doc_cod='" & TipoDoc & "')"
        Call ConfiguraRst(strCadena)
        
        
        
        
        Me.DtpFechaPago.Value = CVDate(rst(3))
        Me.TxtCodEntidad.Text = rst(4)
        Me.TxtEntidad.Text = rst(5)
        Me.TxtDeuda.Text = Format(rst(6), "#,##0.00")
        Me.TxtDeuda.Locked = True
        Me.TxtSerieCredito.Text = "0003"
        Me.TxtNumeroCredito.Text = nuevo
        
        Set rst = Nothing
End If
If FrmDetalleDeudores.Procedencia = nuevo Then
        Me.TxtSerieCredito.Text = "0003"
        Me.TxtNumeroCredito.Text = Nuevo_pago
        
        Me.TxtCodEntidad.Text = FrmDetalleDeudores.HfgDeudas.TextMatrix(FrmDetalleDeudores.HfgDeudas.Row, 0)
        Me.TxtEntidad.Text = FrmDetalleDeudores.HfgDeudas.TextMatrix(FrmDetalleDeudores.HfgDeudas.Row, 1)
        Call MostrarDocumentos
        
End If
End Sub
Sub MostrarDocumentos()
If FrmDetalleDeudores.HfgDeudas.Rows > 0 Then

strCadena = "SELECT SUM(DocumentoVenta.Saldo) FROM DocumentoVenta WHERE (DocumentoVenta.idFormaPago ='" & KEY_CREDITO & "') AND (DocumentoVenta.Anulado = 'F')  AND (DocumentoVenta.Saldo > 0) AND cPersona='" & Trim(Me.TxtCodEntidad.Text) & "'"
Call ConfiguraRst(strCadena)
    If IsNull(rst(0)) = True Then
        Me.TxtDeuda.Text = 0
    Else
        Me.TxtDeuda.Text = Format(rst(0), "###0.00")
    End If
Set rst = Nothing


strCadena = "SELECT DocumentoVenta.dEmisionVenta as EMISION,  (Comprobantes.doc_abrev + ':' + DocumentoVenta.sSerie +'-'+ " & _
"DocumentoVenta.cDocumentoVenta) AS COMPROBANTE , DocumentoVenta.nTotalVenta AS TOTAL, DocumentoVenta.Saldo AS SALDO FROM DocumentoVenta INNER JOIN " & _
"Comprobantes ON DocumentoVenta.doc_cod = Comprobantes.doc_cod  WHERE " & _
            "DocumentoVenta.idFormaPago='" & Trim(KEY_CREDITO) & "' AND DocumentoVenta.anulado='F' AND saldo>0 AND DocumentoVenta.cPersona='" & Trim(Me.TxtCodEntidad.Text) & "'ORDER BY saldo ASC"
            Call LlenarGrillaSalida(strCadena)
End If

End Sub


Private Sub LlenarGrillaSalida(ByVal Cadena As String)
    Call ConfiguraRst(strCadena)
         If rst.RecordCount > 0 Then
        
        
        Me.HfgSalidas.Clear
        Me.HfgSalidas.Rows = 1
        Set Me.HfgSalidas.Recordset = rst
        Me.HfgSalidas.Rows = rst.RecordCount
        HfgSalidas.ColWidth(0) = 1500
        HfgSalidas.ColWidth(1) = 3500
        HfgSalidas.ColWidth(2) = 2000
        HfgSalidas.ColWidth(3) = 2000
        HfgSalidas.ColWidth(4) = 0
        HfgSalidas.ColWidth(5) = 0
        HfgSalidas.ColWidth(6) = 0
        HfgSalidas.ColWidth(7) = 0
        HfgSalidas.ColWidth(8) = 0
        Call DarFormatoFecha(HfgSalidas, 0)
        
    End If
Set rst = Nothing
End Sub
Private Sub TlbAcciones_ButtonClick(ByVal Button As MSComctlLib.Button)
  Select Case Button.key
    Case KEY_SAVE
      
      Call Save
      Call Imprimir_Tiket
        FrmDetalleDeudores.cargar_deudores
        FrmDetalleDeudores.MostrarDocumentos
       
    Case KEY_PRINT
      '  StrCadena = "SELECT     DetallePagoCreditos.FechaPago, Comprobantes.doc_abrev, DetallePagoCreditos.Serie, DetallePagoCreditos.Numero, " & _
        "                      Persona.NombrePersona, Persona.Per_Ruc, Persona.sDireccionCliente1, Comprobantes_1.doc_abrev AS Expr1, " & _
        "                      DocumentoVenta.sSerie, DocumentoVenta.cDocumentoVenta, DocumentoVenta.nTotalVenta, " & _
        "                      DetallePagoCreditos.Monto,DocumentoVenta.saldo,DetallePagoCreditos.letras, Detalle_DocumentoVenta.cProducto, Producto.DescripcionProducto, Unidad.sAbreviatura, " & _
        "                      Detalle_DocumentoVenta.Precio, Detalle_DocumentoVenta.cantidad, Detalle_DocumentoVenta.Total FROM         DocumentoVenta INNER JOIN " & _
        "                      Persona ON DocumentoVenta.cPersona = Persona.cPersona INNER JOIN " & _
        "                      DetallePagoCreditos ON DocumentoVenta.Alm_cod = DetallePagoCreditos.alm_cod INNER JOIN " & _
        "                      Comprobantes ON DetallePagoCreditos.doc_cod = Comprobantes.doc_cod INNER JOIN " & _
        "                      Comprobantes Comprobantes_1 ON DocumentoVenta.doc_cod = Comprobantes_1.doc_cod INNER JOIN " & _
        "                      Detalle_DocumentoVenta ON DocumentoVenta.id_documentoventa = Detalle_DocumentoVenta.id_documentoventa AND " & _
        "                      DocumentoVenta.cDocumentoVenta = Detalle_DocumentoVenta.cDocumentoVenta AND " & _
        "                      DocumentoVenta.doc_cod = Detalle_DocumentoVenta.doc_cod AND " & _
        "                      DocumentoVenta.Alm_cod = Detalle_DocumentoVenta.Alm_Cod AND " & _
        "                      DocumentoVenta.sSerie = Detalle_DocumentoVenta.sSerie INNER JOIN " & _
        "                      Producto ON Detalle_DocumentoVenta.cProducto = Producto.cProducto INNER JOIN " & _
        "                      Unidad ON Producto.cUnidad = Unidad.cUnidad " & _
        "WHERE     (DetallePagoCreditos.doc_cod = '" & Trim(Me.DtcTipoDoc.BoundText) & "') AND (DetallePagoCreditos.Serie = '" & Trim(Me.TxtSerieCredito.Text) & "') AND (DetallePagoCreditos.Numero = '" & Trim(Me.TxtNumeroCredito.Text) & "') AND " & _
        "                      (DocumentoVenta.cDocumentoVenta = '" & Trim(Me.TxtNumero.Text) & "') AND (DocumentoVenta.doc_cod = '" & Trim(Me.DtcTipoDocDeuda.BoundText) & "' AND DocumentoVenta.sSerie='" & Trim(Me.TxtSerie.Text) & "')"
           Call ConfiguraRst(strCadena)
           x = rst.RecordCount
        Ans = ShowMultiReport(rst, "RptPagoCredito1", , App.Path + "\Reportes\")
        
    Case KEY_DELETE
    If MsgBox("Esta Seguro de Eliminar este Pago", vbQuestion + vbYesNo, "Mensaje para el usuario") = vbYes Then
    
    
        
        Me.TlbAcciones.Buttons(KEY_DELETE).Enabled = False
    End If
    Case KEY_EXIT
        Unload Me
  End Select
  
End Sub
Private Sub Imprimir_Tiket()
 Call CargaDefConfigEpsonTM
    Printer.ScaleMode = vbCharacters  'establezco caracteres para controlar la impresion
    Printer.TrackDefault = True 'siempre apunta a la impresora predeter
    Printer.Font.name = "FontB11"
    Printer.Font.Size = "10"
    m_iTamLineaImpresion = Fix(Printer.ScaleWidth / Printer.TextWidth(" "))
    
   
    
    Printer.CurrentX = 0
    Printer.CurrentY = 0
    
    Printer.Print Tab(5); "PEPE'S AUTOSERVICIOS S.A.C"
    Printer.Print Tab(5); "JR. RAMON CASTILLA Nº 155"
    Printer.Print Tab(5); "DELIVERY:521559  RPM:#913647"
    Printer.Print Tab(5); "Tarapoto - San Martin - San Martin"
    Printer.Print Tab(5); "RUC:20493899229"
    Printer.Print Tab(2); "-----------------------------------"
    
    Printer.Print Tab(0); "PAGO CREDITO:"; Space(1); Mid(Me.TxtSerieCredito.Text + Space(50), 1, 4) & "-" & Me.TxtNumeroCredito.Text & Space(3) & Trim(KEY_FECHA)
    Printer.Print Tab(1); "CLIENTE  :"; Mid(Me.TxtEntidad.Text + Space(80), 1, 30)
    Printer.CurrentY = Printer.CurrentY + 0.1
    Printer.Print Tab(1); "===================================="
    Printer.Print Tab(1); "COMPROBANTES CREDITO  :"
    Printer.Print Tab(1); "===================================="
    Printer.Print Tab(1); "DOCUMENTO" + Space(13) + "MONTO" + Space(8) + "SALDO"
    Printer.CurrentY = Printer.CurrentY + 0.1
    Printer.Print Tab(1); "==================================="
    strCadena = "SELECT (Comprobantes.doc_abrev + ':' + DocumentoVenta.sSerie +'-'+ " & _
"DocumentoVenta.cDocumentoVenta) AS COMPROBANTE , DocumentoVenta.nTotalVenta AS TOTAL, DocumentoVenta.Saldo AS SALDO FROM DocumentoVenta INNER JOIN " & _
"Comprobantes ON DocumentoVenta.doc_cod = Comprobantes.doc_cod  WHERE " & _
    "DocumentoVenta.idFormaPago='" & Trim(KEY_CREDITO) & "' AND DocumentoVenta.anulado='F' AND saldo>0 AND DocumentoVenta.cPersona='" & Trim(Me.TxtCodEntidad.Text) & "'ORDER BY 1 ASC"
    Call ConfiguraRst(strCadena)
    For i = 0 To rst.RecordCount - 1
     Printer.Print Tab(1); rst(0) + Space(4) + Format(rst(1), "#,##.00") + Space(4) + Format(rst("saldo"), "#,##0.00")
     rst.MoveNext
    Next i
    Printer.Print Tab(1); "==================================="
    Printer.Print Tab(1); Mid("MONTO ANTERIOR  :" + Space(20), 1, 24) + "S/." + Space(1) + Format(Me.TxtDeuda.Text, "#,##0.00")
    Printer.Print Tab(1); Mid("MONTO PAGADO    :" + Space(20), 1, 24) + "S/." + Space(1) + Format(Me.TxtMonto.Text, "#,##0.00")
    Printer.Print Tab(1); Mid("MONTO RESTANTE  :" + Space(20), 1, 24) + "S/." + Space(1) + Format(Me.txtsaldo.Text, "#,##0.00")
    Printer.Print Tab(1); "==================================="
    Printer.Print Tab(0); "LO ATENDIO:" + Space(1) + KEY_VENDEDOR + Space(1) + "A LAS:" + str(Time)
    Call AbreGaveta
    Printer.EndDoc
    Exit Sub

End Sub



Private Sub TxtDeuda_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        Me.TxtDeuda.Text = Format(Me.TxtDeuda.Text, "#,##0.00")
        Me.TxtMonto.SetFocus
End If
End Sub

Private Sub TxtMonto_Change()
Dim Deuda As Double
Dim Monto As Double
Dim Saldo As Double
Deuda = Me.TxtDeuda.Text
Monto = Val(Me.TxtMonto.Text)
Saldo = Deuda - Monto
Me.txtsaldo.Text = Format(Saldo, "#,##0.00")
End Sub

Private Sub TxtMonto_KeyPress(KeyAscii As Integer)
 
If KeyAscii = 13 Then
 
 Me.TxtMonto.Text = Format(Me.TxtMonto.Text, "###,##0.00")
 
End If
End Sub
Private Sub Resalta(ByVal texto As TextBox)
texto.SelStart = 0
texto.SelLength = Len(Trim(texto))
texto.Text = texto.SelText
texto.SetFocus
End Sub


Private Sub TxtNumeroCredito_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    
    Me.TxtNumeroCredito.Text = FormatosCeros(Me.TxtNumeroCredito.Text, 10)
    strCadena = "SELECT cDocumentoVenta,sSerie FROM DocumentoVenta WHERE (cDocumentoVenta='" & Trim(Me.TxtNumeroCredito.Text) & "' AND sSerie='" & Trim(Me.TxtSerieCredito.Text) & "' AND doc_cod='" & Trim(Me.DtcTipoDoc.BoundText) & "' AND Alm_cod='" & Trim(Me.DtcAlmacen.BoundText) & "')"
    Call ConfiguraRst(strCadena)
    
    
End If
Set rst = Nothing


End Sub



