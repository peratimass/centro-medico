VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form FrmRegistroCompras 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7935
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15015
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7935
   ScaleWidth      =   15015
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmImportacion 
      BackColor       =   &H00FFFFFF&
      Caption         =   "IMPORTACION DE VENTAS"
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
      Height          =   7575
      Left            =   240
      TabIndex        =   11
      Top             =   120
      Visible         =   0   'False
      Width           =   13455
      Begin MSDataListLib.DataCombo DtcPeriodo 
         Height          =   330
         Left            =   960
         TabIndex        =   17
         Top             =   360
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         ForeColor       =   8388608
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
      Begin VitekeySoft.ChameleonBtn cmdImportarDesdeKeyfacil 
         Height          =   495
         Left            =   120
         TabIndex        =   12
         Top             =   6720
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   873
         BTYPE           =   5
         TX              =   "        IMPORTAR COMPRAS       "
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
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
         MICON           =   "FrmRegistroCompras.frx":0000
         PICN            =   "FrmRegistroCompras.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSComctlLib.ProgressBar prog_indicador 
         Height          =   225
         Left            =   120
         TabIndex        =   13
         Top             =   6435
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   397
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VitekeySoft.ChameleonBtn cmdCargarArchivos 
         Height          =   300
         Left            =   960
         TabIndex        =   14
         Top             =   840
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   529
         BTYPE           =   5
         TX              =   "CARGAR ARCHIVO"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   8421631
         BCOLO           =   8421631
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmRegistroCompras.frx":3D5D
         PICN            =   "FrmRegistroCompras.frx":3D79
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfImportacionkeyfacil 
         Height          =   5175
         Left            =   120
         TabIndex        =   15
         Top             =   1200
         Width           =   13215
         _ExtentX        =   23310
         _ExtentY        =   9128
         _Version        =   393216
         ForeColor       =   8388608
         FixedCols       =   0
         ForeColorFixed  =   8388608
         BackColorBkg    =   16777215
         GridColor       =   8388608
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
      Begin VB.Image Image1 
         Height          =   240
         Left            =   13080
         Picture         =   "FrmRegistroCompras.frx":6B0F
         Top             =   240
         Width           =   240
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   240
         TabIndex        =   16
         Top             =   360
         Width           =   690
      End
      Begin VB.Image cmdclose 
         Height          =   240
         Left            =   16440
         Picture         =   "FrmRegistroCompras.frx":99B3
         Top             =   240
         Width           =   240
      End
   End
   Begin VB.TextBox txtEmpresa 
      Height          =   285
      Left            =   480
      TabIndex        =   3
      Text            =   "TxtEmpresa"
      Top             =   6360
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.TextBox txtRuc 
      Height          =   285
      Left            =   3480
      TabIndex        =   2
      Text            =   "txtRuc"
      Top             =   6360
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.TextBox txtAnio 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Buscar"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2760
      TabIndex        =   0
      Top             =   840
      Width           =   1335
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   4680
      TabIndex        =   4
      Top             =   840
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
   End
   Begin MSComctlLib.ImageList ImgIconos 
      Left            =   12840
      Top             =   6600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRegistroCompras.frx":C857
            Key             =   "(Nuevo)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRegistroCompras.frx":CCAB
            Key             =   "(Importar)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRegistroCompras.frx":D37D
            Key             =   "(Modificar)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRegistroCompras.frx":D69D
            Key             =   "(Exportar)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRegistroCompras.frx":DD6F
            Key             =   "(Eliminar)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRegistroCompras.frx":E1C3
            Key             =   "(Salir)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRegistroCompras.frx":E617
            Key             =   "(RCompras)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRegistroCompras.frx":FD89
            Key             =   "(RVentas)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRegistroCompras.frx":101DB
            Key             =   "(ple)"
         EndProperty
      EndProperty
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfdPersona 
      Height          =   6375
      Left            =   240
      TabIndex        =   5
      Top             =   1320
      Width           =   13455
      _ExtentX        =   23733
      _ExtentY        =   11245
      _Version        =   393216
      ForeColor       =   8388608
      FixedCols       =   0
      ForeColorFixed  =   8388608
      BackColorBkg    =   16777215
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
   Begin ComCtl3.CoolBar ClbAcciones 
      Height          =   6345
      Left            =   13800
      TabIndex        =   6
      Top             =   1320
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   11192
      BandCount       =   1
      ForeColor       =   -2147483635
      ImageList       =   "ImgIconos"
      FixedOrder      =   -1  'True
      VariantHeight   =   0   'False
      Orientation     =   1
      EmbossPicture   =   -1  'True
      _CBWidth        =   900
      _CBHeight       =   6345
      _Version        =   "6.0.8169"
      Child1          =   "TlbAcciones"
      MinHeight1      =   840
      Width1          =   3180
      FixedBackground1=   0   'False
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar TlbAcciones 
         Height          =   810
         Left            =   30
         TabIndex        =   7
         Top             =   30
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   1429
         ButtonWidth     =   1588
         ButtonHeight    =   1429
         AllowCustomize  =   0   'False
         Style           =   1
         ImageList       =   "ImgIconos"
         HotImageList    =   "ImgIconos"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   13
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "   Nuevo   "
               Key             =   "(Nuevo)"
               Object.ToolTipText     =   "Nuevo"
               ImageKey        =   "(Nuevo)"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Ingresar"
               Key             =   "(Modificar)"
               Object.ToolTipText     =   "Modificar"
               ImageKey        =   "(Modificar)"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Importar"
               Key             =   "(Importar)"
               ImageKey        =   "(Importar)"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Exportar"
               Key             =   "(Exportar)"
               ImageKey        =   "(Exportar)"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "PLE 3.0"
               Key             =   "(Ple)"
               ImageKey        =   "(ple)"
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Eliminar"
               Key             =   "(Eliminar)"
               Object.ToolTipText     =   "Eliminar"
               ImageKey        =   "(Eliminar)"
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Salir"
               Key             =   "(Salir)"
               ImageKey        =   "(Salir)"
            EndProperty
         EndProperty
         OLEDropMode     =   1
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "REGISTRO DE COMPRAS MENSUAL :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   450
      TabIndex        =   10
      Top             =   240
      Width           =   2925
   End
   Begin VB.Label lblEmpresa 
      BackStyle       =   0  'Transparent
      Caption         =   "Registro de Ventas Mensual:"
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
      Left            =   3720
      TabIndex        =   9
      Top             =   240
      Width           =   7035
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "AÑO:"
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
      Left            =   570
      TabIndex        =   8
      Top             =   840
      Width           =   495
   End
   Begin VB.Shape ShpDatos 
      BackColor       =   &H00DFDFE0&
      BorderStyle     =   3  'Dot
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   435
      Left            =   240
      Top             =   180
      Width           =   13455
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00DFDFE0&
      BorderStyle     =   3  'Dot
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   555
      Left            =   240
      Top             =   720
      Width           =   4215
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   7935
      Left            =   0
      Top             =   0
      Width           =   15015
   End
End
Attribute VB_Name = "FrmRegistroCompras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Procedencia As EnumProcede

Private Sub ChameleonBtn3_Click()

End Sub

Private Sub cmdCargarArchivos_Click()
Call load_compra(Me.DtcPeriodo.BoundText)
      
End Sub
Private Sub load_compra(ByVal in_periodo As String)
Dim Archivo As String
Dim in_mes As String
Dim in_ejercicio As String
strCadena = "SELECT * FROM con_periodo WHERE Id='" & in_periodo & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    in_mes = Format(rst("Mes"), "00")
    in_ejercicio = rst("Ejercicio")
End If


Archivo = Trim("RegistroCompra" & in_mes & in_ejercicio & KEY_RUC) & ".xls"
      'Dim obj As New get_excel
      Set Me.HfImportacionkeyfacil.DataSource = Leer_Excel(App.Path & "\comparar_percy\" & Archivo, "Reporte")
     
      
      
      

End Sub

Private Sub cmdImportarDesdeKeyfacil_Click()
 
For i = 0 To Me.HfImportacionkeyfacil.Rows - 1
    in_doc = Me.HfImportacionkeyfacil.TextMatrix(Me.HfImportacionkeyfacil.Row, 3)
    in_serie = Me.HfImportacionkeyfacil.TextMatrix(Me.HfImportacionkeyfacil.Row, 4)
    in_numero = Me.HfImportacionkeyfacil.TextMatrix(Me.HfImportacionkeyfacil.Row, 5)
    in_proveedor = Me.HfImportacionkeyfacil.TextMatrix(Me.HfImportacionkeyfacil.Row, 6)
    in_emision = Me.HfImportacionkeyfacil.TextMatrix(Me.HfImportacionkeyfacil.Row, 7)
    in_vencimiento = Me.HfImportacionkeyfacil.TextMatrix(Me.HfImportacionkeyfacil.Row, 8)
    
    in_subtotal = Val(Me.HfImportacionkeyfacil.TextMatrix(Me.HfImportacionkeyfacil.Row, 7))
    in_igv = Val(Me.HfImportacionkeyfacil.TextMatrix(Me.HfImportacionkeyfacil.Row, 8))
    in_total = Val(Me.HfImportacionkeyfacil.TextMatrix(Me.HfImportacionkeyfacil.Row, 9))
    
    strCadena = "SELECT * FROM movimiento_compra WHERE id_doc='" & in_doc & "' and serie='" & in_serie & "' and numero='" & in_numero & "' and id_proveedor='" & in_proveedor & "' and ruc='" & KEY_RUC & "' LIMIT 1"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount < 1 Then
       
       
                in_cta_compra = KEY_CTA_COMPRA_SOLES
           
        If in_doc = "0002" Then
             in_cta_compra = KEY_CTA_COMPRA_RH
        End If
        If in_doc = "0417" Then
            in_cta_compra = KEY_CTA_LETRA_PAGAR_SOLES
        End If
        If in_doc = "0418" Then
            in_cta_compra = KEY_CTA_FET_SOLES
        End If
        
        
        If in_doc = "0419" Then
           in_cta_compra = KEY_CTA_ANT_SOLES
        End If
        
        in_responsable = "0"
        
        
        
        If KEY_PAIS = KEY_PERU Then
         '   strCadena = "call P_insert_compra_ultimate('" & in_doc & "','" & KEY_ALM & "','" & Format(CVDate(in_emision), "YYYY-mm-dd") & "','" & Format(CVDate(in_vencimiento), "YYYY-mm-dd") & "','02'," & _
            "'02','--','00001','" & formato_item(Month(in_emision), 2) & "','" & Year(in_emision) & "','" & in_serie & "'," & _
            "'" & Format(Trim(in_numero), "00000000") & "','" & cod_identidad & "','" & Trim(in_idproveedor) & "','" & UCase(in_proveedor) & "','" & Trim(KEY_CAMBIO) & "'," & _
            "'0','" & Val(Me.LblValorVenta.Text) & "','" & Val(Me.LblIgv.Text) & "','" & Val(Me.lblISC.Text) & "','0','" & Val(Me.TxtPecepcion.Text) & "','0','" & Val(Me.lblExonerado.Text) & "','0','" & Val(Me.lblTotal.Text) & "','" & in_saldo & "'," & _
            " '" & KEY_USUARIO & "','" & Trim(Me.txtObservacion.Text) & "','" & Me.DtcTipo.BoundText & "','" & Me.DtcPeriodo.BoundText & "','" & in_cta_compra & "','" & in_responsable & "','" & Val(Me.txtFob.Text) & "','" & Val(Me.txtSeguro.Text) & "','" & Val(Me.TxtFlete.Text) & "','" & Val(Me.TxtCif.Text) & "','" & KEY_RUC & "')"
        End If
        
        Call ConfiguraRstP(strCadena)
        id_compra = rstP(0)
        
If KEY_CONTABILIDAD = "si" Then
    
    
    
        
        If KEY_PAIS = KEY_PERU Then
            strCadena = "call p_insert_compra_emitido_premiun('" & id_compra & "')"
        Else
            strCadena = "call p_insert_compra_emitido_internacional('" & id_compra & "')"
        End If
            CnBd.Execute (strCadena)
    
        
    End If
       
       
       
       
       
       
       
       
       
       
       
       
       
       
       
       
       
       
       
       
       
       
       
    End If
Next

 
 
 
End Sub

Private Sub Command1_Click()
Call actualizar_anio(Me.txtRuc.Text, Me.txtAnio.Text)
End Sub
Public Sub actualizar_anio(ByVal ruc As String, ByVal Anio As String)
strCadena = "SELECT ruc,mes,descripcion AS Periodo,anio, E.descripcion  as estado,razon FROM  registro_compras WHERE ruc='" & Trim(Me.txtRuc.Text) & "' AND anio LIKE '%" & Trim(Me.txtAnio.Text) & "%' ORDER BY anio,mes"
Call llenarGrid(Me.HfdPersona, Me)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    Unload Me
End If
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 50
Me.LblEmpresa.Caption = KEY_EMPRESA + Space(2) + "***[" + "RUC:" + Space(2) + KEY_RUC + "]***"
Me.TxtEmpresa.Text = KEY_EMPRESA
Me.txtRuc.Text = KEY_RUC


Call actualizar


 
End Sub
Public Sub actualizar()

strCadena = "SELECT * FROM view_registro_compras WHERE ruc='" & KEY_RUC & "' "
Call llenarGrid(Me.HfdPersona, Me)
  
  
End Sub

Private Sub HfdPersona_SelChange()
If Len(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)) = 11 Then
    TlbAcciones.Buttons(KEY_UPDATE).Enabled = True
    TlbAcciones.Buttons(KEY_DELETE).Enabled = True
    TlbAcciones.Buttons(KEY_IMPORTAR).Enabled = True
    TlbAcciones.Buttons(KEY_EXPORTAR).Enabled = True
    TlbAcciones.Buttons("(Ple)").Enabled = True
Else
    TlbAcciones.Buttons(KEY_UPDATE).Enabled = False
    TlbAcciones.Buttons(KEY_DELETE).Enabled = False
    TlbAcciones.Buttons(KEY_IMPORTAR).Enabled = True
    TlbAcciones.Buttons(KEY_EXPORTAR).Enabled = True
    TlbAcciones.Buttons("(Ple)").Enabled = False
End If
End Sub
Private Function get_periodo_compra(ByVal in_mes As Integer, ByVal in_anio As String) As String

Dim in_fecha As String
in_fecha = "01-" & Format(in_mes, "00") & "-" & in_anio
get_periodo_compra = get_periodo_actual(Format(in_fecha, "YYYY-mm-dd"))
End Function

Private Sub Image1_Click()
Me.frmImportacion.Visible = False
End Sub


Private Sub importar()


strCadena = "SELECT Id as Codigo,CONCAT(nombre,' - ',Ejercicio) as Descripcion FROM con_periodo ORDER BY ID DESC"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcPeriodo)
Me.DtcPeriodo.BoundText = get_periodo_actual(KEY_FECHA)
Me.frmImportacion.Visible = True


End Sub






Private Sub TlbAcciones_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim afecto As Double, exonerado As Double, igv As Double, Total                 As Double
        
Select Case Button.key
    Case KEY_NEW
      Procedencia = nuevo
      FrmDetalleRegistroCompras.Show
     
    Case KEY_UPDATE
      Procedencia = modificar
      FrmRegistroComprasList.Show
    Case KEY_DELETE
      If MsgBox(MSGELIMINAR, vbQuestion + vbYesNo, MSGVERIFICACION) = vbYes Then
        If MsgBox("Se Borraran Todos los registros relacionados", vbQuestion + vbYesNo) = vbYes Then
            
            strCadena = "DELETE FROM registro_compras WHERE ruc='" & Trim(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)) & "' AND mes='" & Trim(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 1)) & "' AND anio='" & Trim(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 3)) & "'"
            CnBd.Execute (strCadena)
             
            strCadena = "DELETE FROM movimiento_compra WHERE ruc='" & Trim(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)) & "' AND id_mes='" & Trim(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 1)) & "' AND id_anio='" & Trim(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 3)) & "'"
            CnBd.Execute (strCadena)
             
            Call actualizar
        End If
      End If
    Case KEY_IMPORTAR
        Call importar
        Exit Sub
    Case "(Ple)"
        Procedencia = nuevo
        FrmRegistroSunat.Show
        
            
    
    Case KEY_EXPORTAR
        
        Procedencia = nuevo
        frmTempExportExcel.Show
        frmTempExportExcel.txtidperiodo.Text = get_periodo_compra(Val(Trim(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 1))), Trim(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 3)))


    
    Case "(Salir)"
      Unload Me
  End Select
End Sub
Public Sub Importarcompras(ByVal ruc As String)
Dim rstRemoto As New ADODB.Record
Dim valorVenta As Double
Dim igv As Double
Dim Total As Double

  Set rstT = Nothing
  Set rstT = New ADODB.Recordset
  rstT.CursorLocation = adUseClient
  strCadena = "SELECT * FROM RegistroComprasDetalle WHERE mes='01'OR mes='02' AND anio='2013' WHERE ruc='20104050337' ORDER BY  codigounico ASC"
  rstT.Open strCadena, cnbd1, adOpenKeyset, adLockOptimistic
  
If rstT.RecordCount > 0 Then
    rstT.MoveFirst
    For i = 0 To rstT.RecordCount - 1
        If Len(Trim(rstT("RucCliente"))) = 8 Then
            cod_identidad = 1
        End If
        If Len(Trim(rstT("RucCliente"))) = 11 Then
            cod_identidad = 6
        End If
    
         If rstT("tipo_compra") = "01" Then
            valorVenta = rstT("grav1")
            igv = rstT("igv1")
         End If
         If rstT("tipo_compra") = "02" Then
            valorVenta = rstT("grav2")
            igv = rstT("igv2")
         End If
         If rstT("tipo_compra") = "03" Then
            valorVenta = rstT("grav3")
            igv = rstT("igv2")
         End If
         
        strCadena = "P_insert_compra('" & formato_item(rstT("doc_cod"), 4) & "','00001','" & Format(rstT("fecha"), "YYYY-mm-dd") & "','" & Format(rstT("fecha_cancelacion"), "YYYY-mm-dd") & "','02'," & _
        "'" & rstT("tipo_compra") & "','','" & formato_item(rstT("moneda"), 5) & "','" & rstT("mes") & "','" & rstT("anio") & "','" & formato_item(rstT("serie"), 3) & "'," & _
        "'" & formato_item(rstT("numero"), 8) & "','" & cod_identidad & "','" & Trim(rstT("RucCliente")) & "','" & UCase(rstT("NombreCliente")) & "','" & Val(rstT("tc")) & "'," & _
        "'0','" & valorVenta & "','" & igv & "','" & Val(rstT("isc")) & "','0','" & Val(rstT("percepcion")) & "','0','" & Val(rstT("nograv")) & "','0','" & Val(rstT("total")) & "','" & Val(rstT("total")) & "','" & KEY_USUARIO & "','--','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
         
        
        id_compra = LastRegistroRUC("movimiento_compra", "id_compra")
        strCadena = "UPDATE movimiento_compra SET anulado='si',nproveedor='A N U L A D O',id_proveedor='',valor_venta='0',igv='0',isc='0',ivap='0',percepcion='0',retencion='0',exonerado='0',otros='0',total=0,saldo='0',isc='0',percepcion='0',otros='0' WHERE id_compra='" & id_compra & "' AND ruc='" & KEY_RUC & "'"
        rstT.MoveNext
        DoEvents
    Next i
End If

  
End Sub

Private Sub llenarGrid(ByVal Grilla As MSHFlexGrid, ByVal Formulario As Form)
On Error GoTo salir
Dim Acumulado As Double
 Call ConfiguraRst(strCadena)
  If rst.RecordCount < 1 Then
    Grilla.Rows = 0
    Exit Sub

End If
  
  
    Grilla.Rows = 0
    
    
    
      
      
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
            ' -- almacena el ancho de los campos
           Grilla.ColWidth(0) = 1200
           Grilla.ColWidth(1) = 0
           Grilla.ColWidth(2) = 3000
           Grilla.ColWidth(3) = 700
           Grilla.ColWidth(4) = 2500
           Grilla.ColWidth(5) = 1500
           Grilla.ColWidth(6) = 4000
        Next
      
        cabecera = "RUC" & vbTab & "MES" & vbTab & "PERIODO" & vbTab & "AÑO" & vbTab & "RAZON SOCIAL" & vbTab & "ACUMULADO" & vbTab & "ESTADO"
        Grilla.AddItem cabecera
         For k = 0 To 6
             Grilla.col = k
             Grilla.Row = 0
             Grilla.CellBackColor = &HDFDFE0
        Next k

        rst.MoveFirst
        Acumulado = 0
        For i = 0 To rst.RecordCount - 1
        If IsNull(rst("acumulado")) = True Then
            in_acumulado = 0
        Else
            in_acumulado = rst("acumulado")
        End If
                      
            Fila = rst("ruc") & vbTab & rst("mes") & vbTab & rst("periodo") & vbTab & rst("anio") & vbTab & rst("nombre_completo") & vbTab & Format(in_acumulado, "#,##0.000") & vbTab & rst("estado")
            Grilla.AddItem Fila
               
                    For j = i To Grilla.Rows - 1
                       Grilla.col = 5
                       Grilla.Row = i + 1
                       If (rst("estado") = "PENDIENTE") Then
                            Grilla.CellBackColor = &H8080FF
                        Else
                            Grilla.CellBackColor = &HC0FFC0
                        End If
                    Next j
                    
                    
                   
    
        ' Establecemos las Etiquetas de las Columnas
  
    
    
               
        
            rst.MoveNext
        Next i
  Formulario.TlbAcciones.Buttons(KEY_UPDATE).Enabled = False
  Formulario.TlbAcciones.Buttons(KEY_DELETE).Enabled = False
  Formulario.TlbAcciones.Buttons(KEY_IMPORTAR).Enabled = False
  Formulario.TlbAcciones.Buttons(KEY_EXPORTAR).Enabled = False
  Formulario.TlbAcciones.Buttons("(Ple)").Enabled = False
  
    
  Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"

End Sub

Private Sub txtAnio_Change()
If (Len(Me.txtAnio.Text) = 4) Then
    Me.Command1.Enabled = True
Else
    Me.Command1.Enabled = False
End If
End Sub

Private Sub txtAnio_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call actualizar_anio(Me.txtRuc.Text, Me.txtAnio.Text)
End If
End Sub






