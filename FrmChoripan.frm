VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmChoripan 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "REPORTE CHORIPAN"
   ClientHeight    =   9075
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11040
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9075
   ScaleWidth      =   11040
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtProducto 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   8640
      TabIndex        =   17
      Top             =   1120
      Width           =   1815
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfdDetalle 
      Height          =   2895
      Left            =   240
      TabIndex        =   0
      Top             =   2160
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   5106
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
         Name            =   "MS Sans Serif"
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfdDetalle1 
      Height          =   2415
      Left            =   240
      TabIndex        =   4
      Top             =   5520
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   4260
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
         Name            =   "MS Sans Serif"
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
   Begin TabDlg.SSTab SstKardex 
      Height          =   1575
      Left            =   240
      TabIndex        =   8
      Top             =   120
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   2778
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      ForeColor       =   8388608
      TabCaption(0)   =   "Venta"
      TabPicture(0)   =   "FrmChoripan.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Shape1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "LblCantidad"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label8"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label10"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "DtpHasta"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "DtpDesde"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdBuscar"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "ChkFechas"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtCliente"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      Begin VB.TextBox txtCliente 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   8400
         TabIndex        =   16
         Top             =   600
         Width           =   1815
      End
      Begin VB.CheckBox ChkFechas 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Por Fechas"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   13
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "Buscar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5040
         TabIndex        =   12
         Top             =   720
         Width           =   855
      End
      Begin MSComCtl2.DTPicker DtpDesde 
         Height          =   435
         Left            =   1440
         TabIndex        =   9
         Top             =   675
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarForeColor=   -2147483635
         CalendarTitleForeColor=   -2147483635
         Format          =   17104897
         CurrentDate     =   37091
      End
      Begin MSComCtl2.DTPicker DtpHasta 
         Height          =   420
         Left            =   3360
         TabIndex        =   10
         Top             =   675
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   741
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarForeColor=   -2147483635
         CalendarTitleForeColor=   -2147483635
         Format          =   17104897
         CurrentDate     =   37091
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Producto Relacionado:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   240
         Left            =   6120
         TabIndex        =   19
         Top             =   1080
         Width           =   2190
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cliente Relacionado:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   240
         Left            =   6120
         TabIndex        =   18
         Top             =   720
         Width           =   1965
      End
      Begin VB.Label LblCantidad 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Al:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   3120
         TabIndex        =   11
         Top             =   840
         Width           =   255
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderStyle     =   3  'Dot
         FillColor       =   &H00C0C0C0&
         Height          =   975
         Left            =   120
         Top             =   480
         Width           =   10215
      End
   End
   Begin VB.Label lblTotal 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   9120
      TabIndex        =   23
      Top             =   8760
      Width           =   60
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "TOTAL"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   240
      Left            =   7320
      TabIndex        =   22
      Top             =   8760
      Width           =   600
   End
   Begin VB.Label lblTotalstock 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   9120
      TabIndex        =   21
      Top             =   8400
      Width           =   60
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "TOTAL STOCK:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   240
      Left            =   7320
      TabIndex        =   20
      Top             =   8400
      Width           =   1335
   End
   Begin VB.Label lblstock 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   4200
      TabIndex        =   15
      Top             =   8160
      Width           =   945
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "STOCK EN ALMACEN:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   240
      Left            =   2160
      TabIndex        =   14
      Top             =   8160
      Width           =   1935
   End
   Begin VB.Label lblVendido 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   9120
      TabIndex        =   7
      Top             =   8040
      Width           =   60
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "TOTAL VENDIDO:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   240
      Left            =   7320
      TabIndex        =   6
      Top             =   8040
      Width           =   1575
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "COMPROBANTES PERIODO"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   240
      Left            =   240
      TabIndex        =   5
      Top             =   5280
      Width           =   2430
   End
   Begin VB.Label lbltotalventa 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   240
      Left            =   9360
      TabIndex        =   3
      Top             =   5160
      Width           =   1260
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "TOTAL CONSUMIDO :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   240
      Left            =   7320
      TabIndex        =   2
      Top             =   5160
      Width           =   1920
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "COMPROBANTES PERIODO"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   240
      Left            =   240
      TabIndex        =   1
      Top             =   1920
      Width           =   2430
   End
End
Attribute VB_Name = "FrmChoripan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Procedencia As EnumProcede
Private Sub ChkFechas_Click()
If Me.ChkFechas.Value = 1 Then
    Me.DtpDesde.Enabled = True
    Me.DtpHasta.Enabled = True
Else
    Me.DtpDesde.Enabled = False
    Me.DtpHasta.Enabled = False
End If
End Sub

Private Sub CmdBuscar_Click()

If (Me.TxtCliente.Text) = "" Or (Me.txtProducto.Text) = "" Then
    MsgBox "Ingrese Datos de Busqueda", vbInformation, "Mensaje para el usuario"
    Me.TxtCliente.SetFocus
    Exit Sub
End If

If Me.ChkFechas.Value = 1 Then

strCadena = "SELECT sum(nTotalVenta) FROM DocumentoVenta WHERE dEmisionVenta>='" & CVDate(Me.DtpDesde.Value) & "' AND " & _
        "dEmisionVenta<='" & CVDate(Me.DtpHasta.Value) & "' AND Anulado='F' AND idFormaPago='0004' AND saldo>0 AND cPersona='" & Trim(Me.TxtCliente.Text) & "' "
        Call ConfiguraRst(strCadena)
        If IsNull(rst(0)) = True Then
           Me.lbltotalventa.Caption = "S/." & 0
        Else
            Me.lbltotalventa.Caption = "S/." & rst(0)
        End If
        
        Set rst = Nothing
                
strCadena = "SELECT     DocumentoVenta.dEmisionVenta, Comprobantes.doc_abrev + ':' + DocumentoVenta.sSerie + '-'+ DocumentoVenta.cDocumentoVenta, " & _
"                DocumentoVenta.Persona , DocumentoVenta.nTotalVenta, DocumentoVenta.doc_cod,DocumentoVenta.sSerie,DocumentoVenta.cDocumentoVenta, Seguridad.Usuario " & _
"FROM         DocumentoVenta INNER JOIN Seguridad ON DocumentoVenta.id_usuario = Seguridad.IdUsuario INNER JOIN " & _
" Comprobantes ON DocumentoVenta.doc_cod = Comprobantes.doc_cod WHERE dEmisionVenta>='" & CVDate(Me.DtpDesde.Value) & "' AND " & _
        "dEmisionVenta<='" & CVDate(Me.DtpHasta.Value) & "' AND Anulado='F' AND idFormaPago='0004' AND saldo>0 AND cPersona='" & Trim(Me.TxtCliente.Text) & "' ORDER BY DocumentoVenta.sSerie, DocumentoVenta.cDocumentoVenta  ASC"


Else
strCadena = "SELECT sum(nTotalVenta) FROM DocumentoVenta WHERE Anulado='F' AND idFormaPago='0004' AND saldo>0 AND cPersona='" & Trim(Me.TxtCliente.Text) & "' "
        Call ConfiguraRst(strCadena)
        If IsNull(rst(0)) = True Then
           Me.lbltotalventa.Caption = "S/." & 0
        Else
            Me.lbltotalventa.Caption = "S/." & rst(0)
        End If
        
        Set rst = Nothing
        
    strCadena = "SELECT     DocumentoVenta.dEmisionVenta, Comprobantes.doc_abrev + ':' + DocumentoVenta.sSerie + '-'+ DocumentoVenta.cDocumentoVenta, " & _
"                DocumentoVenta.Persona , DocumentoVenta.nTotalVenta, DocumentoVenta.doc_cod,DocumentoVenta.sSerie,DocumentoVenta.cDocumentoVenta ,Seguridad.Usuario " & _
"FROM         DocumentoVenta INNER JOIN Seguridad ON DocumentoVenta.id_usuario = Seguridad.IdUsuario INNER JOIN " & _
" Comprobantes ON DocumentoVenta.doc_cod = Comprobantes.doc_cod AND idFormaPago='0004' AND saldo>0 AND cPersona='" & Trim(Me.TxtCliente.Text) & "' AND cPersona='" & Trim(Me.TxtCliente.Text) & "' ORDER BY DocumentoVenta.sSerie, DocumentoVenta.cDocumentoVenta"
End If

Call LlenarGrilla(Me.HfdDetalle)






If Me.ChkFechas.Value = 1 Then
strCadena = "SELECT     sum(Detalle_DocumentoVenta.Total) " & _
"FROM         Detalle_DocumentoVenta INNER JOIN " & _
"                      DocumentoVenta ON Detalle_DocumentoVenta.id_documentoventa = DocumentoVenta.id_documentoventa AND " & _
"                      Detalle_DocumentoVenta.cDocumentoVenta = DocumentoVenta.cDocumentoVenta And " & _
"                      Detalle_DocumentoVenta.doc_cod = DocumentoVenta.doc_cod AND " & _
"                      Detalle_DocumentoVenta.Alm_Cod = DocumentoVenta.Alm_cod AND Detalle_DocumentoVenta.sSerie = DocumentoVenta.sSerie WHERE " & _
"dEmisionVenta>='" & CVDate(Me.DtpDesde.Value) & "' AND " & _
        "dEmisionVenta<='" & CVDate(Me.DtpHasta.Value) & "' AND Anulado='F' AND dEmisionVenta>='" & CVDate(Me.DtpDesde.Value) & "'  AND Detalle_DocumentoVenta.cProducto='" & Trim(Me.txtProducto.Text) & "'"
        Call ConfiguraRst(strCadena)
        If IsNull(rst(0)) = True Then
            Me.lblVendido.Caption = 0
        Else
            Me.lblVendido.Caption = rst(0)
        End If
        Me.lblTotal.Caption = Val(Val(lblVendido.Caption) + Val(Me.lblTotalstock.Caption))
        Set rst = Nothing

strCadena = "SELECT         DocumentoVenta.dEmisionVenta, Comprobantes.doc_abrev+':'+Detalle_DocumentoVenta.sSerie+'-'+ " & _
"                      Detalle_DocumentoVenta.cDocumentoVenta , DocumentoVenta.Persona, Detalle_DocumentoVenta.total " & _
"FROM         Detalle_DocumentoVenta INNER JOIN " & _
"                      DocumentoVenta ON Detalle_DocumentoVenta.id_documentoventa = DocumentoVenta.id_documentoventa AND " & _
"                      Detalle_DocumentoVenta.cDocumentoVenta = DocumentoVenta.cDocumentoVenta AND " & _
"                      Detalle_DocumentoVenta.doc_cod = DocumentoVenta.doc_cod AND " & _
"                      Detalle_DocumentoVenta.Alm_Cod = DocumentoVenta.Alm_cod AND " & _
"                      Detalle_DocumentoVenta.sSerie = DocumentoVenta.sSerie INNER JOIN " & _
"                      Comprobantes ON Detalle_DocumentoVenta.doc_cod = Comprobantes.doc_cod WHERE dEmisionVenta>='" & CVDate(Me.DtpDesde.Value) & "' AND " & _
        "dEmisionVenta<='" & CVDate(Me.DtpHasta.Value) & "' AND Anulado='F' AND Detalle_DocumentoVenta.cProducto='" & Trim(Me.txtProducto.Text) & "' ORDER BY DocumentoVenta.sSerie, DocumentoVenta.cDocumentoVenta  ASC"

Else

   strCadena = "SELECT     sum(Detalle_DocumentoVenta.Total) " & _
"FROM         Detalle_DocumentoVenta INNER JOIN " & _
"                      DocumentoVenta ON Detalle_DocumentoVenta.id_documentoventa = DocumentoVenta.id_documentoventa AND " & _
"                      Detalle_DocumentoVenta.cDocumentoVenta = DocumentoVenta.cDocumentoVenta And " & _
"                      Detalle_DocumentoVenta.doc_cod = DocumentoVenta.doc_cod AND " & _
"                      Detalle_DocumentoVenta.Alm_Cod = DocumentoVenta.Alm_cod AND Detalle_DocumentoVenta.sSerie = DocumentoVenta.sSerie WHERE " & _
" Anulado='F' AND Detalle_DocumentoVenta.cProducto='" & Trim(Me.txtProducto.Text) & "'"
        Call ConfiguraRst(strCadena)
        If IsNull(rst(0)) = True Then
            Me.lblVendido.Caption = 0
        Else
            Me.lblVendido.Caption = rst(0)
        End If
        Me.lblTotal.Caption = Val(rst(0) + Val(Me.lblTotalstock.Caption))
        Set rst = Nothing
        
    
    strCadena = "SELECT        DocumentoVenta.dEmisionVenta, Comprobantes.doc_abrev+':'+Detalle_DocumentoVenta.sSerie+'-'+ " & _
"                      Detalle_DocumentoVenta.cDocumentoVenta , DocumentoVenta.Persona, Detalle_DocumentoVenta.total " & _
"FROM         Detalle_DocumentoVenta INNER JOIN " & _
"                      DocumentoVenta ON Detalle_DocumentoVenta.id_documentoventa = DocumentoVenta.id_documentoventa AND " & _
"                      Detalle_DocumentoVenta.cDocumentoVenta = DocumentoVenta.cDocumentoVenta AND " & _
"                      Detalle_DocumentoVenta.doc_cod = DocumentoVenta.doc_cod AND " & _
"                      Detalle_DocumentoVenta.Alm_Cod = DocumentoVenta.Alm_cod AND " & _
"                      Detalle_DocumentoVenta.sSerie = DocumentoVenta.sSerie INNER JOIN " & _
"                      Comprobantes ON Detalle_DocumentoVenta.doc_cod = Comprobantes.doc_cod  AND Detalle_DocumentoVenta.cProducto='" & Trim(Me.txtProducto.Text) & "' ORDER BY DocumentoVenta.sSerie, DocumentoVenta.cDocumentoVenta"
End If

Call LlenarGrilla(Me.HfdDetalle1)

End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 10
Me.ChkFechas.Value = 1
Me.DtpDesde.Value = CVDate(Date)
Me.DtpHasta.Value = CVDate(Date)
End Sub
Private Sub LlenarGrilla(ByVal HfdGrillaCont As MSHFlexGrid)
        Call ConfiguraRst(strCadena)
        
        HfdGrillaCont.Rows = 1
        Set HfdGrillaCont.Recordset = rst
        HfdGrillaCont.Rows = rst.RecordCount + 1
        HfdGrillaCont.ColWidth(0) = 1000
        HfdGrillaCont.ColWidth(1) = 2300
        HfdGrillaCont.ColWidth(2) = 3500
        HfdGrillaCont.ColWidth(3) = 1000
        HfdGrillaCont.ColWidth(4) = 0
        HfdGrillaCont.ColWidth(5) = 0
        HfdGrillaCont.ColWidth(6) = 0
        HfdGrillaCont.ColWidth(7) = 2000
        Call DarFormatoFecha(HfdGrillaCont, 0)
        Call DarFormato(HfdGrillaCont, 3)
        Set HfdGrillaCont = Nothing
        Set rst = Nothing
End Sub


Private Sub txtCliente_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Procedencia = buscar
    FrmPersona.Show
End If
End Sub

Private Sub txtProducto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Procedencia = buscar
    FrmProducto.Show
End If
End Sub
