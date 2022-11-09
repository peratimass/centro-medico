VERSION 5.00
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FrmCheques 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7080
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13305
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7080
   ScaleWidth      =   13305
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox TxtidChequera 
      Height          =   285
      Left            =   7920
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   240
      Visible         =   0   'False
      Width           =   1935
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Hfcheques 
      Height          =   6015
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   10610
      _Version        =   393216
      ForeColor       =   8388608
      FixedCols       =   0
      ForeColorFixed  =   8388608
      SelectionMode   =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   12480
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCheques.frx":0000
            Key             =   "(Nuevo)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCheques.frx":0454
            Key             =   "(Modificar)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCheques.frx":0774
            Key             =   "(Eliminar)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCheques.frx":0BC8
            Key             =   "(Salir)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCheques.frx":101C
            Key             =   "(Aceptar)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCheques.frx":133C
            Key             =   "(Buscar)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCheques.frx":165C
            Key             =   "(Reporte)"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar ClbAcciones1 
      Height          =   5985
      Left            =   12360
      TabIndex        =   2
      Top             =   840
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   10557
      BandCount       =   1
      ForeColor       =   8388608
      FixedOrder      =   -1  'True
      VariantHeight   =   0   'False
      Orientation     =   1
      EmbossPicture   =   -1  'True
      _CBWidth        =   900
      _CBHeight       =   5985
      _Version        =   "6.0.8169"
      Caption1        =   "Acciones"
      Child1          =   "TlbAcciones"
      MinHeight1      =   840
      Width1          =   3180
      FixedBackground1=   0   'False
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar TlbAcciones 
         Height          =   780
         Left            =   30
         TabIndex        =   3
         Top             =   420
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   1376
         ButtonWidth     =   1535
         ButtonHeight    =   1376
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   9
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Llenar"
               Key             =   "(Pagar)"
               Object.ToolTipText     =   "LLENAR CHEQUE SELECCIONADO"
               ImageKey        =   "(Buscar)"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Transaccio"
               Key             =   "(Trans)"
               Object.ToolTipText     =   "Modificar"
               ImageKey        =   "(Aceptar)"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Modificar"
               Key             =   "(Modificar)"
               ImageKey        =   "(Modificar)"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Reporte"
               Key             =   "(Imprimir)"
               ImageKey        =   "(Reporte)"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Salir"
               Key             =   "(Salir)"
               ImageKey        =   "(Salir)"
            EndProperty
         EndProperty
         OLEDropMode     =   1
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ALMACEN:"
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
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   255
      Width           =   885
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   3  'Dot
      Height          =   615
      Left            =   120
      Top             =   120
      Width           =   12135
   End
End
Attribute VB_Name = "FrmCheques"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Procedencia As EnumProcede
Private Sub Form_Load()
CenterForm Me
Me.Top = 500
'Me.Label1.Caption = "***** CHEQUES -" + Space(2) + FrmChequeras.HfgDetalle.TextMatrix(FrmChequeras.HfgDetalle.Row, 5) + ":" + Space(2) + FrmChequeras.HfgDetalle.TextMatrix(FrmChequeras.HfgDetalle.Row, 2) + Space(2) + ":" + FrmChequeras.HfgDetalle.TextMatrix(FrmChequeras.HfgDetalle.Row, 4)
'Call actualizar(Me.Hfcheques, Val(Me.TxtidChequera.text))
End Sub
Public Sub actualizar(ByVal Grilla As MSHFlexGrid, ByVal id_chequera As Double)
On Error GoTo SALIR
Dim nbeneficiario As String
Dim ccostos As String
Dim encargado As String
Me.TxtidChequera.text = id_chequera
strCadena = "SELECT MO.descripcion,M.descripcion as cuenta,M.numero_cuenta FROM chequera CH,cheque C,mis_cuentas M,moneda MO WHERE CH.id_chequera=C.id_chequera AND CH.id_cuenta=M.id_cuenta AND C.ruc='" & KEY_RUC & "' AND CH.ruc='" & KEY_RUC & "' AND M.ruc='" & KEY_RUC & "' AND CH.id_chequera='" & id_chequera & "' AND MO.id_moneda=M.id_moneda"
Call ConfiguraRst(strCadena)
Me.Label1.Caption = "***** CHEQUES -" + Space(2) + rst("descripcion") + ":" + Space(2) + rst("cuenta") + Space(2) + ":" + rst("numero_cuenta")

strCadena = "SELECT * FROM cheque WHERE id_chequera='" & id_chequera & "' AND ruc='" & KEY_RUC & "' ORDER BY numero"
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 1
    Grilla.Clear
    Exit Sub
End If
   Grilla.Clear
   Grilla.Refresh
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 1100
           Grilla.ColWidth(2) = 1200
           Grilla.ColWidth(3) = 1200
           Grilla.ColWidth(4) = 3500
           Grilla.ColWidth(5) = 2000
           Grilla.ColWidth(6) = 2800
           Grilla.ColWidth(7) = 0
       Next
        cabecera = "IDCHEQUE" & vbTab & "FECHA" & vbTab & "NUMERO" & vbTab & "MONTO" & vbTab & "BENEFICIARIO" & vbTab & " C.COSTOS" & vbTab & "ENCARGADO" & vbTab & "SALDO"
        Grilla.AddItem cabecera
         For k = 0 To 6
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
         Next k
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
            If Len(rst("id_beneficiario")) >= 8 Then
                nbeneficiario = BDBuscarCampo("persona", "nombre_completo", "dni", rst("id_beneficiario"))
            Else
                nbeneficiario = ""
            End If
            If Len(rst("ccostos")) > 1 Then
                ccostos = BDBuscarCampo("plan_contable_det", "plan_des", "pc_codigo", rst("ccostos"))
            Else
                ccostos = ""
            End If
            If Len(rst("dni_save")) > 1 Then
                encargado = BDBuscarCampo("persona", "nombre_completo", "dni", rst("dni_save"))
            Else
                encargado = ""
            End If
            Fila = rst("id_cheque") & vbTab & rst("fecha_hora") & vbTab & rst("numero") & vbTab & Format(rst("monto"), "#,##0.00") & vbTab & nbeneficiario & vbTab & ccostos & vbTab & UCase(encargado) & vbTab & rst("saldo")
            Grilla.AddItem Fila
            If rst("seleccionado") = "si" Then
                For k = 0 To 6
                    Grilla.col = k
                    Grilla.Row = i + 1
                    Grilla.CellBackColor = &HC0C0FF
                Next k
            End If
            Fila = ""
            rst.MoveNext
        Next i
Exit Sub
SALIR: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub


Private Sub Hfcheques_SelChange()
If Val(Me.Hfcheques.TextMatrix(Me.Hfcheques.Row, 0)) > 0 And Val(Me.Hfcheques.TextMatrix(Me.Hfcheques.Row, 7)) > 0 Then
    Me.TlbAcciones.Buttons("(Trans)").Enabled = True
    Me.TlbAcciones.Buttons(KEY_UPDATE).Enabled = True
Else
    Me.TlbAcciones.Buttons("(Trans)").Enabled = False
    Me.TlbAcciones.Buttons(KEY_UPDATE).Enabled = False
End If
End Sub

Private Sub TlbAcciones_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo Error
  Select Case Button.Key
    Case KEY_PAGAR
        Procedencia = Nuevo
        'Call FrmChequeNuevo.llenar_cheque(Me.Hfcheques.TextMatrix(Me.Hfcheques.Row, 0))
        FrmChequeNuevo.Show
        'frmNuevoComprobante.Show
    Case "(Trans)"
         FrmTransacciones.Show
    Case KEY_DELETE
       If MsgBox(MSGELIMINAR, vbQuestion + vbYesNo, MSGVERIFICACION) = vbYes Then
        Procedencia = Eliminar
        FrmSeguridad.Show
       End If
    Case KEY_EXIT
        Unload Me
  
  Exit Sub
Error:
MsgBox Err.Number & " " & Err.Description, vbCritical, MSGERROR
End Select
End Sub
