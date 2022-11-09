VERSION 5.00
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form FrmChequeras 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6615
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11655
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   11655
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList ImgIconos 
      Left            =   3360
      Top             =   2670
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
            Picture         =   "FrmChequeras.frx":0000
            Key             =   "(Nuevo)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmChequeras.frx":0454
            Key             =   "(Modificar)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmChequeras.frx":0774
            Key             =   "(Eliminar)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmChequeras.frx":0BC8
            Key             =   "(Salir)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmChequeras.frx":101C
            Key             =   "(Aceptar)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmChequeras.frx":133C
            Key             =   "(Cancelar)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmChequeras.frx":165C
            Key             =   "(Quitar)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmChequeras.frx":197C
            Key             =   "(Agregar)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmChequeras.frx":1C9C
            Key             =   "(Buscar)"
         EndProperty
      EndProperty
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfgDetalle 
      Height          =   5775
      Left            =   240
      TabIndex        =   0
      Top             =   390
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   10186
      _Version        =   393216
      ForeColor       =   8388608
      FixedCols       =   0
      ForeColorFixed  =   8388608
      SelectionMode   =   1
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin ComCtl3.CoolBar ClbAcciones 
      Height          =   5745
      Left            =   10605
      TabIndex        =   1
      Top             =   360
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   10134
      BandCount       =   1
      ForeColor       =   8388608
      ImageList       =   "ImgIconos"
      FixedOrder      =   -1  'True
      VariantHeight   =   0   'False
      Orientation     =   1
      EmbossPicture   =   -1  'True
      _CBWidth        =   900
      _CBHeight       =   5745
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
         TabIndex        =   2
         Top             =   420
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   1376
         ButtonWidth     =   1482
         ButtonHeight    =   1376
         Style           =   1
         ImageList       =   "ImgIconos"
         DisabledImageList=   "ImgIconos"
         HotImageList    =   "ImgIconos"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   9
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
               Caption         =   "Cheques"
               Key             =   "(Actualizar)"
               ImageKey        =   "(Aceptar)"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Modificar"
               Key             =   "(Modificar)"
               Object.ToolTipText     =   "Modificar"
               ImageKey        =   "(Modificar)"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Eliminar"
               Key             =   "(Eliminar)"
               Object.ToolTipText     =   "Eliminar"
               ImageKey        =   "(Eliminar)"
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
   Begin VB.Label LblEmpresa 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CHEQUERAS"
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
      Left            =   285
      TabIndex        =   3
      Top             =   120
      Width           =   1005
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   6615
      Left            =   0
      Top             =   0
      Width           =   11655
   End
End
Attribute VB_Name = "FrmChequeras"
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

Private Sub Form_Load()
CenterForm Me
Me.Top = 500
Call actualizar
End Sub

Private Sub HfgDetalle_Click()
If Val(Me.HfgDetalle.TextMatrix(Me.HfgDetalle.Row, 0)) > 0 Then
    Me.TlbAcciones.Buttons(KEY_DELETE).Enabled = True
    Me.TlbAcciones.Buttons(KEY_UPDATE).Enabled = True
    Me.TlbAcciones.Buttons(KEY_ACTUALIZAR).Enabled = True
Else
    Me.TlbAcciones.Buttons(KEY_DELETE).Enabled = False
    Me.TlbAcciones.Buttons(KEY_UPDATE).Enabled = False
    Me.TlbAcciones.Buttons(KEY_ACTUALIZAR).Enabled = False
End If
End Sub

Private Sub HfgDetalle_SelChange()
If Val(Me.HfgDetalle.TextMatrix(Me.HfgDetalle.Row, 0)) > 0 Then
    Me.TlbAcciones.Buttons(KEY_DELETE).Enabled = True
    Me.TlbAcciones.Buttons(KEY_UPDATE).Enabled = True
    Me.TlbAcciones.Buttons(KEY_ACTUALIZAR).Enabled = True
Else
    Me.TlbAcciones.Buttons(KEY_DELETE).Enabled = False
    Me.TlbAcciones.Buttons(KEY_UPDATE).Enabled = False
    Me.TlbAcciones.Buttons(KEY_ACTUALIZAR).Enabled = False
End If
End Sub

Private Sub TlbAcciones_ButtonClick(ByVal Button As MSComctlLib.Button)
  Select Case Button.key
    Case KEY_NEW
      Procedencia = nuevo
      FrmChequeraDet.Show
    Case KEY_ACTUALIZAR
        Call FrmCheques.actualizar(FrmCheques.Hfcheques, Me.HfgDetalle.TextMatrix(Me.HfgDetalle.Row, 0))
        FrmCheques.Show
    Case KEY_UPDATE
      Procedencia = Modificar
       FrmMisCuentasDet.Show
    Case KEY_DELETE
      If MsgBox(MSGELIMINAR + Chr(13) + "Se Eliminaran los cheques Relacionados a esta Chequera", vbQuestion + vbYesNo, MSGVERIFICACION) = vbYes Then
        
        strCadena = "DELETE FROM chequera WHERE id_chequera='" & Val(Me.HfgDetalle.TextMatrix(Me.HfgDetalle.Row, 0)) & "' AND ruc='" & KEY_RUC & "'"
        CnBd.Execute (strCadena)
         
        strCadena = "DELETE FROM cheque WHERE id_chequera='" & Val(Me.HfgDetalle.TextMatrix(Me.HfgDetalle.Row, 0)) & "' AND ruc='" & KEY_RUC & "'"
        CnBd.Execute (strCadena)
         
        Call actualizar
      End If
    Case KEY_EXIT
        Unload Me
  End Select
End Sub
Private Sub llenarGrid(ByVal Grilla As MSHFlexGrid, ByVal Formulario As Form)
On Error GoTo salir
strCadena = "SELECT C.id_chequera,C.id_cuenta,M.descripcion,C.cheque_inicio,C.cheque_fin,M.numero_cuenta,C.cheque_actual,MO.descripcion as moneda FROM chequera C,mis_cuentas M, entidadfinanciera F,moneda MO WHERE M.id_moneda=MO.id_moneda AND  C.id_cuenta=M.id_cuenta AND C.ruc='" & KEY_RUC & "' AND M.id_entidad=F.id"
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 1
    Grilla.Clear
    Exit Sub

End If
  N = 1
   Grilla.Clear
   Grilla.Refresh
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 0
           Grilla.ColWidth(2) = 3500
           Grilla.ColWidth(3) = 2000
           Grilla.ColWidth(4) = 2000
           Grilla.ColWidth(5) = 1200
           Grilla.ColWidth(6) = 1200
           
           
           
           
        Next
         cabecera = "IDCHEQUERA" & vbTab & "IDCUENTA" & vbTab & "ENTIDAD " & vbTab & "CHEQUES" & vbTab & "NUMERO CUENTA" & vbTab & " MONEDA" & vbTab & " Nº CHEQUE"
        Grilla.AddItem cabecera
         For k = 0 To 6
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
                            Next k
                            
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
        
            Fila = rst("id_chequera") & vbTab & rst("id_cuenta") & vbTab & rst("descripcion") & vbTab & "[*** " + rst("cheque_inicio") + "  -  " + rst("cheque_fin") + " ***]" & vbTab & rst("numero_cuenta") & vbTab & rst("moneda") & vbTab & rst("cheque_actual")
            Grilla.AddItem Fila
            Fila = ""
            rst.MoveNext
             
        Next i
    
    
    
    
  Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"


End Sub





