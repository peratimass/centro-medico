VERSION 5.00
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form FrmUnidadesTransporte 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6615
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14250
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   14250
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
            Picture         =   "FrmUnidadesTransporte.frx":0000
            Key             =   "(Nuevo)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmUnidadesTransporte.frx":0454
            Key             =   "(Modificar)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmUnidadesTransporte.frx":0774
            Key             =   "(Eliminar)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmUnidadesTransporte.frx":0BC8
            Key             =   "(Salir)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmUnidadesTransporte.frx":101C
            Key             =   "(Aceptar)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmUnidadesTransporte.frx":133C
            Key             =   "(Cancelar)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmUnidadesTransporte.frx":165C
            Key             =   "(Quitar)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmUnidadesTransporte.frx":197C
            Key             =   "(Agregar)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmUnidadesTransporte.frx":1C9C
            Key             =   "(Buscar)"
         EndProperty
      EndProperty
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfgDetalle 
      Height          =   5655
      Left            =   240
      TabIndex        =   0
      Top             =   390
      Width           =   12735
      _ExtentX        =   22463
      _ExtentY        =   9975
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
   Begin ComCtl3.CoolBar ClbAcciones 
      Height          =   5745
      Left            =   13125
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
         Height          =   810
         Left            =   30
         TabIndex        =   2
         Top             =   375
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   1429
         ButtonWidth     =   1588
         ButtonHeight    =   1429
         Style           =   1
         ImageList       =   "ImgIconos"
         DisabledImageList=   "ImgIconos"
         HotImageList    =   "ImgIconos"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   7
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
               Caption         =   "Modificar"
               Key             =   "(Modificar)"
               Object.ToolTipText     =   "Modificar"
               ImageKey        =   "(Modificar)"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Eliminar"
               Key             =   "(Eliminar)"
               Object.ToolTipText     =   "Eliminar"
               ImageKey        =   "(Eliminar)"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
      Caption         =   "UNIDADES DE TRANSPORTE"
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
      Left            =   375
      TabIndex        =   3
      Top             =   120
      Width           =   2265
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   6615
      Left            =   0
      Top             =   0
      Width           =   14250
   End
End
Attribute VB_Name = "FrmUnidadesTransporte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Procedencia As EnumProcede
Private Sub Form_Load()
CenterForm Me
Me.Top = 500
Call actualizar
End Sub
Public Sub LlenarTransporte(ByVal Grilla As MSHFlexGrid)
On Error GoTo salir
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 1
    Grilla.Clear
    Exit Sub

End If
   Grilla.Clear
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
            Grilla.ColWidth(0) = 800
           Grilla.ColWidth(1) = 2000
           Grilla.ColWidth(2) = 2000
           Grilla.ColWidth(3) = 2000
           Grilla.ColWidth(4) = 1000
           Grilla.ColWidth(5) = 1600
           Grilla.ColWidth(6) = 1700
           Grilla.ColWidth(7) = 1300
         Next
         
        
        cabecera = "CODIGO" & vbTab & "DESCRIPCION" & vbTab & "MARCA" & vbTab & "PLACA" & vbTab & "LLANTAS" & vbTab & "CONSUMO X HORA" & vbTab & "CAP.TANQUE" & vbTab & "MTC"
        Grilla.AddItem cabecera
         For k = 0 To 7
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        rst.MoveFirst
        
        For i = 0 To rst.RecordCount - 1
            Fila = rst("id_transporte") & vbTab & rst("descripcion") & vbTab & rst("marca") & vbTab & rst("placa") & vbTab & rst("llantas") & vbTab & Format(rst("consumo_hora"), "#,##0.00") & Space(2) & " GL" & vbTab & Format(rst("capacidad_tanque"), "#,##0.00") & Space(2) & "GL" & vbTab & rst("mtc")
            Grilla.AddItem Fila
            Fila = ""
            rst.MoveNext
        Next i
     
      
  
  Exit Sub
salir: MsgBox "Ocurrio un Error en la Conexion", vbInformation, "Mensaje para el usuario"
Set rst = Nothing

End Sub

Private Sub HfgLinea_Click()
If HfgLinea.Row > 0 Then
    TlbAcciones.Buttons(KEY_UPDATE).Enabled = True
    TlbAcciones.Buttons(KEY_DELETE).Enabled = True
  End If
End Sub
Private Sub TlbAcciones_ButtonClick(ByVal Button As MSComctlLib.Button)
  Select Case Button.key
    Case KEY_NEW
      Procedencia = nuevo
      FrmUnidadesTransporteDet.Show
    Case KEY_UPDATE
      Procedencia = Modificar
      FrmUnidadesTransporteDet.Show
    Case KEY_DELETE
      If MsgBox(MSGELIMINAR, vbQuestion + vbYesNo, MSGVERIFICACION) = vbYes Then
        strCadena = "DELETE FROM MiTransporte WHERE id_Transporte='" & Me.HfgDetalle.TextMatrix(Me.HfgDetalle.Row, 0) & "'"
        CnBd.Execute (strCadena)
         
        Call actualizar
      End If
    Case KEY_EXIT
        Unload Me
  End Select
End Sub
Public Sub actualizar()
strCadena = "SELECT     T.id_transporte, I.descripcion, T.marca,T.placa,T.llantas,T.mtc ,T.capacidad_tanque,T.consumo_hora  FROM transporte T,transporte_tipo I WHERE T.id_tipo_transporte=I.id_tipo_transporte AND T.ruc='" & KEY_RUC & "' AND I.ruc='" & KEY_RUC & "'"
Call LlenarTransporte(Me.HfgDetalle)
End Sub






