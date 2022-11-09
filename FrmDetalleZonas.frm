VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form FrmDetalleZonas 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5775
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6495
   FillStyle       =   0  'Solid
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   6495
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox TxtDescripcion 
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1560
      MaxLength       =   120
      TabIndex        =   6
      Top             =   1920
      Width           =   2925
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Agregar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4560
      TabIndex        =   5
      Top             =   1920
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Modificar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5400
      TabIndex        =   4
      Top             =   1920
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame FrmActualizar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "MODIFICAR DATOS"
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
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   2520
      Visible         =   0   'False
      Width           =   6135
      Begin VB.CommandButton Command3 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   4320
         TabIndex        =   19
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox TxtDescripcionActualizar 
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   1455
         MaxLength       =   500
         TabIndex        =   2
         Top             =   840
         Width           =   2805
      End
      Begin VB.CommandButton CmdActualizar 
         Caption         =   "Actualizar"
         Height          =   375
         Left            =   4320
         TabIndex        =   1
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label4 
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
         Left            =   360
         TabIndex        =   3
         Top             =   885
         Width           =   1005
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00DFDFE0&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   1275
         Left            =   120
         Top             =   480
         Width           =   5535
      End
   End
   Begin MSComctlLib.ImageList ImgIconos 
      Left            =   5640
      Top             =   2880
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleZonas.frx":0000
            Key             =   "(Aceptar)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleZonas.frx":031C
            Key             =   "(Eliminar)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleZonas.frx":077C
            Key             =   "(Inicio)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleZonas.frx":0BDC
            Key             =   "(Modificar)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleZonas.frx":0EF8
            Key             =   "(Nuevo)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleZonas.frx":1358
            Key             =   "(Quitar)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleZonas.frx":1674
            Key             =   "(Salir)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleZonas.frx":1AD4
            Key             =   "(Red)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleZonas.frx":1F34
            Key             =   "(Grabar)"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleZonas.frx":2814
            Key             =   "(Agregar)"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleZonas.frx":2B30
            Key             =   "(Buscar)"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleZonas.frx":2E4C
            Key             =   "(Cancelar)"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar ClbAcciones 
      Height          =   870
      Left            =   5220
      TabIndex        =   7
      Top             =   4870
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   1535
      BandCount       =   1
      ForeColor       =   -2147483635
      ImageList       =   "ImgIconos"
      FixedOrder      =   -1  'True
      VariantHeight   =   0   'False
      EmbossPicture   =   -1  'True
      _CBWidth        =   1035
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
         TabIndex        =   8
         Top             =   30
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   1429
         ButtonWidth     =   1032
         ButtonHeight    =   1429
         Style           =   1
         ImageList       =   "ImgIconos"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Salir"
               Key             =   "(Cancelar)"
               Object.ToolTipText     =   "Cancelar"
               ImageKey        =   "(Salir)"
            EndProperty
         EndProperty
         OLEDropMode     =   1
      End
   End
   Begin MSDataListLib.DataCombo DtcProvincia 
      Height          =   315
      Left            =   1560
      TabIndex        =   9
      Top             =   480
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      Locked          =   -1  'True
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
   Begin MSDataListLib.DataCombo DtcDistrito 
      Height          =   315
      Left            =   1560
      TabIndex        =   10
      Top             =   840
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      Locked          =   -1  'True
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
   Begin MSDataListLib.DataCombo DtcDepartamento 
      Height          =   315
      Left            =   1560
      TabIndex        =   11
      Top             =   120
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      Locked          =   -1  'True
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfVinculados 
      Height          =   2295
      Left            =   120
      TabIndex        =   12
      Top             =   2520
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   4048
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   8388608
      FixedCols       =   0
      ForeColorFixed  =   8388608
      BackColorBkg    =   16777215
      GridColor       =   12632064
      FocusRect       =   0
      GridLines       =   2
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
   Begin MSDataListLib.DataCombo DtcUrbanizacion 
      Height          =   315
      Left            =   1560
      TabIndex        =   17
      Top             =   1200
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      Locked          =   -1  'True
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
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "URBANIZACION :"
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
      Left            =   315
      TabIndex        =   18
      Top             =   1320
      Width           =   1125
   End
   Begin VB.Label LblDescripcion 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DESCRIPCION ZONA :"
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
      Left            =   15
      TabIndex        =   16
      Top             =   1920
      Width           =   1425
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DEPARTAMENTO :"
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
      Left            =   225
      TabIndex        =   15
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PROVINCIA :"
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
      Left            =   585
      TabIndex        =   14
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DISTRITO :"
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
      Left            =   735
      TabIndex        =   13
      Top             =   960
      Width           =   705
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   5775
      Left            =   0
      Top             =   0
      Width           =   6495
   End
   Begin VB.Shape ShpDatos 
      BackColor       =   &H00DFDFE0&
      FillColor       =   &H00DFDFE0&
      FillStyle       =   0  'Solid
      Height          =   555
      Left            =   120
      Top             =   1800
      Width           =   6135
   End
End
Attribute VB_Name = "FrmDetalleZonas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cCodigo As Double
Private Sub llenarGrid(ByVal Grilla As MSHFlexGrid, ByVal id_Urbanizacion As Double)
'On Error GoTo salir


strCadena = "SELECT id_zona, descripcion_zona FROM zona WHERE id_urbanizacion='" & id_Urbanizacion & "'"

Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 1
    Grilla.Clear
    
    Exit Sub

End If
   n = 1
   Grilla.Clear
   Grilla.Refresh
   Grilla.Rows = 0
      
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
            ' -- almacena el ancho de los campos
           Grilla.ColWidth(0) = 1100
           Grilla.ColWidth(1) = 4500
           
        Next
        cabecera = "CODIGO" & vbTab & "ZONA"
        Grilla.AddItem cabecera
         For k = 0 To 1
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
                            Next k
                            
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
            
            Fila = Fila & rst("id_zona") & vbTab & rst("descripcion_zona")
            If (Fila = "") Then
                X = 1
            End If
          Grilla.AddItem Fila
                    
        Fila = ""
        rst.MoveNext
             
        Next i
End Sub


Private Sub CmdActualizar_Click()
If Me.TxtDescripcionActualizar.Text = "" Then
    MsgBox "Ingrese una Descripcion del Urbanizacion"
    Call Resalta(Me.TxtDescripcionActualizar)
    Exit Sub
Else
    strCadena = "UPDATE Zona SET descripcion_zona='" & Me.TxtDescripcionActualizar.Text & "' where id_zona='" & Val(Me.HfVinculados.TextMatrix(Me.HfVinculados.Row, 0)) & "'"
    CnBd.Execute (strCadena)
     
    Call llenarGrid(Me.HfVinculados, Me.DtcUrbanizacion.BoundText)
    Call FrmZonas.actualizar
    Me.FrmActualizar.Visible = False
End If
End Sub

Private Sub Command1_Click()
Dim idZona As Double
If Me.txtDescripcion.Text = "" Then
    MsgBox "Ingrese una Nombre para la Urbanizacion", vbInformation, "Mensaje para el Usuario"
    Call Resalta(Me.txtDescripcion)
    Exit Sub
Else
    strCadena = "INSERT INTO zona(descripcion_zona,id_Urbanizacion)VALUES('" & Replace(Me.txtDescripcion.Text, "'", "") & "','" & Me.DtcUrbanizacion.BoundText & "')"
   ' CnBd.Execute (strCadena)
     
    idZona = ConsultaUltInsert(strCadena)
    
    Call llenarGrid(Me.HfVinculados, Me.DtcUrbanizacion.BoundText)
    Call FrmZonas.actualizar
    Call Resalta(Me.txtDescripcion)
    Exit Sub
End If
End Sub

Private Sub Command2_Click()

Me.FrmActualizar.Visible = True
cCodigo = Me.HfVinculados.TextMatrix(Me.HfVinculados.Row, 0)
Me.TxtDescripcionActualizar.Text = Me.HfVinculados.TextMatrix(Me.HfVinculados.Row, 1)
Call Resalta(Me.TxtDescripcionActualizar)
 
End Sub

Private Sub Command3_Click()
Me.FrmActualizar.Visible = False
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 700


strCadena = "SELECT id_depa as Codigo,descripcion as Descripcion FROM departamentos ORDER BY descripcion"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcDepartamento)
Me.DtcDepartamento.BoundText = FrmZonas.HfgZona.TextMatrix(FrmZonas.HfgZona.Row, 0)
Set rst = Nothing

strCadena = "SELECT id_provincia as Codigo,descripcion as  Descripcion FROM provincia WHERE id_departamento='" & Me.DtcDepartamento.BoundText & "' ORDER BY descripcion"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcProvincia)
Me.DtcProvincia.BoundText = FrmZonas.HfgZona.TextMatrix(FrmZonas.HfgZona.Row, 1)
Set rst = Nothing

strCadena = "SELECT id_distrito as Codigo,descripcion as  Descripcion FROM distrito WHERE id_provincia='" & Me.DtcProvincia.BoundText & "' ORDER BY descripcion"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcDistrito)
Me.DtcDistrito.BoundText = FrmZonas.HfgZona.TextMatrix(FrmZonas.HfgZona.Row, 2)
Set rst = Nothing


strCadena = "SELECT id_urbanizacion as Codigo,descripcion as  Descripcion FROM urbanizacion WHERE  id_distrito='" & Me.DtcDistrito.BoundText & "' ORDER BY descripcion"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcUrbanizacion)
Me.DtcUrbanizacion.BoundText = FrmZonas.HfgZona.TextMatrix(FrmZonas.HfgZona.Row, 3)
Set rst = Nothing


Me.DtcDepartamento.Enabled = False
Me.DtcDistrito.Enabled = False
Me.DtcProvincia.Enabled = False
Me.DtcUrbanizacion.Enabled = False
Call llenarGrid(Me.HfVinculados, FrmZonas.HfgZona.TextMatrix(FrmZonas.HfgZona.Row, 3))


End Sub

Private Sub HfVinculados_SelChange()
If Val(Me.HfVinculados.TextMatrix(Me.HfVinculados.Row, 0)) > 0 Then
    Me.Command2.Visible = True
Else
    Me.Command2.Visible = False
End If
End Sub

Private Sub TlbAcciones_ButtonClick(ByVal Button As MSComctlLib.Button)
 Select Case Button.key
    Case KEY_NEW
      Procedencia = nuevo
      FrmDetalleUrbanizacion.Show
    Case KEY_UPDATE
      Procedencia = modificar
      FrmDetalleLinea.Show
    Case KEY_DELETE
     ' If MsgBox(MSGELIMINAR, vbQuestion + vbYesNo, MSGVERIFICACION) = vbYes Then
      '  Me.HfgZona.Col = 0
       ' strCadena = "DELETE FROM Linea WHERE clinea='" & Trim(Me.HfgZona.Text) & "'"
        'Call EjecutaRST(strCadena)
        'Set RstEjecuta = Nothing
        'Form_Load
      'End If
    Case KEY_CANCEL
        Unload Me
  End Select
End Sub

Private Sub txtdescripcion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.Command1.SetFocus
End If

End Sub

