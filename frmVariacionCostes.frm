VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmVariacionCostes 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7695
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16830
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   8.25
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
   ScaleHeight     =   7695
   ScaleWidth      =   16830
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtcodigo 
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
      Height          =   315
      Left            =   6240
      TabIndex        =   7
      Top             =   240
      Width           =   1575
   End
   Begin VitekeySoft.ChameleonBtn cmdBuscar 
      Height          =   375
      Left            =   8400
      TabIndex        =   4
      Top             =   240
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      BTYPE           =   5
      TX              =   "BUSCAR"
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
      MICON           =   "frmVariacionCostes.frx":0000
      PICN            =   "frmVariacionCostes.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComCtl2.DTPicker DtpInicio 
      Height          =   345
      Left            =   840
      TabIndex        =   1
      Top             =   240
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   609
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
      Format          =   111738881
      CurrentDate     =   40756
   End
   Begin MSComCtl2.DTPicker DtpFin 
      Height          =   345
      Left            =   3360
      TabIndex        =   3
      Top             =   240
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   609
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
      Format          =   111738881
      CurrentDate     =   40756
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfgFacturas 
      Height          =   6735
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   16455
      _ExtentX        =   29025
      _ExtentY        =   11880
      _Version        =   393216
      ForeColor       =   8388608
      Rows            =   0
      FixedRows       =   0
      FixedCols       =   0
      ForeColorFixed  =   8388608
      BackColorBkg    =   16777215
      GridColor       =   -2147483635
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
   Begin VitekeySoft.ChameleonBtn ChameleonBtn1 
      Height          =   375
      Left            =   16320
      TabIndex        =   6
      Top             =   120
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      BTYPE           =   5
      TX              =   ""
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
      MICON           =   "frmVariacionCostes.frx":2601
      PICN            =   "frmVariacionCostes.frx":261D
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CODIGO :"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   5520
      TabIndex        =   8
      Top             =   300
      Width           =   630
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "HASTA:"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   2640
      TabIndex        =   2
      Top             =   300
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DESDE :"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   300
      Width           =   510
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   6  'Inside Solid
      Height          =   615
      Left            =   120
      Top             =   120
      Width           =   16095
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   7695
      Left            =   0
      Top             =   0
      Width           =   16830
   End
End
Attribute VB_Name = "frmVariacionCostes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Procedencia As EnumProcede

Private Sub Command1_Click()


End Sub

Private Sub ChameleonBtn1_Click()
Unload Me
End Sub

Private Sub cmdbuscar_Click()
strCadena = "SELECT *  FROM view_variacion_precio  WHERE id_producto LIKE '%" & Trim(Me.txtcodigo.Text) & "%' and  fecha>='" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "' and fecha<='" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "' and  ruc='" & KEY_RUC & "'"
Call llenarGrid(Me.HfgFacturas, Me)
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 50
Me.DtpInicio.Value = KEY_FECHA
Me.DtpFin.Value = KEY_FECHA

strCadena = "SELECT *  FROM view_variacion_precio  WHERE fecha<='" & KEY_FECHA & "' and  ruc='" & KEY_RUC & "' LIMIT 20"
Call llenarGrid(Me.HfgFacturas, Me)
End Sub

Private Sub llenarGrid(ByVal Grilla As MSHFlexGrid, ByVal Formulario As Form)
On Error GoTo salir
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 0
    Exit Sub

End If
  
  N = 1
  

   Grilla.Rows = 0
      ' Me.HfdGrilla.Rows = rst.RecordCount - 2
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
            ' -- almacena el ancho de los campos
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 1200
           Grilla.ColWidth(2) = 800
           Grilla.ColWidth(3) = 3000
           Grilla.ColWidth(4) = 3000
           Grilla.ColWidth(5) = 3000
           Grilla.ColWidth(6) = 2500
           Grilla.ColWidth(7) = 1100
           Grilla.ColWidth(8) = 1100
        Next
        cabecera = "CODIGO" & vbTab & "FECHA" & vbTab & "HORA" & vbTab & "OBSERVACION" & vbTab & "OPERADOR" & vbTab & "PRODUCTO" & vbTab & "ALMACEN" & vbTab & "P.ANTERIOR" & vbTab & "P.ACTUAL"
        Grilla.AddItem cabecera
         For k = 0 To 8
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
                            Next k
                            
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
            Fila = rst("id_detalle") & vbTab & Format(rst("fecha"), "dd-mm-YYYY") & vbTab & Format(rst("hora"), "HH:mm:ss AM/PM") & vbTab & rst("observacion") & vbTab & rst("nombre_completo") & vbTab & "[ " & rst("id_producto") & " ]" & Space(2) & rst("nombre_prod") & vbTab & rst("descripcion") & vbTab & Format(rst("precio_venta_anterior"), "#,##0.000") & vbTab & Format(rst("precio_venta_actual"), "#,##0.000")
            Grilla.AddItem Fila
           
            rst.MoveNext
             
        Next i
    Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"


End Sub

Private Sub HfgFacturas_DblClick()
Procedencia = buscar
FrmPrecios.Show
Exit Sub
End Sub

Private Sub TlbAcciones_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.key
    
    Case KEY_EXIT
        Unload Me
    
  End Select

End Sub

Private Sub txtcodigo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Procedencia = Selecionar
    FrmProducto.Show
    Exit Sub
End If
End Sub
