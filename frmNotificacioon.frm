VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmNotificacion 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5085
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12405
   Icon            =   "frmNotificacioon.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   12405
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   300
      Left            =   120
      Top             =   4560
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfCuotas 
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   5953
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
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   3840
      Picture         =   "frmNotificacioon.frx":000C
      Top             =   30
      Width           =   480
   End
   Begin VB.Label lbltelefono 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   345
      Left            =   4305
      TabIndex        =   6
      Top             =   120
      Width           =   7725
   End
   Begin VB.Label lblTotalDeuda 
      Alignment       =   1  'Right Justify
      BackColor       =   &H000000C0&
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Left            =   9720
      TabIndex        =   5
      Top             =   4515
      Width           =   2355
   End
   Begin VB.Label lblEmpresa 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   435
      Left            =   1800
      TabIndex        =   4
      Top             =   540
      Width           =   9090
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sres:"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1080
      TabIndex        =   3
      Top             =   600
      Width           =   615
   End
   Begin VB.Label lblfecha_suspencion 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Left            =   3480
      TabIndex        =   2
      Top             =   4515
      Width           =   2970
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FECHA CORTE :"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   510
      Left            =   600
      TabIndex        =   1
      Top             =   4560
      Width           =   2235
   End
   Begin VB.Image Image2 
      Height          =   240
      Left            =   12120
      Picture         =   "frmNotificacioon.frx":28E6
      Top             =   120
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   585
      Left            =   240
      Picture         =   "frmNotificacioon.frx":578A
      Stretch         =   -1  'True
      Top             =   345
      Width           =   600
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   5085
      Left            =   0
      Top             =   0
      Width           =   12400
   End
End
Attribute VB_Name = "frmNotificacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.LblEmpresa.Caption = KEY_EMPRESA

Me.lbltelefono.Caption = get_telefono_proveedor & Space(2) & "COBRANZA."
Call actualizar
End Sub

Private Sub Image2_Click()

'MDIFrmPrincipal.timer_cobranza.Enabled = True
Unload Me
Exit Sub
End Sub

Private Sub actualizar()
Dim in_mora_dias As Integer

strCadena = "SELECT mora_dias FROM persona_plan_servicio WHERE pago_mensual='si' and dni='" & KEY_RUC & "' and ruc='" & KEY_PROVEEDOR & "' LIMIT 1"
Call ConfiguraRstC(strCadena)
If rstc.RecordCount > 0 Then
    in_mora_dias = rstc("mora_dias")
Else
    in_mora_dias = 0
End If
strCadena = "SELECT mora_monto,id_proveedor_servicio FROM entidad_parametros WHERE cod_unico='" & KEY_PROVEEDOR & "' LIMIT 1"
Call ConfiguraRstC(strCadena)
If rstc.RecordCount > 0 Then
    in_monto_mora = rstc("mora_monto")
Else
    in_monto_mora = 0
End If

strCadena = "call P_nueva_venta_temporal('00000000','" & KEY_PROVEEDOR & "')"
CnBd.Execute (strCadena)
        
strCadena = "call put_temporal_cobranza_ii('" & KEY_RUC & "','00000000','0096','001','" & KEY_ALM & "','" & in_mora_dias & "','" & in_monto_mora & "','" & KEY_PROVEEDOR & "')"
CnBd.Execute (strCadena)
Call llenarGrid_cobranza(Me.HfCuotas)



End Sub
Public Sub llenarGrid_cobranza(ByVal Grilla As MSHFlexGrid)
On Error GoTo salir
Dim tTotal As Double
Dim texonerado As Double
Dim tafecto As Double
Dim in_peso As Single
Dim in_servicio As String
Dim in_contador As Integer
strCadena = "SELECT * FROM view_venta_temporal WHERE id_dni='" & KEY_RUC & "' and  dni_save='00000000'  and  ruc='" & KEY_PROVEEDOR & "'"
Call ConfiguraRstC(strCadena)
If rstc.RecordCount < 1 Then
    Grilla.Rows = 0
   
    Exit Sub
End If
   
       Grilla.Rows = 0
       ReDim arrColWidth(1 To rstc.Fields.Count)
       For Each Campo In rstc.Fields
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 7500
           Grilla.ColWidth(2) = 800
           Grilla.ColWidth(3) = 1400
           Grilla.ColWidth(4) = 1400
          
       Next
        cabecera = "IDTEMPORAL" & vbTab & "DESCRIPCION " & vbTab & "CANT" & vbTab & "PRECIO" & vbTab & "TOTAL"
        Grilla.AddItem cabecera
         For k = 1 To 4
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        rstc.MoveFirst
        tafecto = 0
        For i = 0 To rstc.RecordCount - 1
            Fila = rstc("id") & vbTab & rstc("detalle") & vbTab & Format(rstc("cantidad"), "#,##0.00") & vbTab & Format(rstc("precio"), "#,##0.00") & vbTab & Format(rstc("total"), "#,##0.00")
            Grilla.AddItem Fila
            tafecto = tafecto + rstc("total")
            If Mid(rstc("detalle"), 1, 4) = "MORA" Then
                For k = 2 To 4
                    Grilla.col = k
                    Grilla.Row = i + 1
                    Grilla.CellBackColor = &H8080FF
                Next k
             End If
            rstc.MoveNext
    Next i
    Me.lblTotalDeuda.Caption = "S/." & Format(tafecto, "#,##0.00")
    
    
  
  Grilla.Row = 1
  Grilla.col = 0
  Grilla.ColSel = 5
  Grilla.RowSel = 1
salir:
Exit Sub
Set rst = Nothing
End Sub

Private Sub Timer1_Timer()
If Me.Label2.Visible = True Then
   Me.Label2.Visible = False
Else
    Me.Label2.Visible = True
End If
End Sub

Private Sub get_datos_cobranza()


strCadena = "SELECT * FROM almacen WHERE id_alm='00001' and  ruc='" & KEY_PROVEEDOR & "'"
Call ConfiguraRstL(strCadena)
If rstL.RecordCount > 0 Then
   Me.lbltelefono.Caption = rstL("telefonos") & Space(2) & "AREA DE COBRANZA"
Else
   Me.lbltelefono.Caption = "963819018" & Space(2) & "AREA DE COBRANZA"
End If


End Sub
