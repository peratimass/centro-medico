VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form FrmDetallePagosFactura 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   4815
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10815
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   10815
   ShowInTaskbar   =   0   'False
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfdDetalle 
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   5741
      _Version        =   393216
      ForeColor       =   8388608
      FixedCols       =   0
      ForeColorFixed  =   8388608
      BackColorBkg    =   16777215
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
   Begin VB.Label lblcambio 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   9240
      TabIndex        =   10
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label lblfecha 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   9240
      TabIndex        =   9
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label lbldireccion 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   1680
      TabIndex        =   8
      Top             =   840
      Width           =   5535
   End
   Begin VB.Label lblrazonsocial 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   1680
      TabIndex        =   7
      Top             =   480
      Width           =   5535
   End
   Begin VB.Label lblruc 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   1680
      TabIndex        =   6
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "T.CAMBIO :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   165
      Left            =   8280
      TabIndex        =   5
      Top             =   600
      Width           =   900
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "FECHA EMISION :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   165
      Left            =   7830
      TabIndex        =   4
      Top             =   240
      Width           =   1350
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "RAZON SOCIAL :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   165
      Left            =   180
      TabIndex        =   3
      Top             =   480
      Width           =   1275
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "DIRECCION :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   165
      Left            =   480
      TabIndex        =   2
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "DNI / RUC :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   165
      Left            =   600
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      Height          =   4815
      Left            =   0
      Top             =   0
      Width           =   10815
   End
End
Attribute VB_Name = "FrmDetallePagosFactura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
CenterForm Me
Me.Top = 200
Call llena_pagosVenta(Me.HfdDetalle, FrmReporteRegistroVentas.HfdPersona.TextMatrix(FrmReporteRegistroVentas.HfdPersona.Row, 0))

End Sub
Public Sub llena_pagosVenta(ByVal Grilla As MSHFlexGrid, ByVal idVenta As Double)
On Error GoTo salir
Dim tpago As Double
strCadena = "SELECT * FROM movimiento_venta_monto M,forma_pago_detalle F WHERE M.id_forma_pago=F.id_detalle AND id_venta='" & idVenta & "' AND M.ruc='" & KEY_RUC & "' AND F.ruc='" & KEY_RUC & "' "
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
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 2000
           Grilla.ColWidth(2) = 1500
       Next
        cabecera = "CODIGO" & vbTab & "FORMA PAGO" & vbTab & "MONTO"
        Grilla.AddItem cabecera
         For k = 0 To 2
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        rst.MoveFirst
        tpago = 0
        For i = 0 To rst.RecordCount - 1
            Fila = rst("id_detalle") & vbTab & rst("descripcion") & vbTab & Format(rst("monto"), "###0.00")
            Grilla.AddItem Fila
            tpago = rst("monto") + tpago
            rst.MoveNext
    Next i

Exit Sub
salir: MsgBox "Ocurrio un Error en la Conexion", vbInformation, "Mensaje para el usuario"
Set rst = Nothing
End Sub

