VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmventas_pedidos 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7935
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10575
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7935
   ScaleWidth      =   10575
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtcliente 
      Appearance      =   0  'Flat
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
      Height          =   285
      Left            =   1560
      TabIndex        =   0
      Top             =   480
      Width           =   2775
   End
   Begin VitekeySoft.ChameleonBtn cmdconsultar 
      Height          =   345
      Left            =   8400
      TabIndex        =   1
      Top             =   120
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "CONSULTAR"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   8388608
      FCOLO           =   8388608
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmventas_pedidos.frx":0000
      PICN            =   "frmventas_pedidos.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   300
      Left            =   1560
      TabIndex        =   2
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   529
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
      CalendarForeColor=   8388608
      CalendarTitleForeColor=   8388608
      Format          =   125173761
      CurrentDate     =   42236
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfPendientes 
      Height          =   6735
      Left            =   360
      TabIndex        =   3
      Top             =   960
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   11880
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
   Begin VitekeySoft.ChameleonBtn cmdsalir 
      Height          =   345
      Left            =   8400
      TabIndex        =   4
      Top             =   480
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "SALIR          "
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   8388608
      FCOLO           =   8388608
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmventas_pedidos.frx":2601
      PICN            =   "frmventas_pedidos.frx":261D
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "FECHA "
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "CLIENTE"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   480
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   7935
      Left            =   0
      Top             =   0
      Width           =   10575
   End
End
Attribute VB_Name = "frmventas_pedidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdConsultar_Click()

Dim strpersona As String
strpersona = ""
strpersona = Trim(Me.txtcliente.Text)

strCadena = "SELECT id_venta,ncliente,documento,total,vendedor FROM view_listado_pendientes_II WHERE ncliente LIKE '%" & strpersona & "%' AND  fecha_emision='" & Format(Me.DTPicker1.Value, "YYYY-mm-dd") & "' and id_alm='" & KEY_ALM & "' AND ruc='" & KEY_RUC & "' ORDER BY id_venta DESC"
Call llenar_pendientes(Me.HfPendientes)
End Sub

Private Sub cmdsalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim pantalla As Single
CenterForm Me
'pantalla = Screen.Width
'pantalla1 = (pantalla - FrmVentas.Width) / 2
Me.DTPicker1.Value = KEY_FECHA
'pantalla3 = FrmVentas.Width - pantalla1
'Me.Left = pantalla3 - 500
Me.Top = 50
strCadena = "SELECT id_venta,ncliente,documento,total,vendedor FROM view_listado_pendientes_II WHERE  fecha_emision='" & KEY_FECHA & "' and id_alm='" & KEY_ALM & "' AND ruc='" & KEY_RUC & "' ORDER BY id_venta DESC"
Call llenar_pendientes(Me.HfPendientes)
End Sub
Public Sub llenar_pendientes(ByVal Grilla As MSHFlexGrid)

Call ConfiguraRstI(strCadena)
If rstI.RecordCount < 1 Then
    Grilla.Rows = 1
    Grilla.Clear
    Exit Sub
End If
   Grilla.Clear
   Grilla.Refresh
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rstI.Fields.Count)
       
        For Each Campo In rstI.Fields
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 1200
           Grilla.ColWidth(2) = 2300
           Grilla.ColWidth(3) = 1000
           Grilla.ColWidth(4) = 2200
           Grilla.ColWidth(5) = 500
           
        Next
        cabecera = "IDVENTA" & vbTab & "PROFORMA" & vbTab & "CLIENTE" & vbTab & "MONTO" & vbTab & "VENDEDOR" & vbTab & ""
        Grilla.AddItem cabecera
         For k = 1 To 5
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
         Next k
        rstI.MoveFirst
        c = 5
        NumeroCampo = 5
            
        For i = 0 To rstI.RecordCount - 1
          estado = Chr(168)
          descripcion = ""
            ndocumento = Split(rstI("documento"), ":")
            nproforma = "P:" & ndocumento(1)
            
          Fila = rstI("id_venta") & vbTab & nproforma & vbTab & rstI("ncliente") & vbTab & Format(rstI("total"), "#,##0.00") & vbTab & Mid(rstI("vendedor"), 1, 20) & vbTab & estado
          Grilla.AddItem Fila
        
        If c = NumeroCampo Then
                        With Grilla
                            .Row = i + 1 ' se posiciona en la fila
                            .col = c '  .. en la columna
                            .CellFontName = "Wingdings"
                            .CellFontSize = 14
                            .CellAlignment = flexAlignCenterCenter
                            
                            
                        End With
        End If
        Fila = ""
          
          rstI.MoveNext
      Next i
    
End Sub

Private Sub HfPendientes_DblClick()
If Val(Me.HfPendientes.TextMatrix(Me.HfPendientes.Row, 0)) > 0 Then
    FrmVentas.txt_id_pendiente.Text = Val(Me.HfPendientes.TextMatrix(Me.HfPendientes.Row, 0))
    Call FrmVentas.get_comprobante(Val(Me.HfPendientes.TextMatrix(Me.HfPendientes.Row, 0)))
    Call FrmVentas.llena_pagos(FrmVentas.HfgTipoPagos, FrmVentas.TxtNumeroDoc.Text)
    FrmVentas.timer_pendientes.Enabled = False
    Unload Me
End If
End Sub

