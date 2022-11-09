VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form FrmTipocambio 
   BorderStyle     =   0  'None
   Caption         =   "Tipo Cambio"
   ClientHeight    =   8490
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11070
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   11070
   ShowInTaskbar   =   0   'False
   Begin MSDataListLib.DataCombo DtcPeriodo 
      Height          =   315
      Left            =   6120
      TabIndex        =   12
      Top             =   8040
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
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
   Begin VitekeySoft.ChameleonBtn cmdEditar 
      Height          =   855
      Left            =   9720
      TabIndex        =   6
      Top             =   1320
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1508
      BTYPE           =   5
      TX              =   "MODIFICAR"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmTipocambio.frx":0000
      PICN            =   "FrmTipocambio.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdconsultar 
      Height          =   345
      Left            =   3000
      TabIndex        =   5
      Top             =   7965
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   609
      BTYPE           =   5
      TX              =   "CONSULTAR"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   8388608
      FCOLO           =   8388608
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmTipocambio.frx":32F2
      PICN            =   "FrmTipocambio.frx":330E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComCtl2.DTPicker DtpFecha 
      Height          =   345
      Left            =   1680
      TabIndex        =   0
      Top             =   7965
      Width           =   1335
      _ExtentX        =   2355
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
      Format          =   180027393
      CurrentDate     =   42648
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfgCambio 
      Height          =   7455
      Left            =   240
      TabIndex        =   1
      Top             =   390
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   13150
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
   Begin VitekeySoft.ChameleonBtn cmdCerrar 
      Height          =   855
      Left            =   9720
      TabIndex        =   7
      Top             =   3240
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1508
      BTYPE           =   5
      TX              =   "SALIR"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmTipocambio.frx":5662
      PICN            =   "FrmTipocambio.frx":567E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdsbs 
      Height          =   855
      Left            =   9720
      TabIndex        =   8
      Top             =   2280
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1508
      BTYPE           =   5
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmTipocambio.frx":86A5
      PICN            =   "FrmTipocambio.frx":86C1
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdupdate 
      Height          =   855
      Left            =   9720
      TabIndex        =   9
      Top             =   6960
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1508
      BTYPE           =   5
      TX              =   "UPDATE"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmTipocambio.frx":D3FB
      PICN            =   "FrmTipocambio.frx":D417
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdNuevo 
      Height          =   855
      Left            =   9720
      TabIndex        =   10
      Top             =   360
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1508
      BTYPE           =   5
      TX              =   "NUEVO"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmTipocambio.frx":FA50
      PICN            =   "FrmTipocambio.frx":FA6C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdConsultarPeriodo 
      Height          =   345
      Left            =   8280
      TabIndex        =   13
      Top             =   8040
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   609
      BTYPE           =   5
      TX              =   "CONSULTAR"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   8388608
      FCOLO           =   8388608
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmTipocambio.frx":FEBE
      PICN            =   "FrmTipocambio.frx":FEDA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CONSULTAR FECHA:"
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
      Left            =   4800
      TabIndex        =   11
      Top             =   8040
      Width           =   1335
   End
   Begin VB.Label LblEmpresa 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TIPO DE CAMBIO"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   255
      TabIndex        =   4
      Top             =   120
      Width           =   1305
   End
   Begin VB.Label lblValor 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Left            =   1815
      TabIndex        =   3
      Top             =   5880
      Width           =   105
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CONSULTAR FECHA:"
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
      Left            =   345
      TabIndex        =   2
      Top             =   8040
      Width           =   1335
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   8490
      Left            =   0
      Top             =   0
      Width           =   11070
   End
End
Attribute VB_Name = "FrmTipocambio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Procedencia As EnumProcede

Private Sub cmdCerrar_Click()
 Unload Me
End Sub

Private Sub cmdConsultar_Click()

strCadena = "SELECT *  FROM tipo_cambio WHERE id_creador='" & KEY_RUC & "' and fecha='" & Format(Me.DtpFecha.Value, "YYYY-mm-dd") & "' ORDER BY fecha DESC "
Call llenarGridC(Me.HfgCambio, Me)

End Sub

Private Sub cmdConsultarPeriodo_Click()

strCadena = "SELECT * FROM con_periodo WHERE id='" & Me.DtcPeriodo.BoundText & "' "
Call ConfiguraRst(strCadena)

strCadena = "SELECT *  FROM tipo_cambio WHERE id_creador='" & KEY_RUC & "' and fecha>='" & Format(rst("FechaInicio"), "YYYY-mm-dd") & "' and  fecha<='" & Format(rst("FechaFin"), "YYYY-mm-dd") & "' ORDER BY fecha DESC "
Call llenarGridC(Me.HfgCambio, Me)

End Sub

Private Sub cmdEditar_Click()
      Procedencia = modificar
      Call disabled_form(Me)
      FrmSeguridad.Show
      Exit Sub
End Sub

Private Sub cmdNuevo_Click()
      
      Procedencia = nuevo
      
      Call disabled_form(Me)
      FrmSeguridad.Show
      
      
      
      
      Exit Sub
End Sub

Private Sub cmdSbs_Click()
MsgBox "LOS RESULTADOS DE SBS CARGA AL INICIAR", vbInformation
End Sub

Private Sub cmdupdate_Click()
Dim fecha As Date
fecha = KEY_FECHA
For i = 0 To 100
        
        Call get_cambio_sbs(fecha)
        fecha = DateAdd("d", -1, fecha)
        DoEvents
Next i
MsgBox "Finalizado", vbInformation
End Sub

Public Sub actualizar()
strCadena = "SELECT *  FROM tipo_cambio WHERE id_creador='" & KEY_RUC & "'  ORDER BY fecha DESC LIMIT 31 "
Call llenarGridC(Me.HfgCambio, Me)
End Sub
Public Sub llenarGridC(ByVal Grilla As MSHFlexGrid, ByVal Formulario As Form)
On Error GoTo salir
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 0
    
     
    Exit Sub
End If
    

   Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
            ' -- almacena el ancho de los campos
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 1300
           Grilla.ColWidth(2) = 1400
           Grilla.ColWidth(3) = 1100
           Grilla.ColWidth(4) = 1400
           Grilla.ColWidth(5) = 1100
           Grilla.ColWidth(6) = 1400
           Grilla.ColWidth(7) = 1100
        Next
        cabecera = "CODIGO" & vbTab & "FECHA" & vbTab & "COMPRA " & vbTab & "VALOR" & vbTab & "VENTA" & vbTab & "VALOR" & vbTab & "LOCAL" & vbTab & "VALOR"
        Grilla.AddItem cabecera
         For k = 0 To 7
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
            Fila = rst("id_tipocambio") & vbTab & Format(rst("fecha"), "dd-mmm-YYYY") & vbTab & "PRECIO COMPRA" & vbTab & Format(rst("valor_compra"), "#,##0.0000") & vbTab & "PRECIO VENTA" & vbTab & Format(rst("valor_venta"), "#,##0.0000") & vbTab & "CAMBIO LOCAL" & vbTab & Format(rst("valor_local"), "#,##0.0000")
            Grilla.AddItem Fila
            
            rst.MoveNext
    Next i
        
  
  Grilla.Row = 1
  Grilla.col = 0
  Grilla.ColSel = 1
  Grilla.RowSel = 1
  
  
  Exit Sub
salir: MsgBox "Ocurrio un Error en la Conexion", vbInformation, "Mensaje para el usuario"
Set rst = Nothing
End Sub
Private Sub HfgLinea_Click()
If HfgLinea.Row > 0 Then
     Me.cmdEditar.Enabled = True
  End If
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 150
  strCadena = "SELECT Id as Codigo,CONCAT(Nombre,'-',Ejercicio) as Descripcion  FROM con_periodo ORDER BY Ejercicio DESC,Mes DESC"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcPeriodo)
  Me.DtcPeriodo.BoundText = get_periodo_actual(KEY_FECHA)
  
  
  
Me.DtpFecha.Value = KEY_FECHA
Call actualizar
End Sub

Private Sub HfgCambio_Click()
If Val(Me.HfgCambio.TextMatrix(Me.HfgCambio.Row, 0)) > 0 Then
   Me.cmdEditar.Enabled = True
Else
   Me.cmdEditar.Enabled = False
End If

End Sub


Sub llenarGridME(ByVal Grilla As MSHFlexGrid, ByVal Formulario As Form)
On Error GoTo salir
  Call ConfiguraRst(strCadena)
  Grilla.Clear
  Grilla.Rows = rst.RecordCount
  Set Grilla.Recordset = rst
  Grilla.ColWidth(0) = 1500
  Grilla.ColWidth(1) = 3500
  
Call DarFormatoFecha(Grilla, 0)
Call DarFormato(Grilla, 1)
Grilla.Refresh

  
  Me.cmdEditar.Enabled = False
  strCadena = "SELECT * FROM tipo_cambio WHERE fecha='" & CVDate(Date) & "'"
  Call ConfiguraRst(strCadena)
  Me.lblValor.Caption = "Tipo Cambio:" + Space(3) + Format(rst("valor"), "#,##0.00")
  Set rst = Nothing
  Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub






