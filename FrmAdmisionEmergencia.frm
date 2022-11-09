VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form FrmAdmisionEmergencia 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9240
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   18210
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9240
   ScaleWidth      =   18210
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   13080
      Top             =   2760
   End
   Begin VitekeySoft.ChameleonBtn ChameleonBtn1 
      Height          =   615
      Left            =   12720
      TabIndex        =   7
      Top             =   3480
      Width           =   5400
      _ExtentX        =   9525
      _ExtentY        =   1085
      BTYPE           =   5
      TX              =   "PEDIDOS PENDIENTES"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmAdmisionEmergencia.frx":0000
      PICN            =   "FrmAdmisionEmergencia.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfdPersona 
      Height          =   8685
      Left            =   120
      TabIndex        =   0
      Top             =   525
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   15319
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
   Begin MSComCtl2.MonthView MonthView1 
      Height          =   4710
      Left            =   12720
      TabIndex        =   2
      Top             =   4440
      Width           =   5820
      _ExtentX        =   10266
      _ExtentY        =   8308
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      StartOfWeek     =   134021121
      TitleBackColor  =   8438015
      TrailingForeColor=   12632256
      CurrentDate     =   41323
   End
   Begin VB.TextBox txtNombrePaciente 
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
      Left            =   4680
      TabIndex        =   6
      Top             =   100
      Width           =   2415
   End
   Begin VB.TextBox txtDni 
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
      Left            =   1080
      TabIndex        =   4
      Top             =   100
      Width           =   1935
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   3285
      Left            =   12720
      TabIndex        =   8
      Top             =   0
      Width           =   5400
      _ExtentX        =   9525
      _ExtentY        =   5794
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
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CLIENTE"
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
      Left            =   3600
      TabIndex        =   5
      Top             =   135
      Width           =   660
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RUC/DNI"
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
      Left            =   240
      TabIndex        =   3
      Top             =   135
      Width           =   675
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   450
      Left            =   120
      Top             =   15
      Width           =   12495
   End
   Begin VB.Label lblPacientes 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   18465
      TabIndex        =   1
      Top             =   5685
      Width           =   75
   End
End
Attribute VB_Name = "FrmAdmisionEmergencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Procedencia As EnumProcede

Public Sub buscar_emergencia(ByVal Grilla As MSHFlexGrid, ByVal sql As String)
'On Error GoTo SALIR
'If KEY_TIPO_ATENCION = "03" Then
   ' strCadena = "SELECT A.id_detalle,A.id_hora,A.dni,A.dni_paciente,P.nombre_completo,P.sexo,P.direccion,S.descripcion,A.prioridad,A.emergencia,T.descripcion as destino,A.cie10,A.cie10descripcion,ron_func_edad_v2(P.id_dia,P.id_mes,P.id_anio) as edad,A.id_atencion FROM agenda A ,persona P,seguro_medico_detalle S,triaje_destino T WHERE A.destino=T.codigo  AND A.id_seguro=S.id_detalle AND   A.dni_paciente=P.dni AND A.id_tipo='09' AND A.id_fecha='" & Format(Me.MonthView1.Value, "YYYY-mm-dd") & "' AND A.ruc='" & KEY_RUC & "' AND A.id_seguro<>'00006' ORDER BY A.id_hora ASC"
'Else
    
    
'End If
Call ConfiguraRst(sql)
If rst.RecordCount < 1 Then
    Grilla.rows = 1
    Grilla.Clear
    Exit Sub

End If
  
   Grilla.Clear
   Grilla.Refresh
   Grilla.rows = 0
      ' Me.HfdGrilla.Rows = rst.RecordCount - 2
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
            ' -- almacena el ancho de los campos
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 400
           Grilla.ColWidth(2) = 1200
           Grilla.ColWidth(3) = 800
           Grilla.ColWidth(4) = 2000
           Grilla.ColWidth(5) = 1200
           Grilla.ColWidth(6) = 4500
           Grilla.ColWidth(7) = 2000
           
        Next
         cabecera = "ID VENTA" & vbTab & "Nº ATENCION" & vbTab & "FECHA" & vbTab & "HORA" & vbTab & "COMPROBANTE" & vbTab & "RUC/DNI" & vbTab & "CLIENTE" & vbTab & "RESPONSABLE"
         Grilla.AddItem cabecera
         For k = 1 To 7
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
                            Next k
                            
        rst.MoveFirst
        For i = 1 To rst.RecordCount
             
            If rst("id_almacenero") <> "0" Then
                strCadena = "SELECT nombre_completo FROM persona WHERE dni='" & rst("id_almacenero") & "'"
                Call ConfiguraRstK(strCadena)
                If rstK.RecordCount > 0 Then
                    almacenero = rstK("nombre_completo")
                Else
                    almacenero = "-"
                End If
            Else
                almacenero = "-"
             End If
             
             Fila = str(rst("id_venta")) & vbTab & Format(i, "00") & vbTab & Format(rst("fecha_emision"), "dd-mm-YYYY") & vbTab & Format(rst("hora"), "HH:mm") & vbTab & rst("documento") & vbTab & rst("id_cliente") & vbTab & rst("ncliente") & vbTab & almacenero
             Grilla.AddItem Fila
             
             
             
             If rst("id_almacenero") = "0" Then
                                For k = 0 To 7
                                    Grilla.col = k
                                    Grilla.Row = i
                                    Grilla.CellBackColor = &H8080FF
                                Next k
             
             Else
                
                        For k = 0 To 7
                                    Grilla.col = k
                                    Grilla.Row = i
                                    Grilla.CellBackColor = &H80FF80
                        Next k
                
                  
             
             End If
             
            
        Fila = ""
        rst.MoveNext
        Next i
        
'Exit Sub
'SALIR: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub




















Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
    Unload Me
End If
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 50
Me.MonthView1.Value = KEY_FECHA

strCadena = "SELECT * FROM movimiento_venta WHERE fecha_emision='" & KEY_FECHA & "' AND (id_doc='0001' OR id_doc='0003') AND ruc='" & KEY_RUC & "' "
Call buscar_emergencia(FrmAdmisionEmergencia.HfdPersona, strCadena)
End Sub



Private Sub MonthView1_DateClick(ByVal DateClicked As Date)
strCadena = "SELECT A.id_detalle,A.id_estado,A.id_seguro,A.id_fua,A.id_hora,A.dni,A.dni_paciente,M.nombre_completo as medico,P.nombre_completo,P.sexo,P.direccion,S.descripcion,A.prioridad,A.emergencia,T.descripcion as destino,A.cie10,A.cie10descripcion,ron_func_edad_v2(P.id_dia,P.id_mes,P.id_anio) as edad,A.id_atencion,A.pendiente_sis,A.id_movimiento,A.documento FROM agenda A LEFT JOIN persona M ON A.dni=M.dni,persona P,seguro_medico_detalle S,triaje_destino T WHERE A.destino=T.codigo  AND A.id_seguro=S.id_detalle AND   A.dni_paciente=P.dni AND A.id_tipo='09' AND A.id_fecha='" & Format(Me.MonthView1.Value, "YYYY-mm-dd") & "' AND A.ruc='" & KEY_RUC & "'  ORDER BY A.id_hora ASC"
Call buscar_emergencia(Me.HfdPersona, strCadena)
End Sub



Private Sub Timer1_Timer()
strCadena = "select id_venta,id_doc,serie,numero,id_tipo_factura from movimiento_venta WHERE id_almacenero='0' AND ruc='" & KEY_RUC & "' AND fecha_emision='" & KEY_FECHA & "' AND (id_doc='0001' OR id_doc='0003') LIMIT 0,1"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    strCadena = "UPDATE movimiento_venta SET id_almacenero='" & KEY_USUARIO & "' WHERE id_venta='" & rst("id_venta") & "'"
    CnBd.Execute (strCadena)
    Call insertar_acciones(strCadena)
    
    PlaySound App.Path & "\sonidos\dingding.wav"
    Call Orden_Impresion(rst("id_doc"), rst("serie"), rst("numero"), rst("id_tipo_factura"), "")
              
              
    strCadena = "SELECT * FROM movimiento_venta WHERE fecha_emision='" & KEY_FECHA & "' AND (id_doc='0001' OR id_doc='0003') AND ruc='" & KEY_RUC & "' "
    Call Me.buscar_emergencia(Me.HfdPersona, strCadena)
    
End If

End Sub

Private Sub txtDNI_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    strCadena = "SELECT A.id_detalle,A.id_estado,A.id_seguro,A.id_fua,A.id_hora,A.dni,A.dni_paciente,M.nombre_completo as medico,P.nombre_completo,P.sexo,P.direccion,S.descripcion,A.prioridad,A.emergencia,T.descripcion as destino,A.cie10,A.cie10descripcion,ron_func_edad_v2(P.id_dia,P.id_mes,P.id_anio) as edad,A.id_atencion,A.pendiente_sis,A.id_movimiento,A.documento,A.destino as id_destino FROM agenda A LEFT JOIN persona M ON A.dni=M.dni,persona P,seguro_medico_detalle S,triaje_destino T WHERE A.destino=T.codigo  AND A.id_seguro=S.id_detalle AND   A.dni_paciente=P.dni AND A.id_tipo='09' AND A.id_fecha='" & Format(Me.MonthView1.Value, "YYYY-mm-dd") & "' AND A.ruc='" & KEY_RUC & "' AND A.dni_paciente LIKE '%" & Trim(Me.txtDni.Text) & "%'  ORDER BY A.id_hora ASC"
    Call Me.buscar_emergencia(Me.HfdPersona, strCadena)
End If
End Sub

Private Sub txtNombrePaciente_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    strCadena = "SELECT A.id_detalle,A.id_estado,A.id_seguro,A.id_fua,A.id_hora,A.dni,A.dni_paciente,M.nombre_completo as medico,P.nombre_completo,P.sexo,P.direccion,S.descripcion,A.prioridad,A.emergencia,T.descripcion as destino,A.cie10,A.cie10descripcion,ron_func_edad_v2(P.id_dia,P.id_mes,P.id_anio) as edad,A.id_atencion,A.pendiente_sis,A.id_movimiento,A.documento,A.destino as id_destino FROM agenda A LEFT JOIN persona M ON A.dni=M.dni,persona P,seguro_medico_detalle S,triaje_destino T WHERE A.destino=T.codigo  AND A.id_seguro=S.id_detalle AND   A.dni_paciente=P.dni AND A.id_tipo='09' AND A.id_fecha='" & Format(Me.MonthView1.Value, "YYYY-mm-dd") & "' AND A.ruc='" & KEY_RUC & "' AND P.nombre_completo LIKE '%" & Trim(Me.txtNombrePaciente.Text) & "%'  ORDER BY A.id_hora ASC"
    Call Me.buscar_emergencia(Me.HfdPersona, strCadena)
End If

End Sub
