VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmVigilante1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8745
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16215
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8745
   ScaleWidth      =   16215
   ShowInTaskbar   =   0   'False
   Begin VitekeySoft.ChameleonBtn Command1 
      Height          =   375
      Left            =   4440
      TabIndex        =   9
      Top             =   240
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      BTYPE           =   5
      TX              =   "CONSULTAR"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
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
      MICON           =   "FrmVigilante1.frx":0000
      PICN            =   "FrmVigilante1.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox TxtPersona 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6720
      TabIndex        =   1
      Top             =   240
      Width           =   1575
   End
   Begin VB.TextBox TxtPlaca 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   10800
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker DtPInicio 
      Height          =   375
      Left            =   960
      TabIndex        =   2
      Top             =   240
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   57409537
      CurrentDate     =   40976
   End
   Begin MSComCtl2.DTPicker DtpFin 
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      Top             =   240
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   57409537
      CurrentDate     =   40976
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfOcurrencias 
      Height          =   7695
      Left            =   120
      TabIndex        =   8
      Top             =   960
      Width           =   15975
      _ExtentX        =   28178
      _ExtentY        =   13573
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
   Begin VitekeySoft.ChameleonBtn Command3 
      Height          =   375
      Left            =   8520
      TabIndex        =   10
      Top             =   240
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   5
      TX              =   "BUSCAR"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
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
      MICON           =   "FrmVigilante1.frx":05B6
      PICN            =   "FrmVigilante1.frx":05D2
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdcerrarpantalla 
      Height          =   345
      Left            =   13800
      TabIndex        =   11
      Top             =   240
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   609
      BTYPE           =   5
      TX              =   "CERRAR"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
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
      MICON           =   "FrmVigilante1.frx":0B6C
      PICN            =   "FrmVigilante1.frx":0B88
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn Command2 
      Height          =   375
      Left            =   12120
      TabIndex        =   12
      Top             =   240
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   5
      TX              =   "BUSCAR"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
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
      MICON           =   "FrmVigilante1.frx":3B9D
      PICN            =   "FrmVigilante1.frx":3BB9
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "a"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   2640
      TabIndex        =   7
      Top             =   320
      Width           =   90
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FECHA :"
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
      Height          =   195
      Left            =   240
      TabIndex        =   6
      Top             =   315
      Width           =   600
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DNI :"
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
      Height          =   195
      Left            =   6240
      TabIndex        =   5
      Top             =   315
      Width           =   375
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PLACA:"
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
      Height          =   195
      Left            =   10200
      TabIndex        =   4
      Top             =   315
      Width           =   540
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Height          =   615
      Left            =   6120
      Top             =   120
      Width           =   9975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Height          =   615
      Left            =   120
      Top             =   120
      Width           =   5895
   End
End
Attribute VB_Name = "FrmVigilante1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdcerrarpantalla_Click()
Unload Me
End Sub

Private Sub Command1_Click()
Call llenarOcurrenciasFecha(Me.HfOcurrencias)
End Sub

Private Sub Command4_Click()

End Sub
Private Sub llenarOcurrenciasFecha(ByVal Grilla As MSHFlexGrid)
On Error GoTo SALIR
strCadena = "SELECT A.id,A.fecha,A.hora,C.descripcion,A.dni,P.nombre_completo,C.id_acceso FROM persona_asistencia A,persona P,acceso C WHERE A.dni=P.dni AND A.ruc='" & KEY_RUC & "' AND A.id_acceso=C.id_acceso   AND fecha>='" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "' and fecha<='" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "' ORDER BY A.fecha DESC,A.dni,A.hora ASC"
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
            ' -- almacena el ancho de los campos
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 2000
           Grilla.ColWidth(2) = 1500
           Grilla.ColWidth(3) = 2000
           Grilla.ColWidth(4) = 1500
           Grilla.ColWidth(5) = 5000
           
       Next
         cabecera = "CODIGO" & vbTab & "FECHA" & vbTab & "HORA" & vbTab & "SAL/ING" & vbTab & "DNI" & vbTab & "TRABAJADOR"
         Grilla.AddItem cabecera
         For k = 0 To 5
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
                            Next k
                            
        rst.MoveFirst
        For i = 1 To rst.RecordCount
             
             Fila = rst("id") & vbTab & rst("fecha") & vbTab & Format(rst("hora"), "hh:mm:ss am/pm") & vbTab & rst("descripcion") & vbTab & rst("dni") & vbTab & rst("nombre_completo")
             Grilla.AddItem Fila
             Fila = ""
             
            For j = 0 To 5
                Grilla.col = j
                Grilla.Row = i
                If rst("id_acceso") = "01" Then
                    Grilla.CellBackColor = &HC0FFC0
                Else
                    Grilla.CellBackColor = &HC0C0FF
                End If
            Next j
        
        rst.MoveNext
        Next i
Exit Sub
SALIR: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub


Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
Call llenarOcurrenciasDNI(Me.HfOcurrencias)
End Sub
Private Sub llenarOcurrenciasDNI(ByVal Grilla As MSHFlexGrid)
On Error GoTo SALIR
strCadena = "SELECT A.id,A.fecha,A.hora,C.descripcion,A.dni,P.nombre_completo,V.nombre_completo as vigilante FROM persona_asistencia A,persona P,acceso C,persona V WHERE A.dni=P.dni AND A.ruc='" & KEY_RUC & "' AND A.id_acceso=C.id_acceso AND A.dni_save=V.dni  AND A.dni='" & Trim(Me.TxtPersona.Text) & "' ORDER BY A.id ASC"
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
            ' -- almacena el ancho de los campos
           Grilla.ColWidth(0) = 1000
           Grilla.ColWidth(1) = 2200
           Grilla.ColWidth(2) = 2000
           Grilla.ColWidth(3) = 1200
           Grilla.ColWidth(4) = 3500
           Grilla.ColWidth(5) = 2500
           
       Next
         cabecera = "CODIGO" & vbTab & "HORA" & vbTab & "SAL/ING" & vbTab & "DNI" & vbTab & "ENTIDAD" & vbTab & "VIJILANTE"
         Grilla.AddItem cabecera
         For k = 0 To 5
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
                            Next k
                            
        rst.MoveFirst
        For i = 1 To rst.RecordCount
             
             Fila = rst("id") & vbTab & rst("hora") & vbTab & rst("descripcion") & vbTab & rst("dni") & vbTab & rst("nombre_completo") & vbTab & rst("vigilante")
             Grilla.AddItem Fila
             Fila = ""
        rst.MoveNext
        Next i
Exit Sub
SALIR: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 150
Me.DtpInicio.Value = KEY_FECHA
Me.DtpFin.Value = KEY_FECHA
End Sub
