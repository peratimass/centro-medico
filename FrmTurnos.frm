VERSION 5.00
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form FrmTurnos 
   BorderStyle     =   0  'None
   Caption         =   "Form3"
   ClientHeight    =   8295
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9540
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8295
   ScaleWidth      =   9540
   ShowInTaskbar   =   0   'False
   Begin VB.Frame FrameDetalle 
      BackColor       =   &H00FFFFFF&
      Height          =   4215
      Left            =   600
      TabIndex        =   6
      Top             =   1080
      Visible         =   0   'False
      Width           =   7215
      Begin VitekeySoft.ChameleonBtn cmdProcesar 
         Height          =   615
         Left            =   2160
         TabIndex        =   16
         Top             =   3480
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   1085
         BTYPE           =   3
         TX              =   "PROCESAR"
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
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmTurnos.frx":0000
         PICN            =   "FrmTurnos.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.TextBox TxtIdTurno 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   240
         TabIndex        =   15
         Top             =   2760
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox TxtDescripcion 
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   1935
         MaxLength       =   120
         TabIndex        =   10
         Top             =   480
         Width           =   2685
      End
      Begin VB.TextBox TxtHoraInicial 
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   1890
         MaxLength       =   120
         TabIndex        =   9
         Top             =   1200
         Width           =   1845
      End
      Begin VB.TextBox TxtHoraFinal 
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   1890
         MaxLength       =   120
         TabIndex        =   8
         Top             =   2160
         Width           =   1845
      End
      Begin VB.TextBox txtTotalHoras 
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   1890
         MaxLength       =   120
         TabIndex        =   7
         Top             =   1680
         Width           =   1845
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   5640
         Top             =   1200
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
               Picture         =   "FrmTurnos.frx":3664
               Key             =   "(Aceptar)"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmTurnos.frx":3980
               Key             =   "(Eliminar)"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmTurnos.frx":3DE0
               Key             =   "(Inicio)"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmTurnos.frx":4240
               Key             =   "(Modificar)"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmTurnos.frx":455C
               Key             =   "(Nuevo)"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmTurnos.frx":49BC
               Key             =   "(Quitar)"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmTurnos.frx":4CD8
               Key             =   "(Salir)"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmTurnos.frx":5138
               Key             =   "(Red)"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmTurnos.frx":5598
               Key             =   "(Grabar)"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmTurnos.frx":5E78
               Key             =   "(Agregar)"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmTurnos.frx":6194
               Key             =   "(Buscar)"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmTurnos.frx":64B0
               Key             =   "(Cancelar)"
            EndProperty
         EndProperty
      End
      Begin VitekeySoft.ChameleonBtn cmdcerrar 
         Height          =   615
         Left            =   4320
         TabIndex        =   17
         Top             =   3480
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   1085
         BTYPE           =   3
         TX              =   "CERRAR"
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
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmTurnos.frx":67CC
         PICN            =   "FrmTurnos.frx":67E8
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label LblDescripcion 
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
         Left            =   600
         TabIndex        =   14
         Top             =   525
         Width           =   1005
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "HORA INICIAL :"
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
         Left            =   660
         TabIndex        =   13
         Top             =   1245
         Width           =   1005
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "HORA FINAL :"
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
         Left            =   765
         TabIndex        =   12
         Top             =   2280
         Width           =   915
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "HORAS :"
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
         Left            =   1140
         TabIndex        =   11
         Top             =   1800
         Width           =   555
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H00FFFFFF&
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   6  'Inside Solid
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   2175
         Left            =   120
         Top             =   1080
         Width           =   6255
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00FFFFFF&
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   6  'Inside Solid
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   555
         Left            =   120
         Top             =   360
         Width           =   6255
      End
   End
   Begin VB.TextBox TxtUnidad 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2040
      TabIndex        =   0
      Top             =   7710
      Width           =   3615
   End
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
            Picture         =   "FrmTurnos.frx":991F
            Key             =   "(Nuevo)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTurnos.frx":9D73
            Key             =   "(Modificar)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTurnos.frx":A093
            Key             =   "(Eliminar)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTurnos.frx":A4E7
            Key             =   "(Salir)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTurnos.frx":A93B
            Key             =   "(Aceptar)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTurnos.frx":AC5B
            Key             =   "(Cancelar)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTurnos.frx":AF7B
            Key             =   "(Quitar)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTurnos.frx":B29B
            Key             =   "(Agregar)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTurnos.frx":B5BB
            Key             =   "(Buscar)"
         EndProperty
      EndProperty
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfgLinea 
      Height          =   7095
      Left            =   120
      TabIndex        =   1
      Top             =   390
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   12515
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
      Height          =   7185
      Left            =   8445
      TabIndex        =   2
      Top             =   360
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   12674
      BandCount       =   1
      ForeColor       =   8388608
      ImageList       =   "ImgIconos"
      FixedOrder      =   -1  'True
      VariantHeight   =   0   'False
      Orientation     =   1
      EmbossPicture   =   -1  'True
      _CBWidth        =   900
      _CBHeight       =   7185
      _Version        =   "6.0.8169"
      Caption1        =   "Acciones"
      Child1          =   "TlbAcciones"
      MinHeight1      =   840
      Width1          =   3180
      FixedBackground1=   0   'False
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar TlbAcciones 
         Height          =   5670
         Left            =   30
         TabIndex        =   3
         Top             =   375
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   10001
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
   Begin VB.Label LblFecha 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TURNO :"
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
      Left            =   825
      TabIndex        =   5
      Top             =   7710
      Width           =   585
   End
   Begin VB.Label LblEmpresa 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TURNOS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   90
      TabIndex        =   4
      Top             =   120
      Width           =   675
   End
   Begin VB.Shape ShpDatos 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   555
      Left            =   120
      Top             =   7590
      Width           =   8055
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   8295
      Left            =   0
      Top             =   0
      Width           =   9540
   End
End
Attribute VB_Name = "FrmTurnos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub actualizar()
  strCadena = "SELECT * FROM turno  WHERE  ruc='" & KEY_RUC & "'ORDER BY hora_inicio ASC,hora_final ASC,descripcion"
  Call llenarGrid(Me.HfgLinea, Me)
End Sub
Private Sub llenarGrid(ByVal Grilla As MSHFlexGrid, ByVal Formulario As Form)
Dim strHora As String
On Error GoTo salir
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 0
    
    Exit Sub

End If
  
   
   
   Grilla.Rows = 0
      ' Me.HfdGrilla.Rows = rst.RecordCount - 2
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
            ' -- almacena el ancho de los campos
           Grilla.ColWidth(0) = 1000
           Grilla.ColWidth(1) = 3500
           Grilla.ColWidth(2) = 1800
           Grilla.ColWidth(3) = 1200
           
           
        Next
         cabecera = "CODIGO" & vbTab & "DESCRIPCION" & vbTab & "[ H.INICIO - H.FIN   ]" & vbTab & "HORAS"
         Grilla.AddItem cabecera
         For k = 0 To 3
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
                            Next k
                            
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
        
             
             Fila = rst("id_turno") & vbTab & rst("descripcion") & vbTab & "[ " & Format(rst("hora_inicio"), "hh:mm") & "   -   " & Format(rst("hora_final"), "hh:mm") & "  ]" & vbTab & ""
            
            
          Grilla.AddItem Fila
            
        Fila = ""
        rst.MoveNext
        Next i
Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub


Private Sub cmdcerrar_Click()
Me.FrameDetalle.Visible = False
End Sub

Private Sub cmdprocesar_Click()
        





If Val(Me.TxtIdTurno.Text) > 0 Then
    
    
    strCadena = "UPDATE turno SET descripcion='" & Trim(Me.TxtDescripcion.Text) & "',hora_inicio='" & Trim(Me.TxtHoraInicial.Text) & "',hora_final='" & Trim(Me.TxtHoraFinal.Text) & "',horas='" & Trim(Me.txtTotalHoras.Text) & "' WHERE id_turno='" & Trim(Me.TxtIdTurno.Text) & "' AND ruc='" & KEY_RUC & "'"
    CnBd.Execute (strCadena)
     
Else
        strCadena = "SELECT * FROM turno WHERE ruc='" & KEY_RUC & "' ORDER BY id_turno DESC"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount > 0 Then
            rst.MoveFirst
            strcodigo = formato_item(Val(rst("id_turno")) + 1, 2)
        Else
            strcodigo = "01"
        End If
        strCadena = "INSERT INTO turno (id_turno,descripcion,hora_inicio,hora_final,horas,ruc) VALUES('" & strcodigo & "','" & Trim(Me.TxtDescripcion.Text) & "','" & Trim(Me.TxtHoraInicial.Text) & "','" & Trim(Me.TxtHoraFinal.Text) & "','" & Trim(Me.txtTotalHoras.Text) & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
         
    
End If


Me.TxtDescripcion.Text = ""
Me.TxtIdTurno.Text = ""
Me.TxtHoraFinal.Text = ""
Me.TxtHoraInicial.Text = ""
Me.TxtDescripcion.Text = ""


Me.FrameDetalle.Visible = False
Call actualizar
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 50
Call actualizar
End Sub

Private Sub TlbAcciones_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.key
    Case KEY_NEW
      Me.TxtIdTurno.Text = 0
      Me.TxtDescripcion.Text = ""
      Me.TxtHoraInicial.Text = ""
      Me.TxtHoraFinal.Text = ""
      Me.txtTotalHoras.Text = ""
      Me.FrameDetalle.Visible = True
      Call Resalta(Me.TxtDescripcion)
      Exit Sub
    
    Case KEY_UPDATE
      Me.FrameDetalle.Visible = True
      Me.TxtIdTurno.Text = Me.HfgLinea.TextMatrix(Me.HfgLinea.Row, 0)
      strCadena = "SELECT * FROM turno WHERE id_turno='" & Trim(Me.TxtIdTurno.Text) & "' and ruc='" & KEY_RUC & "'"
      Call ConfiguraRst(strCadena)
      If rst.RecordCount > 0 Then
        Me.TxtDescripcion.Text = rst("descripcion")
        Me.TxtHoraInicial.Text = Format(rst("hora_inicio"), "hh:mm:ss")
        Me.TxtHoraFinal.Text = Format(rst("hora_final"), "hh:mm:ss")
        Me.txtTotalHoras.Text = Format(rst("horas"), "hh:mm:ss")
      End If
   Exit Sub
      
      
    Case KEY_DELETE
      If MsgBox(MSGELIMINAR, vbQuestion + vbYesNo, MSGVERIFICACION) = vbYes Then
        strCadena = "DELETE  FROM turno WHERE id_turno='" & Trim(Me.HfgLinea.TextMatrix(Me.HfgLinea.Row, 0)) & "' AND ruc='" & KEY_RUC & "'"
        CnBd.Execute (strCadena)
         
        Call actualizar
     End If
    
    
    Case KEY_EXIT
        Unload Me
  End Select
End Sub

Private Sub TxtHoraFinal_KeyPress(KeyAscii As Integer)
Dim INICIO As Variant
Dim final As Variant
On Error GoTo salir
If KeyAscii = 13 Then
    Me.TxtHoraFinal.Text = Format(Me.TxtHoraFinal.Text, "hh:mm:ss")
    INICIO = Format(Me.TxtHoraInicial.Text, "hh:mm:ss")
    final = Format(Me.TxtHoraFinal.Text, "hh:mm:ss")
    Me.txtTotalHoras = Format(TimeValue(final) - TimeValue(INICIO), "hh:mm:ss")

End If
salir:
Exit Sub
End Sub

Private Sub TxtHoraInicial_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.TxtHoraInicial.Text = Format(Me.TxtHoraInicial.Text, "hh:mm:ss")
    Call Resalta(Me.TxtHoraFinal)
End If
End Sub

Private Sub txtTotalHoras_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.txtTotalHoras.Text = Format(Me.txtTotalHoras.Text, "hh:mm:ss")
    INICIO = Format(Me.TxtHoraInicial.Text, "hh:mm:ss")
    
    Me.TxtHoraFinal.Text = Format(TimeValue(INICIO) + TimeValue(txtTotalHoras.Text), "hh:mm:ss")
End If

End Sub
