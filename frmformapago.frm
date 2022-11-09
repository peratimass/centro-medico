VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmformapago 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6780
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16635
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   16635
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmdetalle 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5655
      Left            =   3720
      TabIndex        =   7
      Top             =   600
      Visible         =   0   'False
      Width           =   8175
      Begin VB.CheckBox chk_habilitado 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "HABILITADO"
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
         Height          =   255
         Left            =   1800
         TabIndex        =   20
         Top             =   4440
         Width           =   1605
      End
      Begin VB.TextBox txtObservacion 
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
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   1800
         TabIndex        =   18
         Top             =   1320
         Width           =   2295
      End
      Begin VB.TextBox txtid 
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
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   7200
         TabIndex        =   15
         Top             =   480
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox txtcuentacontable 
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
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   1800
         TabIndex        =   13
         Top             =   1800
         Width           =   2295
      End
      Begin MSDataListLib.DataCombo DtcDetallePago 
         Height          =   330
         Left            =   1800
         TabIndex        =   9
         Top             =   840
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         ForeColor       =   8388608
         Text            =   "DataCombo1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VitekeySoft.ChameleonBtn cmdprocear 
         Height          =   555
         Left            =   1800
         TabIndex        =   10
         Top             =   4920
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   979
         BTYPE           =   5
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
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   12582912
         FCOLO           =   12582912
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmformapago.frx":0000
         PICN            =   "frmformapago.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdsalir 
         Height          =   555
         Left            =   3480
         TabIndex        =   11
         Top             =   4920
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   979
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
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   12582912
         FCOLO           =   12582912
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmformapago.frx":3664
         PICN            =   "frmformapago.frx":3680
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSDataListLib.DataCombo DtcMovimiento 
         Height          =   330
         Left            =   1800
         TabIndex        =   16
         Top             =   360
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         ForeColor       =   8388608
         Text            =   "DataCombo1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DtcCuentaCaja 
         Height          =   330
         Left            =   1800
         TabIndex        =   22
         Top             =   2760
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         ForeColor       =   8388608
         Text            =   "DataCombo1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DtcMoneda 
         Height          =   330
         Left            =   1800
         TabIndex        =   24
         Top             =   3240
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         ForeColor       =   8388608
         Text            =   "DataCombo1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DtcSucursal 
         Height          =   330
         Left            =   1800
         TabIndex        =   26
         Top             =   3600
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         ForeColor       =   8388608
         Text            =   "DataCombo1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SUCURSAL:"
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
         Left            =   720
         TabIndex        =   27
         Top             =   3720
         Width           =   750
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MONEDA :"
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
         TabIndex        =   25
         Top             =   3240
         Width           =   690
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CUENTA CAJA :"
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
         TabIndex        =   23
         Top             =   2760
         Width           =   975
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ESTADO :"
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
         Left            =   960
         TabIndex        =   21
         Top             =   4440
         Width           =   600
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "OBSERVACION :"
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
         Left            =   525
         TabIndex        =   19
         Top             =   1365
         Width           =   1035
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " MOVIMIENTO :"
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
         Left            =   540
         TabIndex        =   17
         Top             =   360
         Width           =   1020
      End
      Begin VB.Label lblcuentacontable 
         BackColor       =   &H0080C0FF&
         Caption         =   " "
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
         Height          =   435
         Left            =   1800
         TabIndex        =   14
         Top             =   2160
         Width           =   4950
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CTA CONTABLE :"
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
         Left            =   495
         TabIndex        =   12
         Top             =   1920
         Width           =   1065
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DETALLE :"
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
         Left            =   915
         TabIndex        =   8
         Top             =   840
         Width           =   645
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfdPersona 
      Height          =   6015
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   15135
      _ExtentX        =   26696
      _ExtentY        =   10610
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
   Begin VitekeySoft.ChameleonBtn cmdexit 
      Height          =   855
      Left            =   15480
      TabIndex        =   1
      Top             =   3360
      Width           =   855
      _ExtentX        =   1508
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
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   12582912
      FCOLO           =   12582912
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmformapago.frx":3A70
      PICN            =   "frmformapago.frx":3A8C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmddelete 
      Height          =   855
      Left            =   15480
      TabIndex        =   2
      Top             =   2400
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1508
      BTYPE           =   5
      TX              =   "ELIMINAR"
      ENAB            =   0   'False
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
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   12582912
      FCOLO           =   12582912
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmformapago.frx":3E7C
      PICN            =   "frmformapago.frx":3E98
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
      Left            =   15480
      TabIndex        =   3
      Top             =   1440
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1508
      BTYPE           =   5
      TX              =   "MODIFICAR"
      ENAB            =   0   'False
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
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   12582912
      FCOLO           =   12582912
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmformapago.frx":62E2
      PICN            =   "frmformapago.frx":62FE
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdnuevo 
      Height          =   855
      Left            =   15480
      TabIndex        =   4
      Top             =   480
      Width           =   855
      _ExtentX        =   1508
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
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   12582912
      FCOLO           =   12582912
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmformapago.frx":6618
      PICN            =   "frmformapago.frx":6634
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "FORMAS DE PAGO"
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
      Left            =   240
      TabIndex        =   6
      Top             =   120
      Width           =   1395
   End
   Begin VB.Label lblAcoount 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   18480
      TabIndex        =   5
      Top             =   360
      Width           =   2295
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   6780
      Left            =   0
      Top             =   0
      Width           =   16635
   End
End
Attribute VB_Name = "frmformapago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Procedencia As EnumProcede
Public Sub llenarGrid(ByVal Grilla As MSHFlexGrid)
On Error GoTo salir
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 0
    
    Exit Sub
End If
  
   Grilla.Rows = 0
   ReDim arrColWidth(1 To rst.Fields.Count)
           For Each Campo In rst.Fields
            Grilla.ColWidth(0) = 700
            Grilla.ColWidth(1) = 800
            Grilla.ColWidth(2) = 1400
            Grilla.ColWidth(3) = 2200
            Grilla.ColWidth(4) = 1200
            Grilla.ColWidth(5) = 2500
            Grilla.ColWidth(6) = 1200
            Grilla.ColWidth(7) = 2300
            Grilla.ColWidth(8) = 2300
          Next
         cabecera = "CODIGO" & vbTab & "MOVIMIENTO " & vbTab & "DESCRIPCION " & vbTab & "TIPO " & vbTab & "MONEDA " & vbTab & "OBSERVACION" & vbTab & "CTA CONTABLE" & vbTab & "NRO CUENTA" & vbTab & "CUENTA CAJA"
         Grilla.AddItem cabecera
         For k = 0 To 8
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
         Next k
        rst.MoveFirst
        
        For i = 0 To rst.RecordCount - 1
            If IsNull(rst("cuenta_caja")) = True Then
                in_caja = "[ NO CONFIGURADO  ]"
            Else
                in_caja = rst("cuenta_caja")
            End If
             Fila = rst("id_registro") & vbTab & rst("id_detalle") & vbTab & UCase(rst("forma_pago")) & vbTab & UCase(rst("descripcion")) & vbTab & UCase(rst("moneda")) & vbTab & rst("observacion") & vbTab & rst("cuenta_contable") & vbTab & rst("descripcion_cuenta") & vbTab & in_caja
             Grilla.AddItem Fila
             rst.MoveNext
        Next i

        Me.cmdupdate.Enabled = False
        Me.cmddelete.Enabled = False
  
Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub
Public Sub actualizar()
If KEY_CONTABILIDAD = "si" Then
    strCadena = "SELECT * FROM view_formas_pago WHERE  Ejercicio='" & Year(KEY_FECHA) & "' and ruc='" & KEY_RUC & "'"
Else
    strCadena = "SELECT * FROM view_formas_pago_sc WHERE   ruc='" & KEY_RUC & "'"
End If

Call llenarGrid(Me.HfdPersona)

End Sub

Private Sub cmddelete_Click()
Procedencia = Eliminar
Call disabled_form(Me)
frmsegurity.Show

Exit Sub
End Sub

Private Sub cmdexit_Click()
Unload Me
End Sub

Private Sub cmdNuevo_Click()

Me.txtid.Text = ""
Me.txtcuentacontable.Text = ""
Me.lblcuentacontable.Caption = ""
Me.DtcMovimiento.BoundText = "01"
Me.DtcDetallePago.BoundText = "01"
Me.txtObservacion.Text = ""
Me.frmdetalle.Visible = True


End Sub

Private Sub cmdprocear_Click()
Call Save
End Sub
Private Sub load(ByVal in_registro As String)

strCadena = "SELECT * FROM forma_pago_detalle WHERE id_registro='" & Val(in_registro) & "'  and ruc='" & KEY_RUC & "'"
Call ConfiguraRstT(strCadena)
If rstT.RecordCount > 0 Then
   Me.txtid.Text = rstT("id_registro")
   Me.DtcMovimiento.BoundText = rstT("id")
   Me.DtcDetallePago.BoundText = rstT("id_detalle")
   Me.txtcuentacontable.Text = rstT("cuenta_contable")
   Me.DtcCuentaCaja.BoundText = rstT("id_cuenta_caja")
   Me.DtcMoneda.BoundText = rstT("id_moneda")
   
   If rstT("estado") = "si" Then
       Me.chk_habilitado.Value = 1
   Else
       Me.chk_habilitado.Value = 0
   End If
    
   Me.lblcuentacontable.Caption = get_cuenta(Trim(Me.txtcuentacontable.Text))
   Me.txtObservacion.Text = rstT("observacion")
   Me.DtcSucursal.BoundText = rstT("id_alm")
   Me.frmdetalle.Visible = True
    
End If

End Sub
Private Sub Save()
Dim in_estado As String
If Me.chk_habilitado.Value = 1 Then
   in_estado = "si"
Else
   in_estado = "no"
End If

If Val(Me.txtid.Text) > 0 Then
   strCadena = "UPDATE forma_pago_detalle SET id_alm='" & Me.DtcSucursal.BoundText & "', id_moneda='" & Me.DtcMoneda.BoundText & "',id_cuenta_caja='" & Me.DtcCuentaCaja.BoundText & "',estado='" & in_estado & "', observacion='" & Trim(Me.txtObservacion.Text) & "',cuenta_contable='" & Trim(Me.txtcuentacontable.Text) & "',id_cuenta_caja='" & Me.DtcCuentaCaja.BoundText & "'  WHERE id_registro='" & Val(Me.txtid.Text) & "' and ruc='" & KEY_RUC & "'"
Else
   strCadena = "INSERT INTO forma_pago_detalle (`id_detalle`,`id`,`descripcion`,observacion,`estado`,`cuenta_contable`,id_cuenta_caja,id_moneda,id_alm,`ruc`)VALUES " & _
   "('" & Me.DtcDetallePago.BoundText & "','" & Me.DtcMovimiento.BoundText & "','" & Me.DtcDetallePago.Text & "','" & Trim(Me.txtObservacion.Text) & "','" & in_estado & "','" & Trim(Me.txtcuentacontable.Text) & "','" & Me.DtcCuentaCaja.BoundText & "','" & Me.DtcMoneda.BoundText & "','" & Me.DtcSucursal.BoundText & "','" & KEY_RUC & "')"
End If
CnBd.Execute (strCadena)

Me.frmdetalle.Visible = False
Call actualizar
End Sub

Private Sub cmdsalir_Click()
Me.frmdetalle.Visible = False
End Sub

Private Sub cmdupdate_Click()
Call load(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0))
End Sub

Private Sub DtcMovimiento_Change()

strCadena = "SELECT id_detalle as Codigo,descripcion as Descripcion FROM forma_pago_detalle WHERE id='" & Me.DtcMovimiento.BoundText & "' and  ruc='20479779598' ORDER BY descripcion"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcDetallePago)

End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 150

strCadena = "SELECT id as Codigo,descripcion as Descripcion FROM forma_pago  ORDER BY id"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcMovimiento)


strCadena = "SELECT  DISTINCT id_detalle as Codigo,descripcion as Descripcion FROM forma_pago_detalle WHERE id='" & Me.DtcMovimiento.BoundText & "' and  ruc='" & KEY_RUC & "' ORDER BY descripcion"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcDetallePago)

strCadena = "SELECT id_cuenta as Codigo,CONCAT(descripcion,':',numero_cuenta) as Descripcion FROM mis_cuentas WHERE ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcCuentaCaja)


strCadena = "SELECT id_moneda as Codigo, descripcion as Descripcion FROM moneda ORDER BY id_moneda"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcMoneda)

strCadena = "SELECT id_alm as Codigo,descripcion as Descripcion FROM almacen WHERE ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcSucursal)



Call actualizar

End Sub

Private Sub HfdPersona_SelChange()
If Val(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)) > 0 Then
   Me.cmdupdate.Enabled = True
   Me.cmddelete.Enabled = True
Else
    Me.cmdupdate.Enabled = False
   Me.cmddelete.Enabled = False
End If
End Sub

Private Sub txtCuentaContable_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Procedencia = Selecionar
   FrmPlanContableCuentas.Show
   Exit Sub
End If
End Sub
