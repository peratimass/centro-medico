VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form FrmLoad_web_service 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2655
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5295
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "FrmLoad_web_service.frx":0000
   ScaleHeight     =   2655
   ScaleWidth      =   5295
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   480
      Top             =   120
   End
   Begin VB.PictureBox CB 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1515
      Left            =   0
      ScaleHeight     =   1485
      ScaleWidth      =   4395
      TabIndex        =   0
      Top             =   3240
      Visible         =   0   'False
      Width           =   4425
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   375
      Left            =   3840
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   255
      ExtentX         =   450
      ExtentY         =   661
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0C000&
      Height          =   2650
      Left            =   0
      Top             =   0
      Width           =   5295
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   4920
      Picture         =   "FrmLoad_web_service.frx":8A68
      Top             =   120
      Width           =   240
   End
   Begin VB.Label lblloading 
      BackColor       =   &H00FF0000&
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
      Left            =   495
      TabIndex        =   2
      Top             =   1005
      Width           =   15
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   360
      TabIndex        =   3
      Top             =   960
      Width           =   4455
   End
   Begin VB.Image img_ecuador 
      Height          =   2685
      Left            =   0
      Picture         =   "FrmLoad_web_service.frx":B90C
      Top             =   0
      Visible         =   0   'False
      Width           =   5475
   End
End
Attribute VB_Name = "FrmLoad_web_service"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public FormPadre As Form
Public nom_prcedimiento As String
Public Procedencia As EnumProcede



Public Sub crear_peticion(ByVal in_url As String, ByVal in_metodo As String, ByVal in_json As String, Optional in_headers As String)
Dim Headers As String
Dim PostData() As Byte

Dim p As Object
   
Set p = JSON.parse("{method: '" & in_metodo & "', url: '" & in_url & "', json: true , body: " & in_json & ", headers: " & in_headers & " }")


PostData = JSON.toString(p)

PostData = StrConv(PostData, vbFromUnicode)

Headers = "Content-Type: application/json" & vbCrLf

If KEY_SERVIDOR_CLOUD = "si" Then
    WebBrowser1.Navigate2 "http://api.vitekey.net:3001/intranet/utils/api_console", 0, "", PostData, Headers
Else
    WebBrowser1.Navigate2 "http://192.168.1.241:3030/api_console", 0, "", PostData, Headers
End If





End Sub
Public Sub crear_json_facturacion_electronica(ByVal in_url As String, ByVal in_metodo As String, ByVal in_json As String, Optional in_headers As String)
Dim Headers As String
Dim PostData() As Byte
Dim p As Object
Set p = JSON.parse("{method: '" & in_metodo & "', url: '" & in_url & "', json: true , body: " & in_json & ", headers: " & in_headers & " }")
PostData = JSON.toString(p)
PostData = StrConv(PostData, vbFromUnicode)
Headers = "Content-Type: application/json" & vbCrLf


If KEY_SERVIDOR_CLOUD = "si" Then
   WebBrowser1.Navigate2 "http://api.vitekey.net:3001/intranet/utils/api_console", 0, "", PostData, Headers
Else
    WebBrowser1.Navigate2 "http://192.168.1.241:3030/api/utiles/api_console", 0, "", PostData, Headers
End If






End Sub

Public Sub crear_producto_keyfacil(ByVal in_url As String, ByVal in_metodo As String, ByVal in_json As String, Optional in_headers As String)
Dim Headers As String
Dim PostData() As Byte
Dim p As Object

Set p = JSON.parse("{method: '" & in_metodo & "', url: '" & in_url & "', json: true , body: " & in_json & ", headers: " & in_headers & " }")
PostData = JSON.toString(p)
PostData = StrConv(PostData, vbFromUnicode)
Headers = "Content-Type: application/json" & vbCrLf

WebBrowser1.Navigate2 "http://api.vitekey.net:3001/intranet/utils/api_console", 0, "", PostData, Headers



End Sub

Public Sub get_producto_keyfacil(ByVal in_url As String, ByVal in_metodo As String, ByVal in_json As String, Optional in_headers As String)
Dim Headers As String
Dim PostData() As Byte
Dim p As Object

Set p = JSON.parse("{method: '" & in_metodo & "', url: '" & in_url & "', json: true , body: " & in_json & ", headers: " & in_headers & " }")
PostData = JSON.toString(p)
PostData = StrConv(PostData, vbFromUnicode)
Headers = "Content-Type: application/json" & vbCrLf

WebBrowser1.Navigate2 "http://api.vitekey.net:3001/intranet/utils/api_console", 0, "", PostData, Headers



End Sub

Private Sub cmdCancelar_Click()
'Call cerrar



End Sub




Public Sub cerrar()
If FrmImpTramite.Procedencia = mailenviar Then
   FrmImpTramite.Procedencia = Neutro
   Unload Me
   Call enabled_form(FrmImpTramite)
   Exit Sub
End If
         
If FrmTransferencias.Procedencia = mailenviar Then
   FrmTransferencias.Procedencia = Neutro
   Unload Me
   Call enabled_form(FrmTransferencias)
   Exit Sub
End If
         
         
Unload Me
Call enabled_form(FrmVentas)
Exit Sub
End Sub




Private Sub cmdcerrar_Click()

End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 2500
If KEY_PAIS = KEY_PERU Then
   Me.img_ecuador.Visible = False
Else
   Me.img_ecuador.Visible = True
End If

End Sub

Private Sub timer_rh_Timer()

End Sub

Private Sub Image1_Click()
Call cerrar
End Sub

Private Sub Timer1_Timer()
If Me.lblloading.Width < 4215 Then
   Me.lblloading.Width = Me.lblloading.Width + 5
Else
    Me.lblloading.Width = 0
End If
End Sub
Private Sub WebBrowser1_DownloadBegin()
Me.Caption = "Navegador Web: " & WebBrowser1.LocationName
App.Title = "Navegador Web: " & WebBrowser1.LocationName
'Label2.Caption = "Cargando Página..."
End Sub

Private Sub WebBrowser1_DownloadComplete()
Me.Caption = "Navegador Web: " & WebBrowser1.LocationName
App.Title = "Navegador Web: " & WebBrowser1.LocationName
'Label2.Caption = "Listo"

'strHtml = WebBrowser1.Document.body.innerText

'Set p = JSON.parse(strHtml)
'MsgBox p.Item("method")

'x = WebBrowser1.Document.documentElement.innerHTML
End Sub


Private Sub WebBrowser1_NavigateComplete2(ByVal pDisp As Object, Url As Variant)
'Mostramos la url que se está cargando en el combo
'Agregamos la url al combo

'Mostramos en el la barra de titulo del formulario el title _
 de la página con la propiedad LocationName
Dim in_error_result  As Boolean
Dim in_error As Boolean
Dim in_status As Integer
Dim in_error_lote As String
Dim Y() As Byte
Dim l As String
On Error GoTo salir


Dim strHtml As String
Dim json_r As Object
strHtml = WebBrowser1.Document.body.innerText
Set json_r = JSON.parse(strHtml)
If Len(Trim(strHtml)) = 0 Then
 GoTo salir
    Exit Sub
End If
in_status = json_r.Item("status")

'l = json_r.Item("response").Item("data").Item("pdf147")
'l = Replace(l, "data:image/jpg;base64,", "")
'-------------------

'Dim waka As String


'Convertimos a un arreglo de bytes
'Dim imgBytes() As Byte
'Dim Offset As Long
'Dim Size As Long
'Dim wiki As StdPicture

'imgBytes() = Base64Decode(l)

'Offset = LBound(imgBytes)
'Size = UBound(imgBytes) - LBound(imgBytes) + 1

'Set Me.CB.Picture = ArrayToPicture(imgBytes, Offset, Size)
'Me.CB.Visible = True
'-------------------
'Exit Sub

If in_status = 400 Then



   in_error_mensaje = json_r.Item("response").Item("message")
   If in_error_mensaje = "El documento ya ha sido enviado anteriormente." Then
      Call CallByName(FormPadre, "procesar_firma_electronica_reenvio", VbMethod, strHtml)
      Unload Me
      Exit Sub
   End If
   If in_error_mensaje = "El documento buscado no existe o ya ha sido enviado a sunat" Then
        Call CallByName(FormPadre, "eliminar_firma_electronica", VbMethod, strHtml)
        Unload Me
        Exit Sub
   End If
'If in_error_result = "true" Then
    'MsgBox "Ocurrio un Error:" + Chr(13) + Chr(13) + "[1] Contraseña Incorrecta." + Chr(13) + "[2] Conexion a Internet."
    MsgBox in_error_mensaje, vbInformation
    GoTo salir
End If

Call CallByName(FormPadre, nom_prcedimiento, VbMethod, strHtml)
Unload Me
Exit Sub

salir:
On Error GoTo nn
'MsgBox "Ocurrio un Error:" + "FALLA CONECTIVIDAD INTERNET", vbInformation
in_error_mensaje = strHtml
MsgBox in_error_mensaje


nn:

Me.Timer1.Enabled = False


End Sub







