VERSION 5.00
Begin VB.UserControl UserControl_downloadimg 
   ClientHeight    =   690
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1785
   ScaleHeight     =   690
   ScaleWidth      =   1785
End
Attribute VB_Name = "UserControl_downloadimg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' variables
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private m_Url As String
Private m_Picture As Picture

Private m_AutoSize As Boolean
Private m_BorderStyle As Boolean
' evento para el progreso de la descarga
Public Event progreso(Value As Integer)

' cuando finaliza la descarga se produce este evento.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub UserControl_AsyncReadComplete(AsyncProp As AsyncProperty)
    
On Error GoTo error_Sub
    
    ' carga la imagen
    Set UserControl.Picture = AsyncProp.Value
    ' Establece el Autosize
    If m_AutoSize Then
       Redimensionar_Imagen
    End If
Exit Sub

error_Sub:

MsgBox "Hubo un error al intentar descargar la imagen. " & vbNewLine & _
       "Compruebe la dirección url que sea correcta y la conexión a internet", vbCritical

End Sub

' evento que va lanzando el progreso a mientras se descarga la imagen
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub UserControl_AsyncReadProgress(AsyncProp As AsyncProperty)
    
    On Error Resume Next
    Dim P As Integer
    DoEvents
    P = (AsyncProp.BytesRead / AsyncProp.BytesMax) * 100
    RaiseEvent progreso(P)
    
End Sub

' Sub que ajusta la dimensión del Usercontrol al de la imagen cargada
Private Sub Redimensionar_Imagen()
    Dim AltoImg As Long
    Dim AnchoImg As Long
    If UserControl.Picture <> 0 And m_AutoSize Then
       ' Ancho y alto del picture (cambia la escala )
       AltoImg = ScaleY(UserControl.Picture.Height, vbHimetric, vbTwips)
       AnchoImg = ScaleX(UserControl.Picture.Width, vbHimetric, vbTwips)
       ' ajusta el tamaño del control a la imagen
       Width = AnchoImg
       Height = AltoImg
    End If
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''
' propiedades
'''''''''''''''''''''''''''''''''''''''''''''''''''

Property Get Url() As String
    Url = m_Url
End Property

Property Let Url(Valor As String)
    ' verifica que la dirección tenga el protocolo http por que si no da error
    If LCase(Left(Valor, 7)) <> "http://" Then
       Valor = "http://" & Valor
    End If
    m_Url = Valor
    'vbAsyncTypePicture, es el tipo de dato a descargar, en este caso una imagen
    UserControl.AsyncRead Url, vbAsyncTypePicture, "", vbAsyncReadForceUpdate
    
    PropertyChanged ("Url")
End Property

Property Get Picture() As Picture
    Set Picture = UserControl.Picture
End Property

Property Let Picture(Valor As Picture)
     Set UserControl.Picture = Valor
     PropertyChanged ("Picture")
End Property

' asigna la imagen al picturebox
Property Set Picture(Valor As Picture)
    Set UserControl.Picture = Valor
End Property


Property Get BorderStyle() As Boolean
    BorderStyle = UserControl.BorderStyle
End Property

Property Let BorderStyle(Valor As Boolean)
     If Valor Then
        UserControl.BorderStyle = 1
     Else
        UserControl.BorderStyle = 0
     End If
     PropertyChanged ("BorderStyle")
End Property


Property Get AutoSize() As Boolean
    AutoSize = m_AutoSize
End Property

Property Let AutoSize(Valor As Boolean)
    m_AutoSize = Valor
    PropertyChanged ("AutoSize")
    Redimensionar_Imagen
End Property

' lectura escritura de propiedades en el almacén
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        m_Url = .ReadProperty("Url", "")
        m_AutoSize = .ReadProperty("AutoSize", True)
        UserControl.Picture = .ReadProperty("Picture", Nothing)
        UserControl.BorderStyle = .ReadProperty("BorderStyle", 1)
    End With
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "Picture", UserControl.Picture, Nothing
        .WriteProperty "Url", m_Url, ""
        .WriteProperty "Autosize", m_AutoSize, True
        .WriteProperty "BorderStyle", UserControl.BorderStyle, 1
    End With
End Sub



