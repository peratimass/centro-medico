VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_QrySUNAT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'===========================================================================
Private Const URL_SBusca = "http://www.sunat.gob.pe/cl-ti-itmrconsruc/jcrS00Alias"
Private Const URL_SNumConsulta = "http://www.sunat.gob.pe/cl-ti-itmrconsruc/captcha?accion=random"
'===========================================================================

Public Event DatosObtenidos()
Public Event ErrorEnObtencion()

'===========================================================================

Private m_sRUC As String

Private m_sErrValRUC As String
Private m_sErrConSunat As String

Private m_sEmRazSocial As String
Private m_sEmNomComercial As String
Private m_sDireccion As String

Private WithEvents m_ctlWebExplorer As WebBrowser
Attribute m_ctlWebExplorer.VB_VarHelpID = -1
Private WithEvents m_ctlWebInet As Inet
Attribute m_ctlWebInet.VB_VarHelpID = -1

'===========================================================================

Public Property Get ruc() As String

    ruc = m_sRUC

End Property

Public Property Let ruc(ByVal sRUC As String)

    m_sRUC = sRUC

End Property


Public Property Get ErrValRUC() As String

    ErrValRUC = m_sErrValRUC

End Property


Public Property Get ErrConSunat() As String

    ErrConSunat = m_sErrConSunat

End Property


Public Property Get EmRazSocial() As String

    EmRazSocial = m_sEmRazSocial

End Property


Public Property Get EmDireccion() As String

    EmDireccion = m_sDireccion

End Property


Public Property Let EmRazSocial(ByVal sEmRazSocial As String)

    m_sEmRazSocial = sEmRazSocial

End Property

Public Property Get EmNomComercial() As String

    EmNomComercial = m_sEmNomComercial

End Property


Public Property Set WebExplorer(ctlWebExplorer As WebBrowser)

    Set m_ctlWebExplorer = ctlWebExplorer

End Property

Public Property Set WebInet(ctlWebInet As Inet)

    Set m_ctlWebInet = ctlWebInet

End Property


Private Sub Class_Terminate()

    Set m_ctlWebExplorer = Nothing

    Set m_ctlWebInet = Nothing
End Sub

'===========================================================================

Public Function RUC_EsValido() As Boolean

    If Not IsNumeric(m_sRUC) Then
        m_sErrValRUC = "El valor no es numérico"
    Else
        If Len(m_sRUC) <> 11 Then
            m_sErrValRUC = "Número de dígitos inválido"
        Else
        
            Dim dig01 As Integer
            Dim dig02 As Integer
            Dim dig03 As Integer
            Dim dig04 As Integer
            Dim dig05 As Integer
            Dim dig06 As Integer
            Dim dig07 As Integer
            Dim dig08 As Integer
            Dim dig09 As Integer
            Dim dig10 As Integer
            Dim dig11 As Integer
            
            Dim suma As Integer
            Dim residuo As Integer
            Dim resta As Integer
            
            Dim digChk As Integer
    
            dig01 = CInt(Mid$(m_sRUC, 1, 1)) * 5
            dig02 = CInt(Mid$(m_sRUC, 2, 1)) * 4
            dig03 = CInt(Mid$(m_sRUC, 3, 1)) * 3
            dig04 = CInt(Mid$(m_sRUC, 4, 1)) * 2
            dig05 = CInt(Mid$(m_sRUC, 5, 1)) * 7
            dig06 = CInt(Mid$(m_sRUC, 6, 1)) * 6
            dig07 = CInt(Mid$(m_sRUC, 7, 1)) * 5
            dig08 = CInt(Mid$(m_sRUC, 8, 1)) * 4
            dig09 = CInt(Mid$(m_sRUC, 9, 1)) * 3
            dig10 = CInt(Mid$(m_sRUC, 10, 1)) * 2
            dig11 = CInt(Mid$(m_sRUC, 11, 1))
            
            suma = dig01 + dig02 + dig03 + dig04 + dig05 + dig06 + dig07 + dig08 + dig09 + dig10
            residuo = suma Mod 11
            resta = 11 - residuo
            
            If resta = 11 Then
                digChk = 1
            ElseIf resta = 10 Then
                digChk = 0
            Else
                digChk = resta
            End If
            
            
            If dig11 = digChk Then
                RUC_EsValido = True
            Else
                m_sErrValRUC = "El número de RUC no es válido"
            End If
          
        End If
    End If
End Function

Public Sub CargaWebExplorador()
    
    If RUC_EsValido Then
        Call m_ctlWebInet.Execute(URL_SNumConsulta, "GET")
    Else
        Call m_ctlWebExplorer.Navigate(URL_SBusca)
    End If
End Sub

Private Sub m_ctlWebInet_StateChanged(ByVal State As Integer)
    Dim d_Chunk As Variant
    'En esta variable se almacenan los datos finales
    Dim Datos As String
    Dim RandNum As String

    If State = icResponseCompleted Then
        ' Recupera una porción de datos de 1024 bytes
        d_Chunk = m_ctlWebInet.GetChunk(1024, icString)
        Datos = Datos & d_Chunk
  
        Do
            DoEvents
            ' Recupera una porción de datos de 1024 bytes
            d_Chunk = m_ctlWebInet.GetChunk(1024, icString)
              
            ' Si ya no hay mas sale
            If Len(d_Chunk) = 0 Then
               Exit Do
            Else
              Datos = Datos & d_Chunk
            End If
        Loop

        ' Asigna el dato a la variable local
        RandNum = Datos
        
        '======================================================
        Dim Url As String
        Url = URL_SBusca & "?" & _
            "accion=consPorRuc" & _
            "&nroRuc=" & m_sRUC & _
            "&actReturn=1" & _
            "&numRnd=" & RandNum
            
        m_ctlWebExplorer.Navigate Url
        '======================================================
        
        
    ElseIf State = icError Then
        
        m_sErrConSunat = m_ctlWebInet.ResponseInfo
        
        RaiseEvent ErrorEnObtencion
    End If
        
End Sub

Private Sub m_ctlWebExplorer_DocumentComplete(ByVal pDisp As Object, Url As Variant)
  On Error GoTo Saltar
    Dim PosNomCom As Integer
    Dim PosNomCFin As Integer
    
    m_ctlWebExplorer.Document.onmousedown = vbNullString 'para permitir el click derecho
    
    m_sEmRazSocial = m_ctlWebExplorer.Document.All.forminfoHist.desRuc.Value
    m_sDireccion = m_ctlWebExplorer.Document.All.forminfoHist.desRuc.Value
    
    PosNomCom = InStr(1, m_ctlWebExplorer.Document.documentElement.innerHTML, ">Dirección del Domicilio Fiscal:") + 1
    PosNomCom = InStr(PosNomCom, m_ctlWebExplorer.Document.documentElement.innerHTML, ">") + 1
    PosNomCom = InStr(PosNomCom, m_ctlWebExplorer.Document.documentElement.innerHTML, ">") + 1
    PosNomCFin = InStr(PosNomCom, m_ctlWebExplorer.Document.documentElement.innerHTML, "</")
    
    m_sDireccion = Mid$(m_ctlWebExplorer.Document.documentElement.innerHTML, PosNomCom, PosNomCFin - PosNomCom)
    
    RaiseEvent DatosObtenidos
    Exit Sub
Saltar:
     RaiseEvent DatosObtenidos
End Sub
