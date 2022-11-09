VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_Impresion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_lIdImpresora As Long
Private m_iTamLineaImpresion As Integer
Private m_iEspaciadoGeneral As Integer


'=================================================================================
'   Entradas y salidas unitarias
'=================================================================================

Public Property Get IdImpresora() As Long

    IdImpresora = m_lIdImpresora

End Property

Public Property Let IdImpresora(ByVal lIdImpresora As Long)

    m_lIdImpresora = lIdImpresora

End Property

Public Property Get TamLineaImpresion() As Integer

    TamLineaImpresion = m_iTamLineaImpresion

End Property

Public Property Let TamLineaImpresion(ByVal iTamLineaImpresion As Integer)

    m_iTamLineaImpresion = iTamLineaImpresion

End Property

Public Property Get EspaciadoGeneral() As Integer

    EspaciadoGeneral = m_iEspaciadoGeneral

End Property

Public Property Let EspaciadoGeneral(ByVal iEspaciadoGeneral As Integer)

    m_iEspaciadoGeneral = iEspaciadoGeneral

End Property

'=================================================================================
'   Funciones
'=================================================================================
Public Function AsignaImpresoraById(Optional AutoError As Boolean = True) As Boolean
    Dim NomImpresora As String
    Dim p As VB.Printer
    
    NomImpresora = "nombre impresora" ' BDBuscarCampo("impresora", "ImpNombre", "IdImpresora", m_lIdImpresora)

   
          '  Set Printer = p
            AsignaImpresoraById = True
           
        
End Function

'=================================================================================
Public Sub CargaDefConfigEpsonTM()
    Printer.Font.name = "FontB11"
    Printer.Font.Size = "10"
    m_iTamLineaImpresion = Fix(Printer.ScaleWidth / Printer.TextWidth(" "))
    
End Sub

Public Sub AsignaFuente(Fuente As StdFont)
    Set Printer.Font = Fuente
    m_iTamLineaImpresion = Fix(Printer.ScaleWidth / Printer.TextWidth(" "))
    
End Sub

Public Sub CambiaTamFuente(ByVal TamFuente As Integer)
    Printer.Font.Size = TamFuente
    m_iTamLineaImpresion = Fix(Printer.ScaleWidth / Printer.TextWidth(" "))
    
End Sub

'=================================================================================
Public Sub PrLn( _
                Cadena As String, _
                Optional TamMax As Integer, _
                Optional SeCorta As Boolean, _
                Optional Formato As String, _
                Optional Alinea As prntAlineacion = pAlnIzquierda, _
                Optional CharRelleno As String = " ", _
                Optional EspaciadoPosterior As Integer)

    Dim TamCadena As Integer
    Dim i As Integer
    Dim NumLineas As Integer
    Dim TempCadena As String

    If m_iTamLineaImpresion = 0 Then
        Call AutoCalcularTamMaxImpresion
    End If

    If Formato <> vbNullString Then
        Cadena = Format$(Cadena, Formato)
    End If

    TamCadena = Len(Cadena)

    If TamMax > m_iTamLineaImpresion Then
        TamMax = m_iTamLineaImpresion
    ElseIf TamMax <= 0 Then
        TamMax = m_iTamLineaImpresion
    End If

    If TamCadena >= TamMax Then
        If SeCorta Then
            Cadena = Mid$(Cadena, 1, TamMax)
            Printer.Print AlineaString(Cadena, m_iTamLineaImpresion, Alinea, CharRelleno)

        Else
            NumLineas = Fix(TamCadena / TamMax)
            For i = 1 To NumLineas
                TempCadena = Mid$(Cadena, 1, TamMax)
                Printer.Print AlineaString(TempCadena, m_iTamLineaImpresion, Alinea, CharRelleno)
                Cadena = Mid$(Cadena, TamMax + 1)
            Next i
            TamCadena = Len(Cadena)
            If TamCadena > 0 Then
                Printer.Print AlineaString(Cadena, m_iTamLineaImpresion, Alinea, CharRelleno)
            End If

        End If

    Else
        Printer.Print AlineaString(Cadena, m_iTamLineaImpresion, Alinea, CharRelleno)

    End If
    
    If EspaciadoPosterior <> 0 Then
        Printer.CurrentY = Printer.CurrentY + EspaciadoPosterior
    ElseIf EspaciadoGeneral > 0 Then
        Printer.CurrentY = Printer.CurrentY + EspaciadoGeneral
    End If

End Sub



Private Sub AutoCalcularTamMaxImpresion()
    m_iTamLineaImpresion = Fix(Printer.ScaleWidth / Printer.TextWidth("x"))
End Sub


'=================================================================================
'   Inicio / Fin de la Clase
'=================================================================================
