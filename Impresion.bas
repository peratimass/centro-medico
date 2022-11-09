Attribute VB_Name = "Impresion"
Public Sub CargaDefConfigEpsonTM()
 Printer.Font.name = KEY_TIPO_LETRA    '"FontB11"
    Printer.Font.Size = "10"
    m_iTamLineaImpresion = Fix(Printer.ScaleWidth / Printer.TextWidth(" "))
    
End Sub
Public Sub AbreGaveta()
    Printer.Font.name = "control"
    Printer.Print "A"
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


   

  


