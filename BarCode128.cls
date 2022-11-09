VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BarCode128"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'////////////////////////////////////////////////////////////////////
'//
'// Gestionar el BarCode 128
'//
'// Diseñado por Sergio Fustero Carreras (2006)
'//
'// Crea e imprime código de barras 128 (tipos A y B)
'//     El tipo C está pendiente de implementar
'// Este código de barras se usa, por ejemplo, en las
'//     cartas certificadas, listas de correos, acuses de
'//     recibo, etc. de Correos (España)
'//
'// Puede imprimir directamente a impresora o a un
'//     contexto de dispositivo (p.e. PictureBox)
'//
'////////////////////////////////////////////////////////////////////

'variables locales para almacenar los valores de las propiedades
Private mvarTabla As colC128 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvarFuente As String 'copia local

Public Enum eTipoDeCódigo128
    cC128_A = 1
    cC128_B = 2
    cC128_C = 3
End Enum

Public Function GenerarBarras(Texto As String, ByVal in_activo As String, ByVal in_dependencia As String, ByVal in_fecha As String, DC As Object, Optional Codificación As eTipoDeCódigo128 = cC128_B, Optional x As Single = -1, Optional y As Single = -1, Optional Densidad As Single = 15, Optional alto As Single = 250, Optional ImprimirTexto As Boolean = False, Optional Fuente As String = "System", Optional Tamaño As Integer = 4) As Boolean
Dim des_almacen As String
'------------------------------
' Genera las barras (secuencias de líneas negras y espacios en blanco)
'       en código 128 del tipo A, B ó C para un contexto dado.
' El contexto puede ser un PictureBox, p.e.
'
' Parámetros:
'
'   -- Texto:
'       El texto del que se van a generar las barras
'   -- DC:
'       Contexto donde se imprimirá. Puede ser util para los que uséis vista previa.
'       Para imprimir directamente a la impresora, usa la función GenerarBarrasImpresora
'           (puede generar barras mucho más pequeñas)
'   -- Codificación:
'       Indicar cC128_A, cC128_B ó cC128_C
'   -- X:
'       Coordenada horizontal
'   -- Y:
'       Coordenada vertical
'   -- Densidad:
'       Ancho de las barras (no todos los anchos funcionan bien)
'   -- Alto:
'       Alto de las barras
'   -- ImprimirTexto:
'       Indicamos si queremos que se imprima el texto debajo de las barras o no
'   -- Fuente:
'       Fuente para el texto (si se quiere imprimir)
'   -- Tamaño:
'       Tamaño de fuente para el texto (si se quiere imprimir)
'------------------------------
Dim Aquí As Single, CBarras As String, c As Integer, D As Integer, paso, cx, Ancho
'DC.CurrentX = Densidad
' -- Establece el ancho de las líneas
On Error Resume Next
DC.DrawWidth = 1    'Algunos contextos dan error
On Error GoTo 0
Ancho = Densidad
' -- Establecer posición vertical (fija durante todo el proceso)
If y = -1 Then
    y = DC.CurrentY
End If
' -- Establecer posición horizontal inicial
If x = -1 Then
    Aquí = DC.CurrentX
Else
    Aquí = x
End If
' -- Espacio en blanco antes del código de barras
Aquí = Aquí + 10 * Densidad
' -- Conseguir el código de barras
CBarras = Barras(Texto, Codificación)
' -- Imprimir el código de barras (se intercalan líneas y espacios)

DC.CurrentX = 0
DC.Font.name = "Code128"
DC.Font.Size = 8 'Tamaño
    
DC.FontBold = True
DC.Print KEY_EMPRESA
DC.FontBold = False
y = 180
For c = 1 To Len(CBarras)
    ' -- Imprimir linea o espacio en blanco
    
    For D = 1 To Val(Mid(CBarras, c, 1))
        If (c Mod 2) = 0 Then ' 0 = Falso --> Línea
                                     ' 1 = Verdadero --> Espacio
            DC.Line (Aquí, y)-(Aquí, y + alto), vbWhite
        Else
            DC.Line (Aquí, y)-(Aquí, y + alto), vbBlack
        End If
        Aquí = Aquí + Ancho
    Next
Next
For c = 1 To 10 'Espacio en blanco después
    DC.Line (Aquí, y)-(Aquí, y + alto), vbWhite
    Aquí = Aquí + Ancho
Next
' -- Imprimir el texto, si procede
If ImprimirTexto Then
    DC.Font.name = "Code128"
    DC.Font.name = "Arial"
    DC.Font.Size = 7 'Tamaño
    DC.FontBold = False
    DC.CurrentY = y + alto
    'DC.CurrentX = Cx + (10 * Densidad) + (Aquí - Cx - DC.TextWidth(Texto)) / 2
    DC.CurrentX = 0
    DC.Print Texto & Space(10) & "UC:" & get_last_buy(in_dependencia)
    DC.CurrentX = 0
    
    'strCadena = "SELECT descripcion FROM almacen WHERE id_alm='" & in_dependencia & "' and ruc='" & KEY_RUC & "'"
    'Call ConfiguraRstT(strCadena)
    'If rstT.RecordCount > 0 Then
    '    des_almacen = rstT("descripcion")
    'Else
    '    des_almacen = " "
    'End If
    
    'DC.Print Mid(des_almacen, 1, 23)
    'DC.CurrentX = 0
    
    DC.Print Mid(in_activo, 1, 25)
    
End If
End Function


Public Function GenerarBarrasImpresora(Texto As String, Optional Codificación As eTipoDeCódigo128 = cC128_B, Optional x As Single = -1, Optional y As Single = -1, Optional Densidad As Single = 10, Optional alto As Single = 250, Optional ImprimirTexto As Boolean = False, Optional Fuente As String = "System", Optional Tamaño As Integer = 4) As Boolean
'------------------------------
' Genera las barras (secuencias de líneas negras y espacios en blanco)
'       en código 128 del tipo A, B ó C directamente en la impresora.
'
' Parámetros:
'
'   -- Texto:
'       El texto del que se van a generar las barras
'   -- Codificación:
'       Indicar cC128_A, cC128_B ó cC128_C
'   -- X:
'       Coordenada horizontal
'   -- Y:
'       Coordenada vertical
'   -- Densidad:
'       Ancho de las barras (no todos los anchos funcionan bien)
'   -- Alto:
'       Alto de las barras
'   -- ImprimirTexto:
'       Indicamos si queremos que se imprima el texto debajo de las barras o no
'   -- Fuente:
'       Fuente para el texto (si se quiere imprimir)
'   -- Tamaño:
'       Tamaño de fuente para el texto (si se quiere imprimir)
'------------------------------
Dim Aquí As Single, CBarras As String, c As Integer, D As Integer, paso, cx, Ancho
' -- Establece el ancho de las líneas
Printer.DrawWidth = 5
Ancho = Densidad
' -- Establecer posición vertical (fija durante todo el proceso)
If y = -1 Then
    y = Printer.CurrentY
End If
' -- Establecer posición horizontal inicial
If x = -1 Then
    Aquí = Printer.CurrentX
Else
    Aquí = x
End If
' -- Espacio en blanco antes del código de barras
Aquí = Aquí + 10 * Densidad
' -- Conseguir el código de barras
CBarras = Barras(Texto, Codificación)
' -- Imprimir el código de barras (se intercalan líneas y espacios)
For c = 1 To Len(CBarras)
    ' -- Imprimir linea o espacio en blanco
    For D = 1 To Val(Mid(CBarras, c, 1))
        If (c Mod 2) = 0 Then ' 0 = Falso --> Línea
                                     ' 1 = Verdadero --> Espacio
            Printer.Line (Aquí, y)-(Aquí, y + alto), vbWhite
        Else
            Printer.Line (Aquí, y)-(Aquí, y + alto), vbBlack
        End If
        Aquí = Aquí + Ancho
    Next
Next
' -- Si es necesario, pondremos espacio en blanco después
'For C = 1 To 10 'Espacio en blanco después
'    Printer.Line (Aquí, Y)-(Aquí, Y + Alto), vbWhite
'    Aquí = Aquí + Ancho
'Next
' -- Imprimir el texto, si procede
If ImprimirTexto Then
    Printer.Font.name = Fuente
    Printer.Font.Size = Tamaño
    Printer.CurrentY = y + alto
    Printer.CurrentX = cx + (10 * Densidad) + (Aquí - cx - Printer.TextWidth(Texto)) / 2
    Printer.Print Texto
End If
End Function


Public Function Barras(Texto As String, Optional Codificación As eTipoDeCódigo128 = cC128_B) As String
'------------------------------
' Crea la secuencia de barras y espacios para el texto dado
'
' Falta de implementar el código para el tipo C
'       (solamente admite números y es mucho más reducido)
'------------------------------
Dim c As Integer, letra As Integer, suma As Long
' -- Valor inicial
Select Case Codificación
    Case eTipoDeCódigo128.cC128_A
        Barras = mvarTabla(103 + 1).Patrón
        suma = 103
    Case eTipoDeCódigo128.cC128_B
        Barras = mvarTabla(104 + 1).Patrón
        suma = 104
    Case eTipoDeCódigo128.cC128_C
        Barras = mvarTabla(105 + 1).Patrón
        suma = 105
End Select
' -- Secuencia del texto
If Codificación = cC128_C Then
Else
    For c = 1 To Len(Texto)
        For letra = 1 To 63
            If mvarTabla(letra).Código = Mid(Texto, c, 1) Then
                Barras = Barras & mvarTabla(letra).Patrón
                suma = suma + mvarTabla(letra).Valor * c
                Exit For
            End If
        Next
        
        If Codificación = cC128_A Then
            For letra = 0 To 31
                If letra = Asc(Mid(Texto, c, 1)) Then
                    Barras = Barras & mvarTabla(letra).Patrón
                    suma = suma + mvarTabla(letra).Valor * c
                    Exit For
                End If
            Next
        Else
            For letra = 64 To 94
                If mvarTabla(letra).Código = Mid(Texto, c, 1) Then
                    Barras = Barras & mvarTabla(letra).Patrón
                    suma = suma + mvarTabla(letra).Valor * c
                    Exit For
                End If
            Next
        End If
        
        For letra = 95 To 102
            If Asc(Mid(Texto, c, 1)) = letra + 32 Then
                Barras = Barras & mvarTabla(letra).Patrón
                suma = suma + mvarTabla(letra).Valor * c
                Exit For
            End If
        Next
    Next
End If
' -- Código de control
Barras = Barras & mvarTabla((suma Mod 103) + 1).Patrón
' -- Final
Barras = Barras & mvarTabla(106 + 1).Patrón
End Function


Public Function PonerFuente(Objeto As Object) As Boolean
'---------------------------
' Asigna la fuente (que se debe haber indicado con anterioridad
'       al objeto pasado como parámetro, si la tiene
'---------------------------
Dim c As Integer
PonerFuente = False
On Error Resume Next
For c = 0 To Objeto.FontCount - 1
    If UCase(Trim(mvarFuente)) = UCase(Trim(Objeto.Fonts(c))) Then
        Objeto.Font.name = mvarFuente
        PonerFuente = True
        Exit For
    End If
Next
If Err.Number <> 0 Then
    PonerFuente = False
End If
On Error GoTo 0
End Function

Public Function AsignarFuente(Fuente As String) As Boolean
Dim c As Integer
'------------------------------
' Asigna la fuente para escribir el dato debajo de las barras
'       (si se quiere)
'------------------------------
AsignarFuente = False
'Recorremos todas las fuentes de pantalla
'       para saber si existe la fuente indicada
For c = 0 To Screen.FontCount - 1
    If UCase(Trim(Screen.Fonts(c))) = UCase(Trim(Fuente)) Then
        AsignarFuente = True
        mvarFuente = Fuente
        Exit For
    End If
Next
End Function

Public Property Get Fuente() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Fuente
    Fuente = mvarFuente
End Property

Public Function Texto(Frase As String, Optional Clase As eTipoDeCódigo128 = cC128_A) As String
'*******************************************************
'Devuelve el texto correspondiente para usarlo
'   con la fuente de código de barras
'
'*******************************************************
Dim suma As Long, c As Integer, V As Integer, Resto As Integer
' -- Valor de inicio, según tipo
Select Case Clase
    Case cC128_A
        suma = 203
    Case cC128_B
        suma = 204
    Case cC128_C
        suma = 205
End Select
Texto = Chr(suma)
' -- Cálculo de la suma a codificar
If Clase = cC128_C Then
    For c = 1 To Len(Frase) Step 2
        V = Val(Mid(Frase, c, 2))
        Texto = Texto & mvarTabla(V).Código
        suma = suma + V * c
    Next
Else
    For c = 1 To Len(Frase)
        Select Case Clase
            Case eTipoDeCódigo128.cC128_A
                V = Valor(Mid(Frase, c, 1))
                If V < 64 Then
                    Texto = Texto & Mid(Frase, c, 1)
                End If
            Case eTipoDeCódigo128.cC128_B
                V = Valor(Mid(Frase, c, 1))
                If V < 95 Then
                    Texto = Texto & Mid(Frase, c, 1)
                End If
            Case eTipoDeCódigo128.cC128_C
        End Select
        suma = suma + V * c
    Next
End If
' -- Dígito de control
Resto = suma Mod 103
Texto = Texto + Chr(Resto)
' -- Código final
Texto = Texto & Chr(106)
End Function

Private Function Patrón(Texto As String) As String
'---------------------------
' Devuelve el patrón para un carácter dado
'---------------------------
Dim V As C128
Patrón = ""
For Each V In mvarTabla
    If V.Código = Texto Then
        Patrón = V.Patrón
        Exit For
    End If
Next
End Function

Private Function Valor(Texto As String) As Integer
'---------------------------
' Devuelve el valor numérico correspondiente
'       a un carácter dado
'---------------------------
Dim V As C128
Valor = -1
For Each V In mvarTabla
    If V.Código = Texto Then
        Valor = V.Valor
        Exit For
    End If
Next
End Function

Public Sub InicializarTabla()
'---------------------------
' Crea la tabla con los patrones para cada símbolo
'
' El patrón indica, alternativamente, líneas y espacios
'---------------------------
Dim c As C128
Set c = mvarTabla.Add("A1")
c.Valor = 0
c.Código = " "
c.Patrón = "212222"
Set c = mvarTabla.Add("A2")
c.Valor = 1
c.Código = "!"
c.Patrón = "222122"
Set c = mvarTabla.Add("A3")
c.Valor = 2
c.Código = Chr(34)  ' Comillas dobles
c.Patrón = "222221"
Set c = mvarTabla.Add("A4")
c.Valor = 3
c.Código = "#"
c.Patrón = "121223"
Set c = mvarTabla.Add("A5")
c.Valor = 4
c.Código = "$"
c.Patrón = "121322"
Set c = mvarTabla.Add("A6")
c.Valor = 5
c.Código = "%"
c.Patrón = "131222"
Set c = mvarTabla.Add("A7")
c.Valor = 6
c.Código = "&"
c.Patrón = "122213"
Set c = mvarTabla.Add("A8")
c.Valor = 7
c.Código = "'"
c.Patrón = "122312"
Set c = mvarTabla.Add("A9")
c.Valor = 8
c.Código = "("
c.Patrón = "132212"
Set c = mvarTabla.Add("A10")
c.Valor = 9
c.Código = ")"
c.Patrón = "221213"
Set c = mvarTabla.Add("A11")
'Set C = New C128
c.Valor = 10
c.Código = "*"
c.Patrón = "221312"
Set c = mvarTabla.Add("A12")
c.Valor = 11
c.Código = "+"
c.Patrón = "231212"
Set c = mvarTabla.Add("A13")
c.Valor = 12
c.Código = ","
c.Patrón = "112232"
Set c = mvarTabla.Add("A14")
c.Valor = 13
c.Código = "-"
c.Patrón = "122132"
Set c = mvarTabla.Add("A15")
c.Valor = 14
c.Código = "."
c.Patrón = "122231"
Set c = mvarTabla.Add("A16")
c.Valor = 15
c.Código = "/"
c.Patrón = "113222"
Set c = mvarTabla.Add("A17")
c.Valor = 16
c.Código = "0"
c.Patrón = "123122"
Set c = mvarTabla.Add("A18")
c.Valor = 17
c.Código = "1"
c.Patrón = "123221"
Set c = mvarTabla.Add("A19")
c.Valor = 18
c.Código = "2"
c.Patrón = "223211"
Set c = mvarTabla.Add("A20")
c.Valor = 19
c.Código = "3"
c.Patrón = "221132"
Set c = mvarTabla.Add("A21")
'Set C = New C128
c.Valor = 20
c.Código = "4"
c.Patrón = "221231"
Set c = mvarTabla.Add("A22")
c.Valor = 21
c.Código = "5"
c.Patrón = "213212"
Set c = mvarTabla.Add("A23")
c.Valor = 22
c.Código = "6"
c.Patrón = "223112"
Set c = mvarTabla.Add("A24")
c.Valor = 23
c.Código = "7"
c.Patrón = "312131"
Set c = mvarTabla.Add("A25")
c.Valor = 24
c.Código = "8"
c.Patrón = "311222"
Set c = mvarTabla.Add("A26")
c.Valor = 25
c.Código = "9"
c.Patrón = "321122"
Set c = mvarTabla.Add("A27")
c.Valor = 26
c.Código = ":"
c.Patrón = "321221"
Set c = mvarTabla.Add("A28")
c.Valor = 27
c.Código = ";"
c.Patrón = "312212"
Set c = mvarTabla.Add("A29")
c.Valor = 28
c.Código = "<"
c.Patrón = "322112"
Set c = mvarTabla.Add("A30")
c.Valor = 29
c.Código = "="
c.Patrón = "322211"
Set c = mvarTabla.Add("A31")
'Set C = New C128
c.Valor = 30
c.Código = ">"
c.Patrón = "212123"
Set c = mvarTabla.Add("A32")
c.Valor = 31
c.Código = "?"
c.Patrón = "212321"
Set c = mvarTabla.Add("A33")
c.Valor = 32
c.Código = "@"
c.Patrón = "232121"
Set c = mvarTabla.Add("A34")
c.Valor = 33
c.Código = "A"
c.Patrón = "111323"
Set c = mvarTabla.Add("A35")
c.Valor = 34
c.Código = "B"
c.Patrón = "131123"
Set c = mvarTabla.Add("A36")
c.Valor = 35
c.Código = "C"
c.Patrón = "131321"
Set c = mvarTabla.Add("A37")
c.Valor = 36
c.Código = "D"
c.Patrón = "112313"
Set c = mvarTabla.Add("A38")
c.Valor = 37
c.Código = "E"
c.Patrón = "132113"
Set c = mvarTabla.Add("A39")
c.Valor = 38
c.Código = "F"
c.Patrón = "132311"
Set c = mvarTabla.Add("A40")
c.Valor = 39
c.Código = "G"
c.Patrón = "211313"
Set c = mvarTabla.Add("A41")
c.Valor = 40
c.Código = "H"
c.Patrón = "231113"
Set c = mvarTabla.Add("A42")
c.Valor = 41
c.Código = "I"
c.Patrón = "231311"
Set c = mvarTabla.Add("A43")
c.Valor = 42
c.Código = "J"
c.Patrón = "112133"
Set c = mvarTabla.Add("A44")
c.Valor = 43
c.Código = "K"
c.Patrón = "112331"
Set c = mvarTabla.Add("A45")
c.Valor = 44
c.Código = "L"
c.Patrón = "132131"
Set c = mvarTabla.Add("A46")
c.Valor = 45
c.Código = "M"
c.Patrón = "113123"
Set c = mvarTabla.Add("A47")
c.Valor = 46
c.Código = "N"
c.Patrón = "113321"
Set c = mvarTabla.Add("A48")
c.Valor = 47
c.Código = "O"
c.Patrón = "133121"
Set c = mvarTabla.Add("A49")
c.Valor = 48
c.Código = "P"
c.Patrón = "313121"
Set c = mvarTabla.Add("A50")
c.Valor = 49
c.Código = "Q"
c.Patrón = "211331"
Set c = mvarTabla.Add("A51")
c.Valor = 50
c.Código = "R"
c.Patrón = "231131"
Set c = mvarTabla.Add("A52")
c.Valor = 51
c.Código = "S"
c.Patrón = "213113"
Set c = mvarTabla.Add("A53")
c.Valor = 52
c.Código = "T"
c.Patrón = "213311"
Set c = mvarTabla.Add("A54")
c.Valor = 53
c.Código = "U"
c.Patrón = "213131"
Set c = mvarTabla.Add("A55")
c.Valor = 54
c.Código = "V"
c.Patrón = "311123"
Set c = mvarTabla.Add("A56")
c.Valor = 55
c.Código = "W"
c.Patrón = "311321"
Set c = mvarTabla.Add("A57")
c.Valor = 56
c.Código = "X"
c.Patrón = "331121"
Set c = mvarTabla.Add("A58")
c.Valor = 57
c.Código = "Y"
c.Patrón = "312113"
Set c = mvarTabla.Add("A59")
c.Valor = 58
c.Código = "Z"
c.Patrón = "312311"
Set c = mvarTabla.Add("A60")
c.Valor = 59
c.Código = "["
c.Patrón = "332111"
Set c = mvarTabla.Add("A61")
c.Valor = 60
c.Código = "\"
c.Patrón = "314111"
Set c = mvarTabla.Add("A62")
c.Valor = 61
c.Código = "]"
c.Patrón = "221411"
Set c = mvarTabla.Add("A63")
c.Valor = 62
c.Código = "^"
c.Patrón = "431111"
Set c = mvarTabla.Add("A64")
c.Valor = 63
c.Código = "_"
c.Patrón = "111224"
Set c = mvarTabla.Add("A65")
c.Valor = 64
c.Código = "'"
c.Patrón = "111422"
Set c = mvarTabla.Add("A66")
c.Valor = 65
c.Código = "a"
c.Patrón = "121124"
Set c = mvarTabla.Add("A67")
c.Valor = 66
c.Código = "b"
c.Patrón = "121421"
Set c = mvarTabla.Add("A68")
c.Valor = 67
c.Código = "c"
c.Patrón = "141122"
Set c = mvarTabla.Add("A69")
c.Valor = 68
c.Código = "d"
c.Patrón = "141221"
Set c = mvarTabla.Add("A70")
c.Valor = 69
c.Código = "e"
c.Patrón = "112214"
Set c = mvarTabla.Add("A71")
c.Valor = 70
c.Código = "f"
c.Patrón = "112412"
Set c = mvarTabla.Add("A72")
c.Valor = 71
c.Código = "g"
c.Patrón = "122114"
Set c = mvarTabla.Add("A73")
c.Valor = 72
c.Código = "h"
c.Patrón = "122411"
Set c = mvarTabla.Add("A74")
c.Valor = 73
c.Código = "i"
c.Patrón = "142112"
Set c = mvarTabla.Add("A75")
c.Valor = 74
c.Código = "j"
c.Patrón = "142211"
Set c = mvarTabla.Add("A76")
c.Valor = 75
c.Código = "k"
c.Patrón = "241211"
Set c = mvarTabla.Add("A77")
c.Valor = 76
c.Código = "l"
c.Patrón = "221114"
Set c = mvarTabla.Add("A78")
c.Valor = 77
c.Código = "m"
c.Patrón = "413111"
Set c = mvarTabla.Add("A79")
c.Valor = 78
c.Código = "n"
c.Patrón = "241112"
Set c = mvarTabla.Add("A80")
c.Valor = 79
c.Código = "o"
c.Patrón = "134111"
Set c = mvarTabla.Add("A81")
c.Valor = 80
c.Código = "p"
c.Patrón = "111242"
Set c = mvarTabla.Add("A82")
c.Valor = 81
c.Código = "q"
c.Patrón = "121142"
Set c = mvarTabla.Add("A83")
c.Valor = 82
c.Código = "r"
c.Patrón = "121241"
Set c = mvarTabla.Add("A84")
c.Valor = 83
c.Código = "s"
c.Patrón = "114212"
Set c = mvarTabla.Add("A85")
c.Valor = 84
c.Código = "t"
c.Patrón = "124112"
Set c = mvarTabla.Add("A86")
c.Valor = 85
c.Código = "u"
c.Patrón = "124211"
Set c = mvarTabla.Add("A87")
c.Valor = 86
c.Código = "v"
c.Patrón = "411212"
Set c = mvarTabla.Add("A88")
c.Valor = 87
c.Código = "w"
c.Patrón = "421112"
Set c = mvarTabla.Add("A89")
c.Valor = 88
c.Código = "x"
c.Patrón = "421211"
Set c = mvarTabla.Add("A90")
c.Valor = 89
c.Código = "y"
c.Patrón = "212141"
Set c = mvarTabla.Add("A91")
c.Valor = 90
c.Código = "z"
c.Patrón = "214121"
Set c = mvarTabla.Add("A92")
c.Valor = 91
c.Código = "{"
c.Patrón = "412121"
Set c = mvarTabla.Add("A93")
c.Valor = 92
c.Código = "|"
c.Patrón = "111143"
Set c = mvarTabla.Add("A94")
c.Valor = 93
c.Código = "}"
c.Patrón = "111341"
Set c = mvarTabla.Add("A95")
c.Valor = 94
c.Código = "~"
c.Patrón = "131141"
Set c = mvarTabla.Add("A96")
c.Valor = 95
c.Código = ""
c.Patrón = "114113"
Set c = mvarTabla.Add("A97")
c.Valor = 96
c.Código = ""
c.Patrón = "114311"
Set c = mvarTabla.Add("A98")
c.Valor = 97
c.Código = ""
c.Patrón = "411113"
Set c = mvarTabla.Add("A99")
c.Valor = 98
c.Código = ""
c.Patrón = "411311"
Set c = mvarTabla.Add("A100")
c.Valor = 99
c.Código = ""
c.Patrón = "113141"
Set c = mvarTabla.Add("A101")
c.Valor = 100
c.Código = ""
c.Patrón = "114131"
Set c = mvarTabla.Add("A102")
c.Valor = 101
c.Código = ""
c.Patrón = "311141"
Set c = mvarTabla.Add("A103")
c.Valor = 102
c.Código = ""
c.Patrón = "411131"
Set c = mvarTabla.Add("A104")
c.Valor = 103
c.Código = ""
c.Patrón = "211412"
Set c = mvarTabla.Add("A105")
c.Valor = 104
c.Código = ""
c.Patrón = "211214"
Set c = mvarTabla.Add("A106")
c.Valor = 105
c.Código = ""
c.Patrón = "211232"
Set c = mvarTabla.Add("A107")
c.Valor = 106
c.Código = ""
c.Patrón = "2331112"
'Set C = mvarTabla.Add("A108")
End Sub

Public Property Set Tabla(ByVal vData As colC128)
'se usa al asignar un objeto a la propiedad, en la parte izquierda de una instrucción Set.
'Syntax: Set x.Tabla = Form1
    Set mvarTabla = vData
End Property

Public Property Get Tabla() As colC128
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Tabla
    Set Tabla = mvarTabla
End Property

Private Sub Class_Initialize()
mvarFuente = "PF_C128.ttf"
Set mvarTabla = New colC128
InicializarTabla
End Sub


