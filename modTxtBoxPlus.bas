Attribute VB_Name = "modtxtBox"
Option Explicit

    Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
    Public Const GWL_STYLE = (-16)
    Public Const ES_UPPERCASE As Long = &H8&
    Public Const ES_LOWERCASE As Long = &H10&
    Public Const ES_NUMBER     As Long = &H2000

Public Sub sNumericAlphaBetic(KeyAsciis As Integer, bNumeric As Boolean, txtBoxName As TextBox, Optional intDecimal As Integer, Optional bnegativeRequired As Boolean, Optional intCurrentPos As Integer)
    
    'Keyascii used
    '45 - Minus Sign
    '8-Back Space
    '46-Decimal Point
    '22-Ctrl + V
    '3-Ctrl + C
    
    'For Numeric Only
    Dim intAfterDecimals            As Integer
    If bNumeric = True Then
        Dim intposition As Integer
        
        If KeyAsciis <> 8 And intDecimal > 0 And KeyAsciis <> 45 Then
            intposition = InStr(txtBoxName.Text, ".")   'To check decimal places
            If intposition > 0 And KeyAsciis <> 46 Then
                intAfterDecimals = Len(txtBoxName.Text) + 1 - intposition
                'if no of elemenmts after decimal is greater than intdecimals
                'then keyascii=is set to 0
                If intCurrentPos >= intposition Then
                    If intAfterDecimals > intDecimal Then
                       If txtBoxName.SelLength = 0 Then
                          KeyAsciis = 0
                       End If
                    End If
                End If
            End If
        End If
        
        If KeyAsciis = 46 Then
            intposition = InStr(txtBoxName.Text, ".")
            If intposition > 0 Then
               KeyAsciis = 0
               Exit Sub
            ElseIf Len(Mid$(txtBoxName.Text, txtBoxName.SelStart + 1, Len(txtBoxName))) > intDecimal Then
               KeyAsciis = 0
            End If
        End If
        
        If KeyAsciis = 45 Then
           If bnegativeRequired = False Then
              KeyAsciis = 0
           End If
           intposition = InStr(txtBoxName.Text, "-")
           If intposition > 0 Then
               KeyAsciis = 0
               Exit Sub
           Else
               KeyAsciis = 0
               If Len(txtBoxName.Text) < txtBoxName.MaxLength Then
                    txtBoxName.SelStart = 0
                    txtBoxName.Text = "-" & txtBoxName.Text
                    txtBoxName.SelStart = Len(txtBoxName.Text)
               End If
           End If
        End If
        
        If (KeyAsciis < 48 Or KeyAsciis > 57) _
                    And KeyAsciis <> 46 _
                    And KeyAsciis <> 8 _
                    And KeyAsciis <> 45 _
                    And KeyAsciis <> 3 _
                    And KeyAsciis <> 22 Then
            KeyAsciis = 0
        End If
        
        If KeyAsciis = 22 Then
            If Not IsNumeric(Clipboard.GetText) Then
                KeyAsciis = 0
            End If
        End If
    Else
        'Para admitir solo letras
        If (KeyAsciis >= 65 And KeyAsciis <= 90) Or _
            (KeyAsciis >= 97 And KeyAsciis <= 122) Or _
            KeyAsciis = 8 Or KeyAsciis = 32 Or _
            KeyAsciis = 164 Or KeyAsciis = 165 Then
        Else
            KeyAsciis = 0
        End If
    End If
    
End Sub

