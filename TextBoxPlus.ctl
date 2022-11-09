VERSION 5.00
Begin VB.UserControl TextBoxPlus 
   ClientHeight    =   855
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4515
   ScaleHeight     =   855
   ScaleWidth      =   4515
   ToolboxBitmap   =   "TextBoxPlus.ctx":0000
   Begin VB.TextBox txtGeneral 
      Height          =   690
      Left            =   15
      TabIndex        =   0
      Top             =   0
      Width           =   4455
   End
   Begin VB.Shape cdrReq 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   0  'Transparent
      Height          =   855
      Left            =   0
      Top             =   0
      Width           =   4575
   End
End
Attribute VB_Name = "TextBoxPlus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
    Option Explicit
    
    Enum DatVal
        Numeric = 1
        AlphaNumeric = 2
        AlphabetsOnly = 3
        NumericNoDecimal = 4
    End Enum
    
    Enum LetterCases
        UpperCase = 1
        LowerCase = 2
        MixedCase = 3
    End Enum
    
    Const m_def_Datatype = 1
    Const m_def_LetterCase = 1
    
    Dim enm_newDataType As DatVal
    Dim enm_newLetterCase As LetterCases
    
    'Event Declarations:
    Event Change() 'MappingInfo=txtGeneral,txtGeneral,-1,Change
    Event Click() 'MappingInfo=txtGeneral,txtGeneral,-1,Click
    Event DblClick() 'MappingInfo=txtGeneral,txtGeneral,-1,DblClick
    Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=txtGeneral,txtGeneral,-1,KeyDown
    Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=txtGeneral,txtGeneral,-1,KeyUp
    Event KeyPress(KeyAscii As Integer) 'MappingInfo=txtGeneral,txtGeneral,-1,KeyPress
    Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=txtGeneral,txtGeneral,-1,MouseDown
    Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=txtGeneral,txtGeneral,-1,MouseMove
    Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=txtGeneral,txtGeneral,-1,MouseUp
    
    'Default Property Values:
    Const m_def_Required = False
    Const m_def_DecimalPrecision = 0
    
    'Control BackColors
    Const m_def_BackColorEnabled = &H80000005
    Const m_def_BackColorDisabled = &H8000000F
    
    Dim m_BackColorEnabled  As OLE_COLOR
    Dim m_BackColorDisabled As OLE_COLOR
    
    'Color de required
    Const m_def_ReqBackColor = &HFF00&
    
    'Property Variables:
    Dim m_Required          As Boolean
    Dim m_DecimalPrecision  As Integer
    Dim m_NegativeRequired  As Boolean
    Dim m_formatrequired    As Boolean
    Dim m_DataFormat        As String

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    txtGeneral.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
    m_BackColorEnabled = PropBag.ReadProperty("BackColorEnabled", m_def_BackColorEnabled)
    m_BackColorDisabled = PropBag.ReadProperty("BackColorDisabled", m_def_BackColorDisabled)
    txtGeneral.Alignment = PropBag.ReadProperty("Alignment", 0)
    txtGeneral.Appearance = PropBag.ReadProperty("Appearance", 1)
    txtGeneral.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
    txtGeneral.Enabled = PropBag.ReadProperty("Enabled", True)
    Set txtGeneral.Font = PropBag.ReadProperty("Font", Ambient.Font)
    txtGeneral.ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)
    txtGeneral.Locked = PropBag.ReadProperty("Locked", False)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    txtGeneral.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    txtGeneral.MaxLength = PropBag.ReadProperty("MaxLength", 0)
    txtGeneral.PasswordChar = PropBag.ReadProperty("PasswordChar", "")
    txtGeneral.SelStart = PropBag.ReadProperty("SelStart", 0)
    txtGeneral.SelText = PropBag.ReadProperty("SelText", "")
    txtGeneral.ToolTipText = PropBag.ReadProperty("ToolTipText", "")
    txtGeneral.Text = PropBag.ReadProperty("Text", "")
    enm_newDataType = PropBag.ReadProperty("DataType", m_def_Datatype)
    enm_newLetterCase = PropBag.ReadProperty("LetterCase", m_def_LetterCase)
    m_DecimalPrecision = PropBag.ReadProperty("DecimalPrecision", m_def_DecimalPrecision)
    m_Required = PropBag.ReadProperty("Required", m_def_Required)
    cdrReq.BackColor = PropBag.ReadProperty("ReqBackColor", m_def_ReqBackColor)
    m_NegativeRequired = PropBag.ReadProperty("NegativeRequired", False)
    m_formatrequired = PropBag.ReadProperty("FormatRequired", False)
    m_DataFormat = PropBag.ReadProperty("DataFormat", "")
End Sub


'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BackColor", txtGeneral.BackColor, &H80000005)
    Call PropBag.WriteProperty("BackColorEnabled", m_BackColorEnabled, m_def_BackColorEnabled)
    Call PropBag.WriteProperty("BackColorDisabled", m_BackColorDisabled, m_def_BackColorDisabled)
    Call PropBag.WriteProperty("DataType", enm_newDataType, m_def_Datatype)
    Call PropBag.WriteProperty("LetterCase", enm_newLetterCase, m_def_LetterCase)
    Call PropBag.WriteProperty("Alignment", txtGeneral.Alignment, 0)
    Call PropBag.WriteProperty("Appearance", txtGeneral.Appearance, 1)
    Call PropBag.WriteProperty("BorderStyle", txtGeneral.BorderStyle, 1)
    Call PropBag.WriteProperty("Enabled", txtGeneral.Enabled, True)
    Call PropBag.WriteProperty("Font", txtGeneral.Font, Ambient.Font)
    Call PropBag.WriteProperty("ForeColor", txtGeneral.ForeColor, &H80000008)
    Call PropBag.WriteProperty("Locked", txtGeneral.Locked, False)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", txtGeneral.MousePointer, 0)
    Call PropBag.WriteProperty("MaxLength", txtGeneral.MaxLength, 0)
    Call PropBag.WriteProperty("PasswordChar", txtGeneral.PasswordChar, "")
    Call PropBag.WriteProperty("SelStart", txtGeneral.SelStart, 0)
    Call PropBag.WriteProperty("SelText", txtGeneral.SelText, "")
    Call PropBag.WriteProperty("ToolTipText", txtGeneral.ToolTipText, "")
    Call PropBag.WriteProperty("Text", txtGeneral.Text, "")
    Call PropBag.WriteProperty("DecimalPrecision", m_DecimalPrecision, m_def_DecimalPrecision)
    Call PropBag.WriteProperty("Required", m_Required, m_def_Required)
    Call PropBag.WriteProperty("ReqBackColor", cdrReq.BackColor, m_def_ReqBackColor)
    Call PropBag.WriteProperty("NegativeRequired", m_NegativeRequired, False)
    Call PropBag.WriteProperty("FormatRequired", m_formatrequired, False)
    Call PropBag.WriteProperty("DataFormat", m_DataFormat, "")
 End Sub

Private Sub txtGeneral_GotFocus()
        txtGeneral.SelStart = 0
        txtGeneral.SelLength = Len(txtGeneral.Text)
End Sub

Private Sub txtGeneral_KeyPress(KeyAscii As Integer)
    Dim intDecimal                  As Integer
    On Error Resume Next
    Dim DefStyle&
    Exit Sub
    RaiseEvent KeyPress(KeyAscii)
    If enm_newDataType = Numeric Then
        sNumericAlphaBetic KeyAscii, True, txtGeneral, CInt(m_DecimalPrecision), m_NegativeRequired, txtGeneral.SelStart
    ElseIf enm_newDataType = AlphabetsOnly Then
        If enm_newLetterCase = UpperCase Then
           DefStyle = GetWindowLong(txtGeneral.hWnd, GWL_STYLE)
           Call SetWindowLong(txtGeneral.hWnd, GWL_STYLE, (DefStyle And Not ES_NUMBER And Not ES_LOWERCASE) Or ES_UPPERCASE)
           sNumericAlphaBetic KeyAscii, False, txtGeneral
        ElseIf enm_newLetterCase = LowerCase Then
           sNumericAlphaBetic KeyAscii, False, txtGeneral
           DefStyle = GetWindowLong(txtGeneral.hWnd, GWL_STYLE)
           Call SetWindowLong(txtGeneral.hWnd, GWL_STYLE, (DefStyle And Not ES_NUMBER And Not ES_UPPERCASE) Or ES_LOWERCASE)
        End If
    ElseIf enm_newDataType = NumericNoDecimal Then
        DefStyle = GetWindowLong(txtGeneral.hWnd, GWL_STYLE)
        Call SetWindowLong(txtGeneral.hWnd, GWL_STYLE, (DefStyle And Not ES_UPPERCASE And Not ES_LOWERCASE) Or ES_NUMBER)
    Else
        If enm_newLetterCase = UpperCase Then
           DefStyle = GetWindowLong(txtGeneral.hWnd, GWL_STYLE)
           Call SetWindowLong(txtGeneral.hWnd, GWL_STYLE, (DefStyle And Not ES_NUMBER And Not ES_LOWERCASE) Or ES_UPPERCASE)
        ElseIf enm_newLetterCase = LowerCase Then
           DefStyle = GetWindowLong(txtGeneral.hWnd, GWL_STYLE)
           Call SetWindowLong(txtGeneral.hWnd, GWL_STYLE, (DefStyle And Not ES_NUMBER And Not ES_UPPERCASE) Or ES_LOWERCASE)
        End If
   End If
End Sub

Private Sub txtGeneral_LostFocus()
    On Error Resume Next
    If enm_newDataType = Numeric Then
        If m_formatrequired = True Then
             txtGeneral.Text = Format$(txtGeneral.Text, m_DataFormat)
        End If
    End If
End Sub

Private Sub txtGeneral_Change()
    If m_Required Then
        If txtGeneral.Text = vbNullString Then
            txtGeneral.Left = UserControl.ScaleLeft + 10
            txtGeneral.Top = UserControl.ScaleTop + 10
        Else
            txtGeneral.Left = UserControl.ScaleLeft
            txtGeneral.Top = UserControl.ScaleTop
        End If
    End If
    RaiseEvent Change
End Sub

Private Sub txtGeneral_Click()
    RaiseEvent Click
End Sub

Private Sub txtGeneral_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_Resize()
    With txtGeneral
        .Height = UserControl.ScaleHeight
        .Width = UserControl.ScaleWidth
        .Left = UserControl.ScaleLeft
        .Top = UserControl.ScaleTop
    End With
    
    With cdrReq
        .Left = UserControl.ScaleLeft - 1
        .Height = UserControl.ScaleHeight + 1
        .Top = UserControl.ScaleTop - 1
        .Width = UserControl.ScaleWidth + 1
    End With
End Sub

Private Sub UserControl_Show()
    With txtGeneral
        
        .Height = UserControl.ScaleHeight
        .Width = UserControl.ScaleWidth
        
        If Me.Required Or m_Required Then
            If .Text = vbNullString Then
                .Left = UserControl.ScaleLeft + 10
                .Top = UserControl.ScaleTop + 10
            Else
                .Left = UserControl.ScaleLeft
                .Top = UserControl.ScaleTop
            End If
        Else
            .Left = UserControl.ScaleLeft
            .Top = UserControl.ScaleTop
        End If
    End With
End Sub

Public Property Let DataType(enmDataType As DatVal)
    enm_newDataType = enmDataType
End Property
Public Property Get DataType() As DatVal
    DataType = enm_newDataType
End Property

Public Property Let LetterCase(enmLetterCase As LetterCases)
    enm_newLetterCase = enmLetterCase
End Property
Public Property Get LetterCase() As LetterCases
    LetterCase = enm_newLetterCase
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtGeneral,txtGeneral,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = txtGeneral.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    txtGeneral.BackColor() = New_BackColor
    
    If Me.Enabled Then
        If Me.BackColor <> Me.BackColorEnabled Then
            Me.BackColorEnabled = New_BackColor
        End If
    Else
        If Me.BackColor <> Me.BackColorDisabled Then
            Me.BackColorDisabled = New_BackColor
        End If
    End If
    
    PropertyChanged "BackColor"
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtGeneral,txtGeneral,-1,Alignment
Public Property Get Alignment() As AlignmentConstants
Attribute Alignment.VB_Description = "Returns/sets the alignment of a CheckBox or OptionButton, or a control's text."
    Alignment = txtGeneral.Alignment
End Property

Public Property Let Alignment(ByVal New_Alignment As AlignmentConstants)
    txtGeneral.Alignment() = New_Alignment
    PropertyChanged "Alignment"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtGeneral,txtGeneral,-1,Appearance
Public Property Get Appearance() As AppearanceConstants
    Appearance = txtGeneral.Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As AppearanceConstants)
    txtGeneral.Appearance() = New_Appearance
    PropertyChanged "Appearance"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtGeneral,txtGeneral,-1,BorderStyle
Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = txtGeneral.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    txtGeneral.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtGeneral,txtGeneral,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = txtGeneral.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    If Not New_Enabled Then
        txtGeneral.BackColor = m_BackColorDisabled
    Else
        txtGeneral.BackColor = m_BackColorEnabled
    End If
        
    txtGeneral.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtGeneral,txtGeneral,-1,Font
Public Property Get Font() As Font
    Set Font = txtGeneral.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set txtGeneral.Font = New_Font
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtGeneral,txtGeneral,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = txtGeneral.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    txtGeneral.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

Private Sub txtGeneral_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub
Private Sub txtGeneral_KeyUp(KeyCode As Integer, Shift As Integer)
    'MsgBox txtGeneral.Text
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtGeneral,txtGeneral,-1,Locked
Public Property Get Locked() As Boolean
Attribute Locked.VB_Description = "Determines whether a control can be edited."
    Locked = txtGeneral.Locked
End Property
Public Property Let Locked(ByVal New_Locked As Boolean)
    txtGeneral.Locked() = New_Locked
    PropertyChanged "Locked"
End Property
Private Sub txtGeneral_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseDown(Button, Shift, x, y)
    If Button = 2 Then
        If enm_newDataType = Numeric Or enm_newDataType = NumericNoDecimal Then
           If IsNumeric(Clipboard.GetText) Then
                Clipboard.Clear
           End If
        End If
    End If
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtGeneral,txtGeneral,-1,MouseIcon
Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "Sets a custom mouse icon."
    Set MouseIcon = txtGeneral.MouseIcon
End Property
Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set txtGeneral.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property
Private Sub txtGeneral_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseMove(Button, Shift, x, y)
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtGeneral,txtGeneral,-1,MousePointer
Public Property Get MousePointer() As Integer
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object."
    MousePointer = txtGeneral.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As Integer)
    txtGeneral.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

Private Sub txtGeneral_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtGeneral,txtGeneral,-1,MultiLine
Public Property Get MultiLine() As Boolean
Attribute MultiLine.VB_Description = "Returns/sets a value that determines whether a control can accept multiple lines of text."
    MultiLine = txtGeneral.MultiLine
End Property
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtGeneral,txtGeneral,-1,MaxLength
Public Property Get MaxLength() As Long
Attribute MaxLength.VB_Description = "Returns/sets the maximum number of characters that can be entered in a control."
    MaxLength = txtGeneral.MaxLength
End Property

Public Property Let MaxLength(ByVal New_MaxLength As Long)
    txtGeneral.MaxLength() = New_MaxLength
    PropertyChanged "MaxLength"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtGeneral,txtGeneral,-1,PasswordChar
Public Property Get PasswordChar() As String
Attribute PasswordChar.VB_Description = "Returns/sets a value that determines whether characters typed by a user or placeholder characters are displayed in a control."
    PasswordChar = txtGeneral.PasswordChar
End Property

Public Property Let PasswordChar(ByVal New_PasswordChar As String)
    txtGeneral.PasswordChar() = New_PasswordChar
    PropertyChanged "PasswordChar"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtGeneral,txtGeneral,-1,ScrollBars
Public Property Get ScrollBars() As Integer
Attribute ScrollBars.VB_Description = "Returns/sets a value indicating whether an object has vertical or horizontal scroll bars."
    ScrollBars = txtGeneral.ScrollBars
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtGeneral,txtGeneral,-1,SelStart
Public Property Get SelStart() As Long
Attribute SelStart.VB_Description = "Returns/sets the starting point of text selected."
    SelStart = txtGeneral.SelStart
End Property

Public Property Let SelStart(ByVal New_SelStart As Long)
    txtGeneral.SelStart() = New_SelStart
    PropertyChanged "SelStart"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtGeneral,txtGeneral,-1,SelText
Public Property Get SelText() As String
Attribute SelText.VB_Description = "Returns/sets the string containing the currently selected text."
    SelText = txtGeneral.SelText
End Property

Public Property Let SelText(ByVal New_SelText As String)
    txtGeneral.SelText() = New_SelText
    PropertyChanged "SelText"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtGeneral,txtGeneral,-1,ToolTipText
Public Property Get ToolTipText() As String
Attribute ToolTipText.VB_Description = "Returns/sets the text displayed when the mouse is paused over the control."
    ToolTipText = txtGeneral.ToolTipText
End Property

Public Property Let ToolTipText(ByVal New_ToolTipText As String)
    txtGeneral.ToolTipText() = New_ToolTipText
    PropertyChanged "ToolTipText"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtGeneral,txtGeneral,-1,Text
Public Property Get Text() As String
Attribute Text.VB_Description = "Returns/sets the text contained in the control."
    Text = txtGeneral.Text
End Property

Public Property Let Text(ByVal New_Text As String)
    txtGeneral.Text() = New_Text
    PropertyChanged "Text"
    On Error Resume Next
    If enm_newDataType = Numeric Then
        If m_formatrequired = True Then
             txtGeneral.Text = Format$(txtGeneral.Text, m_DataFormat)
        End If
    End If
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    enm_newDataType = m_def_Datatype
    enm_newLetterCase = m_def_LetterCase
    m_DecimalPrecision = m_def_DecimalPrecision
    m_Required = m_def_Required
    m_BackColorEnabled = m_def_BackColorEnabled
    m_BackColorDisabled = m_def_BackColorDisabled
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get DecimalPrecision() As Integer
    DecimalPrecision = m_DecimalPrecision
End Property

Public Property Let DecimalPrecision(ByVal New_DecimalPrecision As Integer)
    m_DecimalPrecision = New_DecimalPrecision
    PropertyChanged "DecimalPrecision"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get BackColorEnabled() As OLE_COLOR
    BackColorEnabled = m_BackColorEnabled
End Property

Public Property Let BackColorEnabled(ByVal New_BackColorEnabled As OLE_COLOR)
    m_BackColorEnabled = New_BackColorEnabled
    
    If Me.Enabled And Me.BackColor <> Me.BackColorEnabled Then
        Me.BackColor = New_BackColorEnabled
    End If
    
    PropertyChanged "BackColorEnabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get BackColorDisabled() As OLE_COLOR
    BackColorDisabled = m_BackColorDisabled
End Property

Public Property Let BackColorDisabled(ByVal New_BackColorDisabled As OLE_COLOR)
    m_BackColorDisabled = New_BackColorDisabled
    
    If Not Me.Enabled And Me.BackColor <> Me.BackColorDisabled Then
        Me.BackColor = New_BackColorDisabled
    End If
    
    PropertyChanged "BackColorDisabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get Required() As Boolean
    Required = m_Required
End Property

Public Property Let Required(ByVal New_Required As Boolean)
    m_Required = New_Required
    
    If New_Required Then
        If txtGeneral.Text = vbNullString Then
            txtGeneral.Left = UserControl.ScaleLeft + 10
            txtGeneral.Top = UserControl.ScaleTop + 10
        Else
            txtGeneral.Left = UserControl.ScaleLeft
            txtGeneral.Top = UserControl.ScaleTop
        End If
    Else
        txtGeneral.Left = UserControl.ScaleLeft
        txtGeneral.Top = UserControl.ScaleTop
    End If
    
    PropertyChanged "Required"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get ReqBackColor() As OLE_COLOR
    ReqBackColor = cdrReq.BackColor
End Property

Public Property Let ReqBackColor(ByVal New_ReqBackColor As OLE_COLOR)
    cdrReq.BackColor() = New_ReqBackColor
    PropertyChanged "ReqBackColor"
End Property

Public Property Get NegativeRequired() As Boolean
    NegativeRequired = m_NegativeRequired
End Property

Public Property Let NegativeRequired(ByVal New_NegativeRequired As Boolean)
    m_NegativeRequired = New_NegativeRequired
    PropertyChanged "NegativeRequired"
End Property

Public Property Get FormatRequired() As Boolean
    FormatRequired = m_formatrequired
End Property
Public Property Let FormatRequired(ByVal New_FormatRequired As Boolean)
    m_formatrequired = New_FormatRequired
    PropertyChanged "FormatRequired"
End Property
Public Property Get DataFormat() As String
    DataFormat = m_DataFormat
End Property
Public Property Let DataFormat(ByVal New_DataFormat As String)
    m_DataFormat = New_DataFormat
    PropertyChanged "DataFormat"
End Property
Public Sub ApplyFocus()
    txtGeneral.SetFocus
End Sub
