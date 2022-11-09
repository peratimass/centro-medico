VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form UFE30_DatabaseDemo 
   BorderStyle     =   0  'None
   Caption         =   "Suprema PC SDK 3.3 Database Demo (VB 6.0)"
   ClientHeight    =   6075
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   12345
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   405
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   823
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Cerrar"
      Height          =   855
      Left            =   11400
      TabIndex        =   16
      Top             =   2880
      Width           =   735
   End
   Begin VB.ComboBox cbType 
      Height          =   300
      ItemData        =   "UFE30_DatabaseDemo.frx":0000
      Left            =   3240
      List            =   "UFE30_DatabaseDemo.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   4080
      Width           =   1695
   End
   Begin VB.PictureBox pbImageFrame 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3735
      Left            =   1680
      ScaleHeight     =   247
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   223
      TabIndex        =   13
      Top             =   240
      Width           =   3375
   End
   Begin VB.CommandButton btnClear 
      Caption         =   "Clear"
      Height          =   1335
      Left            =   11340
      TabIndex        =   12
      Top             =   4380
      Width           =   795
   End
   Begin VB.CommandButton btnSelectionVerify 
      Caption         =   "Verify"
      Height          =   375
      Left            =   300
      TabIndex        =   11
      Top             =   4980
      Width           =   1035
   End
   Begin VB.CommandButton btnSelectionUpdateTemplate 
      Caption         =   "Update Template"
      Height          =   555
      Left            =   300
      TabIndex        =   10
      Top             =   4320
      Width           =   1035
   End
   Begin VB.CommandButton btnSelectionUpdateUserInfo 
      Caption         =   "Update User Info"
      Height          =   555
      Left            =   300
      TabIndex        =   9
      Top             =   3720
      Width           =   1035
   End
   Begin VB.CommandButton btnSelectionDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   300
      TabIndex        =   8
      Top             =   3300
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Caption         =   "Selection"
      Height          =   2475
      Left            =   180
      TabIndex        =   7
      Top             =   3000
      Width           =   1275
   End
   Begin VB.CommandButton btnDeleteAll 
      Caption         =   "Delete All"
      Height          =   375
      Left            =   180
      TabIndex        =   6
      Top             =   2280
      Width           =   1275
   End
   Begin MSComctlLib.ListView lvDatabaseList 
      Height          =   3795
      Left            =   5400
      TabIndex        =   5
      Top             =   180
      Width           =   5835
      _ExtentX        =   10292
      _ExtentY        =   6694
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton btnIdentify 
      Caption         =   "Identify"
      Height          =   375
      Left            =   180
      TabIndex        =   4
      Top             =   1680
      Width           =   1275
   End
   Begin VB.CommandButton btnEnroll 
      Caption         =   "Enroll"
      Height          =   375
      Left            =   180
      TabIndex        =   3
      Top             =   1260
      Width           =   1275
   End
   Begin VB.CommandButton btnUninit 
      Caption         =   "Uninit"
      Height          =   375
      Left            =   180
      TabIndex        =   2
      Top             =   600
      Width           =   1275
   End
   Begin VB.TextBox txtMessage 
      Height          =   1335
      Left            =   1620
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   4380
      Width           =   9615
   End
   Begin VB.CommandButton btnInit 
      Caption         =   "Init"
      Height          =   375
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   1275
   End
   Begin MSComDlg.CommonDialog cdFileDialog 
      Left            =   0
      Top             =   5400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   3  'Dot
      Height          =   6015
      Left            =   0
      Top             =   0
      Width           =   12300
   End
   Begin VB.Label Label1 
      Caption         =   "Template Type"
      Height          =   255
      Left            =   1680
      TabIndex        =   14
      Top             =   4080
      Width           =   1335
   End
End
Attribute VB_Name = "UFE30_DatabaseDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Activate()
'Call conectar
End Sub

'==========================================================================='
Private Sub Form_Load()
    m_hScanner = 0
    m_hDatabase = 0
    m_hMatcher = 0

    lvDatabaseList.ColumnHeaders.Add , , "Serial", 50, lvwColumnLeft
    lvDatabaseList.ColumnHeaders.Add , , "UserID", 60, lvwColumnLeft
    lvDatabaseList.ColumnHeaders.Add , , "FingerIndex", 80, lvwColumnLeft
    lvDatabaseList.ColumnHeaders.Add , , "Template1", 80, lvwColumnLeft
    lvDatabaseList.ColumnHeaders.Add , , "Template2", 80, lvwColumnLeft
    lvDatabaseList.ColumnHeaders.Add , , "Memo", 60, lvwColumnLeft
    
    cbType.ListIndex = -1
End Sub

Private Sub Form_Terminate()
    btnUninit_Click
End Sub
'==========================================================================='


'==========================================================================='
Public Sub AddMessage(ByVal Text As String)
    txtMessage.SelStart = Len(txtMessage.Text)
    txtMessage.SelText = Text
    txtMessage.Refresh
End Sub

Private Sub btnClear_Click()
    txtMessage.Text = ""
End Sub

Public Sub AddRow(ByVal Serial As Long, ByVal UserID As String, ByVal FingerIndex As Long, ByVal bTemplate1 As Boolean, ByVal bTemplate2 As Boolean, ByVal Memo As String)
    With lvDatabaseList.ListItems.Add(, , Serial)
        .ListSubItems.Add , , UserID
        .ListSubItems.Add , , FingerIndex
        If (bTemplate1) Then
            .ListSubItems.Add , , "O"
        Else
            .ListSubItems.Add , , "X"
        End If
        If (bTemplate2) Then
            .ListSubItems.Add , , "O"
        Else
            .ListSubItems.Add , , "X"
        End If
        .ListSubItems.Add , , Memo
    End With
End Sub

Public Sub UpdateDatabaseList()
    If (m_hDatabase = 0) Then
        Exit Sub
    End If

    Dim ufd_res As UFD_STATUS
    Dim DataNumber As Long
    Dim i As Long

    ufd_res = UFD_GetDataNumber(m_hDatabase, DataNumber)
    If (ufd_res = UFD_STATUS.OK) Then
        AddMessage ("UFD_GetDataNumber: " & DataNumber & vbCrLf)
    Else
        UFD_GetErrorString ufd_res, m_strError
        AddMessage ("UFD_GetDataNumber: " & m_strError & vbCrLf)
        Exit Sub
    End If
            
    lvDatabaseList.ListItems.Clear

    For i = 0 To DataNumber - 1
        ufd_res = UFD_GetDataByIndex(m_hDatabase, i, m_Serial, m_UserID, m_FingerIndex, m_Template1(0), m_Template1Size, m_Template2(0), m_Template2Size, m_Memo)
        If (ufd_res <> UFD_STATUS.OK) Then
            UFD_GetErrorString ufd_res, m_strError
            AddMessage ("UFD_GetDataByIndex: " & m_strError & vbCrLf)
            Exit Sub
        End If
        AddRow m_Serial, m_UserID, m_FingerIndex, (m_Template1Size <> 0), (m_Template2Size <> 0), m_Memo
    Next
End Sub
'==========================================================================='


'==========================================================================='
Private Sub conectar()
    '==========================================================================='
    ' Initilize scanners
    '==========================================================================='
    Dim ufs_res As UFS_STATUS
    Dim ScannerNumber As Long
     Dim ufd_res As UFD_STATUS
    
    Dim Connection As String
    Dim DataSource As String
    
    Screen.MousePointer = vbHourglass
    ufs_res = UFS_Init()
    Screen.MousePointer = vbDefault
    If (ufs_res = UFS_STATUS.OK) Then
        AddMessage ("UFS_Init: OK" & vbCrLf)
    Else
        UFS_GetErrorString ufs_res, m_strError
        AddMessage ("UFS_Init: " & m_strError & vbCrLf)
        Exit Sub
    End If
    
    ufs_res = UFS_GetScannerNumber(ScannerNumber)
    If (ufs_res = UFS_STATUS.OK) Then
        AddMessage ("UFS_GetScannerNumber: " & ScannerNumber & vbCrLf)
    Else
        UFS_GetErrorString ufs_res, m_strError
        AddMessage ("UFS_GetScannerNumber: " & m_strError & vbCrLf)
        Exit Sub
    End If
    
    If (ScannerNumber = 0) Then
        AddMessage ("There's no available scanner" & vbCrLf)
        Exit Sub
    Else
        AddMessage ("First scanner will be used" & vbCrLf)
    End If
    
    ufs_res = UFS_GetScannerHandle(0, m_hScanner)
    If (ufs_res <> UFS_STATUS.OK) Then
        UFS_GetErrorString ufs_res, m_strError
        AddMessage ("UFS_GetScannerHandle: " & m_strError & vbCrLf)
        Exit Sub
    End If
    '==========================================================================='
     
    '==========================================================================='
    ' Open database
    '==========================================================================='
   
    strRuta = App.Path & "\UFDatabase.mdb"
    'DataSource = "UFDatabase.mdb"
    '
    '---- cdFileDialog.Filter = "Database Files (*.mdb)|*.mdb"
    '----cdFileDialog.FileName = "UFDatabase.mdb"
    '----cdFileDialog.DefaultExt = "mdb"
    '----cdFileDialog.Flags = cdlOFNHideReadOnly Or cdlOFNPathMustExist Or cdlOFNOverwritePrompt Or cdlOFNNoReadOnlyReturn
    '----cdFileDialog.ShowOpen
    '----DataSource = cdFileDialog.FileName
            
    Connection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strRuta & ";"
    cnbd1.ConnectionString = Connection
    cnbd1.Open
    
    ufd_res = UFD_Open(cnbd1, vbNullString, vbNullString, m_hDatabase)
    
    If (ufd_res = UFD_STATUS.OK) Then
        AddMessage ("UFD_Open: OK" & vbCrLf)
    Else
        UFD_GetErrorString ufd_res, m_strError
        AddMessage ("UFD_Open: " & m_strError & vbCrLf)
    End If
    
    UpdateDatabaseList
    '==========================================================================='
    
    '==========================================================================='
    ' Create matcher
    '==========================================================================='
    Dim ufm_res As UFM_STATUS

    ufm_res = UFM_Create(m_hMatcher)
    If (ufm_res = UFM_STATUS.OK) Then
        AddMessage ("UFM_Create: OK" & vbCrLf)
    Else
        UFM_GetErrorString ufm_res, m_strError
        AddMessage ("UFM_Create: " & m_strError & vbCrLf)
        Exit Sub
    End If
    '==========================================================================='
    
End Sub
Public Sub conectar_molirey()
    '==========================================================================='
    ' Initilize scanners
    '==========================================================================='
    Dim ufs_res As UFS_STATUS
    Dim ScannerNumber As Long
    
    Screen.MousePointer = vbHourglass
    ufs_res = UFS_Init()
    Screen.MousePointer = vbDefault
    If (ufs_res = UFS_STATUS.OK) Then
        AddMessage ("UFS_Init: OK" & vbCrLf)
    Else
        UFS_GetErrorString ufs_res, m_strError
        AddMessage ("UFS_Init: " & m_strError & vbCrLf)
        Exit Sub
    End If
    
    ufs_res = UFS_GetScannerNumber(ScannerNumber)
    If (ufs_res = UFS_STATUS.OK) Then
        AddMessage ("UFS_GetScannerNumber: " & ScannerNumber & vbCrLf)
    Else
        UFS_GetErrorString ufs_res, m_strError
        AddMessage ("UFS_GetScannerNumber: " & m_strError & vbCrLf)
        Exit Sub
    End If
    
    If (ScannerNumber = 0) Then
        AddMessage ("There's no available scanner" & vbCrLf)
        Exit Sub
    Else
        AddMessage ("First scanner will be used" & vbCrLf)
    End If
    
    ufs_res = UFS_GetScannerHandle(0, m_hScanner)
    If (ufs_res <> UFS_STATUS.OK) Then
        UFS_GetErrorString ufs_res, m_strError
        AddMessage ("UFS_GetScannerHandle: " & m_strError & vbCrLf)
        Exit Sub
    End If
    '==========================================================================='
     
    '==========================================================================='
    ' Open database
    '==========================================================================='
    Dim ufd_res As UFD_STATUS
    
    Dim Connection As String
    Dim DataSource As String
    
    'DataSource = "UFDatabase.mdb"
    '
    '---- cdFileDialog.Filter = "Database Files (*.mdb)|*.mdb"
    '----cdFileDialog.FileName = "UFDatabase.mdb"
    '----cdFileDialog.DefaultExt = "mdb"
    '----cdFileDialog.Flags = cdlOFNHideReadOnly Or cdlOFNPathMustExist Or cdlOFNOverwritePrompt Or cdlOFNNoReadOnlyReturn
    '----cdFileDialog.ShowOpen
    '----DataSource = cdFileDialog.FileName
            
    'Connection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DataSource & ";"
    
    ufd_res = UFD_Open(CnBd, vbNullString, vbNullString, m_hDatabase)
    If (ufd_res = UFD_STATUS.OK) Then
        AddMessage ("UFD_Open: OK" & vbCrLf)
    Else
        UFD_GetErrorString ufd_res, m_strError
        AddMessage ("UFD_Open: " & m_strError & vbCrLf)
    End If
    
    UpdateDatabaseList
    '==========================================================================='
    
    '==========================================================================='
    ' Create matcher
    '==========================================================================='
    Dim ufm_res As UFM_STATUS

    ufm_res = UFM_Create(m_hMatcher)
    If (ufm_res = UFM_STATUS.OK) Then
        AddMessage ("UFM_Create: OK" & vbCrLf)
    Else
        UFM_GetErrorString ufm_res, m_strError
        AddMessage ("UFM_Create: " & m_strError & vbCrLf)
        Exit Sub
    End If
    '==========================================================================='

End Sub

Private Sub btnInit_Click()
Call conectar
End Sub
Private Sub desconectar()
'==========================================================================='
    ' Uninit scanners
    '==========================================================================='
    Dim ufs_res As UFM_STATUS
    
    Screen.MousePointer = vbHourglass
    ufs_res = UFS_Uninit()
    Screen.MousePointer = vbDefault
    If (ufs_res = UFS_STATUS.OK) Then
        AddMessage ("UFS_Uninit: OK" & vbCrLf)
    Else
        UFS_GetErrorString ufs_res, m_strError
        AddMessage ("UFS_Uninit: " & m_strError & vbCrLf)
    End If
    '==========================================================================='
    
    '==========================================================================='
    ' Close database
    '==========================================================================='
    If (m_hDatabase <> 0) Then
        Dim ufd_res As UFD_STATUS
        
        ufd_res = UFD_Close(m_hDatabase)
        If (ufd_res = UFD_STATUS.OK) Then
            AddMessage ("UFD_Close: OK" & vbCrLf)
        Else
            UFD_GetErrorString ufd_res, m_strError
            AddMessage ("UFD_Close: " & m_strError & vbCrLf)
        End If
        m_hDatabase = 0
    End If
    
    lvDatabaseList.ListItems.Clear
    '==========================================================================='
    
    '==========================================================================='
    ' Delete matcher
    '==========================================================================='
    If (m_hMatcher <> 0) Then
        Dim ufm_res As UFM_STATUS

        ufm_res = UFM_Delete(m_hMatcher)
        If (ufm_res = UFM_STATUS.OK) Then
            AddMessage ("UFM_Delete: OK" & vbCrLf)
        Else
            UFM_GetErrorString ufm_res, m_strError
            AddMessage ("UFM_Delete: " & m_strError & vbCrLf)
        End If
        m_hMatcher = 0
    End If
    '==========================================================================='
     cnbd1.Close
End Sub
Private Sub btnUninit_Click()
    Call desconectar
End Sub

Public Function ExtractTemplate(ByRef Template() As Byte, ByRef TemplateSize As Long) As Boolean
    Dim ufs_res As UFS_STATUS
    Dim EnrollQuality As Long

    UFS_ClearCaptureImageBuffer (m_hScanner)
    
    AddMessage ("Place Finger" & vbCrLf)
    
    TemplateSize = 0
    Do
        ufs_res = UFS_CaptureSingleImage(m_hScanner)
         
        If (ufs_res <> UFS_STATUS.OK) Then
            UFS_GetErrorString ufs_res, m_strError
            AddMessage ("UFS_CaptureSingleImage: " & m_strError & vbCrLf)
            ExtractTemplate = False
            Exit Function
        End If

        ufs_res = UFS_Extract(m_hScanner, Template(0), TemplateSize, EnrollQuality)
        If (ufs_res = UFS_STATUS.OK) Then
            AddMessage ("UFS_Extract: OK" & vbCrLf)
            Exit Do
        Else
            UFS_GetErrorString ufs_res, m_strError
            AddMessage ("UFS_Extract: " & m_strError & vbCrLf)
        End If
    Loop

    ExtractTemplate = True
End Function

Private Sub btnEnroll_Click()
Static indexs As Integer
indexs = indexs + 1
    If (cbType.ListIndex = 0) Then
        UFS_SetTemplateType m_hScanner, UFS_TEMPLATE_TYPE.UFS_TEMPLATE_TYPE_SUPREMA
    ElseIf (cbType.ListIndex = 1) Then
        UFS_SetTemplateType m_hScanner, UFS_TEMPLATE_TYPE.UFS_TEMPLATE_TYPE_ISO19794_2
    ElseIf (cbType.ListIndex = 2) Then
        UFS_SetTemplateType m_hScanner, UFS_TEMPLATE_TYPE.UFS_TEMPLATE_TYPE_ANSI378
    End If
    
    If (Not ExtractTemplate(m_Template1, m_Template1Size)) Then
        Exit Sub
    End If
    
    Dim ufd_res As UFD_STATUS
    
    AddMessage ("Input user data" & vbCrLf)
    'UserInfoForm.SetAdd
    'UserInfoForm.Show vbModal
    'If (Not UserInfoForm.DialogResult) Then
     '   AddMessage ("User data input is cancelled by user" & vbCrLf)
      '  Exit Sub
    'Else
        pbImageFrame.Refresh
   ' End If
    
    ufd_res = UFD_AddData(m_hDatabase, FrmPersona.HfdPersona.TextMatrix(FrmPersona.HfdPersona.Row, 0), indexs, m_Template1(0), m_Template1Size, vbNull, 0, "huella")
    
    If (ufd_res <> UFD_STATUS.OK) Then
        UFD_GetErrorString ufd_res, m_strError
        AddMessage ("UFD_AddData: " & m_strError & vbCrLf)
    Else
        cbType.Enabled = False
        cbType.Locked = True
    End If
    
    Dim codigo As Double
    Dim rst1 As New ADODB.Recordset
    Set rst1 = Nothing
    strCadena = "SELECT * FROM fingerprints ORDER BY Serial DESC "
    rst1.Open strCadena, cnbd1, adOpenKeyset, adLockOptimistic
    If rst1.RecordCount > 0 Then
        rst1.MoveFirst
        codigo = Val(rst1("Serial"))
        
    End If
    Dim X As Integer
    X = codigo
    If X > 0 Then
    Set rst1 = Nothing
    strCadena = "SELECT * FROM fingerprints WHERE Serial=" + Str(codigo)
    rst1.Open strCadena, cnbd1, adOpenKeyset, adLockOptimistic
    
    Dim rst2 As New ADODB.Recordset
    strCadena = "SELECT * FROM fingerprints"
    rst2.Open strCadena, CnBd, adOpenKeyset, adLockOptimistic
    rst2.AddNew
    rst2.Fields("UserID").Value = FrmPersona.HfdPersona.TextMatrix(FrmPersona.HfdPersona.Row, 0)
    'rst2.Fields("id_empresa").Value = KEY_RUC
    rst2.Fields("FingerIndex").Value = rst1.Fields("FingerIndex")
    rst2.Fields("Template1").Value = rst1.Fields("Template1")
    rst2.Fields("Template2").Value = rst1.Fields("Template2")
    rst2.Fields("Memo").Value = rst1.Fields("Memo")
    rst2.Update
    'UpdateDatabaseList
    Set rst2 = Nothing
    strCadena = "DELETE FROM fingerprints WHERE Serial=" + Str(codigo)
    cnbd1.Execute (strCadena)
    
    End If
End Sub

Private Sub btnIdentify_Click()
    Dim ufd_res As UFD_STATUS
    Dim ufm_res As UFM_STATUS
    ' Input finger data
    Dim Template(MAX_TEMPLATE_SIZE - 1) As Byte
    Dim TemplateSize As Long
    ' DB data
    Dim DBTemplate() As Byte
    Dim DBTemplateSize() As Long
    Dim DBSerial() As Long
    Dim DBTemplatePtr() As Long
    Dim DBTemplateNum As Long
    '
    Dim MatchIndex As Long
    
    ufd_res = UFD_GetTemplateNumber(m_hDatabase, DBTemplateNum)
    If (ufd_res = UFD_STATUS.OK) Then
        AddMessage ("UFD_GetTemplateNumber: " & DBTemplateNum & vbCrLf)
    Else
        UFD_GetErrorString ufd_res, m_strError
        AddMessage ("UFD_GetTemplateNumber: " & m_strError & vbCrLf)
    End If
    
    ReDim DBTemplate(MAX_TEMPLATE_SIZE - 1, DBTemplateNum - 1) As Byte
    ReDim DBTemplateSize(DBTemplateNum - 1) As Long
    ReDim DBSerial(DBTemplateNum - 1) As Long
    
    ' Make template pointer array to pass two dimensional template data
    ReDim DBTemplatePtr(DBTemplateNum - 1) As Long
    Dim i As Long
    For i = 0 To DBTemplateNum - 1
        DBTemplatePtr(i) = VarPtr(DBTemplate(0, i))
    Next
        
    ufd_res = UFD_GetTemplateListWithSerial(m_hDatabase, DBTemplatePtr(0), DBTemplateSize(0), DBSerial(0))
    If (ufd_res <> UFD_STATUS.OK) Then
        UFD_GetErrorString ufd_res, m_strError
        AddMessage ("UFD_GetTemplateListWithSerial: " & m_strError & vbCrLf)
        Exit Sub
    End If
    
    If (Not ExtractTemplate(Template, TemplateSize)) Then
        Exit Sub
    Else
        pbImageFrame.Refresh
    End If
    
    Screen.MousePointer = vbHourglass
    ufm_res = UFM_Identify(m_hMatcher, Template(0), TemplateSize, DBTemplatePtr(0), DBTemplateSize(0), DBTemplateNum, 5000, MatchIndex)
    Screen.MousePointer = vbDefault
    If (ufm_res <> UFM_STATUS.OK) Then
        UFM_GetErrorString ufm_res, m_strError
        AddMessage ("UFM_Identify: " & m_strError & vbCrLf)
        Exit Sub
    End If
    
    If (MatchIndex <> -1) Then
        AddMessage ("Identification succeed (Serial = " & DBSerial(MatchIndex) & ")" & vbCrLf)
    Else
        AddMessage ("Identification failed" & vbCrLf)
    End If
End Sub

Private Sub btnDeleteAll_Click()
    Dim ufd_res As UFD_STATUS

    ufd_res = UFD_RemoveAllData(m_hDatabase)
    If (ufd_res = UFD_STATUS.OK) Then
        AddMessage ("UFD_RemoveAllData: OK" & vbCrLf)
        UpdateDatabaseList
    Else
        UFD_GetErrorString ufd_res, m_strError
        AddMessage ("UFD_RemoveAllData: " & m_strError & vbCrLf)
    End If
End Sub

Private Sub btnSelectionDelete_Click()
    Dim ufd_res As UFD_STATUS
    Dim Serial As Long

    If (Not lvDatabaseList.SelectedItem.Selected) Then
        AddMessage ("Select data" & vbCrLf)
        Exit Sub
    Else
        Serial = Val(lvDatabaseList.SelectedItem.Text)
    End If

    ufd_res = UFD_RemoveDataBySerial(m_hDatabase, Serial)
    If (ufd_res = UFD_STATUS.OK) Then
        AddMessage ("UFD_RemoveDataBySerial: OK" & vbCrLf)
        UpdateDatabaseList
    Else
        UFD_GetErrorString ufd_res, m_strError
        AddMessage ("UFD_RemoveDataBySerial: " & m_strError & vbCrLf)
    End If
End Sub

Private Sub btnSelectionUpdateUserInfo_Click()
    Dim ufd_res As UFD_STATUS
    Dim Serial As Long
    
    If (Not lvDatabaseList.SelectedItem.Selected) Then
        AddMessage ("Select data" & vbCrLf)
        Exit Sub
    Else
        Serial = Val(lvDatabaseList.SelectedItem.Text)
        UserInfoForm.SetUpdate
        UserInfoForm.txtUserID = lvDatabaseList.SelectedItem.ListSubItems(DATABASE_COL_USERID).Text
        UserInfoForm.txtFingerIndex = lvDatabaseList.SelectedItem.ListSubItems(DATABASE_COL_FINGERINDEX).Text
        UserInfoForm.txtMemo = lvDatabaseList.SelectedItem.ListSubItems(DATABASE_COL_MEMO).Text
    End If
    
    AddMessage ("Update user data" & vbCrLf)
    AddMessage ("UserID and FingerIndex will not be updated" & vbCrLf)
    UserInfoForm.Show vbModal
    If (Not UserInfoForm.DialogResult) Then
        AddMessage ("User data input is cancelled by user" & vbCrLf)
        Exit Sub
    End If
    
    ufd_res = UFD_UpdateDataBySerial(m_hDatabase, Serial, vbNull, 0, vbNull, 0, UserInfoForm.txtMemo)
    If (ufd_res = UFD_STATUS.OK) Then
        AddMessage ("UFD_UpdateDataBySerial: OK" & vbCrLf)
        UpdateDatabaseList
    Else
        UFD_GetErrorString ufd_res, m_strError
        AddMessage ("UFD_UpdateDataBySerial: " & m_strError & vbCrLf)
    End If
End Sub

Private Sub btnSelectionUpdateTemplate_Click()
    Dim ufd_res As UFD_STATUS
    Dim Serial As Long
    
    If (Not lvDatabaseList.SelectedItem.Selected) Then
        AddMessage ("Select data" & vbCrLf)
        Exit Sub
    Else
        Serial = Val(lvDatabaseList.SelectedItem.Text)
    End If
    
    If (Not ExtractTemplate(m_Template1, m_Template1Size)) Then
        Exit Sub
    Else
        pbImageFrame.Refresh
    End If
    
    ufd_res = UFD_UpdateDataBySerial(m_hDatabase, Serial, m_Template1(0), m_Template1Size, vbNull, 0, vbNullString)
    If (ufd_res = UFD_STATUS.OK) Then
        AddMessage ("UFD_UpdateDataBySerial: OK" & vbCrLf)
        UpdateDatabaseList
    Else
        UFD_GetErrorString ufd_res, m_strError
        AddMessage ("UFD_UpdateDataBySerial: " & m_strError & vbCrLf)
    End If
End Sub

Private Sub btnSelectionVerify_Click()
    Dim ufd_res As UFD_STATUS
    Dim ufm_res As UFM_STATUS
    Dim Serial As Long
    ' Input finger data
    Dim Template(MAX_TEMPLATE_SIZE - 1) As Byte
    Dim TemplateSize As Long
    '
    Dim VerifySucceed As Long
    
    If (Not lvDatabaseList.SelectedItem.Selected) Then
        AddMessage ("Select data" & vbCrLf)
        Exit Sub
    Else
        Serial = Val(lvDatabaseList.SelectedItem.Text)
    End If
    
    ufd_res = UFD_GetDataBySerial(m_hDatabase, Serial, m_UserID, m_FingerIndex, m_Template1(0), m_Template1Size, m_Template2(0), m_Template2Size, m_Memo)
    If (ufd_res <> UFD_STATUS.OK) Then
        UFD_GetErrorString ufd_res, m_strError
        AddMessage ("UFD_GetDataBySerial: " & m_strError & vbCrLf)
        Exit Sub
    End If
    
    If (Not ExtractTemplate(Template, TemplateSize)) Then
        Exit Sub
    Else
        pbImageFrame.Refresh
    End If
    
    ufm_res = UFM_Verify(m_hMatcher, Template(0), TemplateSize, m_Template1(0), m_Template1Size, VerifySucceed)
    If (ufm_res <> UFM_STATUS.OK) Then
        UFM_GetErrorString ufm_res, m_strError
        AddMessage ("UFM_Verify: " & m_strError & vbCrLf)
        Exit Sub
    End If

    If (VerifySucceed = 1) Then
        AddMessage ("Verification succeed (Serial = " & Serial & ")" & vbCrLf)
    Else
        AddMessage ("Verification failed" & vbCrLf)
    End If
End Sub
'==========================================================================='
Private Sub lvDatabaseList_BeforeLabelEdit(Cancel As Integer)

End Sub


Private Sub pbImageFrame_Paint()
    Dim ufs_res As Long
    ufs_res = UFS_DrawCaptureImageBuffer(m_hScanner, pbImageFrame.hDC, pbImageFrame.ScaleLeft, pbImageFrame.ScaleTop, pbImageFrame.ScaleLeft + pbImageFrame.ScaleWidth, pbImageFrame.ScaleTop + pbImageFrame.ScaleHeight, 0)
End Sub
