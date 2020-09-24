VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form MainFrm 
   Caption         =   "RegEdit - Control^Zed"
   ClientHeight    =   4404
   ClientLeft      =   48
   ClientTop       =   504
   ClientWidth     =   6348
   LinkTopic       =   "Form1"
   ScaleHeight     =   4404
   ScaleWidth      =   6348
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2124
      Top             =   3708
      _ExtentX        =   804
      _ExtentY        =   804
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainFrm.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainFrm.frx":0454
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainFrm.frx":08A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainFrm.frx":0BFC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Spliter 
      BorderStyle     =   0  'None
      Height          =   3828
      Left            =   2268
      MousePointer    =   9  'Size W E
      ScaleHeight     =   3828
      ScaleWidth      =   120
      TabIndex        =   3
      Top             =   0
      Width           =   120
   End
   Begin MSComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   228
      Left            =   0
      TabIndex        =   2
      Top             =   4176
      Width           =   6348
      _ExtentX        =   11197
      _ExtentY        =   402
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10859
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView Vals 
      Height          =   3792
      Left            =   2340
      TabIndex        =   1
      Top             =   0
      Width           =   3828
      _ExtentX        =   6752
      _ExtentY        =   6689
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Value"
         Object.Width           =   6068
      EndProperty
   End
   Begin MSComctlLib.TreeView Tree 
      Height          =   3792
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2280
      _ExtentX        =   4022
      _ExtentY        =   6689
      _Version        =   393217
      Indentation     =   178
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "ImageList1"
      Appearance      =   1
   End
   Begin VB.Menu MnuAbout 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "MainFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Drag As Boolean
Private Sub LoadRootKeys()
Dim X As Node
    Set X = Tree.Nodes.Add(, , , "HKEY_CLASSES_ROOT", 1, 2)
        AddLoadNode X.Index
    Set X = Tree.Nodes.Add(, , , "HKEY_CURRENT_USER", 1, 2)
        AddLoadNode X.Index
    Set X = Tree.Nodes.Add(, , , "HKEY_LOCAL_MACHINE", 1, 2)
        AddLoadNode X.Index
    Set X = Tree.Nodes.Add(, , , "HKEY_USERS", 1, 2)
        AddLoadNode X.Index
    Set X = Tree.Nodes.Add(, , , "HKEY_CURRENT_CONFIG", 1, 2)
        AddLoadNode X.Index
    Set X = Tree.Nodes.Add(, , , "HKEY_DYN_DATA", 1, 2)
        AddLoadNode X.Index
    Set X = Tree.Nodes.Add(, , , "HKEY_PERFORMANCE_DATA", 1, 2)
        AddLoadNode X.Index
End Sub

Private Sub Form_Load()
    Spliter.Left = 2600
    Form_Resize
    LoadRootKeys
End Sub
Private Function GetHKeyVal(Str As String) As Long
Select Case Str
    Case "HKEY_CLASSES_ROOT"
        GetHKeyVal = HKEY_CLASSES_ROOT
    Case "HKEY_CURRENT_CONFIG"
        GetHKeyVal = HKEY_CURRENT_CONFIG
    Case "HKEY_CURRENT_USER"
        GetHKeyVal = HKEY_CURRENT_USER
    Case "HKEY_DYN_DATA"
        GetHKeyVal = HKEY_DYN_DATA
    Case "HKEY_LOCAL_MACHINE"
        GetHKeyVal = HKEY_LOCAL_MACHINE
    Case "HKEY_PERFORMANCE_DATA"
        GetHKeyVal = HKEY_PERFORMANCE_DATA
    Case "HKEY_USERS"
        GetHKeyVal = HKEY_USERS
End Select
End Function

Private Sub Form_Resize()
On Error Resume Next
    Tree.Move Me.ScaleLeft, Me.ScaleTop, Spliter.Left - Me.ScaleLeft, Me.ScaleHeight - Status.Height
    Spliter.Move Spliter.Left, Me.ScaleTop, 50, Me.ScaleHeight - Status.Height
    Vals.Move Spliter.Left + Spliter.Width, Me.ScaleTop, Me.ScaleWidth - (Spliter.Left + Spliter.Width), Me.ScaleHeight - Status.Height
End Sub

Private Sub MnuAbout_Click()
    FrmAbout.Show vbModal
End Sub

Private Sub Spliter_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Spliter.BackColor = &H808080
    Drag = True
End Sub

Private Sub Spliter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Drag Then _
    Spliter.Left = Spliter.Left + X
End Sub

Private Sub Spliter_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Drag = False
    Spliter.BackColor = &H8000000F
    Form_Resize
End Sub

Private Sub Tree_Expand(ByVal Node As MSComctlLib.Node)
Dim RequesyKey As Long
Dim RequesySubKey As String
Dim Sep As Integer
Dim TempStr As String
Dim RLen As Long
Dim Result As Long
Dim STime As FILETIME
Dim MaxValLen As Long
Dim MaxValNameLen As Long
Dim MaxSubKeyLen As Long
Dim SubKeys As Long
Dim Values As Long
Dim hKey As Long
Dim I As Integer
Dim X As Node
    If Node.Children <> 1 Then Exit Sub
    If Node.Child.Tag <> "L" Then Exit Sub
Sep = IIf(InStr(1, Node.FullPath, "\") > 0, InStr(1, Node.FullPath, "\"), Len(Node.FullPath) + 1)
RequesyKey = GetHKeyVal(Left(Node.FullPath, Sep - 1))
RequesySubKey = Mid(Node.FullPath, Sep + 1)
    Result = RegOpenKeyEx(RequesyKey, RequesySubKey, 0, KEY_QUERY_VALUE, hKey)
    Result = RegQueryInfoKey(hKey, 0, 0, 0, SubKeys, MaxSubKeyLen, 0, Values, MaxValNameLen, MaxValLen, 0, STime)
        For I = 0 To SubKeys - 1
            TempStr = Space(MaxSubKeyLen)
            RLen = LenB(TempStr)
            Result = RegEnumKeyEx(hKey, I, TempStr, RLen, 0, 0, 0, STime)
            Set X = Tree.Nodes.Add(Node.Index, tvwChild, , Left(TempStr, RLen), 1, 2)
            AddLoadNode X.Index
        Next
    Result = RegCloseKey(hKey)
Tree.Nodes.Remove Node.Child.Index
    Node.Sorted = True
End Sub

Private Sub Tree_NodeClick(ByVal Node As MSComctlLib.Node)
Dim RequesyKey As Long
Dim RequesySubKey As String
Dim Sep As Integer
Dim TempStr As String
Dim KData As ByteArray
Dim DataLen As Long
Dim RLen As Long
Dim Result As Long
Dim STime As FILETIME
Dim MaxValLen As Long
Dim MaxValNameLen As Long
Dim MaxSubKeyLen As Long
Dim SubKeys As Long
Dim Values As Long
Dim hKey As Long
Dim I As Integer
Dim Ftype As Long
Status.Panels(1).Text = Node.FullPath
Vals.ListItems.Clear
Vals.Sorted = False
Sep = IIf(InStr(1, Node.FullPath, "\") > 0, InStr(1, Node.FullPath, "\"), Len(Node.FullPath) + 1)
RequesyKey = GetHKeyVal(Left(Node.FullPath, Sep - 1))
RequesySubKey = Mid(Node.FullPath, Sep + 1)
    Result = RegOpenKeyEx(RequesyKey, RequesySubKey, 0, KEY_QUERY_VALUE, hKey)
    Result = RegQueryInfoKey(hKey, 0, 0, 0, SubKeys, MaxSubKeyLen, 0, Values, MaxValNameLen, MaxValLen, 0, STime)
        For I = 0 To Values - 1
            TempStr = Space(MaxValNameLen) + vbNullChar
            RLen = Len(TempStr)
            KData.FirstByte = 0
            DataLen = 101 ' MaxValLen
            Ftype = 0
            'if it's return an error you should enlarge the size of the array in the ByteArray type.
            'you can't use redim statement here becouse it's bytes won't be continues
            Result = RegEnumValue(hKey, I, TempStr, RLen, 0, Ftype, KData.FirstByte, DataLen)
            Vals.ListItems.Add , , IIf(RLen > 0, Left(TempStr, RLen), "(Default)"), , IIf(Ftype = REG_SZ Or Ftype = REG_EXPAND_SZ, 4, 3)
            Vals.ListItems(Vals.ListItems.Count).SubItems(1) = DecodeData(Ftype, KData, DataLen)
        Next
    Result = RegCloseKey(hKey)
    Vals.SortKey = 0
    Vals.SortOrder = lvwAscending
    Vals.Sorted = True
End Sub
Private Sub AddLoadNode(NodeIndex As Integer)
Dim X As Node
    Set X = Tree.Nodes.Add(NodeIndex, tvwChild, , "Loading...")
    X.Tag = "L"
End Sub
Private Function DecodeData(DataType, RegValue As ByteArray, BufferLength2) As String
Dim DummyString2 As String
    Select Case DataType
      'String data.
      Case REG_SZ, REG_EXPAND_SZ:
        DummyString2 = """" & IIf(RegValue.FirstByte <> 0, Chr$(RegValue.FirstByte), "")
        For I = 0 To BufferLength2 - 3
          DummyString2 = DummyString2 & Chr$(RegValue.ByteBuffer(I))
        Next I
        DummyString2 = DummyString2 & """"
      'DWORD data.
      Case REG_DWORD:
        DummyString2 = "0x" & Hex(RegValue.ByteBuffer(2)) _
                & Hex(RegValue.ByteBuffer(1)) _
                & Hex(RegValue.ByteBuffer(0)) _
                & Hex(RegValue.FirstByte)
                
      'Binary data.
      Case REG_BINARY:
        DummyString2 = IIf(Len(Trim(Hex(RegValue.FirstByte))) = 1, "0", "") & Trim(Hex(RegValue.FirstByte))
        For I = 0 To IIf((BufferLength2 - 2 > 100), 100, BufferLength2 - 2)
          DummyString2 = DummyString2 & " " & IIf(Len(Trim(Hex(RegValue.ByteBuffer(I)))) = 1, "0", "") & Trim(Hex(RegValue.ByteBuffer(I)))
        Next I
      Case Else:
        DummyString2 = "Unsupported data type."
    End Select
DecodeData = DummyString2
End Function
