VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmFastReferences 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Fast Reference Manager"
   ClientHeight    =   3660
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6945
   Icon            =   "FastReferences.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   6945
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdFixReferences 
      Caption         =   "&Fix It..."
      Height          =   375
      Left            =   2100
      TabIndex        =   0
      Top             =   3180
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.Timer tmrMakeFormTop 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   780
      Top             =   3180
   End
   Begin VB.CommandButton cmdRemoveReferences 
      Caption         =   "&Remove"
      Height          =   375
      Left            =   4635
      TabIndex        =   2
      Top             =   3180
      Width           =   1155
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Left            =   30
      TabIndex        =   6
      Top             =   2940
      Width           =   6885
   End
   Begin VB.CommandButton cmdLeaveIt 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   375
      Left            =   5895
      TabIndex        =   3
      Top             =   3180
      Width           =   1005
   End
   Begin VB.CommandButton cmdAddReference 
      Caption         =   "&Add..."
      Height          =   375
      Left            =   3345
      TabIndex        =   1
      Top             =   3180
      Width           =   1185
   End
   Begin ComctlLib.ListView lvwReferences 
      Height          =   2625
      Left            =   0
      TabIndex        =   4
      Top             =   300
      Width           =   6945
      _ExtentX        =   12250
      _ExtentY        =   4630
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Description"
         Object.Width           =   6068
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "DLL"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Version"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Path"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   4
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "GUID"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComDlg.CommonDialog cdlFileOpen 
      Left            =   0
      Top             =   3150
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   ".dll"
      DialogTitle     =   "Locate File To Add Reference"
      Filter          =   "ActiveX DLL (*.DLL)|*.dll|Type Library (*.tlb)|*.tlb|All Files|*.*"
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Add or select the references you want to remove:"
      Height          =   195
      Left            =   60
      TabIndex        =   5
      Top             =   60
      Width           =   3480
   End
End
Attribute VB_Name = "frmFastReferences"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public VBInstance As VBIDE.VBE

Private moVBProject As VBIDE.VBProject
Private mlMissingReferencesCount As Long

Private menModalResult As enmModalResult
Private Enum enmModalResult
    emrResponceWaited = 0
    emrClosePressed = 1
    emrFixPressed = 2
End Enum

Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long

Public Function DisplayForm(ByVal voVBProject As VBIDE.VBProject) As Boolean
        
    On Error GoTo ERR_DisplayForm
    
    Screen.MousePointer = vbHourglass
    
    Set moVBProject = voVBProject
    
    Call InitListView(moVBProject)
    
    
    mlMissingReferencesCount = GetMissingReferenceCount(voVBProject)
    If mlMissingReferencesCount = 0 Then
        cmdFixReferences.Visible = False
    Else
        cmdFixReferences.Visible = True
    End If
    
    
    tmrMakeFormTop.Enabled = True
    
    menModalResult = emrResponceWaited
    
    Screen.MousePointer = vbDefault
    
    Me.Show vbModal
    Select Case menModalResult
        Case emrClosePressed
            Unload Me
        Case emrFixPressed
            Unload Me
            Call frmAddIn.ShowFixReferencesDialog(voVBProject, True)
    End Select

Exit Function
ERR_DisplayForm:
    tmrMakeFormTop.Enabled = False
    Set moVBProject = Nothing
    ShowError
End Function

Private Sub InitListView(ByVal voVBProject As VBProject)
    lvwReferences.ListItems.Clear
    
    Dim lsiListItem As ListItem
    Dim oReference As Reference
    Dim sMainListName As String
    Dim sFullPath As String
    Dim lReferenceIndex As Long
    
    For lReferenceIndex = 1 To voVBProject.References.Count
        Set oReference = voVBProject.References.Item(lReferenceIndex)
        
        If Not IsReferenceBroken(oReference) Then
            sMainListName = Trim(oReference.Description)
            If sMainListName = vbNullString Then
                Dim lDotPos As Long
                sMainListName = ExtractFileName(oReference.FullPath)
                lDotPos = InStr(1, sMainListName, ".")
                If lDotPos < Len(sMainListName) Then
                    sMainListName = Mid(sMainListName, 1, lDotPos - 1)
                End If
            End If
            sFullPath = oReference.FullPath
        Else
            sMainListName = "<Missing>"
            sFullPath = "<Not Available>"
        End If
        Set lsiListItem = SafeAddListItem(lvwReferences, oReference.GUID, sMainListName)
        If Not (lsiListItem Is Nothing) Then
            lsiListItem.SubItems(1) = ExtractFileName(sFullPath)
            lsiListItem.SubItems(2) = oReference.Major & "." & oReference.Minor
            lsiListItem.SubItems(3) = sFullPath
            lsiListItem.SubItems(4) = oReference.GUID
        Else
            Call SafeAddListItem(lvwReferences, "", "<There was an error - try using Fix button>")
        End If
    Next lReferenceIndex
End Sub

Private Sub cmdAddReference_Click()
    On Error GoTo ErrorTrap
    
    Dim bLastReadOnlyStatus As Boolean
    bLastReadOnlyStatus = False
    
    cdlFileOpen.Flags = cdlOFNFileMustExist Or cdlOFNPathMustExist Or cdlOFNHideReadOnly ' Or cdlOFNAllowMultiselect ' - This option causes long file name to become short file name!
    cdlFileOpen.fileName = GetSetting(App.EXEName, msREG_SECTION_LAST_VALUES, msREG_KEY_DLL_PATH, "C:\Lamda5\DLLs")
    cdlFileOpen.ShowOpen
    
    bLastReadOnlyStatus = HandleReadOnlyAttrib(moVBProject, True)
    Call moVBProject.References.AddFromFile(cdlFileOpen.fileName)
    Call HandleReadOnlyAttrib(moVBProject, False, bLastReadOnlyStatus)
    Call InitListView(moVBProject)

Exit Sub
ErrorTrap:
    If Err.Number <> cdlCancel Then
        Call SaveErrorObj
        Call HandleReadOnlyAttrib(moVBProject, False, bLastReadOnlyStatus)
        Call RestoreErrorObj
        ShowError
    End If
End Sub

Private Sub cmdFixReferences_Click()
    On Error GoTo ErrorTrap
    
    If mlMissingReferencesCount = 0 Then
        Err.Raise 1000, , "There is no broken references to fix."
    Else
        menModalResult = emrFixPressed
        Me.Hide
    End If
    
Exit Sub
ErrorTrap:
    ShowError
End Sub

Private Sub cmdRemoveReferences_Click()

    On Error GoTo ErrorTrap
    
    Dim lListIndex As Long
    Dim lsiListItem As ListItem
    Dim oReference As Reference
    Dim bLastReadOnlyStatus As Boolean
    bLastReadOnlyStatus = False
    
    For lListIndex = lvwReferences.ListItems.Count To 1 Step -1
        Set lsiListItem = lvwReferences.ListItems(lListIndex)
        If lsiListItem.Selected Then
            Set oReference = GetReferenceObjectFromGUID(moVBProject, lsiListItem.SubItems(4))
            If Not oReference.BuiltIn Then
                bLastReadOnlyStatus = HandleReadOnlyAttrib(moVBProject, True)
                Call RemoveReferenceFromProject(moVBProject, oReference, vbNullString)
                Set oReference = Nothing
                lvwReferences.ListItems.Remove lsiListItem.Index
                Call HandleReadOnlyAttrib(moVBProject, False, bLastReadOnlyStatus)
            Else
                Err.Raise 1000, , "This is built-in default reference which can not be removed"
            End If
        End If
    Next lListIndex

Exit Sub
ErrorTrap:
    Set oReference = Nothing
    Call SaveErrorObj
    Call HandleReadOnlyAttrib(moVBProject, False, bLastReadOnlyStatus)
    Call RestoreErrorObj
    ShowError
End Sub

Private Sub cmdLeaveIt_Click()
    menModalResult = emrClosePressed
    Me.Hide
End Sub

Private Sub Form_Load()
    Me.Caption = Me.Caption & " " & App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Screen.MousePointer = vbDefault
    Set moVBProject = Nothing
End Sub

Private Sub lvwReferences_ColumnClick(ByVal ColumnHeader As ComctlLib.ColumnHeader)
    Call SortListView(lvwReferences, ColumnHeader)
End Sub

Private Sub tmrMakeFormTop_Timer()
    tmrMakeFormTop.Enabled = False
    Call SetForegroundWindow(Me.hwnd)
End Sub



