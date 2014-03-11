VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmAddIn 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Fix Missing References"
   ClientHeight    =   3660
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6945
   Icon            =   "AddIn.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   6945
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrAnimation 
      Enabled         =   0   'False
      Interval        =   30
      Left            =   2160
      Top             =   2700
   End
   Begin VB.Frame fraShadowTravel 
      BorderStyle     =   0  'None
      Height          =   405
      Left            =   30
      TabIndex        =   6
      Top             =   3150
      Visible         =   0   'False
      Width           =   2685
      Begin VB.Line Line2 
         BorderColor     =   &H00000000&
         BorderWidth     =   15
         DrawMode        =   6  'Mask Pen Not
         X1              =   -510
         X2              =   1950
         Y1              =   -360
         Y2              =   480
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Your Shield"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   555
         Left            =   -210
         TabIndex        =   7
         Top             =   -90
         Width           =   2940
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   15
         DrawMode        =   6  'Mask Pen Not
         X1              =   -1740
         X2              =   720
         Y1              =   -90
         Y2              =   750
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Your Shield"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   555
         Left            =   -180
         TabIndex        =   8
         Top             =   -60
         Width           =   2940
      End
   End
   Begin VB.CommandButton cmdManageReferences 
      Caption         =   "&Manage It..."
      Height          =   375
      Left            =   4650
      TabIndex        =   5
      Top             =   2100
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.Timer tmrMakeFormTop 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3120
      Top             =   1260
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Left            =   30
      TabIndex        =   4
      Top             =   2940
      Width           =   6885
   End
   Begin VB.CommandButton cmdLeaveIt 
      Cancel          =   -1  'True
      Caption         =   "&Leave It"
      Height          =   375
      Left            =   5910
      TabIndex        =   1
      Top             =   3180
      Width           =   1005
   End
   Begin VB.CommandButton cmdCorrectAll 
      Caption         =   "&Correct It!"
      Default         =   -1  'True
      Height          =   375
      Left            =   4650
      TabIndex        =   0
      Top             =   3180
      Width           =   1155
   End
   Begin ComctlLib.ListView lvwBrokenRef 
      Height          =   2625
      Left            =   0
      TabIndex        =   2
      Top             =   300
      Width           =   6945
      _ExtentX        =   12250
      _ExtentY        =   4630
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      SmallIcons      =   "imlMain"
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
         Text            =   "Path in VBP"
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
   Begin VB.Label lblInfo 
      Caption         =   "Double Click here for info about this animation!"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   735
      Left            =   2730
      TabIndex        =   9
      Top             =   3090
      Visible         =   0   'False
      Width           =   1920
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblHeading 
      AutoSize        =   -1  'True
      Caption         =   "Reference Corrector has found following references to be missing:"
      Height          =   195
      Left            =   60
      TabIndex        =   3
      Top             =   60
      Width           =   4650
   End
   Begin ComctlLib.ImageList imlMain 
      Left            =   6300
      Top             =   2760
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   3
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "AddIn.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "AddIn.frx":075C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "AddIn.frx":0A76
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmAddIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public VBInstance As VBIDE.VBE

Private moVBProject As VBIDE.VBProject
Private moclBrokenRef As Collection
Private moVBPFile As VBPFile
Private mbDoAnimation As Boolean

Private menModalResult As enmModalResult
Private Enum enmModalResult
    emrResponceWaited = 0
    emrClosePressed = 1
    emrManagePressed = 2
    emrCorrectPressed = 3
End Enum


Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long

Public Sub ShowFixReferencesDialog(ByVal VBProject As VBIDE.VBProject, Optional ByVal vboolRaiseErrorIfNoBrokenReferences As Boolean = False)

    Dim lReferenceIndex As Long
    Dim oReference As Reference
    Dim oBrokenReferencesGUID As Collection
    
    Screen.MousePointer = vbHourglass
    
    Set oBrokenReferencesGUID = New Collection
    
    For lReferenceIndex = 1 To VBProject.References.Count
        Set oReference = VBProject.References.Item(lReferenceIndex)
        If Not oReference.BuiltIn Then
            If IsReferenceBroken(oReference) Then
                oBrokenReferencesGUID.Add oReference.GUID, oReference.GUID
            End If
        End If
    Next lReferenceIndex
    
    
    Screen.MousePointer = vbDefault
    
    If oBrokenReferencesGUID.Count > 0 Then
        Call frmAddIn.DisplayForm(oBrokenReferencesGUID, VBProject)
        Unload frmAddIn
    Else
        If vboolRaiseErrorIfNoBrokenReferences Then
            Err.Raise 1000, , "There are no missing references to fix"
        End If
    End If
    
    Set oReference = Nothing
    Set oBrokenReferencesGUID = Nothing
End Sub

Public Function DisplayForm(ByVal voclBrokenReferences As Collection, ByVal voVBProject As VBIDE.VBProject) As Boolean
        
    On Error GoTo ERR_DisplayForm
    
    Set moVBPFile = New VBPFile
    Call moVBPFile.OpenVBP(voVBProject.fileName)
    
    lvwBrokenRef.ListItems.Clear
    
    Dim lBrokenRefIndex As Long
    Dim oVBPReference As VBPReference
    Dim lsiListItem As ListItem
    
    For lBrokenRefIndex = 1 To voclBrokenReferences.Count
        
        Set oVBPReference = moVBPFile.VBPRefernces(voclBrokenReferences(lBrokenRefIndex))
        
        Set lsiListItem = lvwBrokenRef.ListItems.Add(, oVBPReference.GUID, oVBPReference.Description)
        lsiListItem.SubItems(1) = oVBPReference.DLLName
        lsiListItem.SubItems(2) = oVBPReference.Version
        lsiListItem.SubItems(3) = oVBPReference.File
        lsiListItem.SubItems(4) = oVBPReference.GUID
        lsiListItem.SmallIcon = 1
    Next lBrokenRefIndex
    
    'ToDo - Check if VBP is read-only
    
    Set moVBProject = voVBProject
    Set moclBrokenRef = voclBrokenReferences
    
    menModalResult = emrResponceWaited
    tmrMakeFormTop.Enabled = True
    
    mbDoAnimation = True
    Me.Show vbModal
    Call StopAnimation
    
    If menModalResult = emrManagePressed Then
        Call frmFastReferences.DisplayForm(moVBProject)
    End If
    Set moclBrokenRef = Nothing
    Set moVBProject = Nothing
    
    Set moVBPFile = Nothing

Exit Function
ERR_DisplayForm:
    tmrMakeFormTop.Enabled = False
    Set moVBProject = Nothing
    Set moVBPFile = Nothing
    Set moclBrokenRef = Nothing
    ShowError
End Function

Private Function CorrectReferences(ByVal vsDLLPath As String) As Long

    On Error GoTo ERR_CorrectReferences

    Dim lReferenceIndex As Long
    Dim oReference As Reference
    Dim sDLLFile As String
    Dim lUnCorrectedRefCount As Long
    Dim sGUID As String
    Dim bErrorOccured As Boolean
    
    lUnCorrectedRefCount = 0
    
    Screen.MousePointer = vbHourglass
    
    'lvwBrokenRef.SetFocus
    
    For lReferenceIndex = moVBProject.References.Count To 1 Step -1
        Set oReference = moVBProject.References.Item(lReferenceIndex)
        If IsReferenceBroken(oReference) Then
            'Get the DLL for this ref
            sGUID = oReference.GUID
            lvwBrokenRef.ListItems(sGUID).SmallIcon = 3
            'lvwBrokenRef.ListItems(sGUID).Selected = True
            lvwBrokenRef.ListItems(sGUID).EnsureVisible
            DoEvents
            If vsDLLPath = vbNullString Then
                sDLLFile = moVBPFile.VBPRefernces(sGUID).File
            Else
                sDLLFile = vsDLLPath & "\" & moVBPFile.VBPRefernces(sGUID).DLLName
            End If
            bErrorOccured = False
            Call moVBProject.References.AddFromFile(sDLLFile)
            
            If Not bErrorOccured Then
                Call RemoveReferenceFromProject(moVBProject, oReference, sDLLFile)
                lvwBrokenRef.ListItems(sGUID).SmallIcon = 2
            Else
                lUnCorrectedRefCount = lUnCorrectedRefCount + 1
                lvwBrokenRef.ListItems(sGUID).SmallIcon = 1
            End If
            
            DoEvents
        End If
    Next lReferenceIndex
    
    'cmdCorrectAll.SetFocus
    
    Screen.MousePointer = vbDefault
    
    CorrectReferences = lUnCorrectedRefCount
    
Exit Function
ERR_CorrectReferences:
    Select Case Err.Number
        Case 48 'Message: Error loading DLL in References.AddFromFile function
            'Specified DLL doesn't exist so set the flag that error was occured and go ahead with other missing references
            bErrorOccured = True
            Resume Next
        Case 32813  'Message: Name conflicts with existing module, project or object library
            'Reference we are trying to add already exists. This means unbroken refernece already exists along with broken one. So no need to re-add it. Just go ahead.
            Resume Next
        Case Else
            Screen.MousePointer = vbDefault
            ReRaiseError
    End Select
End Function

Private Sub cmdCorrectAll_Click()

    On Error GoTo ERR_cmdCorrectAll_Click

    Dim lUnCorrectedRefCount As Long
    Dim lReferenceIndex As Long
    Dim oReference As Reference
    Dim sDLLPath As String
    
    Dim bCloseForm As Boolean
    Dim bExitLoop As Boolean
    Dim bLastReadOnlyStatus As Boolean
    bLastReadOnlyStatus = False
    
    
    bLastReadOnlyStatus = HandleReadOnlyAttrib(moVBProject, True) 'Remove
    
    sDLLPath = vbNullString
    
    Do
        bCloseForm = True
        bExitLoop = True
    
        lUnCorrectedRefCount = CorrectReferences(sDLLPath)   'use paths in VBP
        
        If lUnCorrectedRefCount > 0 Then
            Dim bRemoveRef As Boolean
            Dim bUserResponce As Boolean
            Dim bUsingVBPPaths As Boolean
            
            If sDLLPath = vbNullString Then
                bUsingVBPPaths = True
            Else
                bUsingVBPPaths = False
            End If
            
            sDLLPath = GetSetting(App.EXEName, msREG_SECTION_LAST_VALUES, msREG_KEY_DLL_PATH, "C:\Lamda5\DLLs")
            
            bUserResponce = frmUnCorrectableRef.DisplayForm(lUnCorrectedRefCount, bUsingVBPPaths, bRemoveRef, sDLLPath)
            If bUserResponce Then
                If bRemoveRef Then
                    For lReferenceIndex = moVBProject.References.Count To 1 Step -1
                        Set oReference = moVBProject.References.Item(lReferenceIndex)
                        If IsReferenceBroken(oReference) Then
                            lvwBrokenRef.ListItems.Remove oReference.GUID
                            Call RemoveReferenceFromProject(moVBProject, oReference, vbNullString)
                        End If
                    Next lReferenceIndex
                Else
                    Call SaveSetting(App.EXEName, msREG_SECTION_LAST_VALUES, msREG_KEY_DLL_PATH, sDLLPath)
                    bExitLoop = False
                End If
            Else
                bCloseForm = False
            End If
            Unload frmUnCorrectableRef
        Else
            MsgBox "All missing references were successfully corrected.", , "Missing References Fixed"
        End If
    Loop While Not bExitLoop
    
    Call HandleReadOnlyAttrib(moVBProject, False, bLastReadOnlyStatus) 'Set
    
    If bCloseForm Then
        menModalResult = emrCorrectPressed
        Me.Hide
    End If

Exit Sub
ERR_cmdCorrectAll_Click:
    Set oReference = Nothing
    Call SaveErrorObj
    Call HandleReadOnlyAttrib(moVBProject, False, bLastReadOnlyStatus)
    Call RestoreErrorObj
    ShowError
End Sub

Private Sub cmdRemoveReferences_Click()

    'Remove All - removed for the time being

    Dim lReferenceIndex As Long
    Dim oReference As Reference
    
    For lReferenceIndex = moVBProject.References.Count To 1 Step -1
        Set oReference = moVBProject.References.Item(lReferenceIndex)
        If IsReferenceBroken(oReference) Then
            Call RemoveReferenceFromProject(moVBProject, oReference, vbNullString)
        End If
    Next lReferenceIndex
    
End Sub

Private Sub cmdLeaveIt_Click()
    menModalResult = emrClosePressed
    Me.Hide
End Sub

Private Sub cmdManageReferences_Click()
    menModalResult = emrManagePressed
    Me.Hide
End Sub

Private Sub Form_Activate()
    If mbDoAnimation Then
        Call RandomizeAnimation
        Call StartAnimation
        mbDoAnimation = False 'One time animation only
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    tmrAnimation.Enabled = False
    Screen.MousePointer = vbDefault
    Set moVBProject = Nothing
    Set moVBPFile = Nothing
    Set moclBrokenRef = Nothing
End Sub

Private Sub lvwBrokenRef_ColumnClick(ByVal ColumnHeader As ComctlLib.ColumnHeader)
    Call SortListView(lvwBrokenRef, ColumnHeader)
End Sub


Private Sub lblInfo_Click()
'This animation is just to create ammusement for VB
'developers and keep guessing them how it might have been created (:-).
'It is not related with rest of the program's function however.
'This animation was implemented using only 2 labels and
'2 line controls AND without using any graphics command.
End Sub

Private Sub RandomizeAnimation()
    '***************************************
    'NOTE: This code can be safely removed without affecting program function
    Randomize
    If Int((100 * Rnd) Mod 5) = 0 Then
        Label1.Caption = "http://i.am/shital"
        Label1.Font.Size = 16
        Label2.Caption = Label1.Caption
        Label2.Font.Size = Label1.Font.Size
    End If
    Label1.ForeColor = RGB(Int(255 * Rnd), Int(255 * Rnd), Int(255 * Rnd))
    Label2.ForeColor = RGB(Int(255 * Rnd), Int(255 * Rnd), Int(255 * Rnd))
    '***************************************
End Sub

Private Sub StartAnimation()

    On Error GoTo ERR_StartAnimation
    
    fraShadowTravel.Visible = False

    Static x1 As Long
    Static x2 As Long
    Static bInitialValueSaved As Boolean
    If Not bInitialValueSaved Then    'Not initialised
        x1 = Line1.x1
        x2 = Line1.x2
        bInitialValueSaved = True
    Else
        'Restore Line1 pos
        Line1.x1 = x1
        Line1.x2 = x2
    End If
    Line2.x1 = Line1.x1
    Line2.x2 = Line1.x2
    Line2.Y1 = Line1.Y1
    Line2.Y2 = Line1.Y2
    Line1.Visible = True
    Line2.Visible = True
    fraShadowTravel.Visible = True
    tmrAnimation.Enabled = True
Exit Sub
ERR_StartAnimation:
    Call StopAnimation
End Sub

Private Sub tmrAnimation_Timer()

    Static lAnimationPhase As Long
    
    On Error GoTo ERR_tmrAnimation_Timer
    
    Select Case lAnimationPhase
        Case 0
            If Line1.x1 > fraShadowTravel.Width Then
                lAnimationPhase = lAnimationPhase + 1
            Else
                'Advance lines
                Line1.x1 = Line1.x1 + 70
                Line1.x2 = Line1.x2 + 70
                Line2.x1 = Line1.x1
                Line2.x2 = Line1.x2
            End If
'        Case 1
'            If Label1.Left >= Label2.Left Then
'                Label1.Top = Label2.Top
'                lAnimationPhase = lAnimationPhase + 1
'            Else
'                Label1.Left = Label1.Left + 2
'            End If
        Case Else
            Call StopAnimation
            lAnimationPhase = 0
    End Select
Exit Sub
ERR_tmrAnimation_Timer:
    Call StopAnimation
End Sub

Private Sub StopAnimation()
    tmrAnimation.Enabled = False
    'fraShadowTravel.Visible = False
    Line1.Visible = False
    Line2.Visible = False
End Sub

Private Sub tmrMakeFormTop_Timer()
    tmrMakeFormTop.Enabled = False
    Call SetForegroundWindow(Me.hwnd)
End Sub
