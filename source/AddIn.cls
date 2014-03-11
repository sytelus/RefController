VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AddIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Implements IDTExtensibility

Public FormDisplayed          As Boolean
Public VBInstance             As VBIDE.VBE

Public WithEvents moVBProjectsEvents As VBProjectsEvents
Attribute moVBProjectsEvents.VB_VarHelpID = -1

Dim mcbMenuCommandBar         As Office.CommandBarControl
Dim mfrmAddIn                 As New frmAddIn
Public WithEvents MenuHandler As CommandBarEvents          'command bar event handler
Attribute MenuHandler.VB_VarHelpID = -1

Private Const APP_CATEGORY = "Microsoft Visual Basic AddIns"

Sub Hide()
    
    On Error Resume Next
    
    FormDisplayed = False
    mfrmAddIn.Hide
   
End Sub

Sub Show()
      
    On Error GoTo ErrorTrap
    
    If mfrmAddIn Is Nothing Then
        Set mfrmAddIn = New frmAddIn
    End If
    
    Set mfrmAddIn.VBInstance = VBInstance
    'Set mfrmAddIn.Connect = Me
    FormDisplayed = True
    mfrmAddIn.Show
   
Exit Sub
ErrorTrap:
    ShowError
End Sub

'------------------------------------------------------
'this method adds the Add-In to VB
'------------------------------------------------------
Private Sub IDTExtensibility_OnConnection(ByVal VBInst As Object, ByVal ConnectMode As vbext_ConnectMode, ByVal AddInInst As VBIDE.AddIn, custom() As Variant)
    
    On Error GoTo error_handler
    
    'save the vb instance
    Set VBInstance = VBInst
    
    Set moVBProjectsEvents = VBInstance.Events.VBProjectsEvents
    
    'this is a good place to set a breakpoint and
    'test various addin objects, properties and methods

    If ConnectMode = vbext_cm_External Then
        'Used by the wizard toolbar to start this wizard
        'Me.Show
    Else
        Set mcbMenuCommandBar = AddToAddInCommandBar("&Fast References...")
        'sink the event
        Set Me.MenuHandler = VBInst.Events.CommandBarEvents(mcbMenuCommandBar)
        'Call AddToCommandBar
    End If
  
    If ConnectMode = vbext_cm_AfterStartup Then
'        If GetSetting(App.Title, "Settings", "DisplayOnConnect", "0") = "1" Then
            'set this to display the form on connect
            'Me.Show
'        End If
    End If
  
Exit Sub
error_handler:
    
    ShowError
    
End Sub

'------------------------------------------------------
'this method removes the Add-In from VB
'------------------------------------------------------
Private Sub IDTExtensibility_OnDisconnection(ByVal RemoveMode As vbext_DisconnectMode, custom() As Variant)
    
    On Error GoTo error_handler
    
    'delete the command bar entry
    mcbMenuCommandBar.Delete
    
    'shut down the Add-In
    If FormDisplayed Then
        SaveSetting App.Title, "Settings", "DisplayOnConnect", "1"
        FormDisplayed = False
    Else
        SaveSetting App.Title, "Settings", "DisplayOnConnect", "0"
    End If
    
    Unload mfrmAddIn
    Set mfrmAddIn = Nothing

Exit Sub
error_handler:
    
    ShowError
    
End Sub

Private Sub IDTExtensibility_OnStartupComplete(custom() As Variant)
'    If GetSetting(App.Title, "Settings", "DisplayOnConnect", "0") = "1" Then
'        'set this to display the form on connect
'        Me.Show
'    End If
End Sub

Private Sub IDTExtensibility_OnAddInsUpdate(custom() As Variant)
'
End Sub

'this event fires when the menu is clicked in the IDE
Private Sub MenuHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    On Error GoTo ErrorTrap
    If Not (VBInstance.ActiveVBProject Is Nothing) Then
        Call frmFastReferences.DisplayForm(VBInstance.ActiveVBProject)
    Else
        Err.Raise 1000, , "No Active Project in VB IDE"
    End If
Exit Sub
ErrorTrap:
    ShowError
End Sub

Function AddToAddInCommandBar(sCaption As String) As Office.CommandBarControl
    Dim cbMenuCommandBar As Office.CommandBarControl  'command bar object
    Dim cbMenu As Object
  
    On Error GoTo AddToAddInCommandBarErr
    
    'see if we can find the Add-Ins menu
    Set cbMenu = VBInstance.CommandBars("Add-Ins")
    If cbMenu Is Nothing Then
        'not available so we fail
        Exit Function
    End If
    
    'add it to the command bar
    Set cbMenuCommandBar = cbMenu.Controls.Add(1)
    'set the caption
    cbMenuCommandBar.Caption = sCaption
    
    Set AddToAddInCommandBar = cbMenuCommandBar
    
Exit Function
AddToAddInCommandBarErr:
    ShowError
End Function

Sub AddToCommandBar()
  On Error GoTo AddToCommandBarErr
  
  'make sure the standard toolbar is visible
  VBInstance.CommandBars(2).Visible = True
  
  'add it to the command bar
  'the following line will add the TabOrder manager to the
  'Standard toolbar to the right of the ToolBox button
  Set mcbMenuCommandBar = VBInstance.CommandBars(2).Controls.Add(1, , , VBInstance.CommandBars(2).Controls.Count)
  'set the caption
  mcbMenuCommandBar.Caption = "Show Fields"
  'copy the icon to the clipboard
  Clipboard.SetData LoadResPicture(1000, 0)
  'set the icon for the button
  mcbMenuCommandBar.PasteFace
  
  'sink the event
  Set Me.MenuHandler = VBInstance.Events.CommandBarEvents(mcbMenuCommandBar)
  
  'restore the last state
'  If GetSetting(APP_CATEGORY, App.Title, "DisplayOnConnect", "0") = "1" Then
    'set this to display the form on connect
    'Me.Show
'  End If
  
Exit Sub
AddToCommandBarErr:
  ShowError
End Sub

Private Sub moVBProjectsEvents_ItemAdded(ByVal VBProject As VBIDE.VBProject)
    On Error GoTo ErrorTrap
    Call frmAddIn.ShowFixReferencesDialog(VBProject)
Exit Sub
ErrorTrap:
    ShowError
End Sub

