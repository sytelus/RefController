Attribute VB_Name = "modCommon"
Option Explicit

Public Const msREG_SECTION_LAST_VALUES As String = "Last Values"
Public Const msREG_KEY_DLL_PATH As String = "DLLPath"

Public Function SafeGetAttr(ByVal vsFileName As String) As Long
    On Error GoTo ErrorTrap
    SafeGetAttr = GetAttr(vsFileName)
Exit Function
ErrorTrap:
    If Err.Number = 53 Then 'File not found
        SafeGetAttr = 0
    Else
        ReRaiseError
    End If
End Function

Public Sub SafeSetAttr(ByVal vsFileName As String, ByVal vlAttrib As Long)
    On Error GoTo ErrorTrap
    Call SetAttr(vsFileName, vlAttrib)
Exit Sub
ErrorTrap:
    If Err.Number <> 53 Then 'File not found
        ReRaiseError
    End If
End Sub

Public Function IsProjectReadOnly(ByVal voVBProject As VBProject) As Boolean
    IsProjectReadOnly = (GetAttr(voVBProject.fileName) And vbReadOnly = vbReadOnly)
End Function

Public Function HandleReadOnlyAttrib(ByVal voVBProject As VBProject, ByVal vboolRemoveOrSet As Boolean, Optional ByVal vboolLastReadOnlyStatus As Boolean = False) As Boolean
    
    Dim bWasReadOnly As Boolean
    Dim lVBPFileAttrib As Long
    
    If vboolRemoveOrSet Then
        bWasReadOnly = False
        lVBPFileAttrib = SafeGetAttr(voVBProject.fileName)
        If (lVBPFileAttrib And vbReadOnly) = vbReadOnly Then
            bWasReadOnly = True
            Call SafeSetAttr(voVBProject.fileName, lVBPFileAttrib And (Not vbReadOnly))
        End If
    Else
        If vboolLastReadOnlyStatus Then
            bWasReadOnly = False
            lVBPFileAttrib = SafeGetAttr(voVBProject.fileName)
            If Not ((lVBPFileAttrib And vbReadOnly) = vbReadOnly) Then
                Call SafeSetAttr(voVBProject.fileName, lVBPFileAttrib Or vbReadOnly)
            End If
        End If
    End If
    
    HandleReadOnlyAttrib = bWasReadOnly
    
End Function

Public Function GetMissingReferenceCount(ByVal voVBProject As VBProject) As Long
    Dim lReferenceIndex As Long
    Dim oReference As Reference
    Dim lMissingReferenceCount As Long
        
    lMissingReferenceCount = 0
    
    For lReferenceIndex = 1 To voVBProject.References.Count
        Set oReference = voVBProject.References.Item(lReferenceIndex)
        If Not oReference.BuiltIn Then
            If IsReferenceBroken(oReference) Then
                lMissingReferenceCount = lMissingReferenceCount + 1
            End If
        End If
    Next lReferenceIndex
    
    GetMissingReferenceCount = lMissingReferenceCount
    
End Function

Public Function GetReferenceObjectFromGUID(ByVal voVBProject As VBProject, ByVal vsGUID As String) As Reference
    Dim oReference As Reference
    
    Set GetReferenceObjectFromGUID = Nothing
    For Each oReference In voVBProject.References
        If oReference.GUID = vsGUID Then
            Set GetReferenceObjectFromGUID = oReference
            Exit For
        End If
    Next oReference
End Function

Public Function SafeAddListItem(ByVal vlsvListView As ListView, ByVal vsKey As String, ByVal vsText As String) As ListItem
    On Error Resume Next
    
    With vlsvListView
        Set SafeAddListItem = .ListItems.Add(, vsKey, vsText)
        If Err.Number <> 0 Then
            Set SafeAddListItem = .ListItems.Add(, , vsText)
        End If
    End With
End Function

Private Sub CreateDummyKeys(ByVal vsGUID As String, ByVal vsVersion As String, ByVal vsDLLFile As String)
    Call SetRegKey(HKEY_CLASSES_ROOT, "TypeLib\" & vsGUID & "\" & vsVersion & "\0\win32", "", vsDLLFile)
    Call SetRegKey(HKEY_CLASSES_ROOT, "TypeLib\" & vsGUID & "\" & vsVersion & "\FLAGS", "", "0")
    Call SetRegKey(HKEY_CLASSES_ROOT, "TypeLib\" & vsGUID & "\" & vsVersion & "\HELPDIR", "", "")
End Sub

Private Sub DeleteDummyKeys(ByVal vsGUID As String, ByVal vsVersion As String)
    Call RegDeleteKey(HKEY_CLASSES_ROOT, "TypeLib\" & vsGUID & "\" & vsVersion & "\0\win32")
    Call RegDeleteKey(HKEY_CLASSES_ROOT, "TypeLib\" & vsGUID & "\" & vsVersion & "\0")
    Call RegDeleteKey(HKEY_CLASSES_ROOT, "TypeLib\" & vsGUID & "\" & vsVersion & "\FLAGS")
    Call RegDeleteKey(HKEY_CLASSES_ROOT, "TypeLib\" & vsGUID & "\" & vsVersion & "\HELPDIR")
    Call RegDeleteKey(HKEY_CLASSES_ROOT, "TypeLib\" & vsGUID & "\" & vsVersion)
    Call RegDeleteKey(HKEY_CLASSES_ROOT, "TypeLib\" & vsGUID)
End Sub

Public Function IsReferenceBroken(ByVal voReference As Reference) As Boolean
    On Error GoTo ErrorTrap
    
    Dim bIsReferenceBroken As Boolean
    
    bIsReferenceBroken = voReference.IsBroken
    
    If Not bIsReferenceBroken Then
        On Error Resume Next
        Dim sFullPath As String
        
        'Try to access full path property
        sFullPath = voReference.FullPath
        
        bIsReferenceBroken = (Err.Number <> 0)
        
        On Error GoTo ErrorTrap
        
    End If
    IsReferenceBroken = bIsReferenceBroken
Exit Function
ErrorTrap:
    ReRaiseError
End Function

Public Sub RemoveReferenceFromProject(ByVal voVBProject As VBProject, ByVal voReference As Reference, ByVal vsDLLFile As String)

    'VB6 doesn't allows to remove reference unless DLL/OCX is registered - which is
    'not the case when reference is broken. So in all, if reference is broken, VB6 doesn't
    'allows it to remove it using it's object model. However VB5 does allows to do so. To work around this,
    'create dummy entries in registry to make it appeare that DLL/OCX is registered and after
    'removing reference from the project delete those dummy entries.

    On Error GoTo ERR_RemoveReferenceFromProject
    
    Dim sVersion As String
    Dim sGUID As String
    Dim bIsReferenceBroken As Boolean
    Dim bDummyKeysCreated As Boolean
    
    bDummyKeysCreated = False
    bIsReferenceBroken = IsReferenceBroken(voReference)
    
    If bIsReferenceBroken Then
        sVersion = voReference.Major & "." & voReference.Minor
        sGUID = voReference.GUID
        Call CreateDummyKeys(sGUID, sVersion, vsDLLFile)
        bDummyKeysCreated = True
    End If
    
    Call voVBProject.References.Remove(voReference)
    
    If bIsReferenceBroken Then
        Call DeleteDummyKeys(sGUID, sVersion)
        bDummyKeysCreated = False
    End If
    
Exit Sub
ERR_RemoveReferenceFromProject:
    If bDummyKeysCreated = True Then
        Call DeleteDummyKeys(sGUID, sVersion)
        bDummyKeysCreated = False
    End If
    ReRaiseError
End Sub

