Attribute VB_Name = "Utils"
Option Explicit

Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Const BIF_RETURNONLYFSDIRS = 1
Private Type BrowseInfo
    hWndOwner       As Long
    pIDLRoot        As Long
    pszDisplayName  As String
    lpszTitle       As String
    ulFlags         As Long
    lpfnCallBack    As Long
    lparam          As Long
    iImage          As Long
End Type

Private Type udtErrorInfo
    Number As Long
    Description  As String
    HelpContext As Long
    HelpFile As String
    Source As String
End Type

Private muErr As udtErrorInfo
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Const MAX_PATH = 260

Public Function GetPathWithSlash(ByVal vsPath As String, Optional ByVal vsSlashChar As String = "\") As String
    If Right$(vsPath, 1) <> vsSlashChar Then
        GetPathWithSlash = vsPath & vsSlashChar
    Else
        GetPathWithSlash = vsPath
    End If
End Function

Public Function RemoveSlashAtEnd(ByVal vsPath As String, Optional ByVal vsSlashChar As String = "\") As String
    Dim lPathLen As Long
    
    lPathLen = Len(vsPath)
    
    If Right$(vsPath, 1) <> vsSlashChar Then
        RemoveSlashAtEnd = vsPath
    Else
        If lPathLen > 1 Then
            RemoveSlashAtEnd = Left(vsPath, lPathLen - 1)
        Else
            RemoveSlashAtEnd = vbNullString
        End If
    End If

End Function

Public Sub ReRaiseError()
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub

Public Sub DeleteFile(ByVal vsFileName As String)
    On Error Resume Next
    Call Kill(vsFileName)
End Sub

Public Sub ShowError()
    Screen.MousePointer = vbDefault
    MsgBox "Error " & IIf(Err.Number = 1000, vbNullString, Err.Number) & " : " & Err.Description, vbOKOnly, App.Title & " Error"
End Sub

Public Function ExtractFilePath(ByVal vsFileSpec As String, Optional ByVal vsSlashChar As String = "\") As String
    Dim lDotPos As Long
    Dim lLastDotPos As Long
    lDotPos = 0
    Do
        lLastDotPos = lDotPos
        lDotPos = InStr(lDotPos + 1, vsFileSpec, vsSlashChar)
    Loop While lDotPos <> 0
        
    If lLastDotPos <> 0 Then
        ExtractFilePath = Left(vsFileSpec, lLastDotPos)
    Else
        ExtractFilePath = vsFileSpec
    End If
    ExtractFilePath = RemoveSlashAtEnd(ExtractFilePath, vsSlashChar)
End Function

Public Function ExtractFileExtension(ByVal vsFileSpec As String) As String
    Dim lDotPos As Long
    Dim lLastDotPos As Long
    
    Const sDOT As String = "."
    
    lDotPos = 0
    Do
        lLastDotPos = lDotPos
        lDotPos = InStr(lDotPos + 1, vsFileSpec, sDOT)
    Loop While lDotPos <> 0
        
    If lLastDotPos <> 0 Then
        ExtractFileExtension = Mid(vsFileSpec, lLastDotPos + 1)
    Else
        ExtractFileExtension = vbNullString
    End If
End Function

Public Function GetTempDirPath() As String
    Dim sBuffer As String
    Dim lTempDirLen As Long
    
    sBuffer = String$(MAX_PATH, 0)
    lTempDirLen = GetTempPath(MAX_PATH, sBuffer)
    If lTempDirLen > 1 Then
        GetTempDirPath = RemoveSlashAtEnd(Left$(sBuffer, lTempDirLen - 1))
    Else
        GetTempDirPath = RemoveSlashAtEnd(App.Path)
    End If
End Function

Public Function CheckToBool(ByVal venmCheckBoxValue As CheckBoxConstants) As Boolean
    CheckToBool = (venmCheckBoxValue = vbChecked)
End Function

Public Function BoolToCheck(ByVal vboolValue As Boolean) As CheckBoxConstants
    BoolToCheck = IIf(vboolValue, vbChecked, vbUnchecked)
End Function

Public Function ClearCollection(ByVal voclColl As Collection)
    Dim i As Long
    
    For i = voclColl.Count To 1 Step -1
        voclColl.Remove i
    Next i
    
End Function

Public Function IsFileExist(ByVal vsFileName As String) As Boolean
    On Error GoTo ERR_IsFileExist
    Dim lDummy As Long
    lDummy = FileLen(vsFileName)
    IsFileExist = True
        
Exit Function
ERR_IsFileExist:

    If Err.Number = 76 Or Err.Number = 53 Then
        IsFileExist = False
    Else
        ReRaiseError
    End If
End Function

Public Function BrowseForFolder(ByVal vsDefaultFolder As String, Optional ByVal vhWndOwner As Long = 0, Optional ByVal vsPrompt As String = vbNullString) As String
    Dim iNull As Integer
    Dim lpIDList As Long
    Dim lResult As Long
    Dim sPath As String
    Dim udtBI As BrowseInfo

    With udtBI
        .hWndOwner = vhWndOwner
        .lpszTitle = vsPrompt
        .ulFlags = BIF_RETURNONLYFSDIRS
    End With
    
    lpIDList = SHBrowseForFolder(udtBI)
    If lpIDList Then
        sPath = String$(MAX_PATH, 0)
        lResult = SHGetPathFromIDList(lpIDList, sPath)
        Call CoTaskMemFree(lpIDList)
        iNull = InStr(sPath, vbNullChar)
        If iNull Then
            sPath = Left$(sPath, iNull - 1)
        End If
    Else
        sPath = vsDefaultFolder
    End If

    BrowseForFolder = sPath
    
End Function

Public Function IsItemExistInColl(ByVal voclColl As Collection, ByVal vvItemIndex As Variant) As Boolean
    On Error Resume Next
    Dim vItem As Variant
    vItem = voclColl(vvItemIndex)
    IsItemExistInColl = (Err.Number = 0)
End Function

Public Function ExtractFileName(ByVal p_strPath As String) As String
Dim p_strFileName As String
Dim l_intLoop As Integer
Dim l_strTemp As String, l_strCharacter As String
Dim l_blnISaySo As Boolean
  
  p_strFileName = ""
  
  'Check two colon bug in VBP
  Dim lColon1Pos As Long
  Dim lColon2Pos As Long
  Dim lSlashPos As Long
  Dim bIsBuggyName As Boolean
  bIsBuggyName = False
  lColon1Pos = InStr(1, p_strPath, ":")
  lSlashPos = InStr(1, p_strPath, "\")
  If lColon1Pos > 0 Then
    lColon2Pos = InStr(lColon1Pos + 1, p_strPath, ":")
    If lColon2Pos > 0 Then
        bIsBuggyName = True
    Else
        If Mid(p_strPath, lColon1Pos + 1, 1) = "\" And lSlashPos < lColon1Pos Then
            bIsBuggyName = True
        End If
    End If
  End If
  
  If Not bIsBuggyName Then
    If Len(p_strPath) > 0 Then
      
      ' now from the path get the name, work back from the end, avoids relative paths
      If InStr(p_strPath, Chr(92)) > 0 Then
        Do
          l_strCharacter = Mid(p_strPath, Len(p_strPath) - l_intLoop, 1)
          If l_strCharacter = Chr(92) Then
            l_blnISaySo = True
          Else
            p_strFileName = l_strCharacter & p_strFileName
          End If
          l_intLoop = l_intLoop + 1
        Loop While Not l_blnISaySo
      Else
        p_strFileName = p_strPath
      End If
    End If
  Else
    p_strFileName = p_strPath
  End If
  
  If InStr(p_strFileName, Chr(0)) > 0 Then
    p_strFileName = Left(p_strFileName, InStr(p_strFileName, Chr(0)) - 1)
  End If
  
    ExtractFileName = p_strFileName
  
End Function

Public Sub SaveErrorObj()
    With muErr
        .Number = Err.Number
        .Description = Err.Description
        .HelpContext = Err.HelpContext
        .HelpFile = Err.HelpFile
        .Source = Err.Source
    End With
End Sub

Public Sub RestoreErrorObj()
    With Err
        .Number = muErr.Number
        .Description = muErr.Description
        .HelpContext = muErr.HelpContext
        .HelpFile = muErr.HelpFile
        .Source = muErr.Source
    End With
End Sub

Public Sub SortListView(ByVal vlsvListView As ListView, ByVal voColumnHeaderClicked As ColumnHeader)
    With vlsvListView
        .SortKey = voColumnHeaderClicked.Index - 1
        .SortOrder = IIf(.SortOrder = lvwAscending, lvwDescending, lvwAscending)
        .Sorted = False
        .Sorted = True
    End With
End Sub
