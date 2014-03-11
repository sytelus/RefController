VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VBPFile"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public VBPRefernces As VBPReferences
Public fileName As String
Public VBPType As String
Public EXEName32 As String
Public Path32 As String
Public CompatibleEXE32 As String
Public ProjectName As String
Public MajorVersion As String
Public MinorVersion As String
Public BuildNumber As String
Public ProjectDescription As String

Public Sub OpenVBP(ByVal vsVBPFile As String)
    Call ProcessVBPFile(vsVBPFile)
End Sub

Private Function ParseReferenceDetails(ByVal p_strLine As String) As VBPReference
    Dim l_intWhereIs As Integer, l_intCheck As Integer, l_intPart As Integer
    Dim l_strTemp As String
    Dim l_blnUntilISaySo As Boolean
    
    Dim bIsOCX As Boolean
    
    Dim p_strGUID As String
      
      Dim oVBPReference As VBPReference
      
      Set oVBPReference = New VBPReference

      If InStr(1, p_strLine, ".ocx", vbTextCompare) > 0 Then
        bIsOCX = True
      Else
        bIsOCX = False
      End If

      oVBPReference.IsOCX = bIsOCX
        
        
        Do
        
            ' pull off all the details we need
            l_intWhereIs = InStr(p_strLine, "#")
            If l_intWhereIs > 0 Then
              
              l_strTemp = Left(p_strLine, l_intWhereIs - 1)
              p_strLine = Right(p_strLine, Len(p_strLine) - l_intWhereIs)
            Else
              If l_intPart = 0 Then 'No # in line - referncing VBP
                  l_intPart = 3
                  p_strLine = Mid(p_strLine, 3)
              End If
              Dim intsemiColPos As Long
              intsemiColPos = InStr(p_strLine, ";")
              If intsemiColPos > 0 Then
                l_strTemp = Left(p_strLine, intsemiColPos - 1)
                p_strLine = Right(p_strLine, Len(p_strLine) - intsemiColPos)
              Else
                l_strTemp = p_strLine
                l_blnUntilISaySo = True
              End If
            End If
            
            Select Case l_intPart
              
              Case 0
                p_strGUID = l_strTemp
                l_intCheck = InStr(p_strGUID, "{")
                If l_intCheck > 0 Then
                  p_strGUID = Right(p_strGUID, Len(p_strGUID) - l_intCheck + 1)
                End If
                
                oVBPReference.GUID = p_strGUID
                                
              Case 1
                oVBPReference.Version = l_strTemp
              
              Case 2
                If Not bIsOCX Then
                    ' eh, don't know what this means
                    oVBPReference.Unknown = l_strTemp
                Else
                    If InStr(l_strTemp, ";") Then
                        l_strTemp = Trim(Right(l_strTemp, Len(l_strTemp) - InStr(l_strTemp, ";")))
                        oVBPReference.File = l_strTemp
                        oVBPReference.DLLName = GetDLLName(l_strTemp)
                    End If
                End If
                
              Case 3
                    oVBPReference.File = l_strTemp
                    oVBPReference.DLLName = GetDLLName(l_strTemp)
                
              Case 4
                oVBPReference.Description = l_strTemp
            
            End Select
            
            l_intPart = l_intPart + 1
        Loop While Not l_blnUntilISaySo
          
    Set ParseReferenceDetails = oVBPReference
    
End Function

Private Sub ProcessVBPFile(ByVal vsVBPFile As String)

On Error GoTo ErrorTrap

Dim l_strLine As String, l_strName As String, l_strValue As String, l_strType As String, l_strPath As String
Dim l_strCompatabilityExe As String, l_strExeName As String, l_strSQL As String
Dim l_intWhereIs As Integer, l_intFileNumber As Integer, l_intID As Integer
Dim l_intCount As Integer, l_intLoop As Integer
Dim l_varReferences As Variant
  
  ' first open the file,    get a file number
  l_intFileNumber = FreeFile

 
 VBPRefernces.Clear
 fileName = vbNullString
 VBPType = vbNullString
 EXEName32 = vbNullString
 Path32 = vbNullString
 CompatibleEXE32 = vbNullString
 ProjectName = vbNullString
MajorVersion = vbNullString
MinorVersion = vbNullString
BuildNumber = vbNullString
ProjectDescription = vbNullString

    ' now open the vbp
  Open vsVBPFile For Input As l_intFileNumber
    
  ' and read all the information in
  Do While Not EOF(l_intFileNumber)
    Line Input #l_intFileNumber, l_strLine
    
    ' Process the information
    l_intWhereIs = InStr(l_strLine, Chr(61))
    
    If l_intWhereIs > 0 Then
      
      ' pull off the name
      l_strName = Left(l_strLine, l_intWhereIs - 1)
      l_strValue = Right(l_strLine, Len(l_strLine) - l_intWhereIs)
      
      Select Case l_strName
        
        Case "Type"
          VBPType = l_strValue
        
        Case "Reference", "Object"
          Call VBPRefernces.Add(ParseReferenceDetails(l_strLine))
          
        Case "ExeName32"
            EXEName32 = l_strValue
             
        Case "Path32"
          ' build path
          Path32 = l_strValue
          
        Case "CompatibleEXE32"
           CompatibleEXE32 = l_strValue
           
        Case "Name"
           ProjectName = RemoveQuotes(l_strValue)
        Case "MajorVer"
            MajorVersion = l_strValue
        Case "MinorVer"
            MinorVersion = l_strValue
        Case "RevisionVer"
            BuildNumber = l_strValue
        Case "Description"
            ProjectDescription = RemoveQuotes(l_strValue)
      End Select
    End If
  Loop
  
  Close l_intFileNumber
  
Exit Sub
ErrorTrap:
      Call SaveErrorObj
      Call SafeCloseFile(l_intCount)
      Call RestoreErrorObj
      ReRaiseError
End Sub

Private Sub SafeCloseFile(ByVal lFileHandle As Long)
    On Error Resume Next
    Close lFileHandle
End Sub

Private Function RemoveQuotes(ByVal vsStringWithQuotes As String) As String
    Dim sReturn As String
    Dim lStart As Long
    Dim lEnd As Long
    Dim lLen As Long
    lLen = Len(vsStringWithQuotes)
    If lLen > 0 Then
        If Left(vsStringWithQuotes, 1) = """" Then
            lStart = 2
        Else
            lStart = 1
        End If
        
        If Right(vsStringWithQuotes, 1) = """" Then
            lEnd = lLen - 1
        Else
            lStart = lLen
        End If
        
        If lStart <= lEnd Then
            sReturn = Mid(vsStringWithQuotes, lStart, lEnd - lStart + 1)
        Else
            sReturn = vbNullString
        End If
    Else
        sReturn = vbNullString
    End If
    
    RemoveQuotes = sReturn
    
End Function

Private Sub Class_Initialize()
    Set VBPRefernces = New VBPReferences
End Sub

Private Sub Class_Terminate()
    Set VBPRefernces = Nothing
End Sub

Private Function GetDLLName(ByVal p_strPath As String) As String
Dim p_strDLLName As String
Dim l_intLoop As Integer
Dim l_strTemp As String, l_strCharacter As String
Dim l_blnISaySo As Boolean
  
  p_strDLLName = ""
  
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
            p_strDLLName = l_strCharacter & p_strDLLName
          End If
          l_intLoop = l_intLoop + 1
        Loop While Not l_blnISaySo
      Else
        p_strDLLName = p_strPath
      End If
    End If
  Else
    p_strDLLName = p_strPath
  End If
  
  If InStr(p_strDLLName, Chr(0)) > 0 Then
    p_strDLLName = Left(p_strDLLName, InStr(p_strDLLName, Chr(0)) - 1)
  End If
  
    GetDLLName = p_strDLLName
  
End Function

