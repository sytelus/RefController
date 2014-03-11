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
Public FileName As String
Public VBPType As String
Public EXEName32 As String
Public Path32 As String
Public CompatibleEXE32 As String

Public Sub OpenVBP(ByVal vsVBPFile As String)
    Call ProcessVBPFile(vsVBPFile)
End Sub

Private Function ParseReferenceDetails(ByVal p_strLine As String) As VBPReference
    Dim l_intWhereIs As Integer, l_intCheck As Integer, l_intPart As Integer
    Dim l_strTemp As String
    Dim l_blnUntilISaySo As Boolean
    
    Dim p_strGUID As String
      
      Dim oVBPReference As VBPReference
      
      Set oVBPReference = New VBPReference

      ' as usual there is a pain, Controls
      If InStr(p_strLine, ".ocx") > 0 Then
        
        'ToDo - Controls not yet handled
        
      Else
        Do
            ' pull off all the details we need
            l_intWhereIs = InStr(p_strLine, "#")
            If l_intWhereIs > 0 Then
              
              l_strTemp = Left(p_strLine, l_intWhereIs - 1)
              p_strLine = Right(p_strLine, Len(p_strLine) - l_intWhereIs)
            Else
            
              l_strTemp = p_strLine
              l_blnUntilISaySo = True
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
                ' eh, don't know what this means
                oVBPReference.Unknown = l_strTemp
              
              Case 3
                oVBPReference.File = l_strTemp
                oVBPReference.DLLName = GetDLLName(l_strTemp)
                
              Case 4
                oVBPReference.Description = l_strTemp
            
            End Select
            
            l_intPart = l_intPart + 1
        Loop While Not l_blnUntilISaySo
      End If
    
    Set ParseReferenceDetails = oVBPReference
    
End Function

Private Sub ProcessVBPFile(ByVal vsVBPFile As String)
Dim l_strLine As String, l_strName As String, l_strValue As String, l_strType As String, l_strPath As String
Dim l_strCompatabilityExe As String, l_strExeName As String, l_strSQL As String
Dim l_intWhereIs As Integer, l_intFileNumber As Integer, l_intID As Integer
Dim l_intCount As Integer, l_intLoop As Integer
Dim l_varReferences As Variant
  
  ' first open the file,    get a file number
  l_intFileNumber = FreeFile

 VBPRefernces.Clear
    
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
           
      End Select
    End If
  Loop
  
  Close l_intFileNumber
    
End Sub

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
  
  If InStr(p_strDLLName, Chr(0)) > 0 Then
    p_strDLLName = Left(p_strDLLName, InStr(p_strDLLName, Chr(0)) - 1)
  End If
  
    GetDLLName = p_strDLLName
  
End Function

