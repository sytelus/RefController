VERSION 5.00
Begin VB.Form frmRCInstallMain 
   Caption         =   "Reference Corrector Installation"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmRCInstallMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

    Const sAddInDLL As String = "RefCorrect.dll"

    Dim sDLLFileWithPath As String
    
    sDLLFileWithPath = RemoveSlashAtEnd(App.Path) & "\" & sAddInDLL

    If IsFileExist(sDLLFileWithPath) Then
        Call Shell("regsvr32 /s " & """" & sDLLFileWithPath & """")
        AddToINI
        MsgBox "Reference Corrector VB Add-In is now installed." & vbCrLf & "Please restart VB."
    Else
        MsgBox """" & sDLLFileWithPath & """" & " file does not exist. This File is required for installation." & vbCrLf & "Installation aborted."
    End If
    
    End

End Sub
