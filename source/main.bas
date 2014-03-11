Attribute VB_Name = "modMain"
Option Explicit

Public Sub Main()
    If App.StartMode = vbSModeStandalone Then
        Call AddToINI
        'ActiveX EXE registeres itself when it's run but sometime it doesn't so try this manually also.
        Call Shell(App.Path & "\" & App.EXEName & ".exe /REGSERVER", vbMinimizedNoFocus)
        MsgBox "Reference Corrector is now installed." & vbCrLf & "Please restart VB.", , "Reference Corrector"
    End If
End Sub
