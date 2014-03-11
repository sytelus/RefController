VERSION 5.00
Begin VB.Form frmUnCorrectableRef 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Uncorrectable References"
   ClientHeight    =   2250
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "UnCorrectableRef.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   120
      Left            =   90
      TabIndex        =   7
      Top             =   1500
      Width           =   4515
   End
   Begin VB.CommandButton cmdBrowse 
      Height          =   315
      Left            =   4230
      Picture         =   "UnCorrectableRef.frx":0742
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   750
      Width           =   375
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3510
      TabIndex        =   5
      Top             =   1770
      Width           =   1065
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2310
      TabIndex        =   4
      Top             =   1770
      Width           =   1065
   End
   Begin VB.TextBox txtDLLPath 
      Height          =   315
      Left            =   2070
      TabIndex        =   1
      Top             =   750
      Width           =   2115
   End
   Begin VB.OptionButton optRemoveRef 
      Caption         =   "&Remove Uncorrectable References (faster loading)"
      Height          =   315
      Left            =   90
      TabIndex        =   3
      Top             =   1170
      Width           =   4035
   End
   Begin VB.OptionButton optSearchInDLLDir 
      Caption         =   "&Search DLLs in this dir:"
      Height          =   255
      Left            =   90
      TabIndex        =   0
      Top             =   780
      Width           =   1935
   End
   Begin VB.Label lblInfo 
      Caption         =   "Label1"
      Height          =   495
      Left            =   90
      TabIndex        =   6
      Top             =   120
      Width           =   4500
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmUnCorrectableRef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mbOKPressed As Boolean

Private Sub cmdBrowse_Click()
    txtDLLPath.Text = BrowseForFolder(txtDLLPath.Text, Me.hwnd, "Select Folder For DLLs")
End Sub

Private Sub cmdCancel_Click()
    mbOKPressed = False
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    mbOKPressed = True
    Me.Hide
End Sub

Public Function DisplayForm(ByVal vlUncorrectableRefCount As Long, ByVal vboolUsingVBPPath As Boolean, ByRef rbRemoveRef As Boolean, ByRef rsDLLDir As String) As Boolean
    
    mbOKPressed = False
    
    lblInfo.Caption = "There were " & vlUncorrectableRefCount & " references that could not be corrected because DLLs "
    If vboolUsingVBPPath Then
        lblInfo.Caption = lblInfo.Caption & "as in VBP file "
    Else
        lblInfo.Caption = lblInfo.Caption & "as in specified path "
    End If
    lblInfo.Caption = lblInfo.Caption & "doesn't exist on your machine."
    
    txtDLLPath.Enabled = True
    cmdBrowse.Enabled = txtDLLPath.Enabled
    
    If rbRemoveRef Then
        optRemoveRef.Value = True
    Else
        optSearchInDLLDir.Value = True
    End If
    txtDLLPath.Text = rsDLLDir
    
    If Not vboolUsingVBPPath Then
        txtDLLPath.ForeColor = vbRed
        txtDLLPath.BackColor = vbYellow
    Else
        txtDLLPath.ForeColor = vbWindowText
        txtDLLPath.BackColor = vbWindowBackground
    End If
    
    Me.Show vbModal
    
    If optRemoveRef.Value = True Then
        rbRemoveRef = True
    Else
        rbRemoveRef = False
    End If
    rsDLLDir = txtDLLPath.Text
    
    DisplayForm = mbOKPressed
    
End Function

Private Sub optRemoveRef_Click()
    txtDLLPath.Enabled = False
    cmdBrowse.Enabled = txtDLLPath.Enabled
End Sub

Private Sub optSearchInDLLDir_Click()
    txtDLLPath.Enabled = True
    cmdBrowse.Enabled = txtDLLPath.Enabled
End Sub
