VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CredentialsFrm 
   Caption         =   "Change Password"
   ClientHeight    =   3540
   ClientLeft      =   105
   ClientTop       =   435
   ClientWidth     =   7410
   OleObjectBlob   =   "CredentialsFrm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CredentialsFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cancelButton_Click()
    Unload Me
End Sub

Private Sub pwSetButton_Click()
'Check all inputs are valid prior to allowing a change of password
'Removes any currently protected features with old password then re-protects it with new one

    Dim currentPW As String
    Dim oldPW As String
    Dim newPW1 As String
    Dim newPW2 As String
    Dim errorCount As Integer
    Dim sh As Worksheet
    
    errorCount = 0
    'Obtain data/passwords from user input
    oldPW = Me.TextBox1.Value
    newPW1 = Me.TextBox2.Value
    newPW2 = Me.TextBox3.Value
    
    'Ensure user has permission to change the password prior to setting a new password
    'If the user is aware of the old password, it is assumed he/she is authorised.
    
    'If no password has been set yet, set currentPW to empty string and continue
    'Note: The check for whether or not the password already exists should have been
    'performed prior to this form being invoked, preventing user from inputting text
    'in to the old password field.
    
    If (CredentialsFrm.TextBox1.Enabled = True) Then
        currentPW = getPassword()
        If (currentPW <> oldPW) Then
            Call MsgBox("Old password is incorrect!")
            GoTo clean
        End If
    Else
        currentPW = ""
    End If
    
    'Ensure user has not performed a typo before setting new password
    If (newPW1 <> newPW2) Then
        Call MsgBox("Passwords entered do not match. Please ensure New Password and Repeat Password entries are the same.")
        GoTo clean
    End If
    
    'User is not allowed to set the password to an empty string
    If (newPW1 = "") Then
        Call MsgBox("Password cannot be an empty string.")
        GoTo clean
    End If
    
    'If all check conditions above are sucessful, attempt to switch over all
    'protected sheets and store new password
    
    'Can get a bit messy from here if user has been setting protect passwords
    'manually that do not match with the stored password by the pwManagement
    'module
    
    On Error GoTo handler
    
    'Loop through all sheets removing old password and replacing with new password
    Application.ScreenUpdating = False
    For Each sh In ThisWorkbook.Worksheets
        If (sh.ProtectContents = True) Then
            sh.Unprotect password:=currentPW
            sh.Protect password:=newPW1
        End If
    Next sh
    Application.ScreenUpdating = True
    
    '////////////////////////////////////////////////////////////
    '////////////////////////////////////////////////////////////
    'Add other protectable items here e.g. workbook structure etc
    '////////////////////////////////////////////////////////////
    '////////////////////////////////////////////////////////////
    
    'Put new password in to storage location
    Call setPassword(newPW1)

    Call MsgBox("Password has been sucessfully changed " & _
        "with " & errorCount & " errors.")
    
    Unload Me
    
    Exit Sub
    
clean:
    'User has done something silly, clean the form and let them try again
    Me.TextBox1.Value = ""
    Me.TextBox2.Value = ""
    Me.TextBox3.Value = ""
    Exit Sub

handler:
    'Users have been messing with the passwords manually. Can add better error
    'handling later.
    errorCount = errorCount + 1
    MsgBox ("Was unable to change password for " & sh.Name & _
        ". Please remove/change manually.")
    Resume Next
    
End Sub
