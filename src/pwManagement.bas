Attribute VB_Name = "pwManagement"
'Password management and changer module
'Overview:
'Allows user to set and store a password/text in excel not readily accessible by users
'Perfect base for creating apps/workbooks with multiple users that are expected to have
'different privileges without hardcoding a password in to the VBA code. Users with sufficient
'privileges are able to change the password outside of VBA allowing separate protection for
'the VBA code/module.

'Description:
'This module provides a set of functions that can be used inside VBA in conjuction with
'the protect feature to enable/disable certain functions or features, such as blocking
'the editing of certain cells etc unless they are authorised to do so i.e. they know the
'password.

'Main functions/calls:
    'getPassword()
    'Returns the currently stored password in human readable format. If no password is
    'stored, returns an empty string.

    'changePassword()
    'Calls the CredentialsFrm, allowing using to change/set the stored password for the
    'workbook. Create a button and link it's call macro to this function if you want the
    'user to be able to change/set the password. If you don't want the user to be able to
    'change the password this procedure can be invoked manaully through the VBA editor.
    
'Requirements:
    'Form - CredentialsFrm
    'Microsoft XML v3/4 - For encrypt feature


'Additional Notes:
'This module should be implemented on a workbook where no passwords have been set yet
'The user should not be setting protection outside of VBA eg. manually setting a password
'for a given sheet etc.
'Lock this module before sending it to client/users

'Special Note: Excel protection is a bit of a joke as it can be bypassed fairly easily without
'a password. Therefore it is pointless to perform complicated encryption on the password i.e.
'if the user is smart enough to figure out the storage location and unjumble the password then
'they most likely are capable of bypassing Excel's protection all together, assuming you have
'locked this module and they are not aware of the inner workings.


Option Explicit

Sub changePassword()
'Opens form CredentialsFrm to all user to set/change password stored by pwManagement module

Dim currentPassword As String

currentPassword = getPassword()

If (currentPassword = "") Then
    'No password has been stored yet by pwManagement module
    
    'Disable the old password input box before invoking the form
    'The state of the input box will flag to other functions the
    'password was not previously set
    
    CredentialsFrm.TextBox1.Enabled = False
    CredentialsFrm.TextBox1.BackStyle = fmBackStyleTransparent
End If

CredentialsFrm.Show

End Sub


Function getPassword() As String
'Retrieves password from storage location decrypts it and returns human
'readable password string

    Dim storedPassword As String
    Dim actualPassword As String

    storedPassword = retrievePassword()
    actualPassword = decryptText(storedPassword)
    getPassword = actualPassword

End Function

Sub setPassword(password As String)
'Receives human readable password string and encrypts it before storing it

    Dim encryptedPassword As String

    encryptedPassword = encryptText(password)
    Call storePassword(encryptedPassword)

End Sub


Function retrievePassword() As String
'Retrieves password from storage location, if storage location doesn't exist
'yet, password has not been set yet (at least not by the pwManagement module)

    Dim password As String

    On Error GoTo handler

    password = ActiveWorkbook.CustomDocumentProperties("SpecialVal1").Value

    retrievePassword = password

    Exit Function
    
handler:
    retrievePassword = ""
    MsgBox ("No password has been set yet")

End Function

Sub storePassword(password As String)
'Places password in storage location, if storage location doesn't already exist
'creates location

    On Error Resume Next
    ActiveWorkbook.CustomDocumentProperties("SpecialVal1").Value = password

    If Err.Number > 0 Then
        With ActiveWorkbook.CustomDocumentProperties
            .Add Name:="SpecialVal1", _
                LinkToContent:=False, _
                Type:=msoPropertyTypeString, _
                Value:=password
        End With
    End If

End Sub

Function encryptText(text As String) As String
'Encodes text in to base64 format, so it isn't easy readable
'Not actually encryption but Excel protection is a joke anyway
'and can be bypassed regardless of the password/encryption
  
'Requires:
'Microsoft XML v3/4
    
    Dim objXML As MSXML2.DOMDocument
    Dim objNode As MSXML2.IXMLDOMElement
    Dim arrData() As Byte
    
    'Use MSXML to convert text to base64
    Set objXML = New MSXML2.DOMDocument
    Set objNode = objXML.createElement("b64")
    
    arrData = StrConv(text, vbFromUnicode)
    objNode.DataType = "bin.base64"
    objNode.nodeTypedValue = arrData
    
    encryptText = objNode.text

    'Cleanup
    Set objNode = Nothing
    Set objXML = Nothing
    
    Exit Function
    
handler:
    'Provide warning and just pass text as is i.e. without encryption
    MsgBox ("Encryption of text failed! Your text will not be secure")
    encryptText = text
    
End Function

Function decryptText(text As String) As String

'Decodes base64 back in to readable text

'Requires:
'Microsoft XML v3/4

    Dim objXML As MSXML2.DOMDocument
    Dim objNode As MSXML2.IXMLDOMElement
    
    On Error GoTo handler
    
    If (text <> "") Then
        'Use MSXML to convert base64 to text
        Set objXML = New MSXML2.DOMDocument
        Set objNode = objXML.createElement("b64")
    
        objNode.DataType = "bin.base64"
        objNode.text = text
        decryptText = objNode.nodeTypedValue
    
        'Cleanup
        Set objNode = Nothing
        Set objXML = Nothing
        
        'Pass back human readable text
        decryptText = StrConv(decryptText, vbUnicode)
    Else
        decryptText = ""
    End If
    
    Exit Function
    
handler:
    'Function will generally throw an error if text is empty or not in base64
    'Pass back text as is and throw a message to warn the user
    decryptText = text
    MsgBox ("Decryption of text failed! It may not have been encrypted")
    
    
End Function
