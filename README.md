# Excel-VBA-Password-Manager-and-Changer
## Overview
Allows user to set and store a password/text in excel not readily accessible by users.
Perfect base for creating apps/workbooks with multiple users that are expected to have
different privileges without hardcoding a password in to the VBA code. Users with sufficient
privileges are able to change the password outside of VBA allowing separate protection for
the VBA code/module.

## Description
This module provides a set of functions that can be used inside VBA in conjuction with
Excel's protect feature to enable/disable certain functions or features, such as blocking
the editing of certain cells etc unless they are authorised to do so i.e. they know the
password.

## Main functions/calls
    getPassword()
    Returns the currently stored password in human readable format. If no password is
    currently stored, returns an empty string.
    
    changePassword()
    Calls the CredentialsFrm, allowing user to change/set the stored password for the
    workbook. Create a button and link it's call macro to this function if you want the
    user to be able to change/set the password. If you don't want the user to be able to
    change the password this procedure can be invoked manaully through the VBA editor.
    
## Requirements
* Microsoft XML v3/4 - For encrypt feature

## Additional Notes
* This module should be implemented on a workbook where no passwords have been set yet
* The user should not be setting protection outside of VBA eg. manually setting a password
for a given sheet etc.
* Lock this module before sending it to client/users

## Special Note
Excel protection is a bit of a joke as it can be bypassed fairly easily without
a password. Therefore it is pointless to perform complicated encryption on the password i.e.
if the user is smart enough to figure out the storage location and unjumble the password then
they most likely are capable of bypassing Excel's protection all together, assuming you have
locked this module and they are not aware of the inner workings.
