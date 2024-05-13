Option Explicit

Dim objShell, objFSO, objHTTP, strURL, strFile, strTempPath, strDesktopPath, strPuttyPath

' Create objects
Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")

' URL of the file to download
strURL = "https://the.earth.li/~sgtatham/putty/latest/w32/putty.exe"

' File name and path to save the downloaded file
strFile = "putty.exe"
strTempPath = objShell.ExpandEnvironmentStrings("%TEMP%") & "\" & strFile
strDesktopPath = objShell.SpecialFolders("Desktop") & "\" & strFile
strPuttyPath = "C:\Program Files\PuTTY\putty.exe" ' Change this to your PuTTY installation path

' Download the file to temporary folder
objHTTP.open "GET", strURL, False
objHTTP.send

If objHTTP.Status = 200 Then
    Dim objStream
    Set objStream = CreateObject("ADODB.Stream")
    objStream.Open
    objStream.Type = 1
    objStream.Write objHTTP.ResponseBody
    objStream.Position = 0
    objStream.SaveToFile strTempPath
    objStream.Close
    MsgBox "Download completed. File saved to temporary folder.", vbInformation
Else
    MsgBox "Failed to download file.", vbExclamation
End If

' Run PuTTY from temporary folder
If objFSO.FileExists(strTempPath) Then
    objShell.Run Chr(34) & strTempPath & Chr(34), 1, False
Else
    MsgBox "PuTTY executable not found in temporary folder.", vbExclamation
End If

' Move PuTTY back to its original location
If objFSO.FileExists(strTempPath) And objFSO.FolderExists("C:\Program Files\PuTTY\") Then
    objFSO.MoveFile strTempPath, strPuttyPath
    MsgBox "PuTTY executable moved back to its original location.", vbInformation
Else
    MsgBox "Failed to move PuTTY executable.", vbExclamation
End If

' Clean up objects
Set objShell = Nothing
Set objFSO = Nothing
Set objHTTP = Nothing
