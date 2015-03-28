# ============================================================================
# Author: Bill Reed 2013
# Powershell: Mine remote computers for PST files
# ============================================================================
strComputer = Get-Content -Path "C:\computernames.txt"
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
# ============================================================================
# Search entire remote computer
# ============================================================================
Set colFiles = objWMIService.ExecQuery _
    ("Select * from CIM_DataFile Where Extension = 'pst'")
# ============================================================================
# Search only C: of remote computer
# ============================================================================
# Set colFiles = objWMIService.ExecQuery _
#    ("Select * from CIM_DataFile Where Extension = 'pst' AND Drive = 'C:'")
# ============================================================================
If colFiles.Count = 0 Then
    Wscript.Quit
End If
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objTextFile = objFSO.CreateTextFile("C:\Scripts\PST_Data.txt")
For Each objFile in colFiles
    objTextFile.Write(objFile.Drive & objFile.Path & ",") # Path 
    objTextFile.Write(objFile.FileName & "." & objFile.Extension & ",") # File name
    objTextFile.Write(objFile.FileSize & vbCrLf) # File size
Next
objTextFile.Close
