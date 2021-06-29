Attribute VB_Name = "Module1"
Function MoveFiles()

Dim index As Long
Dim FileName As String
Dim extension As String
Dim sourceFolder As String
Dim targetFolder As String

extension = ".png"
sourceFolder = "/Users/sheenapatel/Documents/Projects/Blindness Detection/TRAINING_images"

For index = 2 To Rows.Count
'If statement for No DR
If Cells(index, 2).Value = 0 Then
FileName = Cells(index, 1).Value

FileName = FileName & extension
targetFolder = "/Users/sheenapatel/Documents/Projects/Blindness Detection/TRAINING_images/No_DR"


'If statement for Mild DR
ElseIf Cells(index, 2).Value = 1 Then
FileName = Cells(index, 1).Value

FileName = FileName & extension
targetFolder = "/Users/sheenapatel/Documents/Projects/Blindness Detection/TRAINING_images/Mild_DR"


'If statement for Moderate DR
ElseIf Cells(index, 2).Value = 2 Then
FileName = Cells(index, 1).Value

FileName = FileName & extension
targetFolder = "/Users/sheenapatel/Documents/Projects/Blindness Detection/TRAINING_images/Moderate_DR"


'If statement for Severe DR
ElseIf Cells(index, 2).Value = 3 Then
FileName = Cells(index, 1).Value

FileName = FileName & extension
targetFolder = "/Users/sheenapatel/Documents/Projects/Blindness Detection/TRAINING_images/Severe_DR"



'If statement for Proliferative DR
ElseIf Cells(index, 2).Value = 4 Then
FileName = Cells(index, 1).Value

FileName = FileName & extension
targetFolder = "/Users/sheenapatel/Documents/Projects/Blindness Detection/TRAINING_images/Proliferative_DR"

Else: Debug.Print "Could not find file for some reason"
End If

Call sbCopyingAFile(FileName, targetFolder)

Next index
End Function





'In this Example I am Copying the File From "C:Temp" Folder to "D:Job" Folder
Sub sbCopyingAFile(FileName As String, targetFolder As String)

'Declare Variables
Dim sFile As String
Dim sDFolder As String

'This is Your File Name which you want to Copy
sFile = "/Users/sheenapatel/Documents/Projects/Blindness Detection/TRAINING_images/" & FileName


'Change to match the source folder path
sSFolder = "/Users/sheenapatel/Documents/Projects/Blindness Detection/TRAINING_images"

'Change to match the destination folder path
sDFolder = targetFolder & "/" & FileName

Debug.Print sFile & " " & sDFolder

FileCopy sFile, sDFolder
Debug.Print FileName & " " & "copied successfully to targetfolder"

'Checking If File Is Located in the Source Folder
'If Not FSO.FileExists(sSFolder) Then
   ' MsgBox "Specified File Not Found", vbInformation, "Not Found"
    
'Copying If the Same File is Not Located in the Destination Folder
'ElseIf Not FSO.FileExists(sDFolder & sFile) Then
    'FSO.CopyFile (sSFolder & sFile), sDFolder, True
    'MsgBox "Specified File Copied Successfully", vbInformation, "Done!"
'Else
    'MsgBox "Specified File Already Exists In The Destination Folder", vbExclamation, "File Already Exists"
'End If

End Sub
