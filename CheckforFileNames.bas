Attribute VB_Name = "Module1"
Sub ListAllFilesInFolder()
    Dim folderPath As String
    Dim fileName As String

    ' Specify the folder where your files are located
    folderPath = "\\siwdsntv002\SG_PSC_SG1_PL_08_Control_WHse\Daily Tank Reading\Tanker reading year 2024\Sep 24\" ' Ensure this path is correct!
    
    ' Check if the folder path ends with a backslash
    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    
    ' Debugging: Check if folder path exists
    If Dir(folderPath, vbDirectory) = "" Then
        MsgBox "Folder does not exist: " & folderPath, vbCritical
        Exit Sub
    End If

    ' Loop through each file in the folder
    fileName = Dir(folderPath) ' List all files, regardless of extension
    
    Do While fileName <> ""
        ' Print each file name to the Immediate Window
        Debug.Print "File found: " & fileName
        
        ' Get the next file
        fileName = Dir
    Loop
End Sub
