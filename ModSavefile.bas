Attribute VB_Name = "ModSavefile"
Sub SaveFileAsUWNotes()

Dim FileName As String
    FileName = Selection.Text

CleanName (FilePath)
    Debug.Print "String Selection: " & FileName

'split file name and directory
Dim DirectoryPath As String
    DirectoryPath = Left(FileName, InStrRev(FileName, "\") - 1)
    Debug.Print "Directory Path: " & DirectoryPath

'Making a directory if it doesn't exist
Dim elm As Variant
Dim x As String
    For Each elm In Split(DirectoryPath, "\")
        x = x & elm & "\"
        Debug.Print "Checking if exists: " & x
        If Len(Dir(x, vbDirectory)) = 0 Then
            Debug.Print "Making directory: " & x
            MkDir x
        End If
        
    Next

'On Error GoTo ErrHand:

Debug.Print "File saved as: " & FileName
ActiveDocument.SaveAs FileName:= _
        Left(FileName, InStrRev(FileName, ".") - 1) _
        , FileFormat:=wdFormatXMLDocument, LockComments:=False, Password:="", _
        AddToRecentFiles:=True, WritePassword:="", ReadOnlyRecommended:=False, _
        EmbedTrueTypeFonts:=False, SaveNativePictureFormat:=False, SaveFormsData _
        :=False, SaveAsAOCELetter:=False

Exit Sub

ErrHand:
    
    MsgBox "Error: " & Err.Number & " - " & Err.Description

End Sub

Function CleanName(strName As String) As String
'will clean part # name so it can be made into valid folder name
'may need to add more lines to get rid of other characters

    CleanName = Replace(strName, "/", "")
    CleanName = Replace(CleanName, "*", "")
    CleanName = Replace(CleanName, Chr(10), "")
    

End Function

