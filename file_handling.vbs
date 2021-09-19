' [filesystem object reference](https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/filesystemobject-object)
Sub createFile()
    Dim a As String
    'todo: get the variable short name for home in windows
    a = "C:\Users\ashutoshsingh001\Desktop\hello.xlsx"

    Dim fs
    ' object creation in vba
    Set fs = CreateObject("Scripting.FileSystemObject")
    fs.CreateTextFile a

End Sub


' todo: work on File object and Folder object
' [File object](https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/file-object)
' [Folder object](https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/folder-object)