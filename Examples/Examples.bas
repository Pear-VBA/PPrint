Attribute VB_Name = "Examples"
'@Folder "Examples"
Option Explicit

Public Sub Example()
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")

    Dim c As Collection
    Set c = New Collection

    c.Add "9"
    c.Add 8
    c.Add Array(7, 6, "5")

    d("number") = 1
    d("string") = "c"
    d("array") = Array(1, "2")
    Set d("collection") = c

    ppri_PPrint "number:", 1
    ppri_PPrint "string:", "Hello, World!"
    ppri_PPrint "array:", Array(1, "2")
    ppri_PPrint "collection:", c
    ppri_PPrint "dictionary:", d
    ppri_PPrint "range:", Range("A1:C1")
    ppri_PPrint "user class without repr__ method:", New Class1
    ppri_PPrint "user class with repr__ method:", New Class2
    ppri_PPrint "other objects:", ThisWorkbook, Worksheets, CreateObject("Scripting.FileSystemObject")
End Sub
