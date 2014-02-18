Option Compare Database
Option Explicit

Public Const THE_SOURCE_FOLDER = "C:\ae\aeladdin\aerc\src\"
Public Const THE_XML_FOLDER = "C:\ae\aeladdin\aerc\src\xml\"
'

Public Sub AELADDIN_TEST()

    On Error GoTo PROC_ERR
    'aegitClassTest varSrcFldr:=THE_SOURCE_FOLDER, varXmlFldr:=THE_XML_FOLDER
    aegitClassTest Debugit:="Debugit", varSrcFldr:=THE_SOURCE_FOLDER, varXmlFldr:=THE_XML_FOLDER

PROC_EXIT:
    Exit Sub

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure AELADDIN_TEST"
    Resume Next

End Sub