Option Compare Database
Option Explicit

Public gobjRibbon As IRibbonUI

' Error 438:
' Public Function New_aeNumClass() As aeNumClass
' Ref: http://support.microsoft.com/kb/555159#top
' ====================================================================
' Author:   Peter F. Ennis
' Date:     March 31, 2011
' Comment:  Instantiation of PublicNotCreatable aeNumClass
' Requires: aeNumClass Instancing 2 - PublicNotCreatable
' ====================================================================
' Set New_aeNumClass = New aeNumClass
' End Function



' ==========================================================
' aeRibbon Callbacks
' ==========================================================
' Ref: http://msdn.microsoft.com/en-us/library/aa433869.aspx
' Reference for Access 2010: "Microsoft Office 14.0 Object Library"
' xmlns="http://schemas.microsoft.com/office/2009/07/customui"
' Reference for Access 2007: "Microsoft Office 12.0 Object Library"
' xmlns="http://schemas.microsoft.com/office/2006/01/customui"

Public Sub aeMyAddinInitialize(ribbon As IRibbonUI)
'    Callback name for XML "onLoad"
    Set gobjRibbon = ribbon
End Sub

' Button
Public Function aeNtryPoint(strControl As String, strAction As String)
'    Callback name for XML "onAction"

    Select Case strControl
        Case "btn1"
            aeClassInfo
  
        Case "btn2"
            aeListModules
  
        Case "btn3"
            aeRenumber
  
        Case "btn4"
            aeRunTheSearch
  
        Case Else
            MsgBox "Button """ & strControl & """ clicked", vbInformation, "aeNtryPoint"
    End Select

End Function

Private Function aeClassInfo(Optional strModuleName As Variant)

    On Error GoTo aeClassInfo_Error

    ShowVBA strModuleName
    MyClassInformation

    On Error GoTo 0
    Exit Function

aeClassInfo_Error:

    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure aeClassInfo of Module (_) _COMMANDS_"

End Function

Private Function aeListModules()
    ShowVBA "(_) _COMMANDS_"
    MyListModules
End Function

Private Function aeRenumber()
    ShowVBA "(_) _COMMANDS_"
    MyRenumber
End Function

Private Function aeRunTheSearch()
    DoCmd.OpenForm "Db_SearchThroughObjects"
End Function

Private Function MyClassInformation()
    Dim objNumClass As aeNumClass
    Set objNumClass = New aeNumClass
End Function

Private Sub ShowVBA(Optional strModuleName As Variant)
'    Ref: http://www.officekb.com/Uwe/Forum.aspx/excel-prog/165815/Open-VBA-Editor

    On Error GoTo ShowVBA_Error

    If Application.VBE.MainWindow.visible = False Then
        Application.VBE.MainWindow.visible = True
    End If

    If IsMissing(strModuleName) Then
'        Ref: http://msdn.microsoft.com/en-us/library/aa443989(v=vs.60).aspx
        Set Application.VBE.ActiveCodePane = Application.VBE.CodePanes(1)
        Application.VBE.ActiveCodePane.Show
    Else
        Application.VBE.ActiveVBProject.VBComponents(strModuleName).CodeModule.CodePane.Show
    End If
    
    On Error GoTo 0
    Exit Sub

ShowVBA_Error:

    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure ShowVBA of Module (_) _COMMANDS_"

End Sub

Private Function MySkipModule()

    Dim objNumClass As aeNumClass
    Set objNumClass = New aeNumClass
    Dim str As String

    str = objNumClass.aeSkipModule
    Debug.Print "MySkipModule:aeSkipModule>" & objNumClass.aeSkipModule

End Function

Private Function MyClassName()

    Dim objNumClass As aeNumClass
    Set objNumClass = New aeNumClass

    Dim str As String

    str = objNumClass.aeMyClassName
    Debug.Print "MyClassName:str>" & str

End Function

Private Function MyRenumber()

    Dim objNumClass As aeNumClass
    Set objNumClass = New aeNumClass

    Dim bln As Boolean

    bln = objNumClass.aeReNumAutoIndent

End Function

Private Function MyListModules()

    Dim objNumClass As aeNumClass
    Set objNumClass = New aeNumClass

    Dim astr() As String

    objNumClass.aeDebugIt = True
    astr = objNumClass.aeGetListCodeModules

End Function

Private Sub TestAll()

    Dim objNumClass As aeNumClass
    Set objNumClass = New aeNumClass

    Dim str As String
    str = objNumClass.aeSkipModule
    Debug.Print "TestAll MySkipModule:aeSkipModule>" & str

    str = objNumClass.aeMyClassName
    Debug.Print "TestAll MyClassName:str>" & str

'    Exit Sub

    Dim bln As Boolean
    bln = objNumClass.aeReNumAutoIndent
    Debug.Print "TestAll objNumClass.aeReNumAutoIndent Finished"

    Dim lng As Long
    lng = objNumClass.aeCountLines
    Debug.Print "TestAll objNumClass.aeCountLines Finished"

    Dim astr() As String
    astr = objNumClass.aeGetListCodeModules
    objNumClass.aeDebugIt = True
    astr = objNumClass.aeGetListCodeModules
    Debug.Print "TestAll objNumClass.aeGetListCodeModules Finished"

End Sub

Public Function MyRemoveLineNumbers()

    Dim objNumClass As aeNumClass
    Set objNumClass = New aeNumClass

    Dim bln As Boolean

    bln = objNumClass.aeRemoveLineNums

End Function