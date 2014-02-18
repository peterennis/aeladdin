Option Compare Database
Option Explicit

' RESEARCH
' Using the VBA Extensibility Library = Ref: http://pubs.logicalexpressions.com/pub0009/LPMArticle.asp?ID=307
' Ref: http://www.everythingaccess.com/mdeprotector_example.htm
' Ref: http://www.everythingaccess.com/vbwatchdog.htm
' Ref: http://www.officekb.com/Uwe/Forum.aspx/excel-prog/69455/Print-VBA-code-in-Editor-format-colors
' Ref: http://www.programmerworld.net/resources/visual_basic/visual_basic_addin.php
' Comment Out VBA Code Blocks = Ref: http://www.ozgrid.com/forum/showthread.php?t=10432&page=1
'

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' There is source code here on how to make an extended search
' http://kandkconsulting.tripod.com/VB/VBCode.htm##AddIN

' Ref: http://www.vbforums.com/showthread.php?t=479449
' Made By Michael Ciurescu (CVMichael)
'
' Modifications by Peter F. Ennis for aeladdin(TM)

Private mstrPotentialComment As String
Private maStrModulesList() As String
Private mstrAllCodeObjectsList() As String

Private BlockStart() As String
Private BlockEnd() As String
Private BlockMiddle() As String

Public Sub TestMyReNumberLeftAndAutoIndent(strSkipModule As String)
'    TestMyReNumberLeftAndAutoIndent "(TEST) _COMMANDS_OLD_2"
'
'    Do not renumber the module with the renumber function (i.e. this one)
'    It will crash Access.
'    A copy of the function is in modules "(_) aeNum"
    
       On Error GoTo TestMyReNumberLeftAndAutoIndent_Error
    
       Dim objVBAProject As Object
      Set objVBAProject = Application.VBE.ActiveVBProject
      MyReNumberAndLeftAlign objVBAProject, strSkipModule
      MyReNumberAndLeftAlign objVBAProject, strSkipModule
    
      Dim i As Integer
    
      mstrAllCodeObjectsList = aeListCodeModules(strSkipModule)
    
      Debug.Print "UBound(aeListCodeModules)=" & UBound(aeListCodeModules(strSkipModule))
      Debug.Print "UBound(mstrAllCodeObjectsList)=" & UBound(mstrAllCodeObjectsList)
    
      Debug.Print "TestMyReNumberLeftAndAutoIndent:"
    
      IndentInitialize
    
      For i = 1 To UBound(mstrAllCodeObjectsList)
          If mstrAllCodeObjectsList(i) <> strSkipModule Then
              Debug.Print i, mstrAllCodeObjectsList(i)
              MyIndentCodeModule mstrAllCodeObjectsList(i)
'            Sleep 3000
          End If
      Next i
'    Debug.Print "First time"
      EmptyLines objVBAProject, strSkipModule
'    Second time should have no output in immediate window when testing
'    Debug.Print "Second time"
'    EmptyLines objVBAProject, strSkipModule
    
      On Error GoTo 0
    Exit Sub
    
TestMyReNumberLeftAndAutoIndent_Error:
    
      MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure TestMyReNumberLeftAndAutoIndent of Module (TEST) _COMMANDS_OLD_2"
    
End Sub

Private Sub TestRemoveLineNumbers(strSkipModule As String)
'    TestRemoveLineNumbers "(_) _COMMANDS_OLD_2"
    
       Dim objVBAProject As Object
       Set objVBAProject = Application.VBE.ActiveVBProject
    RemoveLineNumbers objVBAProject, strSkipModule
    
End Sub

Private Sub Test_aeOLDListCodeModules()
    
       Dim i As Integer
       Dim strSkipModule As String
       strSkipModule = "(TEST) _COMMANDS_OLD_2"
    
       mstrAllCodeObjectsList = aeListCodeModules(strSkipModule)
    
       Debug.Print "UBound(mstrAllCodeObjectsList)=" & UBound(mstrAllCodeObjectsList)
    
      For i = 1 To UBound(mstrAllCodeObjectsList)
          Debug.Print mstrAllCodeObjectsList(i)
      Next i
    
End Sub

Private Sub TestMyIndentCodeModule(strTheModuleName As String)
'    Ref: http://www.cpearson.com/excel/vbe.aspx
'    Set a reference to Microsoft Visual Basic for Applications Extensibility Library 5.3
    
       On Error GoTo TestMyIndentCodeModule_Error
    
       Dim objVBComp As VBIDE.VBComponent
       Dim CodeMod As VBIDE.CodeModule
       Dim objVBAProject As Object
    
      Set objVBAProject = Application.VBE.ActiveVBProject
      Set objVBComp = Application.VBE.ActiveVBProject.VBComponents.Item(strTheModuleName)
      Set CodeMod = objVBComp.CodeModule
    
      Debug.Print , objVBComp.Name
    
      IndentInitialize
      IndentCodeModule objVBAProject, CodeMod
    
      On Error GoTo 0
    Exit Sub
    
TestMyIndentCodeModule_Error:
    
      MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure TestMyIndentCodeModule of Module (TEST) _COMMANDS_OLD_2"
    
End Sub

Private Sub TestMyReNumberAndLeftAlign()
'    Do not renumber the module with the renumber function (i.e. this one)
'    It will crash Access.
'    A copy of the function is in modules "(_) aeNum"
    
       Dim objVBAProject As Object
       Set objVBAProject = Application.VBE.ActiveVBProject
       MyReNumberAndLeftAlign objVBAProject, "(TEST) _COMMANDS_OLD_2"
    
End Sub

Private Sub TestSplitLineAtLastApostrophe()
    
       Dim str As String
    
       str = "test it ' here"
    
       If SplitLineAtLastApostrophe(str, mstrPotentialComment) Then
           Debug.Print "TestSplitLineAtLastApostrophe: str>" & str & "<", "mstrPotentialComment>" & mstrPotentialComment & "<"
       End If
    
End Sub

Private Function zzzaeListModules(Optional varDebug As Variant) As String()
'    Ref: http://www.exceltip.com/st/Array_variables_using_VBA_in_Microsoft_Excel/509.html
'    ====================================================================
'    Author:   Peter F. Ennis
'    Date:     March 17, 2011
'    Comment:  Add explicit references for DAO
'    ====================================================================
'
    
       Dim dbs As DAO.Database
      Dim cnt As DAO.Container
      Dim doc As DAO.Document
      Dim i As Integer
      Dim j As Integer
      Dim astr() As String
      Dim blnDebug As Boolean
    
      On Error GoTo aeListModules_Error
    
      If IsMissing(varDebug) Then
          blnDebug = False
      Else
          blnDebug = True
      End If
    
'    Use CurrentDb() to refresh Collections
      Set dbs = CurrentDb()
    
      Set cnt = dbs.Containers("Modules")
      j = cnt.Documents.Count
'    Debug.Print "cnt.Documents.Count=" & j
      ReDim astr(1 To j)
      i = 1
      For Each doc In cnt.Documents
          If Not (Left(doc.Name, 3) = "zzz") Then
'            Application.SaveAsText aTheCodeModule, doc.Name, aegitType.SourceFolder & "\" & doc.Name & ".bas"
'            Debug.Print i, doc.Name
              astr(i) = doc.Name
              i = i + 1
          End If
      Next doc
    
'    For i = 1 To UBound(aStr)
'    Debug.Print aStr(i),
'    Next
    
      zzzaeListModules = astr
    
      Set doc = Nothing
      Set cnt = Nothing
      Set dbs = Nothing
    
      On Error GoTo 0
    Exit Function
    
aeListModules_Error:
    
      MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure aeListModules of Module (_) aeNum"
    
End Function

Private Function SplitLineAtLastApostrophe(ByVal str As String, ByVal strRestOfLine As String) As Boolean
'    Ref: http://vbadud.blogspot.com/2008/11/how-to-return-multiple-values-from-vba.html
'    InStr Ref: http://www.ozgrid.com/forum/showthread.php?t=19252&page=1
'    NOTE: Parsing a code line to get an inline comment is sophisticated
'    This routine returns the string including the last apostrophe in mstrPotentialComment
'    Peter F. Ennis, March 2011
    
       Dim i As Integer
       Dim Pos As Integer
    
'    Debug.Print "str=" & str
      Pos = InStr(str, "'")
      If Pos = 0 Then
          strRestOfLine = ""
          SplitLineAtLastApostrophe = False
      ElseIf Pos = 2 Then            ' A space is forced in the first character position
          strRestOfLine = ""
          SplitLineAtLastApostrophe = False
      Else
          mstrPotentialComment = Mid(str, Pos, Len(str))
'        Debug.Print "SplitLineAtLastApostrophe: InStr(1, mstrPotentialComment, strRestOfLine)>" & InStr(1, mstrPotentialComment, strRestOfLine) & "<"
          If InStr(1, mstrPotentialComment, strRestOfLine) > 0 Then
              SplitLineAtLastApostrophe = True
          Else
              SplitLineAtLastApostrophe = True
          End If
'        Debug.Print "SplitLineAtLastApostrophe: str>" & str & "<", "Pos>" & Pos & "< ", "mstrPotentialComment>" & mstrPotentialComment & " < """
      End If
    
End Function

Private Sub TestKeywordInComment()
    
       Dim str As String
       Dim bln As Boolean
    
       str = "    x = 1             ' i do declare    "
       Debug.Print "TestKeywordInComment: str>" & str & "<"
       bln = KeywordInComment(str, "Declare")
       Debug.Print "TestKeywordInComment: bln>" & bln & "<", "mstrPotentialComment>" & mstrPotentialComment & "<"
    
End Sub

Private Function KeywordInComment(str As String, strKeyWord As String) As Boolean
    
       Dim bln As Boolean
    
       bln = SplitLineAtLastApostrophe(str, strKeyWord)
    
       If bln Then
           KeywordInComment = True
       Else
           KeywordInComment = False
      End If
    
'    Debug.Print "KeywordInComment>" & KeywordInComment & "<", "str>" & str & "<", "strKeyWord>" & strKeyWord & "<"
    
End Function

Private Function VBComponentExists(VBCompName As String, Optional VBProj As VBIDE.VBProject = Nothing) As Boolean
'    Ref: http://www.cpearson.com/excel/vbe.aspx
'    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    This returns True or False indicating whether a VBComponent named
'    VBCompName exists in the VBProject referenced by VBProj. If VBProj
'    is omitted, the VBProject of the Active Application is used.
'    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       Dim VBP As VBIDE.VBProject
       If VBProj Is Nothing Then
           Set VBP = Application.VBE.ActiveVBProject
      Else
          Set VBP = VBProj
      End If
      On Error Resume Next
      VBComponentExists = CBool(Len(VBP.VBComponents(VBCompName).Name))
    
End Function

Private Function aeListCodeModules(strSkipModule As String) As String()
'    Ref: http://www.exceltip.com/st/Array_variables_using_VBA_in_Microsoft_Excel/509.html
'    ====================================================================
'    Author:   Peter F. Ennis
'    Date:     March 17, 2011
'    Comment:  Add explicit references for DAO
'    ====================================================================
'
    
       Dim objAccess As AccessObject  'Each module/form/report.
'    Dim dbs As DAO.Database
'    Dim cnt As DAO.Container
'    Dim doc As DAO.Document
      Dim i As Integer
      Dim j As Integer
      Dim astr() As String
'    Dim blnDebug As Boolean
    
      On Error GoTo aeListCodeModules_Error
    
'    If IsMissing(blnDebug) Then
'    bDebug = False
'    Else
'    bDebug = True
'    End If
    
'    Use CurrentDb() to refresh Collections
'    Set dbs = CurrentDb()
    
'    Set cnt = dbs.Containers("Modules")
      j = CurrentProject.AllModules.Count
'    Debug.Print "cnt.Documents.Count=" & j
      ReDim astr(1 To j)
      i = 1
      For Each objAccess In CurrentProject.AllModules
          If Not (Left(objAccess.Name, 3) = "zzz") And objAccess.Name <> strSkipModule Then
'            Application.SaveAsText aTheCodeModule, objAccess.Name, aegitType.SourceFolder & "\" & doc.Name & ".bas"
'            Debug.Print i, objAccess.Name
              astr(i) = objAccess.Name
              i = i + 1
          End If
      Next objAccess
    
      ReDim Preserve astr(1 To i)
    
'    3811              i = i + 1
      For Each objAccess In CurrentProject.AllForms
          If Not (Left(objAccess.Name, 3) = "zzz") Then
'            Check if a CBF module exists
              If VBComponentExists("Form_" & objAccess.Name) Then
                  ReDim Preserve astr(1 To i)
                  astr(i) = "Form_" & objAccess.Name
                  i = i + 1
              Else
              End If
          End If
      Next objAccess
    
      For Each objAccess In CurrentProject.AllReports
          If Not (Left(objAccess.Name, 3) = "zzz") Then
              If VBComponentExists("Report_" & objAccess.Name) Then
                  ReDim Preserve astr(1 To i)
                  astr(i) = "Report_" & objAccess.Name
                  i = i + 1
              Else
              End If
          End If
      Next objAccess
    
'    If aeDebugIt Then
'    Debug.Print , "aeListCodeModules: aeDebugIt>" & aeDebugIt
'    For i = 1 To UBound(astr)
'    Debug.Print , i, astr(i)
'    Next
'    End If
    
      aeListCodeModules = astr
    
'    Set doc = Nothing
'    Set cnt = Nothing
'    Set dbs = Nothing
    
      On Error GoTo 0
    Exit Function
    
aeListCodeModules_Error:
    
      MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure aeListCodeModules of Module (_) aeNum"
    
End Function

Private Function zzzaeAllCodeObjectsList(Optional varDebug As Variant) As String()
'    ====================================================================
'    Author:   Peter F. Ennis
'    Date:     March 26, 2011
'    Comment:
'    ====================================================================
    
       Dim objVBAProject As Object
       Set objVBAProject = Application.VBE.ActiveVBProject
    
      Dim objVBComponent As VBComponent
      Dim blnDebug As Boolean
      Dim i As Integer
      Dim astr() As String
    
      If IsMissing(varDebug) Then
          blnDebug = False
      Else
          blnDebug = True
      End If
    
      i = 1
      For Each objVBComponent In objVBAProject.VBComponents
          ReDim Preserve astr(1 To i)
          astr(i) = objVBComponent.Name
          i = i + 1
'        Debug.Print objVBComponent.Name
      Next
    
      zzzaeAllCodeObjectsList = astr
    
End Function

Private Sub zzzTest_aeListModules()
    
       Dim i As Integer
    
       maStrModulesList = zzzaeListModules
    
       For i = 1 To UBound(maStrModulesList)
           Debug.Print maStrModulesList(i)
       Next i
    
End Sub

Private Sub TestEmptyLines()
    
       Dim objVBAProject As Object
       Set objVBAProject = Application.VBE.ActiveVBProject
       EmptyLines objVBAProject, "(_) _COMMANDS_"
    
End Sub

Private Sub EmptyLines(oVBP As VBProject, Optional strModuleName As Variant)
'    Peter F. Ennis, March 2011
    
       Dim objVBComponent As VBComponent
       Dim strProcName As String
       Dim strOldName As String
       Dim strLine As String
       Dim iLine As Integer
       Dim lngBodyLine As Long
       Dim lngType As Long
      Dim blnContinue As Boolean
      Dim blnIsSelectCase As Boolean
      Dim i As Integer
      Dim strModName
    
      On Error GoTo EmptyLines_Error
    
'    Do not renumber strModuleName
      If IsMissing(strModuleName) Then
          strModName = ""
      Else
          strModName = strModuleName
      End If
    
      i = 0
      For Each objVBComponent In oVBP.VBComponents
          i = i + 1
'        Debug.Print i, "objVBComponent.Name=" & objVBComponent.Name
          If objVBComponent.CodeModule <> strModName Then
              With objVBComponent.CodeModule
                  strOldName = ""
                  lngBodyLine = 0
                
                  For iLine = 1 To .CountOfLines
                      strLine = .Lines(iLine, 1)
                    
                      strProcName = .ProcOfLine(iLine, lngType)
                    
                      If strProcName <> strOldName Then
                          lngBodyLine = .ProcBodyLine(strProcName, lngType)
                      End If
                    
                      If iLine > lngBodyLine And strProcName <> "" Then
'                        Debug.Print , "strLine>" & strLine & "<"
                        
                          If IsNumeric(Left(strLine, InStr(1, strLine, " "))) Or Left(strLine, 1) = "'" Or Right(strLine, 1) <> " " Then
                          Else
                              .ReplaceLine iLine, Trim(strLine)
'                            For second time testing
'                            Debug.Print iLine, objVBComponent.Name, "strLine>" & strLine & "<"
                          End If
                      End If
                    
                  Next
'                Sleep 1000
              End With
          End If
        
      Next
    
      On Error GoTo 0
    Exit Sub
    
EmptyLines_Error:
    
      MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure EmptyLines of Module (_) aeNum"
    
End Sub

Private Sub RemoveLineNumbers(oVBP As VBProject, Optional strModuleName As Variant)
'    Peter F. Ennis, March 2011
    
       Dim objVBComponent As VBComponent
       Dim strProcName As String
       Dim strOldName As String
       Dim strLine As String
       Dim iLine As Integer
       Dim lngBodyLine As Long
       Dim lngType As Long
      Dim blnContinue As Boolean
      Dim blnIsSelectCase As Boolean
      Dim strModName
      Dim Pos As Integer
    
      On Error GoTo RemoveLineNumbers_Error
    
'    Do not renumber strModuleName
      If IsMissing(strModuleName) Then
          strModName = ""
      Else
          strModName = strModuleName
      End If
    
      For Each objVBComponent In oVBP.VBComponents
          If objVBComponent.CodeModule <> strModName Then
              With objVBComponent.CodeModule
                  strOldName = ""
                  lngBodyLine = 0
                
                  For iLine = 1 To .CountOfLines
                      strLine = .Lines(iLine, 1)
                    
                      strProcName = .ProcOfLine(iLine, lngType)
                    
                      If strProcName <> strOldName Then
                          lngBodyLine = .ProcBodyLine(strProcName, lngType)
                      End If
                    
                      If iLine > lngBodyLine And strProcName <> "" Then
'                        Debug.Print , "strLine>" & strLine & "<"
                        
                          If IsNumeric(Left(strLine, InStr(1, strLine, " "))) Then
                              Pos = InStr(strLine, " ")
'                            If objVBComponent.Name = "(_) Renumber_Bugs_Testing" Then
'                            Debug.Print "iLine>" & iLine, "Pos>" & Pos, "strLine>" & strLine
'                            End If
                              .ReplaceLine iLine, Mid(strLine, Pos, Len(strLine) - Pos + 1)
                          End If
                      End If
                    
                  Next
'                Sleep 1000
              End With
          End If
        
      Next
    
      On Error GoTo 0
    Exit Sub
    
RemoveLineNumbers_Error:
    
      MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure RemoveLineNumbers of Module (_) aeNum"
    
End Sub

Private Sub MyReNumberAndLeftAlign(oVBP As VBProject, strModuleName As String)
'    Ref: https://groups.google.com/group/microsoft.public.excel.programming/browse_frm/thread/33b9b2b3cc22f9e3?hl=en&lr&pli=1
'    Set a reference to Microsoft Visual Basic for Applications Extensibility Library 5.3
'    Code example framework without GoTo
'    Ref: http://www.ozgrid.com/forum/showthread.php?t=120898&page=1
'    Ref: http://allenbrowne.com/vba-CountLines.html
'    Ref: http://sites.google.com/site/msaccesscode/vba-3/showerrorlineinvba
'
'    Peter F. Ennis, March 2011
'    Run it twice to take care of lines that initially have no numbers.
    
      Dim objVBComponent As VBComponent
      Dim strProcName As String
      Dim strOldName As String
      Dim strLine As String
      Dim iLine As Integer
      Dim lngBodyLine As Long
      Dim lngType As Long
      Dim blnContinue As Boolean
      Dim blnIsSelectCase As Boolean
'    Dim i As Integer
'    Dim strModName
    
      On Error GoTo ReNumberAndLeftAlign_Error
    
'    Debug.Print , "MyRenumberAndLeftAlign"
    
'    Do not renumber strModuleName
'    28      If IsMissing(strModuleName) Then
'    29          strModName = ""
'    30      Else
'    31          strModName = strModuleName
'    32      End If
    
'    i = 0
      For Each objVBComponent In oVBP.VBComponents
'        i = i + 1
'        Debug.Print i, "objVBComponent.Name=" & objVBComponent.Name
          If objVBComponent.CodeModule <> strModuleName Then
              With objVBComponent.CodeModule
                  strOldName = ""
                  lngBodyLine = 0
                
                  For iLine = 1 To .CountOfLines
                      strLine = Trim(.Lines(iLine, 1)) & " "
                    
'                    If objVBComponent.Name = "(_) Renumber_Bugs_Testing" Then
'                    Debug.Print "A:" & iLine & ", " & lngBodyLine, "strLine>" & strLine
'                    End If
                    
                      If IsNumeric(Left(strLine, InStr(1, strLine, " "))) Then
                        
'                        Debug.Print iLine & " " & objVBComponent.Name & ">", strLine
'                        Debug.Print Left(strLine, InStr(1, strLine, " "))
'                        Debug.Print Mid(strLine, InStr(1, strLine, " "))
                          .ReplaceLine iLine, Trim(Mid(strLine, InStr(1, strLine, " ")))
'                        If objVBComponent.Name = "(_) Renumber_Bugs_Testing" Then
'                        Debug.Print "B:" & iLine & ", " & lngBodyLine, Trim(Mid(strLine, InStr(1, strLine, " ")))
'                        End If
                      End If
                    
                      strProcName = .ProcOfLine(iLine, lngType)
'                    If objVBComponent.Name = "(_) Renumber_Bugs_Testing" Then
'                    Debug.Print "strProcName=" & strProcName, "iLine=" & iLine, "lngType=" & lngType, ".ProcOfLine(iLine, lngType)=" & .ProcOfLine(iLine, lngType)
'                    End If
                      If strProcName <> strOldName Then
                          lngBodyLine = .ProcBodyLine(strProcName, lngType)
                      End If
                    
                      If iLine > lngBodyLine And strProcName <> "" Then
                        
                        
'                        If objVBComponent.Name = "(_) Renumber_Bugs_Testing" Then
'                        Debug.Print "C:" & iLine & ", " & lngBodyLine, "strLine>" & strLine
'                        Debug.Print "C1:" & iLine & ", " & lngBodyLine, "strLine>" & " " & Trim(.Lines(iLine, 1)) & " "
'                        End If
                        
                        
                          strLine = " " & Trim(.Lines(iLine, 1)) & " "
                          If strLine = "  " Then GoTo ptr_Next
                          If Left(strLine, 2) = " '" Then GoTo ptr_Next
                          If Left(strLine, 2) = " #" Then GoTo ptr_Next
                          If Left(strLine, 4) = " Rem" Then GoTo ptr_Next
'
                          If InStr(1, strLine, " = Split") > 0 And InStr(1, strLine, " strLine,") = 0 Then GoTo ptr_Continue
'                        If objVBComponent.Name = "(_) Renumber_Bugs_Testing" Then
                          If InStr(1, strLine, " Declare ") > 0 And InStr(1, strLine, " strLine,") = 0 And KeywordInComment(strLine, "Declare") Then GoTo ptr_Continue
                          If InStr(1, strLine, " Sub ") > 0 And InStr(1, strLine, " strLine,") = 0 And KeywordInComment(strLine, "Sub") Then GoTo ptr_Continue
                          If InStr(1, strLine, " Function ") > 0 And InStr(1, strLine, " strLine,") = 0 And KeywordInComment(strLine, "Function") Then GoTo ptr_Continue
                          If InStr(1, strLine, " Property ") > 0 And InStr(1, strLine, " strLine,") = 0 And KeywordInComment(strLine, "Property") Then GoTo ptr_Continue
'                        End If
'
                          If InStr(1, strLine, " Declare ") > 0 And InStr(1, strLine, " strLine,") = 0 Then GoTo ptr_Next
                          If InStr(1, strLine, " Sub ") > 0 And InStr(1, strLine, " strLine,") = 0 Then GoTo ptr_Next
                          If InStr(1, strLine, " Function ") > 0 And InStr(1, strLine, " strLine,") = 0 Then GoTo ptr_Next
                          If InStr(1, strLine, " Property ") > 0 And InStr(1, strLine, " strLine,") = 0 Then GoTo ptr_Next
                        
ptr_Continue:
                          strLine = .Lines(iLine, 1) & " "
                          If IsNumeric(Left(strLine, 4)) Then strLine = Mid(strLine, 6)
                         If IsNumeric(Left(strLine, InStr(1, strLine, " ") - 1)) Then strLine = Mid(strLine, InStr(1, strLine, " "))
                         If Trim(strLine) = "" Then
                             strLine = Trim(strLine)
                         Else
                             strLine = Format(iLine - lngBodyLine, "0000") & " " & strLine
                         End If
'                        If objVBComponent.Name = "(_) Renumber_Bugs_Testing" Then
'                        Debug.Print "blnContinue=" & blnContinue, "blnIsSelectCase=" & blnIsSelectCase, iLine, strLine
'                        End If
                         If Not blnContinue And Not blnIsSelectCase Then .ReplaceLine iLine, strLine
                         If InStr(1, strLine, "Select Case") > 0 And InStr(1, strLine, " = Split(") = 0 Then
                             blnIsSelectCase = True
                         ElseIf InStr(1, strLine, "Case") > 0 And InStr(1, strLine, " = Split(") = 0 Then
                             blnIsSelectCase = False
                         ElseIf InStr(1, strLine, "End Select") > 0 And InStr(1, strLine, " = Split(") = 0 Then
                             blnIsSelectCase = False
                         End If
                     End If
ptr_Next:
                     blnContinue = (Right(Trim(strLine), 2) = " _")
                 Next
'                Sleep 1000
             End With
         End If
        
     Next
    
     On Error GoTo 0
    Exit Sub
    
ReNumberAndLeftAlign_Error:
    
     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ReNumberAndLeftAlign of Module (_) aeNum"
    
End Sub

Private Sub IndentInitialize()
       BlockStart = Split("If * Then;For *;Do *;Do;Select Case *;While *;With *;Private Function *;Public Function *;Friend Function *;Function *;Private Sub *;Public Sub *;Friend Sub *;Sub *;Private Property *;Public Property *;Friend Property *;Property *;Private Enum *;Public Enum *;Friend Enum *;Enum *;Private Type *;Public Type *;Friend Type *;Type *", ";")
       BlockEnd = Split("End If;Next*;Loop;Loop *;End Select;Wend;End With;End Function;End Function;End Function;End Function;End Sub;End Sub;End Sub;End Sub;End Property;End Property;End Property;End Property;End Enum;End Enum;End Enum;End Enum;End Type;End Type;End Type;End Type", ";")
       BlockMiddle = Split("ElseIf * Then;Else;Case *", ";")
End Sub

Private Sub MyIndentCodeModule(strTheModuleName As String)
'    Ref: http://www.cpearson.com/excel/vbe.aspx
'    Set a reference to Microsoft Visual Basic for Applications Extensibility Library 5.3
    
       On Error GoTo MyIndentCodeModule_Error
    
       Dim objVBComp As VBIDE.VBComponent
       Dim CodeMod As VBIDE.CodeModule
    
       Dim objVBAProject As Object
      Set objVBAProject = Application.VBE.ActiveVBProject
      Set objVBComp = Application.VBE.ActiveVBProject.VBComponents.Item(strTheModuleName)
      Set CodeMod = objVBComp.CodeModule
    
'    Debug.Print , objVBComp.Name
    
      IndentCodeModule objVBAProject, CodeMod
    
      On Error GoTo 0
    Exit Sub
    
MyIndentCodeModule_Error:
    
      MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure MyIndentCodeModule of Module (TEST) _COMMANDS_OLD_2"
    
End Sub

Private Sub IndentCodeModule(oVBP As VBProject, TheCodeModule As Object)
    
'    110      Dim objVBComponent As VBComponent
'    38          Set objVBComponent = TheCodeModule
    
       On Error GoTo IndentCodeModule_Error
    
       Dim AllCode As String
       Dim AllLines() As String
       Dim LineNumbers() As String
      Dim k As Long
      Dim Q As Long
      Dim VBLine As String
      Dim StartCursorPos As Long
      Dim TopLine As Long
    
      With TheCodeModule
          TopLine = .CodePane.TopLine
          .CodePane.GetSelection StartCursorPos, 0, 0, 0
        
          AllCode = .Lines(1, .CountOfLines)
        
          AllCode = Replace(AllCode, "_" & vbNewLine, "")
          AllLines = Split(AllCode, vbNewLine)
          ReDim LineNumbers(UBound(AllLines))
        
          For k = 0 To UBound(AllLines)
              Q = InStr(1, AllLines(k), " ")
              If Q = 0 Then Q = InStr(1, AllLines(k), vbTab)
            
              If Q > 0 Then
                  If CStr(Val(Left(AllLines(k), Q - 1))) = Left(AllLines(k), Q - 1) Then
                      LineNumbers(k) = Left(AllLines(k), Q - 1)
                      AllLines(k) = Mid(AllLines(k), Q)
                  End If
              End If
            
              AllLines(k) = Trim(AllLines(k))
            
              If Left(AllLines(k), 1) = "'" Then
                  AllLines(k) = "' " & Trim(Mid(AllLines(k), 2))
              End If
          Next k
        
          IndentBlock AllLines
        
          For k = 0 To UBound(AllLines)
              VBLine = RemoveLineComments(AllLines(k))
            
              For Q = 0 To UBound(BlockMiddle)
                  If Replace(VBLine, vbTab, "") Like BlockMiddle(Q) And Left(VBLine, 1) = vbTab Then
                      AllLines(k) = Mid(AllLines(k), 2)
                  End If
              Next Q
          Next k
        
          For k = 0 To UBound(AllLines)
              If Len(LineNumbers(k)) > 0 Then
                  AllLines(k) = LineNumbers(k) & vbTab & AllLines(k)
              End If
          Next k
        
          AllCode = Replace(Join(AllLines, vbNewLine), Chr(1), "")
        
          .DeleteLines 1, .CountOfLines
          .InsertLines 1, AllCode
        
          TheCodeModule.CodePane.SetSelection StartCursorPos, 1, StartCursorPos, 1
          .CodePane.TopLine = TopLine
      End With
    
      On Error GoTo 0
    Exit Sub
    
IndentCodeModule_Error:
    
      MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure IndentCodeModule of Module (_) _COMMANDS_OLD_2"
    
End Sub

Private Sub IndentBlock(VBLines() As String)
    
       On Error GoTo IndentBlock_Error
    
       Dim k As Long
       Dim Q As Long
       Dim EndPos As Long
       Dim VBLine As String
       Dim StartPos As Long
       Dim FoundStartEnd As Boolean
    
      Do
          StartPos = 0
          EndPos = UBound(VBLines)
        
          Do
              FoundStartEnd = False
            
              For k = EndPos To StartPos Step -1
                  VBLine = RemoveLineComments(VBLines(k))
'                Debug.Print "VBLine=" & VBLine
                
                  If Len(VBLine) > 0 Then
                      For Q = 0 To UBound(BlockStart)
                          If VBLine Like BlockStart(Q) Then
                              StartPos = k + 1
                              FoundStartEnd = True
                              Exit For
                          End If
                      Next Q
                    
                      If Q <= UBound(BlockStart) Then Exit For
                  End If
              Next k
            
              If FoundStartEnd Then
                  For k = StartPos To EndPos
                      VBLine = RemoveLineComments(VBLines(k))
                    
                      If Len(VBLine) > 0 Then
                          If VBLine Like BlockEnd(Q) Then
                              EndPos = k - 1
                              FoundStartEnd = True
                              Exit For
                          End If
                      End If
                  Next k
              End If
          Loop While FoundStartEnd
        
          If Not (Not FoundStartEnd And StartPos = 0 And EndPos = UBound(VBLines)) Then
'            Debug.Print StartPos, EndPos
              IndentLineBlock VBLines, StartPos, EndPos
          End If
      Loop Until Not FoundStartEnd And StartPos = 0 And EndPos = UBound(VBLines)
    
      On Error GoTo 0
    Exit Sub
    
IndentBlock_Error:
    
      MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure IndentBlock of Module (_) _COMMANDS_"
    
End Sub

Private Function RemoveLineComments(ByVal VBLine As String) As String
    
       On Error GoTo RemoveLineComments_Error
    
       Dim k As Long
       Dim Q As Long
    
       If InStr(1, VBLine, "'") > 0 Then
           k = 1
           Do
              If Mid(VBLine, k, 1) = """" Then
                  For Q = k + 1 To Len(VBLine)
                      If Mid(VBLine, Q, 1) = """" Then
                          VBLine = Left(VBLine, k) & Mid(VBLine, Q)
                          k = k + 1
                          Exit For
                      End If
                  Next Q
              End If
            
              k = k + 1
          Loop Until k >= Len(VBLine)
        
          For k = 1 To Len(VBLine)
              If Mid(VBLine, k, 1) = "'" Then
                  VBLine = Left(VBLine, k - 1)
                  Exit For
              End If
          Next k
      End If
    
    RemoveLineComments = Trim(VBLine)
    
      On Error GoTo 0
    Exit Function
    
RemoveLineComments_Error:
    
      MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure RemoveLineComments of Module (_) _COMMANDS_"
    
End Function

Private Sub IndentLineBlock(VBLines() As String, ByVal StartLine As Long, ByVal EndLine As Long, Optional NoIndentChar As Byte = 1)
    
       On Error GoTo IndentLineBlock_Error
    
       Dim k As Long
    
       If NoIndentChar > 0 Then
           If StartLine > 0 Then VBLines(StartLine - 1) = Chr(NoIndentChar) & VBLines(StartLine - 1)
           If EndLine < UBound(VBLines) Then VBLines(EndLine + 1) = Chr(NoIndentChar) & VBLines(EndLine + 1)
       End If
    
      For k = StartLine To EndLine
          If Left(VBLines(k), 1) = "'" Then
              VBLines(k) = "' " & vbTab & Mid(VBLines(k), 2)
          Else
              VBLines(k) = vbTab & VBLines(k)
          End If
      Next k
    
      On Error GoTo 0
    Exit Sub
    
IndentLineBlock_Error:
    
      MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure IndentLineBlock of Module (_) _COMMANDS_"
    
End Sub