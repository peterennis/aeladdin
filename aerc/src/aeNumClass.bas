Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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

Private Const AUTHOR As String = "Peter F. Ennis"
Private Const COMPANY As String = "adaept information management"
Private Const VERSION As String = "0.0.5"
Private Const VERSION_DATE As String = "April 11, 2011"
' 20110404 v003 Add DebugIt
' 20110405 v004 Test for VBComponent
' 20110411 v005 VBAversion, Office32or64


Private mstrPotentialComment As String
Private maStrModulesList() As String
Private mstrAllCodeObjectsList() As String

Private BlockStart() As String
Private BlockEnd() As String
Private BlockMiddle() As String

Private Type mySetupType
    SkipModule As String
    MyClassName As String
    Debugit As Boolean
End Type

Private aeNumType As mySetupType
'

Property Get aeCountLines() As Long
1       aeCountLines = CountLines(3)
2       Debug.Print "Get aeCountLines>" & aeCountLines
End Property

Property Get aeDebugIt() As Boolean
1       aeDebugIt = aeNumType.Debugit
End Property

Property Let aeDebugIt(bln As Boolean)
1       aeNumType.Debugit = bln
End Property

Property Get aeGetListCodeModules() As String()
1       aeGetListCodeModules = aeListCodeModules(aeNumType.SkipModule)
End Property

Property Get aeMyClassName() As String
1       aeMyClassName = aeNumType.MyClassName
2       Debug.Print "Get aeMyClassName>" & aeMyClassName
End Property

Property Get aeOffice32or64() As String
1       aeOffice32or64 = Office32or64
2       Debug.Print "Get aeOffice32or64>" & Office32or64
End Property

Property Get aeRemoveLineNums() As Boolean
1       TestRemoveLineNumbers aeNumType.SkipModule
End Property

Property Get aeReNumAutoIndent() As Boolean
1       Debug.Print "Get aeReNumAutoIndent:aeNumType.SkipModule>" & aeNumType.SkipModule
2       TestMyReNumberLeftAndAutoIndent aeNumType.SkipModule
End Property

Property Get aeSkipModule() As String
1       aeSkipModule = aeNumType.SkipModule
2       Debug.Print "Get aeSkipModule>" & aeSkipModule
End Property

Property Get aeVBAversion() As String
1       aeVBAversion = VBAversion
2       Debug.Print "Get aeVBAversion>" & VBAversion
End Property

Private Sub Class_Initialize()
'    Ref: http://www.cadalyst.com/cad/autocad/programming-with-class-part-1-5050
'    Ref: http://www.bigresource.com/Tracker/Track-vb-cyJ1aJEyKj/
'    Ref: http://stackoverflow.com/questions/1731052/is-there-a-way-to-overload-the-constructor-initialize-procedure-for-a-class-in
    
5       If SysCmd(acSysCmdAccessVer) <> "12.0" And SysCmd(acSysCmdAccessVer) <> "14.0" Then     ' "12.0" is Access 2007, "14.0" is Access 2010
6           MsgBox "This add-in is for Access 2007 or 2010 only !!!", vbCritical, "adaept"
        Exit Sub
8       End If
9       aeNumType.SkipModule = "(_) _COMMANDS_"
10      aeNumType.MyClassName = "aeNumClass"
11      aeNumType.Debugit = False
'    Show default values
13      Debug.Print , ">Class_Initialize"
14      Debug.Print , ">AUTHOR = " & AUTHOR
15      Debug.Print , ">COMPANY = " & COMPANY
16      Debug.Print , ">VERSION = " & VERSION
17      Debug.Print , ">VERSION_DATE = " & VERSION_DATE
18      Debug.Print , ">aeMyClassName = " & aeNumType.MyClassName
19      Debug.Print , ">aeSkipModule = " & aeNumType.SkipModule
20      Debug.Print , ">aeDebugIt = " & aeNumType.Debugit
21      Debug.Print , ">aeVBAversion = " & VBAversion
22      Debug.Print , ">aeOffice32or64 = " & Office32or64
    
End Sub

Private Sub Class_Terminate()
1       Debug.Print , ">Class_Terminate"
End Sub

Private Function VBAversion() As String
'    Ref: http://msdn.microsoft.com/en-us/library/ff700513(office.11).aspx
    #If VBA7 Then
3       VBAversion = "VBA7"
    #ElseIf VBA6 Then
5       VBAversion = "VBA6"
    #Else
7       VBAversion = "VBA Not 6 or 7"
    #End If
End Function

Private Function Office32or64() As String
'    Ref: http://msdn.microsoft.com/en-us/library/ff700513(office.11).aspx
    #If Win64 Then
3       Office32or64 = "64"
    #Else
5       Office32or64 = "32"
    #End If
End Function

Private Function VBComponentExists(VBCompName As String, Optional VBProj As VBIDE.VBProject = Nothing) As Boolean
'    Ref: http://www.cpearson.com/excel/vbe.aspx
'    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    This returns True or False indicating whether a VBComponent named
'    VBCompName exists in the VBProject referenced by VBProj. If VBProj
'    is omitted, the VBProject of the Active Application is used.
'    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
7       Dim VBP As VBIDE.VBProject
8       If VBProj Is Nothing Then
9           Set VBP = Application.VBE.ActiveVBProject
10      Else
11          Set VBP = VBProj
12      End If
13      On Error Resume Next
14      VBComponentExists = CBool(Len(VBP.VBComponents(VBCompName).Name))
    
End Function

Private Function aeListCodeModules(strSkipModule As String) As String()
'    Ref: http://www.exceltip.com/st/Array_variables_using_VBA_in_Microsoft_Excel/509.html
'    ====================================================================
'    Author:   Peter F. Ennis
'    Date:     March 17, 2011
'    Comment:  Add explicit references for DAO
'    ====================================================================
'
    
9       Dim objAccess As AccessObject  'Each module/form/report.
10      Dim i As Integer
11      Dim j As Integer
12      Dim astr() As String
    
14      On Error GoTo aeListCodeModules_Error
    
16      j = CurrentProject.AllModules.Count
'    Debug.Print "cnt.Documents.Count=" & j
18      ReDim astr(1 To j)
19      i = 1
20      For Each objAccess In CurrentProject.AllModules
21          If Not (Left(objAccess.Name, 3) = "zzz") And objAccess.Name <> strSkipModule And objAccess.Name <> aeNumType.MyClassName Then
'            Application.SaveAsText aTheCodeModule, objAccess.Name, aegitType.SourceFolder & "\" & doc.Name & ".bas"
'            Debug.Print i, objAccess.Name
24              astr(i) = objAccess.Name
25              i = i + 1
26          End If
27      Next objAccess
    
29      ReDim Preserve astr(1 To i)
    
31      For Each objAccess In CurrentProject.AllForms
32          If Not (Left(objAccess.Name, 3) = "zzz") Then
'            Check if a CBF module exists
34              If VBComponentExists("Form_" & objAccess.Name) Then
35                  ReDim Preserve astr(1 To i)
36                  astr(i) = "Form_" & objAccess.Name
37                  i = i + 1
38              Else
39              End If
40          End If
41      Next objAccess
    
43      For Each objAccess In CurrentProject.AllReports
44          If Not (Left(objAccess.Name, 3) = "zzz") Then
'            Check if a CBF module exists
46              If VBComponentExists("Report_" & objAccess.Name) Then
47                  ReDim Preserve astr(1 To i)
48                  astr(i) = "Report_" & objAccess.Name
49                  i = i + 1
50              Else
51              End If
52          End If
53      Next objAccess
    
55      If aeDebugIt Then
56          Debug.Print , "aeListCodeModules: aeDebugIt>" & aeDebugIt
57          For i = 1 To UBound(astr)
58              Debug.Print , i, astr(i)
59          Next
60      End If
    
62      aeListCodeModules = astr
    
'    Set doc = Nothing
'    Set cnt = Nothing
'    Set dbs = Nothing
    
68      On Error GoTo 0
    Exit Function
    
71 aeListCodeModules_Error:
    
73      MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure aeListCodeModules of Module (_) aeNum"
    
End Function

Private Function SplitLineAtLastApostrophe(ByVal str As String, ByVal strRestOfLine As String) As Boolean
'    Ref: http://vbadud.blogspot.com/2008/11/how-to-return-multiple-values-from-vba.html
'    InStr Ref: http://www.ozgrid.com/forum/showthread.php?t=19252&page=1
'    NOTE: Parsing a code line to get an inline comment is sophisticated
'    This routine returns the string including the last apostrophe in mstrPotentialComment
'    Peter F. Ennis, March 2011
    
7       Dim i As Integer
8       Dim Pos As Integer
    
'    Debug.Print "str=" & str
11      Pos = InStr(str, "'")
12      If Pos = 0 Then
13          strRestOfLine = ""
14          SplitLineAtLastApostrophe = False
15      ElseIf Pos = 2 Then            ' A space is forced in the first character position
16          strRestOfLine = ""
17          SplitLineAtLastApostrophe = False
18      Else
19          mstrPotentialComment = Mid(str, Pos, Len(str))
'        Debug.Print "SplitLineAtLastApostrophe: InStr(1, mstrPotentialComment, strRestOfLine)>" & InStr(1, mstrPotentialComment, strRestOfLine) & "<"
21          If InStr(1, mstrPotentialComment, strRestOfLine) > 0 Then
22              SplitLineAtLastApostrophe = True
23          Else
24              SplitLineAtLastApostrophe = True
25          End If
'        Debug.Print "SplitLineAtLastApostrophe: str>" & str & "<", "Pos>" & Pos & "< ", "mstrPotentialComment>" & mstrPotentialComment & " < """
27      End If
    
End Function

Private Function KeywordInComment(str As String, strKeyWord As String) As Boolean
    
2       Dim bln As Boolean
    
4       bln = SplitLineAtLastApostrophe(str, strKeyWord)
    
6       If bln Then
7           KeywordInComment = True
8       Else
9           KeywordInComment = False
10      End If
    
'    Debug.Print "KeywordInComment>" & KeywordInComment & "<", "str>" & str & "<", "strKeyWord>" & strKeyWord & "<"
    
End Function

Private Function RemoveLineComments(ByVal VBLine As String) As String
    
2       On Error GoTo RemoveLineComments_Error
    
4       Dim k As Long
5       Dim Q As Long
    
7       If InStr(1, VBLine, "'") > 0 Then
8           k = 1
9           Do
10              If Mid(VBLine, k, 1) = """" Then
11                  For Q = k + 1 To Len(VBLine)
12                      If Mid(VBLine, Q, 1) = """" Then
13                          VBLine = Left(VBLine, k) & Mid(VBLine, Q)
14                          k = k + 1
15                          Exit For
16                      End If
17                  Next Q
18              End If
            
20              k = k + 1
21          Loop Until k >= Len(VBLine)
        
23          For k = 1 To Len(VBLine)
24              If Mid(VBLine, k, 1) = "'" Then
25                  VBLine = Left(VBLine, k - 1)
26                  Exit For
27              End If
28          Next k
29      End If
    
    RemoveLineComments = Trim(VBLine)
    
33      On Error GoTo 0
    Exit Function
    
RemoveLineComments_Error:
    
38      MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure RemoveLineComments of Module aeNumClass"
    
End Function

Private Sub TestMyReNumberLeftAndAutoIndent(strSkipModule As String)
'
'    Do not renumber the module with the renumber function (i.e. this one)
'    It will crash Access.
'
    
6       Dim objVBAProject As Object
7       Set objVBAProject = Application.VBE.ActiveVBProject
    
9       Debug.Print "1: Calling MyReNumberAndLeftAlign"
10      MyReNumberAndLeftAlign objVBAProject, strSkipModule
11      Debug.Print "2: Calling MyReNumberAndLeftAlign"
12      MyReNumberAndLeftAlign objVBAProject, strSkipModule
    
14      Dim i As Integer
    
16      Debug.Print "3: Calling aeListCodeModules"
17      mstrAllCodeObjectsList = aeListCodeModules(strSkipModule)
    
19      Debug.Print "TestMyReNumberLeftAndAutoIndent:"
    
21      Debug.Print "4: Calling IndentInitialize"
22      IndentInitialize
23      Debug.Print "IndentInitialize Finished"
24      For i = 1 To UBound(mstrAllCodeObjectsList)
25          If mstrAllCodeObjectsList(i) <> strSkipModule And mstrAllCodeObjectsList(i) <> aeNumType.MyClassName Then
26              Debug.Print mstrAllCodeObjectsList(i)
27              MyIndentCodeModule mstrAllCodeObjectsList(i)
'            Sleep 3000
29          End If
30      Next i
'    Debug.Print "First time"
32      EmptyLines objVBAProject, strSkipModule
'    Second time should have no output in immediate window when testing
'    Debug.Print "Second time"
'    EmptyLines objVBAProject, strSkipModule
    
End Sub

Private Sub TestRemoveLineNumbers(strSkipModule As String)
    
2       Dim objVBAProject As Object
3       Set objVBAProject = Application.VBE.ActiveVBProject
    RemoveLineNumbers objVBAProject, strSkipModule
    
End Sub

Private Sub TestMyReNumberAndLeftAlign()
'    Do not renumber the module with the renumber function (i.e. this one)
'    It will crash Access.
    
4       Dim objVBAProject As Object
5       Set objVBAProject = Application.VBE.ActiveVBProject
6       MyReNumberAndLeftAlign objVBAProject, "aeNumClass"
    
End Sub

Private Sub TestSplitLineAtLastApostrophe()
    
2       Dim str As String
    
4       str = "test it ' here"
    
6       If SplitLineAtLastApostrophe(str, mstrPotentialComment) Then
7           Debug.Print "TestSplitLineAtLastApostrophe: str>" & str & "<", "mstrPotentialComment>" & mstrPotentialComment & "<"
8       End If
    
End Sub

Private Sub TestKeywordInComment()
    
2       Dim str As String
3       Dim bln As Boolean
    
5       str = "    x = 1             ' i do declare    "
6       Debug.Print "TestKeywordInComment: str>" & str & "<"
7       bln = KeywordInComment(str, "Declare")
8       Debug.Print "TestKeywordInComment: bln>" & bln & "<", "mstrPotentialComment>" & mstrPotentialComment & "<"
    
End Sub

Private Sub Test_aeListCodeModules()
'    Used when testing code in a Module and not in a Class
    
3       Dim i As Integer
4       Dim strSkipModule As String
5       strSkipModule = aeNumType.SkipModule
    
7       mstrAllCodeObjectsList = aeListCodeModules(strSkipModule)
    
9       Debug.Print "UBound(mstrAllCodeObjectsList)=" & UBound(mstrAllCodeObjectsList)
    
11      For i = 1 To UBound(mstrAllCodeObjectsList)
12          Debug.Print mstrAllCodeObjectsList(i)
13      Next i
    
End Sub

Private Sub TestEmptyLines()
    
2       Dim objVBAProject As Object
3       Set objVBAProject = Application.VBE.ActiveVBProject
4       EmptyLines objVBAProject, "aeNumClass"
    
End Sub

Private Sub EmptyLines(oVBP As VBProject, Optional strModuleName As Variant)
'    Peter F. Ennis, March 2011
    
3       Dim objVBComponent As VBComponent
4       Dim strProcName As String
5       Dim strOldName As String
6       Dim strLine As String
7       Dim iLine As Integer
8       Dim lngBodyLine As Long
9       Dim lngType As Long
10      Dim blnContinue As Boolean
11      Dim blnIsSelectCase As Boolean
12      Dim i As Integer
13      Dim strModName
    
15      On Error GoTo EmptyLines_Error
    
'    Do not renumber strModuleName
18      If IsMissing(strModuleName) Then
19          strModName = ""
20      Else
21          strModName = strModuleName
22      End If
    
24      i = 0
25      For Each objVBComponent In oVBP.VBComponents
26          i = i + 1
'        Debug.Print i, "objVBComponent.Name=" & objVBComponent.Name
28          If objVBComponent.CodeModule <> strModName Then
29              With objVBComponent.CodeModule
30                  strOldName = ""
31                  lngBodyLine = 0
                
33                  For iLine = 1 To .CountOfLines
34                      strLine = .Lines(iLine, 1)
                    
36                      strProcName = .ProcOfLine(iLine, lngType)
                    
38                      If strProcName <> strOldName Then
39                          lngBodyLine = .ProcBodyLine(strProcName, lngType)
40                      End If
                    
42                      If iLine > lngBodyLine And strProcName <> "" Then
'                        Debug.Print , "strLine>" & strLine & "<"
                        
45                          If IsNumeric(Left(strLine, InStr(1, strLine, " "))) Or Left(strLine, 1) = "'" Or Right(strLine, 1) <> " " Then
46                          Else
47                              .ReplaceLine iLine, Trim(strLine)
'                            For second time testing
'                            Debug.Print iLine, objVBComponent.Name, "strLine>" & strLine & "<"
50                          End If
51                      End If
                    
53                  Next
'                Sleep 1000
55              End With
56          End If
        
58      Next
    
60      On Error GoTo 0
    Exit Sub
    
63 EmptyLines_Error:
    
65      MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure EmptyLines of Module (_) aeNum"
    
End Sub

Private Sub RemoveLineNumbers(oVBP As VBProject, Optional strModuleName As Variant)
'    Peter F. Ennis, March 2011
    
3       Dim objVBComponent As VBComponent
4       Dim strProcName As String
5       Dim strOldName As String
6       Dim strLine As String
7       Dim iLine As Integer
8       Dim lngBodyLine As Long
9       Dim lngType As Long
10      Dim blnContinue As Boolean
11      Dim blnIsSelectCase As Boolean
12      Dim strModName
13      Dim Pos As Integer
    
15      On Error GoTo RemoveLineNumbers_Error
    
'    Do not renumber strModuleName
18      If IsMissing(strModuleName) Then
19          strModName = ""
20      Else
21          strModName = strModuleName
22      End If
    
24      For Each objVBComponent In oVBP.VBComponents
MsgBox objVBComponent.Name
25          If objVBComponent.CodeModule <> strModName And objVBComponent.CodeModule <> aeNumType.MyClassName Then
26              With objVBComponent.CodeModule
27                  strOldName = ""
28                  lngBodyLine = 0
                
30                  For iLine = 1 To .CountOfLines
31                      strLine = .Lines(iLine, 1)
                    
33                      strProcName = .ProcOfLine(iLine, lngType)
                    
35                      If strProcName <> strOldName Then
36                          lngBodyLine = .ProcBodyLine(strProcName, lngType)
37                      End If
                    
39                      If iLine > lngBodyLine And strProcName <> "" Then
'                        Debug.Print , "strLine>" & strLine & "<"
                        
42                          If IsNumeric(Left(strLine, InStr(1, strLine, " "))) Then
43                              Pos = InStr(strLine, " ")
'                            If objVBComponent.Name = "(_) Renumber_Bugs_Testing" Then
'                            Debug.Print "iLine>" & iLine, "Pos>" & Pos, "strLine>" & strLine
'                            End If
47                              .ReplaceLine iLine, Mid(strLine, Pos, Len(strLine) - Pos + 1)
48                          End If
49                      End If
                    
51                  Next
'                Sleep 1000
53              End With
54          End If
        
56      Next
    
58      On Error GoTo 0
    Exit Sub
    
RemoveLineNumbers_Error:
    
63      MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure RemoveLineNumbers of Module (_) aeNum"
    
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
    
11      Dim objVBComponent As VBComponent
12      Dim strProcName As String
13      Dim strOldName As String
14      Dim strLine As String
15      Dim iLine As Integer
16      Dim lngBodyLine As Long
17      Dim lngType As Long
18      Dim blnContinue As Boolean
19      Dim blnIsSelectCase As Boolean
    
21      On Error GoTo ReNumberAndLeftAlign_Error
    
    
24      For Each objVBComponent In oVBP.VBComponents
25          If objVBComponent.CodeModule <> strModuleName And objVBComponent.CodeModule <> aeNumType.MyClassName Then
26              With objVBComponent.CodeModule
27                  strOldName = ""
28                  lngBodyLine = 0
                
30                  For iLine = 1 To .CountOfLines
31                      strLine = Trim(.Lines(iLine, 1)) & " "
                    
'                    If objVBComponent.Name = "(_) Renumber_Bugs_Testing" Then
'                    Debug.Print "A:" & iLine & ", " & lngBodyLine, "strLine>" & strLine
'                    End If
                    
37                      If IsNumeric(Left(strLine, InStr(1, strLine, " "))) Then
                        
'                        Debug.Print iLine & " " & objVBComponent.Name & ">", strLine
'                        Debug.Print Left(strLine, InStr(1, strLine, " "))
'                        Debug.Print Mid(strLine, InStr(1, strLine, " "))
42                          .ReplaceLine iLine, Trim(Mid(strLine, InStr(1, strLine, " ")))
'                        If objVBComponent.Name = "(_) Renumber_Bugs_Testing" Then
'                        Debug.Print "B:" & iLine & ", " & lngBodyLine, Trim(Mid(strLine, InStr(1, strLine, " ")))
'                        End If
46                      End If
                    
48                      strProcName = .ProcOfLine(iLine, lngType)
'                    If objVBComponent.Name = "(_) Renumber_Bugs_Testing" Then
'                    Debug.Print "strProcName=" & strProcName, "iLine=" & iLine, "lngType=" & lngType, ".ProcOfLine(iLine, lngType)=" & .ProcOfLine(iLine, lngType)
'                    End If
52                      If strProcName <> strOldName Then
53                          lngBodyLine = .ProcBodyLine(strProcName, lngType)
54                      End If
                    
56                      If iLine > lngBodyLine And strProcName <> "" Then
                        
'                        If objVBComponent.Name = "(_) Renumber_Bugs_Testing" Then
'                        Debug.Print "C:" & iLine & ", " & lngBodyLine, "strLine>" & strLine
'                        Debug.Print "C1:" & iLine & ", " & lngBodyLine, "strLine>" & " " & Trim(.Lines(iLine, 1)) & " "
'                        End If
                        
63                          strLine = " " & Trim(.Lines(iLine, 1)) & " "
64                          If strLine = "  " Then GoTo ptr_Next
65                          If Left(strLine, 2) = " '" Then GoTo ptr_Next
66                          If Left(strLine, 2) = " #" Then GoTo ptr_Next
67                          If Left(strLine, 4) = " Rem" Then GoTo ptr_Next
'
69                          If InStr(1, strLine, " = Split") > 0 And InStr(1, strLine, " strLine,") = 0 Then GoTo ptr_Continue
'                        If objVBComponent.Name = "(_) Renumber_Bugs_Testing" Then
71                          If InStr(1, strLine, " Declare ") > 0 And InStr(1, strLine, " strLine,") = 0 And KeywordInComment(strLine, "Declare") Then GoTo ptr_Continue
72                          If InStr(1, strLine, " Sub ") > 0 And InStr(1, strLine, " strLine,") = 0 And KeywordInComment(strLine, "Sub") Then GoTo ptr_Continue
73                          If InStr(1, strLine, " Function ") > 0 And InStr(1, strLine, " strLine,") = 0 And KeywordInComment(strLine, "Function") Then GoTo ptr_Continue
74                          If InStr(1, strLine, " Property ") > 0 And InStr(1, strLine, " strLine,") = 0 And KeywordInComment(strLine, "Property") Then GoTo ptr_Continue
'                        End If
'
77                          If InStr(1, strLine, " Declare ") > 0 And InStr(1, strLine, " strLine,") = 0 Then GoTo ptr_Next
78                          If InStr(1, strLine, " Sub ") > 0 And InStr(1, strLine, " strLine,") = 0 Then GoTo ptr_Next
79                          If InStr(1, strLine, " Function ") > 0 And InStr(1, strLine, " strLine,") = 0 Then GoTo ptr_Next
80                          If InStr(1, strLine, " Property ") > 0 And InStr(1, strLine, " strLine,") = 0 Then GoTo ptr_Next
                        
82 ptr_Continue:
83                          strLine = .Lines(iLine, 1) & " "
84                          If IsNumeric(Left(strLine, 4)) Then strLine = Mid(strLine, 6)
85                          If IsNumeric(Left(strLine, InStr(1, strLine, " ") - 1)) Then strLine = Mid(strLine, InStr(1, strLine, " "))
86                          If Trim(strLine) = "" Then
87                              strLine = Trim(strLine)
88                          Else
89                              strLine = Format(iLine - lngBodyLine, "0000") & " " & strLine
90                          End If
'                        If objVBComponent.Name = "(_) Renumber_Bugs_Testing" Then
'                        Debug.Print "blnContinue=" & blnContinue, "blnIsSelectCase=" & blnIsSelectCase, iLine, strLine
'                        End If
94                          If Not blnContinue And Not blnIsSelectCase Then .ReplaceLine iLine, strLine
95                          If InStr(1, strLine, "Select Case") > 0 And InStr(1, strLine, " = Split(") = 0 Then
96                              blnIsSelectCase = True
97                          ElseIf InStr(1, strLine, "Case") > 0 And InStr(1, strLine, " = Split(") = 0 Then
98                              blnIsSelectCase = False
99                          ElseIf InStr(1, strLine, "End Select") > 0 And InStr(1, strLine, " = Split(") = 0 Then
100                             blnIsSelectCase = False
101                         End If
102                     End If
103 ptr_Next:
104                     blnContinue = (Right(Trim(strLine), 2) = " _")
105                 Next
'                Sleep 1000
107             End With
108         End If
        
110     Next
    
112     On Error GoTo 0
    Exit Sub
    
115 ReNumberAndLeftAlign_Error:
    
117     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ReNumberAndLeftAlign of Module (_) aeNum"
    
End Sub

Private Sub IndentInitialize()
1       BlockStart = Split("If * Then;For *;Do *;Do;Select Case *;While *;With *;Private Function *;Public Function *;Friend Function *;Function *;Private Sub *;Public Sub *;Friend Sub *;Sub *;Private Property *;Public Property *;Friend Property *;Property *;Private Enum *;Public Enum *;Friend Enum *;Enum *;Private Type *;Public Type *;Friend Type *;Type *", ";")
2       BlockEnd = Split("End If;Next*;Loop;Loop *;End Select;Wend;End With;End Function;End Function;End Function;End Function;End Sub;End Sub;End Sub;End Sub;End Property;End Property;End Property;End Property;End Enum;End Enum;End Enum;End Enum;End Type;End Type;End Type;End Type", ";")
3       BlockMiddle = Split("ElseIf * Then;Else;Case *", ";")
End Sub

Private Sub MyIndentCodeModule(strTheModuleName As String)
'    Ref: http://www.cpearson.com/excel/vbe.aspx
'    Set a reference to Microsoft Visual Basic for Applications Extensibility Library 5.3
    
4       On Error GoTo MyIndentCodeModule_Error
    
6       Dim objVBComp As VBIDE.VBComponent
7       Dim CodeMod As VBIDE.CodeModule
    
9       Dim objVBAProject As Object
10      Set objVBAProject = Application.VBE.ActiveVBProject
11      Set objVBComp = Application.VBE.ActiveVBProject.VBComponents.Item(strTheModuleName)
12      Set CodeMod = objVBComp.CodeModule
    
'    Debug.Print , objVBComp.Name
    
16      IndentCodeModule objVBAProject, CodeMod
    
18      On Error GoTo 0
    Exit Sub
    
21 MyIndentCodeModule_Error:
    
23      MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure MyIndentCodeModule of Module aeNumClass"
    
End Sub

Private Sub TestMyIndentCodeModule(strTheModuleName As String)
'    Ref: http://www.cpearson.com/excel/vbe.aspx
'    Set a reference to Microsoft Visual Basic for Applications Extensibility Library 5.3
    
4       On Error GoTo TestMyIndentCodeModule_Error
    
6       Dim objVBComp As VBIDE.VBComponent
7       Dim CodeMod As VBIDE.CodeModule
    
9       Dim objVBAProject As Object
10      Set objVBAProject = Application.VBE.ActiveVBProject
11      Set objVBComp = Application.VBE.ActiveVBProject.VBComponents.Item(strTheModuleName)
12      Set CodeMod = objVBComp.CodeModule
    
14      Debug.Print , objVBComp.Name
    
16      IndentInitialize
17      IndentCodeModule objVBAProject, CodeMod
    
19      On Error GoTo 0
    Exit Sub
    
22 TestMyIndentCodeModule_Error:
    
24      MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure TestMyIndentCodeModule of Module aeNumClass"
    
End Sub

Private Sub IndentCodeModule(oVBP As VBProject, TheCodeModule As Object)
'    Ref: Ref: http://www.vbforums.com/showthread.php?t=479449
    
3       On Error GoTo IndentCodeModule_Error
    
5       Dim AllCode As String
6       Dim AllLines() As String
7       Dim LineNumbers() As String
8       Dim k As Long
9       Dim Q As Long
10      Dim VBLine As String
11      Dim StartCursorPos As Long
12      Dim TopLine As Long
    
14      With TheCodeModule
15          TopLine = .CodePane.TopLine
16          .CodePane.GetSelection StartCursorPos, 0, 0, 0
        
18          AllCode = .Lines(1, .CountOfLines)
        
20          AllCode = Replace(AllCode, "_" & vbNewLine, "")
21          AllLines = Split(AllCode, vbNewLine)
22          ReDim LineNumbers(UBound(AllLines))
        
24          For k = 0 To UBound(AllLines)
25              Q = InStr(1, AllLines(k), " ")
26              If Q = 0 Then Q = InStr(1, AllLines(k), vbTab)
            
28              If Q > 0 Then
29                  If CStr(Val(Left(AllLines(k), Q - 1))) = Left(AllLines(k), Q - 1) Then
30                      LineNumbers(k) = Left(AllLines(k), Q - 1)
31                      AllLines(k) = Mid(AllLines(k), Q)
32                  End If
33              End If
            
35              AllLines(k) = Trim(AllLines(k))
            
37              If Left(AllLines(k), 1) = "'" Then
38                  AllLines(k) = "' " & Trim(Mid(AllLines(k), 2))
39              End If
40          Next k
        
42          IndentBlock AllLines
        
44          For k = 0 To UBound(AllLines)
45              VBLine = RemoveLineComments(AllLines(k))
            
47              For Q = 0 To UBound(BlockMiddle)
48                  If Replace(VBLine, vbTab, "") Like BlockMiddle(Q) And Left(VBLine, 1) = vbTab Then
49                      AllLines(k) = Mid(AllLines(k), 2)
50                  End If
51              Next Q
52          Next k
        
54          For k = 0 To UBound(AllLines)
55              If Len(LineNumbers(k)) > 0 Then
56                  AllLines(k) = LineNumbers(k) & vbTab & AllLines(k)
57              End If
58          Next k
        
60          AllCode = Replace(Join(AllLines, vbNewLine), Chr(1), "")
        
62          .DeleteLines 1, .CountOfLines
63          .InsertLines 1, AllCode
        
65          TheCodeModule.CodePane.SetSelection StartCursorPos, 1, StartCursorPos, 1
66          .CodePane.TopLine = TopLine
67      End With
    
69      On Error GoTo 0
    Exit Sub
    
72 IndentCodeModule_Error:
    
74      MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure IndentCodeModule of Module aeNumClass"
    
End Sub

Private Sub IndentBlock(VBLines() As String)
'    Ref: Ref: http://www.vbforums.com/showthread.php?t=479449
    
3       On Error GoTo IndentBlock_Error
    
5       Dim k As Long
6       Dim Q As Long
7       Dim EndPos As Long
8       Dim VBLine As String
9       Dim StartPos As Long
10      Dim FoundStartEnd As Boolean
    
12      Do
13          StartPos = 0
14          EndPos = UBound(VBLines)
        
16          Do
17              FoundStartEnd = False
            
19              For k = EndPos To StartPos Step -1
20                  VBLine = RemoveLineComments(VBLines(k))
'                Debug.Print "VBLine=" & VBLine
                
23                  If Len(VBLine) > 0 Then
24                      For Q = 0 To UBound(BlockStart)
25                          If VBLine Like BlockStart(Q) Then
26                              StartPos = k + 1
27                              FoundStartEnd = True
28                              Exit For
29                          End If
30                      Next Q
                    
32                      If Q <= UBound(BlockStart) Then Exit For
33                  End If
34              Next k
            
36              If FoundStartEnd Then
37                  For k = StartPos To EndPos
38                      VBLine = RemoveLineComments(VBLines(k))
                    
40                      If Len(VBLine) > 0 Then
41                          If VBLine Like BlockEnd(Q) Then
42                              EndPos = k - 1
43                              FoundStartEnd = True
44                              Exit For
45                          End If
46                      End If
47                  Next k
48              End If
49          Loop While FoundStartEnd
        
51          If Not (Not FoundStartEnd And StartPos = 0 And EndPos = UBound(VBLines)) Then
'            Debug.Print StartPos, EndPos
53              IndentLineBlock VBLines, StartPos, EndPos
54          End If
55      Loop Until Not FoundStartEnd And StartPos = 0 And EndPos = UBound(VBLines)
    
57      On Error GoTo 0
    Exit Sub
    
60 IndentBlock_Error:
    
62      MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure IndentBlock of Module aeNumClass"
    
End Sub

Private Sub IndentLineBlock(VBLines() As String, ByVal StartLine As Long, ByVal EndLine As Long, Optional NoIndentChar As Byte = 1)
'    Ref: Ref: http://www.vbforums.com/showthread.php?t=479449
    
3       On Error GoTo IndentLineBlock_Error
    
5       Dim k As Long
    
7       If NoIndentChar > 0 Then
8           If StartLine > 0 Then VBLines(StartLine - 1) = Chr(NoIndentChar) & VBLines(StartLine - 1)
9           If EndLine < UBound(VBLines) Then VBLines(EndLine + 1) = Chr(NoIndentChar) & VBLines(EndLine + 1)
10      End If
    
12      For k = StartLine To EndLine
13          If Left(VBLines(k), 1) = "'" Then
14              VBLines(k) = "' " & vbTab & Mid(VBLines(k), 2)
15          Else
16              VBLines(k) = vbTab & VBLines(k)
17          End If
18      Next k
    
20      On Error GoTo 0
    Exit Sub
    
23 IndentLineBlock_Error:
    
25      MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure IndentLineBlock of Module aeNumClass"
    
End Sub