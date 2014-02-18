Option Compare Database
Option Explicit

Public g_CurrDb    As Object 'The CurrentDb() object has to be keeped in a varibale. It is nesessary for the m_CurrProj_OBjOpen() function
Public g_TableDefs As Object
Public g_QueryDefs As Object

Public Sub m_CurrDb_Init()
    
       If g_CurrDb Is Nothing Then
           Set g_CurrDb = CurrentDb()
           If CurrentProject.ProjectType = acMDB Then
               Set g_TableDefs = g_CurrDb.TableDefs
               Set g_QueryDefs = g_CurrDb.QueryDefs
           End If
       End If
    
End Sub

Public Function m_CurrProj_IsMDE() As Boolean
    
       On Error Resume Next
    
       Dim x As String
    
       m_CurrProj_IsMDE = False
       x = CurrentDb.Properties("MDE") ' Since Access 2007, (Obj.Prop="x") can return true if Prop doesn't exist
       m_CurrProj_IsMDE = (x = "T")
    
End Function

Public Function m_CurrProj_CanReadDef(objType As AcObjectType, ObjName As String) As Boolean
    
       Const dbSecReadDef = 4
    
       Dim p        As Long
       Dim ContName As String
    
       Select Case objType
    Case acTable: ContName = "Tables"
       Case acQuery: ContName = "Tables"
      Case acForm: ContName = "Forms"
      Case acReport: ContName = "Reports"
      Case acMacro: ContName = "Scripts"
      Case acModule: ContName = "Modules"
      Case acDataAccessPage: ContName = "DataAccessPages"
      End Select
    
      p = g_CurrDb.Containers(ContName).Documents(ObjName).AllPermissions
      m_CurrProj_CanReadDef = ((p And dbSecReadDef) = dbSecReadDef)
    
End Function

Public Function m_CurrProj_CanWriteDef(objType As AcObjectType, ObjName As String) As Boolean
    
       Const dbSecWriteDef = 65548
    
       Dim p        As Long
       Dim ContName As String
    
       Select Case objType
    Case acTable: ContName = "Tables"
       Case acQuery: ContName = "Tables"
      Case acForm: ContName = "Forms"
      Case acReport: ContName = "Reports"
      Case acMacro: ContName = "Scripts"
      Case acModule: ContName = "Modules"
      Case acDataAccessPage: ContName = "DataAccessPages"
      End Select
    
      p = g_CurrDb.Containers(ContName).Documents(ObjName).AllPermissions
      m_CurrProj_CanWriteDef = ((p And dbSecWriteDef) = dbSecWriteDef)
    
End Function

Public Sub m_CurrProj_FillObjList(ByRef ObjLst() As String, ByRef ObjNbr As Long, objType As AcObjectType, Optional CountNeverSavedObjects As Boolean)
    
       Dim i      As Long
       Dim j      As Long
       Dim x      As String
    
       ObjNbr = 0
    
       Select Case objType
    Case acTable
          With CurrentData.AllTables
              For i = 0 To (.Count - 1)
                  p_CurrProj_ObjList_Add ObjLst(), ObjNbr, .Item(i).Name
              Next i
          End With
      Case acQuery
          With CurrentData.AllQueries
              For i = 0 To (.Count - 1)
                  p_CurrProj_ObjList_Add ObjLst(), ObjNbr, .Item(i).Name
              Next i
          End With
      Case acForm
          With CurrentProject.AllForms
              For i = 0 To (.Count - 1)
                  p_CurrProj_ObjList_Add ObjLst(), ObjNbr, .Item(i).Name
              Next i
          End With
      Case acReport
          With CurrentProject.AllReports
              For i = 0 To (.Count - 1)
                  p_CurrProj_ObjList_Add ObjLst(), ObjNbr, .Item(i).Name
              Next i
          End With
      Case acMacro
          With CurrentProject.AllMacros
              For i = 0 To (.Count - 1)
                  p_CurrProj_ObjList_Add ObjLst(), ObjNbr, .Item(i).Name
              Next i
          End With
      Case acModule
          With CurrentProject.AllModules ' Modules are sorted descending
              For i = 0 To (.Count - 1)
                  p_CurrProj_ObjList_Add ObjLst(), ObjNbr, .Item(i).Name
              Next i
          End With
      Case acDataAccessPage
          With CurrentProject.AllDataAccessPages
              For i = 0 To (.Count - 1)
                  p_CurrProj_ObjList_Add ObjLst(), ObjNbr, .Item(i).Name
              Next i
          End With
      End Select
    
'    Sort the list because items seems to be sorted a strange way
      For i = 1 To ObjNbr - 1
          j = i
          Do Until j < 1
              If ObjLst(j) < ObjLst(j - 1) Then
                  x = ObjLst(j - 1)
                  ObjLst(j - 1) = ObjLst(j)
                  ObjLst(j) = x
                  j = j - 1
              Else
                  j = 0 ' ends the loop
              End If
          Loop
      Next i
    
'    Add objects not yet saved
      If CountNeverSavedObjects = True Then
        
          Select Case objType
        Case acTable
'            Not supported
          Case acQuery
'            Not supported
          Case acForm
              With Application.Forms
                  For i = 0 To (.Count - 1)
                      If (SysCmd(acSysCmdGetObjectState, objType, .Item(i).Name) And acObjStateNew) <> 0 Then
                          p_CurrProj_ObjList_Add ObjLst(), ObjNbr, .Item(i).Name
                      End If
                  Next i
              End With
          Case acReport
              With Application.Reports
                  For i = 0 To (.Count - 1)
                      If (SysCmd(acSysCmdGetObjectState, objType, .Item(i).Name) And acObjStateNew) <> 0 Then
                          p_CurrProj_ObjList_Add ObjLst(), ObjNbr, .Item(i).Name
                      End If
                  Next i
              End With
          Case acMacro
'            Not supported
          Case acModule
              With Application.Modules
                  For i = 0 To (.Count - 1)
                      If (SysCmd(acSysCmdGetObjectState, objType, .Item(i).Name) And acObjStateNew) <> 0 Then
                          p_CurrProj_ObjList_Add ObjLst(), ObjNbr, .Item(i).Name
                      End If
                 Next i
             End With
         Case acDataAccessPage
             With Application.DataAccessPages
                 For i = 0 To (.Count - 1)
                     If (SysCmd(acSysCmdGetObjectState, objType, .Item(i).Name) And acObjStateNew) <> 0 Then
                         p_CurrProj_ObjList_Add ObjLst(), ObjNbr, .Item(i).Name
                     End If
                 Next i
             End With
         End Select
        
     End If
    
End Sub

Private Sub p_CurrProj_ObjList_Add(ByRef ObjLst() As String, ByRef ObjNbr As Long, Item As String)
    
       ReDim Preserve ObjLst(0 To ObjNbr)
       ObjLst(ObjNbr) = Item
       ObjNbr = ObjNbr + 1
    
End Sub

Public Function m_CurrProj_ObjOpen(objType As AcObjectType, ObjName As String, ByRef WasOpen As Boolean, OpenObject As Object) As Boolean
    
       Set OpenObject = Nothing
       WasOpen = False
    
       On Error Resume Next
    
       Select Case objType
        
    Case acTable
        
          If CurrentProject.ProjectType = acMDB Then
              Set OpenObject = g_CurrDb.TableDefs(ObjName)
          End If
        
      Case acQuery
        
          If CurrentProject.ProjectType = acMDB Then
              Set OpenObject = g_CurrDb.QueryDefs(ObjName)
          End If
        
      Case acForm
        
          Set OpenObject = Forms(ObjName)
          If OpenObject Is Nothing Then
'            SendKeys "{ENTER}"
              DoCmd.OpenForm ObjName, acDesign, , , , acHidden
              Set OpenObject = Forms(ObjName)
          Else
              WasOpen = True
          End If
        
      Case acReport
        
          Set OpenObject = Reports(ObjName)
          If OpenObject Is Nothing Then
'            SendKeys "{ENTER}"
              DoCmd.OpenReport ObjName, acViewDesign
'            DoCmd.OpenReport ObjName, acViewDesign, , , acHidden 'acHidden is available only since Access 2002
              Set OpenObject = Reports(ObjName)
          Else
              WasOpen = True
          End If
        
      Case acModule
        
          Set OpenObject = Modules(ObjName)
          If OpenObject Is Nothing Then
'            SendKeys "{ENTER}"
              DoCmd.OpenModule ObjName
              Set OpenObject = Modules(ObjName)
          Else
              WasOpen = True
          End If
        
      End Select
    
      m_CurrProj_ObjOpen = Not (OpenObject Is Nothing)
    
End Function