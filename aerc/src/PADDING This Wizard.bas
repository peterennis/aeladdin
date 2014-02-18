Option Compare Database
Option Explicit

' Access 2007 supports ADO 2.8 , DAO 12.0
' Access 2010 supports ADO 6.1 , DAO 12.0 (14.0 for DbEngine)

' Things to change betweens Access versions :
' - References: change to the correct Microsoft ActiveX Data Object version (see above)
' - Constant c_vAcceptedAccessVersions: put accepted internal Access versions separated by ;
' Access 2000 => 9.0 , Access 2002 => 10.0 , Access 2003 => 11.0 , Access 2007 => 12.0, Access 2010 => 14.0
' Note: It is better to seperate Access 2002 and 2003 because of a bug (http://support.microsoft.com/kb/897764/en-us/), but it used to be "Access 2002-2003 => 10.0;11.0".

Public Const c_vAcceptedAccessVersions = "15.0"                         'zzz "12.0;14.0;15.0"
Public Const c_vVersion = "0.01"
Public Const c_vName = "aeladdin"
' zzz Public Const cz_vFullName0 = c_vName & "(Access %AccVerE%)" ' bug from the 1.50 to 1.52 version
Public Const cz_vFullName = c_vName                                     'zzz & " (Access %AccVerE%)"

Public Const c_vEditor = "adaept"
Public Const cz_vDefaultProgDir = "\" & c_vEditor & "\" & c_vName       'zzz & " %AccVerE%"
Public Const cz_vProgFile = "aeladdinT"                                 'zzz %AccVerE%"
Public Const cz_vDataFile = "aeladdinT"                                 'zzz %AccVerE%-dat"
Public Const c_vMenuPrefixe0 = "aeT - "
Public Const c_vMenuPrefixe1 = "aeT -- "
Public Const c_vToolIdMax = 8
' zzz Public Const c_vDefaultMenuSel = "0;1;2;3;4;5;6;7;8"
Public Const c_vDefaultMenuSel = "0;;;;4;;;;"
Public Const c_vSetupForm = "aeT_SETUP"
Public Const c_LOGFILE = cz_vProgFile & "_LOG_"
Public gstr_LOGFILE As String
Public Const gblnLOG = False
Public Const c_MainForm = "aeT_ABOUT"

' Registry keys
Public Const c_vReg_LocalMachine = "LM"
Public Const c_vReg_ClassesRoot = "CR"
Public Const c_vReg_RootKey = "Software"

Public Const cz_vReg_EditorPath = c_vReg_RootKey & "\" & c_vEditor
Public Const cz_vReg_ProgKey = c_vName                                  'zzz & "%AccVerE%"
Public Const cz_vReg_ProgPath = cz_vReg_EditorPath & "\" & cz_vReg_ProgKey

Public Const cz_vReg_Uninstall = "Software\Microsoft\Windows\CurrentVersion\Uninstall"
Public Const c_vReg_AccessCommand = "Access.Application.%AccVerI%\Shell\Open\Command"
Public Const cz_vReg_AccessAddIns = "Software\Microsoft\Office\%AccVerI%.0\Access\Menu Add-Ins"

' Registry values's name
Public Const c_vReg_Version = "Version"
Public Const c_vReg_ProgDir = "ProgDir"
Public Const c_vReg_DataDir = "DataDir"
Public Const c_vReg_Language = "Language"
Public Const c_vReg_MenuSel = "MenuSelection"

' Keywords for constants
Public Const c_vAccVer_Internal = "%AccVerI%"
Public Const c_vAccVer_External = "%AccVerE%"
Public Const c_vAccVer_Compatible = "%AccVerC%"

Public Function zzzm_AccVer(ByVal str As String) As String
'    This function merges a string with the current Access version (Internal or Extrenal)
    
       Dim x As String
    
       If InStr(str, c_vAccVer_Compatible) > 0 Then
           Dim Lst As Variant
           Dim i   As Integer
           x = vbNullString
           Lst = Split(c_vAcceptedAccessVersions, ";")
          For i = LBound(Lst) To UBound(Lst)
              If x <> vbNullString Then
                  x = x & " - "
              End If
              x = x & zzzm_AccVersNum("" & Lst(i), False)
          Next i
          x = Replace(str, c_vAccVer_Compatible, x)
      Else
          x = str
      End If
    
      If InStr(x, c_vAccVer_Internal) > 0 Then
          x = Replace(x, c_vAccVer_Internal, zzzm_AccVersNum(SysCmd(acSysCmdAccessVer), True))
      End If
    
      If InStr(x, c_vAccVer_External) > 0 Then
          x = Replace(x, c_vAccVer_External, zzzm_AccVersNum(SysCmd(acSysCmdAccessVer), False))
      End If
    
      zzzm_AccVer = x
    
End Function

Public Function m_AccExt() As String
'    Returns the extension of the current V-Tools Access file.
       Dim p As Long
       p = InStrRev(Application.CodeProject.FullName, ".")
       If p > 0 Then
           m_AccExt = Mid$(Application.CodeProject.FullName, p)
       End If
End Function

Public Function zzzm_AccVersNum(CmdAccessVer As String, Internal As Boolean) As String
    
       Select Case CmdAccessVer
    Case "14.0" 'Access 2010
           zzzm_AccVersNum = IIf(Internal, "14", "2010")
       Case "12.0" 'Access 2007
           zzzm_AccVersNum = IIf(Internal, "12", "2007")
       Case "11.0" 'Access 2003
           zzzm_AccVersNum = IIf(Internal, "11", "2003")
       Case "10.0" 'Access 2002
          zzzm_AccVersNum = IIf(Internal, "10", "2002")
      Case "9.0" 'Access 2000
          zzzm_AccVersNum = IIf(Internal, "9", "2000")
      Case "8.0" 'Access 97
          zzzm_AccVersNum = IIf(Internal, "8", "97")
      Case Else
          zzzm_AccVersNum = "? (" & CmdAccessVer & ")"
      End Select
    
End Function

Public Function m_OpenTool(ToolId As Integer)
    
       Dim FrmName As String
    
       On Error GoTo m_OpenTool_Error
    
'    zzz  FrmName = Choose(ToolId + 1, c_MainForm, "Db_ApplySystemColors", "Db_PictureData", "Db_WorkOnQueries", "Db_SearchThrougthObjects", "Db_Spec", "Db_FormTemplate", "Db_PersonalLibrary", "Db_Containers", c_vSetupForm)
       FrmName = Choose(ToolId + 1, c_MainForm, "Db_SearchThroughObjects", "Form1", "Form2", "Form3", "Form4", "Form5", "Form6", "Form7", "Form8", c_vSetupForm)
    
       gstr_LOGFILE = f_MyDocuments & "\" & c_LOGFILE & Format(Date, "yyyymmdd") & ".txt"
    
      If gblnLOG Then
          Open gstr_LOGFILE For Append As #1
          Print #1, Date, "Now=" & Now, "m_OpenTool"
          Print #1, , "ToolId=" & ToolId
          Print #1, , "FrmName=" & FrmName
          Close #1
      End If
    
      If CodeProject.AllForms(c_MainForm).IsLoaded Then
          DoCmd.Close acForm, c_MainForm
      End If
    
      On Error Resume Next 'This is because in some Forms, the Open event can be cancel if the current Database is an ADP project.
      If FrmName = c_vSetupForm Then
'        The Setup Form is not modal when the code is not activated (Access 2007), thus it is easy to click on the Activate option.
'        But is should be modal in normal usage.
          DoCmd.OpenForm FrmName, , , , , acDialog
      Else
          DoCmd.OpenForm FrmName
      End If
    
      On Error GoTo 0
    Exit Function
    
m_OpenTool_Error:
    
      MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure m_OpenTool of Module This Wizard"
    
End Function

Public Function m_ProgDir(Optional HKey As Long) As String
'    Returns the Program (V) Directory
    
       Dim Folder As String
    
       Debug.Print "HKey=" & HKey
       If HKey = 0 Then
'        zzz        Folder = "" & m_Reg_ValueGetQuick(c_vReg_LocalMachine, m_AccVer(cz_vReg_ProgPath), c_vReg_ProgDir)
           Folder = "" & m_Reg_ValueGetQuick(c_vReg_LocalMachine, cz_vReg_ProgPath, c_vReg_ProgDir)
       Else
          Folder = "" & m_Reg_ValueGet(HKey, c_vReg_ProgDir)
      End If
    
      m_ProgDir = Folder
    
End Function

Public Function m_MenuSel_Get(Optional HKey As Long) As String
'    Returns the current menu selection
'    If a HKey is given, it means the correct Registry Key is already open with this Handle.
    
       Dim MenuSel As String
    
       If HKey = 0 Then
'        zzz        MenuSel = "" & m_Reg_ValueGetQuick(c_vReg_LocalMachine, m_AccVer(cz_vReg_ProgPath), c_vReg_MenuSel)
           MenuSel = "" & m_Reg_ValueGetQuick(c_vReg_LocalMachine, cz_vReg_ProgPath, c_vReg_MenuSel)
       Else
          MenuSel = "" & m_Reg_ValueGet(HKey, c_vReg_MenuSel)
      End If
    
      If MenuSel = "" Then
          MenuSel = c_vDefaultMenuSel
      End If
    
      m_MenuSel_Get = MenuSel
    
End Function

Public Function m_MenuSel_Set(NewLanguage As Integer, NewMenuSel As String, Optional HKeyV As Long, Optional ForceRefresh As Boolean)
'    Actualize the add-in Menu items
    
       Const c_Sep = ";"
    
       Dim HKey        As Long
       Dim OldLanguage As Integer
       Dim OldMenuSel  As String
       Dim ProgPath    As String
       Dim i           As Integer
    
      On Error GoTo m_MenuSel_Set_Error
    
      If HKeyV = 0 Then
'        zzz        HKey = m_Reg_KeyOpen(c_vReg_LocalMachine, m_AccVer(cz_vReg_ProgPath))
          HKey = m_Reg_KeyOpen(c_vReg_LocalMachine, cz_vReg_ProgPath)
      Else
          HKey = HKeyV
      End If
    
'    Get old values and save with new one
    
      OldLanguage = Val("0" & m_Reg_ValueGet(HKey, c_vReg_Language)) 'm_DefaultLanguage(HKey)
      If NewLanguage <> OldLanguage Then
          m_Reg_ValueSet HKey, c_vReg_Language, "" & NewLanguage, REG_SZ
      End If
    
      OldMenuSel = "" & m_Reg_ValueGet(HKey, c_vReg_MenuSel) 'm_MenuSel_Get(HKey)
      If OldMenuSel <> NewMenuSel Then
          m_Reg_ValueSet HKey, c_vReg_MenuSel, NewMenuSel, REG_SZ
      End If
    
      ProgPath = m_ProgDir(HKey)
'    zzz    ProgPath = m_Folder_CheckEnd(ProgPath, True) & m_AccVer(cz_vProgFile) & m_AccExt()
      ProgPath = m_Folder_CheckEnd(ProgPath, True) & cz_vProgFile & m_AccExt()
    
      If HKeyV = 0 Then
          m_Reg_KeyClose HKey
      End If
    
'    First verifying if there is something to modify
      If (NewLanguage <> OldLanguage) Or (NewMenuSel <> OldMenuSel) Or (ForceRefresh = True) Then
        
'        zzz        HKey = m_Reg_KeyOpen(c_vReg_LocalMachine, m_AccVer(cz_vReg_AccessAddIns))
          HKey = m_Reg_KeyOpen(c_vReg_LocalMachine, cz_vReg_AccessAddIns)
        
          If HKey <> 0 Then
            
              Dim SelId  As String
              Dim DelOK  As Boolean
              Dim AddOk  As Boolean
              Dim HKeyM  As Long
              Dim OldLst As Variant
              Dim NewLst As Variant
            
              OldLst = m_ToolList(OldLanguage, True)
              NewLst = m_ToolList(NewLanguage, True)
              NewMenuSel = c_Sep & NewMenuSel & c_Sep
            
              For i = 0 To c_vToolIdMax
                
                  SelId = c_Sep & i & c_Sep
                
                  If ForceRefresh = True Then
                      m_Reg_KeyDelete HKey, vbNullString & OldLst(i)
                      m_Reg_KeyDelete HKey, vbNullString & NewLst(i)
                      AddOk = (InStr(NewMenuSel, SelId) > 0)
                  Else
'                    Look if deleting old item is necessary
                      If OldLanguage = 0 Then
                          DelOK = False
                      Else
                          If InStr(OldMenuSel, SelId) = 0 Then
                              DelOK = False
                          Else
                              DelOK = (NewLanguage <> OldLanguage) Or (InStr(NewMenuSel, SelId) = 0)
                          End If
                      End If
'                    Look if adding new item is necessary
                      If NewLanguage = 0 Then
                          AddOk = False
                      Else
                          If InStr(NewMenuSel, SelId) = 0 Then
                              AddOk = False
                          Else
                              AddOk = (NewLanguage <> OldLanguage) Or (InStr(OldMenuSel, SelId) = 0)
                          End If
                      End If
'                    Delete the old item
                      If DelOK Then
                          m_Reg_KeyDelete HKey, vbNullString & OldLst(i)
                      End If
                  End If
                
'                Add the new item
                  If AddOk Then
'                    zzz                    HKeyM = m_Reg_KeyCreate(c_vReg_LocalMachine, m_AccVer(cz_vReg_AccessAddIns) & "\" & NewLst(i))
                      HKeyM = m_Reg_KeyCreate(c_vReg_LocalMachine, cz_vReg_AccessAddIns & "\" & NewLst(i))
                      If HKeyM <> 0 Then
                          m_Reg_ValueSet HKeyM, "Library", ProgPath, REG_SZ
                         m_Reg_ValueSet HKeyM, "Expression", "=m_OpenTool(" & i & ")", REG_SZ
                         m_Reg_ValueSet HKeyM, "Version", "3", REG_SZ
                         m_Reg_KeyClose HKeyM
                     End If
                 End If
                
             Next i
            
             m_Reg_KeyClose HKey
             Application.ReloadAddIns
             DoEvents
            
         End If
        
     End If
    
     On Error GoTo 0
    Exit Function
    
m_MenuSel_Set_Error:
    
     MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure m_MenuSel_Set of Module This Wizard"
    
End Function

Public Sub m_Debug_OpenUsysTables()
'    For debuging purpose only: open User System tables, which are hidden by default
    
       Const c_Pref = "USYS"
       Dim TblName As String
    
       Dim i As Long
       For i = 0 To CodeDb.TableDefs.Count - 1
           TblName = CodeDb.TableDefs(i).Name
           If Left$(TblName, Len(c_Pref)) = c_Pref Then
              DoCmd.OpenTable TblName, acViewNormal
          End If
      Next i
    
End Sub

Public Sub m_Debug_DisplayVersions()
'    For debuging purpose only: display the versions of Access, ADO and DAO
    
       Debug.Print "Access version = " & Application.VERSION
       Debug.Print "ADO (ActiveX Data Object) version = " & CurrentProject.Connection.VERSION
       Debug.Print "DAO (DbEngine)  version = " & Application.DBEngine.VERSION
       Debug.Print "DAO (CodeDb)    version = " & Application.CodeDb.VERSION
       Debug.Print "DAO (CurrentDb) version = " & Application.CurrentDb.VERSION
    
End Sub