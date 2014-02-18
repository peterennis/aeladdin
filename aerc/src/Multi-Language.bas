Option Compare Database
Option Explicit

' Const dbOpenForwardOnly = 8
' Const dbReadOnly = 4
' Const dbOpenSnapshot = 4

Const c_LanguageNotSet = 0      ' The language is not set yet
Const c_DevLanguage = -1        ' Form's language as it has been created
Const c_NbrLang = 5

Global g_CurrentLanguage As Integer ' Language currently applied to the last open tool
Global g_DefaultLanguage As Integer ' Language to apply to a tool when it is open
Global Const c_LanguagePopup = "cm_VTools_LanguagePopup"

Private Declare Function GetUserDefaultUILanguage Lib "kernel32" () As Long

Public Function m_ApplyLanguageF(Frm As Object, Optional Language As Variant)
'    Must be a function to be called by menus
'    Note: Frm is set as Object instead of Form because is makes a type mistake with DB_SETUP and Access 2000
    
       Dim IsSetupForm As Boolean
    
       On Error GoTo m_ApplyLanguageF_Error
    
       If gblnLOG Then
           Open gstr_LOGFILE For Append As #1
          Print #1, Date, "Now=" & Now, "m_ApplyLanguageF"
          Print #1, , "Frm.Name=" & Frm.Name
          If IsMissing(Language) Then
              Print #1, , "Language=IsMissing"
          Else
              Print #1, , "Language=" & Language
          End If
          Print #1, , "m_DefaultLanguage()=" & m_DefaultLanguage()
          Print #1, , "c_DevLanguage=" & c_DevLanguage
          Print #1, , "c_vSetupForm=" & c_vSetupForm
          Print #1, , "c_vName=" & c_vName
          Print #1, , "c_vVersion=" & c_vVersion
          Print #1, , "Frm.Tag=" & Frm.Tag
          Close #1
      End If
    
      IsSetupForm = (Frm.Name = c_vSetupForm)
    
      If IsMissing(Language) Then
          g_CurrentLanguage = m_DefaultLanguage()
      Else
          g_CurrentLanguage = Language
      End If
    
      If g_CurrentLanguage <> c_DevLanguage Then
        
          Dim rs  As New ADODB.Recordset
          Dim Txt As String
        
          With rs
              .Open "SELECT S_ObjectType,S_ObjectName,S_ObjectProperty,S_Language_" & g_CurrentLanguage & " AS Txt FROM USysMultiLanguage WHERE (S_ObjectParent='" & Frm.Name & "')", CodeProject.Connection, adOpenForwardOnly
              Do Until .EOF
                  Txt = Nz(!Txt, "???")
                  Txt = Replace(Txt, "%vName%", c_vName)
                  Txt = Replace(Txt, "%vVers%", c_vVersion)
                  If IsSetupForm Then
'                    zzz                    Txt = m_AccVer(Txt)
                      Txt = Replace(Txt, "%vRegVers%", Frm.Tag)
                  End If
                  Select Case !S_ObjectType
                Case "F"
                      Frm.Properties(!S_ObjectProperty) = Txt
                  Case "FC"
                      Frm.Controls(!S_ObjectName).Properties(!S_ObjectProperty) = Txt
                  End Select
                  .MoveNext
              Loop
              .Close
          End With
        
      End If
    
      If gblnLOG Then
          Open gstr_LOGFILE For Append As #1
          Print #1, Date, "Now=" & Now, "END> m_ApplyLanguageF"
          Close #1
      End If
    
      On Error GoTo 0
    Exit Function
    
m_ApplyLanguageF_Error:
    
      MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure m_ApplyLanguageF of Module Multi-Language"
    
End Function

Public Function m_ToolList(Optional Language As Integer, Optional WithPrefixe As Boolean) As Variant
'    Returns a variant wich is an array containing the tool labels.
'    Used for registry and in the Reception Form.
    
       Dim i As Integer
       Dim Lst(0 To c_vToolIdMax)
       Dim rs  As New ADODB.Recordset
       Dim Txt As String
    
       On Error GoTo m_ToolList_Error
    
      If Language = 0 Then
          i = m_DefaultLanguage()
      Else
          i = Language
      End If
    
      With rs
          rs.Open "SELECT S_ObjectName,S_Language_" & i & " AS Txt FROM USysMultiLanguage WHERE (S_ObjectParent='Sys') AND (S_ObjectType='TT')", CodeProject.Connection, adOpenForwardOnly
          Do Until .EOF
              i = Val(!S_ObjectName)
              Txt = !Txt
              If WithPrefixe = True Then
                  If i = 0 Then
                      Txt = c_vMenuPrefixe0 & Txt
                  Else
                      Txt = c_vMenuPrefixe1 & Txt
                  End If
              End If
              Lst(i) = Txt
              .MoveNext
          Loop
          .Close
      End With
    
      m_ToolList = Lst()
    
    
      On Error GoTo 0
    Exit Function
    
m_ToolList_Error:
    
      MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure m_ToolList of Module Multi-Language"
    
End Function

Public Function m_CurrentLanguage() As Integer
    
'    If Current Language Code is not set, the value is taken from the Default Language.
       If g_CurrentLanguage = c_LanguageNotSet Then
           g_CurrentLanguage = m_DefaultLanguage()
       End If
    
       m_CurrentLanguage = g_CurrentLanguage
    
End Function

Public Function m_DefaultLanguage(Optional HKey As Long) As Integer
    
'    If Default Language Code is not set, the value is taken from the Registry.
       If g_DefaultLanguage = c_LanguageNotSet Then
           If HKey = 0 Then
'            zzz            g_DefaultLanguage = Val("" & m_Reg_ValueGetQuick(c_vReg_LocalMachine, m_AccVer(cz_vReg_ProgPath), c_vReg_Language))
               g_DefaultLanguage = Val("" & m_Reg_ValueGetQuick(c_vReg_LocalMachine, cz_vReg_ProgPath, c_vReg_Language))
           Else
               g_DefaultLanguage = Val(m_Reg_ValueGet(HKey, c_vReg_Language))
           End If
      End If
    
'    If the language has not been determinated, we take the default language
      If g_DefaultLanguage = c_LanguageNotSet Then
          g_DefaultLanguage = m_WindowsLanguage()
      End If
    
      m_DefaultLanguage = g_DefaultLanguage
    
End Function

Public Function m_ChangeLanguage(NewLanguage As Integer)
    
'    zzz    m_Reg_ValueSetQuick c_vReg_LocalMachine, m_AccVer(cz_vReg_ProgPath), c_vReg_Language, ("" & NewLanguage), REG_SZ
       m_Reg_ValueSetQuick c_vReg_LocalMachine, cz_vReg_ProgPath, c_vReg_Language, ("" & NewLanguage), REG_SZ
       g_CurrentLanguage = NewLanguage
    
End Function

Public Function m_MsgBox(ObjectParent As String, ObjectName As String, Buttons As VbMsgBoxStyle, ParamArray Params() As Variant) As VbMsgBoxResult 'Integer
'    Prompts a Message Box depending to the current language.
    
       Const c_Joker = "%"
    
       Dim rs     As New ADODB.Recordset
       Dim i      As Integer
       Dim Txt    As String
       Dim xTitle As String
       Dim xBody  As String
    
      On Error GoTo m_MsgBox_Error
    
      If gblnLOG Then
          Open gstr_LOGFILE For Append As #1
          Print #1, Date, "Now=" & Now, "m_MsgBox"
          Close #1
      End If
    
      i = m_CurrentLanguage()
    
      If ObjectName = vbNullString Then
          rs.Open "SELECT S_ObjectProperty,S_Language_" & i & " AS Txt FROM USysMultiLanguage WHERE (S_ObjectParent='" & ObjectParent & "') AND (S_ObjectType='M')", CodeProject.Connection, adOpenForwardOnly
      Else
          rs.Open "SELECT S_ObjectProperty,S_Language_" & i & " AS Txt FROM USysMultiLanguage WHERE (S_ObjectParent='" & ObjectParent & "') AND (S_ObjectType='M') AND (S_ObjectName='" & ObjectName & "')", CodeProject.Connection, adOpenForwardOnly
      End If
    
      xBody = "<message not found>"
      xTitle = "<title not found>"
    
      With rs
          Do Until .EOF
'            Get the text
              Txt = Nz(!Txt)
'            Replace the jokeys with the parameters
              For i = LBound(Params) To UBound(Params)
                  Txt = Replace(Txt, c_Joker & (i + 1) & c_Joker, vbNullString & Params(i))
              Next i
              Txt = Replace(Txt, "%VName%", c_vName)
              Txt = Replace(Txt, "%VVers%", c_vVersion)
'            Save in variables
              Select Case !S_ObjectProperty
            Case "Body"
                  xBody = Txt
              Case "Title"
                  xTitle = Txt
              End Select
              .MoveNext
          Loop
          .Close
      End With
    
      m_MsgBox = MsgBox(xBody, Buttons, xTitle)
    
      If gblnLOG Then
          Open gstr_LOGFILE For Append As #1
          Print #1, Date, "Now=" & Now, "END> m_MsgBox"
          Print #1, , "ObjectParent=" & ObjectParent
          Print #1, , "ObjectName=" & ObjectName
          For i = LBound(Params) To UBound(Params)
              Print #1, , "Params(" & i & ")=" & Params(i)
          Next i
          Print #1, , "xBody=" & xBody
          Print #1, , "xTitle=" & xTitle
          Print #1, , "c_vSetupForm=" & c_vSetupForm
          Print #1, , "c_vName=" & c_vName
          Print #1, , "c_vVersion=" & c_vVersion
          Close #1
      End If
    
      On Error GoTo 0
    Exit Function
    
m_MsgBox_Error:
    
      MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure m_MsgBox of Module Multi-Language"
    
End Function

Public Function m_GetText(ObjectParent As String, Optional ObjectName As String, Optional Language As Integer) As String
'    Returns the texte from the String Table depending to the current language.
'    [ObjectParent] is not optional, it's the key for standart text information.
'    But if it's easier to let the Form name for a label, ObjectName can be precised.
    
       Dim rs  As New ADODB.Recordset
       Dim Col As String
       Dim Txt As String
    
       On Error GoTo m_GetText_Error
    
      Select Case Language
    Case -1
          Col = "S_ObjectProperty"
      Case 0
          Col = "S_Language_" & m_CurrentLanguage()
      Case Else
          Col = "S_Language_" & Language
      End Select
    
      With rs
        
          If ObjectName = "" Then
              .Open "SELECT * FROM USysMultiLanguage WHERE (S_ObjectParent='" & ObjectParent & "')", CodeProject.Connection, adOpenForwardOnly
          Else
              .Open "SELECT * FROM USysMultiLanguage WHERE (S_ObjectParent='" & ObjectParent & "') AND (S_ObjectName='" & ObjectName & "')", CodeProject.Connection, adOpenForwardOnly
          End If
        
          If .EOF Then
              Txt = "<" & ObjectParent & "." & ObjectName & ": translation not found>"
          Else
              Txt = vbNullString & .Fields(Col).Value
              Txt = Replace(Txt, "%VName%", c_vName)
              Txt = Replace(Txt, "%VVers%", c_vVersion)
          End If
        
          .Close
        
      End With
    
      m_GetText = Txt
    
    
      On Error GoTo 0
    Exit Function
    
m_GetText_Error:
    
      MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure m_GetText of Module Multi-Language"
    
End Function

Public Function m_LanguagePopup()
'    Show the language popup
    
       On Error Resume Next
       CommandBars(c_LanguagePopup).ShowPopup
    
End Function

Public Function m_WindowsLanguage() As Integer
'    Language used by the current Windows configuration
    
       Dim LId As Long
       Dim Vid As Integer
    
       LId = GetUserDefaultUILanguage()
'    http://msdn2.microsoft.com/en-us/library/ms776324(VS.85).aspx
'    http://msdn2.microsoft.com/en-us/library/ms776208(VS.85).aspx
    
      Select Case Right$(Hex$(LId), 2) ' The primary language is part of the Language Identifier
    Case "0C": Vid = 2 ' French
      Case "07": Vid = 3 ' German
      Case "16": Vid = 4 ' Portuguese
      Case "11": Vid = 5 ' Japanese
      Case Else: Vid = 1 ' English
      End Select
    
      m_WindowsLanguage = Vid
    
End Function