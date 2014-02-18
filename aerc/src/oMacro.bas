Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' This object enables reading of Macro parameters
' From Access 2000, parts of the Access API that used to exist in Acc97 are not available anymore.
' Those functions are replaced by some WizHook function.
' WizHook is a new hidden feature that appears in Access 2000.
' To see it, you have to go int the Object Explorer and choose 'Show hiden members'.
' Then you have to activate it with a key, this key is 51488399.
' If you don't activate the key, all WizHook functions returns 0.

Private Declare Function api_Script_Open Lib "msaccess.exe" Alias "#18" (ByVal lpszScript As String, ByVal lpszLabel As Any, ByVal smode As Long, pgrfExtra As Long, psmv As Long) As Long
Private Declare Sub api_Script_Close Lib "msaccess.exe" Alias "#20" (ByVal Hscr As Long)
Private Declare Sub api_Script_Rewind Lib "msaccess.exe" Alias "#19" (ByVal Hscr As Long)
Private Declare Function api_Script_NextRow Lib "msaccess.exe" Alias "#22" (ByVal Hscr As Long, ByVal fSkipBlank As Long, pfEndOfScript As Long) As Long
' 0& : Label, 1& : Comment, 2& : Condition, 3&-12& : Arguments
Private Declare Function api_Script_GetPart Lib "msaccess.exe" Alias "#23" (ByVal Hscr As Long, ByVal iscc As Long, ByVal lpsz As String, ByVal cchMax As Long) As Long
Private Declare Function api_Script_SetPart Lib "msaccess.exe" Alias "#24" (ByVal Hscr As Long, ByVal iscc As Long, ByVal lpsz As String) As Long
Private Declare Function api_Script_GetActId Lib "msaccess.exe" Alias "#29" (ByVal Hscr As Long) As Long
Private Declare Function api_Script_GetActArgNbr Lib "msaccess.exe" Alias "#30" (ByVal ActId As Long) As Long

Dim Hscr        As Long
Dim EndOfScript As Long
Dim pName       As String
Dim ActId       As Long
Dim ActNbr      As Long

Public Function OpenMacro(MacroName As String) As Boolean
    
       Dim Mac_VVersion    As Long
       Dim Mac_Extra       As Long
       Dim Mac_Label       As String
    
       WizHook.Key = 51488399
    
       If Hscr <> 0 Then
           CloseMacro
      End If
    
      Hscr = 0
      EndOfScript = 0
      ActNbr = 0
      ActId = 0
    
      Hscr = WizHook.OpenScript(MacroName, Mac_Label, 0&, Mac_Extra, Mac_VVersion)
    
      If Hscr <> 0 Then
          pName = MacroName
          Me.NextAction
          OpenMacro = True
      End If
    
End Function

Public Sub CloseMacro()
    
       api_Script_Close Hscr
       Hscr = 0
       EndOfScript = True
       ActNbr = 0
       ActId = 0
       pName = vbNullString
    
End Sub

Public Sub NextAction()
    
       If (Hscr <> 0) And (Not EndOfScript) Then
           api_Script_NextRow Hscr, False, EndOfScript
           ActId = api_Script_GetActId(Hscr)
           ActNbr = ActNbr + 1
       End If
    
End Sub

Public Function EndOfMacro() As Boolean
       EndOfMacro = EndOfScript
End Function

Public Function Name() As String
       Name = pName
End Function

Public Function CurrAct_Index() As Long
       CurrAct_Index = ActNbr
End Function

Public Function CurrAct_ArgNbr() As Long
       CurrAct_ArgNbr = WizHook.ArgsOfActid(ActId)
End Function

Public Function CurrAct_Name() As String
       CurrAct_Name = WizHook.NameFromActid(ActId)
End Function

Public Function CurrAct_Label() As String
       CurrAct_Label = p_GetScriptString(0)
End Function

Public Function CurrAct_Comment() As String
       CurrAct_Comment = p_GetScriptString(1)
End Function

Public Function CurrAct_Condition() As String
       CurrAct_Condition = p_GetScriptString(2)
End Function

Public Function CurrAct_Parameter(PrmIndex As Long) As String
       CurrAct_Parameter = p_GetScriptString(2 + PrmIndex)
End Function

Private Function p_GetScriptString(index As Long) As String
    
       Dim x As String
    
       If WizHook.GetScriptString(Hscr, index, x) Then
       Else
           x = vbNullString
       End If
    
       p_GetScriptString = x
    
End Function