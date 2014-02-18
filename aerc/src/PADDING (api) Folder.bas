Option Compare Database
Option Explicit

Private Type udtBROWSEINFO
    hwndOwner As Long
    pidlRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type

Private Const BIF_RETURNONLYFSDIRS = 1
Private Const MAX_PATH = 260

Private Declare Sub API_CoTaskMemFree Lib "ole32.dll" Alias "CoTaskMemFree" (ByVal hMem As Long)
Private Declare Function API_lstrCat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function API_SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolder" (lpbi As udtBROWSEINFO) As Long
Private Declare Function api_SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDList" (ByVal pidList As Long, ByVal lpBuffer As String) As Long


Public Function m_Folder_SelectBox(sPrompt As String, Optional hwnd As Long) As String
'    Purpose    : Select a directory (or folder) using a Windows standard dialog box
'    Inputs     : strPrompt : Message show in the dialog box
    
       Dim iNull As Integer
       Dim lpidList As Long
       Dim lResult As Long
       Dim sPath As String
       Dim udtBI As udtBROWSEINFO
    
      With udtBI
          If hwnd = 0 Then
              .hwndOwner = Application.hWndAccessApp
          Else
              .hwndOwner = hwnd
          End If
          .lpszTitle = API_lstrCat(sPrompt, "")
          .ulFlags = BIF_RETURNONLYFSDIRS
      End With
    
      lpidList = API_SHBrowseForFolder(udtBI)
      If lpidList Then
          sPath = String$(MAX_PATH, 0)
          lResult = api_SHGetPathFromIDList(lpidList, sPath)
          Call API_CoTaskMemFree(lpidList)
          iNull = InStr(sPath, vbNullChar)
          If iNull Then
              sPath = Left$(sPath, iNull - 1)
          End If
      End If
    
      m_Folder_SelectBox = sPath
    
End Function