Option Compare Database
Option Explicit

Private Type OPENFILENAME
    lStructSize       As Long
    hwndOwner         As Long
    hInstance         As Long
    lpstrFilter       As String
    lpstrCustomFilter As String
    nMaxCustomFilter  As Long
    nFilterIndex      As Long
    lpstrFile         As String
    nMaxFile          As Long
    lpstrFileTitle    As String
    nMaxFileTitle     As Long
    lpstrInitialDir   As String
    lpstrTitle        As String
    flags             As Long
    nFileOffset       As Integer
    nFileExtension    As Integer
    lpstrDefExt       As String
    lCustData         As Long
    lpfnHook          As Long
    lpTemplateName    As String
End Type

Private Const OFN_ALLOWMULTISELECT = &H200
Private Const OFN_CREATEPROMPT = &H2000
Private Const OFN_ENABLEHOOK = &H20
Private Const OFN_ENABLETEMPLATE = &H40
Private Const OFN_ENABLETEMPLATEHANDLE = &H80
Private Const OFN_EXPLORER = &H80000
Private Const OFN_EXTENSIONDIFFERENT = &H400&
Private Const OFN_FILEMUSTEXIST = &H1000
Private Const OFN_HIDEREADONLY = &H4&
Private Const OFN_LONGNAMES = &H200000
Private Const OFN_NOCHANGEDIR = &H8&
Private Const OFN_NODEREFERENCELINKS = &H100000
Private Const OFN_NOLONGNAMES = &H40000
Private Const OFN_NONETWORKBUTTON = &H20000
Private Const OFN_NOTESTFILECREATE = &H10000
Private Const OFN_OVERWRITEPROMPT = &H2&
Private Const OFN_PATHMUSTEXIST = &H800
Private Const OFN_READONLY = &H1
Private Const OFN_SHAREAWARE = &H4000
Private Const OFN_SHAREFALLTHROUGH = 2
Private Const OFN_SHAREWARN = 0
Private Const OFN_SHARENOWARN = 1
Private Const OFN_SHOWHELP = &H10

Private Declare Function api_GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenFileName As OPENFILENAME) As Boolean
Private Declare Function api_GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenFileName As OPENFILENAME) As Boolean

Public Function m_File_Selection(title As String, Optional Filter As String = "All (*.*)|*.*||", Optional MustExist As Boolean, Optional DefaultFileName As String, Optional DefaultDirectory As String, Optional hwnd As Long) As String
    
       Dim xOFN   As OPENFILENAME
       Dim ret    As Boolean
    
       With xOFN
        
           .hwndOwner = hwnd
           .hInstance = 0
           .lpstrFilter = Replace(Filter, "|", vbNullChar)
          .nMaxCustomFilter = 0
          .nFilterIndex = 1
          .lpstrFile = Space(256) & vbNullChar
          .nMaxFile = Len(.lpstrFile)
          .lpstrFileTitle = Space(256) & vbNullChar
          .nMaxFileTitle = Len(.lpstrFileTitle)
          .lpstrInitialDir = DefaultDirectory & vbNullChar
          .lpstrTitle = title & vbNullChar
          .flags = OFN_PATHMUSTEXIST Or OFN_FILEMUSTEXIST Or OFN_HIDEREADONLY
          .nFileOffset = 0
          .nFileExtension = 0
          .lCustData = 0
          .lpfnHook = 0
          .lStructSize = Len(xOFN)
        
          If MustExist Then
              ret = api_GetOpenFileName(xOFN)
          Else
              ret = api_GetSaveFileName(xOFN)
          End If
        
          If ret Then
'            Get the Chr$(0) of wich means the end of the string
              m_File_Selection = Left$(.lpstrFile, InStr(.lpstrFile, vbNullChar) - 1)
          Else
              m_File_Selection = ""
          End If
        
      End With
    
End Function