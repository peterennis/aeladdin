Option Compare Database
Option Explicit

Private Declare Function api_EmptyClipBoard Lib "user32" Alias "EmptyClipboard" () As Long
Private Declare Function api_OpenClipBoard Lib "user32" Alias "OpenClipboard" (ByVal hwnd As Long) As Long
Private Declare Function api_CloseClipBoard Lib "user32" Alias "CloseClipboard" () As Long
Private Declare Function api_SetClipBoardData Lib "user32" Alias "SetClipboardData" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Private Declare Function api_GetClipboardData Lib "user32" Alias "GetClipboardData" (ByVal wFormat As Long) As Long
Private Declare Function api_IsClipboardFormatAvailable Lib "user32" Alias "IsClipboardFormatAvailable" (ByVal uFormat As Long) As Long
Private Declare Function api_GetPriorityClipboardFormat Lib "user32" Alias "GetPriorityClipboardFormat" (lpPriorityList As Long, ByVal nCount As Long) As Long

Private Declare Function api_GlobalAlloc Lib "kernel32" Alias "GlobalAlloc" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function api_GlobalFree Lib "kernel32" Alias "GlobalFree" (ByVal hMem As Long) As Long
Private Declare Function api_GlobalLock Lib "kernel32" Alias "GlobalLock" (ByVal hMem As Long) As Long
Private Declare Function api_GlobalUnlock Lib "kernel32" Alias "GlobalUnlock" (ByVal hMem As Long) As Long
Private Declare Function api_GlobalSize Lib "kernel32" Alias "GlobalSize" (ByVal hMem As Long) As Long

Private Declare Sub api_CopyMem Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

' Global Memory Flags
Private Const GMEM_MOVEABLE = &H2
Private Const GMEM_ZEROINIT = &H40
Private Const GHND = (GMEM_MOVEABLE Or GMEM_ZEROINIT)

' Predefined Clipboard Formats
Private Const CF_TEXT = 1
Private Const CF_BITMAP = 2
Private Const CF_METAFILEPICT = 3
Private Const CF_SYLK = 4
Private Const CF_DIF = 5
Private Const CF_TIFF = 6
Private Const CF_OEMTEXT = 7
Private Const CF_DIB = 8
Private Const CF_PALETTE = 9
Private Const CF_PENDATA = 10
Private Const CF_RIFF = 11
Private Const CF_WAVE = 12
Private Const CF_UNICODETEXT = 13
Private Const CF_ENHMETAFILE = 14
Private Const CF_HDROP = 15
Private Const CF_LOCALE = 16
Private Const CF_MAX = 17
Private Const CF_OWNERDISPLAY = &H80
Private Const CF_DSPTEXT = &H81
Private Const CF_DSPBITMAP = &H82
Private Const CF_DSPMETAFILEPICT = &H83
Private Const CF_DSPENHMETAFILE = &H8E

' http://msdn.microsoft.com/en-us/library/ms776445.aspx
Private Const IS_TEXT_UNICODE_ASCII16 = &H1
Private Const IS_TEXT_UNICODE_CONTROLS = &H4
Private Const IS_TEXT_UNICODE_DBCS_LEADBYTE = &H400
Private Const IS_TEXT_UNICODE_ILLEGAL_CHARS = &H100
Private Const IS_TEXT_UNICODE_NOT_ASCII_MASK = &HF000
Private Const IS_TEXT_UNICODE_NOT_UNICODE_MASK = &HF00
Private Const IS_TEXT_UNICODE_NULL_BYTES = &H1000
Private Const IS_TEXT_UNICODE_ODD_LENGTH = &H200
Private Const IS_TEXT_UNICODE_REVERSE_ASCII16 = &H10
Private Const IS_TEXT_UNICODE_REVERSE_CONTROLS = &H40
Private Const IS_TEXT_UNICODE_REVERSE_MASK = &HF0
Private Const IS_TEXT_UNICODE_REVERSE_SIGNATURE = &H80
Private Const IS_TEXT_UNICODE_REVERSE_STATISTICS = &H20
Private Const IS_TEXT_UNICODE_SIGNATURE = &H8
Private Const IS_TEXT_UNICODE_STATISTICS = &H2
Private Const IS_TEXT_UNICODE_UNICODE_MASK = &HF

Private Declare Function IsTextUnicode Lib "advapi32" (ByVal lpBuffer As String, ByVal cb As Long, lpi As Long) As Long
Private Declare Function IsTextPointerUnicode Lib "advapi32" Alias "IsTextUnicode" (ByVal lpBuffer As Long, ByVal cb As Long, lpi As Long) As Long

Public Sub m_ClipBoard_Set(ByVal Data As String, Optional ByVal cbFormat As Long = 0)
'    Copy the string directly to the Windows clipboard.
    
       Dim hMemAlloc As Long
       Dim hMemLock  As Long
       Dim DataLen   As Long
       Dim Result    As Long
       Dim OpenOk    As Boolean
    
       On Error GoTo Err_ClipBoard_Set
    
      If cbFormat = 0 Then cbFormat = CF_UNICODETEXT
    
      m_Txt_LastNullChar_Add Data ' A null character should signals the end of the data. http://msdn.microsoft.com/en-us/library/ms649013(VS.85).aspx
    
      DataLen = LenB(Data)
      hMemAlloc = api_GlobalAlloc(GHND, DataLen)
    
      If hMemAlloc <> 0 Then
        
          hMemLock = api_GlobalLock(hMemAlloc)
        
          api_CopyMem ByVal hMemLock, ByVal StrPtr(Data), DataLen
        
          If api_GlobalUnlock(hMemAlloc) = 0 Then
              If api_OpenClipBoard(0&) <> 0 Then
                  OpenOk = True
                  Result = api_EmptyClipBoard()
                  Result = api_SetClipBoardData(cbFormat, hMemAlloc)
              End If
          End If
        
      End If
    
End_ClipBoard_Set:
      If OpenOk Then
          Result = api_CloseClipBoard()
      End If
    Exit Sub
    
Err_ClipBoard_Set:
      MsgBox Err.Description
      Resume End_ClipBoard_Set
    
End Sub

Function m_Txt_IsAscii(Txt As String) As Boolean
    
       If Len(Txt) = LenB(Txt) Then
           m_Txt_IsAscii = True
        Exit Function
       End If
    
       Dim i As Long
    
       For i = 1 To Len(Txt)
          If Asc(MidB$(Txt, 2 * i, 1) & vbNullChar) <> 0 Then
              Exit Function ' The result is False
          End If
      Next i
    
      m_Txt_IsAscii = True
    
End Function

Sub m_Txt_LastNullChar_Del(ByRef Txt As String)
'    Delete the last null char if any
    
       Dim p As Long
       p = InStr(1, Txt, vbNullChar, 0)
       If p > 0 Then
           Txt = Mid(Txt, 1, p - 1)
       End If
    
End Sub

Sub m_Txt_LastNullChar_Add(ByRef Txt As String)
'    Ensure that the last char is a null char.
    
       If Right(Txt, 1) <> vbNullChar Then
           Txt = Txt & vbNullChar
       End If
    
End Sub

Public Function m_ClipBoard_Get(Optional cbFormat As Long = 0) As String
'    Get the string content from the Windows clipboard.
    
       Dim hMem      As Long
       Dim lMem      As Long
       Dim Txt       As String
       Dim TxtByte() As Byte
       Dim TxtLen    As Long
       Dim Result    As Long
       Dim OpenOk    As Boolean
    
      On Error GoTo Err_ClipBoard_Get
    
      If cbFormat = 0 Then
          cbFormat = m_ClipBoard_GetPriorityFormat(CF_UNICODETEXT, CF_TEXT, CF_OEMTEXT)
      End If
    
      If cbFormat > 0 Then
          If api_OpenClipBoard(0&) <> 0 Then
              OpenOk = True
              hMem = api_GetClipboardData(cbFormat)
              TxtLen = api_GlobalSize(hMem)
              If TxtLen > 0 Then
                
                  ReDim TxtByte(0 To TxtLen - 1)
                  lMem = api_GlobalLock(hMem)
                  api_CopyMem TxtByte(0), ByVal lMem, TxtLen
                
                  Txt = TxtByte()
                
                  If cbFormat <> CF_UNICODETEXT Then Txt = StrConv(Txt, vbUnicode)
                  m_Txt_LastNullChar_Del Txt ' A null character should signals the end of the data. http://msdn.microsoft.com/en-us/library/ms649013(VS.85).aspx
                
                  Result = api_GlobalUnlock(hMem)
              End If
          End If
      Else
          Txt = "<no text data available in the clipboard>"
      End If
    
      m_ClipBoard_Get = Txt
    
End_ClipBoard_Get:
      If OpenOk Then
          Result = api_CloseClipBoard()
      End If
    Exit Function
    
Err_ClipBoard_Get:
      MsgBox Err.Description
      Resume End_ClipBoard_Get
    
End Function

Public Function m_ClipBoard_GetPriorityFormat(ParamArray Formats()) As Long
    
       Dim Fmts() As Long
       Dim i As Long
       Dim nFmt As Long
    
'    Bail, if no formats were requested
    If UBound(Formats) < 0 Then Exit Function
    
'    Transfer desired formats into a non-variant array
      ReDim Fmts(0 To UBound(Formats)) As Long
      For i = 0 To UBound(Formats)
'        Double conversion, to be safer.
'        Could error trap, but that'd mean the
'        user was a hoser, and we wouldn't want
'        to insinuate *that*, would we?
          Fmts(i) = CLng(Val(Formats(i)))
      Next i
    
      nFmt = api_GetPriorityClipboardFormat(Fmts(0), UBound(Fmts) + 1)
    
'    Return results
      m_ClipBoard_GetPriorityFormat = nFmt
    
End Function