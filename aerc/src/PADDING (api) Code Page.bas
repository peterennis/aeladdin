Option Compare Database
Option Explicit

' Call f_CodePage_Init() to feed the global varibales g_CodePageLst() and g_CodePageNbr
' with the Code Pages supported by the current Windows installation.
' 2007-12-27, Skrol29

Private Const MAX_LEADBYTES = 12
Private Const MAX_DEFAULTCHAR = 2
Private Const MAX_PATH = 260

Private Type CPINFOEX
    MaxCharSize As Long ' max length (Byte) of a char
    DefaultChar(0 To MAX_DEFAULTCHAR - 1) As Byte ' default character
    LeadByte(0 To MAX_LEADBYTES - 1) As Byte ' lead byte ranges
    UnicodeDefaultChar As Byte
    CodePage As Long
    CodePageName(0 To MAX_PATH - 1) As Byte
End Type

Private Declare Function api_EnumSystemCodePages Lib "kernel32.dll" Alias "EnumSystemCodePagesA" (ByVal lpCodePageEnumProc As Long, ByVal dwFlags As Long) As Long
Private Declare Sub api_CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function api_GetCPInfoEx Lib "kernel32" Alias "GetCPInfoExA" (ByVal CodePage As Long, ByVal dwFlags As Long, ByRef lpCPInfoEX As CPINFOEX) As Long

' EnumSystemCodePages
Const CP_INSTALLED = &H1
Const CP_SUPPORTED = &H2

Public Type tCodePage
    id As Long
    Name As String
End Type
Public g_CodePageLst() As tCodePage
Public g_CodePageNbr As Long

Public Function f_CodePage_Init() As Long
    
       If g_CodePageNbr > 0 Then
           f_CodePage_Init = g_CodePageNbr
        Exit Function
       End If
    
'    Call the CallBack as much time as there are CodePage
       api_EnumSystemCodePages AddressOf f_CodePage_CallBack, CP_INSTALLED
    
      f_CodePage_Init = g_CodePageNbr
    
End Function

Public Function f_CodePage_CallBack(CP_Pointer As Long) As Long
'    This function is called by the API EnumSystemCodePages to retrieve all Code Pages installed on Windows.
'    Returns TRUE to continue enumeration or FALSE otherwise.
    
       Dim Buffer As String
       Dim id As Long
    
       Buffer = Space$(255)
       Call api_CopyMemory(ByVal Buffer, CP_Pointer, Len(Buffer))
       Buffer = Left$(Buffer, InStr(Buffer, Chr$(0)) - 1)
    
      id = Val(Buffer)
      If id > 0 Then f_CodePage_Add id, f_CodePage_Name(id)
    
      f_CodePage_CallBack = 1
    
End Function

Public Sub f_CodePage_Debug()
    
       Dim i As Long
    
       For i = 0 To g_CodePageNbr - 1
           Debug.Print g_CodePageLst(i).id & " : " & g_CodePageLst(i).Name
       Next i
    
End Sub

Public Sub f_CodePage_Sort()
    
       Dim i As Long
       Dim j As Long
       Dim CodePage As tCodePage
    
       For i = 0 To g_CodePageNbr - 1
           j = i
           Do While j > 0
               If g_CodePageLst(j).Name < g_CodePageLst(j - 1).Name Then
'                Swap items
                  CodePage = g_CodePageLst(j - 1)
                  g_CodePageLst(j - 1) = g_CodePageLst(j)
                  g_CodePageLst(j) = CodePage
                  j = j - 1
              Else
'                End of the loop
                  j = 0
              End If
          Loop
      Next i
    
End Sub

Public Function f_CodePage_Name(id As Long) As String
'    Returns the name of a given Code Page id.
    
       Dim CpInfo As CPINFOEX
       Dim x As String
       Dim i As Integer
       Dim Ok As Boolean
    
       Call api_GetCPInfoEx(id, 0, CpInfo)
    
      With CpInfo
        
'        Retrieve the full name
          i = LBound(.CodePageName)
          Ok = True
          Do While Ok And (i <= UBound(.CodePageName))
              If .CodePageName(i) = 0 Then
                  Ok = False
              Else
                  x = x & Chr$(.CodePageName(i))
                  i = i + 1
              End If
          Loop
        
'        Take of the number. Names can be formated like "1234 (name)"
          i = InStr(x, "(")
          If (.CodePage > 0) And (i > 0) And (Right(x, 1) = ")") Then
              If Abs(Val(Left$(x, i - 1)) - .CodePage) <= 2 Then ' we use Abs() becasue the name of 50221 is prefixed with 50220
                  x = Mid$(x, i + 1)
                  x = Left$(x, Len(x) - 1)
              End If
          End If
        
'        Debug Vista has some Code Pages with no names
          If x = vbNullString Then
              If id = 1147 Then x = "IBM EBCDIC (France-Euro)"
              If id = 20949 Then x = "Korean Wansung"
          End If
        
'        If no name => display the id
          If x = vbNullString Then
              x = "{" & .CodePage & "}"
          End If
        
      End With
    
      f_CodePage_Name = x
    
End Function

Public Function f_CodePage_CurrentId() As Long
'    Returns the id of the Windows current Code Page
    
       Dim CpInfo As CPINFOEX
    
       Call api_GetCPInfoEx(0, 0, CpInfo)
    
       f_CodePage_CurrentId = CpInfo.CodePage
    
    
End Function

Public Function f_CodePage_Add(id As Long, Name As String) As Long
'    Add a new code page in the list, and return the index if the item
    
       ReDim Preserve g_CodePageLst(0 To g_CodePageNbr)
    
       With g_CodePageLst(g_CodePageNbr)
           .id = id
           .Name = Name
       End With
    
      f_CodePage_Add = g_CodePageNbr
    
      g_CodePageNbr = g_CodePageNbr + 1
    
End Function

Public Function f_CodePage_Index(id As Long) As Long
'    Return the index of the code page in the list, returns -1 if the code page does not exist in the list.
    
       Dim i As Long
       Dim Ok As Boolean
    
       i = 0
       Do Until Ok Or (i >= g_CodePageNbr)
           If g_CodePageLst(i).id = id Then
               Ok = True
          Else
              i = i + 1
          End If
      Loop
    
      If Ok Then
          f_CodePage_Index = i
      Else
          f_CodePage_Index = -1
      End If
    
End Function