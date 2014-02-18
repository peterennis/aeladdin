Option Compare Database
Option Explicit

Private Declare Function api_GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Public Function m_File_GetFolder(FileFullPath As String) As String
'    Returns the folder path of the specified file.
    
       Dim i As Integer
    
       m_File_GetFolder = FileFullPath
       i = Len(FileFullPath)
       Do Until i = 0
           If Mid$(FileFullPath, i, 1) = "\" Then
               m_File_GetFolder = Left$(FileFullPath, i - 1)
              i = 1
          Else
              i = i - 1
          End If
      Loop
    
End Function

Public Function m_File_IsExisting(FileFullPath As String) As Boolean
'    Returns True if the File specified does existes.
    
       Dim x As String
    
       m_File_IsExisting = False
       On Error Resume Next
       x = Dir$(FileFullPath) ' Since Access 2007, (Obj.Prop="x") can return true if Prop doesn't exist
       m_File_IsExisting = (x <> "")
    
End Function

Public Function m_File_Delete(FileFullPath As String) As Boolean
    
       On Error GoTo Err_File_Delete
    
       Kill FileFullPath
       m_File_Delete = True
    
End_File_Delete:
    Exit Function
    
Err_File_Delete:
      m_File_Delete = False
      Resume End_File_Delete
    
End Function

Public Function m_File_Copy(SrcFile As String, DstFile As String)
    
       Const c_Len = 32000
    
       Dim SrcNum As Integer
       Dim DstNum As Integer
       Dim f_Len As Long
       Dim f_End As Long
    
'    Dim x As String * c_Len '(the lenght of the variable is necessary fot the Get instruction)
      ReDim x(c_Len - 1) As Byte ' for Japanese localize, tip sent by Yu-Tang
      Dim i As Long
    
      SrcNum = FreeFile()
      Open SrcFile For Binary Access Read As #SrcNum Len = c_Len
      DstNum = FreeFile()
      Open DstFile For Binary Access Write As #DstNum Len = c_Len '(the Binary mode is necessary in order to not add the lenght of the variable with the Put instruction)
    
'    Save pieces of the data with packages of 32000 bytes
      f_Len = LOF(SrcNum)
      For i = 1 To (f_Len \ c_Len)
          Get #SrcNum, , x
          Put #DstNum, , x
      Next i
    
'    Save the last package
      f_End = (f_Len Mod c_Len)
      If f_End > 0 Then
          ReDim x(f_End - 1) As Byte ' for Japanese localize, tip sent by Yu-Tang
          Get #SrcNum, , x
          Put #DstNum, , x
      End If
    
      Close #SrcNum
      Close #DstNum
    
End Function

Public Function m_Folder_CheckEnd(FolderPath As String, EndWithSep As Boolean) As String
'    This function add the '\' at this end of the Directory path only if it is missing.
    
       Const c_Sep = "\"
    
       If Right(FolderPath, 1) = c_Sep Then
           If EndWithSep Then
               m_Folder_CheckEnd = FolderPath
           Else
               m_Folder_CheckEnd = Left$(FolderPath, Len(FolderPath) - 1)
          End If
      Else
          If EndWithSep Then
              m_Folder_CheckEnd = FolderPath & c_Sep
          Else
              m_Folder_CheckEnd = FolderPath
          End If
      End If
    
End Function

Public Function m_Folder_IsExisting(FolderPath As String) As Boolean
'    Returns True if the specified folder does existes.
    
       Dim x As String
    
       m_Folder_IsExisting = False
       On Error Resume Next
       x = Dir$(FolderPath, vbDirectory) ' Since Access 2007, (Obj.Prop="x") can return true if Prop doesn't exist
       m_Folder_IsExisting = (x <> "")
    
End Function

Public Function m_Folder_GetTmp() As String
    
       Const MAX_PATH = 260
    
       Dim x As String
       Dim r As Long
    
       x = String$(MAX_PATH, 0)
       r = api_GetTempPath(MAX_PATH, x)
    
      If r <> 0 Then
          m_Folder_GetTmp = Left$(x, InStr(x, Chr$(0)) - 1)
      Else
          m_Folder_GetTmp = ""
      End If
    
End Function