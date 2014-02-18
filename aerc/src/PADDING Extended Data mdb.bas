Option Compare Database
Option Explicit

Global Const c_Mdb_FormTemplate = "T_FormTemplate"
Global Const c_Mdb_PersonalObject = "T_PersonalObject"

Public Function m_DataMdb_OpenCnx(Optional ByVal DataMdb As String) As ADODB.Connection
    
       Const dbLangGeneral = ";LANGID=0x0409;CP=1252;COUNTRY=0"
    
       Dim Cnx As New ADODB.Connection
       Dim p   As Integer
       Dim VerTxt As String
    
       If DataMdb = vbNullString Then
           DataMdb = m_DataMdb_FullPath()
      End If
    
'    Creates the mdb if it doesn't exist
      If Not m_File_IsExisting(DataMdb) Then
          Dim DbE As Object
          VerTxt = DBEngine.VERSION
          p = InStr(VerTxt, ".")
          If p > 0 Then VerTxt = Left$(VerTxt, p - 1) & Mid$(VerTxt, p + 1, 1) ' take od the dat and keep only one decimal
          Set DbE = CreateObject("DAO.DbEngine." & VerTxt)
          DbE.CreateDatabase DataMdb, dbLangGeneral
          Set DbE = Nothing
      End If
    
      Cnx.Open "Provider=" & CodeProject.Connection.Provider & ";Data Source=" & DataMdb
    
      Set m_DataMdb_OpenCnx = Cnx
    
End Function

Public Sub m_DataMdb_CheckTable(TableName As String, Optional Cnx As ADODB.Connection, Optional ByVal DataMdb As String)
'    This function check if the specified table exists in the Data mdb.
'    If the mdb doesn't exist, the function creats it.
'    If the tables doesn't exist, the function creates it.
    
       Dim CloseAtEnd As Boolean
    
       If Cnx Is Nothing Then
           CloseAtEnd = True
           Set Cnx = m_DataMdb_OpenCnx(DataMdb)
      End If
    
      If Not p_Table_Existes(Cnx, TableName) Then
          p_Table_Create Cnx, TableName
      End If
    
      If CloseAtEnd = True Then
          Cnx.Close
          Set Cnx = Nothing
      End If
    
End Sub

Public Function m_DataMdb_HasData(Optional ByVal DataMdb As String) As Boolean
'    Returns True if data has been save in any table
'    Usefull only for the Unsinstall process.
    
       Dim Con As New ADODB.Connection
       Dim rs  As New ADODB.Recordset
    
       Dim Nbr As Long
       Dim Sql As String
    
      If DataMdb = vbNullString Then
          DataMdb = m_DataMdb_FullPath()
      End If
      Nbr = 0
    
      On Error Resume Next
    
      Con.Open "Provider=" & CodeProject.Connection.Provider & ";Data Source=" & DataMdb
    
      With rs
          .Open "SELECT COUNT(*) AS Nbr FROM [" & c_Mdb_FormTemplate & "]", Con
          Nbr = Nbr + !Nbr
          .Close
          .Open "SELECT COUNT(*) AS Nbr FROM [" & c_Mdb_PersonalObject & "]", Con
          Nbr = Nbr + !Nbr
          .Close
      End With
    
      Con.Close
    
      m_DataMdb_HasData = (Nbr > 0)
    
End Function

Public Function m_DataMdb_Folder(Optional HKey As Long) As String
'    Gives the folder where the DataMdb file should be.
    
       Dim Folder  As String
    
       If HKey = 0 Then
'        zzz        Folder = "" & m_Reg_ValueGetQuick(c_vReg_LocalMachine, m_AccVer(cz_vReg_ProgPath), c_vReg_DataDir)
           Folder = "" & m_Reg_ValueGetQuick(c_vReg_LocalMachine, cz_vReg_ProgPath, c_vReg_DataDir)
       Else
           Folder = "" & m_Reg_ValueGet(HKey, c_vReg_DataDir) ' Null has happend yet in beta versions
      End If
    
      If Len(Folder) < 3 Then
          Folder = m_File_GetFolder(CodeDb.Name)
      Else
          If Not m_Folder_IsExisting(Folder) Then
              Folder = m_File_GetFolder(CodeDb.Name)
          End If
      End If
    
      m_DataMdb_Folder = Folder
    
End Function

Public Function m_DataMdb_FullPath(Optional HKey As Long) As String
'    Gives the full path where the DataMdb file should be.
    
'    zzz    m_DataMdb_FullPath = m_Folder_CheckEnd(m_DataMdb_Folder(HKey), True) & m_AccVer(cz_vDataFile) & m_AccExt()
       m_DataMdb_FullPath = m_Folder_CheckEnd(m_DataMdb_Folder(HKey), True) & cz_vDataFile & m_AccExt()
    
End Function

Public Sub m_DataMdb_Move(OldPath As String, NewFolder As String)
    
       Dim NewPath As String
    
'    zzz    NewPath = NewFolder & "\" & m_AccVer(cz_vDataFile) & m_AccExt()
       NewPath = NewFolder & "\" & cz_vDataFile & m_AccExt()
    
       If OldPath <> NewPath Then
        
           If m_File_IsExisting(OldPath) Then
              If Not m_File_IsExisting(NewPath) Then
                  m_File_Copy OldPath, NewPath
                  m_File_Delete OldPath
              End If
          End If
        
'        zzz        m_Reg_ValueSetQuick c_vReg_LocalMachine, m_AccVer(cz_vReg_ProgPath), c_vReg_DataDir, NewFolder, REG_SZ
          m_Reg_ValueSetQuick c_vReg_LocalMachine, cz_vReg_ProgPath, c_vReg_DataDir, NewFolder, REG_SZ
        
      End If
    
End Sub

Private Function p_Table_Existes(Con As ADODB.Connection, TableName As String) As Boolean
'    Returns True if the table specified does existes in the database
    
       Dim rs As ADODB.Recordset
    
       p_Table_Existes = False
    
       Set rs = Con.OpenSchema(adSchemaTables) 'Open recordset containing table's information
       With rs
           Do Until .EOF Or (p_Table_Existes = True)
              If !TABLE_NAME = TableName Then
                  p_Table_Existes = True
              End If
              .MoveNext
          Loop
          .Close
      End With
    
End Function

Private Sub p_Table_Create(Con As ADODB.Connection, TableName As String)
'    Creates the specified table in the mdb
    
       Dim Sql As String
    
       Select Case TableName
    Case c_Mdb_FormTemplate
        
           Sql = ""
           Sql = Sql & "CREATE TABLE [" & c_Mdb_FormTemplate & "] ("
          Sql = Sql & "[Id]       INTEGER,"         'Long
          Sql = Sql & "[Property] VARCHAR (255),"   'Text
          Sql = Sql & "[PValue]   VARCHAR (255)"    'Text
          Sql = Sql & ") "
          Con.Execute Sql
        
          Sql = "CREATE INDEX [idx_Id] ON [" & c_Mdb_FormTemplate & "] ([Id])"
          Con.Execute Sql
        
      Case c_Mdb_PersonalObject
        
          Sql = ""
          Sql = Sql & "CREATE TABLE [" & c_Mdb_PersonalObject & "] ("
          Sql = Sql & "[Id]      IDENTITY (1,1) CONSTRAINT primarykey PRIMARY KEY,"
          Sql = Sql & "[Type]    SMALLINT,"        'Integer
          Sql = Sql & "[Name]    VARCHAR (255),"   'Text
          Sql = Sql & "[Version] VARCHAR (255),"   'Text
          Sql = Sql & "[Date]    DATETIME,"        'Date
          Sql = Sql & "[Data]    TEXT"             'LongBinary
          Sql = Sql & ") "
          Con.Execute Sql
        
      End Select
    
End Sub