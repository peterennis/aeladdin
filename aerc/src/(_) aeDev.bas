Option Compare Database
Option Explicit

Public Function VBAversion() As String
'    Ref: http://msdn.microsoft.com/en-us/library/ff700513(office.11).aspx
    
    #If VBA7 Then
       VBAversion = "VBA7"
    #ElseIf VBA6 Then
       VBAversion = "VBA6"
    #Else
       VBAversion = "VBA Not 6 or 7"
    #End If
    
End Function

Public Function Office32or64() As String
'    Ref: http://msdn.microsoft.com/en-us/library/ff700513(office.11).aspx
    
    #If Win64 Then
'    Debug.Print "64"
       Office32or64 = "64"
    #Else
'    Debug.Print "32"
       Office32or64 = "32"
    #End If
    
End Function

Public Sub GetReferences()
'    Ref: http://vbadud.blogspot.com/2008/04/get-references-of-vba-project.html
'    Ref: http://www.pcreview.co.uk/forums/type-property-reference-object-vbulletin-project-t3793816.html
    
       Dim i As Integer
       Dim RefName As String
       Dim RefDesc As String
       Dim blnRefBroken As Boolean
       Dim vbaProj As Object
    
      Set vbaProj = Application.VBE.ActiveVBProject
      Debug.Print vbaProj.Name
      Debug.Print "vbaProj.Type='" & vbaProj.Type & "'"
'    Display the versions of Access, ADO and DAO
      Debug.Print "Access version = " & Application.VERSION
      Debug.Print "ADO (ActiveX Data Object) version = " & CurrentProject.Connection.VERSION
      Debug.Print "DAO (DbEngine)  version = " & Application.DBEngine.VERSION
      Debug.Print "DAO (CodeDb)    version = " & Application.CodeDb.VERSION
      Debug.Print "DAO (CurrentDb) version = " & Application.CurrentDb.VERSION
      Debug.Print "References:"
    
      For i = 1 To vbaProj.References.Count
        
          blnRefBroken = False
        
'        Get the Name of the Reference
          RefName = vbaProj.References(i).Name
        
'        Get the Description of Reference
          RefDesc = vbaProj.References(i).Description
        
          Debug.Print vbaProj.References(i).Name, vbaProj.References(i).Description
        
'        Returns a Boolean value indicating whether or not the Reference object points to a valid reference in the registry. Read-only.
          If Application.VBE.ActiveVBProject.References(i).IsBroken = True Then
              blnRefBroken = True
              Debug.Print vbaProj.References(i).Name, "blnRefBroken=" & blnRefBroken
          End If
      Next
    
End Sub