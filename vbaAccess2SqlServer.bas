Attribute VB_Name = "vbaAccess2SqlServer"
Option Compare Database
Option Explicit

'Tools | References and add the following:
'Microsoft ActiveX Data Objects 2.8 Library (the 2.1 library may already be checked)
'Microsoft ADO Ext. 2.8 for DDL and Security
'Microsoft Scripting Runtime
Public Enum DestType
  [ONE] = 1   'all ddl statements are put in DEST_FILENAME
  [EACH] = 2  'ddl file named for [table].sql put in DEST_PATH
End Enum

Public Function access2sqlserver_ddl()
'This function creates ddl file(s) for Sql Server based on local tables in the mdb
'Access only has one date time that is timestamp, so any dates will be DATETIME
  Dim DEST_TYPE As DestType
  Const ACCESS_NAME_PATTERN As String = "%"
  Const SVR_CONNECTION_STRING As String = "<<CONNECTION STRING HERE>>"
  Const DEST_PATH As String = "\" 'from current path
  Const DEST_FILENAME As String = DEST_PATH & "ddl.sql"
  Const SCHEMA As String = "dbo"
  
  Dim cnn As ADODB.Connection
  Dim rst As ADODB.Recordset, rsttbl As ADODB.Recordset
  Dim fld As ADODB.Field
  Dim cat As ADOX.Catalog
  Dim tbl As ADOX.Table
  Dim col As ADOX.Column
  Dim strType As String, strNull As String, strSql As String
  Dim iMaxNameLen As Integer, iMaxDataTypeLen As Integer, iColNum As Integer
  Dim fso As FileSystemObject
  Dim ts As TextStream
  Dim idxs As String, pkcols As String, hasIdentity As Boolean
  Dim svr As New ADODB.Connection
  
  DEST_TYPE = ONE
  Set fso = New FileSystemObject
  If Not fso.FolderExists(CurrentProject.Path & DEST_PATH) Then
    fso.CreateFolder CurrentProject.Path & DEST_PATH
  End If
  If DEST_TYPE = DestType.ONE Then
    Set ts = fso.CreateTextFile(CurrentProject.Path & DEST_FILENAME, True, False)
  End If
  Set cnn = CurrentProject.Connection
  Set rsttbl = New ADODB.Recordset
  Set cat = New ADOX.Catalog
  cat.ActiveConnection = cnn
  Set rst = New ADODB.Recordset
  svr.Open SVR_CONNECTION_STRING
  rst.Open "SELECT [Name] FROM [MSysObjects] WHERE [Name] like '" & ACCESS_NAME_PATTERN & "' AND [Name] not like 'MSys%' AND [Type]=1", cnn, adOpenForwardOnly, adLockReadOnly, adCmdText
  While Not rst.EOF
    hasIdentity = False
    Set tbl = get_table(cat, rst!Name)
    If Not tbl Is Nothing Then
        If DEST_TYPE = DestType.EACH Then
          Set ts = fso.CreateTextFile(CurrentProject.Path & DEST_PATH & rst!Name & ".sql", True, False)
        End If
        idxs = get_indexes(tbl, pkcols)
        ts.WriteLine "IF OBJECT_ID(N'[" & SCHEMA & "].[" & rst!Name & "]', N'U') IS NOT NULL"
        ts.WriteLine "  BEGIN"
        ts.WriteLine "    DROP TABLE [" & SCHEMA & "].[" & rst!Name & "]"
        ts.WriteLine "    PRINT N'Dropped [" & SCHEMA & "].[" & rst!Name & "]'"
        ts.WriteLine "  END"
        ts.WriteLine "PRINT N'Creating [" & SCHEMA & "].[" & rst!Name & "]'"
        ts.WriteLine "GO"
        ts.WriteLine "CREATE TABLE [" & SCHEMA & "].[" & rst!Name & "] ("
        iMaxNameLen = 0
        iMaxDataTypeLen = 0
        'determine spacing so it prints nicely
        For Each col In tbl.Columns
          If Len(col.Name) > iMaxNameLen Then
            iMaxNameLen = Len(col.Name)
          End If
          If Len(get_datatype(col)) > iMaxDataTypeLen Then
            iMaxDataTypeLen = Len(get_datatype(col))
          End If
        Next col
        iMaxNameLen = iMaxNameLen + 3
        iMaxDataTypeLen = iMaxDataTypeLen + 3
        rsttbl.Open CStr(rst!Name), cnn, adOpenDynamic, adLockOptimistic, adCmdTableDirect
        iColNum = 0
        For Each fld In rsttbl.Fields
          Set col = get_column(tbl, fld.Name)
          iColNum = iColNum + 1
          strType = IIf(iColNum > 1, "     , ", "       ") & "[" & col.Name & "]" & Space(iMaxNameLen - Len(col.Name))
          strNull = IIf(Not col.Properties("Nullable") Or col.Properties("Autoincrement"), "NOT", "   ") & " NULL"
          'Access lets PK fields be null: don't do that
          If InStr(1, pkcols, "[" & col.Name & "]") > 0 Then strNull = "NOT NULL"
          If col.Properties("Autoincrement") Then
            strNull = strNull & " IDENTITY(1, " & col.Properties("Increment") & ")"
            hasIdentity = True
          End If
          If Not IsEmpty(col.Properties("Default")) Then
            strNull = strNull & " DEFAULT " & col.Properties("Default")
          End If
          ts.WriteLine strType & Left(get_datatype(col) & Space(iMaxDataTypeLen), iMaxDataTypeLen) & strNull
        Next fld
        rsttbl.Close
        ts.WriteLine ")" & vbCrLf & "GO" & vbCrLf
        ts.WriteLine idxs
        If DEST_TYPE = DestType.EACH Then
          ts.Close
        End If
        'Uncomment to load
        'Debug.Print rst!Name & " : " & DataLoadDbo(svr, rst!Name, hasIdentity)
    End If
    rst.MoveNext
  Wend
  If DEST_TYPE = DestType.ONE Then
    ts.Close
  End If
  rst.Close
  svr.Close
  Set rst = Nothing
  Set rsttbl = Nothing
  Set cat = Nothing
  Set fso = Nothing
  Set svr = Nothing
End Function

Private Function DataLoadDbo(ByRef svr As ADODB.Connection, ByRef tblName As String, ByRef hasIdentity As Boolean) As Long
  Dim cnn As ADODB.Connection
  Dim dbo As New ADODB.Recordset, tbl As New ADODB.Recordset
  Dim idx As Integer
  
  svr.Execute "TRUNCATE TABLE [" & tblName & "]"
  If hasIdentity Then svr.Execute "SET IDENTITY_INSERT [" & tblName & "] ON"
  dbo.Open tblName, svr, adOpenKeyset, adLockOptimistic, adCmdTableDirect
  Set cnn = CurrentProject.Connection
  tbl.Open tblName, cnn, adOpenKeyset, adLockOptimistic, adCmdTableDirect
  While Not tbl.EOF
    dbo.AddNew
    For idx = 0 To tbl.Fields.Count - 1
      dbo.Fields(idx) = tbl.Fields(idx).Value
    Next idx
    dbo.Update
    DataLoadDbo = DataLoadDbo + 1
    tbl.MoveNext
  Wend
  dbo.Close
  tbl.Close
  If hasIdentity Then svr.Execute "SET IDENTITY_INSERT [" & tblName & "] OFF"
  Set dbo = Nothing
  Set tbl = Nothing
End Function

Private Function get_indexes(ByRef tbl As ADOX.Table, ByRef pkcols As String) As String
  Dim indx As ADOX.Index
  Dim col As ADOX.Column
  Dim ddl As String, cols As String
  
  For Each indx In tbl.Indexes
    cols = "["
    For Each col In indx.Columns
      cols = cols & col.Name & "], ["
    Next col
    cols = Left(cols, Len(cols) - 3)
    If indx.PrimaryKey Then pkcols = "," & cols & ","
    If indx.PrimaryKey Then
      ddl = ddl & "ALTER TABLE dbo." & tbl.Name & " ADD CONSTRAINT PK_" & tbl.Name & " PRIMARY KEY CLUSTERED (" & cols & ")" & vbCrLf
    ElseIf indx.Unique Then
      ddl = ddl & "CREATE UNIQUE NONCLUSTERED INDEX " & indx.Name & " ON dbo." & tbl.Name & " (" & cols & ")" & vbCrLf
      If (indx.IndexNulls And adIndexNullsAllow) = adIndexNullsAllow Then
        ddl = ddl & "WHERE " & Replace(cols, ", ", " IS NOT NULL" & vbCrLf & "  AND ") & " IS NOT NULL" & vbCrLf
      End If
    Else
      ddl = ddl & "CREATE NONCLUSTERED INDEX " & indx.Name & " ON dbo." & tbl.Name & " (" & cols & ")" & vbCrLf
    End If
    ddl = ddl & "GO" & vbCrLf
  Next indx
  get_indexes = ddl
End Function

Private Function get_table(ByRef cat As ADOX.Catalog, ByRef table_name As String) As ADOX.Table
  Dim tbl As ADOX.Table
  For Each tbl In cat.Tables
    If tbl.Name = table_name Then
      Set get_table = tbl
      Exit For
    End If
  Next tbl
End Function

Private Function get_column(ByRef tbl As ADOX.Table, ByRef column_name As String) As ADOX.Column
  Dim col As ADOX.Column
  For Each col In tbl.Columns
    If col.Name = column_name Then
      Set get_column = col
      Exit For
    End If
  Next col
End Function

Private Function get_datatype(ByRef col As ADOX.Column) As String
  Select Case col.Type
    Case Is = adBinary
      get_datatype = "VARBINARY(" & col.DefinedSize & ")"
    Case Is = adBoolean
      get_datatype = "BIT"
    Case Is = adCurrency
      get_datatype = "MONEY"
    Case Is = adDate
      get_datatype = "DATETIME"
    Case Is = adDouble
      get_datatype = "DOUBLE PRECISION" 'aka FLOAT(53)
    Case Is = adGUID
      get_datatype = "UNIQUEIDENTIFIER"
    Case Is = adInteger
      get_datatype = "INT"
    Case Is = adLongVarBinary
      get_datatype = "VARBINARY(MAX)"
    Case Is = adLongVarWChar
      get_datatype = "VARCHAR(MAX)"
    Case Is = adNumeric
      get_datatype = "DECIMAL(" & col.Precision & "," & col.NumericScale & ")"
    Case Is = adSingle
      get_datatype = "REAL" 'aka FLOAT(24)
    Case Is = adSmallInt
      get_datatype = "SMALLINT"
    Case Is = adUnsignedTinyInt
      get_datatype = "TINYINT"
    Case Is = adVarBinary
      get_datatype = "VARBINARY(" & col.DefinedSize & ")"
    Case Is = adVarWChar
      get_datatype = "VARCHAR(" & col.DefinedSize & ")"
    Case Is = adWChar
      get_datatype = "CHAR(" & col.DefinedSize & ")"
    Case Else
      Err.Raise vbObjectError + 3817, "f(n)", "This access type isn't suppose to be supported:  " & col.Type
  End Select
End Function
