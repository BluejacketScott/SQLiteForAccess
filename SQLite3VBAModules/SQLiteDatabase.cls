VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SQLiteDatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private myDbHandle As LongLong

Property Get version() As String
    version = SQLite3LibVersion()
End Property

Private Sub Class_Initialize()
    Dim InitReturn As Long
    
    #If Win64 Then
        
        InitReturn = SQLite3Initialize(ThisWorkbook.Path + "\x64\")
    #Else
        InitReturn = SQLite3Initialize(ThisWorkbook.Path + "\x86\") 
    #End If
    If InitReturn <> SQLITE_INIT_OK Then
        Debug.Print "Error Initializing SQLite. Error: " & Err.LastDllError
        Exit Sub
    End If
    myDbHandle = 0
End Sub

Private Sub Class_Terminate()
    If myDbHandle <> 0 Then SQLite3Close myDbHandle
End Sub

Public Sub OpenDatabase(fileName As String)
    SQLite3Open fileName, myDbHandle
End Sub

Public Sub CloseDatabase()
    SQLite3Close myDbHandle
End Sub

Public Function ExecuteNonQuery(ByVal SqlCommand As String) As Long
    Dim stmtHandle As LongLong
    Dim RetVal As Long
    
    If myDbHandle = 0 Then Err.Raise 99, , "database not open"
    SQLite3PrepareV2 myDbHandle, SqlCommand, stmtHandle
    ErrorCheck
    SQLite3Step stmtHandle
    SQLite3Finalize stmtHandle
    ExecuteNonQuery = SQLite3Changes(myDbHandle)
End Function

Public Sub CreateQueryDef(ByVal queryName As String, ByVal SqlCommand As String)
    ExecuteNonQuery "CREATE VIEW [" & queryName & "] AS " & SqlCommand
End Sub

Public Sub ImportTable(ByVal TableName As String, ByVal newTableName As String)
    If myDbHandle = 0 Then Err.Raise 99, , "database not open"
    ImportQuery "SELECT * FROM" & TableName, newTableName
End Sub

Public Sub ImportQuery(ByVal SqlCommand As String, ByVal newTableName As String)
    Dim myStmtHandle As LongLong
    Dim rs As DAO.Recordset
        
    If myDbHandle = 0 Then Err.Raise 99, , "database not open"
    SQLite3PrepareV2 myDbHandle, SqlCommand, myStmtHandle
    ErrorCheck
    
    If SQLite3Step(myStmtHandle) = SQLITE_ROW Then
        CreateLocalTable myStmtHandle, newTableName
        Set rs = CurrentDb.OpenRecordset(newTableName, dbOpenTable)
        ImportRecord myStmtHandle, rs
        While SQLite3Step(myStmtHandle) = SQLITE_ROW
            ImportRecord myStmtHandle, rs
        Wend
        rs.Close
    Else
'        Debug.Print "SQLite3Step returned " & RetVal
    End If
    
    SQLite3Finalize myStmtHandle
End Sub

Private Sub CreateLocalTable(ByVal stmtHandle As LongLong, newTableName As String)
    Dim td As DAO.TableDef
    Dim colCount As Long
    Dim colValue As Variant
    Dim i As Long
    
    If myDbHandle = 0 Then Err.Raise 99, , "database not open"
    colCount = SQLite3ColumnCount(stmtHandle)
    
    ' Create Table
    Set td = CurrentDb.CreateTableDef(newTableName)
    For i = 0 To colCount - 1
        td.Fields.Append td.CreateField(SQLite3ColumnName(stmtHandle, i), MapTyp(SQLite3ColumnType(stmtHandle, i)))
    Next
    CurrentDb.TableDefs.Append td
End Sub

Private Sub ImportRecord(ByVal stmtHandle As LongLong, rs As DAO.Recordset)
    Dim colValue As Variant
    Dim colCount As Long
    Dim i As Long
    
    colCount = SQLite3ColumnCount(stmtHandle)
    rs.AddNew
    For i = 0 To colCount - 1
        rs.Fields(i).Value = ColumnValue(stmtHandle, i, SQLite3ColumnType(stmtHandle, i))
    Next
    rs.Update
End Sub

Private Function MapTyp(ByVal SQLiteType As Long) As Variant
    Select Case SQLiteType
        Case SQLITE_INTEGER:
            MapTyp = dbLong
        Case SQLITE_FLOAT:
            MapTyp = dbDouble
        Case SQLITE_TEXT:
            MapTyp = dbText
        Case SQLITE_BLOB:
            MapTyp = dbMemo
        Case SQLITE_NULL:
            MapTyp = dbText
    End Select
End Function

Private Function ColumnValue(ByVal stmtHandle As LongLong, ByVal ZeroBasedColIndex As Long, ByVal SQLiteType As Long) As Variant
    Select Case SQLiteType
        Case SQLITE_INTEGER:
            ColumnValue = SQLite3ColumnInt32(stmtHandle, ZeroBasedColIndex)
        Case SQLITE_FLOAT:
            ColumnValue = SQLite3ColumnDouble(stmtHandle, ZeroBasedColIndex)
        Case SQLITE_TEXT:
            ColumnValue = SQLite3ColumnText(stmtHandle, ZeroBasedColIndex)
        Case SQLITE_BLOB:
            ColumnValue = SQLite3ColumnText(stmtHandle, ZeroBasedColIndex)
        Case SQLITE_NULL:
            ColumnValue = Null
    End Select
End Function

Public Sub ExportQuery(ByVal SqlCommand As String, ByVal newTableName As String)
Dim rs As DAO.Recordset
Dim cols As String
Dim i As Integer
Dim RetVal As Long
Dim myStmtHandle As LongLong

    If myDbHandle = 0 Then Err.Raise 99, , "database not open"
    BeginTrans
    ' Create Table
    Set rs = CurrentDb.OpenRecordset(SqlCommand, dbOpenForwardOnly)
    For i = 0 To rs.Fields.Count - 1
        If cols <> "" Then cols = cols & ","
        cols = cols & " [" & rs.Fields(i).Name & "] " & MapTyp2(rs.Fields(i))
    Next i
    ExecuteNonQuery "CREATE TABLE [" & newTableName & "] (" + cols + ")"
    
    cols = ""
    For i = 0 To rs.Fields.Count - 1
        If cols <> "" Then cols = cols & ","
        cols = cols & "?"
    Next i
   
    
    RetVal = SQLite3PrepareV2(myDbHandle, "INSERT INTO [" & newTableName & "] VALUES(" & cols & ")", myStmtHandle)
    ErrorCheck
    
    While Not rs.EOF
        For i = 0 To rs.Fields.Count - 1
            If Not IsNull(rs.Fields(i).Value) Then
            Select Case rs.Fields(i).Type
                Case dbBigInt, dbBoolean, dbByte, dbInteger, dbLong
                    RetVal = SQLite3BindInt32(myStmtHandle, i + 1, rs.Fields(i).Value)
                Case dbDate, dbTimeStamp, Time
                    RetVal = SQLite3BindText(myStmtHandle, i + 1, Format(rs.Fields(i).Value, "yyyy-MM-dd HH:mm:ss"))
                Case dbCurrency, dbDecimal, dbNumeric, dbDouble, dbFloat, dbSingle
                    RetVal = SQLite3BindDouble(myStmtHandle, i + 1, rs.Fields(i).Value)
                Case dbLongBinary, dbVarBinary
                ' ToDo
                Case Else
                    RetVal = SQLite3BindText(myStmtHandle, i + 1, rs.Fields(i).Value)
            End Select
            End If
        Next i
        RetVal = SQLite3Step(myStmtHandle)
        If RetVal <> SQLITE_DONE Then
           Debug.Print "SQLite3Step returned " & RetVal, SQLite3ErrMsg(myDbHandle)
           Beep
        End If
        RetVal = SQLite3Reset(myStmtHandle)
        If RetVal <> SQLITE_OK Then
            Debug.Print "SQLite3Reset returned " & RetVal, SQLite3ErrMsg(myDbHandle)
            Beep
        End If
        rs.MoveNext
    Wend
    rs.Close
    RetVal = SQLite3Finalize(myStmtHandle)
    CommitTrans
'    MsgBox sqlCreate

End Sub

Sub SaveRecord(ByVal TableName As String, ParamArray Values())
Dim cols As String
Dim i As Integer
Dim RetVal As Long
Dim myStmtHandle As LongLong

    ' Create Placeholder
    cols = ""
    For i = 0 To UBound(Values)
        If cols <> "" Then cols = cols & ","
        cols = cols & "?"
    Next i
    ' Create Insert-Statement
    RetVal = SQLite3PrepareV2(myDbHandle, "INSERT INTO [" & TableName & "] VALUES(" & cols & ")", myStmtHandle)
    ErrorCheck

    For i = 0 To UBound(Values)
       Select Case VarType(Values(i))
            Case vbBoolean, vbInteger, vbLong, vbByte
                RetVal = SQLite3BindInt32(myStmtHandle, i + 1, Values(i))
            Case vbDate
                RetVal = SQLite3BindText(myStmtHandle, i + 1, Format(Values(i), "yyyy-MM-dd HH:mm:ss"))
            Case vbSingle, vbDouble, vbCurrency, vbDecimal
                RetVal = SQLite3BindDouble(myStmtHandle, i + 1, Values(i))
            Case vbEmpty, vbNull, vbObject
                ' ToDo
            Case Else
                RetVal = SQLite3BindText(myStmtHandle, i + 1, Values(i))
        End Select
    Next i
    RetVal = SQLite3Step(myStmtHandle)
    If RetVal <> SQLITE_DONE Then
       Debug.Print "SQLite3Step returned " & RetVal, SQLite3ErrMsg(myDbHandle)
    End If
    RetVal = SQLite3Reset(myStmtHandle)
    If RetVal <> SQLITE_OK Then
        Debug.Print "SQLite3Reset returned " & RetVal, SQLite3ErrMsg(myDbHandle)
    End If
    RetVal = SQLite3Finalize(myStmtHandle)
    
End Sub



Private Function MapTyp2(fld As DAO.Field) As String
    If (fld.Attributes And dbAutoIncrField) = dbAutoIncrField Then
       MapTyp2 = "INTEGER NOT NULL PRIMARY KEY AUTOINCREMENT"
    Else
    Select Case fld.Type
        Case dbBigInt, dbBoolean, dbByte, dbInteger, dbLong
            MapTyp2 = "INTEGER"
        Case dbCurrency, dbDecimal, dbNumeric
            MapTyp2 = "DECIMAL"
        Case dbDate, dbTimeStamp
            MapTyp2 = "DATETIME"
        Case dbTime
            MapTyp2 = "TIME"
        Case dbDouble, dbFloat, dbSingle
            MapTyp2 = "DOUBLE"
        Case dbGUID
            MapTyp2 = "GUID"
        Case dbLongBinary, dbVarBinary
            MapTyp2 = "BLOB"
        Case dbMemo:
            MapTyp2 = "MEMO"
        Case Else
            MapTyp2 = "TEXT"
    End Select
    End If
End Function

Public Function Backup(ByVal fileName As String) As Long
    Dim myDbBackupHandle As LongLong
    Dim myBackupHandle As LongLong

    SQLite3Open fileName, myDbBackupHandle
    myBackupHandle = SQLite3BackupInit(myDbBackupHandle, "main", myDbHandle, "main")
    If myBackupHandle <> 0 Then
        SQLite3BackupStep myBackupHandle, -1
        SQLite3BackupFinish myBackupHandle
    End If
    SQLite3ErrCode myDbBackupHandle
    SQLite3Close myDbBackupHandle

End Function

Public Sub Compact()
    SQLite3ExecuteNonQuery myDbHandle, "VACUUM;"
End Sub


Public Sub BeginTrans()
    SQLite3ExecuteNonQuery myDbHandle, "BEGIN TRANSACTION"
End Sub

Public Sub CommitTrans()
    SQLite3ExecuteNonQuery myDbHandle, "COMMIT TRANSACTION"
End Sub

Public Sub RollbackTrans()
    SQLite3ExecuteNonQuery myDbHandle, "ROLLBACK TRANSACTION"
End Sub


Private Sub ErrorCheck()
Dim errcode As Long
    errcode = SQLite3ErrCode(myDbHandle)
    If errcode <> 0 Then Err.Raise errcode, , SQLite3ErrMsg(myDbHandle)
End Sub

