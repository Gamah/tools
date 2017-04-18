Sub main()
    Dim ServerName As String
    Dim Database As String
    
    ServerName = InputBox("Enter Server name\instance or IP:", "Server Select", "localhost")
    Database = InputBox("Enter Database name:", "DB Select", "master")
    Call CleanSheets
    Call GetData(ServerName, Database)
    Call SplitSheets
    Call LinkSheets
End Sub

Private Sub CleanSheets()
    Dim ws As Worksheet
     
    On Error Resume Next
    Set ws = Worksheets("TableIndex")
    On Error GoTo 0
    If Not ws Is Nothing Then
        Application.DisplayAlerts = False
        For Each ws In ThisWorkbook.Worksheets
            If Not ws.Name = "TableIndex" Then ws.Delete
        Next ws
        Application.DisplayAlerts = True
    End If
    
    Sheets("TableIndex").Cells.Clear
     
End Sub

Private Sub GetData(ServerName As String, Database As String)
'Initializes variables

Dim cnn As New ADODB.Connection
Dim rst As New ADODB.Recordset
Dim ConnectionString As String
Dim StrQuery As String
Dim lstr As String
    ConnectionString = "Provider=SQLOLEDB;Data Source=" & ServerName & ";Initial Catalog=" & Database & ";Trusted_connection=yes;"

    cnn.Open ConnectionString
    cnn.CommandTimeout = 900
    
    lstr = lstr + "with basedata as                                                                                                             "
    lstr = lstr + "(                                                                                                                            "
    lstr = lstr + "select                                                                                                                       "
    lstr = lstr + "                                                                                                                             "
    lstr = lstr + " col.name                                as [ColumnName],                                                                    "
    lstr = lstr + " right('0000'+ rtrim(column_id), 4)      as [ColumnID],                                                                      "
    lstr = lstr + " typ.name                                as [Type],                                                                          "
    lstr = lstr + " convert(varchar(20),col.max_length)     as [MaxLength],                                                                     "
    lstr = lstr + " convert(varchar(20),col.precision)      as [Precision],                                                                     "
    lstr = lstr + " convert(varchar(20),col.scale)          as [Scale],                                                                         "
    lstr = lstr + " isnull(com.text,'')                     as [DefaultValue],                                                                  "
    lstr = lstr + " col.is_nullable                         as [IsNullable],                                                                    "
    lstr = lstr + " col.is_identity                         as [IsIdentity],                                                                    "
    lstr = lstr + "    col.is_computed                          as [IsComputed],                                                                "
    lstr = lstr + "    isnull(comcomp.text,'')                  as [ComputedFormula],                                                           "
    lstr = lstr + "    isnull(ext.value,'')                 as [Description],                                                                   "
    lstr = lstr + " sch.name + '.' + tbl.name               as [TableName],                                                                      "
    lstr = lstr + " isnull(exttable.value,'')               as [TableDescription]                                                              "
    lstr = lstr + "from                                                                                                                         "
    lstr = lstr + "       sys.columns col                                                                                                       "
    lstr = lstr + "       join sys.tables tbl on tbl.object_id = col.object_id                                                                  "
    lstr = lstr + "       join sys.schemas sch on sch.schema_id = tbl.Schema_id                                                                 "
    lstr = lstr + "       join sys.types typ on typ.user_type_id = col.user_type_id                                                             "
    lstr = lstr + "       left join sys.objects def on def.object_id = col.default_object_id                                                    "
    lstr = lstr + "       left join sys.syscomments com on com.id = def.object_id                                                               "
    lstr = lstr + "       left join sys.syscomments comcomp on comcomp.id = col.object_id                                                       "
    lstr = lstr + "       and comcomp.number = col.column_id                                                                                    "
    lstr = lstr + "       outer apply fn_listextendedproperty ('MS_Description', 'schema', sch.name, 'table', tbl.name, 'column', col.name) ext "
    lstr = lstr + "       outer apply fn_listextendedproperty (default, 'schema', sch.name, 'table', tbl.name, 'column', col.name) extformula   "
    lstr = lstr + "    outer apply fn_listextendedproperty ('MS_Description', 'schema', sch.name, 'table', tbl.name, null, default) exttable    "
    lstr = lstr + ")                                                                                                                            "
    lstr = lstr + "SELECT * FROM basedata                                                                                                       "
    lstr = lstr + "ORDER BY TableName, ColumnID                                                                                                 "
    StrQuery = lstr
    rst.Open StrQuery, cnn
    For iCols = 0 To rst.Fields.Count - 1
        Worksheets("TableIndex").Cells(1, iCols + 1).Value = rst.Fields(iCols).Name
    Next
    Sheets(1).Range("A2").CopyFromRecordset rst
End Sub

Private Sub SplitSheets()
Dim lr As Long
Dim ws As Worksheet
Dim vcol, i As Integer
Dim icol As Long
Dim myarr As Variant
Dim title As String
Dim titlerow As Integer
vcol = 13
Set ws = Sheets("TableIndex")
lr = ws.Cells(ws.Rows.Count, vcol).End(xlUp).Row
title = "A1:N1"
titlerow = ws.Range(title).Cells(1).Row
icol = ws.Columns.Count
ws.Cells(1, icol) = "Unique"
For i = 2 To lr
    On Error Resume Next
    If ws.Cells(i, vcol) <> "" And Application.WorksheetFunction.Match(ws.Cells(i, vcol), ws.Columns(icol), 0) = 0 Then
        ws.Cells(ws.Rows.Count, icol).End(xlUp).Offset(1) = ws.Cells(i, vcol)
    End If
Next
myarr = Application.WorksheetFunction.Transpose(ws.Columns(icol).SpecialCells(xlCellTypeConstants))
ws.Columns(icol).Clear
For i = 2 To UBound(myarr)
    ws.Range(title).AutoFilter Field:=vcol, Criteria1:=myarr(i) & ""
    If Not Evaluate("=ISREF('" & myarr(i) & "'!A1)") Then
        Sheets.Add(after:=Worksheets(Worksheets.Count)).Name = "Table" & Worksheets.Count
    Else
        Sheets(myarr(i) & "").Move after:=Worksheets(Worksheets.Count)
    End If
ws.Range("A" & titlerow & ":A" & lr).EntireRow.Copy Sheets("Table" & Worksheets.Count - 1).Range("A4")


Sheets("Table" & Worksheets.Count - 1).Range("M4:N5").Cut (Sheets("Table" & Worksheets.Count - 1).Range("A1:B2"))
Sheets("Table" & Worksheets.Count - 1).Range("C1").Select
ActiveCell.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:= _
        "'TableIndex'" & "!A1", TextToDisplay:="Go Back"
ActiveSheet.Range("M:M").Delete
ActiveSheet.Columns.AutoFit

Next
ws.AutoFilterMode = False
ws.Activate
End Sub
Private Sub LinkSheets()
Dim sh As Worksheet
Dim cell As Range
Sheets("TableIndex").Cells.Clear
Sheets("TableIndex").Range("A1").Value = "Table:"
Sheets("TableIndex").Range("B1").Value = "Desc:"
Sheets("TableIndex").Range("A2").Select
For Each sh In ActiveWorkbook.Worksheets
    If ActiveSheet.Name <> sh.Name Then
        ActiveCell.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:= _
        "'" & sh.Name & "'" & "!A1", TextToDisplay:=sh.Range("A2").Value
        ActiveCell.Offset(0, 1).Value = sh.Range("B2").Value
        ActiveCell.Offset(1, 0).Select
   End If
Next sh
Sheets("TableIndex").Columns.AutoFit
Sheets("TableIndex").Range("A1").Select
End Sub

