<div align="center">

## GetDataTableFromExcel\(\) Function


</div>

### Description

Function that returns ADO.NET DataTable from MS Excel file using Microsoft.Jet.OLEDB.4.0 Provider.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Igor Krupitsky](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/igor-krupitsky.md)
**Level**          |Intermediate
**User Rating**    |4.8 (19 globes from 4 users)
**Compatibility**  |VB\.NET
**Category**       |[System Services/ Functions](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/system-services-functions__10-23.md)
**World**          |[\.Net \(C\#, VB\.net\)](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/net-c-vb-net.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/igor-krupitsky-getdatatablefromexcel-function__10-4744/archive/master.zip)





### Source Code

```
Private Function GetDataTableFromExcel(ByVal sFilePath As String) As System.Data.DataTable
    Dim sConnectionString As String = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
    "Data Source=" + sFilePath + ";" & _
    "Extended Properties=""Excel 8.0;"""
    Dim cn As New OleDb.OleDbConnection(sConnectionString)
    cn.Open()
    Dim oTables As DataTable = cn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, Nothing)
    Dim i As Integer
    For i = 0 To oTables.Rows.Count - 1
      Dim sSheetName As String = oTables.Rows(i)("TABLE_NAME").ToString()
      If sSheetName.IndexOf("$") <> -1 Then
        Dim oDataSet As New DataSet
        Dim oAdapter As New OleDbDataAdapter("SELECT * FROM [" + sSheetName + "]", cn)
        oAdapter.TableMappings.Add("Table", sSheetName)
        oAdapter.Fill(oDataSet)
        Dim oDataTable As DataTable = oDataSet.Tables(0)
        cn.Close()
        If oDataTable.Rows.Count > 0 And oDataTable.Columns.Count > 5 Then
          Return oDataTable
        End If
      End If
    Next
  End Function
```

