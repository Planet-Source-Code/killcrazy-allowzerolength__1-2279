<div align="center">

## AllowZeroLength


</div>

### Description

All fields in the selected table are processed and the AllowZeroLength property of the fields are set to either True or False, depending on the Status given as the finaal parameter The function returns a boolean value that can be used by the user to determin other operations.
 
### More Info
 
strDatabase = Full database path

strTableName = Name of table to be processed

Status : True / False

True/False


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Killcrazy](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/killcrazy.md)
**Level**          |Unknown
**User Rating**    |4.3 (166 globes from 39 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Databases/ Data Access/ DAO/ ADO](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases-data-access-dao-ado__1-6.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/killcrazy-allowzerolength__1-2279/archive/master.zip)





### Source Code

```
Function AllowZeroLength(strDatabase As String, strtablename As String, status As Boolean) As Boolean
Dim db As Database
Dim td As TableDef
Dim fd As Field
On Error GoTo Error_Handler
Set db = OpenDatabase(strDatabase)
Set td = db.TableDefs(strtablename)
  'loop through the fields in the selected recordset
  For Each fd In td.Fields
    'Check the field type, and only change the value of text and memo fields
    If fd.Type = dbText Or dbMemo Then
      If status = True Then
         fd.AllowZeroLength = True
      Else
        fd.AllowZeroLength = False
      End If
    End If
  Next fd
  AllowZeroLength = True
  ' Exit Early to avoid error handler.
  Exit Function
Error_Handler:
  ' Raise an error.
  Err.Raise Err.Number, "AllowZeroLength", "Could not process fields.", Err.Description
  AllowZeroLength = False
  ' Reset normal error checking.
  Resume Next
End Function
```

