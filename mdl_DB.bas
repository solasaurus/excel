Attribute VB_Name = "mdl_DB"
Option Explicit

'   Here 's some sample code that I use to connect to CSV files as a database, it allows multiple CSV files in the same folder to be treated as a relational database.
'   It has come in handy quite a few times:
'
'Sub CSV_DB()
'    Sheets("Sheet2").Cells.Clear
'    Dim db As Object
'    Dim rs As Object
'
'    Set db = CreateObject("ADODB.Connection")
'    Set rs = CreateObject("ADODB.Recordset")
'
'    Dim File_Path As String: File_Path = "C:\Users\myname\Desktop\Some folder\Data\"
'    Dim File_Name As String: File_Name = "somefile.csv"
'
'    db.Provider = "Microsoft.Jet.OLEDB.4.0"
'    db.Open "Data Source=" & File_Path & ";" & "Extended Properties=""text;HDR=Yes;FMT=Delimited;"""
'
'    rs.Open "SELECT * FROM [" & File_Name & "]", db
'    rs.MoveFirst
'    Worksheets("Sheet2").Cells(1, 1).CopyFromRecordset rs
'
'    rs.Close: Set rs = Nothing
'    db.Close: Set db = Nothing
'
'End Sub
