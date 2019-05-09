Dim connStr, objConn, getNames

connStr = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source= C:\Users\zizipho.a.mbolekwa\Documents\Test.accdb"

'Define object type
Set objConn = CreateObject("ADODB.Connection")
Set objRecordSet = CreateObject("ADODB.Recordset")

'Open Connection
objConn.open connStr
Set objRecordSet = objConn.OpenSchema(20)

Do Until objRecordset.EOF

    Wscript.Echo "Table name: " & objRecordset.Fields.Item("TABLE_NAME")

    ' Wscript.Echo "Table type: " & objRecordset.Fields.Item("TABLE_TYPE")

    Wscript.Echo

    objRecordset.MoveNext

Loop
