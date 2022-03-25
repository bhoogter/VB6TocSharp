# Recordset

Simulates most of the interface for the ADODB recordset (just continue extending if you need something that's missing).

## Create a new Recordset

- `Recordset RS = new Recordset()`
- `Recordset RS = new Recordset(SQL, [File], Q[uietErrors], [Parameters])`

## Record Access

- `RS.Fields("Name").Value`
- `RS["Name"]`

