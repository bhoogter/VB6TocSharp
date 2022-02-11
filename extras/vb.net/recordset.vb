Imports System.Data
Imports System.Data.OleDb

' Recordset object.  Used to wrap other data objects to 'simulate' VB6.

Public Class Recordset
    Public Source As String = ""
    Public Database As String = ""
    Public QuietErrors As Boolean = False

    Private mAddingRow As Boolean = False
    Public ReadOnly Property AddingRow As Boolean
        Get
            Return mAddingRow
        End Get
    End Property

    Private connection As OleDbConnection
    Private adapter As OleDbDataAdapter
    Private table As DataTable
    Private filteredTable As DataTable
    Private mFilter As String

    Overloads Sub Finalize()
        Close()
    End Sub

    Public Sub New()

    End Sub


    Public Sub New(table As DataTable, adapter As OleDbDataAdapter, connection As OleDbConnection)
        Me.connection = connection
        Me.adapter = adapter
        Me.table = table
    End Sub


    Public Sub New(SQL As String, File As String, Optional QuietErrors As Boolean = False)
        Source = SQL
        Database = File
        Me.QuietErrors = QuietErrors

        Open()
    End Sub

    Public Sub Close()
        Try
            connection?.Close()
        Catch
            ' just suppression
        End Try

        connection = Nothing
        adapter = Nothing
        table = Nothing
        filteredTable = Nothing
    End Sub

    Public Shared Sub sqlExecutionError(mSQL As String, e As Exception)
        Dim T As String = ""
        T &= "getRecordSet Failed: " & e.Message & vbCrLf
        T &= vbCrLf
        T &= mSQL & vbCrLf
        T &= vbCrLf
        T &= "ERROR:" & e.Message

        T = T.Replace("$EDESC", e.Message)
        'ErrMsg = Replace(ErrMsg, "$ENO", Err().Number)
        T = T.Replace("$ESRC", e.Source)
        MsgBox("Database Error: " + T, 0, "Error")
        'CheckStandardErrors() ' Bookmark/updateable query
    End Sub

    Private Function ConnectionString(file As String) As String
        Return "PROVIDER=Microsoft.Jet.OLEDB.4.0Data Source=" + file
    End Function

    Public Property AbsolutePosition As Integer = -1
    Public Property Position As Integer
        Get
            Return AbsolutePosition
        End Get
        Set(value As Integer)
            AbsolutePosition = value
        End Set
    End Property
    Public ReadOnly Property RecordCount As Integer
        Get
            If Not table Is Nothing Then
                If Not table.Rows Is Nothing Then
                    Return table.Rows.Count
                End If
            End If
            Return 0
        End Get
    End Property
    Public ReadOnly Property EOF As Boolean
        Get
            Return AbsolutePosition >= RecordCount
        End Get
    End Property
    Public ReadOnly Property BOF As Boolean
        Get
            Return AbsolutePosition = 0
        End Get
    End Property

    Public Function FieldExists(F As String) As Boolean
        If Not table Is Nothing Then
            If Not table.Columns Is Nothing Then
                Return table.Columns.Contains(F)
            End If
        End If
        Return False
    End Function


    Public Function MoveFirst() As Integer
        AbsolutePosition = 0
        Return 0
    End Function
    Public Function MoveNext() As Integer
        Return If(++AbsolutePosition < RecordCount, AbsolutePosition, AbsolutePosition = RecordCount)
    End Function
    Public Function MovePrevious() As Integer
        Return If(--AbsolutePosition >= 0, AbsolutePosition, AbsolutePosition = 0)
    End Function
    Public Function MoveLast() As Integer
        AbsolutePosition = RecordCount - 1
        Return AbsolutePosition
    End Function

    Public ReadOnly Property Fields As RecordsetFields
        Get
            If AbsolutePosition >= 0 And AbsolutePosition < RecordCount Then Return New RecordsetFields(table.Rows(AbsolutePosition))
            Throw New ArgumentOutOfRangeException("Either EOF or BOF is true.")
        End Get
    End Property

    Public ReadOnly Property FieldNames As List(Of String)
        Get
            If IsNothing(table) Then Return Nothing
            Dim result As List(Of String) = New List(Of String)
            For Each item As DataColumn In table.Columns
                result.Add(item.ColumnName)
            Next
            Return result
        End Get
    End Property

    Public ReadOnly Property Field As PropIndexer(Of Object, Object)
        Get
            Return New PropIndexer(Of Object, Object)(
                Function(k As Object)
                    Return Fields(k).Value
                End Function,
                Function(k As Object, v As Object)
                    Fields(k).Value = v
                End Function
                )
        End Get
    End Property

    Default Property Item(field As Object) As Object
        Get
            Return GetField(field)
        End Get
        Set
            SetField(field, Value)
        End Set
    End Property

    Public Function GetField(key As Object) As Object
        Return Fields(key).Value
    End Function
    Public Sub SetField(key As Object, value As Object)
        Fields(key).Value = value
    End Sub

    Public Function GetRows() As List(Of List(Of Object))
        Dim tableEnumerable As Object = table.AsEnumerable()
        Dim tableList As Object = tableEnumerable.ToArray().ToList()
        Return tableList.ToList() _
        .Select(Function(r As Object)
                    Return r.ItemArray.ToList()
                End Function) _
        .ToList()
    End Function

    Public Property Filter As String
        Get
            Return mFilter
        End Get
        Set(value As String)
            mFilter = value
            If String.IsNullOrEmpty(value) Then
                filteredTable = Nothing
                Return
            End If

            filteredTable = table.Select(mFilter).CopyToDataTable()
        End Set
    End Property

    Protected Function Find(v As String) As Boolean
        Dim temp As DataTable = table.Select(mFilter).CopyToDataTable()
        If temp.Rows.Count = 0 Then Return False
        Dim X As Integer = table.Rows.IndexOf(temp.Rows(0))
        AbsolutePosition = X
        Return True
    End Function

    Private Sub Open()
        Const maxTries = 5

        If Dir(Database) = "" Then
            MsgBox("Database Not Found: " + Database)
            Return
        End If

        Dim result As DataSet = New DataSet()
        connection = New OleDbConnection(ConnectionString(Database))
        Dim Command As OleDbCommand = New OleDbCommand(Source, connection)
        adapter = New OleDbDataAdapter(Command)
        Try
            connection.Open()
            adapter.FillSchema(result, SchemaType.Source)
            adapter.Fill(result, "Default")
        Catch e As Exception
            If Not QuietErrors Then sqlExecutionError(Source, e)
        Finally
            connection.Close()
        End Try

        table = result.Tables("Default")
    End Sub

    Public Sub Update()
        Dim cb As OleDbCommandBuilder = New OleDbCommandBuilder(adapter)
        cb.QuotePrefix = "["
        cb.QuoteSuffix = "]"
        Try
            connection.Open()
            adapter.UpdateCommand = cb.GetUpdateCommand()
            adapter.Update(table)
        Catch e As Exception
            If Not QuietErrors Then sqlExecutionError(adapter.DeleteCommand.ToString(), e)
        Finally
            connection.Close()
        End Try

        mAddingRow = False
    End Sub

    Public Sub AddNew()
        Dim newRow As DataRow = table.NewRow()
        table.Rows.InsertAt(newRow, table.Rows.Count)
        AbsolutePosition = table.Rows.Count - 1
        mAddingRow = True
    End Sub

    Public Sub Delete()
        Dim cb As OleDbCommandBuilder = New OleDbCommandBuilder(adapter)
        Try
            connection.Open()
            adapter.DeleteCommand = cb.GetDeleteCommand()
            adapter.Update(table)
        Catch e As Exception
            If Not QuietErrors Then sqlExecutionError(adapter.UpdateCommand.ToString(), e)
        Finally
            connection.Close()
        End Try
    End Sub

    Public Class RecordsetFields
        Implements ICollection
        Private row As DataRow = Nothing

        Public Sub New(row As DataRow)
            Me.row = row
        End Sub

        Public ReadOnly Property Count As Integer
            Get
                Return row.Table.Columns.Count
            End Get
        End Property
        Public SyncRoot As Object = Nothing
        Public IsSynchronized As Boolean = False

        Private Sub ICollection_CopyTo(array As Array, index As Integer) Implements ICollection.CopyTo
            Throw New InvalidOperationException("Not valid on object")
        End Sub

        Private Function IEnumerable_GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
            Return row.Table.Columns.GetEnumerator()
        End Function

        Default Public ReadOnly Property Item(x As Object) As RecordsetField
            Get
                Dim C As DataColumn = row.Table.Columns(x)
                Return New RecordsetField(row, x)
            End Get
        End Property

        Private ReadOnly Property ICollection_Count As Integer Implements ICollection.Count
            Get
                Throw New NotImplementedException()
            End Get
        End Property

        Private ReadOnly Property ICollection_IsSynchronized As Boolean Implements ICollection.IsSynchronized
            Get
                Throw New NotImplementedException()
            End Get
        End Property

        Private ReadOnly Property ICollection_SyncRoot As Object Implements ICollection.SyncRoot
            Get
                Throw New NotImplementedException()
            End Get
        End Property
    End Class

    Public Class RecordsetField
        Public Const adSmallInt As Integer = 2 ' Integer	SmallInt
        Public Const adInteger As Integer = 3 ' AutoNumber
        Public Const adSingle As Integer = 4 ' Single	Real
        Public Const adDouble As Integer = 5 ' Double	Float	Float
        Public Const adCurrency As Integer = 6 ' Currency	Money
        Public Const adDate As Integer = 7 ' Date	DateTime
        Public Const adIDispatch As Integer = 9 '
        Public Const adBoolean As Integer = 11 '	YesNo	Bit
        Public Const adVariant As Integer = 12 ' Sql_Variant(SQL Server 2000 +)	VarChar2
        Public Const adDecimal As Integer = 14 ' Decimal *
        Public Const adUnsignedTinyInt As Integer = 17 '	Byte	TinyInt
        Public Const adBigInt As Integer = 20 ' BigInt(SQL Server 2000 +)
        Public Const adGUID As Integer = 72 ' ReplicationID(Access 97 (OLEDB)), (Access 2000 (OLEDB))	UniqueIdentifier (SQL Server 7.0 +)
        Public Const adWChar As Integer = 130 ' NChar(SQL Server 7.0 +)
        Public Const adChar As Integer = 129 ' Char	Char
        Public Const adNumeric As Integer = 131 ' Decimal(Access 2000 (OLEDB))	Decimal
        Public Const adBinary As Integer = 128 ' Binary
        Public Const adDBTimeStamp As Integer = 135 ' DateTime(Access 97 (ODBC))	DateTime
        Public Const adVarChar As Integer = 200 ' Text(Access 97)	VarChar	VarChar
        Public Const adLongVarChar As Integer = 201 ' Memo(Access 97)
        Public Const adVarWChar As Integer = 202 ' Text(Access 2000 (OLEDB))	NVarChar (SQL Server 7.0 +)	NVarChar2
        Public Const adLongVarWChar As Integer = 203 ' Memo(Access 2000 (OLEDB))
        Public Const adVarBinary As Integer = 204 ' ReplicationID(Access 97)	VarBinary
        Public Const adLongVarBinary As Integer = 205 ' OLEObject	Image	Long Raw *

        Private Row As DataRow = Nothing
        Public Name As Object = ""
        Public Size As Integer = 0


        Public Sub New(Row As DataRow, Name As Object)
            Me.Row = Row
            Me.Name = Name
        End Sub

        Public Property Value As Object
            Get
                Return Row(Name)
            End Get
            Set(value As Object)
                Row(Name) = value
            End Set
        End Property


        Public ReadOnly Property Type As Object
            Get
                Return Row.Table.Columns(Name).DataType
            End Get
        End Property
    End Class

    Public Class PropIndexer(Of I, V)
        Public Delegate Sub setProperty(idx As I, value As V)
        Public Delegate Function getProperty(idx As I)

        Public getter As getProperty
        Public setter As setProperty

        Public Sub New(g As getProperty, s As setProperty)
            getter = g
            setter = s
        End Sub
        Public Sub New(g As getProperty)
            getter = g
            setter = AddressOf setPropertyNoop
        End Sub
        Public Sub New()
            getter = AddressOf getPropertyNoop
            setter = AddressOf setPropertyNoop
        End Sub

        Private Sub setPropertyNoop(idx As I, value As V)
            ' NOOP.  Intentionally left blank.
        End Sub
        Private Function getPropertyNoop(idx As I) As V
            Return CType(Nothing, V)
        End Function

        Default Public Property Item(ByVal nIndex As I) As V
            Get
                Return getter.Invoke(nIndex)

            End Get
            Set
                setter.Invoke(nIndex, Value)
            End Set
        End Property
    End Class

End Class
