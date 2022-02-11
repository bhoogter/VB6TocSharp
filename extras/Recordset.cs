using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using static Functions;
using static modPath;
using static modStores;
using static modSupportForms;
using static VBExtension;

namespace WinCDS.Classes
{
    public class Recordset
    {
        public string Source = "";
        public Dictionary<dynamic, dynamic> Parameters = null;
        public string Database = "";
        public bool QuietErrors = false;

        private bool mAddingRow = false;
        public bool AddingRow => mAddingRow;

        OleDbConnection connection;
        OleDbDataAdapter adapter;
        DataTable table;
        DataTable filteredTable;
        string mFilter;

        public Recordset() { }

        public Recordset(DataTable table, OleDbDataAdapter adapter, OleDbConnection connection)
        {
            this.connection = connection;
            this.adapter = adapter;
            this.table = table;
        }


        public Recordset(string SQL, string File, bool QuietErrors = false, Dictionary<dynamic, dynamic> Parameters = null)
        {
            Source = SQL;
            this.Parameters = Parameters;
            Database = File;
            this.QuietErrors = QuietErrors;

            Open();
        }

        public void Close()
        {
            try { connection?.Close(); }
            catch { }

            connection = null;
            adapter = null;
            table = null;
            filteredTable = null;
        }

        public static void sqlExecutionError(string mSQL, Exception e)
        {
            string T = "";
            T += "getRecordSet Failed: " + e.Message + vbCrLf1;
            T += vbCrLf1;
            T += mSQL + vbCrLf1;
            T += vbCrLf1;
            T += "ERROR:" + e.Message;

            T = T.Replace("$EDESC", e.Message);
            //ErrMsg = Replace(ErrMsg, "$ENO", Err().Number);
            T = T.Replace("$ESRC", e.Source);
            MsgBox("Database Error: " + T, 0, "Error");
            //CheckStandardErrors(); // Bookmark/updateable query
        }

        private string ConnectionString(string file) { return "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" + file + PasswordProtectedDatabaseString + ";"; }

        public int AbsolutePosition { get; set; }
        public int Position { get => AbsolutePosition; set => AbsolutePosition = value; }
        public int RecordCount => table == null ? 0 : table.Rows == null ? 0 : table.Rows.Count;
        public bool EOF => AbsolutePosition >= RecordCount;
        public bool BOF => AbsolutePosition == 0;

        public bool FieldExists(string F) { return table?.Columns?.Contains(F) ?? false; }


        public int MoveFirst() { return AbsolutePosition = 0; }
        public int MoveNext() { return ++AbsolutePosition < RecordCount ? AbsolutePosition : AbsolutePosition = RecordCount; }
        public int MovePrevious() { return --AbsolutePosition >= 0 ? AbsolutePosition : AbsolutePosition = 0; }
        public int MoveLast() { return AbsolutePosition = RecordCount - 1; }

        public RecordsetFields Fields
        {
            get
            {
                if (AbsolutePosition >= 0 && AbsolutePosition < RecordCount) return new RecordsetFields(table.Rows[AbsolutePosition]);
                throw new ArgumentOutOfRangeException("Either EOF or BOF is true.");
            }
        }

        public List<string> FieldNames
        {
            get
            {
                if (table == null) return null;
                List<string> result = new List<string>();
                foreach (DataColumn item in table.Columns) result.Add(item.ColumnName);
                return result;
            }
        }

        public PropIndexer<dynamic, dynamic> Field
        {
            get => new PropIndexer<dynamic, dynamic>(
                (k) => Fields[k].Value,
                (k, v) => { Fields[k].Value = v; }
                );
        }

        public dynamic this[dynamic field]
        {
            get => GetField(field);
            set { SetField(field, value); }
        }

        public dynamic GetField(dynamic key) => Fields[key].Value;
        public void SetField(dynamic key, dynamic value) => Fields[key].Value = value;

        public List<List<dynamic>> GetRows()
        {
            var tableEnumerable = table.AsEnumerable();
            var tableList = tableEnumerable.ToArray().ToList();
            return tableList.ToList().Select((r) => r.ItemArray.ToList()).ToList();
        }

        public string Filter
        {
            get => mFilter;
            set
            {
                mFilter = value;
                if (string.IsNullOrEmpty(value))
                {
                    filteredTable = null;
                    return;
                }

                filteredTable = table.Select(mFilter).CopyToDataTable();
            }
        }

        internal bool Find(string v)
        {
            DataTable temp = table.Select(mFilter).CopyToDataTable();
            if (temp.Rows.Count == 0) return false;
            int x = table.Rows.IndexOf(temp.Rows[0]);
            AbsolutePosition = x;
            return true;
        }

        private void Open()
        {
            const int maxTries = 5;

            if (!FileExists(Database))
            {
                MsgBox("Database Not Found: " + Database);
                return;
            }

            DataSet result = new DataSet();
            connection = new OleDbConnection(ConnectionString(Database));
            OleDbCommand command = new OleDbCommand(Source, connection);
            foreach(var key in Parameters.Keys)
            {
                OleDbParameter param = command.CreateParameter();
                param.ParameterName = key;
                param.Value = Parameters[key];
            }
            adapter = new OleDbDataAdapter(command);
            try
            {
                connection.Open();
                adapter.FillSchema(result, SchemaType.Source);
                adapter.Fill(result, "Default");
            }
            catch (Exception e)
            {
                if (!QuietErrors) sqlExecutionError(Source, e);
            }
            finally { connection.Close(); }

            table = result.Tables["Default"];
        }

        public void Update()
        {
            OleDbCommandBuilder cb = new OleDbCommandBuilder(adapter);
            cb.QuotePrefix = "[";
            cb.QuoteSuffix = "]";
            try
            {
                connection.Open();
                adapter.UpdateCommand = cb.GetUpdateCommand();
                adapter.Update(table);
            }
            catch (Exception e)
            {
                if (!QuietErrors) sqlExecutionError(adapter.DeleteCommand.ToString(), e);
            }
            finally { connection.Close(); }

            mAddingRow = false;
        }

        public void AddNew()
        {
            DataRow newRow = table.NewRow();
            table.Rows.InsertAt(newRow, table.Rows.Count);
            AbsolutePosition = table.Rows.Count - 1;
            mAddingRow = true;
        }

        public void Delete()
        {
            OleDbCommandBuilder cb = new OleDbCommandBuilder(adapter);
            try
            {
                connection.Open();
                adapter.DeleteCommand = cb.GetDeleteCommand();
                adapter.Update(table);
            }
            catch (Exception e)
            {
                if (!QuietErrors) sqlExecutionError(adapter.UpdateCommand.ToString(), e);
            }
            finally { connection.Close(); }
        }

        public class RecordsetFields : ICollection
        {
            DataRow row = null;

            public RecordsetFields(DataRow row) { this.row = row; }

            public int Count => row.Table.Columns.Count;
            public object SyncRoot => null;
            public bool IsSynchronized => false;

            public void CopyTo(Array array, int index) { throw new InvalidOperationException("Not valid on object"); }

            public IEnumerator GetEnumerator() { return row.Table.Columns.GetEnumerator(); }

            public RecordsetField this[dynamic x]
            {
                get
                {
                    DataColumn C = row.Table.Columns[x];
                    return new RecordsetField(row, x);
                }
            }
        }

        public class RecordsetField
        {
            public const int adSmallInt = 2; //	Integer	SmallInt
            public const int adInteger = 3; //	AutoNumber
            public const int adSingle = 4; //	Single	Real
            public const int adDouble = 5; //	Double	Float	Float
            public const int adCurrency = 6; //	Currency	Money
            public const int adDate = 7; //	Date	DateTime
            public const int adIDispatch = 9; //
            public const int adBoolean = 11; //	YesNo	Bit
            public const int adVariant = 12; //	 	Sql_Variant (SQL Server 2000 +)	VarChar2
            public const int adDecimal = 14; //	 	 	Decimal *
            public const int adUnsignedTinyInt = 17; //	Byte	TinyInt
            public const int adBigInt = 20; //	 	BigInt (SQL Server 2000 +)
            public const int adGUID = 72; //	ReplicationID (Access 97 (OLEDB)), (Access 2000 (OLEDB))	UniqueIdentifier (SQL Server 7.0 +)
            public const int adWChar = 130; //	 	NChar (SQL Server 7.0 +)
            public const int adChar = 129; //	 	Char	Char
            public const int adNumeric = 131; //	Decimal (Access 2000 (OLEDB))	Decimal
            public const int adBinary = 128; //	 	Binary
            public const int adDBTimeStamp = 135; //	DateTime (Access 97 (ODBC))	DateTime
            public const int adVarChar = 200; //	Text (Access 97)	VarChar	VarChar
            public const int adLongVarChar = 201; //	Memo (Access 97)
            public const int adVarWChar = 202; //	Text (Access 2000 (OLEDB))	NVarChar (SQL Server 7.0 +)	NVarChar2
            public const int adLongVarWChar = 203; //	Memo (Access 2000 (OLEDB))
            public const int adVarBinary = 204; //	ReplicationID (Access 97)	VarBinary
            public const int adLongVarBinary = 205; //	OLEObject	Image	Long Raw *

            private DataRow Row = null;
            public dynamic Name = "";
            public int Size = 0;


            public RecordsetField(DataRow Row, dynamic Name)
            {
                this.Row = Row;
                this.Name = Name;
            }

            public dynamic Value
            {
                get => Row[Name];
                set => Row[Name] = value;
            }

            public Type Type => Row.Table.Columns[Name].DataType;

            // private string TypeName
            // {
            //     get
            //     {
            //         switch (Type)
            //         {
            //             case adBinary: return "adBinary(" + Size + ")";
            //             case adBoolean: return "adBoolean";
            //             case adChar: return "adChar(" + Size + ")";
            //             case adCurrency: return "adCurrency";
            //             case adDBTimeStamp: return "adDBTimeStamp";
            //             case adDouble: return "adDouble";
            //             case adInteger: return "adInteger";
            //             case adLongVarBinary: return "adLongVarBinary";
            //             case adLongVarChar: return "adLongVarChar";
            //             case adLongVarWChar: return "adLongVarWChar";
            //             case adNumeric: return "adNumeric";
            //             case adSingle: return "adSingle";
            //             case adSmallInt: return "adSmallInt";
            //             //case adTinyInt: return  "adTinyInt";
            //             case adUnsignedTinyInt: return "adUnsignedTinyInt";
            //             case adVarBinary: return "adVarBinary (" + Size + ")";
            //             case adVarChar: return "adVarChar (" + Size + ")";
            //             case adVarWChar: return "adVarWChar(" + Size + ")";
            //             case adWChar: return "adWChar(" + Size + ")";
            //             default: return "UnKnown Field Type: " + Name + ", " + Type;
            //         }
            //     }
            // }
        }
    }
}
