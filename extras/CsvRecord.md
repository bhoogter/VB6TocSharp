# CSV Record Class

A convenience handler for processing CSV Records.

See Also:  [FieldInfoListSource.cs](FieldInfoListSource.cs)

## Example Usage

```c#
    public class PayComm : CsvRecord
    {
        [RecordField]
        public string Salesman = "";
        [RecordField]
        public string Lease = "";
        [RecordField]
        public string Name = "";
        [RecordField]
        public string Style = "";
        [RecordField]
        public string Landed = "";
        [RecordField]
        public string Sell = "";
        [RecordField]
        public string GM = "";
        [RecordField]
        public string SaleGM = "";
        [RecordField]
        public string Rate = "";
        [RecordField]
        public string Sales = "";
        [RecordField]
        public string Extra = "";
        [RecordField]
        public string Split = "";
        [RecordField]
        public string MargRec = "";

        public PayComm() { }
        public PayComm(
            string salesman = "", string lease = "", string name = "", string style = "", string landed = "",
            string sell = "", string gM = "", string saleGM = "", string rate = "", string sales = "",
            string extra = "", string split = "", string margRec = ""
            )
        {
            Salesman = salesman;
            Lease = lease;
            Name = name;
            Style = style;
            Landed = landed;
            Sell = sell;
            GM = gM;
            SaleGM = saleGM;
            Rate = rate;
            Sales = sales;
            Extra = extra;
            Split = split;
            MargRec = margRec;
        }
    }
```

## Field Config

The RecordField contains the following properties;

- public string name = "";
- public string type = "";
- public int max = 0;
- public int order = 0;

```c#
[RecordField(max = 4)]
public string RDP = "";
```

## Operations

### With Field Definitions

```
PayComm record = new PayComm()
record.Salesman = "field1";
record["Lease"] = "field2 \"with quotes\"";
record.Style = "Field4";
string CsvLine = record.ToString(); // ==> field1,"field2 ""withquotes""",,field4 // ...
PayComm recordCopy = new PayComm(record.ToString());

List<PayComm> fileContents = CsvRecord.FromCsvFile(csvFileContents); // static method call.  File IO is up to you.
String newCsvFileContents = ToCsvFile(fileContents); // Renders list of records as a string.  File IO is up to you.
```

### Without Field Definitions

```
CsvRecord record = new CsvRecord()
record[0] = "field1";
record[1] = "field2 \"with quotes\"";
record[3] = "Field4";
string CsvLine = record.ToString(); // ==> field1,"field2 ""withquotes""",,field4 // ...
```
