# FixedWidthRecord

A Fixed width record utility class using reflection for defined fields.

See Also:  

 - [FieldInfoListSource.cs](FieldInfoListSource.cs)
 - [CsvRecord.md](CsvRecord.md)

## Example

```c#
public class CreditReportHeader : FixedWidthRecord
    {
        [RecordField(max = 4)] public string RecordDescriptorWord = "";
        [RecordField(max = 6)] public string RecordIdentifier = "";
        [RecordField(max = 2)] public string CycleNumber = "";
        [RecordField(max = 10)] public string Reserved1 = "";
        [RecordField(max = 10)] public string EquifaxProgramID = "";
        [RecordField(max = 5)] public string ExperianProgramID = "";
        [RecordField(max = 10)] public string TransUnionProgramID = "";
        [RecordField(max = 8)] public string ActivityDate = "";
        [RecordField(max = 8)] public string DateCreated = "";
        [RecordField(max = 8)] public string ProgramDate = "";
        [RecordField(max = 8)] public string ProgramRevisionDate = "";
        [RecordField(max = 40)] public string ReporterName = "";
        [RecordField(max = 96)] public string ReporterAddress = "";
        [RecordField(max = 10)] public string ReporterPhone = "";
        [RecordField(max = 40)] public string SoftwareVendorName = "";
        [RecordField(max = 5)] public string SoftwareVersion = "";
        [RecordField(max = 156)] public string Reserved2 = "";
    }
```

## Usage

```c#
            CreditReportHeader Hdr = new CreditReportHeader();

            Hdr.RecordDescriptorWord = "0426"; //1-4 Length of header record, in bytes.
            Hdr.RecordIdentifier = "HEADER"; //5-10 Constant value.
            Hdr.CycleNumber = " "; //11-12 Reporting cycle number.  We're not using this.
            Hdr.Reserved1 = Space(10); //13-22 This might be used for Innovis, but not now.
            Hdr.EquifaxProgramID = QueryEquifaxCreditID(); //23-32 This company's Equifax ID.
            Hdr.ExperianProgramID = QueryExperianCreditID(); //33-37 This company's Experian ID.
            Hdr.TransUnionProgramID = QueryTransUnionCreditID(); //38-47 This company's TransUnion ID.
            Hdr.ActivityDate = ExportDate(DateTime.Today); //48-55 Date of last account activity on any account being reported.
            Hdr.DateCreated = ExportDate(DateTime.Today); // Date this export file was created - today.
            Hdr.ProgramDate = "20031108"; // Date the export format was developed.
            Hdr.ProgramRevisionDate = "20031111"; // Date the export format was last revised.
            Hdr.ReporterName = StoreSettings().Name; // Reporting company - the store.
            Hdr.ReporterAddress = StoreSettings().Address; // Reporting company's address.
            Hdr.ReporterPhone = CleanAni(StoreSettings().Phone, 0); // Reporting company's phone number.
            Hdr.SoftwareVendorName = AdminContactCompany; // Software vendor's name - Us.
            Hdr.SoftwareVersion = SoftwareVersion(true, false); // WinCDS Version.
            Hdr.Reserved2 = Space(156); // Blank.
                                        //  Hdr.BlankFill = Space(938)
                                        
            string CreditLine = Hdr.ToString();
```
