using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.Remoting.Messaging;
using System.Security.RightsManagement;
using System.Text;
using System.Threading.Tasks;
using System.Web.ModelBinding;
using System.Windows;
using Microsoft.VisualBasic;
using static modCSV;

namespace WinCDS.Classes
{
    public abstract class CsvRecord : FieldInfoListSource
    {
        protected List<string> extraFields = new List<string>();

        public CsvRecord() { }
        public CsvRecord(string line) { FromLine(line); }

        public string HeaderLine(bool Commented = false, bool addNL = false)
        {
            string S = "";
            if (Commented) S += "# ";
            foreach (var f in FieldInfoList()) S += ProtectCSV(f.Name) + ",";
            if (Strings.Right(S, 1) == ",") S = Strings.Left(S, S.Length - 1);
            if (addNL) S += "\n";
            return S;
        }

        protected string getFieldByIndex(int i)
        {
            FieldInfo f = thisField(i);
            if (f != null) return "" + f.GetValue(this);
            int extraIdx = i - FieldInfoListCount();
            if (extraIdx < extraFields.Count) return extraFields[extraIdx];
            return "";
        }

        protected void setFieldByIndex(int i, string value)
        {
            FieldInfo f = thisField(i);
            if (f != null)
                f.SetValue(this, value);
            else
            {
                int extraIdx = i - FieldInfoListCount();
                while (extraIdx >= extraFields.Count) extraFields.Add("");
                extraFields[extraIdx] = value;
            }
        }

        new public string this[string i]
        {
            get => "" + thisField(i).GetValue(this);
            set => thisField(i).SetValue(this, value);
        }

        new public string this[int i]
        {
            get => getFieldByIndex(i);
            set => setFieldByIndex(i, value);
        }


        public string ToLine()
        { return CSVLine(FieldInfoList().Select(f => f.GetValue(this).ToString()).Concat(extraFields).ToArray()); }

        public void FromLine(string line)
        {
            int i = 0;
            foreach (var f in FieldInfoList()) f.SetValue(this, CSVField(line, i++));
            extraFields = new List<string>();
            for (i = 0; i < CSVFieldCount(line) - FieldInfoListCount(); i++) extraFields.Add("");
        }

        public static List<T> FromCsvFile<T>(string csvContents) where T : CsvRecord, new()
        {
            List<T> res = new List<T>();
            foreach (var l in csvContents.Replace("\r", "").Split('\n'))
            {
                if (l == "") continue;
                if (Strings.Left(l, 1) == "#") continue;
                T item = new T();
                item.FromLine(l);
                res.Add(item);
            }
            return res;
        }

        public static string ToCsvFile<T>(List<T> lines, bool addHeader = false) where T : CsvRecord, new()
        {
            string res = "";
            if (lines.Count == 0) return res;
            if (addHeader) res += lines[0].HeaderLine() + "\r\n";

            foreach (var l in lines) res += l.ToLine() + "\r\n";
            return res;
        }
    }
}
