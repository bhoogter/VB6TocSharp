
namespace WinCDS.Classes
{
    public abstract class FixedWidthRecord : FieldInfoListSource
    {
        public string RecordStart => "";
        public string RecordTerminator => "";

        public new string ToString()
        {
            string s = RecordStart;
            foreach (var f in FieldInfoList())
            {
                RecordField r = thisFieldMod(f.Name);
                int w = r.max;
                s += (f.GetValue(this).ToString() + new string(' ', w)).Substring(0, w);
            }
            s += RecordTerminator;

            return s;
        }

        public void fromString(string l)
        {
            foreach (var f in FieldInfoList())
            {
                RecordField r = thisFieldMod(f.Name);
                int w = r.max;
                string v = l.Substring(0, w);
                l = l.Substring(w);
                f.SetValue(this, l);
            }
        }
    }
}
