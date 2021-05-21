using System.Windows.Controls;


public partial class ctlTestCase : UserControl
{
    // Option Explicit
    private static bool mValue = false;
    private const bool d_Value = true;


    public static bool Value
    {
        get
        {
            bool Value;
            Value = mValue;

            return Value;
        }
        set
        {
            mValue = value;

        }
    }


    private static void UserControl_InitProperties()
    {
        mValue = d_Value;
    }

    private static void UserControl_ReadProperties(ref PropertyBag PropBag)
    {
        // TODO (not supported): On Error Resume Next
        Value = PropBag.ReadProperty("Value", d_Value);
    }

    private static void UserControl_WriteProperties(ref PropertyBag PropBag_UNUSED)
    {
        // TODO (not supported): On Error Resume Next
        PropBag.WriteProperty("Value", Value, d_Value);
    }
}
