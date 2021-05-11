static class modControlProperties
{
    // Option Explicit


    public static string ConvertControlProperty(string Src_UNUSED, string vProp, string cType)
    {
        string ConvertControlProperty = "";
        //If IsInStr(vProp, "SetF") Then Stop
        ConvertControlProperty = vProp;
        switch (vProp)
        {
            case "ListIndex":
                ConvertControlProperty = "SelectedIndex";
                break;
            case "Visible":
                ConvertControlProperty = "Visibility";
                break;
            case "Enabled":
                ConvertControlProperty = "IsEnabled";
                break;
            case "TabStop":
                ConvertControlProperty = "IsTabStop";
                break;
            case "SelStart":
                ConvertControlProperty = "SelectionStart";
                break;
            case "SelLength":
                ConvertControlProperty = "SelectionLength";
                break;
            case "Caption":
                if (cType == "VB.Label")
                {
                    ConvertControlProperty = "Content";
                }
                break;
            case "Value":
                if (cType == "VB.CheckBox")
                {
                    ConvertControlProperty = "IsChecked";
                }
                if (cType == "VB.OptionButton")
                {
                    ConvertControlProperty = "IsChecked";
                }
                if (cType == "MSComCtl2.DTPicker")
                {
                    ConvertControlProperty = "DisplayDate";
                }
                break;
            case "Text":
                if (cType == "VB.ListBox")
                {
                    ConvertControlProperty = "SelectedText.toString()";
                }
                break;
            case "ListCount":
                if (cType == "VB.ListBox")
                {
                    ConvertControlProperty = "Items.Count";
                }
                break;
            case "Default":
                ConvertControlProperty = "IsDefault";
                break;
            case "Cancel":
                ConvertControlProperty = "IsCancel";

                break;
            case "":
                switch (cType)
                {
                    case "VB.Caption":
                        ConvertControlProperty = "Content";
                        break;
                    case "VB.TextBox":
                        ConvertControlProperty = "Text";
                        break;
                    case "VB.ComboBox":
                        ConvertControlProperty = "Text";
                        break;
                    case "VB.PictureBox":
                        ConvertControlProperty = "Source";
                        break;
                    case "VB.Image":
                        ConvertControlProperty = "Source";
                        break;
                    case "VB.ComboBox":
                        ConvertControlProperty = "Text";
                        break;
                    case "VB.OptionButton":
                        ConvertControlProperty = "IsChecked";
                        break;
                    case "VB.CheckBox":
                        ConvertControlProperty = "IsChecked";
                        break;
                    case "VB.Frame":
                        ConvertControlProperty = "Content";
                        break;
                    default:
                        ConvertControlProperty = "DefaultProperty";
                        break;
                }
                break;
        }
        return ConvertControlProperty;
    }
}
