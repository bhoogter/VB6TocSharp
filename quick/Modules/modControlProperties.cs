static class modControlProperties
{
    // control default properties
    public static string ConvertControlProperty(string Src, string vProp, string cType)
    {
        string _ConvertControlProperty = "";
        // If IsInStr(vProp, __S1) Then Stop
        _ConvertControlProperty = vProp;
        switch (vProp)
        {
            case "ListIndex":
                _ConvertControlProperty = "SelectedIndex";
                break;
            case "Visible":
                _ConvertControlProperty = "Visibility";
                break;
            case "Enabled":
                _ConvertControlProperty = "IsEnabled";
                break;
            case "TabStop":
                _ConvertControlProperty = "IsTabStop";
                break;
            case "SelStart":
                _ConvertControlProperty = "SelectionStart";
                break;
            case "SelLength":
                _ConvertControlProperty = "SelectionLength";
                break;
            case "Caption":
                if (cType == "VB.Label") _ConvertControlProperty = "Content";
                break;
            case "Value":
                if (cType == "VB.CheckBox") _ConvertControlProperty = "IsChecked";
                if (cType == "VB.OptionButton") _ConvertControlProperty = "IsChecked";
                if (cType == "MSComCtl2.DTPicker") _ConvertControlProperty = "DisplayDate";
                break;
            case "Text":
                if (cType == "VB.ListBox") _ConvertControlProperty = "SelectedText.toString()";
                break;
            case "ListCount":
                if (cType == "VB.ListBox") _ConvertControlProperty = "Items.Count";
                break;
            case "Default":
                _ConvertControlProperty = "IsDefault";
                break;
            case "Cancel":
                _ConvertControlProperty = "IsCancel";
                break;
            case "LBound":
                _ConvertControlProperty = "LBound()";
                break;
            case "UBound":
                _ConvertControlProperty = "UBound()";
                break;
            case "":
                switch (cType)
                {
                    case "VB.Caption":
                        _ConvertControlProperty = "Content";
                        break;
                    case "VB.TextBox":
                        _ConvertControlProperty = "Text";
                        break;
                    case "VB.ComboBox":
                        _ConvertControlProperty = "Text";
                        break;
                    case "VB.PictureBox":
                        _ConvertControlProperty = "Source";
                        break;
                    case "VB.Image":
                        _ConvertControlProperty = "Source";
                        break;
                    case "VB.OptionButton":
                        _ConvertControlProperty = "IsChecked";
                        break;
                    case "VB.CheckBox":
                        _ConvertControlProperty = "IsChecked";
                        break;
                    case "VB.Frame":
                        _ConvertControlProperty = "Content";
                        break;
                    case "VB.Label":
                        _ConvertControlProperty = "Content";
                        break;
                    default:
                        _ConvertControlProperty = "DefaultProperty";
                        break;
                }
                break;
        }
        return _ConvertControlProperty;
    }

}
