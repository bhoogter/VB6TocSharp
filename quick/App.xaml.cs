using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using VB2CS.Forms;

namespace VB2CS
{
  /// <summary>
  /// Interaction logic for App.xaml
  /// </summary>
  public partial class App : Application
    {
        [STAThread()]
        static void Main()
        {
            frm.instance.ShowDialog();
        }
    }
}
