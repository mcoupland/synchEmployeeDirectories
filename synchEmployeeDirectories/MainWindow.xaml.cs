using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Collections;
using System.Threading;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.Diagnostics;

namespace synchEmployeeDirectories
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public static List<Process> ExcelProcesses = new List<Process>();
        private List<Employee> employees = new List<Employee>();
        public ExcelProcessor ex = new ExcelProcessor();

    public readonly SynchronizationContext ctx = SynchronizationContext.Current;
        public MainWindow()
        {
            InitializeComponent();
            Loaded += MainWindow_Loaded;
        }
        private void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            ex.ProgressUpdated += Ex_ProgressUpdated;
            ex.GetUltiproFile();
            ex.ProcessingComplete += Ex_ProcessingComplete;
            //string word = "OGoggle";
            //char firstmost = word.ToLower().FirstMost();
        }

        private void Ex_ProcessingComplete(object sender, EventArgs e)
        {
            ex.SaveCeridianFormat();
        }

        private void Ex_ProgressUpdated(object sender, ProgressUpdatedArgs e)
        {
            UpdateUI(e.Message);
        }
        public void UpdateUI(string value)
        {
            ctx.Post(
                new SendOrPostCallback(
                o =>
                {
                    lbl.Content = value;
                }),
                value
            );
        }
    }

    public static class StringExtension
    {
        public static int CharCount(this string str)
        {
            return str.ToCharArray().Count();
        }


        public static char FirstMost(this string str)
        {
            Dictionary<char, int> dict = new Dictionary<char, int>();
            for (int i = 0; i < str.ToCharArray().Length; i++)
            {
                char key = str.ToCharArray()[i];
                if (dict.ContainsKey(key))
                {
                    int count = dict[key] + 1;
                    dict[key] = count;
                }
                else
                {
                    dict.Add(key, 1);
                }
            }
            int ii = str.CharCount();
            return dict.OrderByDescending(x => x.Value).First().Key;
        }
    }
}
