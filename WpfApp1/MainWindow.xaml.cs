using CsvFileReader;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using MSExcel = Microsoft.Office.Interop.Excel;

namespace WpfApp1
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window, INotifyPropertyChanged
    {
        Data data;
        Command addExcel;

        public ObservableCollection<string> IsDefects { get; set; }
        public ObservableCollection<Data> Datas { get; set; }

        public Data Data
        {
            get
            {
                return data;
            }
            set
            {
                data = value;
                OnPropertyChanged("Data");
            }
        }

        public MainWindow()
        {
            IsDefects = new ObservableCollection<string> { "no", "yes" };
            Datas = new ObservableCollection<Data>();

            InitializeComponent();
        }

        public Command AddExcel
        {
            get
            {
                return addExcel ??
                  (addExcel = new Command(obj =>
                  {
                      System.Windows.Forms.OpenFileDialog openFileDialog;
                      System.Windows.Forms.DialogResult Form;

                      openFileDialog = new System.Windows.Forms.OpenFileDialog();
                      openFileDialog.Filter = "Excel files|*.xlsx;*.xls;*.csv";
                      Form = openFileDialog.ShowDialog();
                      openFileDialog.RestoreDirectory = true;

                      if (Form == System.Windows.Forms.DialogResult.OK)
                      {
                          string filename = openFileDialog.FileName;
                          var fileInfo = new FileInfo(filename);
                          Datas.Clear();
                          if (fileInfo.Extension == ".csv")
                          {
                              CsvFile csv = new CsvFile(filename);
                              for (int i = 1; i < csv.Reader.DataRows.Count; ++i)
                              {
                                  var array = csv.Reader.DataRows[i].Split(';');

                                  if (array.Length <= 5)
                                  {
                                      System.Windows.Forms.MessageBox.Show($"В файле {fileInfo.Name} представлены не все параметры, импорт не возможен", "Ошибка", (MessageBoxButtons)MessageBoxButton.OK, MessageBoxIcon.Error);
                                      return;
                                  }
                                  Datas.Add(new Data(array[0], array[1], array[2], array[3], array[4], array[5]));
                              }
                          }
                          else
                          {
                              var Excel = new MSExcel.Application();
                              var Book = Excel.Workbooks.Open(filename);
                              var Sheet = (MSExcel.Worksheet)Book.Sheets[1];

                              int lastUsedRow = Excel.Cells.Find("*", System.Reflection.Missing.Value,
                                               System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                                               MSExcel.XlSearchOrder.xlByRows, MSExcel.XlSearchDirection.xlPrevious,
                                               false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;

                              if (lastUsedRow == 0)
                                  return;

                              for (int i = 1; i < lastUsedRow; ++i)
                              {
                                  Datas.Add(new Data(
                                      Sheet.Cells[i + 1, 1].FormulaLocal,
                                      Sheet.Cells[i + 1, 2].FormulaLocal,
                                      Sheet.Cells[i + 1, 3].FormulaLocal,
                                      Sheet.Cells[i + 1, 4].FormulaLocal,
                                      Sheet.Cells[i + 1, 5].FormulaLocal,
                                      Sheet.Cells[i + 1, 6].FormulaLocal
                                      ));

                              }

                              Book.Close(false);
                              Excel.Quit();
                              Excel = null;
                              Book = null;
                              Sheet = null;
                              GC.Collect();
                          }
                      }

                  }));
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        public bool OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
                return true;
            }
            return false;
        }

        private void ProgressBar_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {

        }
    }
}
