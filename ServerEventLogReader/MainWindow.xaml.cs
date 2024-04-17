/* Title:           Server Event Log Reader
 * Date:            10-22-20
 * Author:          Terry Holmes
 * 
 * Description:     This is used to read the event log */

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
using NewEventLogDLL;
using NewEmployeeDLL;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Win32;
using DateSearchDLL;
using CSVFileDLL;

namespace ServerEventLogReader
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        //setting up the classes
        WPFMessagesClass TheMessagesClass = new WPFMessagesClass();
        EventLogClass TheEventLogClass = new EventLogClass();
        EmployeeClass TheEmployeeClass = new EmployeeClass();
        EventLogImportDataSet TheEventLogImportDataSet = new EventLogImportDataSet();
        DateSearchClass TheDateSearchClass = new DateSearchClass();
        
        //setting up the data
        ComboEmployeeDataSet TheComboEmployeeDataSet = new ComboEmployeeDataSet();
        FindServerEventLogContentMatchDataSet TheFindServerEventLogContentMatchDataSet = new FindServerEventLogContentMatchDataSet();

        public MainWindow()
        {
            InitializeComponent();
        }

        private void expCloseProgram_Expanded(object sender, RoutedEventArgs e)
        {
            expCloseProgram.IsExpanded = false;
            TheMessagesClass.CloseTheProgram();
        }

        private void expImportExcel_Expanded(object sender, RoutedEventArgs e)
        {
            Excel.Application xlDropOrder;
            Excel.Workbook xlDropBook;
            Excel.Worksheet xlDropSheet;
            Excel.Range range;

            int intColumnRange = 0;
            int intCounter;
            int intNumberOfRecords;
            string strTransactionDate;
            string strTaskCategory;
            string strNotes;
            DateTime datTransactionDate = DateTime.Now;
            double douTransactionDate;
            int intRemainder;
            int intRecordsReturned;
            

            try
            {
                TheEventLogImportDataSet.importedevents.Rows.Clear();
                expImportExcel.IsExpanded = false;

                Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
                dlg.FileName = "Document"; // Default file name
                dlg.DefaultExt = ".xlsx"; // Default file extension
                dlg.Filter = "Excel (.xlsx)|*.xlsx"; // Filter files by extension

                // Show open file dialog box
                Nullable<bool> result = dlg.ShowDialog();

                // Process open file dialog box results
                if (result == true)
                {
                    // Open document
                    string filename = dlg.FileName;
                }

                PleaseWait PleaseWait = new PleaseWait();
                PleaseWait.Show();

                xlDropOrder = new Excel.Application();
                xlDropBook = xlDropOrder.Workbooks.Open(dlg.FileName, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                xlDropSheet = (Excel.Worksheet)xlDropOrder.Worksheets.get_Item(1);

                range = xlDropSheet.UsedRange;
                intNumberOfRecords = range.Rows.Count;
                intColumnRange = range.Columns.Count;

                for (intCounter = 2; intCounter < intNumberOfRecords; intCounter++)
                {
                    strTaskCategory = Convert.ToString((range.Cells[intCounter, 5] as Excel.Range).Value2).ToUpper();

                    intRemainder = intCounter % 100000;

                    if (intRemainder == 0)
                    {
                        TheMessagesClass.InformationMessage(Convert.ToString(intCounter));
                    }                    
                    
                    strTransactionDate = Convert.ToString((range.Cells[intCounter, 2] as Excel.Range).Value2).ToUpper();
                    douTransactionDate = Convert.ToDouble(strTransactionDate);
                    datTransactionDate = DateTime.FromOADate(douTransactionDate);
                    strNotes = Convert.ToString((range.Cells[intCounter, 6] as Excel.Range).Value2).ToUpper();
                    
                        if (((strNotes.Contains("READDATA") == true) || (strNotes.Contains("WRITEDATA") == true) || (strNotes.Contains("BJC-FILE$") == false)))
                        {
                            EventLogImportDataSet.importedeventsRow NewEventRow = TheEventLogImportDataSet.importedevents.NewimportedeventsRow();

                            NewEventRow.Category = strTaskCategory;
                            NewEventRow.ImportDate = DateTime.Now;
                            NewEventRow.TaskNotes = strNotes;
                            NewEventRow.TransactionDate = datTransactionDate;

                            TheEventLogImportDataSet.importedevents.Rows.Add(NewEventRow);
                        }
                    
                    
                }

                dgrEvents.ItemsSource = TheEventLogImportDataSet.importedevents;

                PleaseWait.Close();


            }
            catch (Exception Ex)
            {
                TheEventLogClass.InsertEventLogEntry(DateTime.Now, "Server Event Log Reader // Main Window // Import Excel  " + Ex.Message);

                TheMessagesClass.ErrorMessage(Ex.ToString());
            }
        }

        private void expProcessImport_Expanded(object sender, RoutedEventArgs e)
        {
            int intCounter;
            int intNumberOfRecords;
            string strEventCategory;
            string strEventNotes;
            DateTime datTransactionDate;
            bool blnFatalError = false;
            int intRemainder;
            int intRecordsReturned;

            PleaseWait PleaseWait = new PleaseWait();
            PleaseWait.Show();

            try
            {
                intNumberOfRecords = TheEventLogImportDataSet.importedevents.Rows.Count;

                if(intNumberOfRecords > 1)
                {
                    for(intCounter = 0; intCounter < intNumberOfRecords; intCounter++)
                    {
                        strEventCategory = TheEventLogImportDataSet.importedevents[intCounter].Category;
                        strEventNotes = TheEventLogImportDataSet.importedevents[intCounter].TaskNotes;
                        datTransactionDate = TheEventLogImportDataSet.importedevents[intCounter].TransactionDate;

                        TheFindServerEventLogContentMatchDataSet = TheEventLogClass.FindServerEventLogContentMatch(strEventNotes, strEventCategory, datTransactionDate);

                       intRecordsReturned = TheFindServerEventLogContentMatchDataSet.FindServerEventLogByContentMatch.Rows.Count;

                        if(intRecordsReturned < 1)
                        {
                            intRemainder = intCounter % 5000;

                            if (intRemainder == 0)
                            {
                                TheMessagesClass.InformationMessage(Convert.ToString(intCounter));
                            }

                            blnFatalError = TheEventLogClass.InsertServerEventLog(datTransactionDate, strEventCategory, strEventNotes);

                            if (blnFatalError == true)
                                throw new Exception();
                        }
                    }
                }

                TheMessagesClass.InformationMessage("All Records Have Been Imported ");
                
            }
            catch (Exception Ex)
            {
                TheEventLogClass.InsertEventLogEntry(DateTime.Now, "Server Event Log Reader // Main Window // Process Import " + Ex.Message);

                TheMessagesClass.ErrorMessage(Ex.ToString());
            }

            PleaseWait.Close();
        }

        private void expImportCSV_Expanded(object sender, RoutedEventArgs e)
        {
            string strFileName = "";

            try
            {
                TheEventLogImportDataSet.importedevents.Rows.Clear();
                expImportExcel.IsExpanded = false;

                Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
                dlg.FileName = "Document"; // Default file name

                // Show open file dialog box
                Nullable<bool> result = dlg.ShowDialog();

                // Process open file dialog box results
                if (result == true)
                {
                    // Open document
                    strFileName = dlg.FileName;
                }

                string[] lines = System.IO.File.ReadAllLines(strFileName);

                foreach (string line in lines)
                {
                    string[] columns = line.Split(',');
                    foreach (string column in columns)
                    {
                        // Do something
                    }
                }
            }
            catch(Exception Ex)
            {
                TheEventLogClass.InsertEventLogEntry(DateTime.Now, "Server Event Log Reader // Main Window // Import CSV Expander " + Ex.Message);

                TheMessagesClass.ErrorMessage(Ex.ToString());
            }
            
           
        }
    }
}
