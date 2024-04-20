/* Title:           Update Work Task Production Value
 * Date:            3-7-19
 * Author:          Terry Holmes
 * 
 * Description:     This used to import values */

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
using WorkTaskDLL;
using Microsoft.Win32;
using Excel = Microsoft.Office.Interop.Excel;
using DataValidationDLL;

namespace UpdateWorkTaskProductionValue
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        //setting up the classes
        WPFMessagesClass TheMessagesClass = new WPFMessagesClass();
        EventLogClass TheEventLogClass = new EventLogClass();
        WorkTaskClass TheWorkTaskClass = new WorkTaskClass();
        DataValidationClass TheDataValidationClass = new DataValidationClass();
        

        //setting up the data
        FindWorkTaskByTaskKeywordDataSet TheFindWorkTaksByTaskKeywordDataSet = new FindWorkTaskByTaskKeywordDataSet();
        ImportedWorkTaskDataSet TheImportedWorkTaskDataSet = new ImportedWorkTaskDataSet();

        WorkTaskDataSet aWorkTaskDataSet;
        WorkTaskDataSetTableAdapters.worktaskTableAdapter aWorkTaskTableAdatper;
        WorkTaskDataSet TheWorkTaskDataSet;

        public MainWindow()
        {
            InitializeComponent();
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            TheMessagesClass.CloseTheProgram();
        }

        private void BtnHelp_Click(object sender, RoutedEventArgs e)
        {
            TheMessagesClass.LaunchHelpSite();
        }

        private void BtnImportExcel_Click(object sender, RoutedEventArgs e)
        {
            Excel.Application xlDropOrder;
            Excel.Workbook xlDropBook;
            Excel.Worksheet xlDropSheet;
            Excel.Range range;

            string strWorkTask;
            int intColumnRange = 0;
            int intCounter;
            int intNumberOfRecords;
            int intRecordsReturned;
            int intWorkTaskID;
            decimal decProductivityRate = 0;
            string strProductivityRate;
            int intRecordCount = 1;
            bool blnError;
            string strFullTask;

            try
            {
                TheImportedWorkTaskDataSet.importedworktask.Rows.Clear();

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

                for (intCounter = 2; intCounter <= intNumberOfRecords; intCounter++)
                {
                    strWorkTask = Convert.ToString((range.Cells[intCounter, 1] as Excel.Range).Value2);

                    if(strWorkTask == null)
                    {
                        
                    }
                    else
                    {
                        strFullTask = strWorkTask;
                        strWorkTask = strWorkTask.Substring(0, 5);

                        TheFindWorkTaksByTaskKeywordDataSet = TheWorkTaskClass.FindWorkTaskByTaskKeyword(strWorkTask);

                        intRecordsReturned = TheFindWorkTaksByTaskKeywordDataSet.FindWorkTaskByTaskKeyword.Rows.Count;

                        if (intRecordsReturned > 0)
                        {
                            strProductivityRate = Convert.ToString((range.Cells[intCounter, 2] as Excel.Range).Value2);

                            blnError = TheDataValidationClass.VerifyDoubleData(strProductivityRate);
                            if(blnError == false)
                            {
                                decProductivityRate = Convert.ToDecimal(strProductivityRate);


                                intWorkTaskID = TheFindWorkTaksByTaskKeywordDataSet.FindWorkTaskByTaskKeyword[0].WorkTaskID;

                                ImportedWorkTaskDataSet.importedworktaskRow NewRateRow = TheImportedWorkTaskDataSet.importedworktask.NewimportedworktaskRow();

                                NewRateRow.WorkTaskID = intWorkTaskID;
                                NewRateRow.WorkTask = strFullTask;
                                NewRateRow.ProductionRate = decProductivityRate;
                                NewRateRow.RecordCount = intRecordCount;
                                intRecordCount++;

                                TheImportedWorkTaskDataSet.importedworktask.Rows.Add(NewRateRow);
                            }
                            
                        }
                    }
                    
                }

                PleaseWait.Close();
                dgrResults.ItemsSource = TheImportedWorkTaskDataSet.importedworktask;
            }
            catch (Exception Ex)
            {
                TheEventLogClass.InsertEventLogEntry(DateTime.Now, "Update Work Task Production Value // Import Excel Button " + Ex.Message);

                TheMessagesClass.ErrorMessage(Ex.ToString());
            }
        }

        private void BtnProcess_Click(object sender, RoutedEventArgs e)
        {
            int intCounter;
            int intNumberOfRecords;
            int intWorkTaskID;
            decimal decProductivityRate;
            bool blnFatalError = false;

            try
            {
                intNumberOfRecords = TheImportedWorkTaskDataSet.importedworktask.Rows.Count - 1;

                for(intCounter = 0; intCounter <= intNumberOfRecords; intCounter++)
                {
                    intWorkTaskID = TheImportedWorkTaskDataSet.importedworktask[intCounter].WorkTaskID;

                    
                        decProductivityRate = TheImportedWorkTaskDataSet.importedworktask[intCounter].ProductionRate;

                        decProductivityRate = Math.Round(decProductivityRate, 2);
                    
                    

                    blnFatalError = TheWorkTaskClass.UpdateWorkTaskProductivityRate(intWorkTaskID, decProductivityRate);

                    if (blnFatalError == true)
                        throw new Exception();
                }
            }
            catch (Exception Ex)
            {
                TheEventLogClass.InsertEventLogEntry(DateTime.Now, "Update Work Task Production Value // Process Button " + Ex.Message);

                TheMessagesClass.ErrorMessage(Ex.ToString());
            }
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            TheWorkTaskDataSet = GetWorkTaskInfo();

            dgrResults.ItemsSource = TheWorkTaskDataSet.worktask;
            
        }
        private WorkTaskDataSet GetWorkTaskInfo()
        {
            aWorkTaskDataSet = new WorkTaskDataSet();
            aWorkTaskTableAdatper = new WorkTaskDataSetTableAdapters.worktaskTableAdapter();
            aWorkTaskTableAdatper.Fill(aWorkTaskDataSet.worktask);

            return aWorkTaskDataSet;
        }
    }
}
