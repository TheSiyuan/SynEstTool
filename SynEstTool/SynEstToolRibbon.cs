using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Text;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Excel = Microsoft.Office.Interop.Excel;
using System.Configuration;
using System.Collections.Specialized;
using System.Timers;
using Microsoft.CSharp.RuntimeBinder;

namespace SynEstTool
{
    public partial class SynEstToolRibbon //functions for button clicks
    {
        //public variables under this class declared here, only here
        //use xml file for config, no longer read text files
        //!!!generic name such as counter or i, can only be used as local variables.!!!

        private void SynEstToolRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            //Excel.Application xlapp = new Microsoft.Office.Interop.Excel.Application();
            //xlapp.DisplayAlerts = false;
        }

        //not used.

        //public void Ribbon_Activation()
        //{
        //    group2.Visible = false;
        //    try
        //    {
        //        Excel.Workbook activeworkbook = Globals.ThisAddIn.Application.ActiveWorkbook;
        //        if (activeworkbook != null)
        //        {
        //            string sattr;
        //            sattr = ConfigurationManager.AppSettings.Get("est_list_sourcesheet");
        //            foreach (Microsoft.Office.Interop.Excel.Worksheet worksheet in activeworkbook.Worksheets)
        //            {
        //                if (worksheet.Name.Contains(sattr))
        //                {
        //                    group2.Visible = true;
        //                    break;
        //                }
        //            }
        //        }
        //    }
        //    catch (NullReferenceException error)
        //    {
        //        Debug.WriteLine("excel not starting");
        //    }
        //}

        
        public string[] strarr1 = new string[] { "1", "2", "3", "1", "2", "3", "1", "2", "3", "1", "2", "3", "1", "2", "3", "1", "2", "3", "1", "2", "3", "1", "2", "3", "1", "2", "3" };
        //being used for testing
        private void BtnStart_Click(object sender, RibbonControlEventArgs e)
        {
            //Act_Est_Functions n = new Act_Est_Functions();
            //Active_Est_List newList = new Active_Est_List(strarr1, n.Est_List_Col_Array());

            Excel.Workbook activeworkbook = Globals.ThisAddIn.Application.ActiveWorkbook;
            //int count = activeworkbook.Worksheets.Count;

            //Excel.Worksheet newsheet = activeworkbook.Worksheets.Add(Type.Missing, activeworkbook.Worksheets[count],Type.Missing,XlSheetType.xlWorksheet);
            //newsheet.Name = "Print List";
            //Excel.Application xlapp = null;
            //xlapp.DisplayAlerts = false;
            Worksheet activeworksheet = activeworkbook.ActiveSheet;
            activeworksheet.PageSetup.Zoom = false;
            activeworksheet.PageSetup.FitToPagesWide = 1;

        }

        private void Consolidate_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Workbook ActiveWorkBook = Globals.ThisAddIn.Application.ActiveWorkbook;
            ActiveWorkBook.Application.DisplayAlerts = false;
            String[] ListofWorksheets = new String[ActiveWorkBook.Worksheets.Count];
            int i = 0;
            foreach (Microsoft.Office.Interop.Excel.Worksheet wSheet in ActiveWorkBook.Worksheets)
            {
                ListofWorksheets[i] = wSheet.Name;
                i++;
            }
            SheetSelector Consolidate_sheet = new SheetSelector(Alist: ListofWorksheets);
            Consolidate_sheet.ShowDialog();
        }

        private void Btn_PrintEstList_Click(object sender, RibbonControlEventArgs e)
        {
            //1. generate a new worksheet, named "PrintList", if not a estimate list, then do nothing.
            //2. reading source sheet line by line
            //3. checking if criteria is met
            //4. if met then make set value in class
            //5. paste class onto the new sheet
            //6. format the data that was pasted, set up page break if necessary
            //7. call out print form with correct settings.
            Stopwatch stopWatch = new Stopwatch();
            stopWatch.Start();
            //TimeSpan ts = stopWatch.Elapsed;
            //Debug.WriteLine(ts);
            Excel.Workbook activeworkbook = Globals.ThisAddIn.Application.ActiveWorkbook;
            //mute alert
            activeworkbook.Application.DisplayAlerts = false;

            Excel.Worksheet worksheet_source, worksheet_target;
            worksheet_source = null;
            worksheet_target = null;
            //check if the workbook is the estimate work book
            string sAttr;
            bool checker1 = false;
            sAttr = ConfigurationManager.AppSettings.Get("Est_List_SourceSheet");

            //step 1
            foreach (Microsoft.Office.Interop.Excel.Worksheet worksheet in activeworkbook.Worksheets)
            {
                if (worksheet.Name.Contains(sAttr))
                {
                    Act_Est_Functions n = new Act_Est_Functions();
                    worksheet_target = n.WorkSheetRemake("Print List", activeworkbook);
                    worksheet_source = worksheet;
                    
                    checker1 = true;
                    break;
                }
            }

            if (checker1 == false)
            {
                MessageBox.Show("No Estimate List Worksheet Found");
                return;
            }
            //step 2

            int EstDataTitleRow = int.Parse(ConfigurationManager.AppSettings.Get("EstDataTitleRow"));
            int DataCount = 0;
            Range rRow = worksheet_source.Rows[EstDataTitleRow];
            //find out how many headers there are
            foreach (Range rCell in rRow.Cells)
            {
                if (rCell.Value2 == null)
                {
                    break;
                }
                else
                {
                    
                    //needs to loop this again as there are empty cell fields in the data table, first needs to get cell counts.
                    //rowData[DataCount] = rCell.ToString();
                    DataCount++;
                    //datacount will be 1 bigger than actual data as the way loop was desigend.
                    //Array.Resize(ref rowData, rowData.Length + 1);
                }
            }


            //step 3 mixed with step 2 a bit
            //in c# apparently that int[4] is from [0] to [3], wtf....
            //define string array to store data in the row
            string[] rowData = new string[DataCount];
            Active_Est_List active_Est_List = null;
            
            //read setting of column number for closing date stored at.
            //sDate needs to be added by 1 as excel column number starts with 1
            int sDate = int.Parse(ConfigurationManager.AppSettings.Get("Est_List_BidCloseDate")) +1;
            
            //use newArray to temporarily store the col mapping, so that this function is not getting called when reading each line.
            Act_Est_Functions m = new Act_Est_Functions();
            var newArray = m.Est_List_Col_Array();

            
            //Initialize where to write data on target sheet.
            int worksheet_target_Row = 13;

            //today's date
            DateTime today = DateTime.Today;

            //datatable to store row #, and date (excel date data) for bid close date in the future, for sorting
            //emm nope, using 2 single array then sort together...
            int[] FutureEst_RowNum = new int[0];
            double[] FutureEst_DateArray = new double[0];
            int TotalLegitRow = 0;

            while (worksheet_source.Cells[EstDataTitleRow, 1].Value2 != null)
            {
                //rRow = worksheet_source.Rows[EstDataTitleRow];
                try
                {
                    var BidCloseDate = worksheet_source.Cells[EstDataTitleRow, sDate].Value;

                    //there is null fields, catch in the catch section
                    Type type = BidCloseDate.GetType();
                    //bool tempbool = (BidCloseDate >= today);
                    //Debug.WriteLine(tempbool);
                    if (type != typeof(System.DateTime))
                    {
                        //if the type is not datetime then the format is wrong
                        //but I haven't decided what happens here.
                        //with current try-catch logic this will roll to next row.
                        throw new DateTimeFormatError("Row " + EstDataTitleRow.ToString() + " Date Format Incorrect");

                    }

                    //step 4
                    else if (DateTime.Compare(BidCloseDate,today) >= 0)
                    {
                        //2020-05-20 code added below to first read all the dates and store the row # in an array.
                        Array.Resize(ref FutureEst_DateArray, FutureEst_DateArray.Length + 1);
                        Array.Resize(ref FutureEst_RowNum, FutureEst_RowNum.Length + 1);
                        FutureEst_RowNum[TotalLegitRow] = EstDataTitleRow;
                        FutureEst_DateArray[TotalLegitRow] = worksheet_source.Cells[EstDataTitleRow, sDate].Value2;
                        TotalLegitRow++;


                    }
                }
                catch (DateTimeFormatError error)
                {
                    //MessageBox.Show("DateTimeFormatError" + EstDataTitleRow);
                }
                catch (RuntimeBinderException error)
                {

                }
                //catch (Exception error)
                //{
                //    MessageBox.Show("exception" + EstDataTitleRow);
                //}
                //Next Row
                EstDataTitleRow++;
            }

            if (FutureEst_DateArray.Length != 0) //if the array is not empty, then add the header, if empty then prompt and exit program
            {
                m.Est_Item_Header(worksheet_target);
            }
            else
            {
                MessageBox.Show("No Future Estimate Found");
                return;
            }

            //sorting 2 arrays together
            Array.Sort(FutureEst_DateArray, FutureEst_RowNum);

            worksheet_target.PageSetup.Zoom = false;
            worksheet_target.PageSetup.FitToPagesTall = false;
            worksheet_target.PageSetup.FitToPagesWide = 1;

            //go through the array to pick up the rows.
            //internal counter to check on the date, and see if a title needs to be generated
            int counter = 0;
            foreach (int RowNum in FutureEst_RowNum)
            {
                //set row
                rRow = worksheet_source.Rows[RowNum];
                
                //now read row into the array cell by cell
                int i = 0;
                foreach (Range rCell in rRow.Cells)
                {
                    string dynamic = rCell.Text;//string type to ease combining and avoid error
                    rowData[i] = dynamic;
                    ++i;
                    if (i >= DataCount)
                    {
                        break;
                        // exit for once reaches the last column, remember DataCount is the size of the array
                        // from 0 to size-1
                        // still like wtf...
                    }

                }

                //step 5
                //the following needs to be looped till sheet runs out.
                active_Est_List = new Active_Est_List(rowData, newArray);
                
                //check if bid close at the same date, if it is not date with previous one then add header
                if (counter == 0 || (FutureEst_DateArray[counter] != FutureEst_DateArray[counter - 1]))
                {
                    Range range = worksheet_target.Range[worksheet_target.Cells[worksheet_target_Row,1], worksheet_target.Cells[worksheet_target_Row, 5]];
                    range.Value = active_Est_List.Est_List_BidCloseDate;
                    range.Merge();
                    range.HorizontalAlignment = XlHAlign.xlHAlignLeft;
                    range.VerticalAlignment = XlVAlign.xlVAlignCenter;
                    range.Font.Name = "Calibri";
                    range.Font.Size = 14;
                    range.Font.Bold = true;
                    range.Font.Color = ColorTranslator.ToOle(Color.White);
                    range.Interior.Color = ColorTranslator.FromHtml("#9c88b3");
                    range.NumberFormat = "dddd, mmmm d, yyyy";
                    worksheet_target_Row++;
                    
                }

                m.Est_Item_CopyPasteFormat(active_Est_List, worksheet_target, worksheet_target_Row, 1);
                worksheet_target_Row += 10;
                
                if ((counter % 5) == 3 || counter == 3) //add page break every 5 items, with exception to the first one, as 2 items added for example and header.
                {
                    worksheet_target.HPageBreaks.Add(worksheet_target.Rows[worksheet_target_Row]);
                }

                counter++;
            }

            

            foreach (VPageBreak vPageBreak in worksheet_target.VPageBreaks)
            {
                vPageBreak.Delete();
            }

            worksheet_target.VPageBreaks.Add(worksheet_target.Columns[5]);

            

            worksheet_target.PageSetup.PaperSize = XlPaperSize.xlPaperLetter;
            
            worksheet_target.PageSetup.PrintArea = "A1:E"+worksheet_target_Row.ToString();

            stopWatch.Stop();
        }

        //Mapping setup
        private void Btn_ColMap_Click(object sender, RibbonControlEventArgs e)
        {
            MappingForm_Make n = new MappingForm_Make();
            n.Est_List_Col_Mapping();
        }

        public class DateTimeFormatError:Exception
        {
            public DateTimeFormatError(string message): base(message)
            {

            }
        }
    }

    public class MappingForm_Make //stores functions for making mapping forms
    {
        //2020-05-10, call out the form to confirm mapping of the columns for active estimate list
        //this form needs to be called out, when the button was clicked
        //2020-05-12 completed and tested no bug found
        public void Est_List_Col_Mapping()
        {

            using (Form Form1 = new Form())
            {
                Form1.Text = "Mapping";
                Type type = typeof(Active_Est_List);
                PropertyInfo[] properties = type.GetProperties();

                //form drawing setting
                int labelstartpoint_v = 10;//vertical start
                int labelstartpoint_h = 10;//horizontal start
                int labelincrement = 26;

                //xml config key value storage
                string sAttr;

                foreach (PropertyInfo property in properties)
                {
                    //add label if type is string, add textbox if type is interger
                    //no longer have interger as the column index no longer properties of class
                    //it will first read value from xml file and then show in textbox.
                    //once the OK key is hit it will then update the xml file.
                    System.Windows.Forms.Label newlabel = new System.Windows.Forms.Label
                    {
                        Size = new System.Drawing.Size(130, 13),
                        Location = new System.Drawing.Point(labelstartpoint_h + 50, labelstartpoint_v),
                        Text = property.Name,
                        AutoSize = true
                    };

                    Form1.Controls.Add(newlabel);

                    sAttr = ConfigurationManager.AppSettings.Get(property.Name);

                    System.Windows.Forms.TextBox newtextbox = new System.Windows.Forms.TextBox()
                    {
                        Size = new Size(30, 13),
                        Location = new System.Drawing.Point(labelstartpoint_h, labelstartpoint_v),
                        //text of the textbox is from public array
                        Text = sAttr,
                        Name = property.Name,
                        
                    };
                    
                    Form1.Controls.Add(newtextbox);
                    labelstartpoint_v += labelincrement;
                }
                Form1.AutoSize = true;
                Form1.AutoSizeMode = AutoSizeMode.GrowAndShrink;

                //add a OK button
                System.Windows.Forms.Button newButton = new System.Windows.Forms.Button()
                {
                    Size = new System.Drawing.Size(130, 40),
                    Location = new System.Drawing.Point(labelstartpoint_h + 40, labelstartpoint_v),
                    Text = "OK",

                };
                Form1.Controls.Add(newButton);
                newButton.Click += new EventHandler(NewButton_Click);
                Form1.ShowDialog();
                //when Ok button click, update to xml file of column mapping
                void NewButton_Click(object sender, EventArgs e)
                {
                    Configuration configuration = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
                    Configuration config = configuration;
                    foreach (Control control in Form1.Controls)
                    {
                        if (control is System.Windows.Forms.TextBox)
                        {
                            config.AppSettings.Settings[control.Name].Value = control.Text;
                        }
                    };
                    config.Save(ConfigurationSaveMode.Modified);
                    ConfigurationManager.RefreshSection("appSettings");
                    Form1.Dispose();
                    Form1.Close();
                }
            }
        }
    }

    public class Act_Est_Functions //functions used in making active estimate list are stored here.
    {
        //2020-05-08 generate a mapping column array based on amount of properties in class.
        public int[] Est_List_Col_Array()
        {
            Type type = typeof(Active_Est_List);
            PropertyInfo[] properties = type.GetProperties();
            int i = properties.Count();
            int[] Est_List_Col_Array = new int[i];
            i = 0;
            string sAttr;
            foreach (PropertyInfo property in properties)
            {
                sAttr = ConfigurationManager.AppSettings.Get(property.Name);
                Est_List_Col_Array[i] = int.Parse(sAttr);
                //Debug.WriteLine(sAttr);
                i++;
            }
            return Est_List_Col_Array;
        }
        public Worksheet WorkSheetRemake(string sheet_name, Excel.Workbook workbook)
        {
            foreach (Worksheet worksheet in workbook.Worksheets)
            {
                if (worksheet.Name.Contains(sheet_name))
                {
                    worksheet.Delete();
                }
            }
            int count = workbook.Worksheets.Count;
            Worksheet newsheet = workbook.Worksheets.Add(Type.Missing, workbook.Worksheets[count], Type.Missing, XlSheetType.xlWorksheet);
            newsheet.Name = sheet_name;
            return newsheet;
        }
        public void Est_Item_Header (Worksheet worksheet)
        {
            var range = worksheet.Range["A1:E2"];
            range.Value = "SUBMISSION SUMMARY";
            range.Merge();
            range.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            range.VerticalAlignment = XlVAlign.xlVAlignCenter;
            range.Font.Name = "Calibri";
            range.Font.Size = 14;
            range.Font.Bold = true;
            range.Font.Italic = true;
            range.Font.Color = ColorTranslator.ToOle(Color.White);
            range.Interior.Color = ColorTranslator.FromHtml("#9c88b3");
            worksheet.Columns[1].ColumnWidth = 32;
            worksheet.Columns[2].ColumnWidth = 32;
            worksheet.Columns[3].ColumnWidth = 3;
            worksheet.Columns[4].ColumnWidth = 16;
            worksheet.Columns[5].ColumnWidth = 15;

            worksheet.Cells[3, 1] = "Synergy Projects Ltd.";
            range = worksheet.Cells[3, 1];
            range.Font.Bold = true;
            range.Font.Size = 9;
            range = worksheet.Cells[3, 5];
            range.Value = DateTime.Today.Date;
            range.NumberFormat = "dddd, mmmm d, yyyy";
            range.Font.Size = 8;
            range = worksheet.Range["A4:E4"];
            range.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;

            worksheet.Rows[4].RowHeight = 4;

            worksheet.Cells[5, 1] = "Project Name";
            worksheet.Cells[5, 2] = "Location";
            worksheet.Cells[5, 4] = "Designated Est. Lead";
            worksheet.Cells[5, 5] = "Designated Est. Lead";

            worksheet.Cells[6, 1] = "Estimated Value / C Term";
            worksheet.Cells[6, 2] = "Owner Name";
            worksheet.Cells[6, 4] = "Chief Estiamtor";
            worksheet.Cells[6, 5] = "Chief Estimator";

            worksheet.Cells[7, 1] = "Submission Type";
            worksheet.Cells[7, 2] = "Source of Project Funds";
            worksheet.Cells[7, 5] = "SPL President";

            worksheet.Cells[8, 1] = "Category / # of Bidders";
            worksheet.Cells[8, 2] = "Designer Name";
            worksheet.Cells[8, 5] = "SGC President";

            worksheet.Cells[9, 1] = "Project Delivery / Contract";
            worksheet.Cells[9, 2] = "Bid Closing Time";
            worksheet.Cells[9, 5] = "Const. Ops. Manager";

            worksheet.Cells[10, 1] = "Estimate No.";
            worksheet.Cells[10, 2] = "Pre-Bid Meeting Date and time";
            worksheet.Cells[10, 4] = "Estimator / Prop Writer";
            worksheet.Cells[10, 5] = "VP / COO";

            range = worksheet.Range["A5:E10"];
            range.Font.Size = 9;

            range = worksheet.Range["A10:E10"];
            range.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;

            worksheet.Rows[11].RowHeight = 4;
            worksheet.Rows[12].RowHeight = 4;
        }
        public void Est_Item_CopyPasteFormat (Active_Est_List list, Worksheet worksheet, int row, int col)//(class, targetsheet, target row, target col)
        {
            int i = col; //store the col value
            Range range;
            
            bool cterm = (list.Est_List_CTerm != "No C-Term");
            //first row
            worksheet.Cells[row, col  ] = list.Est_List_ProjectName;
            worksheet.Cells[row, col].Font.Size = 9;
            worksheet.Cells[row, ++col] = list.Est_List_Location();
            worksheet.Cells[row, ++col] = "";
            worksheet.Cells[row, ++col] = list.Est_List_EstLead;
            worksheet.Cells[row, ++col] = list.Est_List_EstLeadRevDate;
            range = worksheet.Cells[row, col];
            range.Interior.Color = ColorTranslator.FromHtml("#ffcccc");
            worksheet.Rows[row++].Font.Size = 9;
            col = i;
            //second row
            worksheet.Cells[row, col  ] = list.Est_List_EstValue;
            var estvalue = (worksheet.Cells[row, col].Value2);
            try
            {
                estvalue = estvalue / 1000000;
            }
            catch (Exception error)
            {
                estvalue = 0.5;
            }

            worksheet.Cells[row, ++col] = list.Est_List_OwnerName;
            worksheet.Cells[row, ++col] = "";
            worksheet.Cells[row, ++col] = list.Est_List_ChiefEstimator;
            worksheet.Cells[row, ++col] = list.Est_List_ChfEstRevDate;

            //Review critiria
            range = worksheet.Cells[row, col];
            if (cterm == true)
            {
                if (estvalue < 1)
                {
                    range.Interior.Color = ColorTranslator.FromHtml("#fff5cc");
                }
                else
                {
                    range.Interior.Color = ColorTranslator.FromHtml("#fffccc");
                }
            }
            else
            {
                if (estvalue < 5)
                {
                    range.Interior.Color = ColorTranslator.FromHtml("#fff5cc");
                }
                else
                {
                    range.Interior.Color = ColorTranslator.FromHtml("#fffccc");
                }
            }

            worksheet.Rows[row++].Font.Size = 9;
            col = i;
            //third row
            worksheet.Cells[row, col  ] = list.Est_List_SubmissionType;
            worksheet.Cells[row, ++col] = list.Est_List_FundSource;
            worksheet.Cells[row, ++col] = "";
            worksheet.Cells[row, ++col] = "";
            worksheet.Cells[row, ++col] = list.Est_List_SPLPrsdntRevDate;

            //Review critiria
            range = worksheet.Cells[row, col];
            if (cterm == true)
            {
                if (estvalue < 1)
                {
                    range.Value = "";
                }
                else if (estvalue <5)
                {
                    range.Interior.Color = ColorTranslator.FromHtml("#fff5cc");
                }
                else
                {
                    range.Interior.Color = ColorTranslator.FromHtml("#fffccc");
                }
            }
            else
            {
                if (estvalue < 5)
                {
                    range.Value = "";
                }
                else if (estvalue < 15)
                {
                    range.Interior.Color = ColorTranslator.FromHtml("#fff5cc");
                }
                else
                {
                    range.Interior.Color = ColorTranslator.FromHtml("#fffccc");
                }
            }

            worksheet.Rows[row++].Font.Size = 9;
            col = i;
            //forth row
            worksheet.Cells[row, col  ] = list.Est_List_Category;
            worksheet.Cells[row, ++col] = list.Est_List_Designer;
            worksheet.Cells[row, ++col] = "";
            worksheet.Cells[row, ++col] = "";
            worksheet.Cells[row, ++col] = list.Est_List_SGCPrsdntRevDate;
            
            //review critiria
            range = worksheet.Cells[row, col];
            if (cterm == true)
            {
                if (estvalue < 5)
                {
                    range.Value = "";
                }
                else if (estvalue < 15)
                {
                    range.Interior.Color = ColorTranslator.FromHtml("#fff5cc");
                }
                else
                {
                    range.Interior.Color = ColorTranslator.FromHtml("#fffccc");
                }
            }
            else
            {
                if (estvalue < 15)
                {
                    range.Value = "";
                }
                else if (estvalue < 50)
                {
                    range.Interior.Color = ColorTranslator.FromHtml("#fff5cc");
                }
                else
                {
                    range.Interior.Color = ColorTranslator.FromHtml("#fffccc");
                }
            }
            worksheet.Rows[row++].Font.Size = 9;
            col = i;
            //fifth row
            worksheet.Cells[row, col  ] = list.Est_List_ProjDeliveryContractType();
            worksheet.Cells[row, ++col] = list.Est_List_BidCloseDateTime();
            worksheet.Cells[row, ++col] = "";
            worksheet.Cells[row, ++col] = "";
            worksheet.Cells[row, ++col] = list.Est_List_OpsMngrRevDate;
            //review critiria
            range = worksheet.Cells[row, col];
            if (cterm == true)
            {
                if (estvalue < 5)
                {
                    range.Value = "";
                }
                else if (estvalue < 15)
                {
                    range.Interior.Color = ColorTranslator.FromHtml("#fff5cc");
                }
                else
                {
                    range.Interior.Color = ColorTranslator.FromHtml("#fffccc");
                }
            }
            else
            {
                if (estvalue < 15)
                {
                    range.Value = "";
                }
                else if (estvalue < 50)
                {
                    range.Interior.Color = ColorTranslator.FromHtml("#fff5cc");
                }
                else
                {
                    range.Interior.Color = ColorTranslator.FromHtml("#fffccc");
                }
            }
            worksheet.Rows[row++].Font.Size = 9;
            col = i;
            //sixth row
            worksheet.Cells[row, col  ] = list.Est_List_EstNumber;
            worksheet.Cells[row, ++col] = list.Est_List_PreBidDT();
            worksheet.Cells[row, ++col] = "";
            worksheet.Cells[row, ++col] = list.Est_List_EstmtrPrpslWrtr;
            worksheet.Cells[row, ++col] = list.Est_List_VPCOO;
            //review critiria
            range = worksheet.Cells[row, col];
            if (cterm == true)
            {
                if (estvalue < 5)
                {
                    range.Value = "";
                }
                else if (estvalue < 15)
                {
                    range.Value = "";
                }
                else
                {
                    range.Interior.Color = ColorTranslator.FromHtml("#fffccc");
                }
            }
            else
            {
                if (estvalue < 15)
                {
                    range.Value = "";
                }
                else if (estvalue < 50)
                {
                    range.Value = "";
                }
                else
                {
                    range.Interior.Color = ColorTranslator.FromHtml("#fffccc");
                }
            }
            worksheet.Rows[row++].Font.Size = 9;
            worksheet.Rows[row++].RowHeight = 4;
            col = i;
            worksheet.Cells[row, col  ] = "Notes:";
            worksheet.Rows[row].Font.Size = 9;
            worksheet.Rows[++row].RowHeight = 4;
            worksheet.Rows[++row].RowHeight = 4;
        }
    }
    public class Active_Est_List //class for calling on passing info to new sheet.
    {
        //2020-05-07 each property has then a column matched property used for pulling data from the array that was put in, type is "int"
        //2020-05-11 column index will not be class properties, instead they will be stored as a public variables
        public string Est_List_ProjectName { get; set; }
        public string Est_List_City { get; set; }
        public string Est_List_Province { get; set; }
        public string Est_List_EstLead { get; set; } //Designated Estimate Lead
        public string Est_List_EstLeadRevDate { get; set; } //Estimate Lead Review Date
        public string Est_List_EstValue { get; set; } //Estimate Value
        public string Est_List_CTerm { get; set; } //C Terms
        public string Est_List_OwnerName { get; set; }
        public string Est_List_ChiefEstimator { get; set; }
        public string Est_List_ChfEstRevDate { get; set; }
        public string Est_List_SubmissionType { get; set; }
        public string Est_List_FundSource { get; set; }
        public string Est_List_SPLPrsdntRevDate { get; set; }
        public string Est_List_Category { get; set; } //Category
        public string Est_List_NumBidder { get; set; } //# of bidders
        public string Est_List_Designer { get; set; } // Designer Name
        public string Est_List_SGCPrsdntRevDate { get; set; }
        public string Est_List_ProjDelivery { get; set; } //Project Delivery
        public string Est_List_Contract { get; set; } //Contract
        public string Est_List_BidCloseTime { get; set; }
        public string Est_List_BidCloseDate { get; set; }
        public string Est_List_OpsMngrRevDate { get; set; }
        public string Est_List_EstNumber { get; set; }

        public string Est_List_PreBidDate { get; set; }// Prebid meeting Date
        public string Est_List_PreBidTime { get; set; }// Prebid meeting Time
        public string Est_List_EstmtrPrpslWrtr { get; set; } //proposal writer
        public string Est_List_VPCOO { get; set; }

        public string Est_List_BidCloseDateTime()
        {
            string DT = this.Est_List_BidCloseDate + " / " + this.Est_List_BidCloseTime;
            return DT;
        }
        public string Est_List_ProjDeliveryContractType()
        {
            string PDCT = this.Est_List_ProjDelivery + " / " + this.Est_List_Contract;
            return PDCT;
        }
        public string Est_List_EstValueCTerm()
        {
            string EstValueCTerm = this.Est_List_EstValue + " / " + this.Est_List_CTerm;
            return EstValueCTerm;
        }
        public string Est_List_CateNumBidder()
        {
            string CateNumBidder = this.Est_List_Category + " / " + this.Est_List_NumBidder;
            return CateNumBidder;
        }
        public string Est_List_Location()
        {
            string Location = this.Est_List_City + ", " + this.Est_List_Province;
            return Location;
        }
        public string Est_List_PreBidDT()// Prebid meeting Date and Time
        {
            string DateAndTime = this.Est_List_PreBidDate + " @ " + this.Est_List_PreBidTime;
            return DateAndTime;
        }
        public Active_Est_List(string[] Est_Line_Data, int[] Est_Col_Mapping)
        {
            //2020-05-06: expecting a line of excel data passed on to a string type array,
            //then passed on to here
            //values can be passed on to the properties based on the title of the sheet
            //2020-05-10: using a form to confirm the mapping, default value can be stored in file under the directory of the file.
            int i = 0;
            foreach (PropertyInfo property in typeof(Active_Est_List).GetProperties())
            {
                property.SetValue(this, Est_Line_Data[Est_Col_Mapping[i]]);
                //Debug.WriteLine(property.Name);
                //Debug.WriteLine(property.GetValue(this));
                //Debug.WriteLine(Est_Col_Mapping[i]);
                i++;
            }

        }
        
        
    }
}
