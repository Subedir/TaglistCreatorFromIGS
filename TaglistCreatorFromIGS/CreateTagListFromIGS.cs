using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using Microsoft.VisualBasic.FileIO;
using System.IO;
using System.Windows.Forms;
using System.Resources;

namespace TaglistCreatorFromIGS
{
    class CreateTagListFromIGS
    {
        //fields 
        private Excel.Application xlApp = null;
        Dictionary<string, List<ParameterInfo>> csvDataTable = null;
        private string fullPathToCSVDocument = "";
        private string excelFileName = ""; // this is used to name the excel file without the extension (.xlsx)
        private string excelFileFullPath = ""; // this is the full path with the excel file name
        Excel.Workbook xlWorkBook = null;
        Excel.Sheets worksheets = null;
        Excel.Worksheet xlActiveSheet = null;

        /*
         * this is the constructor
         */
        public CreateTagListFromIGS(string FileFullPathName)
        {
            fullPathToCSVDocument = FileFullPathName;

            try
            {
                if (!System.IO.File.Exists(fullPathToCSVDocument))
                { 
                    throw new System.IO.FileNotFoundException();
                }

            }
            catch (FileNotFoundException ex)
            {
                MessageBox.Show("File Not Found Error, Please check if the IGS File exist", "Error");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Error");
            }

        }


        /* Method: createExcelFile
         * this method is used to create an excel file based on the name of the csv file in the same folder as where the 
         * 
        */
        private void createExcelFile()
        {
            try
            { 
            this.excelFileName = Path.GetFileNameWithoutExtension(this.fullPathToCSVDocument);

            string sourcePath = Path.GetDirectoryName(this.fullPathToCSVDocument); // this is directory of where the .csv file is located
            excelFileFullPath = System.IO.Path.Combine(sourcePath, excelFileName + ".xlsx");


            if (IsFileLocked(new FileInfo(excelFileFullPath)))
                {
                    System.IO.File.WriteAllBytes(excelFileFullPath, Properties.Resources.StandardTagList);  // this is used to copy the excel file in this projects resources and create a copy in the folder where the csv is located
                }

            else
                {
                    //throw();
                    throw new Exception(String.Format(@"IOException: The file {0} is currently being used. {1}Follow the following steps: 
                                                                    {1} 1. Delete the file 
                                                                    {1} 2. If unable to delete; Close all programs using this file 
                                                                    {1} 3. Rerun this program", excelFileFullPath,Environment.NewLine));
                }
            
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }
        }



        // Method: initializeExcel
        public void initializeExcel()
        {
            xlApp = new Excel.Application();

            if (xlApp == null)
            {

                MessageBox.Show("ERROR: EXCEL couldn't be started!", "Error");
                System.Windows.Forms.Application.Exit();

            }
            xlApp.DisplayAlerts = false;
            xlWorkBook = xlApp.Workbooks.Open(excelFileFullPath, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            worksheets = xlWorkBook.Worksheets;
            xlActiveSheet = xlWorkBook.Sheets[1];
        }

        // this method is used to create the number of worksheets based on how many unique subcontroller files there are
        // this method also changes the name of the worksheet based on the subcontroller files. 
        private void createWorksheets()
        {

            // this section of the code is used to create X number of excel worksheet based on how many unique subcontrollers there are. 

            //int uniqueSubControllerCount = csvDataTable.Count;

            //for(int i = 0; i < uniqueSubControllerCount-1; i++) // reason for -1 is because the template already has 1 worksheet
            //{
            //    xlSht.Copy(Type.Missing, xlWorkBook.Sheets[xlWorkBook.Sheets.Count]);

            //}




            
            //xlWorkBook.Sheets[xlWorkBook.Sheets.Count].name = xlSheetName;

            // this foreach loop is used to change the name of the worksheet to the name of the subcontrollers 

            foreach (KeyValuePair<string, List<ParameterInfo>> entry in csvDataTable)
            {
                Debug.WriteLine(entry.Key.ToString() +" " + xlWorkBook.Sheets.Count.ToString());
                xlActiveSheet.Copy(Type.Missing, xlWorkBook.Sheets[xlWorkBook.Sheets.Count]);
                xlWorkBook.Sheets[xlWorkBook.Sheets.Count].name = entry.Key;
            }

            xlWorkBook.Sheets[1].delete(); // this is used to delete the first worksheet in the excel file that is marked as template

        }

        /*
         * Method: writeDataToExcelWorksheets
         * Used to write the stringList list to the worksheet that is identified by worksheetName within the workbook object
         */
        private void writeDataToExcelWorksheets(Excel.Workbook xlWB, string worksheetName, List<ParameterInfo> stringList)

        {
            xlActiveSheet =  xlWB.Sheets[worksheetName];
            xlActiveSheet.Activate();

            xlActiveSheet.Cells[3, 1] = worksheetName;
            Debug.WriteLine("test");



            //xlActiveSheet.get_Range("C6", endrange).Value2 = TestList;



            // var TestList = csvDataTable[csvDataTable.Keys.First()].Select(l => l.Tag_Name).ToArray();

            // this is using LINQ to query the csvDataTable data structure and selecting the Tag_Name, Data_type, Address of all the parameter in a given subcontroller file


            var parameterInfoSingle = (from ParameterInfo in csvDataTable[csvDataTable.Keys.First()] select new {  ParameterInfo.Tag_Name,  ParameterInfo.Data_Type ,  ParameterInfo.Address }).ToArray();

            //string[,] parameterInfoSingle = (from ParameterInfo in csvDataTable[csvDataTable.Keys.First()] select ParameterInfo.Tag_Name)
            int NumberOfTags = parameterInfoSingle.Count();
            for (int i = 0; i < NumberOfTags ;i++)
            {
                xlActiveSheet.Cells[i + 6, 2] = i + 2;
                xlActiveSheet.Cells[i + 6, 3] = parameterInfoSingle[i].Tag_Name;
                xlActiveSheet.Cells[i + 6, 4] = parameterInfoSingle[i].Address;
                xlActiveSheet.Cells[i + 6, 6] = parameterInfoSingle[i].Data_Type;
            }
            //foreach (var parameter in parameterInfoSingle)
            //{
                

            //}

            //string endrange = string.Format("C{0}", (TestList.Count() + 5).ToString());


            //string[,] names = new string[,] { { "Matt", null ,"Matt2" }, { "Joanne", null ,"Joanne2" }, { "Robert", null ,"Robert2" } };
            //string[,] names = new string[,] { { "Matt", "Joanne", "Robert" } }; // this does not work
            //xlActiveSheet.get_Range("C6", endrange).Value2 = TestList;
            //xlActiveSheet.get_Range("C6", "D8").Value = names;
        }

        private void testMethod()
        {

            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;
            Excel.Range chartRange;

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            //add data 
            xlWorkSheet.Cells[4, 2] = "";
            xlWorkSheet.Cells[4, 3] = "Student1";
            xlWorkSheet.Cells[4, 4] = "Student2";
            xlWorkSheet.Cells[4, 5] = "Student3";

            xlWorkSheet.Cells[5, 2] = "Term1";
            xlWorkSheet.Cells[5, 3] = "80";
            xlWorkSheet.Cells[5, 4] = "65";
            xlWorkSheet.Cells[5, 5] = "45";

            xlWorkSheet.Cells[6, 2] = "Term2";
            xlWorkSheet.Cells[6, 3] = "78";
            xlWorkSheet.Cells[6, 4] = "72";
            xlWorkSheet.Cells[6, 5] = "60";

            xlWorkSheet.Cells[7, 2] = "Term3";
            xlWorkSheet.Cells[7, 3] = "82";
            xlWorkSheet.Cells[7, 4] = "80";
            xlWorkSheet.Cells[7, 5] = "65";

            xlWorkSheet.Cells[8, 2] = "Term4";
            xlWorkSheet.Cells[8, 3] = "75";
            xlWorkSheet.Cells[8, 4] = "82";
            xlWorkSheet.Cells[8, 5] = "68";

            xlWorkSheet.Cells[9, 2] = "Total";
            xlWorkSheet.Cells[9, 3] = "315";
            xlWorkSheet.Cells[9, 4] = "299";
            xlWorkSheet.Cells[9, 5] = "238";
            xlWorkSheet.get_Range("B2", "E3").Merge(false);

            xlWorkSheet.get_Range("b2", "e3").Merge(false);

            chartRange = xlWorkSheet.get_Range("b2", "e3");
            chartRange.FormulaR1C1 = "MARK LIST";
            chartRange.HorizontalAlignment = 3;
            chartRange.VerticalAlignment = 3;
            chartRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
            chartRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
            chartRange.Font.Size = 20;

            chartRange = xlWorkSheet.get_Range("b4", "e4");
            chartRange.Font.Bold = true;
            chartRange = xlWorkSheet.get_Range("b9", "e9");
            chartRange.Font.Bold = true;

            chartRange = xlWorkSheet.get_Range("b2", "e9");
            chartRange.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);

            xlWorkBook.SaveAs("d:\\csharp.net-informations.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            releaseObject(xlApp);
            releaseObject(xlWorkBook);
            releaseObject(xlWorkSheet);

            MessageBox.Show("File created !");
        }


        /*
         * Method: saveCloseExcelFile
         * this method serves 2 purpose: 
         * 1. used to Save, close the excel workbook,  
         * 2. Releases all the objects created to work with excel file such as worksheets, worksheet, workbook, excel application. 
         * This will free up any memory that is not needed by these objects. 
         * Parameters: Excel.Worksheet xlSht, Excel.Sheets worksheets, Excel.Workbook xlWorkBook, Excel.Application xlApp
         * These are all excel objects need to work with excel files. 
         */
        private void saveCloseExcelFile(Excel.Worksheet xlSht, Excel.Sheets worksheets, Excel.Workbook xlWorkBook, Excel.Application xlApp)
        {
            xlWorkBook.Save();
            xlWorkBook.Close();

            releaseObject(xlSht);
            releaseObject(worksheets);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);
        }

        // this is the overloaded method without any parameters
        // used to close the class objects
        private void saveCloseExcelFile()
        {
            xlWorkBook.Save();
            xlWorkBook.Close();

            releaseObject(xlActiveSheet);
            releaseObject(worksheets);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);
        }


        // Method releaseObject
        // This is used to clear up the object passed as parameter
        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                Debug.WriteLine("Exception Occured while releasing object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

        // this method is used to check if the file is currently being used.
        protected virtual bool IsFileLocked(FileInfo file)
        {
            FileStream stream = null;

            try
            {
                stream = file.Open(FileMode.Open, FileAccess.Read, FileShare.None);
            }
            catch (IOException)
            {
                //the file is unavailable because it is:
                //still being written to
                //or being processed by another thread
                //or does not exist (has already been processed)
                return true;
            }
            finally
            {
                if (stream != null)
                    stream.Close();
            }

            //file is not locked
            return false;
        }


        public void readCSVFile()
        {
            //---------------------------------------------------------------------------

            fullPathToCSVDocument = @"C:\Users\212478881\Desktop\TestCSV Folder\011110TestSiteFullIGSDriver.csv";


            //---------------------------------------------------------------------------


            //List<ParameterInfo> newParameterList = new List<ParameterInfo>();
            //ParameterInfo test = new ParameterInfo();
            //test.Address = "address1212313";
            //newParameterList.Add(test);
            //string addresstest = newParameterList[0].Address;



            // this is used to create a dictionary where the key is the name of the subcontroller file and the value is a ParameterInfo object list
            // Key: Sub Controller File name For example CMN, ZW1
            // Value: List of ParameterInfo object which hold the following parameters: Tag Name, Address, DataType

            csvDataTable = new Dictionary<string, List<ParameterInfo>>();

            // this string is used to add entries to the dictionary using this as a key. 
            string currentSubControllerFile;
            // these strings are used as intermediate holders to hold the current parameter information such as tag name, address, and data type
            string currentTagName;
            string currentAddress;
            string currentDataType;

            using (TextFieldParser parser = new TextFieldParser(fullPathToCSVDocument))
            {
                parser.TextFieldType = FieldType.Delimited;
                parser.SetDelimiters(",");

                //this single readfields will read in the header column from the CSV
                string[] fields = parser.ReadFields();
                Debug.WriteLine(fields.Length);

                string Tag_Name_Header = fields[0];
                string Address_Header = fields[1];
                string Data_Type_Header = fields[2];

                // this while loop is used to read the entire csv file into a list of type ParameterInfo type
                while (!parser.EndOfData)
                {
                    fields = parser.ReadFields();

                    currentSubControllerFile = fields[0].Split('.')[0];
                    currentTagName = fields[0].Split('.')[1];
                    currentAddress = fields[1];
                    currentDataType = fields[2];


                    // this checks to see if the key (SubController File) exist already exist in the dictionary
                    if (csvDataTable.ContainsKey(currentSubControllerFile))
                    {
                        ParameterInfo currentParameterInfo = new ParameterInfo(currentTagName, currentAddress, currentDataType); // this create a new object with TagName, Address, DataType fields
                        csvDataTable[currentSubControllerFile].Add(currentParameterInfo);

                    }

                    else
                    {
                        List<ParameterInfo> newParameterList = new List<ParameterInfo>(); // this creates a new list to be added to the dictionary
                        ParameterInfo currentParameterInfo = new ParameterInfo(currentTagName, currentAddress, currentDataType); // this create a new object with TagName, Address, DataType fields
                        newParameterList.Add(currentParameterInfo); // this adds the ParamaterInfo Object called "currentParameterInfo" to the newly created list
                        csvDataTable.Add(currentSubControllerFile, newParameterList);
                    }

                    
                }

                Debug.WriteLine("test");

                foreach (KeyValuePair<string, List<ParameterInfo>> entry in csvDataTable)
                {
                    Debug.WriteLine(entry.Key);
                }

                //fields = parser.ReadFields();
                //foreach (string field in fields)
                //{
                //    Debug.WriteLine(field);

                //}




                //while (!parser.EndOfData)
                //{
                //    string[] fields = parser.ReadFields();
                //    Debug.WriteLine(fields.ToString());

                //    foreach (string field in fields)
                //    {
                //        System.Console.WriteLine(field);

                //    }
                //}

            }

        }

        public void generateTagList()
        {

            try
            {

            
            readCSVFile();
            createExcelFile();
            initializeExcel();
            createWorksheets();
            string test = csvDataTable.Keys.First();
            writeDataToExcelWorksheets(this.xlWorkBook, test, csvDataTable[test]);
            saveCloseExcelFile();
            }

            catch(Exception ex)
            {
                saveCloseExcelFile();
                MessageBox.Show(string.Format("Program exited with the following error: {0}", ex.ToString()));

            }
        }



    }
    // this ParameterInfo class is used to create a list which will store all the information for a parameter
    class ParameterInfo
    {
        public ParameterInfo(string tagname, string address, string datatype)
        {
            this.Tag_Name = tagname;
            this.Address = address;
            this.Data_Type = datatype;

        }
        public string Tag_Name { get; set; }
        public string Address { get; set; }
        public string Data_Type { get; set; }

        //Dictionary<string,string> parameterfieldValues = new Dictionary<string, string> ();

    }


}
