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
        Excel.Worksheet xlSht = null;


        public CreateTagListFromIGS(string FileFullPathName)
        {
            fullPathToCSVDocument = FileFullPathName;

            try
            {
                if (!System.IO.File.Exists(fullPathToCSVDocument))
                { 
                    throw new System.IO.FileNotFoundException();
                }
                createExcelFile();

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
        private void startExcel()
        {
            xlApp = new Excel.Application();

            if (xlApp == null)
            {

                MessageBox.Show("ERROR: EXCEL couldn't be started!","Error");
                System.Windows.Forms.Application.Exit();

            }

        }

        /* Method: createExcelFile
         * this method is used to create an excel file based on the name of the csv file. 
         * 
        */
        private void createExcelFile()
        {
            this.excelFileName = Path.GetFileNameWithoutExtension(this.fullPathToCSVDocument);

            string sourcePath = Path.GetDirectoryName(this.fullPathToCSVDocument); // this is directory of where the .csv file is located
            string destFile = System.IO.Path.Combine(sourcePath, excelFileName + ".xlsx");
            System.IO.File.WriteAllBytes(destFile, Properties.Resources.StandardTagList);  // this is used to copy the excel file in this projects resources and create a copy in the folder where the csv is located
            
        }
        public void writeDataToExcelWorksheets(string worksheetName, List<string> stringList)

        {


        }


        // Method: initializeExcel
        public void initializeExcel()
        {
            

        }

        // this method is used to create worksheets
        public void createWorksheets(Excel.Application xlApp, string excelFilePath, string xlSheetName)
        {

            xlApp.DisplayAlerts = false;
            string filePath = excelFilePath;
            Excel.Workbook xlWorkBook = xlApp.Workbooks.Open(filePath, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            Excel.Sheets worksheets = xlWorkBook.Worksheets;
            Excel.Worksheet xlSht = xlWorkBook.Sheets[1];
            xlSht.Copy(Type.Missing, xlWorkBook.Sheets[xlWorkBook.Sheets.Count]);
            xlWorkBook.Sheets[xlWorkBook.Sheets.Count].name = xlSheetName;


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




        public void readCSVFile()
        {
            //---------------------------------------------------------------------------

            string csvFilePath = @"C:\Users\212478881\Documents\Visual Studio 2015\Projects\TestConsoleApplication\TestConsoleApplication\011110TestSiteFullIGSDriver.csv";


            //---------------------------------------------------------------------------


            //List<ParameterInfo> newParameterList = new List<ParameterInfo>();
            //ParameterInfo test = new ParameterInfo();
            //test.Address = "address1212313";
            //newParameterList.Add(test);
            //string addresstest = newParameterList[0].Address;



            // this is used to create a dictionary where the key is the name of the subcontroller file and the value is a ParameterInfo object list
            // Key: Sub Controller File name For example CMN, ZW1
            // Value: List of ParameterInfo object which hold the following parameters: Tag Name, Address, DataType

            Dictionary<string, List<ParameterInfo>> csvDataTable = new Dictionary<string, List<ParameterInfo>>();

            // this string is used to add entries to the dictionary using this as a key. 
            string currentSubControllerFile;
            // these strings are used as intermediate holders to hold the current parameter information such as tag name, address, and data type
            string currentTagName;
            string currentAddress;
            string currentDataType;

            using (TextFieldParser parser = new TextFieldParser(csvFilePath))
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
