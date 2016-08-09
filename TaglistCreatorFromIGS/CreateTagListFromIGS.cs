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

    /*
     * Class: CreateTagListFromIGS
     * Inputs: Csv file to read from, ExcelFile name to output to.
     * Outputs: Creates a Excel Taglist with all the parameters
     */

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
         * Constructor:
         * Inputs: Entire file path of the csv file, ExcelFile name without extension
         * Logic: this initializes the object CreateTagListFromIGS while checking if the csv file actually exist. 
         */
        public CreateTagListFromIGS(string FileFullPathName, string excelfile)
        {
            fullPathToCSVDocument = FileFullPathName;
            excelFileName = excelfile;
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
         * this method is used to create an excel file based on the name of the csv file in the same folder where the csv file was located
         * This creates a copy of the standardTaglist excel template file 
        */
        private void createExcelFile()
        {
            try
            { 
            //this.excelFileName = Path.GetFileNameWithoutExtension(this.fullPathToCSVDocument);

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



        /* Method: initializeExcel
         * Logic: used to start the excel application and checks if it can actually be started.
         * Creates worksheets and workbooks 
        */
        private void initializeExcel()
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

        /* Method: createWorksheets
         * Logic: this method is used to create the number of worksheets based on how many unique subcontroller files there are
         * this method also changes the name of the worksheet based on the subcontroller files. 
         */
        // 
        private void createWorksheets()
        {

            // this foreach loop is used to change the name of the worksheet to the name of the subcontrollers 
            //and create X number of excel worksheet based on how many unique subcontrollers there are. 

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
         * Inputs: currently working on workbook, worksheet name in excel, list of ParameterInfo which stores each parameters info ( such as Name, address, data_type)
         * Used to write the stringList list to the worksheet that is identified by worksheetName within the workbook object
         */
        private void writeDataToExcelWorksheets(Excel.Workbook xlWB, string worksheetName, List<ParameterInfo> stringList)

        {
            xlActiveSheet =  xlWB.Sheets[worksheetName]; // select the worksheet based on the name provided. 
            xlActiveSheet.Activate();

            xlActiveSheet.Cells[3, 1] = worksheetName;




            // this is using LINQ to query the csvDataTable data structure and selecting the Tag_Name, Data_type, Address of all the parameter in a given subcontroller file

            var parameterInfoSingle = (from ParameterInfo in stringList select new { ParameterInfo.Tag_Name, ParameterInfo.Data_Type, ParameterInfo.Address }).ToArray();

            // this forloop is actually writing the data to the excel file
            int NumberOfTags = parameterInfoSingle.Count();
            for (int i = 0; i < NumberOfTags; i++)
            {
                xlActiveSheet.Cells[i + 6, 2] = i + 2; // update the Item number field in column B
                xlActiveSheet.Cells[i + 6, 3] = parameterInfoSingle[i].Tag_Name; // updates the tagname into Column C as noted by the number 3 
                xlActiveSheet.Cells[i + 6, 4] = parameterInfoSingle[i].Address; // updates the address into column D
                xlActiveSheet.Cells[i + 6, 6] = parameterInfoSingle[i].Data_Type; // updates the data type into column F


            }

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

        // method: readCSVFile
        // Input: CSV file
        // Output: Creates a data structure of type dictionary where the keys are the subcontroller names and the values are list of ParameterInfo objects
        // Key: Sub Controller File name For example CMN, ZW1 
        // Value: List of ParameterInfo object which hold the following parameters: Tag Name, Address, DataType
        // Dictionary structure shown below:
        // Dictionary[key] = List<ParameterInfo>

        private void readCSVFile()
        {

            csvDataTable = new Dictionary<string, List<ParameterInfo>>();

            // this string is used to add entries to the dictionary using this as a key. 
            string currentSubControllerFile;
            // these strings are used as intermediate holders to hold the current parameter information such as tag name, address, and data type
            string currentTagName;
            string currentAddress;
            string currentDataType;

            // create the data structure dictionary
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

            }

        }

        /* method: generateTagList
         * This method calls the private methods from this class in the correct order to create the excel taglist file
         */ 
        public void generateTagList()
        {

            try
            {
            readCSVFile();
            createExcelFile();
            initializeExcel();
            createWorksheets();

            // this forloop loops throught all the enteries in the dictionary data structure to write them to the excel file
            foreach (KeyValuePair<string, List<ParameterInfo>> entry in csvDataTable)
            {
                writeDataToExcelWorksheets(this.xlWorkBook, entry.Key, entry.Value);
            }


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

        

    }


}





//---------------------- Note: This section of the code is used to store all the old test code that was taken out of the actual code and did not end up being used. --------------



//xlActiveSheet.get_Range("C6", endrange).Value2 = TestList;



// var TestList = csvDataTable[csvDataTable.Keys.First()].Select(l => l.Tag_Name).ToArray();




//foreach (var parameter in parameterInfoSingle)
//{


//}

//string endrange = string.Format("C{0}", (TestList.Count() + 5).ToString());


//string[,] names = new string[,] { { "Matt", null ,"Matt2" }, { "Joanne", null ,"Joanne2" }, { "Robert", null ,"Robert2" } };
//string[,] names = new string[,] { { "Matt", "Joanne", "Robert" } }; // this does not work
//xlActiveSheet.get_Range("C6", endrange).Value2 = TestList;
//xlActiveSheet.get_Range("C6", "D8").Value = names;
//string[,] parameterInfoSingle = (from ParameterInfo in csvDataTable[csvDataTable.Keys.First()] select ParameterInfo.Tag_Name)




//var parameterInfoSingle = (from ParameterInfo in csvDataTable[csvDataTable.Keys.First()] select new { ParameterInfo.Tag_Name, ParameterInfo.Data_Type, ParameterInfo.Address }).ToArray();

////string[,] parameterInfoSingle = (from ParameterInfo in csvDataTable[csvDataTable.Keys.First()] select ParameterInfo.Tag_Name)
//int NumberOfTags = parameterInfoSingle.Count();
//for (int i = 0; i < NumberOfTags; i++)
//{
//    xlActiveSheet.Cells[i + 6, 2] = i + 2;
//    xlActiveSheet.Cells[i + 6, 3] = parameterInfoSingle[i].Tag_Name;
//    xlActiveSheet.Cells[i + 6, 4] = parameterInfoSingle[i].Address;
//    xlActiveSheet.Cells[i + 6, 6] = parameterInfoSingle[i].Data_Type;

//---------------------------------------------------------------------------

//fullPathToCSVDocument = @"C:\Users\212478881\Desktop\TestCSV Folder\011110TestSiteFullIGSDriver.csv";


//---------------------------------------------------------------------------


//List<ParameterInfo> newParameterList = new List<ParameterInfo>();
//ParameterInfo test = new ParameterInfo();
//test.Address = "address1212313";
//newParameterList.Add(test);
//string addresstest = newParameterList[0].Address;


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

//string test = csvDataTable.Keys.First();
//writeDataToExcelWorksheets(this.xlWorkBook, test, csvDataTable[test]);

//Dictionary<string,string> parameterfieldValues = new Dictionary<string, string> ();