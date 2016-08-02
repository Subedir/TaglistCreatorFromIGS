using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using Microsoft.VisualBasic.FileIO;


namespace TaglistCreatorFromIGS
{
    class CreateTagListFromIGS
    {
        //fields 

        // this ParameterInfo class is used to create a list which will store all the information for a parameter

        public static void readCSVFile()
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
