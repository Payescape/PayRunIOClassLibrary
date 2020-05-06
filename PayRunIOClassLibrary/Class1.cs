using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using System.IO;
using System.Data.SqlClient;
using System.Data;
using System.Text.RegularExpressions;
using System.Xml;
using Microsoft.VisualBasic.FileIO;
using System.Globalization;
using System.Net.Mail;
using PayRunIO.CSharp.SDK;
using PicoXLSX;
using DevExpress.XtraReports.UI;
using Amazon;
using Amazon.S3;
using Amazon.S3.Transfer;
using Amazon.S3.Model;

namespace PayRunIOClassLibrary
{
    public class PayRunIOWebGlobeClass
    {
        //Changed by Jim Borland on 29/1/2020 at 10:20 an a bit more.
        public PayRunIOWebGlobeClass() { }

        //Testing making a change to the class 
        
        public void UpdateContactDetails(XDocument xdoc)
        {
            string contactsFolder = xdoc.Root.Element("DataHomeFolder").Value + "PE-Contacts\\";
            string dataSource = xdoc.Root.Element("DataSource").Value;            //"APPSERVER1\\MSSQL";  //"13.69.154.210\\MSSQL"; 
            string dataBase = xdoc.Root.Element("Database").Value;
            string userID = xdoc.Root.Element("Username").Value;
            string password = xdoc.Root.Element("Password").Value;
            string sqlConnectionString = "Server=" + dataSource + ";Database=" + dataBase + ";User ID=" + userID + ";Password=" + password + ";";

            DirectoryInfo dirInfo = new DirectoryInfo(contactsFolder);
            FileInfo[] files = dirInfo.GetFiles("*.csv");
            foreach (FileInfo file in files)
            {
                if (file.FullName.Contains("_contacts_"))
                {
                    //Get a table of contacts from the csv file.
                    DataTable dtContacts = GetDataTableFromCSVFile(xdoc, file.FullName);
                    //Insert the data into an SQL Database.
                    bool success = InsertDataIntoSQLServerUsingSQLBulkCopy(dtContacts, sqlConnectionString, file.FullName, xdoc);
                    if (success)
                    {
                        //We've successfully written the contact data to a temporary table with the name "tmp_CompanyNo_Contacts". e.g. "tmp_2137_Contacts"
                        //Now Insert / Update the contacts table then delete the table.
                        int x = file.FullName.LastIndexOf("\\") + 1;
                        string companyNo = file.FullName.Substring(x, 4);
                        success = InsertUpdateContacts(xdoc, sqlConnectionString, companyNo);
                        if (success)
                        {
                            //Delete the temporary contacts.
                            DeleteTemporaryContacts(xdoc, sqlConnectionString);
                            //Delete the csv file.
                            file.Delete();
                        }
                    }
                }

            }
        }
        private bool InsertDataIntoSQLServerUsingSQLBulkCopy(DataTable csvDataTable, string sqlConnectionString, string csvFileName, XDocument xdoc)
        {
            string configDirName = xdoc.Root.Element("SoftwareHomeFolder").Value;
            string textLine;

            using (SqlConnection sqlConnection = new SqlConnection(sqlConnectionString))
            {

                try
                {
                    sqlConnection.Open();
                    // Check if a table exsists
                    bool tableExists;
                    //
                    // Change the csvFileName to SQL table name here JCB TO DO
                    //
                    string tableName;
                    //This is the contacts file we've received from Web Globe it's named in the following format.
                    //CompanyNo_unity_contacts_export_datetimestamp.csv e.g. 1234_unity_contacts_export_20190630100130001.csv
                    //We just need the company number and contacts for the a table name.

                    tableName = "tmpContacts";  // Create a temporary invoices table and an SQL query will create the live one.


                    string sqlStatement = "SELECT COUNT (*) FROM " + tableName;


                    try
                    {
                        using (SqlCommand sqlCommand = new SqlCommand(sqlStatement, sqlConnection))
                        {
                            sqlCommand.ExecuteScalar();
                            tableExists = true;
                        }
                    }
                    catch
                    {
                        tableExists = false;
                    }

                    if (!tableExists)
                    {
                        // Create the table
                        try
                        {
                            textLine = string.Format("About to create tmpContacts table.");
                            update_Progress(textLine, configDirName, 1);

                            sqlStatement = "CREATE TABLE " + tableName + "(";
                            foreach (DataColumn dataColumn in csvDataTable.Columns)
                            {

                                dataColumn.ColumnName = Regex.Replace(dataColumn.ColumnName, "[^A-Za-z0-9]", "");
                                sqlStatement = sqlStatement + dataColumn.ColumnName + " varchar(150),";
                            }
                            sqlStatement = sqlStatement.Remove(sqlStatement.Length - 1, 1) + ")";
                            SqlCommand createTable = new SqlCommand(sqlStatement, sqlConnection);
                            createTable.ExecuteNonQuery();

                            textLine = string.Format("Sucessfully created tmpContacts table.");
                            update_Progress(textLine, configDirName, 1);
                        }
                        catch (Exception ex)
                        {
                            textLine = string.Format("Failed to create tmpContacts table.");
                            update_Progress(textLine, configDirName, 1);

                            return false;

                        }
                    }
                    try
                    {
                        using (SqlBulkCopy bulkData = new SqlBulkCopy(sqlConnection))
                        {
                            textLine = string.Format("About to bulk write to tmpContacts table.");
                            update_Progress(textLine, configDirName, 1);

                            bulkData.DestinationTableName = tableName;

                            foreach (DataColumn dataColumn in csvDataTable.Columns)
                            {
                                dataColumn.ColumnName = Regex.Replace(dataColumn.ColumnName, "[^A-Za-z0-9]", "");
                                bulkData.ColumnMappings.Add(dataColumn.ToString(), dataColumn.ToString());

                            }
                            //bulkData.BulkCopyTimeout = 600; // 600 seconds
                            bulkData.WriteToServer(csvDataTable);

                            textLine = string.Format("Successfull bulk write to tmpContacts table.");
                            update_Progress(textLine, configDirName, 1);

                            return true;

                        }
                    }
                    catch (Exception ex)
                    {
                        textLine = string.Format("Failed bulk write to tmpContacts table.");
                        update_Progress(textLine, configDirName, 1);

                        return false;

                    }
                }
                catch
                {
                    return false;

                }
                finally
                {
                    sqlConnection.Close();

                }

            }
        }
        private DataTable GetDataTableFromCSVFile(XDocument xdoc, string csvFileName)
        {
            int logOneIn = Convert.ToInt32(xdoc.Root.Element("LogOneIn").Value);
            string configDirName = xdoc.Root.Element("SoftwareHomeFolder").Value;
            string textLine = null;

            string delimiter = ",";
            DataTable csvDataTable = new DataTable();
            try
            {
                using (TextFieldParser csvReader = new TextFieldParser(csvFileName))
                {
                    csvReader.SetDelimiters(new string[] { delimiter });
                    csvReader.HasFieldsEnclosedInQuotes = true;

                    string[] colFields = csvReader.ReadFields();

                    foreach (string column in colFields)
                    {

                        DataColumn datacolumn = new DataColumn(column);
                        datacolumn.AllowDBNull = true;
                        //
                        // Check to make sure we don't have two columns with the same name.
                        //
                        try
                        {
                            csvDataTable.Columns.Add(datacolumn);
                        }
                        catch (Exception ex)
                        {
                            //
                            // We do have a column with this name already.
                            //
                            if (ex.ToString().Contains("already belongs to"))
                            {
                                DateTime dateTimeNow = DateTime.Now;
                                DataColumn dataColumnUnique = new DataColumn(column + dateTimeNow);
                                csvDataTable.Columns.Add(dataColumnUnique);
                            }
                            else
                            {
                                textLine = string.Format("Error getting data from csv file.\r\n{0}.\r\n", ex);
                                update_Progress(textLine, configDirName, logOneIn);
                            }

                        }

                    }

                    while (!csvReader.EndOfData)
                    {
                        string[] fieldData = csvReader.ReadFields();
                        int x = fieldData.Count();
                        string[] tableData = new string[x];
                        for (int i = 0; i < x; i++)
                        {
                            tableData[i] = fieldData[i];
                        }


                        csvDataTable.Rows.Add(tableData);
                    }
                }

            }
            catch (Exception ex)
            {
                textLine = string.Format("Error getting data from csv file.\r\n{0}.\r\n", ex);
                update_Progress(textLine, configDirName, logOneIn);

            }
            return csvDataTable;
        }
        private bool InsertUpdateContacts(XDocument xdoc, string sqlConnectionString, string companyNo)
        {
            int logOneIn = Convert.ToInt32(xdoc.Root.Element("LogOneIn").Value);
            string configDirName = xdoc.Root.Element("SoftwareHomeFolder").Value;
            string textLine = null;

            bool success = false;
            //
            //Try using a stored procedure
            //
            try
            {
                using (var connection = new SqlConnection(sqlConnectionString))
                using (var command = new SqlCommand("InsertUpdateContacts", connection)
                {
                    CommandType = CommandType.StoredProcedure
                })
                {
                    connection.Open();
                    command.Parameters.AddWithValue("CompanyNo", companyNo);
                    command.ExecuteNonQuery();
                    success = true;
                }
            }
            catch (Exception ex)
            {
                textLine = string.Format("Error inserting/updating contacts.\r\n{0}.\r\n", ex);
                update_Progress(textLine, configDirName, logOneIn);
            }


            return success;
        }
        private void DeleteTemporaryContacts(XDocument xdoc, string sqlConnectionString)
        {
            int logOneIn = Convert.ToInt32(xdoc.Root.Element("LogOneIn").Value);
            string configDirName = xdoc.Root.Element("SoftwareHomeFolder").Value;
            string textLine = null;

            //
            //Try using a stored procedure
            //
            try
            {
                using (var connection = new SqlConnection(sqlConnectionString))
                using (var command = new SqlCommand("DeleteTemporaryContacts", connection)
                {
                    CommandType = CommandType.StoredProcedure
                })
                {
                    connection.Open();
                    command.ExecuteNonQuery();

                }
            }
            catch (Exception ex)
            {
                textLine = string.Format("Error deleting temporary contacts.\r\n{0}.\r\n", ex);
                update_Progress(textLine, configDirName, logOneIn);
            }


        }
        public void update_Progress(string textLine, string configDirName, int logOneIn)
        {
            //Get the month and year from today's date
            DateTime now = DateTime.Now;
            string year = now.Year.ToString();
            string month = now.Month.ToString().PadLeft(2, '0');
            string homeFolder = configDirName;

            using (StreamWriter sw = new StreamWriter(homeFolder + "Config\\" + "PRtoWG-Log" + year + month + ".txt", true))
            {
                textLine = string.Format(textLine + " - {0}", now);
                sw.WriteLine(textLine);

            }

        }

        public void ArchiveCompletedPayrollFile(XDocument xdoc, FileInfo completedPayrollFile)
        {
            DateTime now = DateTime.Now;
            string nowString = now.ToString("yyyyMMddHHmmssfff");

            string destFileName = completedPayrollFile.FullName.Replace("Outputs", "PE-ArchivedOutputs").Replace(".xml", "_" + nowString + ".xml");
            //destFileName = destFileName.Replace(".xml", "_" + nowString + ".xml");
            

            File.Move(completedPayrollFile.FullName, destFileName);
        }
        
        
        public XmlDocument RunReport(string rptRef, string prm1, string val1, string prm2, string val2, string prm3, string val3,
                                 string prm4, string val4, string prm5, string val5, string prm6, string val6)
        {
            string url = null;
            if (prm2 == null)
            {
                url = prm1 + "=" + val1;

            }
            else if (prm3 == null)
            {
                url = prm1 + "=" + val1 + "&" + prm2 + "=" + val2;

            }
            else if (prm4 == null)
            {
                url = prm1 + "=" + val1 + "&" + prm2 + "=" + val2
                            + "&" + prm3 + "=" + val3;

            }
            else if (prm5 == null)
            {
                url = prm1 + "=" + val1 + "&" + prm2 + "=" + val2
                            + "&" + prm3 + "=" + val3 + "&" + prm4 + "=" + val4;

            }
            else if (prm6 == null)
            {
                url = prm1 + "=" + val1 + "&" + prm2 + "=" + val2
                            + "&" + prm3 + "=" + val3 + "&" + prm4 + "=" + val4
                            + "&" + prm5 + "=" + val5;

            }
            else
            {
                url = prm1 + "=" + val1 + "&" + prm2 + "=" + val2
                           + "&" + prm3 + "=" + val3 + "&" + prm4 + "=" + val4
                           + "&" + prm5 + "=" + val5 + "&" + prm6 + "=" + val6;
            }
            XmlDocument xmlReport = null;
            try
            {
                //Mark this is the full url = "https://api.test.payrun.io/Report/PayescapeEmployeePeriod/run?EmployerKey=1104&TaxYear=2018&AccPeriodStart=2018/01/01&AccPeriodEnd=2019/03/08&TaxPeriod=49&PayScheduleKey=Weekly"
                var apiHelper = ApiHelper();
                //string testurl = "EmployerKey=1958&TaxYear=2019&AccPeriodStart=2019-04-06&AccPeriodEnd=2020-04-05&TaxPeriod=27&PayScheduleKey=Weekly";
                //xmlReport = apiHelper.GetRawXml("/Report/" + rptRef + "/run?" + testurl);
                xmlReport = apiHelper.GetRawXml("/Report/" + rptRef + "/run?" + url);

            }
            catch (Exception ex)
            {
                //MessageBox.Show("Error running a report.\r\n" + ex);
            }
            return xmlReport;
        }
        private RestApiHelper ApiHelper()
        {
            string consumerKey = "9MaHMtmTUC6iMgymPl94g";                             //Original developer key : "m5lsJMpBnkaJw086zwDw"     "1UH6t3ikiWbdxTNT2Dg"
            string consumerSecret = "44sem3aVCUCxjaFmnolPQhPii7rQQwEyqgTnSJB655Q";   //Original developer secret : "GHM6x3xLEWujpLC5sGXKQ3r2j14RGI0eoLbab8w415Q"     "jKUX3lrQUe4KhEiox6IZw8CXnWUdAkyTl1kthR8ayQ"
            string url = "https://api.test.payrun.io";
            RestApiHelper apiHelper = new PayRunIO.CSharp.SDK.RestApiHelper(
                    new PayRunIO.OAuth1.OAuthSignatureGenerator(),
                    consumerKey,
                    consumerSecret,
                    url,
                    "application/xml",
                    "application/xml");
            return apiHelper;
        }
        public void ArchiveDirectory(XDocument xdoc, string directory, string originalDirName, string archiveDirName)
        {
            int logOneIn = Convert.ToInt32(xdoc.Root.Element("LogOneIn").Value);
            string configDirName = xdoc.Root.Element("SoftwareHomeFolder").Value;
            string textLine = null;

            try
            {
                DateTime now = DateTime.Now;

                int x = directory.LastIndexOf("\\");
                string coNo = directory.Substring(x + 1, 4);
                Directory.CreateDirectory(directory.Replace(originalDirName, archiveDirName));
                DirectoryInfo dirInfo = new DirectoryInfo(directory);
                FileInfo[] files = dirInfo.GetFiles();
                foreach (FileInfo file in files)
                {
                    string destFileName = file.FullName.Replace(originalDirName, archiveDirName);
                    destFileName = destFileName.Replace(".xml", "_" + now.ToString("yyyyMMddHHmmssfff") + ".xml");
                    destFileName = destFileName.Replace(".csv", "_" + now.ToString("yyyyMMddHHmmssfff") + ".csv");
                    File.Move(file.FullName, destFileName);

                }

                Directory.Delete(directory);
            }
            catch (Exception ex)
            {
                textLine = string.Format("Error archiving the Outputs directory, {0}.\r\n{1}.\r\n", directory, ex);
                update_Progress(textLine, configDirName, logOneIn);
            }

        }
        public XmlDocument GetP32Report(RPParameters rpParameters)
        {
            //Run the next period report to get the next pay period.
            string rptRef = "P32";
            string parameter1 = "EmployerKey";
            string parameter2 = "TaxYear";
            
            //Get the P32Sum report
            XmlDocument xmlReport = RunReport(rptRef, parameter1, rpParameters.ErRef, parameter2, rpParameters.TaxYear.ToString(),
                                              null, null, null, null, null, null, null, null);

            
            return xmlReport;
        }

        public void ProcessYtdReport(XDocument xdoc, XmlDocument xmlYTDReport, RPParameters rpParameters)
        {
            List<RPEmployeeYtd> rpEmployeeYtdList = PrepareYTDCSV(xdoc, xmlYTDReport);
            CreateYTDCSV(xdoc, rpEmployeeYtdList, rpParameters);

        }
        private int GetTaxMonth(DateTime thisDate)
        {
            int taxMonth = thisDate.Month - 3;
            if (thisDate.Day < 6)
            {
                taxMonth = taxMonth - 1;
            }
            if (taxMonth < 1)
            {
                taxMonth = taxMonth + 12;
            }
            return taxMonth;
        }
        public string[] GetAListOfDirectories(XDocument xdoc, string source)
        {
            string path = xdoc.Root.Element("DataHomeFolder").Value + source;
            string[] directories = Directory.GetDirectories(path);

            return directories;
        }
        //public FileInfo[] GetAllCompletedPayrollFiles(XDocument xdoc)
        //{
        //    string path = xdoc.Root.Element("DataHomeFolder").Value + "Outputs";
        //    DirectoryInfo folder = new DirectoryInfo(path);
        //    FileInfo[] files = folder.GetFiles("*CompletedPayroll*.xml");
            
        //    return files;
        //}
        
        
        public RPParameters GetRPParameters(XmlDocument xmlReport)
        {
            //Now extract the necessary data and produce the required reports.

            RPParameters rpParameters = new RPParameters();

            var rootElement = XElement.Parse(xmlReport.InnerXml);
            var header = rootElement.Elements("Parameters").ToArray();
            var header1 = rootElement.Elements("Parameters");
            
            foreach (XmlElement parameter in xmlReport.GetElementsByTagName("Parameters"))
            {
                rpParameters.ErRef = GetElementByTagFromXml(parameter, "EmployerCode");
                rpParameters.TaxYear = GetIntElementByTagFromXml(parameter, "TaxYear");
                rpParameters.AccYearStart = Convert.ToDateTime(GetDateElementByTagFromXml(parameter, "AccountingYearStartDate"));
                rpParameters.AccYearEnd = Convert.ToDateTime(GetDateElementByTagFromXml(parameter, "AccountingYearEndDate"));
                rpParameters.TaxPeriod = GetIntElementByTagFromXml(parameter, "TaxPeriod");
                rpParameters.PaySchedule = GetElementByTagFromXml(parameter, "PaySchedule");
            }
            return rpParameters;
        }
        public RPEmployer GetRPEmployer(XmlDocument xmlReport)
        {
            RPEmployer rpEmployer = new RPEmployer();

            foreach (XmlElement employer in xmlReport.GetElementsByTagName("Employer"))
            {
                rpEmployer.Name = GetElementByTagFromXml(employer, "Name");
                rpEmployer.PayeRef = GetElementByTagFromXml(employer, "EmployerPayeRef");
                rpEmployer.P32Required = GetBooleanElementByTagFromXml(employer, "P32Required");
                rpEmployer.BankFileCode = "001";
                rpEmployer.PensionReportCode = "001";
                //Get the bank file code for a table on the database for now. It should be supplied by WebGlobe and then PR eventually.
            }
            return rpEmployer;
        }
        public Tuple<List<RPEmployeeYtd>, RPParameters> PrepareYTDReport(XDocument xdoc, FileInfo file)
        {
            XmlDocument xmlYTDReport = new XmlDocument();
            xmlYTDReport.Load(file.FullName);

            //Now extract the necessary data and produce the required reports.

            RPParameters rpParameters = GetRPParameters(xmlYTDReport);
            List<RPEmployeeYtd> rpEmployeeYtdList = PrepareYTDCSV(xdoc, xmlYTDReport);

            return new Tuple<List<RPEmployeeYtd>, RPParameters>(rpEmployeeYtdList, rpParameters);
        }
        private List<RPEmployeeYtd> PrepareYTDCSV(XDocument xdoc, XmlDocument xmlReport)
        {
            string outgoingFolder = xdoc.Root.Element("DataHomeFolder").Value + "PE-Outgoing";
            List<RPEmployeeYtd> rpEmployeeYtdList = new List<RPEmployeeYtd>();

            foreach (XmlElement employee in xmlReport.GetElementsByTagName("Employee"))
            {
                bool include = false;
                if (GetElementByTagFromXml(employee, "PayRunDate") != "No Pay Run Data Found")
                {
                    //If the employee is a leaver before the start date then don't include.
                    string leaver = GetElementByTagFromXml(employee, "Leaver");
                    DateTime leavingDate = new DateTime();
                    if (GetElementByTagFromXml(employee, "LeavingDate") != "")
                    {
                        leavingDate = DateTime.ParseExact(GetElementByTagFromXml(employee, "LeavingDate"), "yyyy-MM-dd", CultureInfo.InvariantCulture);

                    }
                    DateTime periodStartDate = DateTime.ParseExact(GetElementByTagFromXml(employee, "ThisPeriodStartDate"), "yyyy-MM-dd", CultureInfo.InvariantCulture);
                    //It seems they want to include leaver in the YTD csv file. I think this might change!
                    include = true;
                    //if (leaver.StartsWith("N"))
                    //{
                    //    include = true;
                    //}
                    //else if (leavingDate >= periodStartDate)
                    //{
                    //    include = true;
                    //}
                    
                }

                if (include)
                {
                    RPEmployeeYtd rpEmployeeYtd = new RPEmployeeYtd();

                    rpEmployeeYtd.ThisPeriodStartDate = Convert.ToDateTime(GetDateElementByTagFromXml(employee, "ThisPeriodStartDate"));
                    rpEmployeeYtd.LastPaymentDate = Convert.ToDateTime(GetDateElementByTagFromXml(employee, "LastPaymentDate"));
                    rpEmployeeYtd.EeRef = GetElementByTagFromXml(employee, "EeRef");
                    rpEmployeeYtd.Branch = GetElementByTagFromXml(employee, "Branch");
                    rpEmployeeYtd.CostCentre = GetElementByTagFromXml(employee, "CostCentre");
                    rpEmployeeYtd.Department = GetElementByTagFromXml(employee, "Department");
                    rpEmployeeYtd.LeavingDate = GetDateElementByTagFromXml(employee, "LeavingDate");
                    rpEmployeeYtd.Leaver = GetBooleanElementByTagFromXml(employee, "Leaver");
                    rpEmployeeYtd.TaxPrevEmployment = GetDecimalElementByTagFromXml(employee, "TaxPrevEmployment");
                    rpEmployeeYtd.TaxablePayPrevEmployment = GetDecimalElementByTagFromXml(employee, "TaxablePayPrevEmployment");
                    rpEmployeeYtd.TaxThisEmployment = GetDecimalElementByTagFromXml(employee, "TaxThisEmployment");
                    rpEmployeeYtd.TaxablePayThisEmployment = GetDecimalElementByTagFromXml(employee, "TaxablePayThisEmployment");
                    rpEmployeeYtd.GrossedUp = GetDecimalElementByTagFromXml(employee, "GrossedUp");
                    rpEmployeeYtd.GrossedUpTax = GetDecimalElementByTagFromXml(employee, "GrossedUpTax");
                    rpEmployeeYtd.NetPayYTD = GetDecimalElementByTagFromXml(employee, "NetPayYTD");
                    rpEmployeeYtd.GrossPayYTD = GetDecimalElementByTagFromXml(employee, "GrossPayYTD");
                    rpEmployeeYtd.BenefitInKindYTD = GetDecimalElementByTagFromXml(employee, "BenefitInKindYTD");
                    rpEmployeeYtd.SuperannuationYTD = GetDecimalElementByTagFromXml(employee, "Superannuation");
                    rpEmployeeYtd.HolidayPayYTD = GetDecimalElementByTagFromXml(employee, "HolidayPayYTD");
                    rpEmployeeYtd.PensionablePayYtd = 0;
                    rpEmployeeYtd.EePensionYtd = 0;
                    rpEmployeeYtd.ErPensionYtd = 0;
                    List<RPPensionYtd> rpPensionsYtd = new List<RPPensionYtd>();
                    foreach (XmlElement pension in employee.GetElementsByTagName("Pension"))
                    {
                        RPPensionYtd rpPensionYtd = new RPPensionYtd();
                        rpPensionYtd.Key = Convert.ToInt32(pension.GetAttribute("Key"));
                        rpPensionYtd.Code = GetElementByTagFromXml(pension,"Code");
                        rpPensionYtd.SchemeName = GetElementByTagFromXml(pension, "SchemeName");
                        rpPensionYtd.PensionablePayYtd = GetDecimalElementByTagFromXml(pension, "PensionablePayYtd");
                        rpPensionYtd.EePensionYtd = GetDecimalElementByTagFromXml(pension, "EePensionYtd");
                        rpPensionYtd.ErPensionYtd = GetDecimalElementByTagFromXml(pension, "ErPensionYtd");

                        rpEmployeeYtd.PensionablePayYtd = rpEmployeeYtd.PensionablePayYtd + rpPensionYtd.PensionablePayYtd;
                        rpEmployeeYtd.EePensionYtd = rpEmployeeYtd.EePensionYtd + rpPensionYtd.EePensionYtd;
                        rpEmployeeYtd.ErPensionYtd = rpEmployeeYtd.ErPensionYtd + rpPensionYtd.ErPensionYtd;

                        rpPensionsYtd.Add(rpPensionYtd);
                    }
                    rpEmployeeYtd.Pensions = rpPensionsYtd;

                    rpEmployeeYtd.AeoYTD = GetDecimalElementByTagFromXml(employee, "AeoYTD");
                    rpEmployeeYtd.StudentLoanStartDate = GetDateElementByTagFromXml(employee, "StudentLoanStartDate");
                    rpEmployeeYtd.StudentLoanEndDate = GetDateElementByTagFromXml(employee, "StudentLoanEndDate");
                    rpEmployeeYtd.StudentLoanDeductionsYTD = GetDecimalElementByTagFromXml(employee, "StudentLoanDeductionsYTD");
                    rpEmployeeYtd.PostgraduateLoanStartDate = GetDateElementByTagFromXml(employee, "PostgraduateLoanStartDate");
                    rpEmployeeYtd.PostgraduateLoanEndDate = GetDateElementByTagFromXml(employee, "PostgraduateLoanEndDate");
                    rpEmployeeYtd.PostgraduateLoanDeductionsYTD = GetDecimalElementByTagFromXml(employee, "PostgraduateLoanDeductionsYTD");
                    
                    foreach (XmlElement nicYtd in employee.GetElementsByTagName("NicYtd"))
                    {
                        RPNicYtd rpNicYtd = new RPNicYtd();
                        rpNicYtd.NILetter = nicYtd.GetAttribute("NiLetter");
                        rpNicYtd.NiableYtd = GetDecimalElementByTagFromXml(nicYtd, "NiableYtd");
                        rpNicYtd.EarningsToLEL = GetDecimalElementByTagFromXml(nicYtd, "EarningsToLEL");
                        rpNicYtd.EarningsToSET = GetDecimalElementByTagFromXml(nicYtd, "EarningsToSET");
                        rpNicYtd.EarningsToPET = GetDecimalElementByTagFromXml(nicYtd, "EarningsToPET");
                        rpNicYtd.EarningsToUST = GetDecimalElementByTagFromXml(nicYtd, "EarningsToUST");
                        rpNicYtd.EarningsToAUST = GetDecimalElementByTagFromXml(nicYtd, "EarningsToAUST");
                        rpNicYtd.EarningsToUEL = GetDecimalElementByTagFromXml(nicYtd, "EarningsToUEL");
                        rpNicYtd.EarningsAboveUEL = GetDecimalElementByTagFromXml(nicYtd, "EarningsAboveUEL");
                        rpNicYtd.EeContributionsPt1 = GetDecimalElementByTagFromXml(nicYtd, "EeContributionsPt1");
                        rpNicYtd.EeContributionsPt2 = GetDecimalElementByTagFromXml(nicYtd, "EeContributionsPt2");
                        rpNicYtd.ErContributions = GetDecimalElementByTagFromXml(nicYtd, "ErContributions");
                        rpNicYtd.EeRebate = GetDecimalElementByTagFromXml(nicYtd, "EeRebate");
                        rpNicYtd.ErRebate = GetDecimalElementByTagFromXml(nicYtd, "ErRebate");
                        rpNicYtd.EeReduction = GetDecimalElementByTagFromXml(nicYtd, "EeReduction");
                        rpNicYtd.ErReduction = GetDecimalElementByTagFromXml(nicYtd, "ErReduction");

                        rpEmployeeYtd.NicYtd = rpNicYtd;
                    }
                    foreach (XmlElement nicAccountingPeriod in employee.GetElementsByTagName("NicAccountingPeriod"))
                    {
                        RPNicAccountingPeriod rpNicAccountingPeriod = new RPNicAccountingPeriod();
                        rpNicAccountingPeriod.NILetter = nicAccountingPeriod.GetAttribute("NiLetter");
                        rpNicAccountingPeriod.NiableYtd = GetDecimalElementByTagFromXml(nicAccountingPeriod, "NiableYtd");
                        rpNicAccountingPeriod.EarningsToLEL = GetDecimalElementByTagFromXml(nicAccountingPeriod, "EarningsToLEL");
                        rpNicAccountingPeriod.EarningsToSET = GetDecimalElementByTagFromXml(nicAccountingPeriod, "EarningsToSET");
                        rpNicAccountingPeriod.EarningsToPET = GetDecimalElementByTagFromXml(nicAccountingPeriod, "EarningsToPET");
                        rpNicAccountingPeriod.EarningsToUST = GetDecimalElementByTagFromXml(nicAccountingPeriod, "EarningsToUST");
                        rpNicAccountingPeriod.EarningsToAUST = GetDecimalElementByTagFromXml(nicAccountingPeriod, "EarningsToAUST");
                        rpNicAccountingPeriod.EarningsToUEL = GetDecimalElementByTagFromXml(nicAccountingPeriod, "EarningsToUEL");
                        rpNicAccountingPeriod.EarningsAboveUEL = GetDecimalElementByTagFromXml(nicAccountingPeriod, "EarningsAboveUEL");
                        rpNicAccountingPeriod.EeContributionsPt1 = GetDecimalElementByTagFromXml(nicAccountingPeriod, "EeContributionsPt1");
                        rpNicAccountingPeriod.EeContributionsPt2 = GetDecimalElementByTagFromXml(nicAccountingPeriod, "EeContributionsPt2");
                        rpNicAccountingPeriod.ErContributions = GetDecimalElementByTagFromXml(nicAccountingPeriod, "ErContributions");
                        rpNicAccountingPeriod.EeRebate = GetDecimalElementByTagFromXml(nicAccountingPeriod, "EeRebate");
                        rpNicAccountingPeriod.ErRebate = GetDecimalElementByTagFromXml(nicAccountingPeriod, "ErRebate");
                        rpNicAccountingPeriod.EeReduction = GetDecimalElementByTagFromXml(nicAccountingPeriod, "EeReduction");
                        rpNicAccountingPeriod.ErReduction = GetDecimalElementByTagFromXml(nicAccountingPeriod, "ErReduction");

                        rpNicAccountingPeriod.ErReduction = GetDecimalElementByTagFromXml(nicAccountingPeriod, "ErReduction");

                        rpEmployeeYtd.NicAccountingPeriod = rpNicAccountingPeriod;
                    }

                    rpEmployeeYtd.TaxCode = GetElementByTagFromXml(employee, "TaxCode");
                    rpEmployeeYtd.Week1Month1 = GetBooleanElementByTagFromXml(employee, "Week1Month1");
                    rpEmployeeYtd.WeekNumber = GetIntElementByTagFromXml(employee, "WeekNumber");
                    rpEmployeeYtd.MonthNumber = GetIntElementByTagFromXml(employee, "MonthNumber");
                    rpEmployeeYtd.PeriodNumber = GetIntElementByTagFromXml(employee, "PeriodNumber");
                    rpEmployeeYtd.EeNiPaidByErAccountsAmount = GetDecimalElementByTagFromXml(employee, "EeNiPaidByErAccountsAmount");
                    rpEmployeeYtd.EeNiPaidByErAccountsUnits = GetDecimalElementByTagFromXml(employee, "EeNiPaidByErAccountsUnits");
                    rpEmployeeYtd.EeGuTaxPaidByErAccountsAmount = GetDecimalElementByTagFromXml(employee, "EeGuTaxPaidByErAccountsAmount");
                    rpEmployeeYtd.EeGuTaxPaidByErAccountsUnits = GetDecimalElementByTagFromXml(employee, "EeGuTaxPaidByErAccountsUnit");
                    rpEmployeeYtd.EeNiLERtoUERAccountsAmount = GetDecimalElementByTagFromXml(employee, "EeNiLERtoUERAccountsAmount");
                    rpEmployeeYtd.EeNiLERtoUERAccountsUnits = GetDecimalElementByTagFromXml(employee, "EeNiLERtoUERAccountsUnit");
                    rpEmployeeYtd.ErNiAccountsAmount = GetDecimalElementByTagFromXml(employee, "ErNiAccountAmount");
                    rpEmployeeYtd.ErNiAccountsUnits = GetDecimalElementByTagFromXml(employee, "ErNiAccountUnit");
                    rpEmployeeYtd.EeNiLERtoUERPayeAmount = GetDecimalElementByTagFromXml(employee, "EeNiLERtoUERPayeAmount");
                    rpEmployeeYtd.EeNiLERtoUERPayeUnits = GetDecimalElementByTagFromXml(employee, "EeNiLERtoUERPayeUnit");
                    rpEmployeeYtd.EeNiPaidByErPayeAmount = GetDecimalElementByTagFromXml(employee, "EeNiPaidByErPayeAmount");
                    rpEmployeeYtd.EeNiPaidByErPayeUnits = GetDecimalElementByTagFromXml(employee, "EeNiPaidByErPayeUnits");
                    rpEmployeeYtd.EeGuTaxPaidByErPayeAmount = GetDecimalElementByTagFromXml(employee, "EeGuTaxPaidByErPayeAmount");
                    rpEmployeeYtd.EeGuTaxPaidByErPayeUnits = GetDecimalElementByTagFromXml(employee, "EeGuTaxPaidByErPayeUnit");
                    rpEmployeeYtd.ErNiPayeAmount = GetDecimalElementByTagFromXml(employee, "ErNiPayeAmount");
                    rpEmployeeYtd.ErNiPayeUnits = GetDecimalElementByTagFromXml(employee, "ErNiPayeUnit");

                    //Find the pension pay codes
                    rpEmployeeYtd.PensionPreTaxEeAccounts = 0;
                    rpEmployeeYtd.PensionPreTaxEePaye = 0;
                    rpEmployeeYtd.PensionPostTaxEeAccounts = 0;
                    rpEmployeeYtd.PensionPostTaxEePaye = 0;
                    foreach (XmlElement payCodes in employee.GetElementsByTagName("PayCodes"))
                    {
                        foreach (XmlElement payCode in payCodes.GetElementsByTagName("PayCode"))
                        {
                            string pensionCode = GetElementByTagFromXml(payCode, "Code");
                            if (pensionCode.StartsWith("PENSION"))
                            {
                                if(pensionCode == "PENSIONRAS" || pensionCode == "PENSIONTAXEX")
                                {
                                    rpEmployeeYtd.PensionPostTaxEeAccounts = rpEmployeeYtd.PensionPostTaxEeAccounts + GetDecimalElementByTagFromXml(payCode, "AccountsAmount");
                                    rpEmployeeYtd.PensionPostTaxEePaye = rpEmployeeYtd.PensionPostTaxEePaye + GetDecimalElementByTagFromXml(payCode, "PayeAmount");
                                }
                                else
                                {
                                    rpEmployeeYtd.PensionPreTaxEeAccounts = rpEmployeeYtd.PensionPreTaxEeAccounts + GetDecimalElementByTagFromXml(payCode, "AccountsAmount");
                                    rpEmployeeYtd.PensionPreTaxEePaye = rpEmployeeYtd.PensionPreTaxEePaye + GetDecimalElementByTagFromXml(payCode, "PayeAmount");
                                }
                            }
                            
                        }
                    }
                    rpEmployeeYtd.PensionPreTaxEeAccounts = rpEmployeeYtd.PensionPreTaxEeAccounts * -1;
                    rpEmployeeYtd.PensionPreTaxEePaye = rpEmployeeYtd.PensionPreTaxEePaye *-1;
                    rpEmployeeYtd.PensionPostTaxEeAccounts = rpEmployeeYtd.PensionPostTaxEeAccounts * -1;
                    rpEmployeeYtd.PensionPostTaxEePaye = rpEmployeeYtd.PensionPostTaxEePaye * -1;

                    //These next few fields get treated like pay codes. Use them if they are not zero.
                    //7 pay components EeNiPaidByEr, EeGuTaxPaidByEr, EeNiLERtoUER & ErNi
                    List<RPPayCode> rpPayCodeList = new List<RPPayCode>();

                    for (int i = 0; i < 7; i++)
                    {
                        RPPayCode rpPayCode = new RPPayCode();

                        rpPayCode.EeRef = rpEmployeeYtd.EeRef;
                        rpPayCode.Code = "0";

                        switch (i)
                        {
                            case 0:
                                rpPayCode.PayCode = "EeNIPdByEr";
                                rpPayCode.Description = "Ee NI Paid By Er";
                                rpPayCode.Type = "E";
                                rpPayCode.AccountsAmount = rpEmployeeYtd.EeNiPaidByErAccountsAmount;
                                rpPayCode.PayeAmount = rpEmployeeYtd.EeNiPaidByErPayeAmount;
                                rpPayCode.AccountsUnits = rpEmployeeYtd.EeNiPaidByErAccountsUnits;
                                rpPayCode.PayeUnits = rpEmployeeYtd.EeNiPaidByErPayeUnits;
                                break;
                            case 1:
                                rpPayCode.PayCode = "GUTax";
                                rpPayCode.Description = "Grossed up Tax";
                                rpPayCode.Type = "E";
                                rpPayCode.AccountsAmount = rpEmployeeYtd.EeGuTaxPaidByErAccountsAmount;//GetDecimalElementByTagFromXml(employee, "EeGuTaxPaidByErAccountsAmount");
                                rpPayCode.PayeAmount = rpEmployeeYtd.EeGuTaxPaidByErPayeAmount;//GetDecimalElementByTagFromXml(employee, "EeGuTaxPaidByErPayeAmount");
                                rpPayCode.AccountsUnits = rpEmployeeYtd.EeGuTaxPaidByErAccountsUnits;//GetDecimalElementByTagFromXml(employee, "EeGuTaxPaidByErAccountsUnit");
                                rpPayCode.PayeUnits = rpEmployeeYtd.EeGuTaxPaidByErPayeUnits;//GetDecimalElementByTagFromXml(employee, "EeGuTaxPaidByErPayeUnit");
                                break;
                            case 2:
                                rpPayCode.PayCode = "NIEeeLERtoUER";
                                rpPayCode.Description = "NIEeeLERtoUER-A";
                                rpPayCode.Type = "E";
                                rpPayCode.AccountsAmount = rpEmployeeYtd.EeNiLERtoUERAccountsAmount;//GetDecimalElementByTagFromXml(employee, "EeNiLERtoUERAccountsAmount");
                                rpPayCode.PayeAmount = rpEmployeeYtd.EeNiLERtoUERPayeAmount;//GetDecimalElementByTagFromXml(employee, "EeNiLERtoUERPayeAmount");
                                rpPayCode.AccountsUnits = rpEmployeeYtd.EeNiLERtoUERAccountsUnits;//GetDecimalElementByTagFromXml(employee, "EeNiLERtoUERAccountsUnit");
                                rpPayCode.PayeUnits = rpEmployeeYtd.EeNiLERtoUERPayeUnits;//GetDecimalElementByTagFromXml(employee, "EeNiLERtoUERPayeUnit");
                                break;
                            case 3:
                                rpPayCode.PayCode = "NIEr";
                                rpPayCode.Description = "NIEr-A";
                                rpPayCode.Type = "T";
                                rpPayCode.AccountsAmount = rpEmployeeYtd.ErNiAccountsAmount;//GetDecimalElementByTagFromXml(employee, "ErNiAccountAmount");
                                rpPayCode.PayeAmount = rpEmployeeYtd.ErNiPayeAmount;//GetDecimalElementByTagFromXml(employee, "ErNiPayeAmount");
                                rpPayCode.AccountsUnits = rpEmployeeYtd.ErNiAccountsUnits;//GetDecimalElementByTagFromXml(employee, "ErNiAccountUnit");
                                rpPayCode.PayeUnits = rpEmployeeYtd.ErNiPayeUnits;//GetDecimalElementByTagFromXml(employee, "ErNiPayeUnit");
                                break;
                            case 4:
                                rpPayCode.PayCode = "PenEr";
                                rpPayCode.Description = "PenEr";
                                rpPayCode.Type = "D";
                                rpPayCode.AccountsAmount = rpEmployeeYtd.ErPensionYtd;//GetDecimalElementByTagFromXml(employee, "ErPensionYTD");
                                rpPayCode.PayeAmount = rpEmployeeYtd.ErPensionYtd;//GetDecimalElementByTagFromXml(employee, "ErPensionYTD");
                                rpPayCode.AccountsUnits = 0;
                                rpPayCode.PayeUnits = 0;
                                break;
                            case 5:
                                rpPayCode.PayCode = "PenPreTaxEe";
                                rpPayCode.Description = "PenPreTaxEe";
                                rpPayCode.Type = "D";
                                rpPayCode.AccountsAmount = rpEmployeeYtd.PensionPreTaxEeAccounts;//GetDecimalElementByTagFromXml(employee, "EePensionYTD");
                                rpPayCode.PayeAmount = rpEmployeeYtd.PensionPreTaxEePaye;//GetDecimalElementByTagFromXml(employee, "EePensionYTD");
                                rpPayCode.AccountsUnits = 0;
                                rpPayCode.PayeUnits = 0;
                                break;
                            default:
                                rpPayCode.PayCode = "PenPostTaxEe";
                                rpPayCode.Description = "PenPostTaxEe";
                                rpPayCode.Type = "D";
                                rpPayCode.AccountsAmount = rpEmployeeYtd.PensionPostTaxEeAccounts;//GetDecimalElementByTagFromXml(employee, "EePensionYTD");
                                rpPayCode.PayeAmount = rpEmployeeYtd.PensionPostTaxEePaye;//GetDecimalElementByTagFromXml(employee, "EePensionYTD");
                                rpPayCode.AccountsUnits = 0;
                                rpPayCode.PayeUnits = 0;
                                break;
                        }

                        //
                        //Check if any of the values are not zero. If so write the first employee record
                        //
                        bool allZeros = false;
                        if (rpPayCode.AccountsAmount == 0 && rpPayCode.AccountsUnits == 0 &&
                            rpPayCode.PayeUnits == 0 && rpPayCode.PayeUnits == 0)
                        {
                            allZeros = true;

                        }
                        if (!allZeros)
                        {
                            //Add employee record to the list
                            rpPayCodeList.Add(rpPayCode);
                            //rpEmployeeYtd.PayCodes.Add(rpPayCode);
                        }
                    }
                    //Add in the pension schemes
                    foreach(RPPensionYtd rpPensionYtd in rpEmployeeYtd.Pensions)
                    {
                        //Ee pension
                        RPPayCode rpPayCode = new RPPayCode();

                        rpPayCode.EeRef = rpEmployeeYtd.EeRef;
                        rpPayCode.Code = "0";
                        rpPayCode.PayCode = rpPensionYtd.Code + "-" + rpPensionYtd.SchemeName + "-Ee";
                        rpPayCode.Description = rpPayCode.PayCode;
                        rpPayCode.Type = "P";
                        rpPayCode.AccountsAmount = rpPensionYtd.EePensionYtd;
                        rpPayCode.PayeAmount = rpPensionYtd.EePensionYtd;
                        rpPayCode.AccountsUnits = 0;
                        rpPayCode.PayeUnits = 0;

                        rpPayCodeList.Add(rpPayCode);

                        //Er pension
                        rpPayCode = new RPPayCode();

                        rpPayCode.EeRef = rpEmployeeYtd.EeRef;
                        rpPayCode.Code = "0";
                        rpPayCode.PayCode = rpPensionYtd.Code + "-" + rpPensionYtd.SchemeName + "-Er";
                        rpPayCode.Description = rpPayCode.PayCode;
                        rpPayCode.Type = "P";
                        rpPayCode.AccountsAmount = rpPensionYtd.ErPensionYtd;
                        rpPayCode.PayeAmount = rpPensionYtd.ErPensionYtd;
                        rpPayCode.AccountsUnits = 0;
                        rpPayCode.PayeUnits = 0;

                        rpPayCodeList.Add(rpPayCode);

                        //Pensionable pay
                        rpPayCode = new RPPayCode();

                        rpPayCode.EeRef = rpEmployeeYtd.EeRef;
                        rpPayCode.Code = "0";
                        rpPayCode.PayCode = rpPensionYtd.Code + "-" + rpPensionYtd.SchemeName + "-Pay";
                        rpPayCode.Description = rpPayCode.PayCode;
                        rpPayCode.Type = "P";
                        rpPayCode.AccountsAmount = rpPensionYtd.PensionablePayYtd;
                        rpPayCode.PayeAmount = rpPensionYtd.PensionablePayYtd;
                        rpPayCode.AccountsUnits = 0;
                        rpPayCode.PayeUnits = 0;

                        rpPayCodeList.Add(rpPayCode);
                    }

                    foreach (XmlElement payCodes in employee.GetElementsByTagName("PayCodes"))
                    {
                        foreach (XmlElement payCode in payCodes.GetElementsByTagName("PayCode"))
                        {
                            RPPayCode rpPayCode = new RPPayCode();

                            rpPayCode.EeRef = rpEmployeeYtd.EeRef;
                            rpPayCode.Code = GetElementByTagFromXml(payCode, "Code");
                            rpPayCode.PayCode = GetElementByTagFromXml(payCode, "Code");
                            rpPayCode.Description = GetElementByTagFromXml(payCode, "Description");
                            bool isPayCode = GetBooleanElementByTagFromXml(payCode, "IsPayCode");
                            if (isPayCode)
                            {
                                rpPayCode.Type = "E";
                            }
                            else
                            {
                                rpPayCode.Type = "D";
                            }

                            rpPayCode.AccountsAmount = GetDecimalElementByTagFromXml(payCode, "AccountsAmount");
                            rpPayCode.PayeAmount = GetDecimalElementByTagFromXml(payCode, "PayeAmount");
                            rpPayCode.AccountsUnits = GetDecimalElementByTagFromXml(payCode, "AccountsUnits");
                            rpPayCode.PayeUnits = GetDecimalElementByTagFromXml(payCode, "PayeUnits");

                            //
                            //Check if any of the values are not zero. If so write the first employee record
                            //
                            bool allZeros = false;
                            if (rpPayCode.AccountsAmount == 0 && rpPayCode.AccountsUnits == 0 &&
                                rpPayCode.PayeAmount == 0 && rpPayCode.PayeUnits == 0)
                            {
                                allZeros = true;

                            }
                            if (!allZeros)
                            {
                                //I don't require TAX, NI or PENSION
                                if (rpPayCode.Code != "TAX" && rpPayCode.Code != "NI" && !rpPayCode.Code.StartsWith("PENSION"))
                                {
                                    if (rpPayCode.Type == "D")
                                    {
                                        //Deduction so multiply by -1
                                        rpPayCode.AccountsAmount = rpPayCode.AccountsAmount * -1;
                                        //rpPayCode.AccountsUnits = rpPayCode.AccountsUnits * -1;
                                        rpPayCode.PayeAmount = rpPayCode.PayeAmount * -1;
                                        //rpPayCode.PayeUnits = rpPayCode.PayeUnits * -1;

                                    }
                                    if (rpPayCode.Code == "UNPDM")
                                    {
                                        //Change UNPDM back to UNPD£. WG uses UNPD£ PR doesn't like symbols like £ in pay codes.
                                        rpPayCode.Code = "";// "UNPD£";
                                        rpPayCode.PayCode = "UNPD£";
                                    }
                                    else
                                    {
                                        rpPayCode.Code = "";
                                    }
                                    //Add to employee record
                                    rpPayCodeList.Add(rpPayCode);
                                    //rpEmployeeYtd.PayCodes.Add(rpPayCode);
                                }



                            }

                        }
                        rpEmployeeYtd.PayCodes = rpPayCodeList;
                    }
                    rpEmployeeYtdList.Add(rpEmployeeYtd);
                }

            }
            //Sort the list of employees into EeRef sequence before returning them.
            rpEmployeeYtdList.Sort(delegate (RPEmployeeYtd x, RPEmployeeYtd y)
            {
                if (x.EeRef == null && y.EeRef == null) return 0;
                else if (x.EeRef == null) return -1;
                else if (y.EeRef == null) return 1;
                else return x.EeRef.CompareTo(y.EeRef);
            });

            return rpEmployeeYtdList;
        }
        public void CreateYTDCSV(XDocument xdoc, List<RPEmployeeYtd> rpEmployeeYtdList, RPParameters rpParameters)
        {
            string outgoingFolder = xdoc.Root.Element("DataHomeFolder").Value + "PE-Outgoing";

            string coNo = rpParameters.ErRef;
            //Create csv version and write it to the same folder.
            //string csvFileName = "V:\\Payescape\\PayRunIO\\WG\\" + coNo + "_YearToDates_" + DateTime.Now.ToString("yyyyMMddHHmmssfff") + ".csv";
            string csvFileName = outgoingFolder + "\\" + coNo + "\\" + coNo + "_YearToDates_" + DateTime.Now.ToString("yyyyMMddHHmmssfff") + ".csv";
            bool writeHeader = true;
            using (StreamWriter sw = new StreamWriter(csvFileName))
            {

                //Loop through each employee and write the csv file.
                string[] payYTDDetails = new string[45];


                foreach (RPEmployeeYtd rpEmployeeYtd in rpEmployeeYtdList)
                {
                    payYTDDetails[0] = rpEmployeeYtd.LastPaymentDate.ToString("dd/MM/yy", CultureInfo.InvariantCulture);
                    //I'm using the rpParameters from the "EmployeePeriod" report.
                    payYTDDetails[0] = rpParameters.PayRunDate.ToString("dd/MM/yy", CultureInfo.InvariantCulture);
                    payYTDDetails[1] = rpEmployeeYtd.EeRef;
                    if (rpEmployeeYtd.LeavingDate != null)
                    {
                        payYTDDetails[2] = Convert.ToDateTime(rpEmployeeYtd.LeavingDate).ToString("dd/MM/yy", CultureInfo.InvariantCulture);
                    }
                    else
                    {
                        payYTDDetails[2] = "";
                    }
                    if (rpEmployeeYtd.Leaver)
                    {
                        payYTDDetails[3] = "Y";
                    }
                    else
                    {
                        payYTDDetails[3] = "N";
                    }
                    payYTDDetails[4] = rpEmployeeYtd.TaxPrevEmployment.ToString();
                    payYTDDetails[5] = rpEmployeeYtd.TaxablePayPrevEmployment.ToString();
                    payYTDDetails[6] = rpEmployeeYtd.TaxThisEmployment.ToString();
                    payYTDDetails[7] = rpEmployeeYtd.TaxablePayThisEmployment.ToString();
                    payYTDDetails[8] = rpEmployeeYtd.GrossedUp.ToString();
                    payYTDDetails[9] = rpEmployeeYtd.GrossedUpTax.ToString();
                    payYTDDetails[10] = rpEmployeeYtd.NetPayYTD.ToString();
                    payYTDDetails[11] = rpEmployeeYtd.GrossPayYTD.ToString();
                    payYTDDetails[12] = rpEmployeeYtd.BenefitInKindYTD.ToString();
                    payYTDDetails[13] = rpEmployeeYtd.SuperannuationYTD.ToString();
                    payYTDDetails[14] = rpEmployeeYtd.HolidayPayYTD.ToString();
                    //Add the pensions from the list of pensions
                    decimal erPensionYtd = 0;
                    decimal eePensionYtd = 0;
                    foreach(RPPensionYtd pensionYtd in rpEmployeeYtd.Pensions)
                    {
                        erPensionYtd = erPensionYtd + pensionYtd.ErPensionYtd;
                        eePensionYtd = eePensionYtd + pensionYtd.EePensionYtd;
                    }
                    payYTDDetails[15] = erPensionYtd.ToString();
                    payYTDDetails[16] = eePensionYtd.ToString();
                    payYTDDetails[17] = rpEmployeeYtd.AeoYTD.ToString();
                    if (rpEmployeeYtd.StudentLoanStartDate != null)
                    {
                        payYTDDetails[18] = Convert.ToDateTime(rpEmployeeYtd.StudentLoanStartDate).ToString("dd/MM/yy", CultureInfo.InvariantCulture);
                    }
                    else
                    {
                        payYTDDetails[18] = "";
                    }
                    if (rpEmployeeYtd.StudentLoanEndDate != null)
                    {
                        payYTDDetails[19] = Convert.ToDateTime(rpEmployeeYtd.StudentLoanEndDate).ToString("dd/MM/yy", CultureInfo.InvariantCulture);
                    }
                    else
                    {
                        payYTDDetails[19] = "";
                    }
                    payYTDDetails[20] = rpEmployeeYtd.StudentLoanDeductionsYTD.ToString();
                    payYTDDetails[21] = rpEmployeeYtd.NicYtd.NILetter;
                    payYTDDetails[22] = rpEmployeeYtd.NicYtd.NiableYtd.ToString();
                    payYTDDetails[23] = rpEmployeeYtd.NicYtd.EarningsToLEL.ToString();
                    payYTDDetails[24] = rpEmployeeYtd.NicYtd.EarningsToSET.ToString();
                    payYTDDetails[25] = rpEmployeeYtd.NicYtd.EarningsToPET.ToString();
                    payYTDDetails[26] = rpEmployeeYtd.NicYtd.EarningsToUST.ToString();
                    payYTDDetails[27] = rpEmployeeYtd.NicYtd.EarningsToAUST.ToString();
                    payYTDDetails[28] = rpEmployeeYtd.NicYtd.EarningsToUEL.ToString();
                    payYTDDetails[29] = rpEmployeeYtd.NicYtd.EarningsAboveUEL.ToString();
                    payYTDDetails[30] = rpEmployeeYtd.NicYtd.EeContributionsPt1.ToString();
                    payYTDDetails[31] = rpEmployeeYtd.NicYtd.EeContributionsPt2.ToString();
                    payYTDDetails[32] = rpEmployeeYtd.NicYtd.ErContributions.ToString();
                    payYTDDetails[33] = rpEmployeeYtd.NicYtd.EeRebate.ToString();
                    payYTDDetails[34] = rpEmployeeYtd.NicYtd.ErRebate.ToString();
                    payYTDDetails[35] = rpEmployeeYtd.NicYtd.EeReduction.ToString();
                    payYTDDetails[36] = rpEmployeeYtd.TaxCode;
                    if (rpEmployeeYtd.Week1Month1)
                    {
                        payYTDDetails[37] = "Y";
                    }
                    else
                    {
                        payYTDDetails[37] = "N";
                    }
                    payYTDDetails[38] = rpEmployeeYtd.WeekNumber.ToString();
                    payYTDDetails[39] = rpEmployeeYtd.MonthNumber.ToString();
                    payYTDDetails[40] = rpEmployeeYtd.NicYtd.NiableYtd.ToString();
                    payYTDDetails[41] = rpEmployeeYtd.StudentLoanPlanType; //Student Loan Plan Type

                    if (rpEmployeeYtd.PostgraduateLoanStartDate != null)
                    {
                        payYTDDetails[42] = Convert.ToDateTime(rpEmployeeYtd.PostgraduateLoanStartDate).ToString("dd/MM/yy", CultureInfo.InvariantCulture); //Postgraduate Loan Start Date
                    }
                    else
                    {
                        payYTDDetails[42] = "";
                    }
                    if (rpEmployeeYtd.PostgraduateLoanEndDate != null)
                    {
                        payYTDDetails[43] = Convert.ToDateTime(rpEmployeeYtd.PostgraduateLoanEndDate).ToString("dd/MM/yy", CultureInfo.InvariantCulture); //Postgraduate Loan End Date
                    }
                    else
                    {
                        payYTDDetails[43] = "";
                    }

                    payYTDDetails[44] = rpEmployeeYtd.PostgraduateLoanDeductionsYTD.ToString()  ; //Postgraduate Loan Deducted

                    foreach (RPPayCode rpPayCode in rpEmployeeYtd.PayCodes)
                    {
                        //Don't use pay codes TAX, NI or any that begin with PENSION
                        if (rpPayCode.Code != "TAX" && rpPayCode.Code != "NI" && !rpPayCode.Code.StartsWith("PENSION"))
                        {
                            string[] payCodeDetails = new string[8];
                            payCodeDetails[0] = "";// rpPayCode.Code;
                            payCodeDetails[1] = rpPayCode.Type;
                            payCodeDetails[2] = rpPayCode.PayCode;
                            payCodeDetails[3] = rpPayCode.Description;
                            payCodeDetails[4] = rpPayCode.AccountsAmount.ToString();
                            payCodeDetails[5] = rpPayCode.PayeAmount.ToString();
                            payCodeDetails[6] = rpPayCode.AccountsUnits.ToString();
                            payCodeDetails[7] = rpPayCode.PayeUnits.ToString();

                            switch (payCodeDetails[2])
                            {
                                case "UNPDM":
                                    //Change UNPDM back to UNPD£. WG uses UNPD£ PR doesn't like symbols like £ in pay codes.
                                    payCodeDetails[2] = "UNPD£";
                                    break;
                                case "SLOAN":
                                    payCodeDetails[2] = "StudentLoan";
                                    payCodeDetails[3] = "StudentLoan";
                                    break;
                            }
                            
                            //Write employee record
                            WritePayYTDCSV(rpParameters, payYTDDetails, payCodeDetails, sw, writeHeader);
                            writeHeader = false;
                        }

                    }

                }

            }

        }
        private void WritePayYTDCSV(RPParameters rpParameters, string[] payYTDDetails, string[] payCodeDetails, StreamWriter sw, bool writeHeader)
        {
            string csvLine = null;
            if (writeHeader)
            {
                string csvHeader = "Co,RunDate,process,Batch,EeRef,LeaveDate,Leaver,Tax Previous Emt," +
                              "Taxable Pay Previous Emt,Tax This Emt,Taxable Pay This Emt,Grossed Up," +
                              "Grossed Up Tax,Net Pay,GrossYTD,Benefit in Kind,Superannuation," +
                              "Holiday Pay,ErPensionYTD,EePensionYTD,AEOYTD,StudentLoanStartDate," +
                              "StudentLoanEndDate,StudentLoanDeductions,NI Letter,Total," +
                              "Earnings To LEL,Earnings To SET,Earnings To PET,Earnings To UST," +
                              "Earnings To AUST,Earnings To UEL,Earnings Above UEL," +
                              "Ee Contributions Pt1,Ee Contributions Pt2,Er Contributions," +
                              "Ee Rebate,Er Rebate,Ee Reduction,PayCode,det,payCodeValue," +
                              "payCodeDesc,Acc Year Bal,PAYE Year Bal,Acc Year Units," +
                              "PAYE Year Units,Tax Code,Week1/Month 1,Week Number,Month Number," +
                              "NI Earnings YTD,Student Loan Plan Type,Postgraduate Loan Start Date," +
                              "Postgraduate Loan End Date,Postgraduate Loan Deducted";
                csvLine = csvHeader;
                sw.WriteLine(csvLine);
                csvLine = null;

            }
            string batch = null;
            switch (rpParameters.PaySchedule)
            {
                case "Monthly":
                    batch = "M";
                    break;
                case "TwoWeekly":
                    batch = "M";
                    break;
                case "FourWeekly":
                    batch = "M";
                    break;
                case "Yearly":
                    batch = "M";
                    break;
                default:
                    batch = "W";
                    break;
            }
            if (rpParameters.PaySchedule == "Monthly")
            {
                batch = "M";
            }

            string process = null;
            process = "20" + payYTDDetails[0].Substring(6, 2) + payYTDDetails[0].Substring(3, 2) + payYTDDetails[0].Substring(0, 2) + "01";
            csvLine = csvLine + "\"" + rpParameters.ErRef + "\"" + "," +                                     //Co. Number
                            "\"" + payYTDDetails[0] + "\"" + "," +                                            //Run Date / Last Payment Date
                            "\"" + process + "\"" + "," +                                                     //Process
                            "\"" + batch + "\"" + ",";                                                        //Batch


            //From payYTDDetails[1] (EeRef) to payYTDDetails[35] (EeReduction)
            for (int i = 1; i < 36; i++)
            {
                csvLine = csvLine + "\"" + payYTDDetails[i] + "\"" + ",";
            }
            //From payCodeDetails[0] (PayCode) to payCodeDetails[7] (PAYE Year Units)
            for (int i = 0; i < 8; i++)
            {
                csvLine = csvLine + "\"" + payCodeDetails[i] + "\"" + ",";
            }
            //From payYTDDetails[36] (TaxCode) to payYTDDetails[45] (Postgraduate Loan Deducted)
            for (int i = 36; i < 44; i++)
            {
                csvLine = csvLine + "\"" + payYTDDetails[i] + "\"" + ",";
            }

            csvLine = csvLine.TrimEnd(',');

            sw.WriteLine(csvLine);

        }
        

        public void CreateHistoryCSV(XDocument xdoc, RPParameters rpParameters, RPEmployer rpEmployer, List<RPEmployeePeriod> rpEmployeePeriodList)
        {
            string outgoingFolder = xdoc.Root.Element("DataHomeFolder").Value + "PE-Outgoing";
            string coNo = rpParameters.ErRef;
            //Write the whole xml file to the folder.
            //string xmlFileName = "V:\\Payescape\\PayRunIO\\WG\\" + coNo + "_PayHistory_" + DateTime.Now.ToString("yyyyMMddHHmmssfff") + ".xml";
            string dirName = outgoingFolder + "\\" + coNo + "\\";
            Directory.CreateDirectory(dirName);
            //Create csv version and write it to the same folder.
            string csvFileName = outgoingFolder + "\\" + coNo + "\\" + coNo + "_PayHistory_" + DateTime.Now.ToString("yyyyMMddHHmmssfff") + ".csv";
            bool writeHeader = true;
            using (StreamWriter sw = new StreamWriter(csvFileName))
            {

                //Loop through each employee and write the csv file.
                string[] payHistoryDetails = new string[54];

                foreach (RPEmployeePeriod rpEmployeePeriod in rpEmployeePeriodList)
                {
                    bool include = false;

                    //If the employee is a leaver before the start date then don't include.
                    if (!rpEmployeePeriod.Leaver)
                    {
                        include = true;
                    }
                    else if (rpEmployeePeriod.LeavingDate >= rpEmployeePeriod.PeriodStartDate)
                    {
                        include = true;
                    }

                    if (include)
                    {
                        payHistoryDetails[0] = rpEmployeePeriod.Reference;
                        payHistoryDetails[1] = rpEmployeePeriod.PayRunDate.ToString("dd/MM/yy", CultureInfo.InvariantCulture);
                        payHistoryDetails[2] = rpEmployeePeriod.PeriodStartDate.ToString("dd/MM/yy", CultureInfo.InvariantCulture);
                        payHistoryDetails[3] = rpEmployeePeriod.PeriodEndDate.ToString("dd/MM/yy", CultureInfo.InvariantCulture);
                        payHistoryDetails[4] = rpEmployeePeriod.PayrollYear.ToString();
                        payHistoryDetails[5] = rpEmployeePeriod.Gross.ToString();
                        payHistoryDetails[6] = rpEmployeePeriod.NetPayTP.ToString();
                        payHistoryDetails[7] = rpEmployeePeriod.DayHours.ToString();
                        if (rpEmployeePeriod.StudentLoanStartDate != null)
                        {
                            payHistoryDetails[8] = Convert.ToDateTime(rpEmployeePeriod.StudentLoanStartDate).ToString("dd/MM/yy", CultureInfo.InvariantCulture);
                        }
                        else
                        {
                            payHistoryDetails[8] = "";
                        }
                        if (rpEmployeePeriod.StudentLoanEndDate != null)
                        {
                            payHistoryDetails[9] = Convert.ToDateTime(rpEmployeePeriod.StudentLoanEndDate).ToString("dd/MM/yy", CultureInfo.InvariantCulture);
                        }
                        else
                        {
                            payHistoryDetails[9] = "";
                        }
                        //decimal studentLoan = rpEmployeePeriod.StudentLoan * -1;
                        //payHistoryDetails[10] = studentLoan.ToString();
                        payHistoryDetails[10] = (rpEmployeePeriod.StudentLoan * -1).ToString();
                        payHistoryDetails[11] = rpEmployeePeriod.NILetter;
                        payHistoryDetails[12] = rpEmployeePeriod.CalculationBasis;
                        payHistoryDetails[13] = rpEmployeePeriod.Total.ToString();
                        payHistoryDetails[14] = rpEmployeePeriod.EarningsToLEL.ToString();
                        payHistoryDetails[15] = rpEmployeePeriod.EarningsToSET.ToString();
                        payHistoryDetails[16] = rpEmployeePeriod.EarningsToPET.ToString();
                        payHistoryDetails[17] = rpEmployeePeriod.EarningsToUST.ToString(); ;
                        payHistoryDetails[18] = rpEmployeePeriod.EarningsToAUST.ToString();
                        payHistoryDetails[19] = rpEmployeePeriod.EarningsToUEL.ToString();
                        payHistoryDetails[20] = rpEmployeePeriod.EarningsAboveUEL.ToString();
                        payHistoryDetails[21] = rpEmployeePeriod.EeContributionsPt1.ToString();
                        payHistoryDetails[22] = rpEmployeePeriod.EeContributionsPt2.ToString();
                        payHistoryDetails[23] = rpEmployeePeriod.ErNICYTD.ToString();
                        payHistoryDetails[24] = rpEmployeePeriod.EeRebate.ToString();
                        payHistoryDetails[25] = rpEmployeePeriod.ErRebate.ToString();
                        payHistoryDetails[26] = rpEmployeePeriod.EeReduction.ToString();
                        if(rpEmployeePeriod.LeavingDate != null)
                        {
                            payHistoryDetails[27] = Convert.ToDateTime(rpEmployeePeriod.LeavingDate).ToString("dd/MM/yy", CultureInfo.InvariantCulture);
                        }
                        else
                        {
                            payHistoryDetails[27] = "";
                        }
                        
                        if (rpEmployeePeriod.Leaver)
                        {
                            payHistoryDetails[28] = "Y";
                        }
                        else
                        {
                            payHistoryDetails[28] = "N";
                        }

                        payHistoryDetails[29] = rpEmployeePeriod.TaxCode.ToString();
                        if (rpEmployeePeriod.Week1Month1)
                        {
                            payHistoryDetails[30] = "Y";
                            //Remove the " W1" from the tax code
                            payHistoryDetails[29] = payHistoryDetails[29].Replace(" W1","");
                        }
                        else
                        {
                            payHistoryDetails[30] = "N";
                        }
                        payHistoryDetails[31] = "0";   //rpEmployeePeriod.TaxCodeChangeTypeID;
                        payHistoryDetails[32] = "UNKNOWN"; //rpEmployeePeriod.TaxCodeChangeType;
                        payHistoryDetails[33] = rpEmployeePeriod.TaxPrev.ToString();
                        payHistoryDetails[34] = rpEmployeePeriod.TaxablePayPrevious.ToString();
                        payHistoryDetails[35] = rpEmployeePeriod.TaxThis.ToString();
                        payHistoryDetails[36] = rpEmployeePeriod.TaxablePayYTD.ToString();
                        payHistoryDetails[37] = rpEmployeePeriod.HolidayAccruedTd.ToString();

                        decimal erPensionYtd = 0;
                        decimal eePensionYtd = 0;
                        decimal erPensionTp = 0;
                        decimal eePensionTp = 0;
                        decimal erPensionPrd = 0;
                        decimal eePensionPrd = 0;
                        foreach(RPPensionPeriod pensionPeriod in rpEmployeePeriod.Pensions)
                        {
                            erPensionYtd = erPensionYtd + pensionPeriod.ErPensionYtd;
                            eePensionYtd = eePensionYtd + pensionPeriod.EePensionYtd;
                            erPensionTp = erPensionTp + pensionPeriod.ErPensionTaxPeriod;
                            eePensionTp = eePensionTp + pensionPeriod.EePensionTaxPeriod;
                            erPensionPrd = erPensionPrd + pensionPeriod.ErPensionPayRunDate;
                            eePensionPrd = eePensionPrd + pensionPeriod.EePensionPayRunDate;
                        }
                        payHistoryDetails[38] = erPensionYtd.ToString();
                        payHistoryDetails[39] = eePensionYtd.ToString();
                        payHistoryDetails[40] = erPensionTp.ToString();
                        payHistoryDetails[41] = eePensionTp.ToString();
                        payHistoryDetails[42] = erPensionPrd.ToString();
                        payHistoryDetails[43] = eePensionPrd.ToString();

                        payHistoryDetails[44] = rpEmployeePeriod.DirectorshipAppointmentDate.ToString("dd/MM/yy", CultureInfo.InvariantCulture);
                        if (rpEmployeePeriod.Director)
                        {
                            payHistoryDetails[45] = "Y";
                        }
                        else
                        {
                            payHistoryDetails[45] = "N";
                        }
                        if (payHistoryDetails[45] == "N")               //Director
                        {
                            //They're not a director
                            payHistoryDetails[12] = "E";                //They're not a director
                        }
                        else
                        {
                            //They're a director
                            if (payHistoryDetails[12] == "Cumulative")  //Calculation basis
                            {
                                payHistoryDetails[12] = "C";            //Calculation Basis is Cumulative and they're a director
                            }
                            else
                            {
                                payHistoryDetails[12] = "N";            //Calculation Basis is Week1Month1 and they're a director
                            }

                        }
                        payHistoryDetails[46] = rpEmployeePeriod.EeContributionsTaxPeriodPt1.ToString();
                        payHistoryDetails[47] = rpEmployeePeriod.EeContributionsTaxPeriodPt2.ToString();
                        payHistoryDetails[48] = rpEmployeePeriod.ErNICTP.ToString();
                        payHistoryDetails[49] = Convert.ToDateTime(rpEmployeePeriod.AEAssessment.AssessmentDate).ToString("dd/MM/yy", CultureInfo.InvariantCulture);
                        payHistoryDetails[50] = rpEmployeePeriod.AEAssessment.AssessmentCode;
                        payHistoryDetails[51] = rpEmployeePeriod.AEAssessment.AssessmentEvent;
                        payHistoryDetails[52] = rpEmployeePeriod.AEAssessment.TaxPeriod.ToString();
                        payHistoryDetails[53] = rpEmployeePeriod.AEAssessment.TaxYear.ToString();

                       //Er NI & Er Pension
                        for (int i = 0; i < 2; i++)
                        {
                            string[] payCodeDetails = new string[12];

                            switch (i)
                            {
                                case 0:
                                    payCodeDetails[1] = "NIEr-A";
                                    payCodeDetails[2] = "NIEr";
                                    payCodeDetails[3] = "T";
                                    payCodeDetails[6] = rpEmployeePeriod.ErNICTP.ToString();
                                    break;
                                case 1:
                                    payCodeDetails[1] = "PenEr";
                                    payCodeDetails[2] = "PenEr";
                                    payCodeDetails[3] = "M";
                                    payCodeDetails[6] = erPensionTp.ToString();
                                    break;
                                
                            }
                            payCodeDetails[0] = "0";
                            payCodeDetails[4] = "0";
                            payCodeDetails[5] = "0";
                            payCodeDetails[7] = "0";
                            payCodeDetails[8] = "0";
                            payCodeDetails[9] = "0";
                            payCodeDetails[10] = "0";
                            payCodeDetails[11] = "0";
                            
                                //
                                //Check if any of the values are not zero. If so write the first employee record
                                //
                            bool allZeros = false;
                            if (Convert.ToDecimal(payCodeDetails[5]) == 0 && Convert.ToDecimal(payCodeDetails[6]) == 0 &&
                                Convert.ToDecimal(payCodeDetails[7]) == 0 && Convert.ToDecimal(payCodeDetails[8]) == 0 &&
                                Convert.ToDecimal(payCodeDetails[9]) == 0)
                            {
                                allZeros = true;

                            }
                            if (!allZeros)
                            {
                                //Write employee record
                                WritePayHistoryCSV(rpParameters, payHistoryDetails, payCodeDetails, sw, writeHeader);
                                writeHeader = false;

                            }

                        }
                        foreach (RPAddition rpAddition in rpEmployeePeriod.Additions)
                        {
                            string[] payCodeDetails = new string[12];
                            payCodeDetails = new string[12];
                            payCodeDetails[0] = "";
                            payCodeDetails[1] = rpAddition.Description;
                            payCodeDetails[2] = rpAddition.Code.TrimStart(' ');
                            payCodeDetails[3] = "E"; //Earnings
                            payCodeDetails[4] = rpAddition.Rate.ToString();
                            payCodeDetails[5] = rpAddition.Units.ToString();
                            payCodeDetails[6] = rpAddition.AmountTP.ToString();
                            if (payCodeDetails[4] == "0.00")
                            {
                                payCodeDetails[4] = payCodeDetails[6];  // Make Rate equal to amount if rate is zero.
                            }
                            payCodeDetails[7] = rpAddition.AccountsYearBalance.ToString();
                            payCodeDetails[8] = rpAddition.AmountYTD.ToString();
                            payCodeDetails[9] = rpAddition.AccountsYearUnits.ToString();
                            payCodeDetails[10] = rpAddition.PayeYearUnits.ToString();
                            payCodeDetails[11] = rpAddition.PayrollAccrued.ToString();



                            //
                            //Check if any of the values are not zero. If so write the first employee record
                            //
                            bool allZeros = false;
                            if (Convert.ToDecimal(payCodeDetails[5]) == 0 && Convert.ToDecimal(payCodeDetails[6]) == 0 &&
                                Convert.ToDecimal(payCodeDetails[7]) == 0 && Convert.ToDecimal(payCodeDetails[8]) == 0 &&
                                Convert.ToDecimal(payCodeDetails[9]) == 0)
                            {
                                allZeros = true;

                            }
                            if (!allZeros)
                            {
                                //Write employee record
                                WritePayHistoryCSV(rpParameters, payHistoryDetails, payCodeDetails, sw, writeHeader);
                                writeHeader = false;

                            }


                        }
                        decimal penPreAmount = 0, penPostAmount = 0;
                        bool wait = false;
                        foreach (RPDeduction rpDeduction in rpEmployeePeriod.Deductions)
                        {
                            string[] payCodeDetails = new string[12];
                            payCodeDetails = new string[12];
                            payCodeDetails[0] = "";
                            if(rpDeduction.Code == "SLOAN")
                            {
                                payCodeDetails[1] = "StudentLoan";
                                payCodeDetails[2] = "StudentLoan";
                            }
                            else
                            {
                                payCodeDetails[1] = rpDeduction.Description;
                                payCodeDetails[2] = rpDeduction.Code.TrimStart(' ');
                            }
                            
                            payCodeDetails[3] = "D"; //Earnings
                            payCodeDetails[4] = rpDeduction.Rate.ToString();
                            payCodeDetails[5] = rpDeduction.Units.ToString();
                            payCodeDetails[6] = rpDeduction.AmountTP.ToString();
                            if (payCodeDetails[4] == "0.00")
                            {
                                payCodeDetails[4] = payCodeDetails[6];  // Make Rate equal to amount if rate is zero.
                            }
                            payCodeDetails[7] = rpDeduction.AccountsYearBalance.ToString();
                            payCodeDetails[8] = rpDeduction.AmountYTD.ToString();
                            payCodeDetails[9] = rpDeduction.AccountsYearUnits.ToString();
                            payCodeDetails[10] = rpDeduction.PayeYearUnits.ToString();
                            payCodeDetails[11] = rpDeduction.PayrollAccrued.ToString();
                            switch (payCodeDetails[2]) //PayCode
                            {
                                case "TAX":
                                    payCodeDetails[0] = "0";
                                    payCodeDetails[1] = payHistoryDetails[29];  // Tax Code
                                    payCodeDetails[2] = payHistoryDetails[29];  // Tax Code
                                    payCodeDetails[7] = "0";
                                    payCodeDetails[8] = "0";
                                    payCodeDetails[3] = "T";                    // Tax    
                                    break;
                                case "NI":
                                    payCodeDetails[0] = "0";
                                    payCodeDetails[1] = "NIEeeLERtoUER-A";      // Ee NI
                                    payCodeDetails[2] = "NIEeeLERtoUER";        // Ee NI
                                    payCodeDetails[7] = "0";
                                    payCodeDetails[8] = "0";
                                    payCodeDetails[3] = "T";                    // Tax    
                                    break;
                                case "PENSION":
                                    penPreAmount = rpDeduction.AmountTP;
                                    wait = true;
                                    break;
                                case "PENSIONSS":
                                    payCodeDetails[0] = "0";
                                    payCodeDetails[1] = "PenPreTaxEe";         // Ee Pension
                                    payCodeDetails[2] = "PenPreTaxEe";         // Ee Pension
                                    payCodeDetails[6] = (penPreAmount + rpDeduction.AmountTP).ToString();
                                    payCodeDetails[7] = "0";
                                    payCodeDetails[8] = "0";
                                    payCodeDetails[9] = "0";
                                    payCodeDetails[10] = "0";
                                    payCodeDetails[11] = "0";
                                    wait = false;
                                    break;
                                case "PENSIONRAS":
                                    penPostAmount = rpDeduction.AmountTP;
                                    wait = true;
                                    break;
                                case "PENSIONTAXEX":
                                    payCodeDetails[0] = "0";
                                    payCodeDetails[1] = "PenPostTaxEe";         // Ee Pension
                                    payCodeDetails[2] = "PenPostTaxEe";         // Ee Pension
                                    payCodeDetails[6] = (penPostAmount + rpDeduction.AmountTP).ToString();
                                    payCodeDetails[7] = "0";
                                    payCodeDetails[8] = "0";
                                    payCodeDetails[9] = "0";
                                    payCodeDetails[10] = "0";
                                    payCodeDetails[11] = "0";
                                    wait = false;
                                    break;
                                default:
                                    payCodeDetails[0] = "";
                                    break;

                            }
                            if(!wait)
                            {
                                //
                                //Check if any of the values are not zero. If so write the first employee record
                                //
                                bool allZeros = false;
                                if (Convert.ToDecimal(payCodeDetails[5]) == 0 && Convert.ToDecimal(payCodeDetails[6]) == 0 &&
                                    Convert.ToDecimal(payCodeDetails[7]) == 0 && Convert.ToDecimal(payCodeDetails[8]) == 0 &&
                                    Convert.ToDecimal(payCodeDetails[9]) == 0)
                                {
                                    allZeros = true;

                                }
                                if (!allZeros)
                                {
                                    //Write employee record
                                    WritePayHistoryCSV(rpParameters, payHistoryDetails, payCodeDetails, sw, writeHeader);
                                    writeHeader = false;

                                }
                            }
                            


                        }

                    }


                }

            }

        }
        private void WritePayHistoryCSV(RPParameters rpParameters, string[] payHistoryDetails, string[] payCodeDetails, StreamWriter sw, bool writeHeader)
        {

            string csvLine = null;
            if (writeHeader)
            {
                string csvHeader = "co,runDate,Period_Start_Date,Period_End_Date,process,PayrollYear," +
                              "EEid,Gross,NetPay,Batch,CheckVoucher,Account,Transit,DeptName," +
                              "CostCentreName,branchName,Days/Hours,StudentLoanStartDate," +
                              "StudentLoanEndDate,StudentLoanDeductions,NI Letter,Calculation Basis," +
                              "Total,Earnings To LEL,Earnings To SET,Earnings To PET,Earnings To UST," +
                              "Earnings To AUST,Earnings To UEL,Earnings Above UEL,Ee Contributions Pt1," +
                              "Ee Contributions Pt2,Er Contributions,Ee Rebate,Er Rebate,Ee Reduction," +
                              "LeaveDate,Leaver,Tax Code,Week1/Month 1,Tax Code Change Type ID," +
                              "Tax Code Change Type,Tax Previous Emt,Taxable Pay Previous Emt,Tax This Emt," +
                              "Taxable Pay This Emt,PayCode,payCodeDesc,payCodeValue,det,rate,hours,Amount," +
                              "Acc Year Bal,PAYE Year Bal,Acc Year Units,PAYE Year Units,Payroll Accrued," +
                              "LastAutoEnrolmentAssessmentDate,AutoEnrolmentAssessment,AutoEnrolmentAssessmentEvent," +
                              "AssessmentTaxPeriod,AssessmentTaxYear";
                csvLine = csvHeader;
                sw.WriteLine(csvLine);
                csvLine = null;

            }
            string batch = null;
            switch (rpParameters.PaySchedule)
            {
                case "Monthly":
                    batch = "M";
                    break;
                case "TwoWeekly":
                    batch = "M";
                    break;
                case "FourWeekly":
                    batch = "M";
                    break;
                case "Yearly":
                    batch = "M";
                    break;
                default:
                    batch = "W";
                    break;
            }
            
            string process = null;
            process = "20" + payHistoryDetails[1].Substring(6, 2) + payHistoryDetails[1].Substring(3, 2) + payHistoryDetails[1].Substring(0, 2) + "01";
            csvLine = csvLine + "\"" + rpParameters.ErRef + "\"" + "," +                                   //Co. Number
                            "\"" + payHistoryDetails[1] + "\"" + "," +                                  //Run Date
                            "\"" + payHistoryDetails[2] + "\"" + "," +                                  //Period Start Date
                            "\"" + payHistoryDetails[3] + "\"" + "," +                                  //Period End Date
                            "\"" + process + "\"" + "," +                                               //Process
                            "\"" + payHistoryDetails[4] + "\"" + "," +                                  //Payroll Year
                            "\"" + payHistoryDetails[0] + "\"" + "," +                     //Ee ID
                            "\"" + payHistoryDetails[5] + "\"" + "," +                                  //Gross
                            "\"" + payHistoryDetails[6] + "\"" + "," +                                  //Net
                            "\"" + batch + "\"" + "," +                                                 //batch
                            "\"" + "0" + "\"" + "," +                                                   //CheckVoucher
                            "\"" + "0" + "\"" + "," +                                                   //Account
                            "\"" + "0" + "\"" + "," +                                                   //Transit
                            "\"" + "[Default]" + "\"" + "," +                                           //DeptName
                            "\"" + "[Default]" + "\"" + "," +                                           //CostCentreName
                            "\"" + "[Default]" + "\"" + ",";                                            //branchName

            //From payHistoryDetails[7] (DayHours) to payHistoryDetails[36] (Taxable Pay This)
            for (int i = 7; i < 37; i++)
            {
                csvLine = csvLine + "\"" + payHistoryDetails[i] + "\"" + ",";
            }
            //From payCodeDetails[0] (PayCode) to payCodeDetails[11] (Payroll Accrued)
            for (int i = 0; i < 12; i++)
            {
                csvLine = csvLine + "\"" + payCodeDetails[i] + "\"" + ",";
            }
            //From payHistoryDetails[??] (LastAutoEnrolmentAssessmentDate) to payHistoryDetails[??] (Assessment Tax Year)
            for(int i = 49; i < 54; i++)
            {
                csvLine = csvLine + "\"" + payHistoryDetails[i] + "\"" + ",";
            }

            csvLine = csvLine.TrimEnd(',');

            sw.WriteLine(csvLine);

        }
        public decimal GetDecimalElementByTagFromXml(XmlElement xmlElement, string tag)
        {
            decimal decimalValue = 0;
            string element = GetElementByTagFromXml(xmlElement, tag);
            if (element != null)
            {
                try
                {
                    decimalValue = Convert.ToDecimal(element);
                }
                catch
                {

                }
            }

            return decimalValue;
        }
        public bool GetBooleanElementByTagFromXml(XmlElement xmlElement, string tag)
        {
            bool boolValue = false;
            string element = GetElementByTagFromXml(xmlElement, tag);
            if(element=="Y" || element=="Yes")
            {
                element = "true";
            }
            if (element != null)
            {
                try
                {
                    boolValue = Convert.ToBoolean(element);
                }
                catch
                {

                }
            }

            return boolValue;
        }
        public int GetIntElementByTagFromXml(XmlElement xmlElement, string tag)
        {
            int intValue = 0;

            string element = GetElementByTagFromXml(xmlElement, tag);
            if (element != null)
            {
                try
                {
                    intValue = Convert.ToInt32(element);
                }
                catch
                {

                }
            }

            return intValue;
        }
        public DateTime? GetDateElementByTagFromXml(XmlElement xmlElement, string tag)
        {
            DateTime? dateValue = null;
            string element = GetElementByTagFromXml(xmlElement, tag);
            if (element != null)
            {
                try
                {
                    dateValue = Convert.ToDateTime(element);
                }
                catch
                {

                }
            }


            return dateValue;
        }
        public string GetElementByTagFromXml(XmlElement xmlElement, string tag)
        {
            string element = null;
            int i = xmlElement.GetElementsByTagName(tag).Count;
            if (i > 0)
            {
                element = xmlElement.GetElementsByTagName(tag).Item(0).InnerText;
            }
            return element;
        }


        
        public void ProcessBankAndPensionFiles(XDocument xdoc, List<RPEmployeePeriod> rpEmployeePeriodList, List<RPPensionContribution> rpPensionContributions, RPEmployer rpEmployer, RPParameters rpParameters)
        {
            string outgoingFolder = xdoc.Root.Element("DataHomeFolder").Value + "PE-Reports" + "\\" + rpParameters.ErRef;
            string dataSource = xdoc.Root.Element("DataSource").Value;            //"APPSERVER1\\MSSQL";  //"13.69.154.210\\MSSQL";  
            string dataBase = xdoc.Root.Element("Database").Value;
            string userID = xdoc.Root.Element("Username").Value;
            string password = xdoc.Root.Element("Password").Value;
            string sqlConnectionString = "Server=" + dataSource + ";Database=" + dataBase + ";User ID=" + userID + ";Password=" + password + ";";
            try
            {
                DataRow drCompanyReportCodes = GetCompanyReportCodes(xdoc, sqlConnectionString, rpParameters);
                rpEmployer.BankFileCode = drCompanyReportCodes.ItemArray[0].ToString();                 //BankFileCode
                rpEmployer.PensionReportCode = drCompanyReportCodes.ItemArray[1].ToString();            //PensionReportCode
            }
            catch
            {
                rpEmployer.BankFileCode = "000";
            }
            
            //Bank file code is not equal to "001" so a bank file is required.
            switch (rpEmployer.BankFileCode)
            {
                case "001":
                    //Barclays
                    CreateBarclaysBankFile(outgoingFolder, rpEmployeePeriodList, rpEmployer);
                    break;
                case "002":
                    //Eagle
                    CreateEagleBankFile(outgoingFolder, rpEmployeePeriodList, rpEmployer);
                    break;
                default:
                    //No bank file required
                    break;
            }
            string previousSchemeName = null;
            //Create a list of pension file objects for each scheme name, then use it to write the pension
            //file for that scheme then move on to the next scheme name.
            //The rpPensionContributions object is already sorted in scheme name sequence
            RPPensionFileScheme rpPensionFileScheme = new RPPensionFileScheme();
            List<RPPensionContribution> rpPensionFileSchemePensionContributions = new List<RPPensionContribution>();
            List<RPPensionFileScheme> rpPensionFileSchemes = new List<RPPensionFileScheme>();
            foreach (RPPensionContribution rpPensionContribution in rpPensionContributions)
            {
                if (rpPensionContribution.RPPensionPeriod.SchemeName != previousSchemeName)
                {
                    //We've moved to a new scheme.
                    if (previousSchemeName != null)
                    {
                        //The rpPensionFileScheme object we've create should now contain a scheme name plus a list for employee contributions
                        rpPensionFileScheme.RPPensionContributions = rpPensionFileSchemePensionContributions;
                        rpPensionFileSchemes.Add(rpPensionFileScheme);
                        rpPensionFileScheme = new RPPensionFileScheme();
                        rpPensionFileSchemePensionContributions = new List<RPPensionContribution>();
                    }
                    previousSchemeName = rpPensionContribution.RPPensionPeriod.SchemeName;
                    rpPensionFileScheme.SchemeName = rpPensionContribution.RPPensionPeriod.SchemeName;
                    if(rpPensionFileScheme.SchemeName.ToUpper().Contains("AVIVA"))
                    {
                        rpPensionFileScheme.SchemeProvider = "AVIVA";
                    }
                    else if (rpPensionFileScheme.SchemeName.ToUpper().Contains("NEST"))
                    {
                        rpPensionFileScheme.SchemeProvider = "NEST";
                    }
                    else if (rpPensionFileScheme.SchemeName.ToUpper().Contains("WORKERS PENSION TRUST"))
                    {
                        rpPensionFileScheme.SchemeProvider = "WORKERS PENSION TRUST";
                    }
                    else
                    {
                        rpPensionFileScheme.SchemeProvider = "UNKOWN";
                    }
                }
                rpPensionFileSchemePensionContributions.Add(rpPensionContribution);
            }
            //After the last rpPensionContribution create the final pensionFileScheme and add it to the list.
            //The rpPensionFileScheme object we've create should now contain a scheme name plus a list for employee contributions
            rpPensionFileScheme.RPPensionContributions = rpPensionFileSchemePensionContributions;
            rpPensionFileSchemes.Add(rpPensionFileScheme);
            ProcessPensionFileSchemes(outgoingFolder, rpPensionFileSchemes, rpEmployer);
          }
        private void ProcessPensionFileSchemes(string outgoingFolder, List<RPPensionFileScheme> rpPensionFileSchemes, RPEmployer rpEmployer)
        {
            foreach(RPPensionFileScheme rpPensionFileScheme in rpPensionFileSchemes)
            {
                switch(rpPensionFileScheme.SchemeProvider)
                {
                    case "AVIVA":
                        CreateTheAvivaPensionFile(outgoingFolder, rpPensionFileScheme, rpEmployer);
                        break;
                    case "NEST":
                        CreateTheNestPensionFile(outgoingFolder, rpPensionFileScheme, rpEmployer);
                        break;
                    case "WORKERS PENSION TRUST":
                        CreateTheWorkersPensionTrustPensionFile(outgoingFolder, rpPensionFileScheme, rpEmployer);
                        break;
                    case "UNKNOWN":
                        break;
                }
            }
        }
        private void CreateBarclaysBankFile(string outgoingFolder, List<RPEmployeePeriod> rpEmployeePeriodList, RPEmployer rpEmployer)
        {
            string bankFileName = outgoingFolder + "\\" + "BarclaysBankFile.txt";
            string quotes = "\"";
            string comma = ",";
            //Create the Barclays bank file which does not have a header row.
            using (StreamWriter sw = new StreamWriter(bankFileName))
            {
                string csvLine = null;
                foreach (RPEmployeePeriod rpEmployeePeriod in rpEmployeePeriodList)
                {
                    if (rpEmployeePeriod.PaymentMethod == "BACS")
                    {
                        string fullName = rpEmployeePeriod.Forename + " " + rpEmployeePeriod.Surname;
                        fullName = fullName.ToUpper();
                        csvLine = quotes + rpEmployeePeriod.SortCode + quotes + comma +
                                  quotes + fullName + quotes + comma +
                                  quotes + rpEmployeePeriod.BankAccNo + quotes + comma +
                                  quotes + rpEmployeePeriod.NetPayTP.ToString() + quotes + comma +
                                  quotes + rpEmployer.Name.ToUpper() + quotes + comma +
                                  quotes + "99" + quotes;
                        sw.WriteLine(csvLine);
                    }
                }
            }
        }
        private void CreateEagleBankFile(string outgoingFolder, List<RPEmployeePeriod> rpEmployeePeriodList, RPEmployer rpEmployer)
        {
            string bankFileName = outgoingFolder + "\\" + "EagleBankFile.csv";
            string comma = ",";
            //Create the Eagle bank file which does have a header row.
            using (StreamWriter sw = new StreamWriter(bankFileName))
            {
                //Write the header row
                string csvLine = "AccName,SortCode,AccNumber,Amount,Ref";
                sw.WriteLine(csvLine);
                foreach (RPEmployeePeriod rpEmployeePeriod in rpEmployeePeriodList)
                {
                    if(rpEmployeePeriod.PaymentMethod == "BACS")
                    {
                        string fullName = rpEmployeePeriod.Forename + " " + rpEmployeePeriod.Surname;
                        fullName = fullName.ToUpper();
                        csvLine = fullName + comma +
                                  rpEmployeePeriod.SortCode + comma +
                                  rpEmployeePeriod.BankAccNo + comma +
                                  rpEmployeePeriod.NetPayTP.ToString() + comma +
                                  fullName;
                        sw.WriteLine(csvLine);
                    }
               }
            }
        }
        
        private void CreateTheNestPensionFile(string outgoingFolder, RPPensionFileScheme rpPensionFileScheme, RPEmployer rpEmployer)
        {
            string pensionFileName = outgoingFolder + "\\" + rpPensionFileScheme.SchemeName + "PensionFile.csv";
            string comma = ",";
            string providerEmployerReference = rpPensionFileScheme.RPPensionContributions[0].RPPensionPeriod.ProviderEmployerReference;
            string startDate = rpPensionFileScheme.RPPensionContributions[0].StartDate.ToString("yyyy-MM-dd");
            string endDate = rpPensionFileScheme.RPPensionContributions[0].EndDate.ToString("yyyy-MM-dd");
            string frequency = rpPensionFileScheme.RPPensionContributions[0].Freq;
            string blank = "";
            string zeroContributions = "";
            string header = 'H' + comma + providerEmployerReference + comma +
                                            "CS" + comma + endDate + comma + "My Source" +
                                            comma + blank + comma + frequency + comma + blank +
                                            comma + blank + comma + startDate;

            using (StreamWriter sw = new StreamWriter(pensionFileName))
            {
                sw.WriteLine(header);
                string csvLine = null;

                foreach (RPPensionContribution rpPensionContribution in rpPensionFileScheme.RPPensionContributions)
                {
                    zeroContributions = ""; //need to reset the value else it will always be 5 
                    if (rpPensionContribution.RPPensionPeriod.EePensionTaxPeriod == 0 && rpPensionContribution.RPPensionPeriod.ErPensionTaxPeriod == 0)
                    {
                        zeroContributions = "5";
                    }
                    csvLine = 'D' + comma + rpPensionContribution.Surname + comma + rpPensionContribution.NINumber +
                    comma + rpPensionContribution.EeRef + comma + rpPensionContribution.RPPensionPeriod.PensionablePayTaxPeriod + comma +
                    blank + comma + rpPensionContribution.RPPensionPeriod.ErPensionTaxPeriod + comma + rpPensionContribution.RPPensionPeriod.EePensionTaxPeriod +
                    comma + zeroContributions;
                    sw.WriteLine(csvLine);
                }
                string footer = 'T' + comma + rpPensionFileScheme.RPPensionContributions.Count + comma + '3';
                sw.WriteLine(footer);
            }
        }
        
        private void CreateTheWorkersPensionTrustPensionFile(string outgoingFolder, RPPensionFileScheme rpPensionFileScheme, RPEmployer rpEmployer)
        {
            string pensionFileName = outgoingFolder + "\\" + rpPensionFileScheme.SchemeName + "PensionFile.csv";
            string comma = ",";
            
            using (StreamWriter sw = new StreamWriter(pensionFileName))
            {
                string csvLine = null;

                foreach (RPPensionContribution rpPensionContribution in rpPensionFileScheme.RPPensionContributions)
                {
                    if (rpPensionContribution.RPPensionPeriod.EePensionTaxPeriod != 0 || rpPensionContribution.RPPensionPeriod.ErPensionTaxPeriod != 0) //if ee has either Ee or Er contributions
                    {
                        csvLine = rpPensionContribution.NINumber + comma + rpPensionContribution.ForenameSurname + comma +
                                        rpPensionContribution.PayRunDate.ToString("yyyy/MM/dd") + comma + rpPensionContribution.RPPensionPeriod.ErPensionTaxPeriod +
                                        comma + rpPensionContribution.RPPensionPeriod.EePensionTaxPeriod;

                        sw.WriteLine(csvLine);
                    }
                }
            }
        }
        //private void CreateAvivaPensionFile(string outgoingFolder, List<RPEmployeePeriod> rpEmployeePeriodList, RPEmployer rpEmployer)
        //{
        //    string pensionFileName = outgoingFolder + "\\" + "AvivaPensionFile.csv";
        //    string comma = ",";
        //    string pension = "AVIVA";
        //    string header = "PayrollMonth,Name,NInumber,AlternativeuniqueID,Employerregularcontributionamount,Employeeregulardeduction,Reasonforpartialornon-payment,Employerregularcontributionamount,Employeeoneoffcontribution,NewcategoryID";

        //    using (StreamWriter sw = new StreamWriter(pensionFileName))
        //    {
        //        sw.WriteLine(header);
        //        string csvLine = null;

        //        foreach (RPEmployeePeriod rpEmployeePeriod in rpEmployeePeriodList)
        //        {
        //            foreach (RPPensionPeriod rpPensionPeriod in rpEmployeePeriod.Pensions)
        //            {
        //                if (rpPensionPeriod.SchemeName.ToUpper().Contains(pension))
        //                {

        //                }
        //                bool contains = rpPensionPeriod.SchemeName.IndexOf(pension, StringComparison.OrdinalIgnoreCase) >= 0;
        //                if (contains)
        //                {
        //                    if (rpPensionPeriod.EePensionTaxPeriod != 0 || rpPensionPeriod.ErPensionTaxPeriod != 0) //if ee has either Ee or Er contributions
        //                    {
        //                        csvLine = rpEmployeePeriod.PayRunDate.ToString("MMM-yy") + comma + rpEmployeePeriod.Surname + comma + rpEmployeePeriod.NINumber +
        //                                  comma + rpEmployeePeriod.Reference + comma + rpPensionPeriod.ErPensionTaxPeriod + comma + rpPensionPeriod.EePensionTaxPeriod +
        //                                  comma + comma + comma + comma;

        //                        sw.WriteLine(csvLine);
        //                    }
        //                }
        //            }
        //        }
        //    }
        //}
        private void CreateTheAvivaPensionFile(string outgoingFolder, RPPensionFileScheme rpPensionFileScheme, RPEmployer rpEmployer)
        {
            string pensionFileName = outgoingFolder + "\\" + rpPensionFileScheme.SchemeName + "PensionFile.csv";
            string comma = ",";
            string header = "PayrollMonth,Name,NInumber,AlternativeuniqueID,Employerregularcontributionamount,Employeeregulardeduction,Reasonforpartialornon-payment,Employerregularcontributionamount,Employeeoneoffcontribution,NewcategoryID";

            using (StreamWriter sw = new StreamWriter(pensionFileName))
            {
                sw.WriteLine(header);
                string csvLine = null;

                foreach (RPPensionContribution rpPensionContribution in rpPensionFileScheme.RPPensionContributions)
                {
                    if (rpPensionContribution.RPPensionPeriod.EePensionTaxPeriod != 0 || rpPensionContribution.RPPensionPeriod.ErPensionTaxPeriod != 0) //if ee has either Ee or Er contributions
                    {
                        csvLine = rpPensionContribution.PayRunDate.ToString("MMM-yy") + comma + rpPensionContribution.Surname + comma + rpPensionContribution.NINumber +
                                    comma + rpPensionContribution.EeRef + comma + rpPensionContribution.RPPensionPeriod.ErPensionTaxPeriod + comma + 
                                    rpPensionContribution.RPPensionPeriod.EePensionTaxPeriod +
                                    comma + comma + comma + comma;

                        sw.WriteLine(csvLine);
                    }

                }
            }
        }

        
        public void PrintStandardReports(XDocument xdoc, List<RPEmployeePeriod> rpEmployeePeriodList, RPEmployer rpEmployer, RPParameters rpParameters, 
                                         List<P45> p45s, List<RPPayComponent> rpPayComponents, List<RPPensionContribution> rpPensionContributions)
        {
            PrintPayslips(xdoc, rpEmployeePeriodList, rpEmployer, rpParameters);
            PrintPayslipsSimple(xdoc, rpEmployeePeriodList, rpEmployer, rpParameters);
            PrintPaymentsDueByMethodReport(xdoc, rpEmployeePeriodList, rpEmployer, rpParameters);
            PrintComponentAnalysisReport(xdoc, rpPayComponents, rpEmployer, rpParameters);
            PrintPensionContributionsReport(xdoc, rpPensionContributions, rpEmployer, rpParameters);
            PrintPayrollRunDetailsReport(xdoc, rpEmployeePeriodList, rpEmployer, rpParameters);
            if (p45s.Count > 0)
            {
                PrintP45s(xdoc, p45s, rpParameters);
            }
        }
        
        public string[] RemoveBlankAddressLines(string[] oldAddress)
        {
            string[] newAddress = new string[6];
            int x = 0;
            for (int i = 0; i < 6; i++)
            {
                if (oldAddress[i] != "" && oldAddress[i] != " " && oldAddress[i] != null)
                {
                    newAddress[x] = oldAddress[i];
                    x++;
                }
            }
            for (int i = x; i < 6; i++)
            {
                newAddress[i] = "";
            }
            return newAddress;
        }
        public decimal CalculateHMRCTotal(RPP32Report rpP32Report, int payeMonth)
        {
            decimal hmrcTotal = 0;
            foreach(RPP32ReportMonth rpP32ReportMonth in rpP32Report.RPP32ReportMonths)
            {
                if(rpP32ReportMonth.PeriodNo==payeMonth)
                {
                    hmrcTotal = rpP32ReportMonth.RPP32Summary.AmountDue;
                }
            }
            return hmrcTotal;
        }
        private void PrintPayslips(XDocument xdoc, List<RPEmployeePeriod> rpEmployeePeriodList, RPEmployer rpEmployer, RPParameters rpParameters)
        {
            string softwareHomeFolder = xdoc.Root.Element("SoftwareHomeFolder").Value + "Programs\\";
            string outgoingFolder = xdoc.Root.Element("DataHomeFolder").Value + "PE-Reports";
            string coNo = rpParameters.ErRef;
            string coName = rpEmployer.Name;
            int taxYear = rpParameters.TaxYear;
            int taxPeriod = rpParameters.TaxPeriod;
            string freq = rpParameters.PaySchedule;

            //Main payslip report
            XtraReport report1 = XtraReport.FromFile(softwareHomeFolder + "Payslip.repx", true);
            report1.Parameters["CmpName"].Value = rpEmployer.Name;
            report1.Parameters["PayeRef"].Value = rpEmployer.PayeRef;
            report1.Parameters["Date"].Value = rpParameters.PayRunDate; //.AccYearEnd;
            report1.Parameters["Period"].Value = rpParameters.TaxPeriod;
            report1.DataSource = rpEmployeePeriodList;
            //// To show the report designer. You need to uncomment this to design the report.
            //// You also need to comment out the report.ExportToPDF line below
            ////
            bool designMode = false;
            if (designMode)
            {
                report1.ShowDesigner();
                //report1.ShowPreview();

            }
            else
            {
                // Export to pdf file.

                //
                // I'm going to remove spaces from the document name. I'll replace them with dashes
                //
                //string dirName = "V:\\Payescape\\PayRunIO\\WG\\";

                string dirName = outgoingFolder + "\\" + coNo + "\\";
                Directory.CreateDirectory(dirName);
                string docName = coNo + "_PayslipReportFor_TaxYear_" + taxYear + "_Period_" + taxPeriod + ".pdf";

                report1.ExportToPdf(dirName + docName);
                //docName = docName.Replace(".pdf", ".xlsx");
                //report1.ExportToXlsx(dirName + docName);

            }

        }
        private void PrintPayslipsSimple(XDocument xdoc, List<RPEmployeePeriod> rpEmployeePeriodList, RPEmployer rpEmployer, RPParameters rpParameters)
        {
            string softwareHomeFolder = xdoc.Root.Element("SoftwareHomeFolder").Value + "Programs\\";
            string outgoingFolder = xdoc.Root.Element("DataHomeFolder").Value + "PE-Reports";
            string coNo = rpParameters.ErRef;
            string coName = rpEmployer.Name;
            int taxYear = rpParameters.TaxYear;
            int taxPeriod = rpParameters.TaxPeriod;
            string freq = rpParameters.PaySchedule;

            //Main payslip report
            XtraReport report1 = XtraReport.FromFile(softwareHomeFolder + "PayslipSimple.repx", true);
            report1.Parameters["CmpName"].Value = rpEmployer.Name;
            report1.Parameters["PayeRef"].Value = rpEmployer.PayeRef;
            report1.Parameters["Date"].Value = rpParameters.PayRunDate; //.AccYearEnd;
            report1.Parameters["Period"].Value = rpParameters.TaxPeriod;
            report1.DataSource = rpEmployeePeriodList;
            //// To show the report designer. You need to uncomment this to design the report.
            //// You also need to comment out the report.ExportToPDF line below
            ////
            bool designMode = false;
            if (designMode)
            {
                report1.ShowDesigner();
                //report1.ShowPreview();

            }
            else
            {
                // Export to pdf file.

                //
                // I'm going to remove spaces from the document name. I'll replace them with dashes
                //
                //string dirName = "V:\\Payescape\\PayRunIO\\WG\\";

                string dirName = outgoingFolder + "\\" + coNo + "\\";
                Directory.CreateDirectory(dirName);
                string docName = coNo + "_PayslipReportFor_TaxYear_" + taxYear + "_Period_" + taxPeriod + ".xlsx";

                //report1.ExportToPdf(dirName + docName);
                //docName = docName.Replace(".pdf", ".xlsx");
                report1.ExportToXlsx(dirName + docName);

            }

        }
        private void PrintPaymentsDueByMethodReport(XDocument xdoc, List<RPEmployeePeriod> rpEmployeePeriodList, RPEmployer rpEmployer, RPParameters rpParameters)
        {
            string softwareHomeFolder = xdoc.Root.Element("SoftwareHomeFolder").Value + "Programs\\";
            string outgoingFolder = xdoc.Root.Element("DataHomeFolder").Value + "PE-Reports";
            string coNo = rpParameters.ErRef;
            string coName = rpEmployer.Name;
            int taxYear = rpParameters.TaxYear;
            int taxPeriod = rpParameters.TaxPeriod;
            string freq = rpParameters.PaySchedule;
            //var payeMonth = rpParameters.AccYearEnd.Day < 6 ? rpParameters.AccYearEnd.Month - 4 : rpParameters.AccYearEnd.Month - 3;
            var payeMonth = rpParameters.PayRunDate.Day < 6 ? rpParameters.PayRunDate.Month - 4 : rpParameters.PayRunDate.Month - 3;
            if (payeMonth <= 0)
            {
                payeMonth += 12;
            }

            //Main payslip report
            XtraReport report1 = XtraReport.FromFile(softwareHomeFolder + "PaymentsDueByMethodsReport.repx", true);         //"PaymentsDueByMethodReport.repx"

            report1.Parameters["CmpName"].Value = rpEmployer.Name;
            report1.Parameters["PayeRef"].Value = rpEmployer.PayeRef;
            report1.Parameters["Date"].Value = rpParameters.PayRunDate;
            report1.Parameters["Period"].Value = rpParameters.TaxPeriod;
            report1.Parameters["Freq"].Value = rpParameters.PaySchedule;
            report1.Parameters["PAYEMonth"].Value = payeMonth;
            report1.DataSource = rpEmployeePeriodList;
            //// To show the report designer. You need to uncomment this to design the report.
            //// You also need to comment out the report.ExportToPDF line below
            ////
            bool designMode = false;
            if (designMode)
            {
                report1.ShowDesigner();
                //report1.ShowPreview();

            }
            else
            {
                // Export to pdf file.

                //
                // I'm going to remove spaces from the document name. I'll replace them with dashes
                //
                //string dirName = "V:\\Payescape\\PayRunIO\\WG\\";
                string dirName = outgoingFolder + "\\" + coNo + "\\";
                Directory.CreateDirectory(dirName);
                string docName = coNo + "_PaymentsDueByMethodReportFor_TaxYear_" + taxYear + "_Period_" + taxPeriod + ".pdf";

                report1.ExportToPdf(dirName + docName);
                docName = docName.Replace(".pdf", ".xlsx");
                report1.ExportToXlsx(dirName + docName);
            }

        }
        private void PrintPensionContributionsReport(XDocument xdoc, List<RPPensionContribution> rpPensionContributions, RPEmployer rpEmployer, RPParameters rpParameters)
        {
            string softwareHomeFolder = xdoc.Root.Element("SoftwareHomeFolder").Value + "Programs\\";
            string outgoingFolder = xdoc.Root.Element("DataHomeFolder").Value + "PE-Reports";
            string coNo = rpParameters.ErRef;
            string coName = rpEmployer.Name;
            int taxYear = rpParameters.TaxYear;
            int taxPeriod = rpParameters.TaxPeriod;
            string freq = rpParameters.PaySchedule;
            //var payeMonth = rpParameters.AccYearEnd.Day < 6 ? rpParameters.AccYearEnd.Month - 4 : rpParameters.AccYearEnd.Month - 3;
            var payeMonth = rpParameters.PayRunDate.Day < 6 ? rpParameters.PayRunDate.Month - 4 : rpParameters.PayRunDate.Month - 3;
            if (payeMonth <= 0)
            {
                payeMonth += 12;
            }

            //Main payslip report
            XtraReport report1 = XtraReport.FromFile(softwareHomeFolder + "PensionContributionsReport.repx", true);         //"PensionContributionsReport.repx"

            report1.Parameters["CmpName"].Value = rpEmployer.Name;
            report1.Parameters["PayeRef"].Value = rpEmployer.PayeRef;
            report1.Parameters["Date"].Value = rpParameters.PayRunDate;
            report1.Parameters["Period"].Value = rpParameters.TaxPeriod;
            report1.Parameters["Freq"].Value = rpParameters.PaySchedule;
            report1.Parameters["PAYEMonth"].Value = payeMonth;
            report1.DataSource = rpPensionContributions;
            //// To show the report designer. You need to uncomment this to design the report.
            //// You also need to comment out the report.ExportToPDF line below
            ////
            bool designMode = false;
            if (designMode)
            {
                report1.ShowDesigner();
                //report1.ShowPreview();

            }
            else
            {
                // Export to pdf file.

                //
                // I'm going to remove spaces from the document name. I'll replace them with dashes
                //
                //string dirName = "V:\\Payescape\\PayRunIO\\WG\\";
                string dirName = outgoingFolder + "\\" + coNo + "\\";
                Directory.CreateDirectory(dirName);
                string docName = coNo + "_PensionContributionsReportFor_TaxYear_" + taxYear + "_Period_" + taxPeriod + ".pdf";

                report1.ExportToPdf(dirName + docName);
                docName = docName.Replace(".pdf", ".xlsx");
                report1.ExportToXlsx(dirName + docName);
            }

        }
        private void PrintComponentAnalysisReport(XDocument xdoc, List<RPPayComponent> rpPayComponents, RPEmployer rpEmployer, RPParameters rpParameters)
        {
            string softwareHomeFolder = xdoc.Root.Element("SoftwareHomeFolder").Value + "Programs\\";
            string outgoingFolder = xdoc.Root.Element("DataHomeFolder").Value + "PE-Reports";
            string coNo = rpParameters.ErRef;
            string coName = rpEmployer.Name;
            int taxYear = rpParameters.TaxYear;
            int taxPeriod = rpParameters.TaxPeriod;
            string freq = rpParameters.PaySchedule;
            //var payeMonth = rpParameters.AccYearEnd.Day < 6 ? rpParameters.AccYearEnd.Month - 4 : rpParameters.AccYearEnd.Month - 3;
            var payeMonth = rpParameters.PayRunDate.Day < 6 ? rpParameters.PayRunDate.Month - 4 : rpParameters.PayRunDate.Month - 3;
            if (payeMonth <= 0)
            {
                payeMonth += 12;
            }

            //Main payslip report
            XtraReport report1 = XtraReport.FromFile(softwareHomeFolder + "ComponentAnalysisReport.repx", true);         //"ComponentAnalysisReport.repx"

            report1.Parameters["CmpName"].Value = rpEmployer.Name;
            report1.Parameters["PayeRef"].Value = rpEmployer.PayeRef;
            report1.Parameters["Date"].Value = rpParameters.PayRunDate;
            report1.Parameters["Period"].Value = rpParameters.TaxPeriod;
            report1.Parameters["Freq"].Value = rpParameters.PaySchedule;
            report1.Parameters["PAYEMonth"].Value = payeMonth;
            report1.DataSource = rpPayComponents;
            //// To show the report designer. You need to uncomment this to design the report.
            //// You also need to comment out the report.ExportToPDF line below
            ////
            bool designMode = false;
            if (designMode)
            {
                report1.ShowDesigner();
                //report1.ShowPreview();

            }
            else
            {
                // Export to pdf file.

                //
                // I'm going to remove spaces from the document name. I'll replace them with dashes
                //
                //string dirName = "V:\\Payescape\\PayRunIO\\WG\\";
                string dirName = outgoingFolder + "\\" + coNo + "\\";
                Directory.CreateDirectory(dirName);
                string docName = coNo + "_ComponentAnalysisReportFor_TaxYear_" + taxYear + "_Period_" + taxPeriod + ".pdf";

                report1.ExportToPdf(dirName + docName);
                docName = docName.Replace(".pdf", ".xlsx");
                report1.ExportToXlsx(dirName + docName);
            }

        }
        private void PrintPayrollRunDetailsReport(XDocument xdoc, List<RPEmployeePeriod> rpEmployeePeriodList, RPEmployer rpEmployer, RPParameters rpParameters)
        {
            string softwareHomeFolder = xdoc.Root.Element("SoftwareHomeFolder").Value + "Programs\\";
            string outgoingFolder = xdoc.Root.Element("DataHomeFolder").Value + "PE-Reports";
            string coNo = rpParameters.ErRef;
            string coName = rpEmployer.Name;
            int taxYear = rpParameters.TaxYear;
            int taxPeriod = rpParameters.TaxPeriod;
            string freq = rpParameters.PaySchedule;
            //var payeMonth = rpParameters.AccYearEnd.Day < 6 ? rpParameters.AccYearEnd.Month - 4 : rpParameters.AccYearEnd.Month - 3;
            var payeMonth = rpParameters.PayRunDate.Day < 6 ? rpParameters.PayRunDate.Month - 4 : rpParameters.PayRunDate.Month - 3;
            if (payeMonth <= 0)
            {
                payeMonth += 12;
            }

            //Main payslip report
            XtraReport report1 = XtraReport.FromFile(softwareHomeFolder + "PayrollRunDetailsReport.repx", true);         //"PayrollRunDetailsReport.repx"

            report1.Parameters["CmpName"].Value = rpEmployer.Name;
            report1.Parameters["PayeRef"].Value = rpEmployer.PayeRef;
            report1.Parameters["Date"].Value = rpParameters.PayRunDate;
            report1.Parameters["Period"].Value = rpParameters.TaxPeriod;
            report1.Parameters["Freq"].Value = rpParameters.PaySchedule;
            report1.Parameters["PAYEMonth"].Value = payeMonth;
            report1.DataSource = rpEmployeePeriodList;
            //// To show the report designer. You need to uncomment this to design the report.
            //// You also need to comment out the report.ExportToPDF line below
            ////
            bool designMode = false;
            if (designMode)
            {
                report1.ShowDesigner();
                //report1.ShowPreview();

            }
            else
            {
                // Export to pdf file.

                //
                // I'm going to remove spaces from the document name. I'll replace them with dashes
                //
                //string dirName = "V:\\Payescape\\PayRunIO\\WG\\";
                string dirName = outgoingFolder + "\\" + coNo + "\\";
                Directory.CreateDirectory(dirName);
                string docName = coNo + "_PayrollRunDetailsReportFor_TaxYear_" + taxYear + "_Period_" + taxPeriod + ".pdf";

                report1.ExportToPdf(dirName + docName);
                docName = docName.Replace(".pdf", ".xlsx");
                report1.ExportToXlsx(dirName + docName);
            }

        }
        
        private void PrintP45s(XDocument xdoc, List<P45> p45s, RPParameters rpParameters)
        {
            string softwareHomeFolder = xdoc.Root.Element("SoftwareHomeFolder").Value + "Programs\\";
            string outgoingFolder = xdoc.Root.Element("DataHomeFolder").Value + "PE-Reports";
            string coNo = rpParameters.ErRef;
            int taxYear = rpParameters.TaxYear;
            int taxPeriod = rpParameters.TaxPeriod;
            //P45 report
            XtraReport report1 = XtraReport.FromFile(softwareHomeFolder + "P45.repx", true);
            report1.DataSource = p45s;
            //// To show the report designer. You need to uncomment this to design the report.
            //// You also need to comment out the report.ExportToPDF line below
            ////
            bool designMode = false;
            if (designMode)
            {
                report1.ShowDesigner();
                //report1.ShowPreview();

            }
            else
            {
                // Export to pdf file.

                //
                // I'm going to remove spaces from the document name. I'll replace them with dashes
                //
                string dirName = outgoingFolder + "\\" + coNo + "\\";
                Directory.CreateDirectory(dirName);
                string docName = coNo + "_P45ReportFor_TaxYear_" + taxYear + "_Period_" + taxPeriod + ".pdf";

                report1.ExportToPdf(dirName + docName);
                docName = docName.Replace(".pdf", ".xlsx");
                report1.ExportToXlsx(dirName + docName);
            }

        }
        public void PrintP32Report(XDocument xdoc, RPP32Report rpP32Report, RPParameters rpParameters)
        {
            string softwareHomeFolder = xdoc.Root.Element("SoftwareHomeFolder").Value + "Programs\\";
            string outgoingFolder = xdoc.Root.Element("DataHomeFolder").Value + "PE-Reports";
            string coNo = rpParameters.ErRef;
            string coName = rpP32Report.EmployerName;
            int taxYear = rpP32Report.TaxYear;
            int taxPeriod = rpParameters.TaxPeriod;
            string freq = rpParameters.PaySchedule;
            var payeMonth = rpParameters.PayRunDate.Day < 6 ? rpParameters.PayRunDate.Month - 4 : rpParameters.PayRunDate.Month - 3;
            if (payeMonth <= 0)
            {
                payeMonth += 12;
            }

            //Main P32 report
            XtraReport report1 = XtraReport.FromFile(softwareHomeFolder + "P32Report.repx", true);
            report1.Parameters["CmpName"].Value = coName;
            report1.Parameters["PayeRef"].Value = rpP32Report.EmployerPayeRef;
            report1.Parameters["Date"].Value = rpParameters.PayRunDate; //.AccYearEnd;
            report1.Parameters["Period"].Value = rpParameters.TaxPeriod;
            report1.Parameters["Freq"].Value = rpParameters.PaySchedule;
            report1.Parameters["PAYEMonth"].Value = payeMonth;
            report1.Parameters["AnnualEmploymentAllowance"].Value = rpP32Report.AnnualEmploymentAllowance;
            report1.Parameters["PaymentRef"].Value = rpP32Report.PaymentRef;
            report1.Parameters["TaxYearStartDate"].Value = rpP32Report.TaxYearStartDate;
            report1.Parameters["TaxYearEndDate"].Value = rpP32Report.TaxYearEndDate;
            report1.DataSource = rpP32Report.RPP32ReportMonths;
            //// To show the report designer. You need to uncomment this to design the report.
            //// You also need to comment out the report.ExportToPDF line below
            ////
            bool designMode = false;
            if (designMode)
            {
                report1.ShowDesigner();
                //report1.ShowPreview();

            }
            else
            {
                // Export to pdf file.

                //
                // I'm going to remove spaces from the document name. I'll replace them with dashes
                //
                //string dirName = "V:\\Payescape\\PayRunIO\\WG\\";

                string dirName = outgoingFolder + "\\" + coNo + "\\";
                Directory.CreateDirectory(dirName);
                string docName = coNo + "_P32ReportFor_TaxYear_" + taxYear + "_Period_" + taxPeriod + ".pdf";

                report1.ExportToPdf(dirName + docName);
                docName = docName.Replace(".pdf", ".xlsx");
                report1.ExportToXlsx(dirName + docName);

            }
        }
        public void ZipReports(XDocument xdoc, RPEmployer rpEmployer, RPParameters rpParameters)
        {
            int logOneIn = Convert.ToInt32(xdoc.Root.Element("LogOneIn").Value);
            string configDirName = xdoc.Root.Element("SoftwareHomeFolder").Value;
            string textLine = null;

            //
            // Zip the folder.
            //
            string dateTimeStamp = DateTime.Now.ToString("yyyyMMddhhmmssfff");
            string sourceFolder = xdoc.Root.Element("DataHomeFolder").Value + "PE-Reports\\" + rpParameters.ErRef;
            string zipFileName = xdoc.Root.Element("DataHomeFolder").Value + "PE-Reports\\" + rpParameters.ErRef + "_PDF_Reports_" + rpEmployer.HMRCDesc + "_" + dateTimeStamp + ".zip";
            string zipFileNameNoPassword = xdoc.Root.Element("DataHomeFolder").Value + "PE-ReportsNoPassword\\" + rpParameters.ErRef + "_PDF_Reports_" + rpEmployer.HMRCDesc + "_" + dateTimeStamp + ".zip";
            string password = null;
            password = rpEmployer.Name.ToLower().Replace(" ", "");
            if (password.Length >= 4)
            {
                password = password.Substring(0, 4);
            }
            password = password + rpParameters.ErRef;
            try
            {
                using (Ionic.Zip.ZipFile zip = new Ionic.Zip.ZipFile())
                {
                    zip.Password = password;
                    zip.AddDirectory(sourceFolder);
                    zip.Save(zipFileName);
                }
                //Create a copy of the reports with no password for Emer & Mark
                using (Ionic.Zip.ZipFile zip = new Ionic.Zip.ZipFile())
                {
                    zip.AddDirectory(sourceFolder);
                    zip.Save(zipFileNameNoPassword);
                }

                DeleteFilesThenFolder(xdoc, sourceFolder);

            }
            catch (Exception ex)
            {
                textLine = string.Format("Error zipping pdf reports for source folder, {0}.\r\n{1}.\r\n", sourceFolder, ex);
                update_Progress(textLine, configDirName, logOneIn);
            }

        }
        private void DeleteFilesThenFolder(XDocument xdoc, string sourceFolder)
        {
            int logOneIn = Convert.ToInt32(xdoc.Root.Element("LogOneIn").Value);
            string configDirName = xdoc.Root.Element("SoftwareHomeFolder").Value;
            string textLine = null;

            try
            {
                DirectoryInfo dirInfo = new DirectoryInfo(sourceFolder);
                FileInfo[] files = dirInfo.GetFiles();
                foreach (FileInfo file in files)
                {
                    file.Delete();
                }
                Directory.Delete(sourceFolder);
            }
            catch (Exception ex)
            {
                textLine = string.Format("Error deleting files from source folder, {0}.\r\n{1}.\r\n", sourceFolder, ex);
                update_Progress(textLine, configDirName, logOneIn);
            }
        }
        public void EmailZippedReports(XDocument xdoc, RPEmployer rpEmployer, RPParameters rpParameters)
        {
            int logOneIn = Convert.ToInt32(xdoc.Root.Element("LogOneIn").Value);
            string configDirName = xdoc.Root.Element("SoftwareHomeFolder").Value;
            string textLine = null;

            string reportFolder = xdoc.Root.Element("DataHomeFolder").Value + "PE-Reports";
            try
            {
                DirectoryInfo dirInfo = new DirectoryInfo(reportFolder);
                FileInfo[] files = dirInfo.GetFiles();
                foreach (FileInfo file in files)
                {
                    EmailZippedReport(xdoc, file, rpParameters, rpEmployer);
                    file.MoveTo(file.FullName.Replace("PE-Reports", "PE-Reports\\Archive"));
                }

            }
            catch (Exception ex)
            {
                textLine = string.Format("Error emailing zipped pdf reports for report folder, {0}.\r\n{1}.\r\n", reportFolder, ex);
                update_Progress(textLine, configDirName, logOneIn);
            }
        }
        public void UploadZippedReportsToAmazonS3(XDocument xdoc, RPEmployer rpEmployer, RPParameters rpParameters)
        {
            int logOneIn = Convert.ToInt32(xdoc.Root.Element("LogOneIn").Value);
            string configDirName = xdoc.Root.Element("SoftwareHomeFolder").Value;
            string textLine = null;

            string reportFolder = xdoc.Root.Element("DataHomeFolder").Value + "PE-ReportsNoPassword";
            try
            {
                DirectoryInfo dirInfo = new DirectoryInfo(reportFolder);
                FileInfo[] files = dirInfo.GetFiles();
                foreach (FileInfo file in files)
                {
                    UploadZippedReportToAmazonS3(xdoc, file, rpParameters, rpEmployer);
                    file.MoveTo(file.FullName.Replace("PE-ReportsNoPassword", "PE-ReportsNoPassword\\Archive"));
                }

            }
            catch (Exception ex)
            {
                textLine = string.Format("Error uploading zipped pdf reports to Amazon S3 for report folder, {0}.\r\n{1}.\r\n", reportFolder, ex);
                update_Progress(textLine, configDirName, logOneIn);
            }
        }
        private void UploadZippedReportToAmazonS3(XDocument xdoc, FileInfo file, RPParameters rpParameters, RPEmployer rpEmployer)
        {
            int logOneIn = Convert.ToInt32(xdoc.Root.Element("LogOneIn").Value);
            string configDirName = xdoc.Root.Element("SoftwareHomeFolder").Value;
            string reportFolder = xdoc.Root.Element("DataHomeFolder").Value + "PE-Reports";
            string awsBucketName = xdoc.Root.Element("AwsBucketName").Value;
            string awsAccessKey = xdoc.Root.Element("AwsAccessKey").Value;
            string awsAccessSecret = xdoc.Root.Element("AwsAccessSecret").Value;
            bool awsInDevelopment = Convert.ToBoolean(xdoc.Root.Element("InDevelopment").Value);

            bool live = Convert.ToBoolean(xdoc.Root.Element("Live").Value);
            string bucketName = awsBucketName;
            RegionEndpoint bucketRegion = RegionEndpoint.EUWest2;
            IAmazonS3 s3Client;
            if(awsInDevelopment)
            {
                s3Client = new AmazonS3Client(awsAccessKey, awsAccessSecret, RegionEndpoint.EUWest2);
            }
            else
            {
                s3Client = new AmazonS3Client(RegionEndpoint.EUWest2);
            }
            string folderPath;
            if (live)
            {
                folderPath = "PE-ReportsLive/";
            }
            else
            {
                folderPath = "PE-ReportsTest/";
            }
            
            PutObjectRequest request = new PutObjectRequest()
            {
                InputStream = file.OpenRead(),
                BucketName = bucketName,
                Key = folderPath + file.ToString()
            };
            PutObjectResponse response = s3Client.PutObject(request);
            
            //Create a folder in S3 - Don't really need this I don't think since I'll create all the reports in one folder. Each zipped file is a set of reports.
            //PutObjectRequest request = new PutObjectRequest()
            //{
            //    BucketName=bucketName,
            //    Key=folderPath
            //};
            //PutObjectResponse response = s3Client.PutObject(request);
            //Copy a file into S3 bucket.
            
            //UploadFileAsync(s3Client, file, bucketName).Wait();

        }
        //private static async Task UploadFileAsync(IAmazonS3 s3Client, FileInfo file, string bucketName)
        //{
        //    try
        //    {
        //        var fileTransferUtility =
        //            new TransferUtility(s3Client);

        //        // Option 1. Upload a file. The file name is used as the object key name.
        //        await fileTransferUtility.UploadAsync(file.FullName, bucketName);
        //        string message = "Upload complete.";

                
        //    }
        //    catch (AmazonS3Exception e)
        //    {
        //        string message = e.Message;
        //    }
        //    catch (Exception e)
        //    {
        //        string message = e.Message;
        //    }

        //}
        private void EmailZippedReport(XDocument xdoc, FileInfo file, RPParameters rpParameters, RPEmployer rpEmployer)
        {
            int logOneIn = Convert.ToInt32(xdoc.Root.Element("LogOneIn").Value);
            string configDirName = xdoc.Root.Element("SoftwareHomeFolder").Value;
            string textLine = null;

            try
            {
                //
                // Send an email.
                //
                bool validEmailAddress = false;
                //Find amount due to HMRC in the file name.
                int x = file.FullName.LastIndexOf('[');
                int y = file.FullName.LastIndexOf(']');
                string hmrcDesc = null;
                if(rpEmployer.P32Required)
                {
                    hmrcDesc = file.FullName.Substring(x, y - x);
                    hmrcDesc = hmrcDesc.Replace("[", "£");
                }
                
                DateTime runDate = rpParameters.PayRunDate;
                runDate = runDate.AddMonths(1);
                int day = runDate.Day;
                day = 20 - day;
                runDate = runDate.AddDays(day);
                string dueDate = runDate.ToLongDateString();
                string taxYear = rpParameters.TaxYear.ToString() + "/" + (rpParameters.TaxYear + 1).ToString().Substring(2, 2);
                //string emailPassword = "fjbykfgxxkdgclfp"; //fjbykfgxxkdgclfp
                string mailSubject = String.Format("Payroll reports for tax year {0}, pay period {1}.", taxYear, rpParameters.TaxPeriod);
                string mailBody = null;
                
                // Get currrent day of week.
                DayOfWeek today = DateTime.Today.DayOfWeek;
                string dataSource = xdoc.Root.Element("DataSource").Value;            //"APPSERVER1\\MSSQL";  //"13.69.154.210\\MSSQL";  
                string dataBase = xdoc.Root.Element("Database").Value;
                string userID = xdoc.Root.Element("Username").Value;
                string password = xdoc.Root.Element("Password").Value;
                string sqlConnectionString = "Server=" + dataSource + ";Database=" + dataBase + ";User ID=" + userID + ";Password=" + password + ";";
                //
                //Get the SMTP email settings from the database
                //
                SMTPEmailSettings smtpEmailSettings = new SMTPEmailSettings();
                smtpEmailSettings = GetEmailSettings(xdoc, sqlConnectionString);
                //
                //Get a list of email addresses to send the reports to
                //
                List<ContactInfo> contactInfoList = new List<ContactInfo>();
                contactInfoList = GetListOfContactInfo(xdoc, sqlConnectionString, rpParameters);
                foreach (ContactInfo contactInfo in contactInfoList)
                {
                    RegexUtilities regexUtilities = new RegexUtilities();
                    validEmailAddress = regexUtilities.IsValidEmail(contactInfo.EmailAddress);
                    if (validEmailAddress)
                    {
                        mailBody = String.Format("Hi {0},\r\n\r\nPlease find attached payroll reports for tax year {1}, pay period {2}.\r\n\r\n"
                                                 , contactInfo.FirstName, taxYear, rpParameters.TaxPeriod);
                        if(rpEmployer.P32Required)
                        {
                            mailBody = mailBody + string.Format("The amount payable to HMRC this month is {0}, this payment is due on or before {1}.\r\n\r\n"
                                                                 , hmrcDesc, dueDate);
                        }
                        mailBody = mailBody + string.Format("Please review and confirm if all is correct.\r\n\r\nKind Regards,\r\n\r\nThe Payescape Team.");
                        MailMessage mailMessage = new MailMessage();
                        mailMessage.To.Add(new MailAddress(contactInfo.EmailAddress));
                        mailMessage.From = new MailAddress(smtpEmailSettings.FromAddress);
                        //mailMessage.From = new MailAddress("jcborland@jbsoftwareservices.onmicrosoft.com");
                        mailMessage.Subject = mailSubject;
                        mailMessage.Body = mailBody;
                        //mailMessage.Attachments.Add(new Attachment(file.FullName));
                        using (Attachment attachment = new Attachment(file.FullName))
                        {
                            mailMessage.Attachments.Add(attachment);

                            //emailPassword = "@LI20sserluss16:";

                            SmtpClient smtpClient = new SmtpClient();
                            smtpClient.UseDefaultCredentials = smtpEmailSettings.SMTPUserDefaultCredentials;
                            smtpClient.Credentials = new System.Net.NetworkCredential(smtpEmailSettings.SMTPUsername, smtpEmailSettings.SMTPPassword);

                            //smtpClient.Credentials = new System.Net.NetworkCredential("jcborland@jbsoftwareservices.onmicrosoft.com", "JB20soft14");
                            smtpClient.Port = smtpEmailSettings.SMTPPort;
                            smtpClient.Host = smtpEmailSettings.SMTPHost;
                            //smtpClient.Host = "outlook-emeawest4.office365.com";
                            smtpClient.EnableSsl = smtpEmailSettings.SMTPEnableSSL;

                            bool emailSent = false;
                            try
                            {
                                textLine = string.Format("Attempting sending an email to, {0} from {1} with password:{2}, port:{3}, host:{4}.", contactInfo.EmailAddress,
                                                          smtpEmailSettings.SMTPUsername, smtpEmailSettings.SMTPPassword, smtpEmailSettings.SMTPPort, smtpEmailSettings.SMTPHost);
                                update_Progress(textLine, configDirName, logOneIn);

                                smtpClient.Send(mailMessage);
                                emailSent = true;


                            }
                            catch (Exception ex)
                            {
                                textLine = string.Format("Error sending an email to, {0}.\r\n{1}.\r\n", contactInfo.EmailAddress, ex);
                                update_Progress(textLine, configDirName, logOneIn);
                            }

                            if (emailSent)
                            {


                            }
                            else
                            {

                            }
                        }



                    }
                }

            }
            catch (Exception ex)
            {
                textLine = string.Format("Error sending email for file, {0}.\r\n{1}.\r\n", file.FullName, ex);
                update_Progress(textLine, configDirName, logOneIn);
            }
            finally
            {

            }

        }
        private SMTPEmailSettings GetEmailSettings(XDocument xdoc, string sqlConnectionString)
        {
            int x = sqlConnectionString.LastIndexOf(";Password=") + 10;
            int y = sqlConnectionString.LastIndexOf(";");
            string logConnectionString = sqlConnectionString.Substring(0, x + 2) + "*********" + sqlConnectionString.Substring(y - 2);


            int logOneIn = Convert.ToInt32(xdoc.Root.Element("LogOneIn").Value);
            string configDirName = xdoc.Root.Element("SoftwareHomeFolder").Value;
            string textLine = null;
            bool success = false;
            SMTPEmailSettings smtpEmailSettings = new SMTPEmailSettings();
            DataTable dtSMTPEmailSettings = new DataTable();
            //
            //Try using a stored procedure
            //
            try
            {
                using (var connection = new SqlConnection(sqlConnectionString))
                using (var command = new SqlCommand("SelectSMTPEmailSettings", connection)
                {
                    CommandType = CommandType.StoredProcedure
                })
                {
                    connection.Open();
                    SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(command);
                    sqlDataAdapter.Fill(dtSMTPEmailSettings);
                }
                success = true;
            }
            catch (Exception ex)
            {
                textLine = string.Format("Error getting email settings with SQL connection string, {0}.\r\n{1}.\r\n", logConnectionString, ex);
                update_Progress(textLine, configDirName, logOneIn);
            }
            if (success)
            {
                //
                //There should only be one record.
                //
                DataRow drSMTPEmailSettings;
                drSMTPEmailSettings = dtSMTPEmailSettings.Rows[0];
                smtpEmailSettings.Body = null;                  //I'm not using this yet. May never use it.
                smtpEmailSettings.FromAddress = drSMTPEmailSettings.ItemArray[0].ToString();
                smtpEmailSettings.SMTPEnableSSL = Convert.ToBoolean(drSMTPEmailSettings.ItemArray[6]);
                smtpEmailSettings.SMTPHost = drSMTPEmailSettings.ItemArray[2].ToString();
                smtpEmailSettings.SMTPPassword = drSMTPEmailSettings.ItemArray[4].ToString();
                smtpEmailSettings.SMTPPort = Convert.ToInt32(drSMTPEmailSettings.ItemArray[5]);
                smtpEmailSettings.SMTPUserDefaultCredentials = Convert.ToBoolean(drSMTPEmailSettings.ItemArray[1]);
                smtpEmailSettings.SMTPUsername = drSMTPEmailSettings.ItemArray[3].ToString();
                smtpEmailSettings.Subject = null;               //I'm not using this yet. May never use it.

                string softwareHomeFolder = xdoc.Root.Element("SoftwareHomeFolder").Value;

                textLine = string.Format("Getting SMTP email settings with connection string : {0}.", logConnectionString);
                update_Progress(textLine, configDirName, logOneIn);

                textLine = string.Format("Got SMTP email settings, host is : {0}.", smtpEmailSettings.SMTPHost);
                update_Progress(textLine, configDirName, logOneIn);

            }

            return smtpEmailSettings;

        }
        private List<ContactInfo> GetListOfContactInfo(XDocument xdoc, string sqlConnectionString, RPParameters rpParameters)
        {
            int logOneIn = Convert.ToInt32(xdoc.Root.Element("LogOneIn").Value);
            string configDirName = xdoc.Root.Element("SoftwareHomeFolder").Value;
            string textLine = null;

            int x = sqlConnectionString.LastIndexOf(";Password=") + 10;
            int y = sqlConnectionString.LastIndexOf(";");
            string logConnectionString = sqlConnectionString.Substring(0, x + 2) + "*********" + sqlConnectionString.Substring(y - 2);

            textLine = string.Format("Start getting a list of email addresses with connection string : {0}.", logConnectionString);
            update_Progress(textLine, configDirName, logOneIn);

            List<ContactInfo> contactInfoList = new List<ContactInfo>();
            string companyNo = rpParameters.ErRef;                  //file.FullName.Substring(0, 4);
            DataTable dtContactInfo = new DataTable();
            try
            {
                using (var connection = new SqlConnection(sqlConnectionString))
                using (var command = new SqlCommand("SelectPayrollReportsContacts", connection)
                {
                    CommandType = CommandType.StoredProcedure
                })
                {
                    connection.Open();
                    command.Parameters.AddWithValue("CompanyNo", companyNo);
                    SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(command);
                    sqlDataAdapter.Fill(dtContactInfo);
                }
            }
            catch (Exception ex)
            {
                textLine = string.Format("Error getting the list of email addresses.\r\n{0}.\r\n", ex);
                update_Progress(textLine, configDirName, logOneIn);
            }
            foreach (DataRow drContactInfo in dtContactInfo.Rows)
            {
                ContactInfo contactInfo = new ContactInfo();
                contactInfo.EmailAddress = drContactInfo.ItemArray[0].ToString();
                contactInfo.FirstName = drContactInfo.ItemArray[1].ToString();
                contactInfoList.Add(contactInfo);
            }

            textLine = string.Format("Finished getting a list of email addresses with connection string : {0}.", logConnectionString);
            update_Progress(textLine, configDirName, logOneIn);

            return contactInfoList;
        }
        private DataRow GetCompanyReportCodes(XDocument xdoc, string sqlConnectionString, RPParameters rpParameters)
        {
            int logOneIn = Convert.ToInt32(xdoc.Root.Element("LogOneIn").Value);
            string configDirName = xdoc.Root.Element("SoftwareHomeFolder").Value;
            string textLine = null;

            int x = sqlConnectionString.LastIndexOf(";Password=") + 10;
            int y = sqlConnectionString.LastIndexOf(";");
            string logConnectionString = sqlConnectionString.Substring(0, x + 2) + "*********" + sqlConnectionString.Substring(y - 2);

            textLine = string.Format("Start getting the company report codes with connection string : {0}.", logConnectionString);
            update_Progress(textLine, configDirName, logOneIn);

            string companyNo = rpParameters.ErRef;                  //file.FullName.Substring(0, 4);
            DataTable dtCompanyReportCodes = new DataTable();
            try
            {
                using (var connection = new SqlConnection(sqlConnectionString))
                using (var command = new SqlCommand("SelectCompanyReportCodes", connection)
                {
                    CommandType = CommandType.StoredProcedure
                })
                {
                    connection.Open();
                    command.Parameters.AddWithValue("CompanyNo", companyNo);
                    SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(command);
                    sqlDataAdapter.Fill(dtCompanyReportCodes);
                }
            }
            catch (Exception ex)
            {
                textLine = string.Format("Error getting the company report codes.\r\n{0}.\r\n", ex);
                update_Progress(textLine, configDirName, logOneIn);
            }
            
            DataRow drCompanyReportCodes = dtCompanyReportCodes.Rows[0];

            textLine = string.Format("Finished getting company report codes with connection string : {0}.", logConnectionString);
            update_Progress(textLine, configDirName, logOneIn);

            return drCompanyReportCodes;
        }
        public bool GetIsUnity(XDocument xdoc, string sqlConnectionString, int companyNo)
        {
            bool isUnity = false;
            int logOneIn = Convert.ToInt32(xdoc.Root.Element("LogOneIn").Value);
            string configDirName = xdoc.Root.Element("SoftwareHomeFolder").Value;
            string textLine = null;

            int x = sqlConnectionString.LastIndexOf(";Password=") + 10;
            int y = sqlConnectionString.LastIndexOf(";");
            string logConnectionString = sqlConnectionString.Substring(0, x + 2) + "*********" + sqlConnectionString.Substring(y - 2);

            textLine = string.Format("Start getting IsUnity with connection string : {0}.", logConnectionString);
            update_Progress(textLine, configDirName, logOneIn);

            DataTable dtCompanyReportCodes = new DataTable();
            try
            {
                using (var connection = new SqlConnection(sqlConnectionString))
                using (var command = new SqlCommand("SelectIsUnity", connection)
                {
                    CommandType = CommandType.StoredProcedure
                })
                {
                    connection.Open();
                    command.Parameters.AddWithValue("CompanyNo", companyNo);
                    SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(command);
                    sqlDataAdapter.Fill(dtCompanyReportCodes);
                }
            }
            catch (Exception ex)
            {
                textLine = string.Format("Error getting IsUnity.\r\n{0}.\r\n", ex);
                update_Progress(textLine, configDirName, logOneIn);
            }

            isUnity = Convert.ToBoolean(dtCompanyReportCodes.Rows[0].ItemArray[0]);
            
            textLine = string.Format("Finished getting IsUnity with connection string : {0}.", logConnectionString);
            update_Progress(textLine, configDirName, logOneIn);

            return isUnity;
        }

    }
    public class ReadConfigFile
    {
        //
        // Using XDocument instead of XmlReader
        //
        string fileName = "PayescapeWGtoPR.xml";
        XDocument xdoc = new XDocument();

        public ReadConfigFile() { }


        public XDocument ConfigRecord(string dirName)
        {
            string fullName = dirName + fileName;
            try
            {
                xdoc = XDocument.Load(fullName);
            }
            catch
            {
            }

            return xdoc;
        }
        
    }

    //Report (RP) Parameters
    public class RPParameters
    {
        public string ErRef { get; set; }
        public int TaxYear { get; set; }
        public DateTime AccYearStart { get; set; }
        public DateTime AccYearEnd { get; set; }
        public int TaxPeriod { get; set; }
        public string PaySchedule { get; set; }
        public DateTime PayRunDate { get; set; }


        public RPParameters() { }
        public RPParameters(string erRef, int taxYear, DateTime accYearStart,
                            DateTime accYearEnd, int taxPeriod, string paySchedule, DateTime payRundate)
        {
            ErRef = erRef;
            TaxYear = taxYear;
            AccYearStart = accYearStart;
            AccYearEnd = accYearEnd;
            TaxPeriod = taxPeriod;
            PaySchedule = paySchedule;
            PayRunDate = payRundate;
        }
    }
    //Report (RP) Employer
    public class RPEmployer
    {
        public string Name { get; set; }
        public string PayeRef { get; set; }
        public string HMRCDesc { get; set; }
        public string BankFileCode { get; set; }
        public string PensionReportCode { get; set; }
        public bool P32Required { get; set; }

        public RPEmployer() { }
        public RPEmployer(string name, string payeRef, string hmrcDesc,
                           string bankFileCode, string pensionReportCode,
                           bool p32Required)
        {
            Name = name;
            PayeRef = payeRef;
            HMRCDesc = hmrcDesc;
            BankFileCode = bankFileCode;
            PensionReportCode = pensionReportCode;
            P32Required = p32Required;
        }
    }

    //Report (RP) Employee
    public class RPEmployeePeriod
    {
        public string Reference { get; set; }
        public string Title { get; set; }
        public string Forename { get; set; }
        public string Surname { get; set; }
        public string Fullname { get; set; }
        public string RefFullname { get; set; }
        public string SurnameForename { get; set; }
        public string Address1 { get; set; }
        public string Address2 { get; set; }
        public string Address3 { get; set; }
        public string Address4 { get; set; }
        public string Postcode { get; set; }
        public string Country { get; set; }
        public string SortCode { get; set; }
        public string BankAccNo { get; set; }
        public DateTime DateOfBirth { get; set; }
        public string Gender { get; set; }
        public string BuildingSocRef { get; set; }
        public string NINumber { get; set; }
        public string PaymentMethod { get; set; }
        public DateTime PayRunDate { get; set; }
        public DateTime PeriodStartDate { get; set; }
        public DateTime PeriodEndDate { get; set; }
        public int PayrollYear { get; set; }
        public decimal Gross { get; set; }
        public decimal NetPayTP { get; set; }
        public decimal DayHours { get; set; }
        public DateTime? StudentLoanStartDate { get; set; }
        public DateTime? StudentLoanEndDate { get; set; }
        public decimal StudentLoan { get; set; }
        public string NILetter { get; set; }
        public string CalculationBasis { get; set; }
        public decimal Total { get; set; }
        public decimal EarningsToLEL { get; set; }
        public decimal EarningsToSET { get; set; }
        public decimal EarningsToPET { get; set; }
        public decimal EarningsToUST { get; set; }
        public decimal EarningsToAUST { get; set; }
        public decimal EarningsToUEL { get; set; }
        public decimal EarningsAboveUEL { get; set; }
        public decimal EeContributionsPt1 { get; set; }
        public decimal EeContributionsPt2 { get; set; }
        public decimal ErNICYTD { get; set; }
        public decimal EeRebate { get; set; }
        public decimal ErRebate { get; set; }
        public decimal EeReduction { get; set; }
        public DateTime? LeavingDate { get; set; }
        public bool Leaver { get; set; }
        public string TaxCode { get; set; }
        public bool Week1Month1 { get; set; }
        public string TaxCodeChangeTypeID { get; set; }
        public string TaxCodeChangeType { get; set; }
        public decimal TaxPrev { get; set; }
        public decimal TaxablePayPrevious { get; set; }
        public decimal TaxThis { get; set; }
        public decimal TaxablePayYTD { get; set; }
        public decimal TaxablePayTP { get; set; }
        public decimal HolidayAccruedTd { get; set; }
        public RPAEAssessment AEAssessment { get; set; }
        public List<RPPensionPeriod> Pensions { get; set; }
        public decimal ErPensionTotalTP { get; set; }
        public decimal ErPensionTotalYtd {get;set;}
        public DateTime DirectorshipAppointmentDate { get; set; }
        public bool Director { get; set; }
        public decimal EeContributionsTaxPeriodPt1 { get; set; }
        public decimal EeContributionsTaxPeriodPt2 { get; set; }
        public decimal ErNICTP { get; set; }
        public string Frequency { get; set; }
        public decimal NetPayYTD { get; set; }
        public decimal TotalPayTP { get; set; }
        public decimal TotalPayYTD { get; set; }
        public decimal TotalDedTP { get; set; }
        public decimal TotalDedYTD { get; set; }
        public string PensionCode { get; set; }
        public decimal PreTaxAddDed { get; set; }
        public decimal GUCosts { get; set; }
        public decimal AbsencePay { get; set; }
        public decimal HolidayPay { get; set; }
        public decimal PreTaxPension { get; set; }
        public decimal Tax { get; set; }
        public decimal NetNI { get; set; }
        public decimal PostTaxAddDed { get; set; }
        public decimal PostTaxPension { get; set; }
        public decimal AOE { get; set; }
        public List<RPAddition> Additions { get; set; }
        public List<RPDeduction> Deductions { get; set; }
        public List<RPPayslipDeduction> PayslipDeductions { get; set; }
        public RPEmployeePeriod() { }
        public RPEmployeePeriod(string reference, string title, string forename, string surname, string fullname, string refFullname, string surnameForename,
                          string address1, string address2, string address3, string address4, string postcode,
                          string country, string sortCode, string bankAccNo, DateTime dateOfBirth, string gender, string buildingSocRef,
                          string niNumber, string paymentMethod, DateTime payRunDate, DateTime periodStartDate, DateTime periodEndDate, int payrollYear,
                          decimal gross, decimal netPayTP, decimal dayHours, DateTime? studentLoanStartDate, DateTime? studentLoanEndDate,
                          decimal studentLoan, string niLetter, string calculationBasis, decimal total,
                          decimal earningsToLEL, decimal earningsToSET, decimal earningsToPET, decimal earningsToUST, decimal earningsToAUST,
                          decimal earningsToUEL, decimal earningsAboveUEL, decimal eeContributionsPt1, decimal eeContributionsPt2,
                          decimal erNICYTD, decimal eeRebate, decimal erRebate, decimal eeReduction, DateTime leavingDate, bool leaver,
                          string taxCode, bool week1Month1, string taxCodeChangeTypeID, string taxCodeChangeType, decimal taxPrev,
                          decimal taxablePayPrevious, decimal taxThis, decimal taxablePayYTD, decimal taxablePayTP, decimal holidayAccruedTd,
                          RPAEAssessment aeAssessment,
                          List<RPPensionPeriod> pensions,
                          decimal erPensionTotalTP, decimal erPensionTotalYtd,
                          //decimal erPensionYTD, decimal eePensionYTD, decimal erPensionTP, decimal eePensionTP, decimal erContributionPercent,
                          //decimal eeContributionPercent, decimal pensionablePay, DateTime erPensionPayRunDate, DateTime eePensionPayRunDate,
                          DateTime directorshipAppointmentDate, bool director, decimal eeContributionsTaxPeriodPt1, decimal eeContributionsTaxPeriodPt2,
                          decimal erNICTP, string frequency, decimal netPayYTD, decimal totalPayTP, decimal totalPayYTD, decimal totalDedTP, 
                          decimal totalDedYTD, string pensionCode, decimal preTaxAddDed, decimal guCosts, decimal absencePay,
                          decimal holidayPay, decimal preTaxPension, decimal tax, decimal netNI,
                          decimal postTaxAddDed, decimal postTaxPension, decimal aoe, 
                          List<RPAddition> additions, List<RPDeduction> deductions, List<RPPayslipDeduction> payslipDeductions)
        {
            Reference = reference;
            Title = title;
            Forename = forename;
            Surname = surname;
            Fullname = fullname;
            RefFullname = refFullname;
            SurnameForename = surnameForename;
            Address1 = address1;
            Address2 = address2;
            Address3 = address3;
            Address4 = address4;
            Postcode = postcode;
            Country = country;
            SortCode = sortCode;
            BankAccNo = bankAccNo;
            DateOfBirth = dateOfBirth;
            Gender = gender;
            BuildingSocRef = buildingSocRef;
            NINumber = niNumber;
            PaymentMethod = paymentMethod;
            PayRunDate = payRunDate;
            PeriodStartDate = periodStartDate;
            PeriodEndDate = periodEndDate;
            PayrollYear = payrollYear;
            Gross = gross;
            NetPayTP = netPayTP;
            DayHours = dayHours;
            StudentLoanStartDate = studentLoanStartDate;
            StudentLoanEndDate = studentLoanEndDate;
            StudentLoan = studentLoan;
            NILetter = niLetter;
            CalculationBasis = calculationBasis;
            Total = total;
            EarningsToLEL = earningsToLEL;
            EarningsToSET = earningsToSET;
            EarningsToPET = earningsToPET;
            EarningsToUST = earningsToUST;
            EarningsToAUST = earningsToAUST;
            EarningsToUEL = earningsToUEL;
            EarningsAboveUEL = earningsAboveUEL;
            EeContributionsPt1 = eeContributionsPt1;
            EeContributionsPt2 = eeContributionsPt2;
            ErNICYTD = erNICYTD;
            EeRebate = eeRebate;
            ErRebate = erRebate;
            EeReduction = eeReduction;
            LeavingDate = leavingDate;
            Leaver = leaver;
            TaxCode = taxCode;
            Week1Month1 = week1Month1;
            TaxCodeChangeTypeID = taxCodeChangeTypeID;
            TaxCodeChangeType = taxCodeChangeType;
            TaxPrev = taxPrev;
            TaxablePayPrevious = taxablePayPrevious;
            TaxThis = taxThis;
            TaxablePayYTD = taxablePayYTD;
            TaxablePayTP = taxablePayTP;
            HolidayAccruedTd = holidayAccruedTd;
            AEAssessment = aeAssessment;
            Pensions = pensions;
            ErPensionTotalTP = erPensionTotalTP;
            ErPensionTotalYtd = erPensionTotalYtd;
            DirectorshipAppointmentDate = directorshipAppointmentDate;
            Director = director;
            EeContributionsTaxPeriodPt1 = eeContributionsTaxPeriodPt1;
            EeContributionsTaxPeriodPt2 = eeContributionsTaxPeriodPt2;
            ErNICTP = erNICTP;
            Frequency = frequency;
            NetPayYTD = netPayYTD;
            TotalPayTP = totalPayTP;
            TotalPayYTD = totalPayYTD;
            TotalDedTP = totalDedTP;
            TotalDedYTD = totalDedYTD;
            PensionCode = pensionCode;
            PreTaxAddDed = preTaxAddDed;
            GUCosts = guCosts;
            AbsencePay = absencePay;
            HolidayPay = holidayPay;
            PreTaxPension = preTaxPension;
            Tax = tax;
            NetNI = netNI;
            PostTaxAddDed = postTaxAddDed;
            PostTaxPension = postTaxPension;
            AOE = aoe;
            Additions = additions;
            Deductions = deductions;
            PayslipDeductions = payslipDeductions;
        }

    }
    public class RPEmployeeYtd
    {
        public DateTime ThisPeriodStartDate { get; set; }
        public DateTime LastPaymentDate { get; set; }
        public string EeRef { get; set; }
        public string Branch { get; set; }
        public string CostCentre { get; set; }
        public string Department { get; set; }
        public DateTime? LeavingDate { get; set; }
        public bool Leaver { get; set; }
        public decimal TaxPrevEmployment { get; set; }
        public decimal TaxablePayPrevEmployment { get; set; }
        public decimal TaxThisEmployment { get; set; }
        public decimal TaxablePayThisEmployment { get; set; }
        public decimal GrossedUp { get; set; }
        public decimal GrossedUpTax { get; set; }
        public decimal NetPayYTD { get; set; }
        public decimal GrossPayYTD { get; set; }
        public decimal BenefitInKindYTD { get; set; }
        public decimal SuperannuationYTD { get; set; }
        public decimal HolidayPayYTD { get; set; }
        public decimal PensionablePayYtd { get; set; }
        public decimal EePensionYtd { get; set; }
        public decimal ErPensionYtd { get; set; }
        public decimal PensionPreTaxEeAccounts { get; set; }
        public decimal PensionPreTaxEePaye { get; set; }
        public decimal PensionPostTaxEeAccounts { get; set; }
        public decimal PensionPostTaxEePaye { get; set; }
        public List<RPPensionYtd> Pensions { get; set; }
        public decimal AeoYTD { get; set; }
        public DateTime? StudentLoanStartDate { get; set; }
        public DateTime? StudentLoanEndDate { get; set; }
        public string StudentLoanPlanType { get; set; }
        public decimal StudentLoanDeductionsYTD { get; set; }
        public DateTime? PostgraduateLoanStartDate { get; set; }
        public DateTime? PostgraduateLoanEndDate { get; set; }
        public decimal PostgraduateLoanDeductionsYTD { get; set; }
        public RPNicYtd NicYtd { get; set; }
        public RPNicAccountingPeriod NicAccountingPeriod { get; set; }
        public string TaxCode { get; set; }
        public bool Week1Month1 { get; set; }
        public int WeekNumber { get; set; }
        public int MonthNumber { get; set; }
        public int PeriodNumber { get; set; }
        public decimal EeNiPaidByErAccountsAmount { get; set; }
        public decimal EeNiPaidByErAccountsUnits { get; set; }
        public decimal EeGuTaxPaidByErAccountsAmount { get; set; }
        public decimal EeGuTaxPaidByErAccountsUnits { get; set; }
        public decimal EeNiLERtoUERAccountsAmount { get; set; }
        public decimal EeNiLERtoUERAccountsUnits { get; set; }
        public decimal ErNiAccountsAmount { get; set; }
        public decimal ErNiAccountsUnits { get; set; }
        public decimal EeNiLERtoUERPayeAmount { get; set; }
        public decimal EeNiLERtoUERPayeUnits { get; set; }
        public decimal EeNiPaidByErPayeAmount { get; set; }
        public decimal EeNiPaidByErPayeUnits { get; set; }
        public decimal EeGuTaxPaidByErPayeAmount { get; set; }
        public decimal EeGuTaxPaidByErPayeUnits { get; set; }
        public decimal ErNiPayeAmount { get; set; }
        public decimal ErNiPayeUnits { get; set; }
        public List<RPPayCode> PayCodes { get; set; }
        public RPEmployeeYtd() { }
        public RPEmployeeYtd(DateTime thisPeriodStartDate, DateTime lastPaymentDate, string eeRef, string branch, string costCentre, string department,
                          DateTime? leavingDate, bool leaver, decimal taxPrevEmployment,
                          decimal taxablePayPrevEmployment, decimal taxThisEmployemnt, decimal taxablePayThisEmployment, decimal grossedUp, decimal grossedUpTax,
                          decimal netPayYTD, decimal grossPayYTD, decimal benefitInKindYTD, decimal superannuationYTD, decimal holidayPayYTD,
                          decimal pensionablePayYtd, decimal eePensionYtd, decimal erPensionYtd, decimal pensionPreTaxEeAccounts, decimal pensionPreTaxEePaye,
                          decimal pensionPostTaxEeAccounts, decimal pensionPostTaxEePaye, List<RPPensionYtd> pensions,
                          decimal aeoYTD, DateTime? studentLoanStartDate, DateTime? studentLoanEndDate,
                          string studentLoanPlanType, decimal studentLoanDeductionsYTD, DateTime? postgraduateLoanStartDate, DateTime? postgraduateLoanEndDate,
                          decimal postgraduateLoanDeductionsYTD,
                          RPNicYtd nicYtd, RPNicAccountingPeriod nicAccountingPeriod,
                          decimal eeReduction, 
                          string taxCode, bool week1Month1, int weekNumber, int monthNumber, int periodNumber,
                          decimal eeNiPaidByErAccountsAmount, decimal eeNiPaidByErAccountsUnits, decimal eeGuTaxPaidByErAccountsAmount, decimal eeGuTaxPaidByErAccountsUnits,
                          decimal eeNiLERtoUERAccountsAmount, decimal eeNiLERtoUERAccountsUnits, decimal eeNiLERtoUERPayeAmount, decimal eeNiLERtoUERPayeUnits,
                          decimal erNiAccountsAmount, decimal erNiAccountsUnits, decimal erNiLERtoUERPayeAmount, decimal erNiLERtoUERPayeUnits, decimal eeNiPaidByErPayeAmount,
                          decimal eeNiPaidByErPayeUnits, decimal eeGuTaxPaidByErPayeAmount, decimal eeGuTaxPaidByErPayeUnits, decimal erNiPayeAmount, decimal erNiPayeUnits,
                          List<RPPayCode> payCodes)
                          
        {
            ThisPeriodStartDate = thisPeriodStartDate;
            LastPaymentDate = lastPaymentDate;
            EeRef = eeRef;
            Branch = Branch;
            CostCentre = CostCentre;
            Department = Department;
            LeavingDate = leavingDate;
            Leaver = leaver;
            TaxPrevEmployment = taxPrevEmployment;
            TaxablePayPrevEmployment = taxablePayPrevEmployment;
            TaxThisEmployment = taxThisEmployemnt;
            TaxablePayThisEmployment = taxablePayThisEmployment;
            GrossedUp = grossedUp;
            GrossedUpTax = grossedUpTax;
            NetPayYTD = netPayYTD;
            GrossPayYTD = grossPayYTD;
            BenefitInKindYTD = benefitInKindYTD;
            SuperannuationYTD = superannuationYTD;
            HolidayPayYTD = holidayPayYTD;
            PensionablePayYtd = pensionablePayYtd;
            EePensionYtd = eePensionYtd;
            ErPensionYtd = erPensionYtd;
            PensionPreTaxEeAccounts = pensionPreTaxEeAccounts;
            PensionPreTaxEePaye = pensionPreTaxEePaye;
            PensionPostTaxEeAccounts = pensionPostTaxEeAccounts;
            PensionPostTaxEePaye = pensionPostTaxEePaye;
            Pensions = pensions;
            AeoYTD = aeoYTD;
            StudentLoanStartDate = studentLoanStartDate;
            StudentLoanEndDate = studentLoanEndDate;
            StudentLoanPlanType = studentLoanPlanType;
            StudentLoanDeductionsYTD = studentLoanDeductionsYTD;
            PostgraduateLoanStartDate = postgraduateLoanStartDate;
            PostgraduateLoanEndDate = postgraduateLoanEndDate;
            PostgraduateLoanDeductionsYTD = postgraduateLoanDeductionsYTD;
            NicYtd = nicYtd;
            NicAccountingPeriod = nicAccountingPeriod;
            TaxCode = taxCode;
            Week1Month1 = week1Month1;
            WeekNumber = weekNumber;
            MonthNumber = monthNumber;
            PeriodNumber = periodNumber;
            EeNiPaidByErAccountsAmount = eeNiPaidByErAccountsAmount;
            EeNiPaidByErAccountsUnits = eeNiPaidByErAccountsUnits;
            EeGuTaxPaidByErAccountsAmount = eeGuTaxPaidByErAccountsAmount;
            EeGuTaxPaidByErAccountsUnits = eeGuTaxPaidByErAccountsUnits;
            EeNiLERtoUERAccountsAmount = eeNiLERtoUERAccountsAmount;
            EeNiLERtoUERAccountsUnits = eeNiLERtoUERAccountsUnits;
            ErNiAccountsAmount = erNiAccountsAmount;
            ErNiAccountsUnits = erNiAccountsUnits;
            EeNiLERtoUERPayeAmount = eeNiLERtoUERPayeAmount;
            EeNiLERtoUERPayeUnits = eeNiLERtoUERPayeUnits;
            EeNiPaidByErPayeAmount = eeNiPaidByErPayeAmount;
            EeNiPaidByErPayeUnits = eeNiPaidByErPayeUnits;
            EeGuTaxPaidByErPayeAmount = eeGuTaxPaidByErPayeAmount;
            EeGuTaxPaidByErPayeUnits = eeGuTaxPaidByErPayeUnits;
            ErNiPayeAmount = erNiPayeAmount;
            ErNiPayeUnits = erNiPayeUnits;
            PayCodes = payCodes;
        }

    }
    public class RPPensionPeriod
    {
        public int Key { get; set; }
        public string Code { get; set; }
        public string SchemeName { get; set; }
        public DateTime? StartJoinDate { get; set; }
        public bool IsJoiner { get; set; }
        public string ProviderEmployerReference { get; set; }
        public decimal EePensionYtd { get; set; }
        public decimal ErPensionYtd { get; set; }
        public decimal PensionablePayYtd { get; set; }
        public decimal EePensionTaxPeriod { get; set; }
        public decimal ErPensionTaxPeriod { get; set; }
        public decimal PensionablePayTaxPeriod { get; set; }
        public decimal EePensionPayRunDate { get; set; }
        public decimal ErPensionPayRunDate { get; set; }
        public decimal PensionablePayPayRunDate { get; set; }
        public decimal EeContibutionPercent { get; set; }
        public decimal ErContributionPercent { get; set; }
        
        public RPPensionPeriod() { }
        public RPPensionPeriod(int key, string code, string schemeName, DateTime? startJoinDate, bool isJoiner,
                               string providerEmployerReference,
                               decimal eePensionYtd, decimal erPensionYtd,
                               decimal pensionablePayYtd, decimal eePensionTaxPeriod, decimal erPensionTaxPeriod,
                               decimal pensionPayTaxPeriod, decimal eePensionPayRunDate, decimal erPensionPayRunDate,
                               decimal pensionablePayPayRunDate, decimal eeContributionPercent,
                               decimal erContributionPercent)
        {
            Key = key;
            Code = code;
            SchemeName = schemeName;
            StartJoinDate = startJoinDate;
            IsJoiner = isJoiner;
            ProviderEmployerReference = providerEmployerReference;
            EePensionYtd = eePensionYtd;
            ErPensionYtd = erPensionYtd;
            PensionablePayYtd = pensionablePayYtd;
            EePensionTaxPeriod = eePensionTaxPeriod;
            ErPensionTaxPeriod = erPensionTaxPeriod;
            PensionablePayTaxPeriod = pensionPayTaxPeriod;
            EePensionPayRunDate = eePensionPayRunDate;
            ErPensionPayRunDate = erPensionPayRunDate;
            PensionablePayPayRunDate = pensionablePayPayRunDate;
            EeContibutionPercent = eeContributionPercent;
            ErContributionPercent = erContributionPercent;
        }
    }
    
    public class RPPensionYtd
    {
        public int Key { get; set; }
        public string Code { get; set; }
        public string SchemeName { get; set; }
        public decimal PensionablePayYtd { get; set; }
        public decimal EePensionYtd { get; set; }
        public decimal ErPensionYtd { get; set; }
       public RPPensionYtd() { }
        public RPPensionYtd(int key, string code, string schemeName, 
                            decimal pensionablePayYtd, decimal eePensionYtd, decimal erPensionYtd)
        {
            Key=key;
            Code = code;
            SchemeName=schemeName;
            PensionablePayYtd = pensionablePayYtd;
            EePensionYtd=eePensionYtd;
            ErPensionYtd=erPensionYtd;
        }
    }
    public class RPAddress
    {
        public string Line1 { get; set; }
        public string Line2 { get; set; }
        public string Line3 { get; set; }
        public string Line4 { get; set; }
        public string Postcode { get; set; }
        public string Country { get; set; }
        
        public RPAddress() { }
        public RPAddress(string line1, string line2, string line3, string line4,
                         string postcode, string country)
        {
            Line1 = line1;
            Line2 = line2;
            Line3 = line3;
            Line4 = line4;
            Postcode = postcode;
            Country = country;
        }
    }
    public class RPPensionContribution
    {
        public string EeRef { get; set; }
        public string Title { get; set; }
        public string Forename { get; set; }
        public string Surname { get; set; } 
        public string Fullname { get; set; }
        public string SurnameForename { get; set; }
        public string ForenameSurname { get; set; }
        public DateTime DOB { get; set; }
        public RPAddress RPAddress { get; set; }
        public string EmailAddress { get; set; }
        public string Gender { get; set; }
        public string NINumber { get; set; }
        public string Freq { get; set; }
        public DateTime StartDate { get; set; }
        public DateTime EndDate { get; set; }
        public DateTime PayRunDate { get; set; }
        public RPPensionPeriod RPPensionPeriod { get; set; }

        public RPPensionContribution() { }
        public RPPensionContribution(string eeRef, string title, string forename,
                                     string surname, string fullname, string surnameForename, 
                                     string forenameSurname, DateTime dob, RPAddress rpAddress,
                                     string emailAddress, string gender,
                                     string niNumber, string freq,
                                     DateTime startDate, DateTime endDate,DateTime payRunDate,
                                     RPPensionPeriod rpPensionPeriod)
        {
            EeRef = eeRef;
            Title = title;
            Forename = forename;
            Surname = surname;
            Fullname = fullname;
            SurnameForename = surnameForename;
            ForenameSurname = forenameSurname;
            DOB = dob;
            RPAddress = rpAddress;
            EmailAddress = emailAddress;
            Gender = gender;
            NINumber = niNumber;
            Freq = freq;
            StartDate = startDate;
            EndDate = endDate;
            PayRunDate = payRunDate;
            RPPensionPeriod = rpPensionPeriod;
        }
    }
    public class RPPensionFileScheme
    {
        public string SchemeName { get; set; }
        public string SchemeProvider { get; set; }
        public List<RPPensionContribution> RPPensionContributions { get; set; }

        public RPPensionFileScheme() { }
        public RPPensionFileScheme(string schemeName, string schemeProvider,
                                   List<RPPensionContribution> rpPensionContributions)
        {
            SchemeName = schemeName;
            SchemeProvider = schemeProvider;
            RPPensionContributions = rpPensionContributions;
        }
    }

    public class RPNicYtd
    {
        public string NILetter { get; set; }
        public decimal NiableYtd { get; set; }
        public decimal EarningsToLEL { get; set; }
        public decimal EarningsToSET { get; set; }
        public decimal EarningsToPET { get; set; }
        public decimal EarningsToUST { get; set; }
        public decimal EarningsToAUST { get; set; }
        public decimal EarningsToUEL { get; set; }
        public decimal EarningsAboveUEL { get; set; }
        public decimal EeContributionsPt1 { get; set; }
        public decimal EeContributionsPt2 { get; set; }
        public decimal ErContributions { get; set; }
        public decimal EeRebate { get; set; }
        public decimal ErRebate { get; set; }
        public decimal EeReduction { get; set; }
        public decimal ErReduction { get; set; }
        public RPNicYtd() { }
        public RPNicYtd(string niLetter, decimal niableYtd, decimal earningsToLEL, decimal earningsToSET, decimal earningsToPET,
                        decimal earningsToUST, decimal earningsToAUST, decimal earningsToUEL, decimal earningsAboveUEL,
                        decimal eeContributionsPt1, decimal eeContributionsPt2, decimal erContributions, decimal eeRebate,
                        decimal erRebate, decimal eeReduction, decimal erReduction)
        {
            NILetter=niLetter;
            NiableYtd=niableYtd;
            EarningsToLEL=earningsToLEL;
            EarningsToSET=earningsToSET;
            EarningsToPET=earningsToPET;
            EarningsToUST = earningsToUST;
            EarningsToAUST = earningsToAUST;
            EarningsToUEL = earningsToUEL;
            EarningsAboveUEL = earningsAboveUEL;
            EeContributionsPt1 = eeContributionsPt1;
            EeContributionsPt2 = eeContributionsPt2;
            ErContributions = erContributions;
            EeRebate=eeRebate;
            ErRebate=erRebate;
            EeReduction=eeReduction;
            ErReduction = erReduction;
        }
    }
    public class RPNicAccountingPeriod
    {
        public string NILetter { get; set; }
        public decimal NiableYtd { get; set; }
        public decimal EarningsToLEL { get; set; }
        public decimal EarningsToSET { get; set; }
        public decimal EarningsToPET { get; set; }
        public decimal EarningsToUST { get; set; }
        public decimal EarningsToAUST { get; set; }
        public decimal EarningsToUEL { get; set; }
        public decimal EarningsAboveUEL { get; set; }
        public decimal EeContributionsPt1 { get; set; }
        public decimal EeContributionsPt2 { get; set; }
        public decimal ErContributions { get; set; }
        public decimal EeRebate { get; set; }
        public decimal ErRebate { get; set; }
        public decimal ErReduction { get; set; }
        public decimal EeReduction { get; set; }
        public RPNicAccountingPeriod() { }
        public RPNicAccountingPeriod(string niLetter, decimal niableYtd, decimal earningsToLEL, decimal earningsToSET, decimal earningsToPET,
                        decimal earningsToUST, decimal earningsToAUST, decimal earningsToUEL, decimal earningsAboveUEL,
                        decimal eeContributionsPt1, decimal eeContributionsPt2, decimal erContributions, decimal eeRebate,
                        decimal erRebate, decimal eeReduction, decimal erReduction)
        {
            NILetter = niLetter;
            NiableYtd = niableYtd;
            EarningsToLEL = earningsToLEL;
            EarningsToSET = earningsToSET;
            EarningsToPET = earningsToPET;
            EarningsToUST = earningsToUST;
            EarningsToAUST = earningsToAUST;
            EarningsToUEL = earningsToUEL;
            EarningsAboveUEL = earningsAboveUEL;
            EeContributionsPt1 = eeContributionsPt1;
            EeContributionsPt2 = eeContributionsPt2;
            ErContributions = erContributions;
            EeRebate = eeRebate;
            ErRebate = erRebate;
            EeReduction = eeReduction;
            ErReduction = erReduction;
        }
    }

    public class P45
    {
        public string ErOfficeNo { get; set; }
        public string ErRefNo { get; set; }
        public string NINumber { get; set; }
        public string Title { get; set; }
        public string Surname { get; set; }
        public string FirstNames { get; set; }
        public DateTime LeavingDate { get; set; }
        public bool StudentLoansDeductionToContinue { get; set; }
        public string TaxCode { get; set; }
        public bool Week1Month1 { get; set; }
        public int WeekNo { get; set; }
        public int MonthNo { get; set; }
        public decimal PayToDate { get; set; }
        public decimal TaxToDate { get; set; }
        public decimal PayThis { get; set; }
        public decimal TaxThis { get; set; }
        public string EeRef { get; set; }
        public bool IsMale { get; set; }
        public DateTime DateOfBirth { get; set; }
        public string Address1 { get; set; }
        public string Address2 { get; set; }
        public string Address3 { get; set; }
        public string Address4 { get; set; }
        public string Postcode { get; set; }
        public string Country { get; set; }
        public string ErName { get; set; }
        public string ErAddress1 { get; set; }
        public string ErAddress2 { get; set; }
        public string ErAddress3 { get; set; }
        public string ErAddress4 { get; set; }
        public string ErPostcode { get; set; }
        public string ErCountry { get; set; }
        public DateTime Now { get; set; }

        public P45() { }
        public P45(string erOfficeNo, string erRefNo, string niNumber, string title, string surname, string firstNames,
                    DateTime leavingDate,
                    bool studentLoansDedustionToContinue, string taxCode, int weekNo, int monthNo,
                    decimal payToDate, decimal taxToDate, decimal payThis, decimal taxThis, string eeRef, bool isMale,
                    string erName, string address1,
                    string address2, string address3, string address4, string postcode, string country,
                    DateTime dateOfBirth, string erAddress1,
                    string erAddress2, string erAddress3, string erAddress4, string erPostcode, string erCountry,
                    DateTime now)


        {
            ErOfficeNo = erOfficeNo;
            ErRefNo = erRefNo;
            NINumber = niNumber;
            Title = title;
            Surname = surname;
            FirstNames = firstNames;
            LeavingDate = leavingDate;
            StudentLoansDeductionToContinue = studentLoansDedustionToContinue;
            TaxCode = taxCode;
            WeekNo = weekNo;
            MonthNo = monthNo;
            PayToDate = payToDate;
            TaxToDate = taxToDate;
            PayThis = payThis;
            TaxThis = TaxThis;
            EeRef = eeRef;
            IsMale = isMale;
            DateOfBirth = dateOfBirth;
            Address1 = address1;
            Address2 = address2;
            Address3 = address3;
            Address4 = address4;
            Postcode = postcode;
            Country = country;
            ErName = erName;
            ErAddress1 = erAddress1;
            ErAddress2 = erAddress2;
            ErAddress3 = erAddress3;
            ErAddress4 = erAddress4;
            ErPostcode = erPostcode;
            ErCountry = erCountry;
            Now = now;
        }

    }

    //Report (RP) Additions
    public class RPAddition
    {
        public string EeRef { get; set; }

        public string Code { get; set; }
        public string Description { get; set; }
        public decimal Rate { get; set; }
        public decimal Units { get; set; }
        public decimal AmountTP { get; set; }
        public decimal AmountYTD { get; set; }
        public decimal AccountsYearBalance { get; set; }
        public decimal AccountsYearUnits { get; set; }
        public decimal PayeYearUnits { get; set; }
        public decimal PayrollAccrued { get; set; }
        public RPAddition() { }
        public RPAddition(string eeRef, string code, string description, decimal rate, decimal units,
                           decimal amountTP, decimal amountYTD, decimal accountsYearBalance,
                           decimal accountsYearUnits, decimal payeYearUnits, decimal payrollAccrued)
        {
            EeRef = eeRef;
            Code = code;
            Description = description;
            Rate = rate;
            Units = units;
            AmountTP = amountTP; //Amount
            AmountYTD = amountYTD; //PayeYearBalance
            AccountsYearBalance = accountsYearBalance;
            AccountsYearUnits = accountsYearUnits;
            PayeYearUnits = payeYearUnits;
            PayrollAccrued = payrollAccrued;
        }
    }

    //Report (RP) Deductions
    public class RPDeduction
    {
        public string EeRef { get; set; }
        public string Seq { get; set; }
        public string Code { get; set; }
        public string Description { get; set; }
        public bool IsTaxable { get; set; }
        public decimal Rate { get; set; }
        public decimal Units { get; set; }
        public decimal AmountTP { get; set; }
        public decimal AmountYTD { get; set; }
        public decimal AccountsYearBalance { get; set; }
        public decimal AccountsYearUnits { get; set; }
        public decimal PayeYearUnits { get; set; }
        public decimal PayrollAccrued { get; set; }
        public RPDeduction() { }
        public RPDeduction(string eeRef, string seq, string code, string description, bool isTaxable, decimal rate,
                           decimal units, decimal amountTP, decimal amountYTD, decimal accountsYearBalance,
                           decimal accountsYearUnits, decimal payeYearUnits, decimal payrollAccrued)
        {
            EeRef = eeRef;
            Seq = seq;
            Code = code;
            Description = description;
            IsTaxable = isTaxable;
            Rate = rate;
            Units = units;
            AmountTP = amountTP; //Amount
            AmountYTD = amountYTD; //PayeYearBalance
            AccountsYearBalance = accountsYearBalance;
            AccountsYearUnits = accountsYearUnits;
            PayeYearUnits = payeYearUnits;
            PayrollAccrued = payrollAccrued;
        }
    }
    //Payslip Report (RP) Deductions
    public class RPPayslipDeduction
    {
        public string EeRef { get; set; }
        public string Seq { get; set; }
        public string Code { get; set; }
        public string Description { get; set; }
        public decimal AmountTP { get; set; }
        public decimal AmountYTD { get; set; }
        public RPPayslipDeduction() { }
        public RPPayslipDeduction(string eeRef, string seq, string code, string description,
                                  decimal amountTP, decimal amountYTD)
        {
            EeRef = eeRef;
            Seq = seq;
            Code = code;
            Description = description;
            AmountTP = amountTP; //Amount
            AmountYTD = amountYTD; //PayeYearBalance
        }
    }

    //Report (RP) Pay Code
    public class RPPayCode
    {
        public string EeRef { get; set; }
        public string Code { get; set; }
        public string PayCode { get; set; }
        public string Description { get; set; }
        public string Type { get; set; }
        public decimal AccountsAmount{ get; set; }
        public decimal PayeAmount{ get; set; }
        public decimal AccountsUnits { get; set; }
        public decimal PayeUnits { get; set; }
        public RPPayCode() { }
        public RPPayCode(string eeRef, string code, string payCode, string description, string type,
                         decimal accountsAmount, decimal payeAmount, decimal accountsUnits, decimal payeUnits)
        {
            EeRef = eeRef;
            Code = code;
            PayCode = payCode;
            Description = description;
            Type=type;
            AccountsAmount = accountsAmount;
            PayeAmount = payeAmount;
            AccountsUnits = accountsUnits;
            PayeUnits = payeUnits;
        }
    }
    //Report (RP) PreSamplePayCode
    public class RPPreSamplePayCode
    {
        public string Code { get; set; }
        public string Description { get; set; }
        public bool InUse { get; set; }
        public RPPreSamplePayCode() { }
        public RPPreSamplePayCode(string code,string description, bool inUse)
        {
            Code = code;
            Description = description;
            InUse = inUse;
        }
    }
    public class RPPayComponent
    {
        public string PayCode { get; set; }
        public string Description { get; set; }
        public string EeRef { get; set; }
        public string Fullname { get; set; }
        public string Surname { get; set; }
        public string SurnameForename { get; set; }
        public decimal Rate { get; set; }
        public decimal UnitsTP { get; set; }
        public decimal AmountTP { get; set; }
        public decimal UnitsYTD { get; set; }
        public decimal AmountYTD { get; set; }
        public decimal AccountsYearBalance { get; set; }
        public decimal AccountsYearUnits { get; set; }
        public decimal PayrollAccrued { get; set; }
        public bool IsTaxable { get; set; }
        public bool IsPayCode { get; set; }
        public string EarningOrDeduction { get; set; }
        public RPPayComponent() { }
        public RPPayComponent(string payCode, string description, string eeRef, string fullname,
                              string surname, string surnameForename, decimal rate, decimal unitsTP, decimal amountTP,
                               decimal unitsYTD, decimal amountYTD, decimal accountsYearBalance, decimal accountsYearUnits,
                               decimal payrollAccrued, bool isTaxable, bool isPayCode,
                               string earningOrDeduction)
        {
            PayCode = payCode;
            Description = description;
            EeRef = eeRef;
            Fullname = fullname;
            Surname = surname;
            SurnameForename = surnameForename;
            Rate = rate;
            UnitsTP = unitsTP;
            AmountTP = amountTP;
            UnitsYTD = unitsYTD;
            AmountYTD = amountYTD;
            AccountsYearBalance = accountsYearBalance;
            AccountsYearUnits = accountsYearUnits;
            PayrollAccrued = payrollAccrued;
            IsTaxable = isTaxable;
            IsPayCode = isPayCode;
            EarningOrDeduction = earningOrDeduction;
        }
    }
    public class RPP32Report
    {
        public string EmployerName { get; set; }
        public string EmployerPayeRef { get; set; }
        public string PaymentRef { get; set; }
        public int TaxYear { get; set; }
        public DateTime TaxYearStartDate { get; set; }
        public DateTime TaxYearEndDate { get; set; }
        public int AnnualEmploymentAllowance { get; set; }
        public List<RPP32ReportMonth> RPP32ReportMonths { get; set; }
        public RPP32Report() { }
        public RPP32Report(string employerName, string employerPayeRef, string paymentRef,
                                  int taxYear, DateTime taxYearStartDate, DateTime taxYearEndDate,
                                  int annualEmploymentAllowance,
                                  List<RPP32ReportMonth> rpP32ReportMonths)
        {
            EmployerName = employerName;
            EmployerPayeRef = employerPayeRef;
            PaymentRef = paymentRef;
            TaxYear = taxYear;
            TaxYearStartDate = taxYearStartDate;
            TaxYearEndDate = taxYearEndDate;
            AnnualEmploymentAllowance = annualEmploymentAllowance;
            RPP32ReportMonths = rpP32ReportMonths;
        }
    }
    public class RPP32ReportMonth
    {
        public int PeriodNo { get; set; }
        public string RPPeriodNo { get; set; }
        public string RPPeriodText { get; set; }
        public string PeriodName { get; set; }
        public RPP32Breakdown RPP32Breakdown { get; set; }
        public RPP32Summary RPP32Summary { get; set; }
        public RPP32ReportMonth() { }
        public RPP32ReportMonth(int periodNo, string rpPeriodNo, 
                                string rpPeriodText, string periodName,
                                RPP32Breakdown rpP32Breakdown,
                                RPP32Summary rpP32Summary)
        {
            PeriodNo = periodNo;
            RPPeriodNo = rpPeriodNo;
            RPPeriodText = rpPeriodText;
            PeriodName = periodName;
            RPP32Breakdown = rpP32Breakdown;
            RPP32Summary = rpP32Summary;
        }
    }
    public class RPP32Summary
    {
        //Period 0 equals opening balance & period 13 equals annual total
        public decimal Tax { get; set; }
        public decimal StudentLoan { get; set; }
        public decimal PostGraduateLoan { get; set; }
        public decimal NetTax { get; set; }
        public decimal EmployerNI { get; set; }
        public decimal EmployeeNI { get; set; }
        public decimal GrossNICs { get; set; }
        public decimal SmpRecovered { get; set; }
        public decimal SmpComp { get; set; }
        public decimal SppRecovered { get; set; }
        public decimal SppComp { get; set; }
        public decimal ShppRecovered { get; set; }
        public decimal ShppComp { get; set; }
        public decimal SapRecovered { get; set; }
        public decimal SapComp { get; set; }
        public decimal TotalDeductions { get; set; }
        public decimal AppLevy { get; set; }
        public decimal CisDeducted { get; set; }
        public decimal CisSuffered { get; set; }
        public decimal NetNICs { get; set; }
        public decimal EmploymentAllowance { get; set; }
        public decimal AmountDue { get; set; }
        public decimal AmountPaid { get; set; }
        public decimal RemainingBalance { get; set; }
        public RPP32Summary() { }
        public RPP32Summary(decimal tax, decimal studentLoan, decimal postGraduateLoan, decimal netTax,
                            decimal employerNI, decimal employeeNI, decimal grossNICs, decimal smpRecovered,
                            decimal smpComp, decimal sppRecovered, decimal sppComp, decimal shppRecovered,
                            decimal shppComp, decimal sapRecovered, decimal sapComp, decimal totalDeductions,
                            decimal appLevy, decimal cisDeducted, decimal cisSuffered, decimal netNICs,
                            decimal employmentAllowance, decimal amountDue, decimal amountPaid, decimal remainingBalance)
        {
            Tax = tax;
            StudentLoan = studentLoan;
            PostGraduateLoan = postGraduateLoan;
            NetTax = netTax;
            EmployerNI = employerNI;
            EmployeeNI = employeeNI;
            GrossNICs = grossNICs;
            SmpRecovered = smpRecovered;
            SmpComp = smpComp;
            SppRecovered = sppRecovered;
            SppComp = sppComp;
            ShppRecovered = shppRecovered;
            ShppComp = shppComp;
            SapRecovered = sapRecovered;
            SapComp = sapComp;
            TotalDeductions = totalDeductions;
            AppLevy = appLevy;
            CisDeducted = cisDeducted;
            CisSuffered = cisSuffered;
            NetNICs = netNICs;
            EmploymentAllowance = employmentAllowance;
            AmountDue = amountDue;
            AmountPaid = amountPaid;
            RemainingBalance = remainingBalance;
        }
    }
    public class RPP32Breakdown
    {
        public List<RPP32Schedule> RPP32Schedules { get; set; }
        public RPP32Breakdown() { }
        public RPP32Breakdown(List<RPP32Schedule> rpP32Schedules)
        {
            RPP32Schedules = rpP32Schedules;
        }
    }
    public class RPP32Schedule
    {
        public string PayScheduleName { get; set; }
        public string PayScheduleFrequency { get; set; }
        public List<RPP32PayRun> RPP32PayRuns { get; set; }
        public RPP32Schedule() { }
        public RPP32Schedule(string payScheduleName, string payScheduleFrequency,
                              List<RPP32PayRun> rpP32PayRuns)
        {
            PayScheduleName = payScheduleName;
            PayScheduleFrequency = payScheduleFrequency;
            RPP32PayRuns = rpP32PayRuns;
        }
    }

    public class RPP32PayRun
    {
        public DateTime PayDate { get; set; }
        public int PayPeriod { get; set; }
        public decimal IncomeTax { get; set; }
        public decimal StudentLoan { get; set; }
        public decimal PostGraduateLoan { get; set; }
        public decimal NetIncomeTax { get; set; }
        public decimal GrossNICs { get; set; }
        public RPP32PayRun() { }
        public RPP32PayRun(DateTime payDate, int payPeriod, decimal incomeTax,
                           decimal studentLoan, decimal postGraduateLoan,
                           decimal netIncomeTax, decimal grossNICs)
        {
            PayDate = payDate;
            PayPeriod = payPeriod;
            IncomeTax = incomeTax;
            StudentLoan = studentLoan;
            PostGraduateLoan = postGraduateLoan;
            NetIncomeTax = netIncomeTax;
            GrossNICs = grossNICs;
        }
    }
    public class RPAEAssessment
    {
        public int Age { get; set; }
        public int StatePensionAge { get; set; }
        public DateTime? StatePensionDate { get; set; }
        public DateTime? AssessmentDate { get; set; }
        public decimal QualifyingEarnings { get; set; }
        public string AssessmentCode { get; set; }
        public string AssessmentEvent { get; set; }
        public string AssessmentResult { get; set; }
        public string AssessmentOverride { get; set; }
        public DateTime? OptOutWindowEndDate { get; set; }
        public DateTime? ReenrolmentDate { get; set; }
        public bool IsMemberOfAlternativePensionScheme { get; set; }
        public int TaxYear { get; set; }
        public int TaxPeriod { get; set; }
        public RPAEAssessment() { }
        public RPAEAssessment(int age, int statePensionAge, DateTime? statePensionDate, DateTime? assessmentDate, decimal qualifyingEarnings,
                              string assessmentCode, string assessmentEvent, string assessmentResult,
                              string assessmentOverride, DateTime? optOutWindowEndDate, DateTime? reenrolmentDate,
                              bool isMemberOfAlternativePensionScheme, int taxYear, int taxPeriod)
        {
            Age = age;
            StatePensionAge = statePensionAge;
            StatePensionDate = statePensionDate;
            AssessmentDate = assessmentDate;
            QualifyingEarnings = qualifyingEarnings;
            AssessmentCode = assessmentCode;
            AssessmentEvent = assessmentEvent;
            AssessmentResult = assessmentResult;
            AssessmentOverride = assessmentOverride;
            OptOutWindowEndDate = optOutWindowEndDate;
            ReenrolmentDate = reenrolmentDate;
            IsMemberOfAlternativePensionScheme = isMemberOfAlternativePensionScheme;
            TaxYear = taxYear;
            TaxPeriod = taxPeriod;
        }
    }
    public class ContactInfo
    {
        public string FirstName { get; set; }
        public string EmailAddress { get; set; }
        public ContactInfo() { }
        public ContactInfo(string firstName, string emailAddress)
        {
            FirstName = firstName;
            EmailAddress = emailAddress;
        }
    }
    public class SMTPEmailSettings
    {
        public string Subject { get; set; }
        public string Body { get; set; }
        public string FromAddress { get; set; }
        public bool SMTPUserDefaultCredentials { get; set; }
        public string SMTPUsername { get; set; }
        public string SMTPPassword { get; set; }
        public int SMTPPort { get; set; }
        public string SMTPHost { get; set; }
        public bool SMTPEnableSSL { get; set; }
        public SMTPEmailSettings() { }
        public SMTPEmailSettings(string subject, string body, string fromAddress, bool smtpUserDefaultCredentials,
                                 string smtpUsername, string smtpPassword, int smtpPort, string smtpHost,
                                 bool smtpEnableSSL)
        {
            Subject = subject;
            Body = body;
            FromAddress = fromAddress;
            SMTPUserDefaultCredentials = smtpUserDefaultCredentials;
            SMTPUsername = smtpUsername;
            SMTPPassword = smtpPassword;
            SMTPPort = smtpPort;
            SMTPHost = smtpHost;
            SMTPEnableSSL = smtpEnableSSL;
        }
    }
    public class RegexUtilities
    {
        bool invalid = false;

        public bool IsValidEmail(string strIn)
        {
            invalid = false;
            if (String.IsNullOrEmpty(strIn))
                return false;

            // Use IdnMapping class to convert Unicode domain names.
            strIn = Regex.Replace(strIn, @"(@)(.+)$", this.DomainMapper);
            if (invalid)
                return false;

            // Return true if strIn is in valid e-mail format.
            return Regex.IsMatch(strIn,
                   @"^(?("")(""[^""]+?""@)|(([0-9a-z]((\.(?!\.))|[-!#\$%&'\*\+/=\?\^`\{\}\|~\w])*)(?<=[0-9a-z])@))" +
                   @"(?(\[)(\[(\d{1,3}\.){3}\d{1,3}\])|(([0-9a-z][-\w]*[0-9a-z]*\.)+[a-z0-9]{2,17}))$",
                   RegexOptions.IgnoreCase);
        }
        public bool IsValidPostcode(string strIn)
        {
            invalid = false;
            if (String.IsNullOrEmpty(strIn))
                return false;

            // Return true if strIn is in valid Postcode format.
            return Regex.IsMatch(strIn,
                   "(^gir\\s?0aa$)|(^[a-z-[qvx]](\\d{1,2}|[a-hk-y]\\d{1,2}|\\d[a-hjks-uw]|[a-hk-y]\\d[abehmnprv-y])\\s?\\d[a-z-[cikmov]]{2}$)",
                   RegexOptions.IgnoreCase);
        }
        public bool IsValidNINumber(string strIn)
        {
            invalid = false;
            if (String.IsNullOrEmpty(strIn))
                return false;

            // Return true if strIn is in valid NI Number format.
            return Regex.IsMatch(strIn,
                   @"^([a-zA-Z]){2}( )?([0-9]){2}( )?([0-9]){2}( )?([0-9]){2}( )?([a-zA-Z]){1}?$",
                   RegexOptions.IgnoreCase);
        }
        private string DomainMapper(Match match)
        {
            // IdnMapping class with default property values.
            IdnMapping idn = new IdnMapping();

            string domainName = match.Groups[2].Value;
            try
            {
                domainName = idn.GetAscii(domainName);
            }
            catch (ArgumentException)
            {
                invalid = true;
            }
            return match.Groups[1].Value + domainName;

        }
    }
}
