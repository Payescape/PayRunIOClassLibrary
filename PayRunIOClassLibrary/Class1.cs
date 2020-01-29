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
using DevExpress.XtraReports.UI;
using PayRunIO.CSharp.SDK;
using PicoXLSX;

namespace PayRunIOClassLibrary
{
    public class PayRunIOWebGlobeClass
    {
        //Changed by Jim Borland on 29/1/2020 at 10:20 an a bit more.
        public PayRunIOWebGlobeClass() { }

        
        public void UpdateContactDetails(XDocument xdoc)
        {
            //Changes
            int j = 0;
            for (int i=0; i < 100; i++)
            {
                j = i;
            }
            
            string newString = "NewString";
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
            string consumerKey = "1UH6t3ikiWbdxTNT2Dg";                             //Original developer key : "m5lsJMpBnkaJw086zwDw"
            string consumerSecret = "jKUX3lrQUe4KhEiox6IZw8CXnWUdAkyTl1kthR8ayQ";   //Original developer secret : "GHM6x3xLEWujpLC5sGXKQ3r2j14RGI0eoLbab8w415Q"
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
        public void ArchiveDirectory(XDocument xdoc, string directory)
        {
            int logOneIn = Convert.ToInt32(xdoc.Root.Element("LogOneIn").Value);
            string configDirName = xdoc.Root.Element("SoftwareHomeFolder").Value;
            string textLine = null;

            try
            {
                DateTime now = DateTime.Now;

                int x = directory.LastIndexOf("\\");
                string coNo = directory.Substring(x + 1, 4);
                Directory.CreateDirectory(directory.Replace("Outputs", "PE-ArchivedOutputs"));
                DirectoryInfo dirInfo = new DirectoryInfo(directory);
                FileInfo[] files = dirInfo.GetFiles();
                foreach (FileInfo file in files)
                {
                    string destFileName = file.FullName.Replace("Outputs", "PE-ArchivedOutputs");
                    destFileName = destFileName.Replace(".xml", "_" + now.ToString("yyyyMMddHHmmssfff") + ".xml");
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
        public bool CheckIfP32Required(RPParameters rpParameters)
        {
            //Run the next period report to get the next pay period.
            string rptRef = "NEXTPERIOD";
            string parameter1 = "EmployerKey";
            string parameter2 = "PayScheduleKey";
            DateTime currentPayRunDate = DateTime.MinValue;
            DateTime nextPayRunDate = DateTime.MinValue;

            //Get the next period report
            XmlDocument xmlReport = RunReport(rptRef, parameter1, rpParameters.ErRef, parameter2, rpParameters.PaySchedule,
                                              null, null, null, null, null, null, null, null);

            foreach (XmlElement nextPayPeriod in xmlReport.GetElementsByTagName("NextPayPeriod"))
            {
                currentPayRunDate = Convert.ToDateTime(GetDateElementByTagFromXml(nextPayPeriod, "LastPayDay"));
                nextPayRunDate = Convert.ToDateTime(GetDateElementByTagFromXml(nextPayPeriod, "NextPayDay"));
            }
            bool p32Required = false;
            if (currentPayRunDate != DateTime.MinValue)
            {
                int currentTaxMonth = GetTaxMonth(currentPayRunDate);
                int nextTaxMonth = GetTaxMonth(nextPayRunDate);
                if (currentTaxMonth != nextTaxMonth)
                {
                    p32Required = true;
                }

            }


            return p32Required;
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
        public string[] GetAListOfDirectories(XDocument xdoc)
        {
            string path = xdoc.Root.Element("DataHomeFolder").Value + "Outputs";
            string[] directories = Directory.GetDirectories(path);

            return directories;
        }
        public FileInfo[] GetAllCompletedPayrollFiles(XDocument xdoc)
        {
            string path = xdoc.Root.Element("DataHomeFolder").Value + "Outputs";
            DirectoryInfo folder = new DirectoryInfo(path);
            FileInfo[] files = folder.GetFiles("*CompletedPayroll*.xml");
            
            return files;
        }
        
        
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
                    rpEmployeeYtd.ErPensionYTD = GetDecimalElementByTagFromXml(employee, "ErPensionYTD");
                    rpEmployeeYtd.EePensionYTD = GetDecimalElementByTagFromXml(employee, "EePensionYTD");
                    rpEmployeeYtd.AeoYTD = GetDecimalElementByTagFromXml(employee, "AeoYTD");
                    rpEmployeeYtd.StudentLoanStartDate = GetDateElementByTagFromXml(employee, "StudentLoanStartDate");
                    rpEmployeeYtd.StudentLoanEndDate = GetDateElementByTagFromXml(employee, "StudentLoanEndDate");
                    rpEmployeeYtd.StudentLoanDeductionsYTD = GetDecimalElementByTagFromXml(employee, "StudentLoanDeductionsYTD");
                    rpEmployeeYtd.NiLetter = GetElementByTagFromXml(employee, "NiLetter");
                    rpEmployeeYtd.NiableYTD = GetDecimalElementByTagFromXml(employee, "NiableYtd");
                    rpEmployeeYtd.EarningsToLEL = GetDecimalElementByTagFromXml(employee, "EarningsToLEL");
                    rpEmployeeYtd.EarningsToSET = GetDecimalElementByTagFromXml(employee, "EarningsToSET");
                    rpEmployeeYtd.EarningsToPET = GetDecimalElementByTagFromXml(employee, "EarningsToPET");
                    rpEmployeeYtd.EarningsToUST = GetDecimalElementByTagFromXml(employee, "EarningsToUST");
                    rpEmployeeYtd.EarningsToAUST = GetDecimalElementByTagFromXml(employee, "EarningsToAUST");
                    rpEmployeeYtd.EarningsToUEL = GetDecimalElementByTagFromXml(employee, "EarningsToUEL");
                    rpEmployeeYtd.EarningsAboveUEL = GetDecimalElementByTagFromXml(employee, "EarningsAboveUEL");
                    rpEmployeeYtd.EeContributionsPt1 = GetDecimalElementByTagFromXml(employee, "EeContributionsPt1");
                    rpEmployeeYtd.EeContributionsPt2 = GetDecimalElementByTagFromXml(employee, "EeContributionsPt2");
                    rpEmployeeYtd.ErContributions = GetDecimalElementByTagFromXml(employee, "ErContributions");
                    rpEmployeeYtd.EeRebate = GetDecimalElementByTagFromXml(employee, "EeRebate");
                    rpEmployeeYtd.ErRebate = GetDecimalElementByTagFromXml(employee, "ErRebate");
                    rpEmployeeYtd.ErReduction = GetDecimalElementByTagFromXml(employee, "ErReduction");
                    rpEmployeeYtd.EeReduction = GetDecimalElementByTagFromXml(employee, "EeReduction");
                    rpEmployeeYtd.TaxCode = GetElementByTagFromXml(employee, "TaxCode");
                    rpEmployeeYtd.Week1Month1 = GetBooleanElementByTagFromXml(employee, "Week1Month1");
                    rpEmployeeYtd.WeekNumber = GetIntElementByTagFromXml(employee, "WeekNumber");
                    rpEmployeeYtd.MonthNumber = GetIntElementByTagFromXml(employee, "MonthNumber");
                    rpEmployeeYtd.PeriodNumber = GetIntElementByTagFromXml(employee, "PeriodNumber");
                    rpEmployeeYtd.EeNiPaidByErAccountsAmount = GetDecimalElementByTagFromXml(employee, "EeNiPaidByErAccountsAmount");
                    rpEmployeeYtd.EeNiPaidByErAccountsUnits = GetDecimalElementByTagFromXml(employee, "EeNiPaidByErAccountsUnits");
                    rpEmployeeYtd.EeGuTaxPaidByErAccountsAmount = GetDecimalElementByTagFromXml(employee, "EeGuTaxPaidByErAccountsAmount");
                    rpEmployeeYtd.EeGuTaxPaidByErAccountsUnits = GetDecimalElementByTagFromXml(employee, "EeGuTaxPaidByErAccountsUnits");
                    rpEmployeeYtd.EeNiLERtoUERAccountsAmount = GetDecimalElementByTagFromXml(employee, "EeNiLERtoUERAccountsAmount");
                    rpEmployeeYtd.EeNiLERtoUERAccountsUnits = GetDecimalElementByTagFromXml(employee, "EeNiLERtoUERAccountsUnits");
                    rpEmployeeYtd.ErNiAccountsAmount = GetDecimalElementByTagFromXml(employee, "ErNiAccountsAmount");
                    rpEmployeeYtd.ErNiAccountsUnits = GetDecimalElementByTagFromXml(employee, "ErNiAccountsUnits");
                    rpEmployeeYtd.EeNiLERtoUERPayeAmount = GetDecimalElementByTagFromXml(employee, "EeNiLERtoUERPayeAmount");
                    rpEmployeeYtd.EeNiLERtoUERPayeUnits = GetDecimalElementByTagFromXml(employee, "EeNiLERtoUERPayeUnits");
                    rpEmployeeYtd.EeNiPaidByErPayeAmount = GetDecimalElementByTagFromXml(employee, "EeNiPaidByErPayeAmount");
                    rpEmployeeYtd.EeNiPaidByErPayeUnits = GetDecimalElementByTagFromXml(employee, "EeNiPaidByErPayeUnits");
                    rpEmployeeYtd.EeGuTaxPaidByErPayeAmount = GetDecimalElementByTagFromXml(employee, "EeGuTaxPaidByErPayeAmount");
                    rpEmployeeYtd.EeGuTaxPaidByErPayeUnits = GetDecimalElementByTagFromXml(employee, "EeGuTaxPaidByErPayeUnits");
                    rpEmployeeYtd.ErNiPayeAmount = GetDecimalElementByTagFromXml(employee, "ErNiPayeAmount");
                    rpEmployeeYtd.ErNiPayeUnits = GetDecimalElementByTagFromXml(employee, "ErNiPayeUnits");


                    //These next few fields get treated like pay codes. Use them if they are not zero.
                    //4 pay components EeNiPaidByEr, EeGuTaxPaidByEr, EeNiLERtoUER & ErNi
                    List<RPPayCode> rpPayCodeList = new List<RPPayCode>();

                    for (int i = 0; i < 6; i++)
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
                                rpPayCode.AccountsAmount = GetDecimalElementByTagFromXml(employee, "EeGuTaxPaidByErAccountsAmount");
                                rpPayCode.PayeAmount = GetDecimalElementByTagFromXml(employee, "EeGuTaxPaidByErPayeAmount");
                                rpPayCode.AccountsUnits = GetDecimalElementByTagFromXml(employee, "EeGuTaxPaidByErAccountsUnit");
                                rpPayCode.PayeUnits = GetDecimalElementByTagFromXml(employee, "EeGuTaxPaidByErPayeUnit");
                                break;
                            case 2:
                                rpPayCode.PayCode = "NIEeeLERtoUER";
                                rpPayCode.Description = "NIEeeLERtoUER-A";
                                rpPayCode.Type = "E";
                                rpPayCode.AccountsAmount = GetDecimalElementByTagFromXml(employee, "EeNiLERtoUERAccountsAmount");
                                rpPayCode.PayeAmount = GetDecimalElementByTagFromXml(employee, "EeNiLERtoUERPayeAmount");
                                rpPayCode.AccountsUnits = GetDecimalElementByTagFromXml(employee, "EeNiLERtoUERAccountsUnit");
                                rpPayCode.PayeUnits = GetDecimalElementByTagFromXml(employee, "EeNiLERtoUERPayeUnit");
                                break;
                            case 3:
                                rpPayCode.PayCode = "NIEr";
                                rpPayCode.Description = "NIEr-A";
                                rpPayCode.Type = "T";
                                rpPayCode.AccountsAmount = GetDecimalElementByTagFromXml(employee, "ErNiAccountAmount");
                                rpPayCode.PayeAmount = GetDecimalElementByTagFromXml(employee, "ErNiPayeAmount");
                                rpPayCode.AccountsUnits = GetDecimalElementByTagFromXml(employee, "ErNiAccountUnit");
                                rpPayCode.PayeUnits = GetDecimalElementByTagFromXml(employee, "ErNiPayeUnit");
                                break;
                            case 4:
                                rpPayCode.PayCode = "PenEr";
                                rpPayCode.Description = "PenEr";
                                rpPayCode.Type = "D";
                                rpPayCode.AccountsAmount = GetDecimalElementByTagFromXml(employee, "ErPensionYTD");
                                rpPayCode.PayeAmount = GetDecimalElementByTagFromXml(employee, "ErPensionYTD");
                                rpPayCode.AccountsUnits = 0;
                                rpPayCode.PayeUnits = 0;
                                break;
                            default:
                                rpPayCode.PayCode = "PenPostTaxEe";
                                rpPayCode.Description = "PenPostTaxEe";
                                rpPayCode.Type = "D";
                                rpPayCode.AccountsAmount = GetDecimalElementByTagFromXml(employee, "EePensionYTD");
                                rpPayCode.PayeAmount = GetDecimalElementByTagFromXml(employee, "EePensionYTD");
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
                                        rpPayCode.AccountsUnits = rpPayCode.AccountsUnits * -1;
                                        rpPayCode.PayeAmount = rpPayCode.PayeAmount * -1;
                                        rpPayCode.PayeUnits = rpPayCode.PayeUnits * -1;

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
                    payYTDDetails[15] = rpEmployeeYtd.ErPensionYTD.ToString();
                    payYTDDetails[16] = rpEmployeeYtd.EePensionYTD.ToString();
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
                    payYTDDetails[21] = rpEmployeeYtd.NiLetter;
                    payYTDDetails[22] = rpEmployeeYtd.NiableYTD.ToString();
                    payYTDDetails[23] = rpEmployeeYtd.EarningsToLEL.ToString();
                    payYTDDetails[24] = rpEmployeeYtd.EarningsToSET.ToString();
                    payYTDDetails[25] = rpEmployeeYtd.EarningsToPET.ToString();
                    payYTDDetails[26] = rpEmployeeYtd.EarningsToUST.ToString();
                    payYTDDetails[27] = rpEmployeeYtd.EarningsToAUST.ToString();
                    payYTDDetails[28] = rpEmployeeYtd.EarningsToUEL.ToString();
                    payYTDDetails[29] = rpEmployeeYtd.EarningsAboveUEL.ToString();
                    payYTDDetails[30] = rpEmployeeYtd.EeContributionsPt1.ToString();
                    payYTDDetails[31] = rpEmployeeYtd.EeContributionsPt2.ToString();
                    payYTDDetails[32] = rpEmployeeYtd.ErContributions.ToString();
                    payYTDDetails[33] = rpEmployeeYtd.EeRebate.ToString();
                    payYTDDetails[34] = rpEmployeeYtd.ErRebate.ToString();
                    payYTDDetails[35] = rpEmployeeYtd.EeReduction.ToString();
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
                    //payYTDDetails[40] = rpEmployeeYtd.PeriodNumber.ToString(); //PeriodNumber was here but never seemed to be used.
                    payYTDDetails[40] = rpEmployeeYtd.NiableYTD.ToString();
                    payYTDDetails[41] = ""; //Student Loan Plan Type
                    payYTDDetails[42] = ""; //Postgraduate Loan Start Date
                    payYTDDetails[43] = ""; //Postgraduate Loan End Date
                    payYTDDetails[44] = ""; //Postgraduate Loan Deducted

                    foreach (RPPayCode rpPayCode in rpEmployeeYtd.PayCodes)
                    {
                        //Don't use pay codes TAX, NI or any that begin with PENSION
                        if (rpPayCode.Code != "TAX" && rpPayCode.Code != "NI" && !rpPayCode.Code.StartsWith("PENSION"))
                        {
                            string[] payCodeDetails = new string[8];
                            payCodeDetails[0] = rpPayCode.Code;
                            payCodeDetails[1] = rpPayCode.Type;
                            payCodeDetails[2] = rpPayCode.PayCode;
                            payCodeDetails[3] = rpPayCode.Description;
                            payCodeDetails[4] = rpPayCode.AccountsAmount.ToString();
                            payCodeDetails[5] = rpPayCode.PayeAmount.ToString();
                            payCodeDetails[6] = rpPayCode.AccountsUnits.ToString();
                            payCodeDetails[7] = rpPayCode.PayeUnits.ToString();
                            if (rpPayCode.Type == "D" && rpPayCode.PayCode != "PenEr" && rpPayCode.PayCode != "PenPostTaxEe")
                            {
                                //Deduction so multiply by -1 unless it's "PenEr" or "PenPostTaxEe" which are already positive.
                                for (int i = 4; i < 8; i++)
                                {
                                    payCodeDetails[i] = (Convert.ToDecimal(payCodeDetails[i]) * -1).ToString();
                                }
                            }
                            if (payCodeDetails[0] == "UNPDM")
                            {
                                //Change UNPDM back to UNPD£. WG uses UNPD£ PR doesn't like symbols like £ in pay codes.
                                payCodeDetails[0] = "";// "UNPD£";
                                payCodeDetails[2] = "UNPD£";
                            }
                            else
                            {
                                payCodeDetails[0] = "";
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
                              "Taxable Pay Previous Emt, Tax This Emt, Taxable Pay This Emt,Grossed Up," +
                              "Grossed Up Tax, Net Pay,GrossYTD,Benefit in Kind,Superannuation," +
                              "Holiday Pay, ErPensionYTD, EePensionYTD, AEOYTD, StudentLoanStartDate," +
                              "StudentLoanEndDate, StudentLoanDeductions, NI Letter,Total," +
                              "Earnings To LEL,Earnings To SET,Earnings To PET,Earnings To UST," +
                              "Earnings To AUST,Earnings To UEL,Earnings Above UEL," +
                              "Ee Contributions Pt1,Ee Contributions Pt2,Er Contributions," +
                              "Ee Rebate,Er Rebate, Ee Reduction,PayCode,det,payCodeValue," +
                              "payCodeDesc,Acc Year Bal,PAYE Year Bal,Acc Year Units," +
                              "PAYE Year Units,Tax Code, Week1/ Month 1,Week Number, Month Number," +
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
                string[] payHistoryDetails = new string[49];

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
                        payHistoryDetails[10] = rpEmployeePeriod.StudentLoan.ToString();
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
                        payHistoryDetails[38] = rpEmployeePeriod.ErPensionYTD.ToString();
                        payHistoryDetails[39] = rpEmployeePeriod.EePensionYTD.ToString();
                        payHistoryDetails[40] = rpEmployeePeriod.ErPensionTP.ToString();
                        payHistoryDetails[41] = rpEmployeePeriod.EePensionTP.ToString();
                        payHistoryDetails[42] = rpEmployeePeriod.ErPensionPayRunDate.ToString("dd/MM/yy", CultureInfo.InvariantCulture);
                        payHistoryDetails[43] = rpEmployeePeriod.EePensionPayRunDate.ToString("dd/MM/yy", CultureInfo.InvariantCulture);
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
                                default:
                                    payCodeDetails[1] = "PenEr";
                                    payCodeDetails[2] = "PenEr";
                                    payCodeDetails[3] = "M";
                                    payCodeDetails[6] = rpEmployeePeriod.ErPensionTP.ToString();
                                    break;
                            }
                            payCodeDetails[0] = "0";
                            payCodeDetails[4] = "0";
                            payCodeDetails[5] = "0.00";
                            payCodeDetails[7] = "0.00";
                            payCodeDetails[8] = "0.00";
                            payCodeDetails[9] = "0.00";
                            payCodeDetails[10] = "0";
                            payCodeDetails[11] = "0";

                            //
                            //Check if any of the values are not zero. If so write the first employee record
                            //
                            bool allZeros = false;
                            if (payCodeDetails[5] == "0.00" && payCodeDetails[6] == "0.00" &&
                                payCodeDetails[7] == "0.00" && payCodeDetails[8] == "0.00" &&
                                payCodeDetails[9] == "0.00")
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
                            payCodeDetails[0] = rpAddition.Code.TrimStart(' ');
                            payCodeDetails[1] = rpAddition.Description;
                            payCodeDetails[2] = payCodeDetails[0];
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
                        foreach (RPDeduction rpDeduction in rpEmployeePeriod.Deductions)
                        {
                            string[] payCodeDetails = new string[12];
                            payCodeDetails = new string[12];
                            payCodeDetails[0] = rpDeduction.Code.TrimStart(' ');
                            payCodeDetails[1] = rpDeduction.Description;
                            payCodeDetails[2] = payCodeDetails[0];
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
                            switch (payCodeDetails[0]) //PayCode
                            {
                                case "TAX":
                                    payCodeDetails[0] = "0";
                                    payCodeDetails[1] = payHistoryDetails[29];  // Tax Code
                                    payCodeDetails[2] = payHistoryDetails[29];  // Tax Code
                                    payCodeDetails[3] = "T";                    // Tax    
                                    break;
                                case "NI":
                                    payCodeDetails[0] = "0";
                                    payCodeDetails[1] = "NIEeeLERtoUER-A";      // Ee NI
                                    payCodeDetails[2] = "NIEeeLERtoUER";        // Ee NI
                                    payCodeDetails[3] = "T";                    // Tax    
                                    break;
                                case "PENSION":
                                case "PENSIONSS":
                                case "PENSIONRAS":
                                case "PENSIONTAXEX":
                                    payCodeDetails[0] = "0";
                                    payCodeDetails[1] = "PenPostTaxEe";         // Ee Pension
                                    payCodeDetails[2] = "PenPostTaxEe";         // Ee Pension
                                    break;
                                default:
                                    payCodeDetails[0] = "";
                                    break;

                            }
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
                              "Acc Year Bal,PAYE Year Bal,Acc Year Units,PAYE Year Units,Payroll Accrued";
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


        
        public void ProcessBankReports(XDocument xdoc, List<RPEmployeePeriod> rpEmployeePeriodList, RPEmployer rpEmployer, RPParameters rpParameters)
        {
            string outgoingFolder = xdoc.Root.Element("DataHomeFolder").Value + "PE-Reports" + "\\" + rpParameters.ErRef;
            string dataSource = xdoc.Root.Element("DataSource").Value;            //"APPSERVER1\\MSSQL";  //"13.69.154.210\\MSSQL";  
            string dataBase = xdoc.Root.Element("Database").Value;
            string userID = xdoc.Root.Element("Username").Value;
            string password = xdoc.Root.Element("Password").Value;
            string sqlConnectionString = "Server=" + dataSource + ";Database=" + dataBase + ";User ID=" + userID + ";Password=" + password + ";";
            DataRow drCompanyReportCodes = GetCompanyReportCodes(xdoc, sqlConnectionString, rpParameters);
            rpEmployer.BankFileCode = drCompanyReportCodes.ItemArray[0].ToString();                 //BankFileCode
            rpEmployer.PensionReportCode = drCompanyReportCodes.ItemArray[1].ToString();            //PensionReportCode
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
        public void PrintStandardReports(XDocument xdoc, List<RPEmployeePeriod> rpEmployeePeriodList, RPEmployer rpEmployer, RPParameters rpParameters, List<P45> p45s, List<RPPayComponent> rpPayComponents)
        {
            PrintPayslips(xdoc, rpEmployeePeriodList, rpEmployer, rpParameters);
            PrintPayslipsSimple(xdoc, rpEmployeePeriodList, rpEmployer, rpParameters);
            PrintPaymentsDueByMethodReport(xdoc, rpEmployeePeriodList, rpEmployer, rpParameters);
            PrintComponentAnalysisReport(xdoc, rpPayComponents, rpEmployer, rpParameters);
            PrintPensionContributionsReport(xdoc, rpEmployeePeriodList, rpEmployer, rpParameters);
            PrintPayrollRunDetailsReport(xdoc, rpEmployeePeriodList, rpEmployer, rpParameters);
            //PrintPayrollNewReport(xdoc, rpEmployeePeriodList, rpEmployer, rpParameters);
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
        public decimal CalculateHMRCTotal(List<RPEmployeePeriod> rpEmployeePeriodList)
        {
            decimal hmrcTotal = 0;
            foreach (RPEmployeePeriod employee in rpEmployeePeriodList)
            {
                hmrcTotal = hmrcTotal + employee.Tax + employee.NetNI + employee.ErNICTP;
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
        private void PrintPensionContributionsReport(XDocument xdoc, List<RPEmployeePeriod> rpEmployeePeriodList, RPEmployer rpEmployer, RPParameters rpParameters)
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
                    EmailZippedReport(xdoc, file, rpParameters);
                    file.MoveTo(file.FullName.Replace("PE-Reports", "PE-Reports\\Archive"));
                }

            }
            catch (Exception ex)
            {
                textLine = string.Format("Error emailing zipped pdf reports for report folder, {0}.\r\n{1}.\r\n", reportFolder, ex);
                update_Progress(textLine, configDirName, logOneIn);
            }
        }
        private void EmailZippedReport(XDocument xdoc, FileInfo file, RPParameters rpParameters)
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
                string hmrcDesc = file.FullName.Substring(x, y - x);
                hmrcDesc = hmrcDesc.Replace("[", "£");
                DateTime runDate = rpParameters.PayRunDate;
                runDate = runDate.AddMonths(1);
                int day = runDate.Day;
                day = 20 - day;
                runDate = runDate.AddDays(day);
                string dueDate = runDate.ToLongDateString();
                //string emailPassword = "fjbykfgxxkdgclfp"; //fjbykfgxxkdgclfp
                string mailSubject = String.Format("Payroll reports for tax year {0}, pay period {1}.", rpParameters.TaxYear, rpParameters.TaxPeriod);

                string mailBody = String.Format("Hi,\r\n\r\nPlease find attached payroll reports for tax year {0}, pay period {1}.\r\n\r\n" +
                                                "The amount payable to HMRC this month is {2}, this payment is due on or before {3}.\r\n\r\n" +
                                                "Please review and confirm if all is correct.\r\n\r\nKind Regards,\r\n\r\nThe Payescape Team."
                                                , rpParameters.TaxYear, rpParameters.TaxPeriod, hmrcDesc, dueDate);
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
                List<string> emailAddresses = new List<string>();
                emailAddresses = GetListOfEmailAddresses(xdoc, sqlConnectionString, rpParameters);
                foreach (string emailAddress in emailAddresses)
                {
                    RegexUtilities regexUtilities = new RegexUtilities();
                    validEmailAddress = regexUtilities.IsValidEmail(emailAddress);
                    if (validEmailAddress)
                    {


                        MailMessage mailMessage = new MailMessage();
                        mailMessage.To.Add(new MailAddress(emailAddress));
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


                                smtpClient.Send(mailMessage);
                                emailSent = true;


                            }
                            catch (Exception ex)
                            {
                                textLine = string.Format("Error sending an email to, {0}.\r\n{1}.\r\n", emailAddress, ex);
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
        private List<string> GetListOfEmailAddresses(XDocument xdoc, string sqlConnectionString, RPParameters rpParameters)
        {
            int logOneIn = Convert.ToInt32(xdoc.Root.Element("LogOneIn").Value);
            string configDirName = xdoc.Root.Element("SoftwareHomeFolder").Value;
            string textLine = null;

            int x = sqlConnectionString.LastIndexOf(";Password=") + 10;
            int y = sqlConnectionString.LastIndexOf(";");
            string logConnectionString = sqlConnectionString.Substring(0, x + 2) + "*********" + sqlConnectionString.Substring(y - 2);

            textLine = string.Format("Start getting a list of email addresses with connection string : {0}.", logConnectionString);
            update_Progress(textLine, configDirName, logOneIn);

            List<string> emailAddresses = new List<string>();
            string companyNo = rpParameters.ErRef;                  //file.FullName.Substring(0, 4);
            DataTable dtEmailAddresses = new DataTable();
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
                    sqlDataAdapter.Fill(dtEmailAddresses);
                }
            }
            catch (Exception ex)
            {
                textLine = string.Format("Error getting the list of email addresses.\r\n{0}.\r\n", ex);
                update_Progress(textLine, configDirName, logOneIn);
            }
            foreach (DataRow drEmailAddresses in dtEmailAddresses.Rows)
            {
                emailAddresses.Add(drEmailAddresses.ItemArray[0].ToString());
            }

            textLine = string.Format("Finished getting a list of email addresses with connection string : {0}.", logConnectionString);
            update_Progress(textLine, configDirName, logOneIn);

            return emailAddresses;
        }
        private DataRow GetCompanyReportCodes(XDocument xdoc, string sqlConnectionString, RPParameters rpParameters)
        {
            int logOneIn = Convert.ToInt32(xdoc.Root.Element("LogOneIn").Value);
            string configDirName = xdoc.Root.Element("SoftwareHomeFolder").Value;
            string textLine = null;

            int x = sqlConnectionString.LastIndexOf(";Password=") + 10;
            int y = sqlConnectionString.LastIndexOf(";");
            string logConnectionString = sqlConnectionString.Substring(0, x + 2) + "*********" + sqlConnectionString.Substring(y - 2);

            textLine = string.Format("Start getting the company report codes connection string : {0}.", logConnectionString);
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
            //string[] companyReportCodes = new string[2];
            DataRow drCompanyReportCodes = dtCompanyReportCodes.Rows[0];
            //foreach (DataRow drCompanyReportCodes in dtCompanyReportCodes.Rows)
            //{
            //    //Should only be 1 row
            //    companyReportCodes[0] = drCompanyReportCodes.ItemArray[0].ToString();
            //    companyReportCodes[1] = drCompanyReportCodes.ItemArray[1].ToString();
                
            //}

            textLine = string.Format("Finished getting company report codes with connection string : {0}.", logConnectionString);
            update_Progress(textLine, configDirName, logOneIn);

            return drCompanyReportCodes;
        }
       
    }
    public class ReadConfigFile
    {
        //
        // Using XDocument instead of XmlReader
        //
        string fileName = "PayescapeWGtoPR.xml";
        string xmlSoftwareHomeFolder = "C:\\Payescape\\Service\\";
        string xmlDataHomeFolder = "C:\\Payescape\\Data\\";
        string xmlSFTPHostName = "sftp.bluemarblepayroll.com";
        string xmlUser = "payescape123";
        string xmlPasswordFile = "payescape.ppk";
        string xmlInterval = "10";
        string xmlLogOneIn = "100";
        string xmlOffFrom = "22:30:00";
        string xmlOffTo = "00:30:00";
        string xmlRunConstantly = "False";
        string xmlFilePrefix = "WGtoPR_";
        string xmlArchive = "True";
        string xmlDataSource = "APPSERVER1\\MSSQL";
        string xmlDatabase = "Payescape";
        string xmlUsername = "PayrollEngineLogin";
        string xmlPassword = "JB20soft14";
        XDocument xdoc = new XDocument();

        public ReadConfigFile() { }


        public XDocument ConfigRecord(string dirName)
        {
            string fullName = dirName + fileName;
            string passwordFile = dirName + xmlPasswordFile;

            try
            {
                bool updateRequired = false;
                bool exists = false;
                xdoc = XDocument.Load(fullName);
                exists = xdoc.Root.Descendants("SoftwareHomeFolder").Any();
                if (exists)
                {
                    xmlSoftwareHomeFolder = xdoc.Root.Element("SoftwareHomeFolder").Value;
                }
                else
                {
                    updateRequired = true;
                }
                exists = xdoc.Root.Descendants("DataHomeFolder").Any();
                if (exists)
                {
                    xmlDataHomeFolder = xdoc.Root.Element("DataHomeFolder").Value;
                }
                else
                {
                    updateRequired = true;
                }
                exists = xdoc.Root.Descendants("OffFrom").Any();
                if (exists)
                {
                    xmlOffFrom = xdoc.Root.Element("OffFrom").Value;
                }
                else
                {
                    updateRequired = true;
                }
                exists = xdoc.Root.Descendants("OffTo").Any();
                if (exists)
                {
                    xmlOffTo = xdoc.Root.Element("OffTo").Value;
                }
                else
                {
                    updateRequired = true;
                }
                exists = xdoc.Root.Descendants("RunConstantly").Any();
                if (exists)
                {
                    xmlRunConstantly = xdoc.Root.Element("RunConstantly").Value;
                }
                else
                {
                    updateRequired = true;
                }
                exists = xdoc.Root.Descendants("SFTPHostName").Any();
                if (exists)
                {
                    xmlSFTPHostName = xdoc.Root.Element("SFTPHostName").Value;
                }
                else
                {
                    updateRequired = true;
                }
                exists = xdoc.Root.Descendants("User").Any();
                if (exists)
                {
                    xmlUser = xdoc.Root.Element("User").Value;
                }
                else
                {
                    updateRequired = true;
                }
                exists = xdoc.Root.Descendants("PasswordFile").Any();
                if (exists)
                {
                    xmlPasswordFile = xdoc.Root.Element("PasswordFile").Value;
                }
                else
                {
                    updateRequired = true;
                }
                exists = xdoc.Root.Descendants("Interval").Any();
                if (exists)
                {
                    xmlInterval = xdoc.Root.Element("Interval").Value;
                }
                else
                {
                    updateRequired = true;
                }
                exists = xdoc.Root.Descendants("LogOneIn").Any();
                if (exists)
                {
                    xmlLogOneIn = xdoc.Root.Element("LogOneIn").Value;
                }
                else
                {
                    updateRequired = true;
                }
                exists = xdoc.Root.Descendants("FilePrefix").Any();
                if (exists)
                {
                    xmlFilePrefix = xdoc.Root.Element("FilePrefix").Value;
                }
                else
                {
                    updateRequired = true;
                }
                exists = xdoc.Root.Descendants("Archive").Any();
                if (exists)
                {
                    xmlArchive = xdoc.Root.Element("Archive").Value;
                }
                else
                {
                    updateRequired = true;
                }
                if (updateRequired)
                {
                    CreateConfigFile(dirName, fullName);
                }
                exists = xdoc.Root.Descendants("DataSource").Any();
                if (exists)
                {
                    xmlDataSource = xdoc.Root.Element("DataSource").Value;
                }
                else
                {
                    updateRequired = true;
                }
                if (updateRequired)
                {
                    CreateConfigFile(dirName, fullName);
                }
                exists = xdoc.Root.Descendants("Database").Any();
                if (exists)
                {
                    xmlDatabase = xdoc.Root.Element("Database").Value;
                }
                else
                {
                    updateRequired = true;
                }
                if (updateRequired)
                {
                    CreateConfigFile(dirName, fullName);
                }
                exists = xdoc.Root.Descendants("Username").Any();
                if (exists)
                {
                    xmlUsername = xdoc.Root.Element("Username").Value;
                }
                else
                {
                    updateRequired = true;
                }
                if (updateRequired)
                {
                    CreateConfigFile(dirName, fullName);
                }
                exists = xdoc.Root.Descendants("Password").Any();
                if (exists)
                {
                    xmlPassword = xdoc.Root.Element("Password").Value;
                }
                else
                {
                    updateRequired = true;
                }
                if (updateRequired)
                {
                    CreateConfigFile(dirName, fullName);
                }
            }

            catch (Exception ex)
            {

                if (ex.ToString().Contains("Could not find a part of the path") || ex.ToString().Contains("Could not find file"))
                {
                    CreateConfigFile(dirName, fullName);

                }
                xdoc = XDocument.Load(fullName);
            }

            return xdoc;
        }
        private void CreateConfigFile(string dirName, string fullName)
        {
            // Create Folder and dummy config xml file.
            Directory.CreateDirectory(dirName);


            // Create a dummy config xml file.
            new XDocument
                (
                new XElement
                    ("Configuration",
                     new XElement("SoftwareHomeFolder", xmlSoftwareHomeFolder),
                     new XElement("DataHomeFolder", xmlDataHomeFolder),
                     new XElement("Interval", xmlInterval),
                     new XElement("LogOneIn", xmlLogOneIn),
                     new XElement("RunConstantly", xmlRunConstantly),
                     new XElement("OffFrom", xmlOffFrom),
                     new XElement("OffTo", xmlOffTo),
                     new XElement("SFTPHostName", xmlSFTPHostName),
                     new XElement("User", xmlUser),
                     new XElement("PasswordFile", xmlPasswordFile),
                     new XElement("FilePrefix", xmlFilePrefix),
                     new XElement("Archive", xmlArchive),
                     new XElement("DataSource", xmlDataSource),
                     new XElement("Database", xmlDatabase),
                     new XElement("Username", xmlUsername),
                     new XElement("Password", xmlPassword)
                    )
                   )
             .Save(fullName);
            xdoc = XDocument.Load(fullName);
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

        public RPEmployer() { }
        public RPEmployer(string name, string payeRef, string hmrcDesc,
                           string bankFileCode, string pensionReportCode)
        {
            Name = name;
            PayeRef = payeRef;
            HMRCDesc = hmrcDesc;
            BankFileCode = bankFileCode;
            PensionReportCode = pensionReportCode;

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
        public decimal ErPensionYTD { get; set; }
        public decimal EePensionYTD { get; set; }
        public decimal ErPensionTP { get; set; }
        public decimal EePensionTP { get; set; }
        public decimal ErContributionPercent { get; set; }
        public decimal EeContributionPercent { get; set; }
        public decimal PensionablePay { get; set; }
        public DateTime ErPensionPayRunDate { get; set; }
        public DateTime EePensionPayRunDate { get; set; }
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
                          decimal erPensionYTD, decimal eePensionYTD, decimal erPensionTP, decimal eePensionTP, decimal erContributionPercent,
                          decimal eeContributionPercent, decimal pensionablePay, DateTime erPensionPayRunDate, DateTime eePensionPayRunDate,
                          DateTime directorshipAppointmentDate, bool director, decimal eeContributionsTaxPeriodPt1, decimal eeContributionsTaxPeriodPt2,
                          decimal erNICTP, string frequency, decimal netPayYTD, decimal totalPayTP, decimal totalPayYTD, decimal totalDedTP, 
                          decimal totalDedYTD, string pensionCode, decimal preTaxAddDed, decimal guCosts, decimal absencePay,
                          decimal holidayPay, decimal preTaxPension, decimal tax, decimal netNI,
                          decimal postTaxAddDed, decimal postTaxPension, decimal aoe, 
                          List<RPAddition> additions, List<RPDeduction> deductions)
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
            ErPensionYTD = erPensionYTD;
            EePensionYTD = eePensionYTD;
            ErPensionTP = erPensionTP;
            EePensionTP = eePensionTP;
            ErContributionPercent = erContributionPercent;
            EeContributionPercent = eeContributionPercent;
            PensionablePay = pensionablePay;
            ErPensionPayRunDate = erPensionPayRunDate;
            EePensionPayRunDate = eePensionPayRunDate;
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
        }

    }
    public class RPEmployeeYtd
    {
        public DateTime ThisPeriodStartDate { get; set; }
        public DateTime LastPaymentDate { get; set; }
        public string EeRef { get; set; }
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
        public decimal ErPensionYTD { get; set; }
        public decimal EePensionYTD { get; set; }
        public decimal AeoYTD { get; set; }
        public DateTime? StudentLoanStartDate { get; set; }
        public DateTime? StudentLoanEndDate { get; set; }
        public decimal StudentLoanDeductionsYTD { get; set; }
        public string NiLetter { get; set; }
        public decimal NiableYTD { get; set; }
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
        public RPEmployeeYtd(DateTime thisPeriodStartDate, DateTime lastPaymentDate, string eeRef, DateTime? leavingDate, bool leaver, decimal taxPrevEmployment,
                          decimal taxablePayPrevEmployment, decimal taxThisEmployemnt, decimal taxablePayThisEmployment, decimal grossedUp, decimal grossedUpTax,
                          decimal netPayYTD, decimal grossPayYTD, decimal benefitInKindYTD, decimal superannuationYTD, decimal holidayPayYTD,
                          decimal erPensionYTD, decimal eePensionYTD, decimal aeoYTD, DateTime? studentLoanStartDate, DateTime? studentLoanEndDate,
                          decimal studentLoanDeductionsYTD, string niLetter, decimal niableYTD, decimal earningsToLEL, decimal earningsToSET,
                          decimal earningsToPET, decimal earningsToUST, decimal earningsToAUST, decimal earningsToUEL, decimal earningsAboveUEL,
                          decimal eeContributionsPt1, decimal eeContributionsPt2, decimal erContributions, decimal eeRebate, decimal erRebate,
                          decimal erReduction, decimal eeReduction, string taxCode, bool week1Month1, int weekNumber, int monthNumber, int periodNumber,
                          decimal eeNiPaidByErAccountsAmount, decimal eeNiPaidByErAccountsUnits, decimal eeGuTaxPaidByErAccountsAmount, decimal eeGuTaxPaidByErAccountsUnits,
                          decimal eeNiLERtoUERAccountsAmount, decimal eeNiLERtoUERAccountsUnits, decimal eeNiLERtoUERPayeAmount, decimal eeNiLERtoUERPayeUnits,
                          decimal erNiAccountsAmount, decimal erNiAccountsUnits, decimal erNiLERtoUERPayeAmount, decimal erNiLERtoUERPayeUnits, decimal eeNiPaidByErPayeAmount,
                          decimal eeNiPaidByErPayeUnits, decimal eeGuTaxPaidByErPayeAmount, decimal eeGuTaxPaidByErPayeUnits, decimal erNiPayeAmount, decimal erNiPayeUnits,
                          List<RPPayCode> payCodes)
                          
        {
            ThisPeriodStartDate = thisPeriodStartDate;
            LastPaymentDate = lastPaymentDate;
            EeRef = eeRef;
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
            ErPensionYTD = erPensionYTD;
            EePensionYTD = eePensionYTD;
            AeoYTD = aeoYTD;
            StudentLoanStartDate = studentLoanStartDate;
            StudentLoanEndDate = studentLoanEndDate;
            StudentLoanDeductionsYTD = studentLoanDeductionsYTD;
            NiLetter = niLetter;
            NiableYTD = niableYTD;
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
            ErReduction = erReduction;
            EeReduction = eeReduction;
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
        public RPDeduction() { }
        public RPDeduction(string eeRef, string code, string description, decimal rate, decimal units,
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
        public bool IsTaxable { get; set; }
        public bool IsPayCode { get; set; }
        public string EarningOrDeduction { get; set; }
        public RPPayComponent() { }
        public RPPayComponent(string payCode, string description, string eeRef, string fullname,
                              string surname, string surnameForename, decimal rate, decimal unitsTP, decimal amountTP,
                               decimal unitsYTD, decimal amountYTD, bool isTaxable, bool isPayCode,
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
            IsTaxable = isTaxable;
            IsPayCode = isPayCode;
            EarningOrDeduction = earningOrDeduction;
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
