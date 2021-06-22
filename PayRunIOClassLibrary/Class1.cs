using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using System.IO;
using System.Data.SqlClient;
using System.Data;
using System.Text.RegularExpressions;
using System.Xml;
using Microsoft.VisualBasic.FileIO;
using System.Globalization;
using PayRunIO.CSharp.SDK;
using DevExpress.XtraReports.UI;
using PicoXLSX;

namespace PayRunIOClassLibrary
{
    public class PayRunIOWebGlobeClass
    {
        public PayRunIOWebGlobeClass() { }

        //Testing making a change to the class 
        // Contacts to move to Payrun.io meta data:https://payrun.atlassian.net/browse/PEINT-330
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
                    bool success = InsertDataIntoSQLServerUsingSQLBulkCopy(dtContacts, sqlConnectionString, xdoc);
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
        private bool InsertDataIntoSQLServerUsingSQLBulkCopy(DataTable csvDataTable, string sqlConnectionString, XDocument xdoc)
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
                            Update_Progress(textLine, configDirName);

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
                            Update_Progress(textLine, configDirName);
                        }
                        catch
                        {
                            textLine = string.Format("Failed to create tmpContacts table.");
                            Update_Progress(textLine, configDirName);

                            return false;

                        }
                    }
                    try
                    {
                        using (SqlBulkCopy bulkData = new SqlBulkCopy(sqlConnection))
                        {
                            textLine = string.Format("About to bulk write to tmpContacts table.");
                            Update_Progress(textLine, configDirName);

                            bulkData.DestinationTableName = tableName;

                            foreach (DataColumn dataColumn in csvDataTable.Columns)
                            {
                                dataColumn.ColumnName = Regex.Replace(dataColumn.ColumnName, "[^A-Za-z0-9]", "");
                                bulkData.ColumnMappings.Add(dataColumn.ToString(), dataColumn.ToString());

                            }
                            //bulkData.BulkCopyTimeout = 600; // 600 seconds
                            bulkData.WriteToServer(csvDataTable);

                            textLine = string.Format("Successfull bulk write to tmpContacts table.");
                            Update_Progress(textLine, configDirName);

                            return true;

                        }
                    }
                    catch
                    {
                        textLine = string.Format("Failed bulk write to tmpContacts table.");
                        Update_Progress(textLine, configDirName);

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
            string configDirName = xdoc.Root.Element("SoftwareHomeFolder").Value;
            string textLine;

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

                        DataColumn datacolumn = new DataColumn(column)
                        {
                            AllowDBNull = true
                        };
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
                                Update_Progress(textLine, configDirName);
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
                Update_Progress(textLine, configDirName);

            }
            return csvDataTable;
        }
        private bool InsertUpdateContacts(XDocument xdoc, string sqlConnectionString, string companyNo)
        {
            string configDirName = xdoc.Root.Element("SoftwareHomeFolder").Value;
            string textLine;

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
                Update_Progress(textLine, configDirName);
            }


            return success;
        }
        private void DeleteTemporaryContacts(XDocument xdoc, string sqlConnectionString)
        {
            string configDirName = xdoc.Root.Element("SoftwareHomeFolder").Value;
            string textLine;

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
                Update_Progress(textLine, configDirName);
            }


        }
        public void Update_Progress(string textLine, string configDirName)
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


        public XmlDocument RunReport(XDocument xdoc, string rptRef, string prm1, string val1, string prm2, string val2, string prm3, string val3,
                                 string prm4, string val4, string prm5, string val5, string prm6, string val6)
        {
            string url;
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
                var apiHelper = ApiHelper(xdoc);
                //string testurl = "EmployerKey=1958&TaxYear=2019&AccPeriodStart=2019-04-06&AccPeriodEnd=2020-04-05&TaxPeriod=27&PayScheduleKey=Weekly";
                //xmlReport = apiHelper.GetRawXml("/Report/" + rptRef + "/run?" + testurl);
                xmlReport = apiHelper.GetRawXml("/Report/" + rptRef + "/run?" + url);

            }
            catch
            {
                //MessageBox.Show("Error running a report.\r\n" + ex);
            }
            return xmlReport;
        }
        public XmlDocument GetPayRunIOObject(XDocument xdoc, string objectType, string erNo)
        {
            XmlDocument xmlObject = null;
            try
            {
                var apiHelper = ApiHelper(xdoc);
                xmlObject = apiHelper.GetRawXml("/" + objectType + "/" + erNo);

            }
            catch(Exception ex)
            {
                
            }
            return xmlObject;
        }
        public string RunTransformReport(XDocument xdoc, string rptRef, string prm1, string val1, string prm2, string val2, string prm3, string val3,
                                 string prm4, string val4, string prm5, string val5, string prm6, string val6)
        {
            string url;
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
            string csvReport = null;
            try
            {
                var apiHelper = ApiHelper(xdoc);
                csvReport = apiHelper.GetRawText("/Report/" + rptRef + "/run?" + url);
                
            }
            catch
            {
                //MessageBox.Show("Error running a report.\r\n" + ex);
            }
            return csvReport;
        }
        private RestApiHelper ApiHelper(XDocument xdoc)
        {
            string consumerKey = xdoc.Root.Element("PayRunConsumerKey").Value;
            string consumerSecret = xdoc.Root.Element("PayRunConsumerSecret").Value;
            string url = xdoc.Root.Element("PayRunUrl").Value;
            //Get reports from the live server when necessary
            //consumerKey = "OcMeDpYcUaacLPyztnMLA";
            //consumerSecret = "MpR3BR0OwkOmqgn9ZROWbglc55uEO6jEar58SPYpCkA";
            //url = "https://api.payrun.io";
            
            RestApiHelper apiHelper = new RestApiHelper(
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
            string configDirName = xdoc.Root.Element("SoftwareHomeFolder").Value;
            string textLine;

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
                    //destFileName = destFileName.Replace(".xml", "_" + now.ToString("yyyyMMddHHmmssfff") + ".xml");
                    //destFileName = destFileName.Replace(".csv", "_" + now.ToString("yyyyMMddHHmmssfff") + ".csv");
                    try
                    {
                        File.Move(file.FullName, destFileName);
                    }
                    catch
                    {
                        file.Delete();
                    }
                    

                }

                Directory.Delete(directory);
            }
            catch (Exception ex)
            {
                textLine = string.Format("Error archiving the Outputs directory, {0}.\r\n{1}.\r\n", directory, ex);
                Update_Progress(textLine, configDirName);
            }

        }
        public XmlDocument GetP32Report(XDocument xdoc, RPParameters rpParameters)
        {
            string rptRef = "P32";
            string parameter1 = "EmployerKey";
            string parameter2 = "TaxYear";
            
            //Get the P32 report
            XmlDocument xmlReport = RunReport(xdoc, rptRef, parameter1, rpParameters.ErRef, parameter2, rpParameters.TaxYear.ToString(),
                                              null, null, null, null, null, null, null, null);

            
            return xmlReport;
        }
        public XmlDocument GetApprenticeshipLevyReport(XDocument xdoc, RPParameters rpParameters)
        {
            string rptRef = "PE-AppLevyAnnual";
            string parameter1 = "EmployerKey";
            string parameter2 = "TaxYear";
            
            //Get the Note And Coin Requirement report
            XmlDocument xmlReport = RunReport(xdoc, rptRef, parameter1, rpParameters.ErRef,
                                              parameter2, rpParameters.TaxYear.ToString(),
                                              null, null, null, null, null, null, null, null);


            return xmlReport;
        }
        public XmlDocument GetNoteAndCoinRequirementReport(XDocument xdoc, RPParameters rpParameters)
        {
            string rptRef = "PSCOIN2";
            string parameter1 = "EmployerKey";
            string parameter2 = "PayScheduleKey";
            string parameter3 = "PaymentDate";

            //Get the Note And Coin Requirement report
            XmlDocument xmlReport = RunReport(xdoc, rptRef, parameter1, rpParameters.ErRef, 
                                              parameter2, rpParameters.PaySchedule,
                                              parameter3, rpParameters.PayRunDate.ToString("yyyy-MM-dd"),
                                              null, null, null, null, null, null);


            return xmlReport;
        }
        public XmlDocument GetCurrentAttachmentOfEarningsOrders(XDocument xdoc, RPParameters rpParameters)
        {
            string rptRef = "PSAEORUN";
            string parameter1 = "EmployerKey";
            string parameter2 = "PayScheduleKey";
            string parameter3 = "PaymentDate";

            //Get the current attachment of earnings orders
            XmlDocument xmlReport = RunReport(xdoc, rptRef, parameter1, rpParameters.ErRef,
                                              parameter2, rpParameters.PaySchedule,
                                              parameter3, rpParameters.PayRunDate.ToString("yyyy-MM-dd"),
                                              null, null, null, null, null, null);


            return xmlReport;
        }
        public XmlDocument GetBankFileReport(XDocument xdoc, RPParameters rpParameters)
        {
            string rptRef = "PEBankFile";
            string parameter1 = "EmployerKey";
            string parameter2 = "PayScheduleKey";
            string parameter3 = "PaymentDate";

            //Get the Bank File report
            XmlDocument xmlReport = RunReport(xdoc, rptRef, parameter1, rpParameters.ErRef,
                                              parameter2, rpParameters.PaySchedule,
                                              parameter3, rpParameters.PayRunDate.ToString("yyyy-MM-dd"),
                                              null, null, null, null, null, null);


            return xmlReport;
        }
        public string GetCsvPensionsReport(XDocument xdoc, RPParameters rpParameters, RPPensionFileScheme rpPensionFileScheme, bool isJoiner)
        {
            string rptRef = "PAPDIS";
            string parameter1 = "EmployerKey";
            string parameter2 = "PayScheduleKey";
            string parameter3 = "TaxYear";
            string parameter4 = "PaymentDate";
            string parameter5 = "PensionKey";
            string parameter6 = "TransformDefinitionKey";

            string transformKey;
            switch(rpPensionFileScheme.ProviderName)
            {
                case "ROYAL LONDON PENSION":
                    //Get the Royal London Pensions report
                    transformKey = "RL-PENSION-CSV";
                    break;
                case "FRIENDS LIFE PENSION":
                    //Get the Friends Life Pensions report
                    transformKey = "RL-PENSION-CSV";
                    break;
                case "STANDARD LIFE PENSION":
                    //Get the Standard Life Pensions report
                    transformKey = "RL-PENSION-CSV";
                    break;
                case "THE AMBER PENSION TRUST":
                    //Get The Amber Pension Trust Pensions report
                    transformKey = "RL-PENSION-CSV";
                    break;
                case "SCOTTISH WIDOWS PENSION":
                    //Get the Scottish Widows Pensions report
                    transformKey = "RL-PENSION-CSV";
                    break;
                case "NOW PENSION":
                    //Get the Now Pensions report
                    transformKey = "NOW-PENSION-CSV";
                    break;
                case "LEGAL & GENERAL":
                    //Get the Legal & General Pensions report
                    transformKey = "LG-CONTRIBS-CSV"; //TODO change this to LG-PENSION-CSV or whatever Tim calls it.
                    break;
                case "AEGON":
                    //Get the Aegon Pensions report
                    transformKey = "AEGON-CONTRIBS-CSV"; //TODO change this to AE-PENSION-CSV or whatever Tim calls it.
                    break;
                default:
                    //Get the Smart Pensions report
                    transformKey = "PE-PAPDIS1-CSV";
                    break;
            }
            if(isJoiner)
            {
                transformKey = transformKey.Replace("CONTRIBS-CSV", "JOINERS-CSV");
            }
            string csvReport = RunTransformReport(xdoc, rptRef,
                                parameter1, rpParameters.ErRef,
                                parameter2, rpParameters.PaySchedule,
                                parameter3, rpParameters.TaxYear.ToString(),
                                parameter4, rpParameters.PayRunDate.ToString("yyyy-MM-dd"),
                                parameter5, rpPensionFileScheme.Key.ToString(),
                                parameter6, transformKey);
            

            return csvReport;
        }
        public XmlDocument GetCombinedPayrollRunReport(XDocument xdoc, RPParameters rpParameters)
        {
            string rptRef = "CombinedPayrollRun";
            string parameter1 = "EmployerKey";
            string parameter2 = "PayScheduleKey";
            string parameter3 = "StartDate";
            string parameter4 = "EndDate";

            //Get the Combined Payroll Run report
            XmlDocument xmlReport = RunReport(xdoc, rptRef, parameter1, rpParameters.ErRef, parameter2, rpParameters.PaySchedule,
                                              parameter3, "2020/04/06", parameter4, "2021/04/05", null, null, null, null);

            
            return xmlReport;
        }
        public string[] GetAListOfDirectories(XDocument xdoc, string source)
        {
            string path = xdoc.Root.Element("DataHomeFolder").Value + source;
            string[] directories = Directory.GetDirectories(path);

            return directories;
        }
        public RPParameters GetRPParameters(XmlDocument xmlReport)
        {
            //Now extract the necessary data and produce the required reports.

            RPParameters rpParameters = new RPParameters();

            foreach (XmlElement parameter in xmlReport.GetElementsByTagName("Parameters"))
            {
                rpParameters.ErRef = GetElementByTagFromXml(parameter, "EmployerCode");
                rpParameters.TaxYear = GetIntElementByTagFromXml(parameter, "TaxYear");
                rpParameters.AccYearStart = Convert.ToDateTime(GetDateElementByTagFromXml(parameter, "AccountingYearStartDate"));
                rpParameters.AccYearEnd = Convert.ToDateTime(GetDateElementByTagFromXml(parameter, "AccountingYearEndDate"));
                rpParameters.TaxPeriod = GetIntElementByTagFromXml(parameter, "TaxPeriod");
                rpParameters.PeriodNo = GetIntElementByTagFromXml(parameter, "PeriodNumber");
                rpParameters.PaySchedule = GetElementByTagFromXml(parameter, "PaySchedule");
                rpParameters.PayRunDate = Convert.ToDateTime(GetDateElementByTagFromXml(parameter, "PaymentDate"));
            }
            return rpParameters;
        }
        public RPEmployer GetRPEmployer(XDocument xdoc, XmlDocument xmlReport, RPParameters rpParameters)
        {
            string objectType = "Employer";
            XmlDocument xmlEmployer = GetPayRunIOObject(xdoc, objectType, rpParameters.ErRef);
            RPEmployer rpEmployer = new RPEmployer();
            string dataSource = xdoc?.Root?.Element("DataSource").Value;            //"APPSERVER1\\MSSQL";  //"13.69.154.210\\MSSQL";  
            string dataBase = xdoc?.Root?.Element("Database").Value;
            string userID = xdoc?.Root?.Element("Username").Value;
            string password = xdoc?.Root?.Element("Password").Value;
            string sqlConnectionString = "Server=" + dataSource + ";Database=" + dataBase + ";User ID=" + userID + ";Password=" + password + ";";
            
            foreach (XmlElement employer in xmlReport.GetElementsByTagName("Employer"))
            {
                rpEmployer.Name = GetElementByTagFromXml(employer, "Name");
                rpEmployer.PayeRef = GetElementByTagFromXml(employer, "EmployerPayeRef");
                rpEmployer.P32Required = GetBooleanElementByTagFromXml(employer, "P32Required");
                rpEmployer.CalculateApprenticeshipLevy = GetBooleanElementByTagFromXml(employer, "CalculateApprenticeshipLevy");
            }
            foreach (XmlElement address in xmlEmployer.GetElementsByTagName("Address"))
            {
                rpEmployer.Address1 = GetElementByTagFromXml(address, "Address1");
                rpEmployer.Address2 = GetElementByTagFromXml(address, "Address2");
                rpEmployer.Address3 = GetElementByTagFromXml(address, "Address3");
                rpEmployer.Address4 = GetElementByTagFromXml(address, "Address4");
                rpEmployer.Postcode = GetElementByTagFromXml(address, "Postcode");
                rpEmployer.Country = GetElementByTagFromXml(address, "Country");
            }

            rpEmployer.BankFileCode = "000";
            rpEmployer.PensionReportFileType = "Unknown";
            rpEmployer.PensionReportAEWorkersGroup = "A";
            rpEmployer.NESTPensionText = "My source";
            rpEmployer.HREscapeCompanyNo = null;
            rpEmployer.ReportPassword = null;
            rpEmployer.ZipReports = true;
            rpEmployer.ReportsInExcelFormat = true;
            rpEmployer.PayRunDetailsYTDRequired = false;
            rpEmployer.PayrollTotalsSummaryRequired = false;
            rpEmployer.NoteAndCoinRequired = false;
            rpEmployer.HoldPayHistory = false;

            if (xdoc != null && xdoc.Root != null)
            {
                //Get the bank file code for a table on the database for now. It should be supplied by WebGlobe and then PR eventually.
                try
                {
                    DataRow drCompanyReportCodes = GetCompanyReportCodes(xdoc, sqlConnectionString, rpParameters);
                    if(drCompanyReportCodes.ItemArray[0] != System.DBNull.Value)
                    {
                        rpEmployer.BankFileCode = drCompanyReportCodes.ItemArray[0].ToString();
                    }
                    if (drCompanyReportCodes.ItemArray[1] != System.DBNull.Value)
                    {
                        rpEmployer.PensionReportFileType = drCompanyReportCodes.ItemArray[1].ToString();
                    }
                    if (drCompanyReportCodes.ItemArray[2] != System.DBNull.Value)
                    {
                        rpEmployer.PensionReportAEWorkersGroup = drCompanyReportCodes.ItemArray[2].ToString();
                    }
                    if (drCompanyReportCodes.ItemArray[3] != System.DBNull.Value)
                    {
                        rpEmployer.NESTPensionText = drCompanyReportCodes.ItemArray[3].ToString();
                    }
                    if (drCompanyReportCodes.ItemArray[4] != System.DBNull.Value)
                    {
                        rpEmployer.HREscapeCompanyNo = Convert.ToInt32(drCompanyReportCodes.ItemArray[4]);
                    }
                    if (drCompanyReportCodes.ItemArray[5] != System.DBNull.Value)
                    {
                        rpEmployer.ReportPassword = drCompanyReportCodes.ItemArray[5].ToString();
                    }
                    if (drCompanyReportCodes.ItemArray[6] != System.DBNull.Value)
                    {
                        rpEmployer.ZipReports = Convert.ToBoolean(drCompanyReportCodes.ItemArray[6]);
                    }
                    if (drCompanyReportCodes.ItemArray[7] != System.DBNull.Value)
                    {
                        rpEmployer.ReportsInExcelFormat = Convert.ToBoolean(drCompanyReportCodes.ItemArray[7]);
                    }
                    if (drCompanyReportCodes.ItemArray[8] != System.DBNull.Value)
                    {
                        rpEmployer.PayRunDetailsYTDRequired = Convert.ToBoolean(drCompanyReportCodes.ItemArray[8]);
                    }
                    if (drCompanyReportCodes.ItemArray[9] != System.DBNull.Value)
                    {
                        rpEmployer.PayrollTotalsSummaryRequired = Convert.ToBoolean(drCompanyReportCodes.ItemArray[9]);
                    }
                    if (drCompanyReportCodes.ItemArray[10] != System.DBNull.Value)
                    {
                        rpEmployer.NoteAndCoinRequired = Convert.ToBoolean(drCompanyReportCodes.ItemArray[10]);
                    }
                    if (drCompanyReportCodes.ItemArray[11] != System.DBNull.Value)
                    {
                        rpEmployer.HoldPayHistory = Convert.ToBoolean(drCompanyReportCodes.ItemArray[11]);
                    }
                }
                catch
                {
                   
                }
            }
            return rpEmployer;
        }
        public void ArchiveRTIOutputs(string directory, FileInfo file)
        {
            //Move RTI file to PE-ArchivedRTI from Outputs
            string archiveDirName = directory.Replace("Outputs", "PE-ArchivedRTI");
            Directory.CreateDirectory(directory.Replace(directory, archiveDirName));
            string destinationFilename = file.FullName.Replace("Outputs", "PE-ArchivedRTI");
            File.Move(file.FullName, destinationFilename);
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
            if(element == "Y" || element == "Yes")
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
        public static void EmptyDirectory(DirectoryInfo directory)
        {
            foreach (System.IO.FileInfo file in directory.GetFiles()) file.Delete();
            foreach (System.IO.DirectoryInfo subDirectory in directory.GetDirectories()) subDirectory.Delete(true);
        }
        public void DeleteFilesThenFolder(XDocument xdoc, string sourceFolder)
        {
            string configDirName = xdoc.Root.Element("SoftwareHomeFolder").Value;
            string textLine;

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
                Update_Progress(textLine, configDirName);
            }
        }
        
        public SMTPEmailSettings GetEmailSettings(XDocument xdoc, string sqlConnectionString)
        {
            int x = sqlConnectionString.LastIndexOf(";Password=") + 10;
            int y = sqlConnectionString.LastIndexOf(";");
            string logConnectionString = sqlConnectionString.Substring(0, x + 2) + "*********" + sqlConnectionString.Substring(y - 2);


            string configDirName = xdoc.Root.Element("SoftwareHomeFolder").Value;
            string textLine;
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
                Update_Progress(textLine, configDirName);
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

                
                textLine = string.Format("Getting SMTP email settings with connection string : {0}.", logConnectionString);
                Update_Progress(textLine, configDirName);

                textLine = string.Format("Got SMTP email settings, host is : {0}.", smtpEmailSettings.SMTPHost);
                Update_Progress(textLine, configDirName);

            }

            return smtpEmailSettings;

        }
        public List<ContactInfo> GetListOfContactInfo(XDocument xdoc, string sqlConnectionString, RPParameters rpParameters)
        {
            string configDirName = xdoc.Root.Element("SoftwareHomeFolder").Value;
            string textLine;

            int x = sqlConnectionString.LastIndexOf(";Password=") + 10;
            int y = sqlConnectionString.LastIndexOf(";");
            string logConnectionString = sqlConnectionString.Substring(0, x + 2) + "*********" + sqlConnectionString.Substring(y - 2);

            textLine = string.Format("Start getting a list of email addresses with connection string : {0}.", logConnectionString);
            Update_Progress(textLine, configDirName);

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
                Update_Progress(textLine, configDirName);
            }
            foreach (DataRow drContactInfo in dtContactInfo.Rows)
            {
                ContactInfo contactInfo = new ContactInfo()
                {
                    EmailAddress = drContactInfo.ItemArray[0].ToString(),
                    FirstName = drContactInfo.ItemArray[1].ToString()
                };
                contactInfoList.Add(contactInfo);
            }

            textLine = string.Format("Finished getting a list of email addresses with connection string : {0}.", logConnectionString);
            Update_Progress(textLine, configDirName);

            return contactInfoList;
        }
        private DataRow GetCompanyReportCodes(XDocument xdoc, string sqlConnectionString, RPParameters rpParameters)
        {
            string configDirName = xdoc.Root.Element("SoftwareHomeFolder").Value;
            string textLine;

            int x = sqlConnectionString.LastIndexOf(";Password=") + 10;
            int y = sqlConnectionString.LastIndexOf(";");
            string logConnectionString = sqlConnectionString.Substring(0, x + 2) + "*********" + sqlConnectionString.Substring(y - 2);

            textLine = string.Format("Start getting the company report codes with connection string : {0}.", logConnectionString);
            Update_Progress(textLine, configDirName);

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
                Update_Progress(textLine, configDirName);
            }
            
            DataRow drCompanyReportCodes = dtCompanyReportCodes.Rows[0];

            textLine = string.Format("Finished getting company report codes with connection string : {0}.", logConnectionString);
            Update_Progress(textLine, configDirName);

            return drCompanyReportCodes;
        }
        public bool GetIsUnity(XDocument xdoc, string sqlConnectionString, int companyNo)
        {
            bool isUnity;
            string configDirName = xdoc.Root.Element("SoftwareHomeFolder").Value;
            string textLine;

            int x = sqlConnectionString.LastIndexOf(";Password=") + 10;
            int y = sqlConnectionString.LastIndexOf(";");
            string logConnectionString = sqlConnectionString.Substring(0, x + 2) + "*********" + sqlConnectionString.Substring(y - 2);

            textLine = string.Format("Start getting IsUnity with connection string : {0}.", logConnectionString);
            Update_Progress(textLine, configDirName);

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
                Update_Progress(textLine, configDirName);
            }

            isUnity = Convert.ToBoolean(dtCompanyReportCodes.Rows[0].ItemArray[0]);
            
            textLine = string.Format("Finished getting IsUnity with connection string : {0}.", logConnectionString);
            Update_Progress(textLine, configDirName);

            return isUnity;
        }
        public XtraReport CreatePDFReport(XmlDocument xmlReport, string reportName, string assemblyName)
        {
            //Load report
            reportName += ".repx";
            var reportLayout = ResourceHelper.ReadResourceFileToStream(
                assemblyName, reportName);
            XtraReport xtraReport = XtraReport.FromStream(reportLayout);
            XmlReader xmlReader = new XmlNodeReader(xmlReport);
            DataSet set = new DataSet();
            set.ReadXml(xmlReader);

            xtraReport.DataSource = set;

            return xtraReport;
        }
        public XtraReport CreatePDFReport(List<RPEmployeePeriod> rpEmployeePeriodList, RPEmployer rpEmployer, RPParameters rpParameters, string reportName, string assemblyName)
        {
            var payeMonth = rpParameters.PayRunDate.Day < 6 ? rpParameters.PayRunDate.Month - 4 : rpParameters.PayRunDate.Month - 3;
            if (payeMonth <= 0)
            {
                payeMonth += 12;
            }
            //Load report
            reportName += ".repx";
            var reportLayout = ResourceHelper.ReadResourceFileToStream(
                assemblyName, reportName);
            XtraReport xtraReport = XtraReport.FromStream(reportLayout);
            xtraReport.Parameters["CmpName"].Value = rpEmployer.Name;
            xtraReport.Parameters["PayeRef"].Value = rpEmployer.PayeRef;
            xtraReport.Parameters["Date"].Value = rpParameters.PayRunDate;
            xtraReport.Parameters["Period"].Value = rpParameters.PeriodNo;
            xtraReport.Parameters["Freq"].Value = rpParameters.PaySchedule;
            xtraReport.Parameters["PAYEMonth"].Value = payeMonth;
            xtraReport.DataSource = rpEmployeePeriodList;

            return xtraReport;
        }
        public XtraReport CreatePDFReport(List<RPPensionContribution> rpPensionContributions, RPEmployer rpEmployer, RPParameters rpParameters, string reportName, string assemblyName)
        {
            var payeMonth = rpParameters.PayRunDate.Day < 6 ? rpParameters.PayRunDate.Month - 4 : rpParameters.PayRunDate.Month - 3;
            if (payeMonth <= 0)
            {
                payeMonth += 12;
            }
            //Load report
            reportName += ".repx";
            var reportLayout = ResourceHelper.ReadResourceFileToStream(
                assemblyName, reportName);
            XtraReport xtraReport = XtraReport.FromStream(reportLayout);
            xtraReport.Parameters["CmpName"].Value = rpEmployer.Name;
            xtraReport.Parameters["PayeRef"].Value = rpEmployer.PayeRef;
            xtraReport.Parameters["Date"].Value = rpParameters.PayRunDate;
            xtraReport.Parameters["Period"].Value = rpParameters.PeriodNo;
            xtraReport.Parameters["Freq"].Value = rpParameters.PaySchedule;
            xtraReport.Parameters["PAYEMonth"].Value = payeMonth;
            xtraReport.DataSource = rpPensionContributions;

            return xtraReport;
        }
        public XtraReport CreatePDFReport(List<RPPayComponent> rpPayComponents, RPEmployer rpEmployer, RPParameters rpParameters, bool showDetail, string reportName, string assemblyName)
        {
            var payeMonth = rpParameters.PayRunDate.Day < 6 ? rpParameters.PayRunDate.Month - 4 : rpParameters.PayRunDate.Month - 3;
            if (payeMonth <= 0)
            {
                payeMonth += 12;
            }
            //Load report
            reportName += ".repx";
            var reportLayout = ResourceHelper.ReadResourceFileToStream(
                assemblyName, reportName);
            XtraReport xtraReport = XtraReport.FromStream(reportLayout);
            xtraReport.Parameters["CmpName"].Value = rpEmployer.Name;
            xtraReport.Parameters["PayeRef"].Value = rpEmployer.PayeRef;
            xtraReport.Parameters["Date"].Value = rpParameters.PayRunDate;
            xtraReport.Parameters["Period"].Value = rpParameters.PeriodNo;
            xtraReport.Parameters["Freq"].Value = rpParameters.PaySchedule;
            xtraReport.Parameters["PAYEMonth"].Value = payeMonth;
            xtraReport.Parameters["ShowDetailBand"].Value = showDetail;
            xtraReport.DataSource = rpPayComponents;

            return xtraReport;
        }
        public XtraReport CreatePDFReport(List<P45> p45s, string reportName, string assemblyName)
        {
            //Load report
            reportName += ".repx";
            var reportLayout = ResourceHelper.ReadResourceFileToStream(
                assemblyName, reportName);
            XtraReport xtraReport = XtraReport.FromStream(reportLayout);
            xtraReport.DataSource = p45s;

            return xtraReport;
        }
        public XtraReport CreatePDFReport(RPP32Report rpP32Report, RPEmployer rpEmployer, RPParameters rpParameters, string reportName, string assemblyName)
        {
            var payeMonth = rpParameters.PayRunDate.Day < 6 ? rpParameters.PayRunDate.Month - 4 : rpParameters.PayRunDate.Month - 3;
            if (payeMonth <= 0)
            {
                payeMonth += 12;
            }
            //Load report
            reportName += ".repx";
            var reportLayout = ResourceHelper.ReadResourceFileToStream(
                assemblyName, reportName);
            XtraReport xtraReport = XtraReport.FromStream(reportLayout);
            xtraReport.Parameters["CmpName"].Value = rpEmployer.Name;
            xtraReport.Parameters["PayeRef"].Value = rpEmployer.PayeRef;
            xtraReport.Parameters["Date"].Value = rpParameters.PayRunDate;
            xtraReport.Parameters["Period"].Value = rpParameters.PeriodNo;
            xtraReport.Parameters["Freq"].Value = rpParameters.PaySchedule;
            xtraReport.Parameters["PAYEMonth"].Value = payeMonth;
            xtraReport.Parameters["AnnualEmploymentAllowance"].Value = rpP32Report.AnnualEmploymentAllowance;
            xtraReport.Parameters["PaymentRef"].Value = rpP32Report.PaymentRef;
            xtraReport.Parameters["TaxYearStartDate"].Value = rpP32Report.TaxYearStartDate;
            xtraReport.Parameters["TaxYearEndDate"].Value = rpP32Report.TaxYearEndDate;
            xtraReport.Parameters["TaxYear"].Value = rpP32Report.TaxYear;
            xtraReport.DataSource = rpP32Report.RPP32ReportMonths;

            return xtraReport;
        }
        private List<string> CreateListOfFixedColumns()
        {
            //Create a list of the required fixed columns.
            List<string> fixCol = new List<string>()
            {
                "PayRunDate",
                "EeRef",
                "Name",
                "Dept",
                "CostCentre",
                "Branch",
                "Status",
                "TaxCode",
                "NILetter",
                "PreTaxAddDed",
                "GrossedUpTaxThisRun",
                "EeNIPdByEr",
                "GUStudentLoan",
                "GUNIReduction",
                "PenPreTaxEeGU",
                "TotalAbsencePay",
                "HolidayPay",
                "PenPreTaxEe",
                "TaxablePay",
                "Tax",
                "NI",
                "PostTaxAddDed",
                "PostTaxPension",
                "AEO",
                "StudentLoan",
                "NetPay",
                "ErNI",
                "PenEr",
                "TotalGrossUp",
                "TotalNICs",
                "TotalPens"
            };
            
            return fixCol;
        }
        private List<string> CreateListOfVariableColumns(List<RPPreSamplePayCode> rpPreSamplePayCodes)
        {
            //Create a list of the required variable columns.
            List<string> varCol = new List<string>();

            foreach (RPPreSamplePayCode rpPreSamplePayCode in rpPreSamplePayCodes)
            {
                if (rpPreSamplePayCode.Code != "TAX" && rpPreSamplePayCode.Code != "NI")
                {
                    if (rpPreSamplePayCode.InUse)
                    {
                        varCol.Add(rpPreSamplePayCode.Description);
                    }
                }
            }

            return varCol;
        }
        private Workbook CreateFixedWorkbookColumns(Workbook workbook, RPEmployeePeriod rpEmployeePeriod)
        {
            workbook.CurrentWorksheet.AddNextCell(rpEmployeePeriod.PayRunDate);
            workbook.CurrentWorksheet.AddNextCell(rpEmployeePeriod.Reference);
            workbook.CurrentWorksheet.AddNextCell(rpEmployeePeriod.Fullname);
            workbook.CurrentWorksheet.AddNextCell(rpEmployeePeriod.Department);
            workbook.CurrentWorksheet.AddNextCell(rpEmployeePeriod.CostCentre);
            workbook.CurrentWorksheet.AddNextCell(rpEmployeePeriod.Branch);
            workbook.CurrentWorksheet.AddNextCell("Calc");
            workbook.CurrentWorksheet.AddNextCell(rpEmployeePeriod.TaxCode);
            workbook.CurrentWorksheet.AddNextCell(rpEmployeePeriod.NILetter);
            workbook.CurrentWorksheet.AddNextCell(rpEmployeePeriod.PreTaxAddDed);
            workbook.CurrentWorksheet.AddNextCell(0.00);//GrossedUpTaxThisRun
            workbook.CurrentWorksheet.AddNextCell(0.00);//EeNIPdByEr
            workbook.CurrentWorksheet.AddNextCell(0.00);//GUStudentLoan
            workbook.CurrentWorksheet.AddNextCell(0.00);//GUNIReduction
            workbook.CurrentWorksheet.AddNextCell(0.00);//PenPreTaxEeGU
            workbook.CurrentWorksheet.AddNextCell(rpEmployeePeriod.AbsencePay);
            workbook.CurrentWorksheet.AddNextCell(rpEmployeePeriod.HolidayPay);
            workbook.CurrentWorksheet.AddNextCell(rpEmployeePeriod.PreTaxPension);
            workbook.CurrentWorksheet.AddNextCell(rpEmployeePeriod.TaxablePayTP);
            workbook.CurrentWorksheet.AddNextCell(rpEmployeePeriod.Tax);
            workbook.CurrentWorksheet.AddNextCell(rpEmployeePeriod.NetNI);
            workbook.CurrentWorksheet.AddNextCell(rpEmployeePeriod.PostTaxAddDed);
            workbook.CurrentWorksheet.AddNextCell(rpEmployeePeriod.PostTaxPension);
            workbook.CurrentWorksheet.AddNextCell(rpEmployeePeriod.AEO);
            workbook.CurrentWorksheet.AddNextCell(rpEmployeePeriod.StudentLoan);
            workbook.CurrentWorksheet.AddNextCell(rpEmployeePeriod.NetPayTP);
            workbook.CurrentWorksheet.AddNextCell(rpEmployeePeriod.ErNICTP);

            decimal erPensionTP = 0;
            decimal eePensionTP = 0;
            foreach (RPPensionPeriod pensionPeriod in rpEmployeePeriod.Pensions)
            {
                erPensionTP += pensionPeriod.ErPensionTaxPeriod;
                eePensionTP += pensionPeriod.EePensionTaxPeriod;
            }
            workbook.CurrentWorksheet.AddNextCell(erPensionTP);
            workbook.CurrentWorksheet.AddNextCell(0.00);//TotalGrossUP
            workbook.CurrentWorksheet.AddNextCell(rpEmployeePeriod.ErNICTP + rpEmployeePeriod.NetNI);
            workbook.CurrentWorksheet.AddNextCell(erPensionTP + eePensionTP);

            return workbook;
        }
        private Workbook CreateVariableWorkbookColumns(Workbook workbook, RPEmployeePeriod rpEmployeePeriod, List<string> varCol)
        {
            foreach (string col in varCol)
            {
                //Add in the variable additions.
                bool colFound = false;
                foreach (RPAddition rpAddition in rpEmployeePeriod.Additions)
                {
                    if (col == rpAddition.Description)
                    {
                        workbook.CurrentWorksheet.AddNextCell(rpAddition.AmountTP);
                        colFound = true;
                        break;
                    }

                }
                //If the column has not been found in additions check the variable deductions.
                if (!colFound)
                {
                    foreach (RPDeduction rpDeduction in rpEmployeePeriod.Deductions)
                    {
                        if (col == rpDeduction.Description)
                        {
                            workbook.CurrentWorksheet.AddNextCell(rpDeduction.AmountTP);
                            colFound = true;
                            break;
                        }

                    }
                    //If the column hasn't been found in additions or deduction set it to zero.
                    if (!colFound)
                    {
                        workbook.CurrentWorksheet.AddNextCell(0.00m);
                    }
                }



            }

            return workbook;
        }
        private List<RPPreSamplePayCode> CreateListOfRequiredColumns(XmlDocument xmlReport)
        {
            RPPreSamplePayCode rpPreSamplePayCode;
            //Create a list of all possible pay codes. For this purpose pensions can be turned into pay codes.
            List<RPPreSamplePayCode> rpPreSamplePayCodes = new List<RPPreSamplePayCode>();

            //There could be multiple payruns in this xml file.
            foreach (XmlElement payRun in xmlReport.GetElementsByTagName("PayRun"))
            {
                DateTime payRunDate = Convert.ToDateTime(GetDateElementByTagFromXml(payRun, "PayRunDate"));
                //There could be multiple employees in each pay run.
                foreach (XmlElement employee in payRun.GetElementsByTagName("Employee"))
                {
                    //There could be multiple pensions in each employee.
                    foreach (XmlElement pension in employee.GetElementsByTagName("Pension"))
                    {
                        string eeCode = GetElementByTagFromXml(pension, "Code") +
                                   GetElementByTagFromXml(pension, "ProviderName");
                        string eeDesc = GetElementByTagFromXml(pension, "SchemeName");
                        string erCode = eeCode + "(Er)";
                        string erDesc = eeDesc + "(Er)";
                        //Add the Er pension
                        rpPreSamplePayCode = new RPPreSamplePayCode()
                        {
                            Code = erCode,
                            Description = erDesc,
                            InUse = true
                        };
                        rpPreSamplePayCodes = CheckAddToList(rpPreSamplePayCodes, rpPreSamplePayCode);
                        //Add the Ee Pension
                        rpPreSamplePayCode = new RPPreSamplePayCode()
                        {
                            Code = eeCode,
                            Description = eeDesc,
                            InUse = true
                        };
                        rpPreSamplePayCodes = CheckAddToList(rpPreSamplePayCodes, rpPreSamplePayCode);
                    }
                    //There could be multiple pay codes in each employee.
                    foreach (XmlElement payCode in employee.GetElementsByTagName("PayCode"))
                    {
                        rpPreSamplePayCode = new RPPreSamplePayCode()
                        {
                            Code = GetElementByTagFromXml(payCode, "Code"),
                            Description = GetElementByTagFromXml(payCode, "Description"),
                            InUse = true
                        };
                        rpPreSamplePayCodes = CheckAddToList(rpPreSamplePayCodes, rpPreSamplePayCode);
                    }
                }
            }
            return rpPreSamplePayCodes;
        }
        private List<RPEmployeePeriod> CreateListOfEmployeePeriods(XmlDocument xmlReport)
        {
            //Create a list of all the employees within each pay run date.
            List<RPEmployeePeriod> rpEmployeePeriods = new List<RPEmployeePeriod>();
            //There could be multiple payruns in this xml file.
            foreach (XmlElement payRun in xmlReport.GetElementsByTagName("PayRun"))
            {
                DateTime payRunDate = Convert.ToDateTime(GetDateElementByTagFromXml(payRun, "PaymentDate"));
                //There could be multiple employees in each pay run.
                foreach (XmlElement employee in payRun.GetElementsByTagName("Employee"))
                {
                    RPEmployeePeriod rpEmployeePeriod = new RPEmployeePeriod();
                    rpEmployeePeriod.PayRunDate = payRunDate;
                    rpEmployeePeriod.Reference = GetElementByTagFromXml(employee, "Code");
                    rpEmployeePeriod.Fullname = GetElementByTagFromXml(employee, "LastName") +
                                                       " " +
                                                       GetElementByTagFromXml(employee, "FirstName");
                    rpEmployeePeriod.TaxCode = GetElementByTagFromXml(employee, "PayLineTaxCode");
                    rpEmployeePeriod.NILetter = GetElementByTagFromXml(employee, "NiLetter");
                    rpEmployeePeriod.PreTaxAddDed = GetDecimalElementByTagFromXml(employee, "PreTaxAddDed");
                    rpEmployeePeriod.AbsencePay = GetDecimalElementByTagFromXml(employee, "AbsencePay");
                    rpEmployeePeriod.HolidayPay = GetDecimalElementByTagFromXml(employee, "HolidayPay");
                    rpEmployeePeriod.PreTaxPension = GetDecimalElementByTagFromXml(employee, "PreTaxPension");
                    rpEmployeePeriod.TaxablePayTP = GetDecimalElementByTagFromXml(employee, "TaxablePay");
                    rpEmployeePeriod.Tax = GetDecimalElementByTagFromXml(employee, "Tax");
                    rpEmployeePeriod.NetNI = GetDecimalElementByTagFromXml(employee, "NetEeNi");
                    rpEmployeePeriod.PostTaxAddDed = GetDecimalElementByTagFromXml(employee, "PostTaxAddDed");
                    rpEmployeePeriod.PostTaxPension = GetDecimalElementByTagFromXml(employee, "PostTaxPension");
                    rpEmployeePeriod.AEO = GetDecimalElementByTagFromXml(employee, "AEO");
                    rpEmployeePeriod.StudentLoan = GetDecimalElementByTagFromXml(employee, "StudentLoan");
                    rpEmployeePeriod.NetPayTP = GetDecimalElementByTagFromXml(employee, "NetPay");
                    rpEmployeePeriod.ErNICTP = GetDecimalElementByTagFromXml(employee, "ErNi");
                    rpEmployeePeriod.ErPensionTotalTP = GetDecimalElementByTagFromXml(employee, "ErPension");
                    rpEmployeePeriod.Branch = GetElementByTagFromXml(employee, "Branch");
                    rpEmployeePeriod.Branch = GetElementByTagFromXml(employee, "Department");
                    rpEmployeePeriod.Branch = GetElementByTagFromXml(employee, "CostCentre");

                    List<RPAddition> rpAdditions = new List<RPAddition>();
                    List<RPDeduction> rpDeductions = new List<RPDeduction>();
                    List<RPPensionPeriod> rpPensionPeriods = new List<RPPensionPeriod>();
                    //There could be multiple pensions in each employee.
                    foreach (XmlElement pension in employee.GetElementsByTagName("Pension"))
                    {
                        RPPensionPeriod rpPensionPeriod = new RPPensionPeriod();
                        rpPensionPeriod.Key = Convert.ToInt32(pension.GetAttribute("Key"));
                        rpPensionPeriod.Code = GetElementByTagFromXml(pension, "Code");
                        rpPensionPeriod.ProviderName = GetElementByTagFromXml(pension, "ProviderName");
                        rpPensionPeriod.SchemeName = GetElementByTagFromXml(pension, "SchemeName");
                        rpPensionPeriod.StartJoinDate = GetDateElementByTagFromXml(pension, "StartJoinDate");
                        rpPensionPeriod.IsJoiner = GetBooleanElementByTagFromXml(pension, "IsJoiner");
                        rpPensionPeriod.ProviderEmployerReference = GetElementByTagFromXml(pension, "ProviderEmployerRef");
                        rpPensionPeriod.EePensionYtd = GetDecimalElementByTagFromXml(pension, "EePensionYtd");
                        rpPensionPeriod.ErPensionYtd = GetDecimalElementByTagFromXml(pension, "ErPensionYtd");
                        rpPensionPeriod.PensionablePayYtd = GetDecimalElementByTagFromXml(pension, "PensionablePayYtd");
                        rpPensionPeriod.EePensionTaxPeriod = GetDecimalElementByTagFromXml(pension, "EePensionTaxPeriod");
                        rpPensionPeriod.ErPensionTaxPeriod = GetDecimalElementByTagFromXml(pension, "ErPensionTaxPeriod");
                        rpPensionPeriod.PensionablePayTaxPeriod = GetDecimalElementByTagFromXml(pension, "PensionablePayTaxPeriod");
                        rpPensionPeriod.EePensionPayRunDate = GetDecimalElementByTagFromXml(pension, "EePensionPayRunDate");
                        rpPensionPeriod.ErPensionPayRunDate = GetDecimalElementByTagFromXml(pension, "ErPensionPayRunDate");
                        rpPensionPeriod.PensionablePayPayRunDate = GetDecimalElementByTagFromXml(pension, "PensionablePayDate");
                        rpPensionPeriod.EeContibutionPercent = GetDecimalElementByTagFromXml(pension, "EeContributionPercent") * 100;
                        rpPensionPeriod.ErContributionPercent = GetDecimalElementByTagFromXml(pension, "ErContributionPercent") * 100;
                        rpPensionPeriod.TotalPayTaxPeriod = rpEmployeePeriod.Gross;

                        rpPensionPeriods.Add(rpPensionPeriod);

                        string eeCode = GetElementByTagFromXml(pension, "Code") +
                                   GetElementByTagFromXml(pension, "ProviderName");
                        string eeDesc = GetElementByTagFromXml(pension, "SchemeName");
                        string erCode = eeCode + "(Er)";
                        string erDesc = eeDesc + "(Er)";
                        //Add as an addition to the employee period object for Ee pension
                        RPAddition rpAddition = new RPAddition()
                        {
                            Code = erCode,
                            Description = erDesc,
                            AmountTP = GetDecimalElementByTagFromXml(pension, "ErPensionTaxPeriod")
                        };
                        rpAdditions.Add(rpAddition);
                        //Add as an addition to the employee period object for Ee pension
                        rpAddition = new RPAddition()
                        {
                            Code = eeCode,
                            Description = eeDesc,
                            AmountTP = GetDecimalElementByTagFromXml(pension, "EePensionTaxPeriod")
                        };
                        rpAdditions.Add(rpAddition);
                    }
                    rpEmployeePeriod.Pensions = rpPensionPeriods;
                    //There could be multiple pay codes in each employee.
                    foreach (XmlElement payCode in employee.GetElementsByTagName("PayCode"))
                    {
                        //Add them all as additions for the purposes of this report
                        RPAddition rpAddition = new RPAddition()
                        {
                            Code = GetElementByTagFromXml(payCode, "Code"),
                            Description = GetElementByTagFromXml(payCode, "Description"),
                            AmountTP = GetDecimalElementByTagFromXml(payCode, "Amount")
                        };
                        //If it's a deduction multiply by -1
                        if (GetElementByTagFromXml(payCode, "EarningOrDeduction") == "D")
                        {
                            rpAddition.AmountTP *= -1;
                        }

                        rpAdditions.Add(rpAddition);

                    }
                    rpEmployeePeriod.Additions = rpAdditions;
                    rpEmployeePeriod.Deductions = rpDeductions;
                    rpEmployeePeriods.Add(rpEmployeePeriod);
                }

            }
            return rpEmployeePeriods;
        }
        private List<RPPreSamplePayCode> CheckAddToList(List<RPPreSamplePayCode> rpPreSamplePayCodes, RPPreSamplePayCode rpPreSampleNewPayCode)
        {
            bool inList = false;
            foreach (RPPreSamplePayCode rpPreSamplePayCode in rpPreSamplePayCodes)
            {
                if (rpPreSampleNewPayCode.Code == rpPreSamplePayCode.Code)
                {
                    inList = true;
                    break;
                }
            }
            if (!inList)
            {
                rpPreSamplePayCodes.Add(rpPreSampleNewPayCode);
            }
            return rpPreSamplePayCodes;
        }
        public PicoXLSX.Workbook PrepareBottomlineReport(XmlDocument xmlReport, PicoXLSX.Workbook workbook)
        {
            DateTime creationDate = DateTime.Now;
            DateTime processDate = DateTime.Now;
            string bacsServiceUserNumber = null;
            string erName;
            string erBankAccountName = null;
            string erBankSortCode = null;
            string erBankAccountNumber = null;
            decimal totalAmount = 0;
            foreach(XmlElement employer in xmlReport.GetElementsByTagName("Employer"))
            {
                processDate=Convert.ToDateTime(GetElementByTagFromXml(employer,"PaymentDate"));
                erName = GetElementByTagFromXml(employer, "Name");
                erBankAccountName = GetElementByTagFromXml(employer, "BankAccountName");
                erBankAccountNumber = GetElementByTagFromXml(employer, "BankAccountNumber");
                erBankSortCode = GetElementByTagFromXml(employer, "BankAccountSortCode");
                bacsServiceUserNumber = GetElementByTagFromXml(employer, "BacsServiceUserNumber");
                
            }
            processDate = GetPreviousWorkingDay(processDate);
            //First row
            workbook = CreateBottomlineRow(workbook, "BACS File Submission", "", "", "Creation date:", creationDate.ToString("dd/MM/yyyy"));
            //Next row
            workbook = CreateBottomlineRow(workbook, "", "", "", "Process date:", processDate.ToString("dd/MM/yyyy"));
            //Next row
            workbook = CreateBottomlineRow(workbook, "", "", "", "Value date:", creationDate.ToString("dd/MM/yyyy"));
            //Next row
            workbook = CreateBottomlineRow(workbook, "Service User Number", bacsServiceUserNumber, "", "", "");
            //Next row
            workbook.CurrentWorksheet.GoToNextRow();
            //Next row
            workbook = CreateBottomlineRow(workbook, "AccName", "SortCode", "AccNumber", "Amount", "Ref");
            //Loop through each employee
            string eeFullName;
            string eeBankAccountName;
            string eeBankAccountNumber;
            string eeBankSortCode;
            string eeBankAccountReference;
            decimal eeNetPay;
            foreach(XmlElement payRun in xmlReport.GetElementsByTagName("PayRuns"))
            {
                foreach(XmlElement employee in payRun.GetElementsByTagName("Employee"))
                {
                    eeFullName = GetElementByTagFromXml(employee, "FullName");
                    eeBankAccountName = GetElementByTagFromXml(employee, "BankAccountName");
                    eeBankAccountNumber = GetElementByTagFromXml(employee, "BankAccountNumber");
                    eeBankSortCode = GetElementByTagFromXml(employee, "BankAccountSortCode");
                    eeBankAccountReference = GetElementByTagFromXml(employee, "BankAccountReference");
                    eeNetPay = GetDecimalElementByTagFromXml(employee, "NetPay");
                    //Next row
                    workbook = CreateBottomlineRow(workbook, eeBankAccountName, eeBankSortCode, eeBankAccountNumber, eeNetPay.ToString(), eeBankAccountReference);
                    totalAmount += eeNetPay;
                }
            }
            workbook = CreateBottomlineRow(workbook, erBankAccountName, erBankSortCode, erBankAccountNumber, (totalAmount * -1).ToString(), "CONTRA");
            
            return workbook;
        }
        private PicoXLSX.Workbook CreateBottomlineRow(PicoXLSX.Workbook workbook, string col1, string col2, string col3, string col4, string col5)
        {
            workbook.CurrentWorksheet.AddNextCell(col1);
            workbook.CurrentWorksheet.AddNextCell(col2);
            workbook.CurrentWorksheet.AddNextCell(col3);
            workbook.CurrentWorksheet.AddNextCell(col4);
            workbook.CurrentWorksheet.AddNextCell(col5);
            workbook.CurrentWorksheet.GoToNextRow();
            return workbook;
        }
        public DateTime GetPreviousWorkingDay(DateTime date)
        {
            DateTime prevDay = date.AddDays(-1);
            
            while(prevDay.DayOfWeek == DayOfWeek.Saturday || prevDay.DayOfWeek == DayOfWeek.Sunday)
            {
                prevDay = prevDay.AddDays(-1);
            }
            return prevDay;
        }
        public PicoXLSX.Workbook PreparePreReport(XmlDocument xmlReport, PicoXLSX.Workbook workbook)
        {
            //Create a list of pay codes that are in use.
            List<RPPreSamplePayCode> rpPreSamplePayCodes = CreateListOfRequiredColumns(xmlReport);
            //Create a list of the fixed columns required.
            List<string> fixCol = CreateListOfFixedColumns();
            //Create a list of the variable columns required.
            List<string> varCol = CreateListOfVariableColumns(rpPreSamplePayCodes);
            //Create a list of employee period object within each pay run.
            List<RPEmployeePeriod> rpEmployeePeriods = CreateListOfEmployeePeriods(xmlReport);

            //Create a workbook.
            workbook = CreatePreXLSX(rpEmployeePeriods, fixCol, varCol, workbook);

            return workbook;
        }
        public PicoXLSX.Workbook CreatePreXLSX(List<RPEmployeePeriod> rpEmployeePeriodList,
                                       List<string> fixCol, List<string> varCol,
                                       PicoXLSX.Workbook workbook)
        {
            //Add the fixed headings
            foreach (string col in fixCol)
            {
                workbook.CurrentWorksheet.AddNextCell(col, PicoXLSX.Style.BasicStyles.Bold);

            }
            //Add the variable headings
            foreach (string col in varCol)
            {
                workbook.CurrentWorksheet.AddNextCell(col, PicoXLSX.Style.BasicStyles.Bold);
            }

            //Now for each employee create a row and add in the values for each column
            foreach (RPEmployeePeriod rpEmployeePeriod in rpEmployeePeriodList)
            {
                workbook.CurrentWorksheet.GoToNextRow();

                workbook = CreateFixedWorkbookColumns(workbook, rpEmployeePeriod);
                workbook = CreateVariableWorkbookColumns(workbook, rpEmployeePeriod, varCol);

            }
            //Try adding a formula
            workbook.CurrentWorksheet.GoToNextRow();
            workbook.CurrentWorksheet.GoToNextRow();

            workbook.CurrentWorksheet.AddNextCell("Totals", PicoXLSX.Style.BasicStyles.ColorizedText("990000"));

            //From Reference column to NILetter column
            for (int i = 0; i < 8; i++)
            {
                workbook.CurrentWorksheet.AddNextCell("");
            }
            //The first 9 columns are text and cannot be summed. 
            //The can be summed using a formula in the form =SUM(J2:J61). Column J is column 10 and is the first summable column.
            int rows = rpEmployeePeriodList.Count + 1;
            int cols = fixCol.Count + varCol.Count - 9;
            for (int i = 10; i < cols + 10; i++)
            {
                string colName = GetExcelColumnName(i);
                string formula = "=SUM(" + colName + "2:" + colName + rows + ")";
                workbook.WS.Formula(formula, PicoXLSX.Style.BasicStyles.ColorizedText("990000"));
            }

            return workbook;

        }
        public string GetEmployerNumber(XmlDocument xmlReport)
        {
            string coNo = null;
            foreach (XmlElement parameters in xmlReport.GetElementsByTagName("Parameters"))
            {
                coNo = GetElementByTagFromXml(parameters, "EmployerCode");
            }
            return coNo;
        }
        private string GetExcelColumnName(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnName;
        }
        public PicoXLSX.Workbook CreatePreReportWorkbook(XmlDocument xmlReport, string workBookName)
        {
            PicoXLSX.Workbook workbook = new PicoXLSX.Workbook(workBookName, "Pre");

            //Will need to return the xlsx file
            workbook = PreparePreReport(xmlReport, workbook);

            return workbook;
        }
        public PicoXLSX.Workbook CreateBottomlineReportWorkbook(XmlDocument xmlReport, string workBookName)
        {
            PicoXLSX.Workbook workbook = new PicoXLSX.Workbook(workBookName, "BACSDetails");

            //Will need to return the xlsx file
            workbook = PrepareBottomlineReport(xmlReport, workbook);

            return workbook;
        }
    }
    public class ReadConfigFile
    {
        //
        // Using XDocument instead of XmlReader
        //
        readonly string fileName = "PayescapeWGtoPR.xml";
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
        public int PeriodNo { get; set; }
        public string PaySchedule { get; set; }
        public DateTime PayRunDate { get; set; }
        public int PensionKey { get; set; }
        public bool PaidInCash { get; set; }
        public bool AOERequired { get; set; }

        public RPParameters() { }
        public RPParameters(string erRef, int taxYear, DateTime accYearStart,
                            DateTime accYearEnd, int taxPeriod, int periodNo,
                            string paySchedule, DateTime payRundate,
                            int pensionKey, bool paidInCash, bool aoeRequired)
        {
            ErRef = erRef;
            TaxYear = taxYear;
            AccYearStart = accYearStart;
            AccYearEnd = accYearEnd;
            TaxPeriod = taxPeriod;
            PeriodNo = periodNo;
            PaySchedule = paySchedule;
            PayRunDate = payRundate;
            PensionKey = pensionKey;
            PaidInCash = paidInCash;
            AOERequired = aoeRequired;
        }
    }
    //Report (RP) Employer
    public class RPEmployer
    {
        public string Name { get; set; }
        public string Address1 { get; set; }
        public string Address2 { get; set; }
        public string Address3 { get; set; }
        public string Address4 { get; set; }
        public string Postcode { get; set; }
        public string Country { get; set; }
        public string PayeRef { get; set; }
        public string HMRCDesc { get; set; }
        public string BankFileCode { get; set; }
        public string PensionReportFileType { get; set; }
        public string PensionReportAEWorkersGroup { get; set; }
        public bool P32Required { get; set; }
        public string NESTPensionText { get; set; }
        public int? HREscapeCompanyNo { get; set; }
        public string ReportPassword { get; set; }
        public bool ZipReports { get; set; }
        public bool ReportsInExcelFormat { get; set; }
        public bool PayRunDetailsYTDRequired { get; set; }
        public bool PayrollTotalsSummaryRequired { get; set; }
        public bool NoteAndCoinRequired { get; set; }
        public bool HoldPayHistory { get; set; }
        public bool CalculateApprenticeshipLevy { get; set; }

        public RPEmployer() { }
        public RPEmployer(string name, string address1, string address2, string address3,
                          string address4, string postcode, string country,
                          string payeRef, string hmrcDesc,
                          string bankFileCode,
                          string pensionReportFileType, string pensionReportAEWorkersGroup,
                          bool p32Required, string nestPensionText, int? hrEscapeCompanyNo,
                          string reportPassword, bool zipReports, bool reportsInExcelFormat,
                          bool payRunDetailsYTDRequired, bool payrollTotalsSummaryRequired,
                          bool noteAndCoinRequired, bool holdPayHistory,
                          bool calculateApprenticeshipLevy)
        {
            Name = name;
            Address1 = address1;
            Address2 = address2;
            Address3 = address3;
            Address4 = address4;
            Postcode = postcode;
            Country = country;
            PayeRef = payeRef;
            HMRCDesc = hmrcDesc;
            BankFileCode = bankFileCode;
            PensionReportFileType = pensionReportFileType;
            PensionReportAEWorkersGroup = pensionReportAEWorkersGroup; ;
            P32Required = p32Required;
            NESTPensionText = nestPensionText;
            HREscapeCompanyNo = hrEscapeCompanyNo;
            ReportPassword = reportPassword;
            ZipReports = zipReports;
            ReportsInExcelFormat = reportsInExcelFormat;
            PayRunDetailsYTDRequired = payRunDetailsYTDRequired;
            PayrollTotalsSummaryRequired = payrollTotalsSummaryRequired;
            NoteAndCoinRequired = noteAndCoinRequired;
            HoldPayHistory = holdPayHistory;
            CalculateApprenticeshipLevy = calculateApprenticeshipLevy;
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
        public DateTime StartingDate { get; set; }
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
        public decimal StudentLoanYtd { get; set; }
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
        public decimal EePensionTotalTP { get; set; }
        public decimal EePensionTotalYtd { get; set; }
        public decimal ErPensionTotalTP { get; set; }
        public decimal ErPensionTotalYtd { get; set; }
        public List<RPPensionPeriod> Pensions { get; set; }
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
        public decimal TotalOtherDedTP { get; set; }        //For the Pay Run Details YTD report there is an Other Deduction column. Include all deductions excluding Pension, Tax, NI, AEO & Student Loans
        public decimal TotalOtherDedYTD { get; set; }
        public string PensionCode { get; set; }
        public decimal PreTaxAddDed { get; set; }
        public decimal GUCosts { get; set; }
        public decimal AbsencePay { get; set; }
        public decimal AbsencePayYtd { get; set; }
        public decimal HolidayPay { get; set; }
        public decimal PreTaxPension { get; set; }
        public decimal Tax { get; set; }
        public decimal NetNI { get; set; }
        public decimal PostTaxAddDed { get; set; }
        public decimal PostTaxPension { get; set; }
        public decimal AEO { get; set; }
        public decimal AEOYtd { get; set; }
        public decimal TotalPayComponentAdditions { get; set; }
        public decimal TotalPayComponentDeductions { get; set; }
        public decimal BenefitsInKind { get; set; }
        public decimal SSPSetOff { get; set; }
        public decimal SSPAdd { get; set; }
        public decimal SMPSetOff { get; set; }
        public decimal SMPAdd { get; set; }
        public decimal OSPPSetOff { get; set; }
        public decimal OSPPAdd { get; set; }
        public decimal SAPSetOff { get; set; }
        public decimal SAPAdd { get; set; }
        public decimal ShPPSetOff { get; set; }
        public decimal ShPPAdd { get; set; }
        public decimal SPBPSetOff { get; set; }
        public decimal SPBPAdd { get; set; }
        
        public decimal Zero { get; set; }
        public string Branch { get; set; }
        public string Department { get; set; }
        public string CostCentre { get; set; }
        public List<RPAddition> Additions { get; set; }
        public List<RPDeduction> Deductions { get; set; }
        public List<RPPayslipDeduction> PayslipDeductions { get; set; }
        public RPEmployeePeriod() { }
        public RPEmployeePeriod(string reference, string title, string forename, string surname, string fullname, string refFullname, string surnameForename,
                          string address1, string address2, string address3, string address4, string postcode,
                          string country, string sortCode, string bankAccNo, DateTime dateOfBirth, DateTime startingDate, string gender, string buildingSocRef,
                          string niNumber, string paymentMethod, DateTime payRunDate, DateTime periodStartDate, DateTime periodEndDate, int payrollYear,
                          decimal gross, decimal netPayTP, decimal dayHours, DateTime? studentLoanStartDate, DateTime? studentLoanEndDate,
                          decimal studentLoan, decimal studentLoanYtd, string niLetter, string calculationBasis, decimal total,
                          decimal earningsToLEL, decimal earningsToSET, decimal earningsToPET, decimal earningsToUST, decimal earningsToAUST,
                          decimal earningsToUEL, decimal earningsAboveUEL, decimal eeContributionsPt1, decimal eeContributionsPt2,
                          decimal erNICYTD, decimal eeRebate, decimal erRebate, decimal eeReduction, DateTime leavingDate, bool leaver,
                          string taxCode, bool week1Month1, string taxCodeChangeTypeID, string taxCodeChangeType, decimal taxPrev,
                          decimal taxablePayPrevious, decimal taxThis, decimal taxablePayYTD, decimal taxablePayTP, decimal holidayAccruedTd,
                          RPAEAssessment aeAssessment,
                          decimal eePensionTotalTP, decimal eePensionTotalYtd, decimal erPensionTotalTP, decimal erPensionTotalYtd, List<RPPensionPeriod> pensions,
                          DateTime directorshipAppointmentDate, bool director, decimal eeContributionsTaxPeriodPt1, decimal eeContributionsTaxPeriodPt2,
                          decimal erNICTP, string frequency, decimal netPayYTD, decimal totalPayTP, decimal totalPayYTD, decimal totalDedTP, decimal totalDedYTD,
                          decimal totalOtherDedTP, decimal totalOtherDedYTD, string pensionCode, decimal preTaxAddDed, decimal guCosts, decimal absencePay, decimal absencePayYtd,
                          decimal holidayPay, decimal preTaxPension, decimal tax, decimal netNI,
                          decimal postTaxAddDed, decimal postTaxPension, decimal aeo, decimal aeoYtd, 
                          decimal totalPayComponentAdditions, decimal totalPayComponentDeductions, decimal benefitsInKind,
                          decimal sspSetOff, decimal sspAdd, decimal smpSetOff, decimal smpAdd, decimal osppSetOff, decimal osppAdd, decimal sapSetOff, decimal sapAdd,
                          decimal shppSetOff, decimal shppAdd, decimal spbpSetOff, decimal spbpAdd, decimal zero, string branch, string department, string costCentre,
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
            StartingDate = startingDate;
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
            StudentLoanYtd = studentLoanYtd;
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
            EePensionTotalTP = eePensionTotalTP;
            EePensionTotalYtd = eePensionTotalYtd;
            ErPensionTotalTP = erPensionTotalTP;
            ErPensionTotalYtd = erPensionTotalYtd;
            Pensions = pensions;
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
            TotalOtherDedTP = totalOtherDedTP;
            TotalOtherDedYTD = totalOtherDedYTD;
            PensionCode = pensionCode;
            PreTaxAddDed = preTaxAddDed;
            GUCosts = guCosts;
            AbsencePay = absencePay;
            AbsencePayYtd = absencePayYtd;
            HolidayPay = holidayPay;
            PreTaxPension = preTaxPension;
            Tax = tax;
            NetNI = netNI;
            PostTaxAddDed = postTaxAddDed;
            PostTaxPension = postTaxPension;
            AEO = aeo;
            AEOYtd = aeoYtd;
            TotalPayComponentAdditions = totalPayComponentAdditions;
            TotalPayComponentDeductions = totalPayComponentDeductions;
            BenefitsInKind = benefitsInKind;
            SSPSetOff = sspSetOff;
            SSPAdd = sspAdd;
            SMPSetOff = smpSetOff;
            SMPAdd = smpAdd;
            OSPPSetOff = osppSetOff;
            OSPPAdd = osppAdd;
            SAPSetOff = sapSetOff;
            SAPAdd = sapAdd;
            ShPPSetOff = shppSetOff;
            ShPPAdd = shppAdd;
            SPBPSetOff = spbpSetOff;
            SPBPAdd = spbpAdd;
            Zero = zero;
            Branch = branch;
            Department = department;
            CostCentre = costCentre;
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
                          string taxCode, bool week1Month1, int weekNumber, int monthNumber, int periodNumber,
                          decimal eeNiPaidByErAccountsAmount, decimal eeNiPaidByErAccountsUnits, decimal eeGuTaxPaidByErAccountsAmount, decimal eeGuTaxPaidByErAccountsUnits,
                          decimal eeNiLERtoUERAccountsAmount, decimal eeNiLERtoUERAccountsUnits, decimal eeNiLERtoUERPayeAmount, decimal eeNiLERtoUERPayeUnits,
                          decimal erNiAccountsAmount, decimal erNiAccountsUnits, decimal eeNiPaidByErPayeAmount,
                          decimal eeNiPaidByErPayeUnits, decimal eeGuTaxPaidByErPayeAmount, decimal eeGuTaxPaidByErPayeUnits, decimal erNiPayeAmount, decimal erNiPayeUnits,
                          List<RPPayCode> payCodes)
                          
        {
            ThisPeriodStartDate = thisPeriodStartDate;
            LastPaymentDate = lastPaymentDate;
            EeRef = eeRef;
            Branch = branch;
            CostCentre = costCentre;
            Department = department;
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
        public string ProviderName { get; set; }
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
        public DateTime? AEAssessmentDate {get;set;}
        public string AEWorkerGroup { get; set; }
        public string AEStatus { get; set; }
        public decimal TotalPayTaxPeriod { get; set; }
        public int StatePensionAge { get; set; }
        public RPPensionPeriod() { }
        public RPPensionPeriod(int key, string code, string providerName, string schemeName, DateTime? startJoinDate, bool isJoiner,
                               string providerEmployerReference,
                               decimal eePensionYtd, decimal erPensionYtd,
                               decimal pensionablePayYtd, decimal eePensionTaxPeriod, decimal erPensionTaxPeriod,
                               decimal pensionPayTaxPeriod, decimal eePensionPayRunDate, decimal erPensionPayRunDate,
                               decimal pensionablePayPayRunDate, decimal eeContributionPercent,
                               decimal erContributionPercent,
                               DateTime? aeAssessmentDate, string aeWorkerGroup, string aeStatus,
                               decimal totalPayTaxPeriod, int statePensionAge)
        {
            Key = key;
            Code = code;
            SchemeName = schemeName;
            ProviderName = providerName;
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
            AEAssessmentDate = aeAssessmentDate;
            AEWorkerGroup = aeWorkerGroup;
            AEStatus = aeStatus;
            TotalPayTaxPeriod = totalPayTaxPeriod;
            StatePensionAge = statePensionAge;
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
        public DateTime StartingDate { get; set; }
        public DateTime? LeavingDate { get; set; }
        public RPAddress RPAddress { get; set; }
        public string EmailAddress { get; set; }
        public string Gender { get; set; }
        public string NINumber { get; set; }
        public string Freq { get; set; }
        public DateTime StartDate { get; set; }
        public DateTime EndDate { get; set; }
        public DateTime PayRunDate { get; set; }
        public string SchemeFileType { get; set; }
        public string AEAssessmentOverride { get; set; }
        public RPPensionPeriod RPPensionPeriod { get; set; }

        public RPPensionContribution() { }
        public RPPensionContribution(string eeRef, string title, string forename,
                                     string surname, string fullname, string surnameForename, 
                                     string forenameSurname, DateTime dob,
                                     DateTime startingDate, DateTime? leavingDate, 
                                     RPAddress rpAddress,
                                     string emailAddress, string gender,
                                     string niNumber, string freq,
                                     DateTime startDate, DateTime endDate,DateTime payRunDate,
                                     string schemeFileType, string aeAssessmentOverride,
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
            StartingDate=startingDate;
            LeavingDate=leavingDate;
            RPAddress = rpAddress;
            EmailAddress = emailAddress;
            Gender = gender;
            NINumber = niNumber;
            Freq = freq;
            StartDate = startDate;
            EndDate = endDate;
            PayRunDate = payRunDate;
            SchemeFileType = schemeFileType;
            AEAssessmentOverride = aeAssessmentOverride;
            RPPensionPeriod = rpPensionPeriod;
        }
    }
    public class RPPensionFileScheme
    {
        public int Key { get; set; }
        public string SchemeName { get; set; }
        public string ProviderName { get; set; }
        public List<RPPensionContribution> RPPensionContributions { get; set; }

        public RPPensionFileScheme() { }
        public RPPensionFileScheme(int key, string schemeName, string providerName, List<RPPensionContribution> rpPensionContributions)
        {
            Key = key;
            SchemeName = schemeName;
            ProviderName = providerName;
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
            TaxThis = taxThis;
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
        public bool IsPayCode { get; set; }
        public RPAddition() { }
        public RPAddition(string eeRef, string code, string description, decimal rate, decimal units,
                           decimal amountTP, decimal amountYTD, decimal accountsYearBalance,
                           decimal accountsYearUnits, decimal payeYearUnits, decimal payrollAccrued,
                           bool isPayCode)
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
            IsPayCode = isPayCode;
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
        public bool IsPayCode { get; set; }
        public RPDeduction() { }
        public RPDeduction(string eeRef, string seq, string code, string description, bool isTaxable, decimal rate,
                           decimal units, decimal amountTP, decimal amountYTD, decimal accountsYearBalance,
                           decimal accountsYearUnits, decimal payeYearUnits, decimal payrollAccrued,
                           bool isPayCode)
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
            IsPayCode = isPayCode;
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
        public decimal TotalAmount{ get; set; }
        public decimal AccountsAmount { get; set; }
        public decimal PayeAmount{ get; set; }
        public decimal AccountsUnits { get; set; }
        public decimal PayeUnits { get; set; }
        public bool IsPayCode { get; set; }
        public RPPayCode() { }
        public RPPayCode(string eeRef, string code, string payCode, string description, string type, decimal totalAmount,
                         decimal accountsAmount, decimal payeAmount, decimal accountsUnits, decimal payeUnits,
                         bool isPayCode)
        {
            EeRef = eeRef;
            Code = code;
            PayCode = payCode;
            Description = description;
            Type=type;
            TotalAmount = totalAmount;
            AccountsAmount = accountsAmount;
            PayeAmount = payeAmount;
            AccountsUnits = accountsUnits;
            PayeUnits = payeUnits;
            IsPayCode = isPayCode;
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
        public decimal ApprenticeshipLevyAllowance { get; set; }
        public decimal AnnualEmploymentAllowance { get; set; }
        public bool OpeningBalancesRequired { get; set; }
        public List<RPP32ReportMonth> RPP32ReportMonths { get; set; }
        public RPP32Report() { }
        public RPP32Report(string employerName, string employerPayeRef, string paymentRef,
                                  int taxYear, DateTime taxYearStartDate, DateTime taxYearEndDate,
                                  decimal apprenticeshipLevyAllowance, decimal annualEmploymentAllowance,
                                  bool openBalancesRequired,
                                  List<RPP32ReportMonth> rpP32ReportMonths)
        {
            EmployerName = employerName;
            EmployerPayeRef = employerPayeRef;
            PaymentRef = paymentRef;
            TaxYear = taxYear;
            TaxYearStartDate = taxYearStartDate;
            TaxYearEndDate = taxYearEndDate;
            ApprenticeshipLevyAllowance = apprenticeshipLevyAllowance;
            AnnualEmploymentAllowance = annualEmploymentAllowance;
            OpeningBalancesRequired = openBalancesRequired;
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
        public decimal SpbpRecovered { get; set; }
        public decimal SpbpComp { get; set; }
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
                            decimal shppComp, decimal sapRecovered, decimal sapComp,
                            decimal spbpRecovered, decimal spbpComp, decimal totalDeductions,
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
            SpbpRecovered = spbpRecovered;
            SpbpComp = spbpComp;
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
        public int TaxPeriod { get; set; }
        public decimal IncomeTax { get; set; }
        public decimal StudentLoan { get; set; }
        public decimal PostGraduateLoan { get; set; }
        public decimal NetIncomeTax { get; set; }
        public decimal GrossNICs { get; set; }
        public RPP32PayRun() { }
        public RPP32PayRun(DateTime payDate, int payPeriod, int taxPeriod, decimal incomeTax,
                           decimal studentLoan, decimal postGraduateLoan,
                           decimal netIncomeTax, decimal grossNICs)
        {
            PayDate = payDate;
            PayPeriod = payPeriod;
            TaxPeriod = taxPeriod;
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
        public string WorkersGroup { get; set; }
        public string Status { get; set; }
        public RPAEAssessment() { }
        public RPAEAssessment(int age, int statePensionAge, DateTime? statePensionDate, DateTime? assessmentDate, decimal qualifyingEarnings,
                              string assessmentCode, string assessmentEvent, string assessmentResult,
                              string assessmentOverride, DateTime? optOutWindowEndDate, DateTime? reenrolmentDate,
                              bool isMemberOfAlternativePensionScheme, int taxYear, int taxPeriod,
                              string workersGroup, string status)
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
            WorkersGroup = workersGroup;
            Status = status;
        }
    }
    public class RPSummaryEmployee
    {
        public string Code { get; set; }
        public string LastName { get; set; }
        public string FirstName { get; set; }
        public string LastNameFirstName { get; set; }
        public string FirstNameLastName { get; set; }
        public string Branch { get; set; }
        public string Department { get; set; }
        public string TaxCode { get; set; }
        public string TaxBasis { get; set; }
        public string NiLetter { get; set; }
        public decimal PreTaxAddDed { get; set; }
        public decimal GUCosts { get; set; }
        public decimal AbsencePay { get; set; }
        public decimal HolidayPay { get; set; }
        public decimal PreTaxPension { get; set; }
        public decimal TaxablePay { get; set; }
        public decimal Tax { get; set; }
        public decimal NetEeNi { get; set; }
        public decimal PostTaxAddDed { get; set; }
        public decimal PostTaxPension { get; set; }
        public decimal AEO { get; set; }
        public decimal StudentLoan { get; set; }
        public RPNetPayCashAnalysis NetPay { get; set; }
        public decimal ErNi { get; set; }
        public decimal ErPension { get; set; }
        public string PaymentType { get; set; }

        public RPSummaryEmployee() { }
        public RPSummaryEmployee(string code, string lastrName, string firstName, string lastNameFirstName, string firstNameLastName,
                        string branch, string department,
                        string taxCode, string taxBasis, string niLetter,
                        decimal preTaxAddDed, decimal guCosts, decimal absencePay, decimal holidayPay, decimal preTaxPension,
                        decimal taxablePay, decimal tax, decimal netEeNi, decimal postTaxAddDed,
                        decimal postTaxPension, decimal aeo, decimal studentLoan, RPNetPayCashAnalysis netPay,
                        decimal erNi, decimal erPension, string paymentType)
        {
            Code = code;
            LastName = lastrName;
            FirstName = firstName;
            LastNameFirstName = lastNameFirstName;
            FirstNameLastName = firstNameLastName;
            Branch = branch;
            Department = department;
            TaxCode = taxCode;
            TaxBasis = taxBasis;
            NiLetter = niLetter;
            PreTaxAddDed = preTaxAddDed;
            GUCosts = guCosts;
            AbsencePay = absencePay;
            HolidayPay = holidayPay;
            PreTaxPension = preTaxPension;
            TaxablePay = taxablePay;
            Tax = tax;
            NetEeNi = netEeNi;
            PostTaxAddDed = postTaxAddDed;
            PostTaxPension = postTaxPension;
            AEO = aeo;
            StudentLoan = studentLoan;
            NetPay = netPay;
            ErNi = erNi;
            ErPension = erPension;
            PaymentType = paymentType;
        }
    }
    public class RPNetPayCashAnalysis
    {
        public Decimal NetPay { get; set; }
        public int TwentyPounds { get; set; }
        public int TenPounds { get; set; }
        public int FivePounds { get; set; }
        public int TwoPounds { get; set; }
        public int OnePounds { get; set; }
        public int FiftyPence { get; set; }
        public int TwentyPence { get; set; }
        public int TenPence { get; set; }
        public int FivePence { get; set; }
        public int TwoPence { get; set; }
        public int OnePence { get; set; }

        public RPNetPayCashAnalysis() { }
        public RPNetPayCashAnalysis(decimal netPay, int twentyPounds, int tenPounds, int fivePounds,
                                  int twoPounds, int onePounds,
                                  int fiftyPence, int twentyPence, int tenPence, int fivePence,
                                  int twoPence, int onePence)
        {
            NetPay = netPay;
            TwentyPounds = twentyPounds;
            TenPounds = tenPounds;
            FivePounds = fivePounds;
            TwoPounds = twoPounds;
            OnePounds = onePounds;
            FiftyPence = fiftyPence;
            TwentyPence = twentyPence;
            TenPence = tenPence;
            FivePence = fivePence;
            TwoPence = twoPence;
            OnePence = onePence;
        }
    }
    public class RPSummaryPayRuns
    {
        public int MaxDepartments { get; set; }
        public List<RPSummaryPayRun> RPSummaryPayRun { get; set; }
        public RPSummaryPayRuns() { }
        public RPSummaryPayRuns(int maxDepartments, List<RPSummaryPayRun> rpSummaryPayRun)
        {
            MaxDepartments = maxDepartments;
            RPSummaryPayRun = rpSummaryPayRun;
        }
    }
    public class RPSummaryPayRun
    {
        public DateTime PaymentDate { get; set; }
        public int TaxPeriod { get; set; }
        public int TaxYear { get; set; }
        public int PAYEMonth { get; set; }
        public List<RPBranch> RPBranches { get; set; }
        public RPSummaryPayRun() { }
        public RPSummaryPayRun(DateTime paymentDate, int taxPeriod, int taxYear,
                      int payeMonth, List<RPBranch> rpBranches)
        {
            PaymentDate = paymentDate;
            TaxPeriod = taxPeriod;
            TaxYear = taxYear;
            RPBranches = rpBranches;
            PAYEMonth = payeMonth;
        }
    }

    public class RPBranch
    {
        public string Name { get; set; }
        public List<RPDepartment> RPDepartments { get; set; }

        public RPBranch() { }
        public RPBranch(string name, List<RPDepartment> rpDepartments)
        {
            Name = name;
            RPDepartments = rpDepartments;
        }
    }
    public class RPDepartment
    {
        public string Name { get; set; }
        public List<RPSummaryEmployee> Employees { get; set; }

        public RPDepartment() { }
        public RPDepartment(string name, List<RPSummaryEmployee> employees)
        {
            Name = name;
            Employees = employees;

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
        public string RemoveNonAlphaNumericChars(string text)
        {
            Regex alphaNumeric = new Regex("[^a-zA-Z0-9]");
            text = alphaNumeric.Replace(text, "");
            return text;
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
