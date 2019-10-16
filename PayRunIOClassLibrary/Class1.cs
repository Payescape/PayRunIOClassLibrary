using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Globalization;
using System.Xml.Linq;
using System.IO;

namespace PayRunIOClassLibrary
{
    public class PayRunIOWebGlobeClass
    {
        public PayRunIOWebGlobeClass() { }

        public FileInfo[] GetAllCompletedPayrollFiles(XDocument xdoc)
        {
            string path = xdoc.Root.Element("DataHomeFolder").Value + "Outputs";
            DirectoryInfo folder = new DirectoryInfo(path);
            FileInfo[] files = folder.GetFiles("*CompletedPayroll*.xml");

            return files;
        }

    }
    public class ReadConfigFile
    {
        string change = null;
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

        public RPEmployer() { }
        public RPEmployer(string name, string payeRef)
        {
            Name = name;
            PayeRef = payeRef;


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
        public int DayHours { get; set; }
        public DateTime StudentLoanStartDate { get; set; }
        public DateTime StudentLoanEndDate { get; set; }
        public decimal StudentLoan { get; set; }
        public string NILetter { get; set; }
        public string CalculationBasis { get; set; }
        public decimal TotalPayTP { get; set; }
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
        public DateTime LeavingDate { get; set; }
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
        public decimal HolidayAccruedYTD { get; set; }
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
        public decimal EeContributionTaxPeriodPt1 { get; set; }
        public decimal EeContributionTaxPeriodPt2 { get; set; }
        public decimal ErNICTP { get; set; }
        public string Frequency { get; set; }
        public decimal NetPayYTD { get; set; }
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
        public RPEmployeePeriod(string reference, string title, string forename, string surname, string fullname, string refFullname,
                          string address1, string address2, string address3, string address4, string postcode,
                          string country, DateTime dateOfBirth, string gender, bool leaver, DateTime leavingDate,
                          string niNumber, string niLetter, string taxCode, bool week1Month1, string frequency,
                          string paymentMethod, DateTime payRunDate,
                          decimal netPayTP, decimal netPayYTD, decimal taxablePayTP, decimal taxablePayYTD,
                          decimal taxablePayPrevious, decimal totalPayTP, decimal totalPayYTD, decimal totalDedTP, decimal totalDedYTD,
                          decimal erNICTP, decimal erNICYTD, decimal erPensionTP, decimal eePensionTP, decimal erPensionYTD,
                          decimal eePensionYTD, decimal pensionablePay, string pensionCode, string sortCode, string bankAccNo, string buildingSocRef,
                          decimal erContributionPercent, decimal preTaxAddDed, decimal guCosts, decimal absencePay,
                          decimal holidayPay, decimal preTaxPension, decimal tax, decimal taxPrev, decimal taxThis, decimal netNI,
                          decimal postTaxAddDed, decimal postTaxPension, decimal aoe, decimal studentLoan,
                          decimal eeContributionPercent, List<RPAddition> additions, List<RPDeduction> deductions)
        {
            Reference = reference;
            Title = title;
            Forename = forename;
            Surname = surname;
            Fullname = fullname;
            RefFullname = refFullname;
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
            PayRunDate = PayRunDate;
            //PeriodStartDate
            //PeriodEndDate
            //PayrollYear
            //Gross
            NetPayTP = netPayTP;
            //DayHours
            //StudentLoanStartDate
            //StundentLoanEndDate
            StudentLoan = studentLoan;
            NILetter = niLetter;
            //CalculationBasis
            TotalPayTP = totalPayTP;
            //EarningsToLEL
            //EarningsToSET
            //EarningsToPET
            //EarningsToUST
            //EarningsToAUST
            //EarningsToUEL
            //EarningsAboveUel
            //EeContributionsPt1
            //EeContributionsPt2
            ErNICYTD = erNICYTD;
            //EeRebate
            //ErRebate
            //EeReduction
            LeavingDate = leavingDate;
            Leaver = leaver;
            TaxCode = taxCode;
            Week1Month1 = week1Month1;
            //TaxCodeChangeTypeID
            //TaxCodeChangeType
            TaxPrev = taxPrev;
            TaxablePayPrevious = taxablePayPrevious;
            TaxThis = taxThis;
            TaxablePayYTD = taxablePayYTD;
            TaxablePayTP = taxablePayTP;
            //Holiday AccruedTd
            ErPensionYTD = erPensionYTD;
            EePensionYTD = eePensionYTD;
            ErPensionTP = erPensionTP;
            EePensionTP = eePensionTP;
            ErContributionPercent = erContributionPercent;
            EeContributionPercent = eeContributionPercent;
            PensionablePay = pensionablePay;
            //ErPensionPayRunDate
            //EePensionPayRunDate
            //DirectorshipAppointmentDate
            //Director
            //EeContributionsTaxPeriodPt1
            //EeContributionsTaxPeriodPt2
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
        public string Description { get; set; }
        public decimal Rate { get; set; }
        public decimal Units { get; set; }
        public decimal AmountTP { get; set; }
        public decimal AmountYTD { get; set; }
        public RPAddition() { }
        public RPAddition(string eeRef, string description, decimal rate, decimal units,
                           decimal amountTP, decimal amountYTD)
        {
            EeRef = eeRef;
            Description = description;
            Rate = rate;
            Units = units;
            AmountTP = amountTP;
            AmountYTD = amountYTD;

        }
    }

    //Report (RP) Deductions
    public class RPDeduction
    {
        public string EeRef { get; set; }
        public string Description { get; set; }
        public decimal AmountTP { get; set; }
        public decimal AmountYTD { get; set; }
        public RPDeduction() { }
        public RPDeduction(string eeRef, string description,
                           decimal amountTP, decimal amountYTD)
        {
            EeRef = eeRef;
            Description = description;
            AmountTP = amountTP;
            AmountYTD = amountYTD;

        }
    }
    public class RPPayComponent
    {
        public string PayCode { get; set; }
        public string Description { get; set; }
        public string EeRef { get; set; }
        public string Fullname { get; set; }
        public string Surname { get; set; }
        public decimal Rate { get; set; }
        public decimal UnitsTP { get; set; }
        public decimal AmountTP { get; set; }
        public decimal UnitsYTD { get; set; }
        public decimal AmountYTD { get; set; }
        public RPPayComponent() { }
        public RPPayComponent(string payCode, string description, string eeRef, string fullname,
                              string surname, decimal rate, decimal unitsTP, decimal amountTP,
                               decimal unitsYTD, decimal amountYTD)
        {
            PayCode = payCode;
            Description = description;
            EeRef = eeRef;
            Fullname = fullname;
            Surname = surname;
            Rate = rate;
            UnitsTP = unitsTP;
            AmountTP = amountTP;
            UnitsYTD = unitsYTD;
            AmountYTD = amountYTD;

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
