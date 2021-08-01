
using System.IO;
using System;
using System.Linq;
using Microsoft.Office.Interop.Excel;
using GeneratePayslip.Model;
using System.Data.Entity;
using System.Diagnostics;
using System.Globalization;
using GeneratePayslip.Resources.Constants;
using System.Net.Mail;
using System.Net;
using System.Web;
using PdfSharp.Pdf;
using PdfSharp.Pdf.Security;
using PdfReader = PdfSharp.Pdf.IO.PdfReader;

//using iTextSharp.text.pdf;


namespace GeneratePayslip
{
    class Program
    {

        public static void Main(string[] args)
        {
            try
            {

                foreach (Process proc in Process.GetProcessesByName("Excel"))
                {
                    proc.Kill();
                }

                foreach (Process proc in Process.GetProcessesByName("Adobe Acrobat Reader DC"))
                {
                    proc.Kill();
                }


                #region A Code to create a payslip

                PaymentModel _context = new PaymentModel();


                var employeeList = _context.Employees.Include(p => p.Payslips).ToList();

                var empList = (from em in _context.Employees
                    join pay in _context.Payslips
                        on em.EmployeeId equals pay.EmployeeId
                    select new
                    {
                        em,
                        pay,
                    }).ToList();


                for (int i = 0; i < empList.Count; i++)
                {
                    int employeeId = employeeList[i].EmployeeId;

                    string name = employeeList[i].FullName;
                    string jobTitle = employeeList[i].JobTitle;
                    string department = employeeList[i].Department;
                    string idNumber = employeeList[i].IdNumber;
                    string email = employeeList[i].Email;


                    string bankName = employeeList[i].BankName;
                    string bankAccountNo = employeeList[i].BankAccountNo;
                    string accountHolder = employeeList[i].AccountHolder;

                    var basicSalary = empList[i].pay.BasicSalary;


                    if (i < employeeList.Count)
                    {
                       // string directory = "C:\\Users\\Noncedo\\source\\repos\\GeneratePayslip\\GeneratePayslip\\";
                       string directory = AppDomain.CurrentDomain.BaseDirectory;

                        string sourceFolder = directory + @"Template\Copy of Template 2.xlsx";

                        //"
                        string fileName = name + "_" + DateTime.Now.ToString("yyyy-MM-dd") + ".xlsx";


                        var excelApp = new Application();



                        string destinationFolder = directory + "Payslips";

                        //string pdfFolder = directory + "\\PDF Payslips\\";



                        destinationFolder = System.IO.Path.Combine(destinationFolder, fileName);


                        if (File.Exists(destinationFolder))
                        {
                            File.Delete(destinationFolder);
                        }

                        //copying file to Destination folder
                        File.Copy(sourceFolder, destinationFolder, true);

                        excelApp.Workbooks.Open(destinationFolder,
                            Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing,
                            true, Type.Missing,
                            Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing);

                        var ws = excelApp.Worksheets;
                        var worksheet = (Worksheet) ws.get_Item("Sheet1");



                        Range range = worksheet.UsedRange;

                        if (employeeList[i].CreatedDate != null)
                        {
                            DateTime paymentDate = employeeList[i].CreatedDate.Value;
                        }

                        //DateTime currentDate = new DateTime("2020-01-01");
                        //var twentyFive =currentDat

                       


                        range.Cells.set_Item(7, 4, name);
                        range.Cells.set_Item(7, 7, jobTitle);
                        range.Cells.set_Item(8, 7, department);

                        range.Cells.set_Item(24, 4, bankName);
                        range.Cells.set_Item(25, 4, accountHolder);
                        range.Cells.set_Item(26, 4, bankAccountNo);

                        range.Cells.set_Item(13, 5, basicSalary);



                        int monthNumber = int.Parse(DateTime.Now.ToString("MM"));

                        var year = DateTime.Now.ToString("yyyy");

                        var monthName = CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(monthNumber);

                        string dateNow = monthName + " " + year;


                        range.Cells.set_Item(4, 4, dateNow);




                        // Disable the compatibility checker
                        excelApp.DisplayAlerts = false;

                        excelApp.DisplayAlerts = true;


                        #endregion

                        #region A Code to Convert payslip to PDF format

                        // var pdfDoc = fileName.Replace(".xls", ".pdf");

                        //excelApp.Workbooks.Close();

                        var thisFileWorkbook = excelApp.Application.Workbooks.Open(destinationFolder,Type.Missing
                        ,Type.Missing,Type.Missing
                        ,"password",Type.Missing
                        ,Type.Missing,Type.Missing
                        ,Type.Missing,false,
                        Type.Missing);

                     

                        string newPdfFilePath = Path.Combine(

                            Path.GetDirectoryName(destinationFolder),

                            $"{Path.GetFileNameWithoutExtension(destinationFolder)}.pdf");



                        thisFileWorkbook.Save();

                        thisFileWorkbook.ExportAsFixedFormat(

                            Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF,

                            newPdfFilePath);


                        thisFileWorkbook.Close(false, destinationFolder);

                       
                        #endregion

                        #region Send Email
                        SendEmail(email, newPdfFilePath,idNumber);

                        File.Delete(destinationFolder);

                        //foreach (Process proc in Process.GetProcessesByName("Mail"))
                        //{
                        //    proc.Kill();
                        //}


                       // File.Delete(newPdfFilePath);

                        #endregion
                    }

                }

            }
            
            catch (Exception e)
            {
               // var result = "Documents not generated ";

                StreamWriter sw = null;
                sw = new StreamWriter(AppDomain.CurrentDomain.BaseDirectory + "\\LogFile.txt", true);
                sw.WriteLine(DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss")  + "\n" + e);
                sw.Flush();
                sw.Close();

            }
           
        }

        public static void SendEmail(string email, string attachment,string password)

        {
            try
            {
                //senders gmail address
                string SendersAddress = EmailInfo.FROM_EMAIL_ACCOUNT;
                //receiver address
                string ReceiversAddress = email;
                //senders password
                string SendersPassword = EmailInfo.FROM_EMAIL_PASSWORD;
                //Email subject
                const string subject = "Copy of your Payslip";


                SmtpClient smtp = new SmtpClient
                {
                    UseDefaultCredentials = false,
                    Host = EmailInfo.SMTP_HOST_OUTLOOK,
                    //use port 25 if 587 is blocked
                    Port = 587,
                    EnableSsl = true,
                    DeliveryMethod = SmtpDeliveryMethod.Network,
                    Credentials = new NetworkCredential(SendersAddress, SendersPassword),
                    //Timeout = 3000
                };

                //set the message
                MailMessage message = new MailMessage(SendersAddress, ReceiversAddress, subject,
                    "Attached, Please find the copy of your payslip. The password to open the document is your Id number.");

               
                PdfDocument document = PdfReader.Open(attachment);

                PdfSecuritySettings securitySettings = document.SecuritySettings;

                // Setting one of the passwords automatically sets the security level to 
                // PdfDocumentSecurityLevel.Encrypted128Bit.
                securitySettings.UserPassword = password;
                securitySettings.OwnerPassword = password;

                // Don't use 40 bit encryption unless needed for compatibility reasons
                //securitySettings.DocumentSecurityLevel = PdfDocumentSecurityLevel.Encrypted40Bit;

                // Restrict some rights.
                securitySettings.PermitAccessibilityExtractContent = false;
                securitySettings.PermitAnnotations = false;
                securitySettings.PermitAssembleDocument = false;
                securitySettings.PermitExtractContent = false;
                securitySettings.PermitFormsFill = true;
                securitySettings.PermitFullQualityPrint = false;
                securitySettings.PermitModifyDocument = true;
                securitySettings.PermitPrint = false;

                // Save the document...
                document.Save(attachment);


                message.Attachments.Add(new Attachment(attachment));

                //use smtp sever we specified above to send the message(MailMessage message)
                smtp.Send(message);


                var result = "Email sent successfully to: ";

                result = result + " " + email;
                StreamWriter sw = null;
                sw = new StreamWriter(AppDomain.CurrentDomain.BaseDirectory + "\\LogFile.txt", true);
                sw.WriteLine(DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + ": "+ result);
                sw.Flush();
                sw.Close();


            }
            catch (Exception e)
            {
                var result = "Email not sent to: ";
                result = result + " " + email;
                
                StreamWriter sw = null;
                sw = new StreamWriter(AppDomain.CurrentDomain.BaseDirectory + "\\LogFile.txt", true);
                sw.WriteLine(DateTime.Now.ToString("yyyy-MM-dd") + ": " + result+"\n" + e );
                sw.Flush();
                sw.Close();
            }




        }


    }
}
   
