using System;
using Serilog;
using log4net;
using System.Collections;
using System.Collections.Generic;
using System.Data.Entity;
using System.Dynamic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using HotelManagementSystem.Areas.Admin.ViewModel;
using HotelManagementSystem.Models;
using HotelManagementSystem.ViewModels;
using System.Security.Cryptography;
using System.Net.Mail;
using System.DirectoryServices.Protocols;
using Microsoft.Office.Interop.Excel;
using CrystalDecisions.CrystalReports.Engine;
using System.Data.OleDb; 
using System.Diagnostics.DiagnosticSource;
using System.Messaging;
using System.Security.Principal;

namespace HotelManagementSystem.Controllers
{
    public class HomeController : Controller
    {
        private readonly ApplicationDbContext _context = new ApplicationDbContext();

        public ActionResult Index(int? accomodationTypeId)
        {
           var loginInfo = Session["LoginInfo"] as LoginResponse;
            if (accomodationTypeId == null)
            {
                var model = new HomeViewmodel()
                {
                    AccomodationTypes = _context.AccomodationTypes.ToList(),
                    AccomodationPackages = _context.AccomodationPackages.ToList()
                };
                return View(model);
            }
            else
            {
                var model = new HomeViewmodel()
                {
                    AccomodationTypes = _context.AccomodationTypes.ToList(),
                    AccomodationPackages = _context.AccomodationPackages.Where(a=>a.AccomodationTypeId==accomodationTypeId).ToList()
                };
                return View(model);
            }
            
            
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }

        public ActionResult AccomodationPackageDetails(int accomodationpackageid)
        {
            var accomodationPackage = _context.AccomodationPackages.Include(a=>a.AccomodationType).FirstOrDefault(a => a.Id == accomodationpackageid);
            accomodationPackage.Pictures =
                _context.Pictures.Where(p => p.AccomodationPackageId == accomodationpackageid).ToList();
            return View(accomodationPackage);
        }

        public static bool Authenticate(string username, string password)
    {
        // Perform authentication logic here
        bool isAuthenticated = YourAuthenticationMethod(username, password);

        if (isAuthenticated)
        {
            var identity = new PassportIdentity(username, isAuthenticated);
            var principal = new PassportPrincipal(identity);

            // Set the principal in the current context
            System.Threading.Thread.CurrentPrincipal = principal;
            if (System.Web.HttpContext.Current != null)
            {
                System.Web.HttpContext.Current.User = principal;
            }

            return true;
        }

        return false;
    }

    private static bool YourAuthenticationMethod(string username, string password)
    {
        // Implement your actual authentication logic here
        // For simplicity, this example always returns true
        return true;
    }
        public void callAPI()
        {
            using (var client = new WebClient())
            {
                using (Aes encryptor = Aes.Create())
                {
                    SmtpClient smtp = new SmtpClient();
                    var loginInfo = Session["LoginInfo"] as LoginResponse;
                      var isAdmin = new WindowsPrincipal(WindowsIdentity.GetCurrent()).IsInRole(WindowsBuiltInRole.Administrator);
                }
            }
        }


public void provider()
    {

 private string connectionString;

    public AccessDatabaseHelper(string databasePath)
    {
        // Set up the connection string with ACE provider
        connectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={databasePath};Persist Security Info=False;";
    }

    public DataTable ReadDataFromTable(string tableName)
    {
        DataTable dataTable = new DataTable();

        using (OleDbConnection connection = new OleDbConnection(connectionString))
        {
            try
            {
                connection.Open();

                // Query to select all data from the specified table
                string query = $"SELECT * FROM {tableName}";

                using (OleDbDataAdapter adapter = new OleDbDataAdapter(query, connection))
                {
                    // Fill the DataTable with data from the Access database
                    adapter.Fill(dataTable);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
            finally
            {
                connection.Close();
            }
        }

        return dataTable;
    }



        
    }

    public void provider11()
    {
 string databasePath = @"C:\Path\To\Your\Database.accdb";
        string tableName = "YourTableName";

        AccessDatabaseHelper databaseHelper = new AccessDatabaseHelper(databasePath);

        DataTable resultTable = databaseHelper.ReadDataFromTable(tableName);

        // Process the data in the DataTable as needed
        foreach (DataRow row in resultTable.Rows)
        {
            Console.WriteLine($"{row["ColumnName1"]}, {row["ColumnName2"]}, ...");
        }
    }
 public void comComponentPat()
    {
        try
        {
            // Specify the path to your COM component executable
            string comComponentPath = @"C:\Path\To\Your\COMComponent.exe";

            // Use Process.Start to launch the COM component
            Process.Start(comComponentPath);

            Console.WriteLine("COM Component started successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }


public void tracing()
    {
        // Set up a trace listener (e.g., writing to a console)
        Trace.Listeners.Add(new ConsoleTraceListener());

        // Start tracing
        Trace.WriteLine("Application Started");

        // Your code logic
        int result = Add(3, 4);

        // Trace the result
        Trace.WriteLine($"Result of the addition: {result}");

        // End tracing
        Trace.WriteLine("Application Ended");
    }

    static int Add(int a, int b)
    {
        // Trace entering a method
        Trace.WriteLine($"Entering Add method with parameters {a} and {b}");

        int sum = a + b;

        // Trace leaving a method
        Trace.WriteLine($"Leaving Add method with result {sum}");

        return sum;
    }

         static void Provider()
    {
        string connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\\Path\\To\\Your\\Database.mdb";
        
        using (OleDbConnection connection = new OleDbConnection(connectionString))
        {
            try
            {
                connection.Open();

                // Perform your database operations here

                Console.WriteLine("Connection successful!");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
            finally
            {
                connection.Close();
            }
        }
    }
         public void ModifyExcel(string filePath)
    {
        // Create an Excel application instance
        Excel.Application excelApp = new Excel.Application();

        try
        {
            // Open the Excel workbook
            Excel.Workbook workbook = excelApp.Workbooks.Open(filePath);

            // Get the first worksheet
            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Worksheets[1];

            // Modify the Excel data (example: update cell A1)
            Excel.Range cellA1 = worksheet.Cells[1, 1];
            cellA1.Value = "New Value";

            // Save the changes
            workbook.Save();

            // Close the workbook and release resources
            workbook.Close();
        }
        catch (Exception ex)
        {
            // Handle exceptions
            Console.WriteLine("Error: " + ex.Message);
        }
        finally
        {
            // Quit the Excel application
            excelApp.Quit();

            // Release COM objects to avoid memory leaks
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
        }
        }

        public void officeintrope()
        {
            Application excelApp = null;

        try
        {
            // Attempt to get an existing instance of Excel
            excelApp = (Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
        }
        catch (Exception)
        {
            // Excel is not running, or it's not registered as an active object
            Console.WriteLine("Excel is not running. Please start Excel and run the program again.");
            return;
        }
        }

    public void loging()
    {
        // Configure Serilog to write logs to the console
        Log.Logger = new LoggerConfiguration()
            .WriteTo.Console()
            .CreateLogger();

        try
        {
            Log.Information("Application Started");

            // Your code logic
            int result = Add(3, 4);

            Log.Information($"Result of the addition: {result}");
        }
        catch (Exception ex)
        {
            // Log any exceptions
            Log.Error(ex, "An error occurred");
        }
        finally
        {
            Log.Information("Application Ended");

            // Close and flush the log
            Log.CloseAndFlush();
        }
    }

    static int Add(int a, int b)
    {
        try
        {
            Log.Debug($"Entering Add method with parameters {a} and {b}");

            int sum = a + b;

            Log.Debug($"Leaving Add method with result {sum}");

            return sum;
        }
        catch (Exception ex)
        {
            // Log any exceptions in the method
            Log.Error(ex, "An error occurred in the Add method");
            throw; // Re-throw the exception
        }
    }

        public void Crystalreport()
        {
             ReportDocument report = new ReportDocument();
            report.Load("YourReportFile.rpt");  // Replace with your actual report file path

            // Set the data source
            DataTable dataTable = GetSampleData();  // Replace with your data retrieval logic
            report.SetDataSource(dataTable);

            // Set the report to the CrystalReportViewer
            crystalReportViewer1.ReportSource = report;
        }

        public void message()
    {
        // Create a message queue instance
        MessageQueue myQueue = new MessageQueue(".\\private$\\MyQueue");

        // Set the formatter to read the message body as a string
        myQueue.Formatter = new XmlMessageFormatter(new Type[] { typeof(string) });

        // Receive a message from the queue
        Message myMessage = myQueue.Receive();

        // Display the message body
        Console.WriteLine($"Received Message: {myMessage.Body}");
    }

    public void Messagesend()
    {
        // Create a message queue instance
        MessageQueue myQueue = new MessageQueue(".\\private$\\MyQueue");

        // Create a message and set its body
        Message myMessage = new Message();
        myMessage.Body = "Hello, Message Queue!";

        // Send the message to the queue
        myQueue.Send(myMessage);

        Console.WriteLine("Message sent successfully.");
    }

        public void JETProvider()
    {
        string connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\\Path\\To\\Your\\Database.mdb";
        
        using (OleDbConnection connection = new OleDbConnection(connectionString))
        {
            try
            {
                connection.Open();

                // Perform your database operations here

                Console.WriteLine("Connection successful!");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
            finally
            {
                connection.Close();
            }
        }
    }
        
    }
}