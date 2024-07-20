using Microsoft.Office.Interop.Excel;
using System;
using System.Runtime.InteropServices;

public class ExcelTest
{
    public static bool IsExcelWorking(string excelFilePath)
    {
        try
        {
            // Create a new Excel application object
            Application excelApp = new Application();

            // Open the Excel file
            Workbook workbook = excelApp.Workbooks.Open(excelFilePath);

            // Close the Excel application
            excelApp.Quit();

            // Release the COM objects
            Marshal.ReleaseComObject(workbook);
            Marshal.ReleaseComObject(excelApp);

            // Return true if no exception was thrown
            return true;
        }
        catch (Exception ex)
        {
            // Handle the exception, e.g., log it
            Console.WriteLine("Error opening Excel file: " + ex.Message);
            return false;
        }
    }

    public static void Main(string[] args)
    {
        // string excelFilePath = @"a.xlsx"; // Replace with your Excel file path
        // Console.WriteLine("start.");

        // if (IsExcelWorking(excelFilePath))
        // {
        //     Console.WriteLine("Excel file is working.");
        // }
        // else
        // {
        //     Console.WriteLine("There was a problem opening the Excel file.");
        // }
        string excelFilePath = @"x.xlsx"; // Replace with your Excel file path
        Console.WriteLine("start.");
        Console.WriteLine("start2.");
        Console.WriteLine("Object3");
        try
        {

            Console.WriteLine("Object1");
            // Create a new Excel application object
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();

            Console.WriteLine("Object");
            // Open the Excel file
            Workbook workbook = excelApp.Workbooks.Open(excelFilePath);

            // Close the Excel application
            excelApp.Quit();

            // Release the COM objects
            Marshal.ReleaseComObject(workbook);
            Marshal.ReleaseComObject(excelApp);

            // Return true if no exception was thrown
            Console.WriteLine("no error");
        }
        catch (Exception ex)
        {
            // Handle the exception, e.g., log it
            Console.WriteLine("Error opening Excel file: " + ex.Message);
        }
         Console.WriteLine("Press any key to exit...");  
    Console.ReadKey();  
    }
}