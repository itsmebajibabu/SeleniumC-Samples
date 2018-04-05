using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Interactions;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace CSharpDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                DataTable dt = Common.ReadExcel(@"C:\Users\Admin\Desktop\test.xlsx");
                DataTable d1 = new DataTable();
                d1.Columns.Add("UserId", typeof(string));
                d1.Columns.Add("Password", typeof(string));
                d1.Columns.Add("Result", typeof(string));

                foreach (DataRow drr in dt.Rows)
                {
                    DataRow dr = d1.NewRow();
                    dr[0] = drr[1].ToString();
                    dr[1] = drr[2].ToString();
                    string baseurl = "http://orangehrm.qedgetech.com/symfony/web/index.php/auth/login";
                    var driver = new ChromeDriver();
                    driver.Url = baseurl;
                    driver.FindElement(By.Id("txtUsername")).SendKeys(drr[1].ToString());
                    driver.FindElement(By.Id("txtPassword")).SendKeys(drr[2].ToString());
                    driver.FindElement(By.Id("btnLogin")).Click();
                    Thread.Sleep(1500);

                    if (driver.Url.Equals("http://orangehrm.qedgetech.com/symfony/web/index.php/dashboard"))
                    {
                        Console.WriteLine("{0} | {1} | Passed", drr[1], drr[2]);
                        dr[2] = "Passed";
                    }
                    else
                    {
                        Console.WriteLine("{0} | {1} | Failed", drr[1], drr[2]);
                        dr[2] = "Failed";
                        Common.ScreenShot(driver);
                    }
                    Thread.Sleep(1500);
                    driver.Close();
                    d1.Rows.Add(dr);
                }
                bool x = Common.WriteToExcel(d1, @"C:\Users\Admin\Desktop\testOutput.xlsx");
                if (x)
                {
                    Console.WriteLine("Result Saved");
                }
                else
                {
                    Console.WriteLine("Result Unable to Save");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                Console.ReadKey();
            }
        }
    }
}
