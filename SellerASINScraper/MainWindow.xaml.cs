using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;


namespace SellerASINScraper
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private bool IsElementPresent(By by, IWebDriver driver) 
        {                                                       
            try
            {
                driver.FindElement(by);
                return true;
            }
            catch (NoSuchElementException)
            {
                return false;
            }
            catch (ElementNotVisibleException)
            {
                return false;
            }

        }
        public void Run_Click(object sender, RoutedEventArgs e) //Run button!
        {

            Close();
            var driverService = ChromeDriverService.CreateDefaultService(@"\\brmpro\MACAPPS\ClickOnce\CustomerServiceAutomationTool");
            driverService.HideCommandPromptWindow = true;
            var driver = new ChromeDriver(driverService, new ChromeOptions());
            
            string sellerID = SellerIDx.Text;

            driver.Url = "https://www.amazon.com/s?me=" + sellerID + "&rh=p_4%3ABob%27s+Red+Mill";


            Excel excel = new Excel();
            excel.CreateNewFile();
            //excel.CreateNewSheet();
           

            var cellCount = 1;
            var pageItemCount = 1;
            var totalItemCount = 1; //exists because pageItemCount will be reset every new page
            var pageNumber = 2;
            var sellerName = driver.FindElement(By.XPath("//*[@id='search']/span[2]/h1/div/div[1]/div/div/a/span")).Text;

            while (IsElementPresent(By.XPath("//*[@id='search']/div[1]/div[2]/div/span[7]/div/span/div/div/ul/li[" + pageNumber + "]"), driver) == true)
            { 

                while (IsElementPresent(By.XPath("//*[@id='search']/div[1]/div[2]/div/span[3]/div[1]/div[" + pageItemCount + "]"), driver) == true)
                {
                    string result = driver.FindElement(By.XPath("//*[@id='search']/div[1]/div[2]/div/span[3]/div[1]/div[" + pageItemCount + "]")).GetAttribute("data-asin");
                    excel.WriteToCell(cellCount, 1, result);

                    cellCount++;
                    totalItemCount++;
                    pageItemCount++;
                }


                pageItemCount = 1;
                pageNumber++;

                if (IsElementPresent(By.XPath("//*[@id='search']/div[1]/div[2]/div/span[7]/div/span/div/div/ul/li[" + pageNumber + "]"), driver) == true)
                    {
                    var clickPG = driver.FindElement(By.XPath("//*[@id='search']/div[1]/div[2]/div/span[7]/div/span/div/div/ul/li[" + pageNumber + "]"));
                    clickPG.Click();
                    }

            }


            excel.SaveAs(@"\\brmserver\company\eComm\SellerASINScraper\Reports\ASIN Report for " + sellerID + " - " + sellerName);
            excel.Quit();
            driver.Quit();
        }
    }
}
