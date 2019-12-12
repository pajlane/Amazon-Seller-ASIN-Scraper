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
            var driverService = ChromeDriverService.CreateDefaultService(@"\\brmserver\company\eComm\SellerASINScraper");
            driverService.HideCommandPromptWindow = true;
            var driver = new ChromeDriver(driverService, new ChromeOptions());
            
            string sellerID = SellerIDx.Text;

            driver.Url = "https://www.amazon.com/s?me=" + sellerID + "&rh=p_4%3ABob%27s+Red+Mill";


            Excel excel = new Excel();
            excel.CreateNewFile();
            //excel.CreateNewSheet();
           

            var cellCount = 1;
            var pageItemCount = 1;
            var pageNumber = 2;
            int countPageNumber = 0;
            var checker = "";
            var sellerName = driver.FindElement(By.XPath("//*[@id='search']/span[2]/h1/div/div[1]/div/div/a/span")).Text;

            while (IsElementPresent(By.XPath("//*[@id='search']/div[1]/div[2]/div/span[8]/div/span/div/div/ul/li["+ pageNumber + "]/a"), driver) == true) //the page numbers at the bottom
            {
                //pageItemCount = 1;

                while (IsElementPresent(By.XPath("//*[@id='search']/div[1]/div[2]/div/span[4]/div[1]/div[" + pageItemCount + "]"), driver) == true) //the actual products
                { 
                    string result = driver.FindElement(By.XPath("//*[@id='search']/div[1]/div[2]/div/span[4]/div[1]/div[" + pageItemCount + "]")).GetAttribute("data-asin");
                    excel.WriteToCell(cellCount, 1, result);

                    cellCount++;
                    pageItemCount++; //maybe use this to tell program to stop. 16 Per full page. If less than 16 it doesn't need to run again. Downside is if the last page is 16 it won't shut off
                }
                pageItemCount = 1;
                pageNumber++; 
                if (pageNumber > 6) //all elements are set to 6 after page 5, so this is necessary to keep the program running till the last page.
                {
                    pageNumber--;
                }


                

                if (IsElementPresent(By.XPath("//*[@id='search']/div[1]/div[2]/div/span[8]/div/span/div/div/ul/li[" + pageNumber + "]/a"), driver) == true)
                {
                    checker = driver.FindElement(By.XPath("//*[@id='search']/div[1]/div[2]/div/span[8]/div/span/div/div/ul/li[" + pageNumber + "]/a")).Text; //supposed to get the actual page number from within the element
                }

                int actualPageNumber = Convert.ToInt32(checker); //the only true count of what page you are on is collected by var checker. 

                if (actualPageNumber == countPageNumber) //Ends the program after it's reached the last page. The only way to know it's scraped all the asins is if it repeats itself
                {
                    pageNumber = 999;
                }
                else
                {
                    countPageNumber = actualPageNumber;
                }


                if (IsElementPresent(By.XPath("//*[@id='search']/div[1]/div[2]/div/span[8]/div/span/div/div/ul/li[" + pageNumber + "]/a"), driver) == true) //clicks on the next page if it exists
                    {
                    var clickPG = driver.FindElement(By.XPath("//*[@id='search']/div[1]/div[2]/div/span[8]/div/span/div/div/ul/li[" + pageNumber + "]/a"));
                    clickPG.Click();
                    }

            }

            //if (IsElementPresent(By.XPath("//*[@id='search']/div[1]/div[2]/div/span[8]/div/span/div/div/ul/li[" + pageNumber + "]/a"), driver) == false) //for sellers with only one page of our products
            //{

               // while (IsElementPresent(By.XPath("//*[@id='search']/div[1]/div[2]/div/span[4]/div[1]/div[" + pageItemCount + "]"), driver) == true)
               // {
                   // string result = driver.FindElement(By.XPath("//*[@id='search']/div[1]/div[2]/div/span[3]/div[1]/div[" + pageItemCount + "]")).GetAttribute("data-asin");
                   // excel.WriteToCell(cellCount, 1, result);

                   // cellCount++;
                    //pageItemCount++;
               //}

            //}

            excel.SaveAs(@"\\brmserver\company\eComm\SellerASINScraper\Reports\ASIN Report for " + sellerID + " - " + sellerName);
            excel.Quit();
            driver.Quit();
        }
    }
}
