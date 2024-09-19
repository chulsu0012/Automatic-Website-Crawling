using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Diagnostics;
using System.Text.RegularExpressions;
//Selenium Library
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
//EPPlus Library
using OfficeOpenXml;

namespace DataCrawling
{
    public partial class Form1 : Form
    {
        protected ChromeDriverService driverService = null;
        protected ChromeOptions options = null;
        protected ChromeDriver driver = null;

        String url = "https://www.g2b.go.kr/index.jsp";
        String keyword = "RPA";
        String startDate;
        String endDate;

        public Form1()
        {
            InitializeComponent();

            try
            {
                driverService = ChromeDriverService.CreateDefaultService();
                driverService.HideCommandPromptWindow = true;

                options = new ChromeOptions();
                options.AddArgument("disable-gpu");
            }
            catch (Exception ex)
            {
                Trace.WriteLine(ex.Message);
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            dateTimePicker1.Value = DateTime.Today.AddDays(-4);
            dateTimePicker2.Value = DateTime.Today;
        }
        
        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            DateTime date = dateTimePicker1.Value;
            startDate = date.ToString("yyyy'/'MM'/'dd");
        }
        private void dateTimePicker2_ValueChanged(object sender, EventArgs e)
        {
            DateTime date = dateTimePicker2.Value;
            endDate = date.ToString("yyyy'/'MM'/'dd");
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            if (dateTimePicker1.Value > dateTimePicker2.Value)
            {
                MessageBox.Show("시작 날짜가 종료 날짜보다 큽니다. 다시 선택하세요.", "날짜 오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                Search();
                List<List<String>> data = getData();
                if (data.Count > 0) SaveExcel(data);
                else MessageBox.Show("기간 내 데이터가 존재하지 않습니다.");
                Application.Exit();
            }
        }
        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {

        }

        private void Search()
        {
            try
            {
                // Chrome으로 검색 진행
                driver = new ChromeDriver(driverService, options);

                // 링크의 웹 사이트로 이동
                driver.Navigate().GoToUrl(url);
                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);

                // 공고명
                var searchBox = driver.FindElement(By.XPath("//*[@id='bidNm']"));
                searchBox.SendKeys(keyword);

                // 공고일자
                var startInput = driver.FindElement(By.XPath("//*[@id='fromBidDt']"));
                startInput.Clear();
                startInput.SendKeys(startDate);

                var endInput = driver.FindElement(By.XPath("//*[@id='toBidDt']"));
                endInput.Clear();
                endInput.SendKeys(endDate);

                // 검색
                var searchButton = driver.FindElement(By.XPath("//*[@id='searchForm']/div/fieldset[1]/ul/li[4]/dl/dd[3]/a/strong"));
                searchButton.Click();
            }
            catch (Exception exc)
            {
                Trace.WriteLine(exc.Message);
            }
        }

        private List<List<String>> getData()
        {
            List<List<String>> resultData = new List<List<String>>();

            driver.SwitchTo().Frame("sub");
            driver.SwitchTo().Frame("main");

            var table = driver.FindElement(By.XPath("//*[@id='resultForm']/div[2]/table"));
            var tbody = table.FindElement(By.TagName("tbody"));
            var trs = tbody.FindElements(By.TagName("tr"));

            foreach (var tr in trs)
            {
                List<String> data = new List<String>();
                var tds = tr.FindElements(By.TagName("td"));

                if (tds.Count() <= 1)
                {
                    break;
                }

                data.Add(tds[3].Text);
                data.Add(keyword);
                for (int i = 4; i < 6; i++)
                {
                    data.Add(tds[i].Text);
                }
                data.Add(Regex.Replace(tds[7].Text, @"\s*\(.*?\)\s*", String.Empty).Replace('/', '-'));
                data.Add(DateTime.Now.ToString("yyyy-MM-dd HH:mm"));

                resultData.Add(data);
            }

            driver.Quit();
            return resultData;
        }

        private void SaveExcel(List<List<String>> data)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                openFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx";
                openFileDialog.FilterIndex = 1;
                openFileDialog.RestoreDirectory = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    FileInfo fileInfo = new FileInfo(openFileDialog.FileName);

                    using (ExcelPackage package = new ExcelPackage(fileInfo))
                    {
                        ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

                        int startRow = worksheet.Dimension.End.Row + 1;

                        for (int i = 0; i < data.Count; i++)
                        {
                            for (int j = 0; j < data[i].Count; j++)
                            {
                                worksheet.Cells[startRow + i, j + 1].Value = data[i][j];
                            }
                        }

                        package.Save();
                    }
                }
            }
        }
    }
}
