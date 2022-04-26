using System;
using OpenQA.Selenium;
using OpenQA.Selenium.Winium;
using OpenQA.Selenium.Chrome;
using System.Security.Policy;
using System.Threading;
using System.IO;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace Automation_FiltRU
{
    class Program
    {            
        static void Main(string[] args)
        {
        //Console.WriteLine("Hello World!");
        ////Acessar OTRs 
            OverHere:
            IWebDriver driver = new ChromeDriver();            
            driver.Navigate().GoToUrl("https://servicedesk.bemol.com.br/otrs/index.pl");
            IWebElement campoUser = driver.FindElement(By.Id("User"));
            IWebElement campoSenha = driver.FindElement(By.Id("Password"));
            IWebElement btnLogin = driver.FindElement(By.Id("LoginButton"));
                       

            // FAzer login
            campoUser.SendKeys("16683");
            campoSenha.SendKeys("jW26uReb");
            btnLogin.Click();

            driver.Navigate().GoToUrl("https://servicedesk.bemol.com.br/otrs/index.pl?Action=AgentStatistics;Subaction=Overview");
            GetOver:
            driver.Navigate().GoToUrl("https://servicedesk.bemol.com.br/otrs/index.pl?Action=AgentStatistics;Subaction=View;StatID=81");
            //Baixar Arquivo Excel
            try
            {
                IWebElement btnFormato = driver.FindElement(By.Id("Format_Search"));
                btnFormato.SendKeys("Excel");
                IWebElement btnExcel = driver.FindElement(By.CssSelector("div.InputField_TreeContainer"));
                btnExcel.Click();
                Thread.Sleep(100);
                IWebElement btnStart = driver.FindElement(By.Id("StartStatistic"));
                btnStart.Click();
                //Captura a data logo quando ocorre o Download do Arquivo
                
                
            }
            catch{
                driver.Quit();
                goto OverHere;
            }
            DateTime today = DateTime.Now;
            String date = today.ToString("yyyy-MM-dd_HH_mm");
             
            Thread.Sleep(40000);

            
            string path = $@"C:\Users\dashboard.di\Downloads\Chamados_Infraestrutura_Created_{date}_TimeZone_America_Manaus.xlsx";

            
            //verifia se existe a planilha e apaga caso exista.
            if (File.Exists(@"C:\Users\dashboard.di\Downloads\bruto.xlsx"))
            {
                System.IO.File.Delete(@"C:\Users\dashboard.di\Downloads\bruto.xlsx");
            }
            //if (File.Exists(@"C:\Users\15992\Downloads\testecs\planilhabase.xlsx"))
            //{
            //    System.IO.File.Delete(@"C:\Users\15992\Downloads\testecs\planilhabase.xlsx");
            //}

            Thread.Sleep(1000);

            //Renomeia o arquivo para bruto.
            bool result = File.Exists(path);
            if (result)
            {
                string oldName = path;
                string newName = @"C:\Users\dashboard.di\Downloads\bruto.xlsx";
                System.IO.File.Move(oldName, newName);
            }
            else
            {
                goto GetOver;
            }

            //Executa o script Python via CMD
            //Nos testes demorou 51 segs para ser efetuada a conversão via python.
            string command = @"/C C:\Users\dashboard.di\Documents\filtRU.exe";
            Process.Start("cmd.exe", command);

            //Fecha o navegardor e os processos abertos com ele
            

            Thread.Sleep(60000);
            driver.Quit();
        //Environment.Exit(0);            

        //Abre o Power BI
        
         String addPath = @"C:\Users\dashboard.di\Documents\BI\BI\DashBoard Infraestrutura.pbix";

            DesktopOptions dop = new DesktopOptions();
            dop.ApplicationPath = addPath;

            String url = @"C:\Users\dashboard.di\Documents\Winium";
            IWebDriver openDriver = new WiniumDriver(url, dop);

            Mouse:
            //SetCursorPosition(852,135);

            Thread.Sleep(5000);
            try
            {
                openDriver.FindElement(By.Name("Atualizar")).Click();
            }
            catch
            {
                goto Mouse;
            }

            Thread.Sleep(10000);
            openDriver.Close();
            Environment.Exit(0);
        }
        //Classe para setar o mouse em alguma posição
        [DllImport("user32.dll")]
        static extern bool SetCursorPos(int X, int Y);

        public static void SetCursorPosition(int x, int y)
        {
            SetCursorPos(x, y);
        }
        


    }
}
