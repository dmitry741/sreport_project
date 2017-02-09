using System;

namespace JoinReport
{
    class Program
    {
        static void Main(string[] args)
        {
            SecondVar();
        }

        static void SecondVar()
        {
            string path = @"D:\Dmitry\SReport\StrengthReport\StrengthReport\bin\Debug\reports\new\";
            Stimulsoft.Report.StiReport mainReport = new Stimulsoft.Report.StiReport();
            mainReport.Load(path + "myreport_0_new.mrt");

            // create file collection
            System.Collections.ArrayList reportList = new System.Collections.ArrayList();

            // fill file collection
            for (int i = 1; i <= 2; i++)
            {
                reportList.Add(string.Format("{0}myreport_{1}_new.mrt", path, i));
            }
            
            int index = 0;

            foreach (string file in reportList)
            {
                Stimulsoft.Report.StiReport report = new Stimulsoft.Report.StiReport();
                report.Load(file);

                Stimulsoft.Report.Components.StiPagesCollection pageCol = report.Pages;
                System.Collections.IEnumerator pageEn = pageCol.GetEnumerator();

                while (pageEn.MoveNext())
                {
                    Stimulsoft.Report.Components.StiPage page = pageEn.Current as Stimulsoft.Report.Components.StiPage;
                    Stimulsoft.Report.Components.StiComponentsCollection cc = page.Components;
                    System.Collections.IEnumerator compEn = cc.GetEnumerator();

                    while (compEn.MoveNext())
                    {
                        Stimulsoft.Report.Components.StiComponent component = compEn.Current as Stimulsoft.Report.Components.StiComponent;

                        string Name = component.Name + "_QWER_" + (index++).ToString();
                        component.Name = Name;
                    }

                    Stimulsoft.Report.Components.StiPage newPage = new Stimulsoft.Report.Components.StiPage();
                    newPage.Components.AddRange(cc);
                    //newPage.
                    mainReport.Pages.Add(newPage);

                    Console.WriteLine("Page: " + index.ToString());
                }
            }

            // compile
            mainReport.Compile();

            //mainReport.Render();
            //mainReport.ExportDocument(Stimulsoft.Report.StiExportFormat.Pdf, "result.pdf");

            // save
            mainReport.Save("result.mrt");

            // copy
            System.IO.File.Copy("result.mrt", "D:\\Dmitry\\SReport\\StrengthReport\\StrengthReport\\bin\\Debug\\result.mrt", true);

            // finish
            Console.WriteLine("All ok: to close application hit Enter");
            Console.ReadLine();
        }

        static void FirstVar()
        {
            string path = @"D:\Dmitry\SReport\StrengthReport\StrengthReport\bin\Debug\reports\new";

            // declare main report
            Stimulsoft.Report.StiReport mainReport = new Stimulsoft.Report.StiReport();
            mainReport.Load(path + "myreport_0_new.mrt");

            // get page collection
            Stimulsoft.Report.Components.StiPagesCollection pcMain = mainReport.Pages;

            // create file collection
            System.Collections.ArrayList reportList = new System.Collections.ArrayList();

            // fill file collection
            for (int i = 1; i <= 1; i++)
            {
                reportList.Add(string.Format("{0}myreport_{1}_new.mrt", path, i));
            }

            //reportList.Add("myreport_3.mrt");
            //reportList.Add("myreport_6.mrt");
            //reportList.Add("myreport_8.mrt");
            //reportList.Add("myreport_10.mrt");
            //reportList.Add("myreport_12.mrt");

            //reportList.Add("myreport_14.mrt");
            //reportList.Add("myreport_15.mrt");
            //reportList.Add("myreport_17.mrt");
            //reportList.Add("myreport_20.mrt");
            //reportList.Add("myreport_23.mrt");
            //reportList.Add("myreport_25.mrt");

            //reportList.Add("myreport_27.mrt");
            //reportList.Add("myreport_30.mrt");
            //reportList.Add("myreport_33.mrt");

            int index = 2;

            foreach (string file in reportList)
            {
                Stimulsoft.Report.StiReport report = new Stimulsoft.Report.StiReport();
                report.Load(file);

                Stimulsoft.Report.Components.StiPagesCollection pageCol = report.Pages;

                System.Collections.IEnumerator pageEn = pageCol.GetEnumerator();

                while (pageEn.MoveNext())
                {
                    Stimulsoft.Report.Components.StiPage page = pageEn.Current as Stimulsoft.Report.Components.StiPage;
                    Stimulsoft.Report.Components.StiComponentsCollection cc = page.Components;

                    System.Collections.IEnumerator compEn = cc.GetEnumerator();

                    while (compEn.MoveNext())
                    {
                        Stimulsoft.Report.Components.StiComponent component = compEn.Current as Stimulsoft.Report.Components.StiComponent;

                        string Name = "R" + index.ToString() + component.Name;
                        component.Name = Name;

                        Console.WriteLine(component.Name);
                    }

                    Stimulsoft.Report.Components.StiComponentsCollection ccMain = pcMain[index].Components;
                    ccMain.AddRange(cc);

                    pcMain[index].Components = ccMain;

                    index++;
                    Console.WriteLine("Page: " + index.ToString());
                }
            }

            // compile
            mainReport.Compile();

            //mainReport.Render();
            //mainReport.ExportDocument(Stimulsoft.Report.StiExportFormat.Pdf, "result.pdf");

            // save
            mainReport.Save("result.mrt");

            // copy
            System.IO.File.Copy("result.mrt", "D:\\Dmitry\\SReport\\StrengthReport\\StrengthReport\\bin\\Debug\\result.mrt", true);

            // finish
            Console.WriteLine("All ok: to close application hit Enter");
            Console.ReadLine();
        }
    }
}
