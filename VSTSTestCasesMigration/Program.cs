using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Configuration;

namespace VSTSTestCasesMigration
{
    class Program
    {
        static void Main(string[] args)
        {
            string uri = ConfigurationManager.AppSettings["Uri"];
            string projectName = ConfigurationManager.AppSettings["ProjectName"];
            string productName = ConfigurationManager.AppSettings["ProductName"];
            string assignee = ConfigurationManager.AppSettings["Assignee"];
            string PAT = ConfigurationManager.AppSettings["PAT"];
            string baseDirectory = ConfigurationManager.AppSettings["BaseDirectory"];
            string tcFileName = ConfigurationManager.AppSettings["TestCasesFile"];
            bool createAreaPaths = Convert.ToBoolean(ConfigurationManager.AppSettings["CreateAreaPaths"]);
            var project = VSTSOperations.GetTeamProject(uri, projectName);

            int logFileStartIndex = baseDirectory.LastIndexOf('\\', baseDirectory.Length - 2) + 1;
            MyLogger.FileName = AppDomain.CurrentDomain.BaseDirectory + baseDirectory.Substring(logFileStartIndex, baseDirectory.Length - logFileStartIndex - 1) + DateTime.Now.ToString("yyMMddHHmmss") + ".txt";
            //MyLogger.Log("abbcddd");
            //VSTSOperations.ManageTestPlans(uri, project, ConfigurationManager.AppSettings["TestPlanName"]);
            List<int> ids = new List<int>();
            for (int i = 24201; i <= 24209; i++)
            {
                ids.Add(i);
            }
            VSTSOperations.DeleteWorkItemsOneByOne(uri, ids);
            VSTSOperations.DeleteWorkItems(uri, ids);

            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(baseDirectory + tcFileName);
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[2];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;

            Stack<string> areaPathStack = new Stack<string>();
            areaPathStack.Push(projectName + "\\" + productName);

            if (createAreaPaths)
            {
                for (int i = 2; i <= rowCount; i++)
                {
                    try
                    {
                        string suiteName = string.Empty;
                        int depth = 0;
                        if (xlRange.Cells[i, 1] != null && xlRange.Cells[i, 1].Value2 != null)
                        {
                            depth = Convert.ToInt32(xlRange.Cells[i, 1].Value2.ToString());

                            if (depth == -1)
                                break;

                            if (xlRange.Cells[i, 2] == null || xlRange.Cells[i, 2].Value2 == null)
                            {
                                throw new Exception("excel is not formatted correctly.");
                            }
                            suiteName = xlRange.Cells[i, 2].Value2.ToString();
                            suiteName = suiteName.Substring(suiteName.IndexOf(' ') + 1).Trim();

                            suiteName = Helpers.ReplaceReservedChars(suiteName);

                            while (depth <= areaPathStack.Count - 1)
                            {
                                areaPathStack.Pop();
                            }

                            areaPathStack.Push(areaPathStack.Peek() + "\\" + suiteName);
                            VSTSOperations.CreateAreaPath(areaPathStack.Peek(), uri, project.Guid, PAT);
                            project = VSTSOperations.GetTeamProject(uri, projectName);
                        }
                    }
                    catch (Exception ex)
                    {
                        MyLogger.Log("Error while creating area path : " + areaPathStack.Peek() + " on row : " + i.ToString());
                        MyLogger.Log(ex.Message);
                    }
                    Console.WriteLine("Area Path Processed Row : " + i.ToString() + " of " + rowCount.ToString());
                }
            }

            while (areaPathStack.Count != 0)
                areaPathStack.Pop();

            areaPathStack.Push(projectName + "\\" + productName);

            for (int i = 2; i <= rowCount; i++)
            {
                try
                {
                    string suiteName = string.Empty;
                    int depth = 0;
                    if (xlRange.Cells[i, 1] != null && xlRange.Cells[i, 1].Value2 != null)
                    {
                        depth = Convert.ToInt32(xlRange.Cells[i, 1].Value2.ToString());

                        if (depth == -1)
                            break;

                        if (xlRange.Cells[i, 2] == null || xlRange.Cells[i, 2].Value2 == null)
                        {
                            throw new Exception("excel is not formatted correctly.");
                        }
                        suiteName = xlRange.Cells[i, 2].Value2.ToString();
                        suiteName = suiteName.Substring(suiteName.IndexOf(' ') + 1).Trim();

                        suiteName = Helpers.ReplaceReservedChars(suiteName);

                        while (depth <= areaPathStack.Count - 1)
                        {
                            areaPathStack.Pop();
                        }

                        areaPathStack.Push(areaPathStack.Peek() + "\\" + suiteName);
                        //VSTSOperations.CreateAreaPath(areaPathStack.Peek(), uri, project.Guid, PAT);
                        //project = VSTSOperations.GetTeamProject(uri, projectName);
                    }

                    Tuple<string, string>[] steps = new Tuple<string, string>[1];
                    string title = string.Empty;
                    string description = string.Empty;
                    string precondition = string.Empty;

                    string Product = string.Empty;
                    string Component = string.Empty;
                    string SubComponent = string.Empty;
                    string Priority = string.Empty;
                    string Release = string.Empty;
                    string AutomationStatus = string.Empty;
                    string AutomationClassName = string.Empty;
                    string AutomationID = string.Empty;
                    string ReasonforNotAutomating = string.Empty;
                    string TestCategory = string.Empty;
                    string DeploymentMode = string.Empty;
                    string PAMTags = string.Empty;
                    string Test_Data = string.Empty;
                    string SolutionName = string.Empty;
                    string Applicable = string.Empty;
                    string EpicIds = string.Empty;

                    List<Tuple<string, string>> extraFields = new List<Tuple<string, string>>();
                    List<string> attachments = new List<string>();

                    if (xlRange.Cells[i, 3] != null && xlRange.Cells[i, 3].Value2 != null)
                    {
                        if (xlRange.Cells[i, 9] != null && xlRange.Cells[i, 9].Value2 != null)
                        {
                            precondition = "preconditions \r\n" + xlRange.Cells[i, 9].Value2.ToString() + "\r\n\r\n\r\n";
                        }
                        string allsteps = precondition;
                        string allexpectedresults = string.Empty;
                        int iBeforeSteps = i;
                        int iAfterSteps = i;
                        if (xlRange.Cells[i, 11] != null && xlRange.Cells[i, 11].Value2 != null)
                        {
                            do
                            {
                                allsteps += xlRange.Cells[i, 11].Value2.ToString();
                                allexpectedresults += xlRange.Cells[i, 12].Value2.ToString();
                                i++;
                            } while ((xlRange.Cells[i, 3] == null || xlRange.Cells[i, 3].Value2 == null) && (xlRange.Cells[i, 11] != null && xlRange.Cells[i, 11].Value2 != null));
                            iAfterSteps = i - 1;
                            i = iBeforeSteps;
                        }

                        steps[0] = new Tuple<string, string>(allsteps, allexpectedresults);
                        //if (xlRange.Cells[i, 11] != null && xlRange.Cells[i, 11].Value2 != null && xlRange.Cells[i, 12] != null && xlRange.Cells[i, 12].Value2 != null)
                        //{
                        //    steps[0] = new Tuple<string, string>(precondition + xlRange.Cells[i, 11].Value2.ToString(), xlRange.Cells[i, 12].Value2.ToString());
                        //}
                        //else if (xlRange.Cells[i, 11] == null || xlRange.Cells[i, 11].Value2 == null)
                        //{
                        //    steps[0] = new Tuple<string, string>(precondition, xlRange.Cells[i, 12].Value2.ToString());
                        //}
                        //else if (xlRange.Cells[i, 12] == null || xlRange.Cells[i, 12].Value2 == null)
                        //{
                        //    steps[0] = new Tuple<string, string>(precondition + xlRange.Cells[i, 11].Value2.ToString(), string.Empty);
                        //}

                        if (xlRange.Cells[i, 4] != null && xlRange.Cells[i, 4].Value2 != null)
                        {
                            title = xlRange.Cells[i, 4].Value2.ToString();
                            title = Helpers.ReplaceReservedChars(title);
                            title = (title.Length > 255) ? title.Substring(0, 255) : title;
                        }
                        if (xlRange.Cells[i, 8] != null && xlRange.Cells[i, 8].Value2 != null)
                        {
                            description = xlRange.Cells[i, 8].Value2.ToString();
                            description = Helpers.ReplaceReservedChars(description);
                        }

                        if (xlRange.Cells[i, 14] != null && xlRange.Cells[i, 14].Value2 != null)
                        {
                            Product = xlRange.Cells[i, 14].Value2.ToString();
                            Product = Helpers.ReplaceReservedChars(Product);
                        }
                        if (xlRange.Cells[i, 15] != null && xlRange.Cells[i, 15].Value2 != null)
                        {
                            Component = xlRange.Cells[i, 15].Value2.ToString();
                            Component = Helpers.ReplaceReservedChars(Component);
                        }
                        if (xlRange.Cells[i, 16] != null && xlRange.Cells[i, 16].Value2 != null)
                        {
                            SubComponent = xlRange.Cells[i, 16].Value2.ToString();
                            SubComponent = Helpers.ReplaceReservedChars(SubComponent);
                        }
                        if (xlRange.Cells[i, 17] != null && xlRange.Cells[i, 17].Value2 != null)
                        {
                            Priority = xlRange.Cells[i, 17].Value2.ToString();
                            Priority = Helpers.ReplaceReservedChars(Priority);
                        }
                        if (xlRange.Cells[i, 18] != null && xlRange.Cells[i, 18].Value2 != null)
                        {
                            Release = xlRange.Cells[i, 18].Value2.ToString();
                            Release = Helpers.ReplaceReservedChars(Release);
                        }
                        if (xlRange.Cells[i, 19] != null && xlRange.Cells[i, 19].Value2 != null)
                        {
                            AutomationStatus = xlRange.Cells[i, 19].Value2.ToString();
                            AutomationStatus = Helpers.ReplaceReservedChars(AutomationStatus);
                        }
                        if (xlRange.Cells[i, 20] != null && xlRange.Cells[i, 20].Value2 != null)
                        {
                            AutomationClassName = xlRange.Cells[i, 20].Value2.ToString();
                            AutomationClassName = Helpers.ReplaceReservedChars(AutomationClassName);
                        }
                        if (xlRange.Cells[i, 21] != null && xlRange.Cells[i, 21].Value2 != null)
                        {
                            AutomationID = xlRange.Cells[i, 21].Value2.ToString();
                            AutomationID = Helpers.ReplaceReservedChars(AutomationID);
                        }
                        if (xlRange.Cells[i, 22] != null && xlRange.Cells[i, 22].Value2 != null)
                        {
                            ReasonforNotAutomating = xlRange.Cells[i, 22].Value2.ToString();
                            ReasonforNotAutomating = Helpers.ReplaceReservedChars(ReasonforNotAutomating);
                        }
                        if (xlRange.Cells[i, 23] != null && xlRange.Cells[i, 23].Value2 != null)
                        {
                            TestCategory = xlRange.Cells[i, 23].Value2.ToString();
                            TestCategory = Helpers.ReplaceReservedChars(TestCategory);
                        }
                        if (xlRange.Cells[i, 24] != null && xlRange.Cells[i, 24].Value2 != null)
                        {
                            DeploymentMode = xlRange.Cells[i, 24].Value2.ToString();
                            DeploymentMode = Helpers.ReplaceReservedChars(DeploymentMode);
                        }
                        if (xlRange.Cells[i, 25] != null && xlRange.Cells[i, 25].Value2 != null)
                        {
                            PAMTags = xlRange.Cells[i, 25].Value2.ToString();
                            PAMTags = Helpers.ReplaceReservedChars(PAMTags);
                        }
                        if (xlRange.Cells[i, 26] != null && xlRange.Cells[i, 26].Value2 != null)
                        {
                            Test_Data = xlRange.Cells[i, 26].Value2.ToString();
                            Test_Data = Helpers.ReplaceReservedChars(Test_Data);
                        }
                        if (xlRange.Cells[i, 27] != null && xlRange.Cells[i, 27].Value2 != null)
                        {
                            SolutionName = xlRange.Cells[i, 27].Value2.ToString();
                            SolutionName = Helpers.ReplaceReservedChars(SolutionName);
                        }
                        if (xlRange.Cells[i, 28] != null && xlRange.Cells[i, 28].Value2 != null)
                        {
                            Applicable = xlRange.Cells[i, 28].Value2.ToString();
                            Applicable = Helpers.ReplaceReservedChars(Applicable);
                        }
                        if (xlRange.Cells[i, 29] != null && xlRange.Cells[i, 29].Value2 != null)
                        {
                            EpicIds = xlRange.Cells[i, 29].Value2.ToString();
                            EpicIds = Helpers.ReplaceReservedChars(EpicIds);
                        }

                        extraFields.Add(new Tuple<string, string>("Product", Product));
                        extraFields.Add(new Tuple<string, string>("Component", Component));
                        extraFields.Add(new Tuple<string, string>("Sub Component", SubComponent));
                        extraFields.Add(new Tuple<string, string>("qtestPriority", Priority));
                        extraFields.Add(new Tuple<string, string>("Release", Release));
                        extraFields.Add(new Tuple<string, string>("qTestAutomation Status", AutomationStatus));
                        extraFields.Add(new Tuple<string, string>("AutomationClassName", AutomationClassName));
                        extraFields.Add(new Tuple<string, string>("Automation ID", AutomationID));
                        extraFields.Add(new Tuple<string, string>("Reason for Not Automating", ReasonforNotAutomating));
                        extraFields.Add(new Tuple<string, string>("Test Category", TestCategory));
                        extraFields.Add(new Tuple<string, string>("Deployment Mode", DeploymentMode));
                        extraFields.Add(new Tuple<string, string>("PAM Tags", PAMTags));
                        extraFields.Add(new Tuple<string, string>("Test_Data", Test_Data));
                        extraFields.Add(new Tuple<string, string>("Solution Name", SolutionName));
                        extraFields.Add(new Tuple<string, string>("Applicable", Applicable));
                        extraFields.Add(new Tuple<string, string>("EPIC_IDs", EpicIds));

                        if (xlRange.Cells[i, 5] != null && xlRange.Cells[i, 5].Value2 != null)
                        {
                            do
                            {
                                string attachmentUrl = ((Excel.Range)xlRange.Cells[i, 5]).Cells.Hyperlinks[1].Address;
                                attachments.Add(baseDirectory + attachmentUrl);
                                i++;
                            } while ((xlRange.Cells[i, 3] == null || xlRange.Cells[i, 3].Value2 == null) && (xlRange.Cells[i, 5] != null && xlRange.Cells[i, 5].Value2 != null));
                            i--;
                            i = (i > iAfterSteps) ? i : iAfterSteps;
                        }

                        VSTSOperations.CreateNewTestCase(uri, project, title, areaPathStack.Peek(), projectName, description, assignee, steps, extraFields, attachments);

                    }
                    Console.WriteLine("Processed Row : " + i.ToString() + " of " + rowCount.ToString());
                }
                catch (Exception ex)
                {
                    MyLogger.Log("Error while processing row : " + i.ToString());
                    MyLogger.Log(ex.Message);
                }
            }

            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);

            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
        }
    }
}
