using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace VSTSTestCasesMigration
{
    class Program
    {
        static void Main(string[] args)
        {
            string uri = "https://dev.azure.com/VarinderKumar";
            var project = VSTSOperations.GetTeamProject(uri, "RahulTCMigration");

            //VSTSOperations.ManageTestPlans(uri, project, "rahultpNew");

            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"\\10.26.1.19\Common-Data\Hydra\RahulRao\Foundation-Test Case-20190905\Foundation-Test Case-20190904.xls");
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[2];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;

            Stack<string> areaPathStack = new Stack<string>();
            areaPathStack.Push("RahulTCMigration\\Foundation");
           
            string AttachmentBaseDirectory = @"\\10.26.1.19\Common-Data\Hydra\RahulRao\Foundation-Test Case-20190905\";

            for (int i = 2; i <= rowCount; i++)
            {                
                string suiteName = string.Empty;
                int depth = 0;
                if (xlRange.Cells[i, 1] != null && xlRange.Cells[i, 1].Value2 != null)
                {
                    if (xlRange.Cells[i, 2] == null || xlRange.Cells[i, 2].Value2 == null)
                    {
                        throw new Exception("excel is not formatted correctly.");
                    }
                    suiteName = xlRange.Cells[i, 2].Value2.ToString();
                    suiteName = suiteName.Substring(suiteName.IndexOf(' ') + 1).Trim();

                    depth = Convert.ToInt32(xlRange.Cells[i, 1].Value2.ToString());

                    if (depth == -1)
                        break;

                    while (depth <= areaPathStack.Count - 1)
                    {
                        areaPathStack.Pop();
                    }

                    areaPathStack.Push(areaPathStack.Peek() + "\\" + suiteName);
                }

                Tuple<string, string>[] steps = new Tuple<string, string>[1];
                string title = string.Empty;
                string description = string.Empty;
                string precondition = string.Empty;

                string Product = string.Empty;
                string Component = string.Empty;
                string AutomationClassName = string.Empty;
                string AutomationID = string.Empty;
                string ReasonforNotAutomating = string.Empty;
                string DeploymentMode = string.Empty;
                string PAMTags = string.Empty;
                string Applicable = string.Empty;

                List<Tuple<string, string>> extraFields = new List<Tuple<string, string>>();
                List<string> attachments = new List<string>();

                if (xlRange.Cells[i, 3] != null && xlRange.Cells[i, 3].Value2 != null)
                {
                    if (xlRange.Cells[i, 9] != null && xlRange.Cells[i, 9].Value2 != null)
                    {
                        precondition = "preconditions \r\n" + xlRange.Cells[i, 9].Value2.ToString() + "\r\n\r\n\r\n";
                    }

                    if (xlRange.Cells[i, 11] != null && xlRange.Cells[i, 11].Value2 != null && xlRange.Cells[i, 12] != null && xlRange.Cells[i, 12].Value2 != null)
                    {
                        steps[0] = new Tuple<string, string>(precondition + xlRange.Cells[i, 11].Value2.ToString(), xlRange.Cells[i, 12].Value2.ToString());
                    }
                    else if (xlRange.Cells[i, 11] == null || xlRange.Cells[i, 11].Value2 == null)
                    {
                        steps[0] = new Tuple<string, string>(precondition, xlRange.Cells[i, 12].Value2.ToString());
                    }
                    else if (xlRange.Cells[i, 12] == null || xlRange.Cells[i, 12].Value2 == null)
                    {
                        steps[0] = new Tuple<string, string>(precondition + xlRange.Cells[i, 11].Value2.ToString(), string.Empty);
                    }

                    if (xlRange.Cells[i, 4] != null && xlRange.Cells[i, 4].Value2 != null)
                    {
                        title = xlRange.Cells[i, 4].Value2.ToString();
                    }
                    if (xlRange.Cells[i, 8] != null && xlRange.Cells[i, 8].Value2 != null)
                    {
                        description = xlRange.Cells[i, 8].Value2.ToString();
                    }

                    if (xlRange.Cells[i, 14] != null && xlRange.Cells[i, 14].Value2 != null)
                    {
                        Product = xlRange.Cells[i, 14].Value2.ToString();
                    }
                    if (xlRange.Cells[i, 15] != null && xlRange.Cells[i, 15].Value2 != null)
                    {
                        Component = xlRange.Cells[i, 15].Value2.ToString();
                    }
                    //if (xlRange.Cells[i, 16] != null && xlRange.Cells[i, 16].Value2 != null)
                    //{
                    //    AutomationStatus = xlRange.Cells[i, 16].Value2.ToString();
                    //}
                    if (xlRange.Cells[i, 17] != null && xlRange.Cells[i, 17].Value2 != null)
                    {
                        AutomationClassName = xlRange.Cells[i, 17].Value2.ToString();
                    }
                    if (xlRange.Cells[i, 18] != null && xlRange.Cells[i, 18].Value2 != null)
                    {
                        AutomationID = xlRange.Cells[i, 18].Value2.ToString();
                    }
                    if (xlRange.Cells[i, 19] != null && xlRange.Cells[i, 19].Value2 != null)
                    {
                        ReasonforNotAutomating = xlRange.Cells[i, 19].Value2.ToString();
                    }
                    if (xlRange.Cells[i, 20] != null && xlRange.Cells[i, 20].Value2 != null)
                    {
                        DeploymentMode = xlRange.Cells[i, 20].Value2.ToString();
                    }
                    if (xlRange.Cells[i, 21] != null && xlRange.Cells[i, 21].Value2 != null)
                    {
                        PAMTags = xlRange.Cells[i, 21].Value2.ToString();
                    }
                    if (xlRange.Cells[i, 22] != null && xlRange.Cells[i, 22].Value2 != null)
                    {
                        Applicable = xlRange.Cells[i, 22].Value2.ToString();
                    }
                    extraFields.Add(new Tuple<string, string>("Product", Product));
                    extraFields.Add(new Tuple<string, string>("Component", Component));
                    extraFields.Add(new Tuple<string, string>("AutomationClassName", AutomationClassName));
                    extraFields.Add(new Tuple<string, string>("Automation ID", AutomationID));
                    extraFields.Add(new Tuple<string, string>("Reason for Not Automating", ReasonforNotAutomating));
                    extraFields.Add(new Tuple<string, string>("Deployment Mode", DeploymentMode));
                    extraFields.Add(new Tuple<string, string>("PAM Tags", PAMTags));
                    extraFields.Add(new Tuple<string, string>("Applicable", Applicable));

                    if(xlRange.Cells[i, 5] != null && xlRange.Cells[i, 5].Value2 != null)
                    {
                        do
                        {
                            string attachmentUrl = ((Excel.Range)xlRange.Cells[i, 5]).Cells.Hyperlinks[1].Address;
                            attachments.Add(AttachmentBaseDirectory + attachmentUrl);
                            i++;
                        } while ((xlRange.Cells[i, 3] == null || xlRange.Cells[i, 3].Value2 == null) && (xlRange.Cells[i, 5] != null && xlRange.Cells[i, 5].Value2 != null));
                        i--;
                    }

                    VSTSOperations.CreateAreaPath(areaPathStack.Peek(), uri, project.Guid);
                    VSTSOperations.CreateNewTestCase(uri, project, title, areaPathStack.Peek(), "RahulTCMigration", description, "Rahul Rao", steps, extraFields,attachments);

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
