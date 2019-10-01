using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.TestManagement.Client;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using Microsoft.TeamFoundation.WorkItemTracking.Internals;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi;
using Microsoft.VisualStudio.Services.Common;
using Microsoft.VisualStudio.Services.WebApi;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace VSTSTestCasesMigration
{
    public class VSTSOperations
    {
        public class ActionResult
        {
            public bool Success { get; set; }
            public List<string> ErrorCodes { get; set; }
            public int Id { get; set; }
        }
        private static ActionResult CheckValidationResult(WorkItem workItem)
        {
            var validationResult = workItem.Validate();

            ActionResult result = null;
            if (validationResult.Count == 0)
            {
                // Save the new work item.
                workItem.Save();

                result = new ActionResult()
                {
                    Success = true,
                    Id = workItem.Id
                };
            }
            else
            {
                //result = ParseValidation(validationResult);
                throw new Exception("Invalid Fields");
            }

            return result;
        }
        private static WorkItem GetWorkItem(string uri, int testedWorkItemId)
        {
            TfsTeamProjectCollection tfs;

            tfs = TfsTeamProjectCollectionFactory.GetTeamProjectCollection(new Uri(uri)); // https://mytfs.visualstudio.com/DefaultCollection
            tfs.Authenticate();

            var workItemStore = new WorkItemStore(tfs);
            WorkItem workItem = workItemStore.GetWorkItem(testedWorkItemId);

            return workItem;

        }
        public static Project GetTeamProject(string uri, string name)
        {
            TfsTeamProjectCollection tfs;

            tfs = TfsTeamProjectCollectionFactory.GetTeamProjectCollection(new Uri(uri)); // https://mytfs.visualstudio.com/DefaultCollection
            tfs.Authenticate();

            var workItemStore = new WorkItemStore(tfs);

            var project = (from Project pr in workItemStore.Projects
                           where pr.Name == name
                           select pr).FirstOrDefault();
            if (project == null)
                throw new Exception($"Unable to find {name} in {uri}");

            return project;
        }
        public static void DeleteWorkItemsOneByOne(string uri, List<int> ids)
        {
            try
            {
                TfsTeamProjectCollection tfs;

                tfs = TfsTeamProjectCollectionFactory.GetTeamProjectCollection(new Uri(uri)); // https://mytfs.visualstudio.com/DefaultCollection
                tfs.Authenticate();

                var workItemStore = new WorkItemStore(tfs);

                foreach (var id in ids)
                {
                    List<int> idList = new List<int>();
                    idList.Add(id);
                    workItemStore.DestroyWorkItems(idList);
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine("Error while deleting workitems one by one " + ex.Message);
            }
        }
        public static void DeleteWorkItems(string uri, List<int> ids)
        {
            try
            {
                TfsTeamProjectCollection tfs;

                tfs = TfsTeamProjectCollectionFactory.GetTeamProjectCollection(new Uri(uri)); // https://mytfs.visualstudio.com/DefaultCollection
                tfs.Authenticate();

                var workItemStore = new WorkItemStore(tfs);

                workItemStore.DestroyWorkItems(ids);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error while deleting workitems " + ex.Message);
            }
        }
        public static void ManageTestPlans(string uri, Project project, string testPlanName)
        {
            try
            {
                TfsTeamProjectCollection tfs;

                tfs = TfsTeamProjectCollectionFactory.GetTeamProjectCollection(new Uri(uri)); // https://mytfs.visualstudio.com/DefaultCollection
                tfs.Authenticate();

                ITestManagementService service = (ITestManagementService)tfs.GetService(typeof(ITestManagementService));
                ITestManagementTeamProject testProject = service.GetTeamProject(project);

                //Create TestPlan
                ITestPlan newTestPlan = testProject.TestPlans.Create();
                newTestPlan.Name = testPlanName;
                newTestPlan.Owner = tfs.AuthorizedIdentity;
                newTestPlan.Save();

                //Get all are paths
                NodeCollection nodeCollection = testProject.WitProject.AreaRootNodes;

                foreach (Node node in nodeCollection)
                {
                    // Create Static Test Suite and set name based on area path
                    IStaticTestSuite stsuite = testProject.TestSuites.CreateStatic();
                    stsuite.Title = node.Name;

                    newTestPlan.RootSuite.Entries.Add(stsuite);
                    newTestPlan.Save();

                    if (node.HasChildNodes)
                    {
                        CreateTestSuites(testProject, newTestPlan, node, stsuite, testProject.TeamProjectName + "\\" + node.Name);
                    }
                }
            }
            catch (Exception ex)
            {

            }

        }

        private static void CreateTestSuites(ITestManagementTeamProject testProject, ITestPlan newTestPlan, Node node, IStaticTestSuite parentSuite, string parentAreaPath)
        {
            NodeCollection childNodes = node.ChildNodes;

            foreach (Node childNode in childNodes)
            {
                IStaticTestSuite newsuite = null;
                newsuite = (IStaticTestSuite)testProject.TestSuites.Find(68227);

                newsuite = testProject.TestSuites.CreateStatic();
                newsuite.Title = childNode.Name;
                parentSuite.Entries.Add(newsuite);
                newTestPlan.Save();

                try
                {
                    //find tc's based on areapath
                    IEnumerable<ITestCase> tcs = testProject.TestCases.Query(string.Format("SELECT [System.Id], [System.Title] FROM WorkItems WHERE [System.TeamProject]='{0}' AND [System.AreaPath]='{1}' AND [System.WorkItemType]='Test Case'", testProject.TeamProjectName, parentAreaPath + "\\" + childNode.Name));
                    //Add the above entries to the static test suite.
                    foreach (ITestCase tc in tcs)
                    {
                        newsuite.Entries.Add(tc);
                    }
                }
                catch
                {
                    MyLogger.Log("Error while adding tcs for node : " + childNode.Name + " whose parent area path is : " + parentAreaPath);
                }

                newTestPlan.Save();

                if (childNode.HasChildNodes)
                {
                    CreateTestSuites(testProject, newTestPlan, childNode, newsuite, parentAreaPath + "\\" + childNode.Name);
                }
            }
        }

        private static void AddTestCaseSteps(string uri, Project project, int testCaseId, Tuple<string, string>[] steps)
        {
            TfsTeamProjectCollection tfs;

            tfs = TfsTeamProjectCollectionFactory.GetTeamProjectCollection(new Uri(uri)); // https://mytfs.visualstudio.com/DefaultCollection
            tfs.Authenticate();

            ITestManagementService service = (ITestManagementService)tfs.GetService(typeof(ITestManagementService));
            ITestManagementTeamProject testProject = service.GetTeamProject(project);
            ITestCase testCase = testProject.TestCases.Find(testCaseId);

            foreach (var step in steps)
            {
                ITestStep newStep = testCase.CreateTestStep();
                newStep.Title = Helpers.ReplaceReservedChars(step.Item1);
                newStep.ExpectedResult = step.Item2;

                testCase.Actions.Add(newStep);
            }
            testCase.Save();
        }

        public static void CreateAreaPath(string areaPath, string uri, Guid projectGUID, string PAT)
        {
            try
            {
                var node = new Microsoft.TeamFoundation.WorkItemTracking.WebApi.Models.WorkItemClassificationNode();
                node.StructureType = Microsoft.TeamFoundation.WorkItemTracking.WebApi.Models.TreeNodeStructureType.Area;
                string areaName = areaPath.Substring(areaPath.LastIndexOf('\\') + 1);
                string pathToParentNode = areaPath.Substring(areaPath.IndexOf('\\') + 1, areaPath.LastIndexOf('\\') - areaPath.IndexOf('\\'));
                node.Name = areaName;

                VssConnection connection = new VssConnection(new Uri(uri), new VssBasicCredential(string.Empty, PAT));
                var client = connection.GetClient<WorkItemTrackingHttpClient>();

                var result = client.CreateOrUpdateClassificationNodeAsync(node, projectGUID, Microsoft.TeamFoundation.WorkItemTracking.WebApi.Models.TreeStructureGroup.Areas, pathToParentNode).Result;
            }
            catch (AggregateException ex)
            {
            }
            catch (Exception ex)
            {
                MyLogger.Log("Error in AreaPath Creation : " + areaPath);
                throw;
            }
        }
        public static ActionResult CreateNewTestCase(string uri, Project project, string title, string areaPath, string iterationPath, string description, string assignee, Tuple<string, string>[] reproductionSteps, List<Tuple<string, string>> extraFields, List<string> attachments, object tags = null)
        {
            WorkItemType workItemType = project.WorkItemTypes["Test Case"];

            // Create the work item. 
            WorkItem newTestCase = new WorkItem(workItemType);
            newTestCase.Title = string.IsNullOrEmpty(title) ? "Untitled" : title;
            newTestCase.Description = string.IsNullOrEmpty(description) ? "no desc" : description;
            newTestCase.AreaPath = areaPath;
            newTestCase.IterationPath = iterationPath;
            newTestCase.Fields["Assigned To"].Value = assignee;

            foreach (var attachment in attachments)
            {
                newTestCase.Attachments.Add(new Attachment(attachment, "added by rahul rao"));
            }

            foreach (var extrafield in extraFields)
            {
                newTestCase.Fields[extrafield.Item1].Value = extrafield.Item2;
            }

            // Copy tags
            if (tags != null)
                newTestCase.Fields["Tags"].Value = tags;

            ActionResult result = CheckValidationResult(newTestCase);
            if (result.Success)
            {
                AddTestCaseSteps(uri, project, result.Id, reproductionSteps);
            }

            return result;
        }
    }
}
