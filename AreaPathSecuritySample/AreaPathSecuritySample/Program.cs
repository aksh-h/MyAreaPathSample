using CommandLine;
using Microsoft.TeamFoundation.Core.WebApi;
using Microsoft.TeamFoundation.SourceControl.WebApi;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi.Models;
using Microsoft.VisualStudio.Services.Client;
using Microsoft.VisualStudio.Services.Graph.Client;
using Microsoft.VisualStudio.Services.Identity;
using Microsoft.VisualStudio.Services.Identity.Client;
using Microsoft.VisualStudio.Services.Security;
using Microsoft.VisualStudio.Services.Security.Client;
using Microsoft.VisualStudio.Services.WebApi;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace AddUserToAreaPath
{
    internal class Program
    {
        private static Guid securityNamespaceId = new Guid("83e28ad4-2d72-4ceb-97b0-c7726d5502c3");
        private static Guid projectSecurityNamespaceId = new Guid("52d39943-cb85-4d7f-8fa8-c6baac873819");
        private static Guid gitRepoNamespaceId = new Guid("2e9eb7ed-3c0a-47d4-87c1-0ffdd275fd87");
        private static void Main(string[] args)
        {
            try
            {
                var result = Parser.Default.ParseArguments<Options>(args);
                //Console.WriteLine("Enter Organization Name");
                string accountUrl = "culater";
                //string accountUrl = Console.ReadLine();
                accountUrl = "https://dev.azure.com/" + accountUrl;
                //Console.WriteLine("Enter Project Name");

                string projectName = "ContosoAir";
                //string projectName = Console.ReadLine();
                string areaPathName = string.Empty;
                string groupName = string.Empty;
                string projectId = string.Empty;
                Console.WriteLine("You might see a login screen if you have never signed in to your account using this app.");
                VssConnection connection = new VssConnection(new Uri(accountUrl), new VssClientCredentials());

                // Get the team project
                TeamProject project = GetProject(connection, projectName);
                // Create Group at project level
                Dictionary<string, string> groupArea = new Dictionary<string, string>();
                string[] staticGroups = new string[] { "Dev", "QA", "Dev Lead", "Tech Lead", "Manager" };
                string _area = string.Empty;
                string _subArea = string.Empty;
                string _group = string.Empty;

                WorkItemTrackingHttpClient workClient = connection.GetClient<WorkItemTrackingHttpClient>();
                WorkItemClassificationNode areaPaths = workClient.GetClassificationNodeAsync(project.Id, TreeStructureGroup.Areas, depth: 5).Result;

                GitHttpClient gitClient = connection.GetClient<GitHttpClient>();
                List<GitRepository> repos = gitClient.GetRepositoriesAsync(project.Id).Result;

                if (areaPaths.HasChildren == true)
                {
                    List<string> listAreas = new List<string>();
                    string areaName = string.Empty;
                    //create areas and add it to list
                    foreach (var child in areaPaths.Children)
                    {
                        areaName = child.Name;
                        if (child.HasChildren == false)
                        {
                            listAreas.Add(areaName);
                        }

                        if (child.HasChildren == true)
                        {
                            foreach (var ch in child.Children)
                            {
                                string subArea = ch.Name;
                                string newSubAreaName = areaName + "\\" + subArea;
                                listAreas.Add(newSubAreaName);
                            }
                        }
                    }
                    foreach (var lArea in listAreas)
                    {
                        string subAreaName = string.Empty;
                        string[] splitArea = lArea.Split('\\');
                        string newgroupName = string.Empty;
                        if (splitArea.Length == 2)
                        {
                            subAreaName = splitArea[1];
                        }
                        else
                        {
                            subAreaName = splitArea[0];
                        }
                        foreach (var _sGroups in staticGroups)
                        {
                            newgroupName = subAreaName + "_" + _sGroups;
                            groupArea.Add(newgroupName, lArea);
                        }
                    }

                }

                if (groupArea.Count > 0)
                {
                    foreach (var grp in groupArea)
                    {
                        CreateProjectVSTSGroup(connection, project.Id, grp.Key);
                    }
                }
                Console.WriteLine("Mapping groups to area, please wait..");

                if (groupArea.Count > 0)
                {
                    foreach (var grp in groupArea)
                    {
                        areaPathName = grp.Value;
                        groupName = grp.Key;
                        WorkItemClassificationNode areaPath = workClient.GetClassificationNodeAsync(project.Id, TreeStructureGroup.Areas, path: areaPathName).Result;
                        // Get the group
                        Identity group = GetProjectGroup(connection, groupName, projectName);
                        string areaUnderGroupsToMove = string.Empty;
                        string[] _areaName = areaPathName.Split('\\');
                        if (_areaName.Length == 2)
                        {
                            areaUnderGroupsToMove = _areaName[1];
                        }
                        else
                        {
                            areaUnderGroupsToMove = _areaName[0];
                        }
                        var repo_GrouopToBeMovedUnder = repos.Where(r => r.Name == areaUnderGroupsToMove).SingleOrDefault();
                        if (repo_GrouopToBeMovedUnder != null)
                        {
                            Console.WriteLine($"Mapping {groupName} group to {repo_GrouopToBeMovedUnder.Name} repository");
                        }
                        // Get the acls for the area path
                        // Add group to the area path security with read/write perms for work items in this area path

                        SecurityHttpClient securityClient = connection.GetClient<SecurityHttpClient>();
                        IEnumerable<AccessControlList> acls = securityClient.QueryAccessControlListsAsync(securityNamespaceId, null, null, false, false).Result;
                        AccessControlList areaPathAcl = acls.FirstOrDefault(x => x.Token.Contains(areaPath.Identifier.ToString()));
                        AccessControlEntry entry = new AccessControlEntry(group.Descriptor, 48, 0, null);
                        var aces = securityClient.SetAccessControlEntriesAsync(securityNamespaceId, areaPathAcl.Token, new List<AccessControlEntry> { entry }, false).Result;

                        // Get acls for project
                        IEnumerable<AccessControlList> aclsProject = securityClient.QueryAccessControlListsAsync(projectSecurityNamespaceId, null, null, false, false).Result;
                        AccessControlList projectAcl = aclsProject.FirstOrDefault(x => x.Token.Contains(project.Id.ToString()));
                        AccessControlEntry projectEntry = new AccessControlEntry(group.Descriptor, 1, 0, null);
                        var acesP = securityClient.SetAccessControlEntriesAsync(projectSecurityNamespaceId, projectAcl.Token, new List<AccessControlEntry> { projectEntry }, false).Result;

                        //GetACL for repository
                        if (repo_GrouopToBeMovedUnder != null && !string.IsNullOrEmpty(repo_GrouopToBeMovedUnder.Name))
                        {
                            IEnumerable<AccessControlList> aclsRepo = securityClient.QueryAccessControlListsAsync(gitRepoNamespaceId, null, null, false, false).Result;
                            AccessControlList RepoAcl = aclsRepo.FirstOrDefault(x => x.Token.Contains(repo_GrouopToBeMovedUnder.Id.ToString()));
                            AccessControlEntry repoEntry = new AccessControlEntry(group.Descriptor, 16502, 0, null);
                            var acesRepo = securityClient.SetAccessControlEntriesAsync(gitRepoNamespaceId, RepoAcl.Token, new List<AccessControlEntry> { repoEntry }, false).Result;
                        }
                    }
                }
                // Get the area path
                Console.WriteLine("Successfully added your group to the area path.");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        private static TeamProject GetProject(VssConnection connection, string projectName)
        {
            ProjectHttpClient projectClient = connection.GetClient<ProjectHttpClient>();
            IEnumerable<TeamProjectReference> projects = projectClient.GetProjects(top: 10000).Result;

            TeamProjectReference project = projects.FirstOrDefault(p => p.Name.Equals(projectName, StringComparison.OrdinalIgnoreCase));

            return projectClient.GetProject(project.Id.ToString(), true).Result;
        }

        private static Identity GetProjectGroup(VssConnection connection, string groupName, string projectName)
        {
            GraphHttpClient graphClient = connection.GetClient<GraphHttpClient>();
            PagedGraphGroups groups = graphClient.ListGroupsAsync().Result;
            IdentityHttpClient _identityClient;

            var projectClient = connection.GetClient<ProjectHttpClient>();
            TeamProject teamProject = projectClient.GetProject(projectName).Result;

            _identityClient = connection.GetClient<IdentityHttpClient>();
            IdentitiesCollection groups1 = _identityClient.ListGroupsAsync(new Guid[] { teamProject.Id }).Result;
            var group = groups1.Where(x => x.DisplayName.EndsWith(groupName)).SingleOrDefault();

            return group;
        }
        private static Identity GetGroup(VssConnection connection, string groupName, string projectName)
        {
            GraphHttpClient graphClient = connection.GetClient<GraphHttpClient>();

            PagedGraphGroups groups = graphClient.ListGroupsAsync().Result;
            // This program assumes that the group we need is in the first batch of groups returned by the api. Ideally you need to page through
            // the api results to find your group.
            //GraphGroup group = groups.GraphGroups.FirstOrDefault(x => x.PrincipalName.Equals(groupName, StringComparison.OrdinalIgnoreCase));
            GraphGroup group = groups.GraphGroups.FirstOrDefault(x => x.PrincipalName.Equals("[" + projectName + "]\\" + groupName, StringComparison.OrdinalIgnoreCase));
            string xm = JsonConvert.SerializeObject(groups.GraphGroups);
            GraphStorageKeyResult storageKey = graphClient.GetStorageKeyAsync(group.Descriptor).Result;

            Guid id = storageKey.Value;

            IdentityHttpClient identityClient = connection.GetClient<IdentityHttpClient>();
            return identityClient.ReadIdentityAsync(id).Result;
        }

        //public static List<Groups> ExportAreas(bool hasHeader = true)
        //{
        //    string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Setting\\Groups.xlsx");

        //    using (var pck = new ExcelPackage())
        //    {
        //        using (var stream = File.OpenRead(path))
        //        {
        //            pck.Load(stream);
        //        }
        //        var ws = pck.Workbook.Worksheets.First();
        //        DataTable tbl = new DataTable();
        //        foreach (var firstRowCell in ws.Cells[1, 1, 1, ws.Dimension.End.Column])
        //        {
        //            tbl.Columns.Add(hasHeader ? firstRowCell.Text : string.Format("Column {0}", firstRowCell.Start.Column));
        //        }
        //        var startRow = hasHeader ? 2 : 1;
        //        for (int rowNum = startRow; rowNum <= ws.Dimension.End.Row; rowNum++)
        //        {
        //            var wsRow = ws.Cells[rowNum, 1, rowNum, ws.Dimension.End.Column];
        //            DataRow row = tbl.Rows.Add();
        //            foreach (var cell in wsRow)
        //            {
        //                row[cell.Start.Column - 1] = cell.Text;
        //            }
        //        }
        //        string JSONString = string.Empty;
        //        var abc = JsonConvert.DeserializeObject<List<Groups>>(JsonConvert.SerializeObject(tbl));
        //        return JsonConvert.DeserializeObject<List<Groups>>(JsonConvert.SerializeObject(tbl));
        //    }
        //}

        //public static string ConcatValues(string area, string subarea, string group)
        //{
        //    if (!string.IsNullOrEmpty(area) && !string.IsNullOrEmpty(subarea) && !string.IsNullOrEmpty(group))
        //    {
        //        return string.Format("{0}_{1}", subarea, group);
        //    }
        //    else if (!string.IsNullOrEmpty(area) && string.IsNullOrEmpty(subarea) && !string.IsNullOrEmpty(group))
        //    {
        //        return string.Format("{0}_{1}", area, group);
        //    }
        //    else
        //    {
        //        return "";
        //    }
        //}

        //public static string ConcatAreaValues(string area, string subarea)
        //{
        //    if (!string.IsNullOrEmpty(area) && !string.IsNullOrEmpty(subarea))
        //    {
        //        return string.Format("{0}\\{1}", area, subarea);
        //    }
        //    else if (!string.IsNullOrEmpty(area) && string.IsNullOrEmpty(subarea))
        //    {
        //        return string.Format("{0}", area);
        //    }
        //    else
        //    {
        //        return "";
        //    }
        //}

        #region permission
        //Area Path Level
        //3, 0 Edit this node
        //     View permission for this node
        //17,0 View permissions for this node
        //      View work items in this node
        //48, 0 View work items in this node
        //      Edit work items in this node
        //49, 0 View permissions for this node
        //241, 0 Edit work items in this node
        //Manage test plans	
        //Manage test suites
        //View permissions for this node
        //View work items in this node

        //Project Level

        //513 
        //View project-level information	Allow
        //View test runs

        //15989759
        //Bypass rules on work item updates   Allow
        //Change process of team project.Allow
        //Create tag definition Not set
        //Create test runs    Allow
        //Delete and restore work items   Allow
        //Delete team project Allow
        //Delete test runs    Allow
        //Edit project-level information  Allow
        //Manage project properties   Allow
        //Manage test configurations  Allow
        //Manage test environments    Allow
        //Move work items out of this project Allow
        //Permanently delete work items   Allow
        //Rename team project Allow
        //Suppress notifications for work item updates Allow
        //Update project visibility Allow
        //View project-level information  Allow
        //View test runs

        //1
        //View project-level information  Allow

        //7033
        //Create test runs    Allow
        //Delete test runs    Allow
        //Manage test configurations  Allow
        //Manage test environments    Allow
        //View project-level information  Allow
        //View test runs  Allow


        // 15145
        // Create test runs Allow
        //Delete and restore work items Allow
        //Delete test runs Allow
        //Manage test configurations Allow
        //Manage test environments Allow
        //View project-level information  Allow
        //View test runs  Allow

        //112
        //all not set
        #endregion

        public static void CreateProjectVSTSGroup(VssConnection connection, Guid projectId, string groupName)
        {
            // get the project scope descriptor
            //
            GraphHttpClient graphClient = connection.GetClient<GraphHttpClient>();
            GraphDescriptorResult projectDescriptor = graphClient.GetDescriptorAsync(projectId).Result;

            // create a group at the project level
            // 
            GraphGroupCreationContext createGroupContext = new GraphGroupVstsCreationContext
            {
                DisplayName = groupName,
                Description = "Group at project level created via client library"
            };

            GraphGroup newGroup = graphClient.CreateGroupAsync(createGroupContext, projectDescriptor.Value).Result;
            string groupDescriptor = newGroup.Descriptor;
            Console.WriteLine(groupDescriptor);
        }
    }
}
