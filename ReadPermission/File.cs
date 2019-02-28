using Coll = System.Data;
using Microsoft.SharePoint.Client;
using System;
using System.Text;
using System.Collections.Generic;
using System.Collections;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Net;
using Microsoft.Office.Interop.Excel;
using File = System.IO.File;


namespace MakeSiteReadOnly
{
    public class SiteCollectionData
    {
        public string SiteCollection { get; set; }
        public string SkipSubSites { get; set; }
    }

    class Program
    {
        static string path = @"E:\AJ\PoC_MakeSiteReadOnly\";//not used
        static System.Data.DataTable daTable = new System.Data.DataTable();

        static void Main(string[] args)
        {

            SetSiteCollectionPermission();
            createExcel(daTable);
            Console.ReadLine();

        }
        static void createExcel(Coll.DataTable daTable)
        {
            try
            {
                Application excel = new Microsoft.Office.Interop.Excel.Application();
                excel.Visible = false;
                excel.DisplayAlerts = false;
                Workbook worKbooK = excel.Workbooks.Add(Type.Missing);


                Worksheet worKsheeT = (Microsoft.Office.Interop.Excel.Worksheet)worKbooK.ActiveSheet;
                worKsheeT.Name = "Permissions";

                worKsheeT.Range[worKsheeT.Cells[1, 1], worKsheeT.Cells[1, 8]].Merge();
                worKsheeT.Cells[1, 1] = "Site Permisison Data";
                worKsheeT.Cells.Font.Size = 15;
                Range celLrangE = null;

                int rowcount = 2;

                foreach (Coll.DataRow datarow in daTable.Rows)
                {
                    rowcount += 1;
                    for (int i = 1; i <= daTable.Columns.Count; i++)
                    {

                        if (rowcount == 3)
                        {
                            worKsheeT.Cells[2, i] = daTable.Columns[i - 1].ColumnName;
                            worKsheeT.Cells.Font.Color = System.Drawing.Color.Black;

                        }

                        worKsheeT.Cells[rowcount, i] = datarow[i - 1].ToString();

                        if (rowcount > 3)
                        {
                            if (i == daTable.Columns.Count)
                            {
                                if (rowcount % 2 == 0)
                                {
                                    celLrangE = worKsheeT.Range[worKsheeT.Cells[rowcount, 1], worKsheeT.Cells[rowcount, daTable.Columns.Count]];
                                }

                            }
                        }

                    }

                }

                celLrangE = worKsheeT.Range[worKsheeT.Cells[1, 1], worKsheeT.Cells[rowcount, daTable.Columns.Count]];
                celLrangE.EntireColumn.AutoFit();
                Microsoft.Office.Interop.Excel.Borders border = celLrangE.Borders;
                border.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                border.Weight = 2d;

                celLrangE = worKsheeT.Range[worKsheeT.Cells[1, 1], worKsheeT.Cells[2, daTable.Columns.Count]];

                worKbooK.SaveAs("E:\\Testing.xlsx"); ;
                worKbooK.Close();
                excel.Quit();

            }
            catch (Exception ex)
            {
                Console.WriteLine("Error Found in " + ex.StackTrace + "and the error is" + ex.Message.ToString());

            }
            finally
            {
                //worKsheeT = null;
                //celLrangE = null;
                //worKbooK = null;
            }

        }
        static void SetSiteCollectionPermission()
        {
            try
            {

                daTable.Columns.Add("SiteName", typeof(string));
                daTable.Columns.Add("SiteURL", typeof(string));
                daTable.Columns.Add("GroupName/UserName", typeof(string));

                daTable.Columns.Add("List/Library/FolderName", typeof(string));
                daTable.Columns.Add("PermissionName", typeof(string));


                List<SiteCollectionData> siteData = GetRecordsfromCsv();
                foreach (SiteCollectionData _data in siteData)
                {
                    string[] arrSubSites = _data.SkipSubSites.Split(',');
                    ClientContext clientContext = new ClientContext(_data.SiteCollection);
                    //clientContext.Credentials = new NetworkCredential(@"ASSOCIATES\FastTrackMS10","F@stTr@ckMS10#123","hcltech");
                    LogSummery("Under the context of SiteCollection : " + _data.SiteCollection);
                    Web oWebsite = clientContext.Web;
                    clientContext.Load(oWebsite,
                    website => website.Webs,
                    website => website.Title,
                    website => website.Lists
                    );
                    clientContext.ExecuteQuery();
                    //Get SiteCollection Level things first
                    Console.WriteLine("Working for Site-Collection: " + oWebsite.Title + " ....");
                    LogSummery("Calling function to make the site collection: " + oWebsite.Title + " read only.");
                    AssignGroupReadPermission(_data.SiteCollection);
                    AssignFolderReadPermission(_data.SiteCollection);
                    Console.WriteLine("Completed for Site-Collection: " + oWebsite.Title);
                    Console.WriteLine("=============x====================x==========================");

                    // Get all groups of web and assign them read permission
                    LogSummery("Start reading the sub-sites to break the inheritance of selected sub-sites.");
                    // Stop inheritance for selected sub-site..
                    foreach (Web oweb in oWebsite.Webs)
                        SetReadPermissionOnSubSites(arrSubSites, clientContext, oweb);

                }
                Console.WriteLine("Operation Completed successfullyy..");
                Console.ReadLine();

            }
            catch (Exception ex)
            {
                Console.WriteLine("Error Found in " + ex.StackTrace + "and the error is" + ex.Message.ToString());
                Console.WriteLine("Enter to continue..");
                Console.ReadLine();
                throw;
            }

        }

        private static void SetReadPermissionOnSubSites(string[] arrSubSites, ClientContext clientContext, Web oweb)
        {
            var subsitetitle = oweb.ServerRelativeUrl.Split('/');
            LogSummery("Working for sub-site: " + subsitetitle);
            var title = subsitetitle[subsitetitle.Length - 1];
            string match = arrSubSites.FirstOrDefault(s => title.Equals(s));
            LogSummery("Breaking inheritance for sub-site :" + match);
            if (!String.IsNullOrEmpty(match))
            {
                //// If match for skip then do nothing..
                //Below code is to break the inheritance for a web..
                //oweb.BreakRoleInheritance(true, true);
                //try
                //{
                //    clientContext.Load(oweb);
                //    clientContext.ExecuteQuery();
                //    LogSummery("Breake inheritance process done for sub-site :" + match);
                //}
                //catch (Exception ex) { LogSummery("Breake inheritance process fail for sub-site :" + match + ". Because " + ex.Message.ToString()); }
            }
            else
            {
                try
                {
                    LogSummery("Checking for unique role assignment for sub-site :" + title);
                    clientContext.Load(oweb, w => w.HasUniqueRoleAssignments);
                    clientContext.ExecuteQuery();
                    LogSummery("sub-site :" + subsitetitle + " has unique role assignment.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error Found in " + ex.StackTrace + "and the error is" + ex.Message.ToString());
                    LogSummery("fail to get unique role assignment for sub-site :" + oweb.Title + " Because " + ex.Message.ToString());
                }

                if (String.IsNullOrEmpty(match) && oweb.HasUniqueRoleAssignments)
                {
                    Console.WriteLine("Working for Sub-Site: " + title + " ......");
                    LogSummery("Sub-site which has unique permission and not in excluded list, assign them read permission :" + oweb.Title);
                    AssignGroupReadPermission(oweb.Url);
                    AssignFolderReadPermission(oweb.Url);
                    Console.WriteLine("Completed permission for Sub-Site: " + title);
                }
                if (String.IsNullOrEmpty(match) && (!oweb.HasUniqueRoleAssignments))
                {
                    Console.WriteLine("Working for Sub-Site: " + title + " ......");
                    LogSummery("Sub-site which inherit permission and not in exclusion list. Checking for any library or folder with unique permission to make them Read only :" + oweb.Title);
                    AssignFolderReadPermission(oweb.Url);
                    Console.WriteLine("Completed permission for Sub-Site: " + title);
                }
            }
        }
        static List<SiteCollectionData> GetRecordsfromCsv()
        {
            List<SiteCollectionData> siteData = new List<SiteCollectionData>();

            Application ap = new Application();
            string exeDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            Workbook wb = ap.Workbooks.Open(Path.Combine(exeDir, "Report.xlsx"));
            Worksheet ws = wb.ActiveSheet;
            bool flag = true;
            int counter = 2;
            int index = 0;
            while (flag)
            {
                string columnHeader = "A" + counter.ToString();
                Range row = ws.get_Range(columnHeader);
                row.Value2 = row.Text.ToString().Trim();
                if (row.Value2 != null)
                {
                    SiteCollectionData listItem = new SiteCollectionData();
                    Range Exprow = ws.get_Range("B" + counter.ToString());
                    string lstSubSite = Convert.ToString(Exprow.Value2);
                    listItem.SiteCollection = row.Value2;
                    listItem.SkipSubSites = lstSubSite;
                    siteData.Insert(index, listItem);
                }
                else
                { flag = false; }
                index++;
                counter++;
            }
            return siteData;

        }
        static void LogSummery(string Message)
        {
            string date = string.Format("{0:dd-MM-yyyy}", DateTime.Now);
            string fileName = "log_" + date + ".txt";
            System.IO.File.AppendAllText(fileName, DateTime.Now.ToString() + "\t" + "Summery : " + "\t" + Message.ToString() + Environment.NewLine);

        }

        static void AssignGroupReadPermission(string oWebUrl) //,Web oWebs)
        {
            //Set groups Read on site collection level.
            LogSummery("==============Site Group -- Start Reading group for site-collection to assign them Read Only permission.====================");
            ClientContext ctx = new ClientContext(oWebUrl);
            Web oWebsite = ctx.Web;
            ctx.Load(oWebsite, w => w.Title);
            ctx.ExecuteQuery();
            #region Group and User Read Permission
            //var groups = oWebsite.SiteGroups;
            RoleAssignmentCollection roleAssignments = oWebsite.RoleAssignments;
            ctx.Load(roleAssignments);
            try
            {
                ctx.ExecuteQuery();
            }
            catch (Exception ex) { Console.WriteLine("Error Found in " + ex.StackTrace + "and the error is" + ex.Message.ToString()); }
            foreach (RoleAssignment roleAssignment in roleAssignments.ToList())
            {
                ctx.Load(roleAssignment, s => s.Member.Title, s => s.Member.LoginName, s => s.Member.Id, s => s.Member.PrincipalType, s => s.RoleDefinitionBindings.Include(d => d.Name));
                try
                {
                    ctx.ExecuteQuery();
                }
                catch (Exception ex) { Console.WriteLine("Error Found in " + ex.StackTrace + "and the error is" + ex.Message.ToString()); }
                if (roleAssignment.Member.PrincipalType.ToString() == "SharePointGroup")
                {
                    LogSummery("Site Group -- Working on group : " + roleAssignment.Member.Title + " has permission: " + (roleAssignment.RoleDefinitionBindings[0].Name));
                    daTable.Rows.Add(oWebsite.Title, oWebUrl, roleAssignment.Member.Title, "SharePointGroup", roleAssignment.RoleDefinitionBindings[0].Name);

                }
                else if (roleAssignment.Member.PrincipalType.ToString() == "User")
                {
                    User user = null;
                    var userLoginName = roleAssignment.Member.LoginName.Split('|')[1];
                    if (userLoginName != "hcltech\\devadmin")
                    {
                        user = ctx.Web.EnsureUser(userLoginName);

                        ctx.Load(user);
                        ctx.ExecuteQuery();
                    }
                    LogSummery("Site Group -- Working on User : " + roleAssignment.Member.Title + " has permission:" + (roleAssignment.RoleDefinitionBindings[0].Name));
                    daTable.Rows.Add(oWebsite.Title, oWebUrl, userLoginName = (userLoginName == "hcltech\\devadmin") ? userLoginName : userLoginName + "(" + user.Email + ")", "User", roleAssignment.RoleDefinitionBindings[0].Name);
                }
                //roleAssignment.RoleDefinitionBindings.RemoveAll();
                //roleAssignment.Update();
                try
                {
                    ctx.Load(roleAssignment);
                    ctx.ExecuteQuery();
                    LogSummery("Site Group -- Successfully removed the existing permission for : " + roleAssignment.Member.Title);
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error Found in " + ex.StackTrace + "and the error is" + ex.Message.ToString());
                    LogSummery("Site Group -- Fail to remove the existing permission for : " + roleAssignment.Member.Title + ". Baucase " + ex.Message.ToString());
                }
                //Assign Read permission..
                //RoleDefinition roleDefinition = ctx.Web.RoleDefinitions.GetByType(RoleType.Reader);
                //RoleDefinitionBindingCollection roleDefinitionBindingColl = new RoleDefinitionBindingCollection(ctx);
                // roleDefinitionBindingColl.Add(roleDefinition);
                //RoleAssignment roleAssign = ctx.Web.RoleAssignments.Add(roleAssignment.Member, roleDefinitionBindingColl);

                try
                {
                    //ctx.ExecuteQuery();
                    LogSummery("Site Group -- Successfully assigned Read permission on group : " + roleAssignment.Member.Title);
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error Found in " + ex.StackTrace + "and the error is" + ex.Message.ToString());
                    LogSummery("Site Group -- Fail to assign Read permission on group : " + roleAssignment.Member.Title + ". Baucause " + ex.Message.ToString());
                }
            #endregion
                #region User Read permission
                //LogSummery("Start assigning read permission to Users of web..");
                //UserCollection oUsers = ctx.Web.SiteUsers;
                //try
                //{
                //    ctx.Load(oUsers);
                //    ctx.ExecuteQuery();
                //    LogSummery("Found total users : " + oUsers.Count);
                //}
                //catch (Exception ex)
                //{
                //    LogSummery("Fail in getting users bacause " + ex.Message.ToString());
                //}

                //foreach (User oUser in oUsers)
                //{
                //    if (!oUser.IsSiteAdmin)
                //    {
                //        LogSummery("Working for user : " + oUser.Title);
                //        //Remove the existing permission..
                //        RoleAssignment userRoleAssiment = oWebsite.RoleAssignments.GetByPrincipal(oUser);
                //        userRoleAssiment.RoleDefinitionBindings.RemoveAll();
                //        userRoleAssiment.Update();
                //        try
                //        {
                //            ctx.Load(userRoleAssiment);
                //            ctx.ExecuteQuery();

                //        }
                //        catch (Exception ex) { LogSummery("Fail in getting user role assignment because : " + ex.Message.ToString() + "Error on line number 268."); }
                //        // Assign Read permission..
                //        RoleDefinition userRoleDefinition = ctx.Web.RoleDefinitions.GetByType(RoleType.Reader);
                //        RoleDefinitionBindingCollection userRoleDefinitionBindingColl = new RoleDefinitionBindingCollection(ctx);
                //        userRoleDefinitionBindingColl.Add(userRoleDefinition);
                //        RoleAssignment userRoleAssign = ctx.Web.RoleAssignments.Add(oUser, userRoleDefinitionBindingColl);

                //        try
                //        {
                //            ctx.ExecuteQuery();

                //        }
                //        catch (Exception ex) { LogSummery("Error in setting read permission for User " + oUser.Title + ". Because " + ex.Message.ToString() + "Error on line number 281."); }
                //    }
                //}
                //daTable.Rows.Add(oWebsite.Title, oWebUrl, roleAssignment.Member.Title, "","", roleAssignment.RoleDefinitionBindings[0].Name);
                #endregion

            }
            LogSummery("==============Site Group -- Completed Reading group for site-collection to assign them Read Only permission.====================");
        }
        //static void ResetToRead(ClientObject obj, string objType, ClientContext ctx)
        //{
        //    //Remove the existing permission..
        //    var objPrinciple;
        //    if(objType=="User"){
        //        objPrinciple = obj as User;
        //    }
        //    else if (objType == "Group")
        //    {
        //        objPrinciple = ctx.CastTo<Group>(obj);
        //    }
        //    RoleAssignment groupRoleAssiment = oWebsite.RoleAssignments.GetByPrincipal(group);
        //    groupRoleAssiment.RoleDefinitionBindings.RemoveAll();
        //    groupRoleAssiment.Update();
        //    try
        //    {
        //        ctx.Load(groupRoleAssiment);
        //        ctx.ExecuteQuery();
        //        LogSummery("Successfully removed the existing permission on group : " + group.Title);
        //    }
        //    catch (Exception ex) { LogSummery("Fail to remove the existing permission on group : " + group.Title + ". Baucase " + ex.Message.ToString()); }
        //    //Assign Read permission..
        //    RoleDefinition roleDefinition = ctx.Web.RoleDefinitions.GetByType(RoleType.Reader);
        //    RoleDefinitionBindingCollection roleDefinitionBindingColl = new RoleDefinitionBindingCollection(ctx);
        //    roleDefinitionBindingColl.Add(roleDefinition);
        //    RoleAssignment roleAssign = ctx.Web.RoleAssignments.Add(group, roleDefinitionBindingColl);

        //    try
        //    {
        //        ctx.ExecuteQuery();
        //        LogSummery("Successfully assigned Read permission on group : " + group.Title);
        //    }
        //    catch (Exception ex) { LogSummery("Fail to assign Read permission on group : " + group.Title + ". Baucause " + ex.Message.ToString()); }
        //}

        static void AssignFolderReadPermission(string oWebUrl)
        {
            LogSummery("=============Lib & Folder -- Start Working on unique Libraries & folders to make them read only.===============");
            ClientContext ctxFolder = new ClientContext(oWebUrl);
            Web oWebsite = ctxFolder.Web;
            ctxFolder.Load(oWebsite,
            website => website.Webs,
            website => website.Title,
            website => website.Url
            );
            ctxFolder.ExecuteQuery();
            LogSummery("Library -- Getting document libraries for web " + oWebsite.Url);
            //var Libraries = ctxFolder.LoadQuery(ctxFolder.Web.Lists.Where(l => l.BaseTemplate == 101).Include(l => l.HasUniqueRoleAssignments, l => l.Title, l => l.RoleAssignments));
            var Libraries = ctxFolder.LoadQuery(ctxFolder.Web.Lists.Include(l => l.HasUniqueRoleAssignments, l => l.Title, l => l.RoleAssignments));
            ctxFolder.ExecuteQuery();
            LogSummery("Library -- Checking all libraries for unique folders under ther web " + oWebsite.Url);
            foreach (List _oDoc in Libraries)
            {
                LogSummery("Library -- Checking for library " + _oDoc.Title + " if it has unique permission then setting read only.");
                ctxFolder.Load(_oDoc, o => o.RootFolder.ServerRelativeUrl);
                ctxFolder.ExecuteQuery();
                //Console.WriteLine("VipinTest" + _oDoc.RootFolder.ServerRelativeUrl);
                #region Unique Folder to read only.
                if (_oDoc.HasUniqueRoleAssignments)
                {
                    LogSummery("Library -- Found folder " + _oDoc.Title + " as unique library. Now setting it to read only.");
                    // RoleAssignmentCollection roleDocAssignments = _oDoc.RoleAssignments;

                    var roleDocAssignments = ctxFolder.LoadQuery(_oDoc.RoleAssignments.Include(l => l.Member.Title, l => l.Member.LoginName, l => l.Member.Id, l => l.PrincipalId, l => l.Member.PrincipalType, l => l.RoleDefinitionBindings.Include(d => d.Name)));
                    ctxFolder.ExecuteQuery();
                    foreach (var assignemtn in roleDocAssignments)
                    {
                        //Console.WriteLine(assignemtn.Member.Title);
                        //Console.WriteLine(assignemtn.Member.LoginName);
                        LogSummery("Library -- Working on group : " + assignemtn.Member.Title + " has permission: " + (assignemtn.RoleDefinitionBindings[0].Name));
                        if (assignemtn.Member.PrincipalType.ToString() == "SharePointGroup")
                        {
                            daTable.Rows.Add(oWebsite.Title, oWebUrl, assignemtn.Member.Title, _oDoc.RootFolder.ServerRelativeUrl, assignemtn.RoleDefinitionBindings[0].Name);

                        }
                        else if (assignemtn.Member.PrincipalType.ToString() == "User")
                        {
                            User user = null;
                            var userLoginName = assignemtn.Member.LoginName.Split('|')[1];
                            if (userLoginName != "hcltech\\devadmin")
                            {
                                user = ctxFolder.Web.EnsureUser(userLoginName);
                                ctxFolder.Load(user);
                                ctxFolder.ExecuteQuery();
                            }

                            LogSummery("Site Group -- Working on User : " + assignemtn.Member.Title + " has permission:" + (assignemtn.RoleDefinitionBindings[0].Name));

                            daTable.Rows.Add(oWebsite.Title, oWebUrl, userLoginName = (userLoginName == "hcltech\\devadmin") ? userLoginName : userLoginName + "(" + user.Email + ")", _oDoc.RootFolder.ServerRelativeUrl, assignemtn.RoleDefinitionBindings[0].Name);
                        }
                        RoleAssignment groupRoleAssiment = _oDoc.RoleAssignments.GetByPrincipal(assignemtn.Member);
                        //groupRoleAssiment.RoleDefinitionBindings.RemoveAll();
                        //groupRoleAssiment.Update();
                        try
                        {
                            ctxFolder.Load(groupRoleAssiment);
                            ctxFolder.ExecuteQuery();

                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("Error Found in " + ex.StackTrace + "and the error is" + ex.Message.ToString());
                            LogSummery("Library -- Fail in removing existing role assignment on library " + _oDoc.Title + " because : " + ex.Message.ToString());
                        }
                        // Assign Read permission..
                        //RoleDefinition userRoleDefinition = ctxFolder.Web.RoleDefinitions.GetByType(RoleType.Reader);
                        //RoleDefinitionBindingCollection userRoleDefinitionBindingColl = new RoleDefinitionBindingCollection(ctxFolder);
                        //userRoleDefinitionBindingColl.Add(userRoleDefinition);
                        //RoleAssignment userRoleAssign = _oDoc.RoleAssignments.Add(assignemtn.Member, userRoleDefinitionBindingColl);

                        try
                        {
                            ctxFolder.ExecuteQuery();

                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("Error Found in " + ex.StackTrace + "and the error is" + ex.Message.ToString());
                            LogSummery("Library -- Fail in setting library to read-only because : " + ex.Message.ToString() + "Error on line number 268.");
                        }

                    }
                }
                #endregion
                ctxFolder.Load(_oDoc, l => l.BaseType);
                try
                {
                    ctxFolder.ExecuteQuery();
                }
                catch (Exception ex) { }
                if (_oDoc.BaseType == BaseType.DocumentLibrary)
                {
                    LogSummery("Folders -- Getting all folders for library " + _oDoc.Title);
                    var folders = _oDoc.GetItems(CamlQuery.CreateAllFoldersQuery());
                    //var folders = _oDoc.GetItems(_query);
                    ctxFolder.Load(folders, icol => icol.Include(i => i.RoleAssignments.Include(ra => ra.Member, ra => ra.Member.LoginName, ra => ra.Member.PrincipalType), i => i.DisplayName, i => i.Folder));
                    try
                    {
                        ctxFolder.ExecuteQuery();
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Error Found in " + ex.StackTrace + "and the error is" + ex.Message.ToString());
                        LogSummery("Folders -- Fail to load folder " + _oDoc.Title + " baucause " + ex.Message);
                    }
                    LogSummery("Folders -- Checking all folders of library " + _oDoc.Title + "for unique permission to male them read only");
                    LogSummery("Folders -- Found total folders " + folders.Count + " to work.");
                    foreach (var _oFolder in folders)
                    {
                        try
                        {


                            ctxFolder.Load(_oFolder, f => f.HasUniqueRoleAssignments, f => f["FileDirRef"]);
                            try
                            {
                                ctxFolder.ExecuteQuery();
                            }
                            catch (Exception ex) { }
                            if (_oFolder.HasUniqueRoleAssignments)
                            {
                                Console.WriteLine((string)_oFolder["FileDirRef"]);
                                var parentFolder = _oFolder.ParentList.ParentWeb.GetFolderByServerRelativeUrl((string)_oFolder["FileDirRef"]);
                                _oFolder.Context.Load(parentFolder);
                                _oFolder.Context.ExecuteQuery();
                                Console.WriteLine(parentFolder.ServerRelativeUrl);
                                LogSummery("Folders -- Folder: " + _oFolder.DisplayName + " found unique permission. Working on to make it read only.");

                                //Console.WriteLine(_oFolder.DisplayName);

                                var roleAssignments = ctxFolder.LoadQuery(_oFolder.RoleAssignments.Include(i => i.Member.Id, i => i.Member.LoginName, i => i.Member.PrincipalType, i => i.Member.Title, i => i.RoleDefinitionBindings.Include(s => s.Name)));
                                ctxFolder.ExecuteQuery();
                                foreach (var assignemtn in roleAssignments.ToList())
                                {
                                    try
                                    {
                                        //daTable.Rows.Add(oWebsite.Title, oWebsite.Url, sGroupName, 78, 59, 72, 95, 83, 77);
                                        LogSummery("Folders -- Working on group : " + assignemtn.Member.Title + " has permission: " + (assignemtn.RoleDefinitionBindings[0].Name));
                                        if (assignemtn.Member.PrincipalType.ToString() == "SharePointGroup")
                                        {
                                            daTable.Rows.Add(oWebsite.Title, oWebUrl, assignemtn.Member.Title, parentFolder.ServerRelativeUrl + "/" + _oFolder.DisplayName, assignemtn.RoleDefinitionBindings[0].Name);
                                        }
                                        else if (assignemtn.Member.PrincipalType.ToString() == "User")
                                        {
                                            User user = null;
                                            var userLoginName = assignemtn.Member.LoginName.Split('|')[1];
                                            if (userLoginName != "hcltech\\devadmin")
                                            {
                                                user = ctxFolder.Web.EnsureUser(userLoginName);
                                                ctxFolder.Load(user);
                                                ctxFolder.ExecuteQuery();
                                            }

                                            LogSummery("Site Group -- Working on User : " + assignemtn.Member.Title + " has permission:" + (assignemtn.RoleDefinitionBindings[0].Name));

                                            daTable.Rows.Add(oWebsite.Title, oWebUrl, userLoginName = (userLoginName == "hcltech\\devadmin") ? userLoginName : userLoginName + "(" + user.Email + ")", parentFolder.ServerRelativeUrl + "/" + _oFolder.DisplayName, assignemtn.RoleDefinitionBindings[0].Name);
                                        }
                                        //Console.WriteLine(assignemtn.Member.Title);
                                        //Console.WriteLine(assignemtn.Member.LoginName);
                                        RoleAssignment groupRoleAssiment = _oFolder.RoleAssignments.GetByPrincipal(assignemtn.Member);
                                        //groupRoleAssiment.RoleDefinitionBindings.RemoveAll();
                                        // groupRoleAssiment.Update();
                                        ctxFolder.Load(groupRoleAssiment);
                                        ctxFolder.ExecuteQuery();
                                        // Assign Read permission..
                                        // RoleDefinition userRoleDefinition = ctxFolder.Web.RoleDefinitions.GetByType(RoleType.Reader);
                                        //RoleDefinitionBindingCollection userRoleDefinitionBindingColl = new RoleDefinitionBindingCollection(ctxFolder);
                                        //userRoleDefinitionBindingColl.Add(userRoleDefinition);
                                        //RoleAssignment userRoleAssign = _oFolder.RoleAssignments.Add(assignemtn.Member, userRoleDefinitionBindingColl);
                                        //ctxFolder.ExecuteQuery();
                                        LogSummery("Folders -- Folder read enabled operation successfull for fodler " + _oFolder.DisplayName);
                                    }
                                    catch (Exception ex)
                                    {
                                        Console.WriteLine("Error Found in " + ex.StackTrace + "and the error is" + ex.Message.ToString());
                                    }
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("Error Found in " + ex.StackTrace + "and the error is" + ex.Message.ToString());
                            LogSummery("Folders -- Fail in role assignment because : " + ex.Message.ToString() + " for folder " + _oFolder.DisplayName);
                        }
                    }
                }
                else
                {
                    LogSummery("Folders -- Getting all folders for List " + _oDoc.Title);
                    var folders = _oDoc.GetItems(CamlQuery.CreateAllFoldersQuery());
                    //var folders = _oDoc.GetItems(_query);
                    ctxFolder.Load(folders, icol => icol.Include(i => i.RoleAssignments.Include(ra => ra.Member), i => i.DisplayName, i => i.Folder));
                    try
                    {
                        ctxFolder.ExecuteQuery();
                    }
                    catch (Exception ex)
                    {
                        LogSummery("Folders -- Fail to load folder " + _oDoc.Title + " baucause " + ex.Message);
                    }
                    LogSummery("Folders -- Checking all folders of list " + _oDoc.Title + "for unique permission to make them read only");
                    LogSummery("Folders -- Found total folders " + folders.Count + " to work.");
                    foreach (var _oFolder in folders)
                    {
                        try
                        {


                            ctxFolder.Load(_oFolder, f => f.HasUniqueRoleAssignments, f => f["FileDirRef"]);
                            try
                            {
                                ctxFolder.ExecuteQuery();
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine("Error Found in " + ex.StackTrace + "and the error is" + ex.Message.ToString());
                            }
                            if (_oFolder.HasUniqueRoleAssignments)
                            {
                                Console.WriteLine((string)_oFolder["FileDirRef"]);
                                var parentFolder = _oFolder.ParentList.ParentWeb.GetFolderByServerRelativeUrl((string)_oFolder["FileDirRef"]);
                                _oFolder.Context.Load(parentFolder);
                                _oFolder.Context.ExecuteQuery();
                                Console.WriteLine(parentFolder.ServerRelativeUrl);
                                LogSummery("Folders -- Folder: " + _oFolder.DisplayName + " found unique permission. Working on to make it read only.");

                                //Console.WriteLine(_oFolder.DisplayName);

                                var roleAssignments = ctxFolder.LoadQuery(_oFolder.RoleAssignments.Include(i => i.Member.Id, i => i.Member.LoginName, i => i.Member.PrincipalType, i => i.Member.Title, i => i.RoleDefinitionBindings.Include(s => s.Name)));
                                ctxFolder.ExecuteQuery();
                                foreach (var assignemtn in roleAssignments.ToList())
                                {
                                    try
                                    {
                                        //daTable.Rows.Add(oWebsite.Title, oWebsite.Url, sGroupName, 78, 59, 72, 95, 83, 77);
                                        LogSummery("Folders -- Working on group : " + assignemtn.Member.Title + " has permission: " + (assignemtn.RoleDefinitionBindings[0].Name));

                                        if (assignemtn.Member.PrincipalType.ToString() == "SharePointGroup")
                                        {
                                            daTable.Rows.Add(oWebsite.Title, oWebUrl, assignemtn.Member.Title, parentFolder.ServerRelativeUrl + "/" + _oFolder.DisplayName, assignemtn.RoleDefinitionBindings[0].Name);
                                        }
                                        else if (assignemtn.Member.PrincipalType.ToString() == "User")
                                        {
                                            User user = null;
                                            var userLoginName = assignemtn.Member.LoginName.Split('|')[1];
                                            if (userLoginName != "hcltech\\devadmin")
                                            {
                                                user = ctxFolder.Web.EnsureUser(userLoginName);
                                                ctxFolder.Load(user);
                                                ctxFolder.ExecuteQuery();
                                            }

                                            LogSummery("Site Group -- Working on User : " + assignemtn.Member.Title + " has permission:" + (assignemtn.RoleDefinitionBindings[0].Name));

                                            daTable.Rows.Add(oWebsite.Title, oWebUrl, userLoginName = (userLoginName == "hcltech\\devadmin") ? userLoginName : userLoginName + "(" + user.Email + ")", parentFolder.ServerRelativeUrl + "/" + _oFolder.DisplayName, assignemtn.RoleDefinitionBindings[0].Name);
                                        }
                                        //Console.WriteLine(assignemtn.Member.Title);
                                        //Console.WriteLine(assignemtn.Member.LoginName);
                                        RoleAssignment groupRoleAssiment = _oFolder.RoleAssignments.GetByPrincipal(assignemtn.Member);
                                        //groupRoleAssiment.RoleDefinitionBindings.RemoveAll();
                                        // groupRoleAssiment.Update();
                                        ctxFolder.Load(groupRoleAssiment);
                                        ctxFolder.ExecuteQuery();
                                        // Assign Read permission..
                                        // RoleDefinition userRoleDefinition = ctxFolder.Web.RoleDefinitions.GetByType(RoleType.Reader);
                                        //RoleDefinitionBindingCollection userRoleDefinitionBindingColl = new RoleDefinitionBindingCollection(ctxFolder);
                                        //userRoleDefinitionBindingColl.Add(userRoleDefinition);
                                        //RoleAssignment userRoleAssign = _oFolder.RoleAssignments.Add(assignemtn.Member, userRoleDefinitionBindingColl);
                                        //ctxFolder.ExecuteQuery();
                                        LogSummery("Folders -- Folder read enabled operation successfull for fodler " + _oFolder.DisplayName);
                                    }
                                    catch (Exception ex)
                                    {
                                        Console.WriteLine("Error Found in " + ex.StackTrace + "and the error is" + ex.Message.ToString());
                                    }
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("Error Found in " + ex.StackTrace + "and the error is" + ex.Message.ToString());
                            LogSummery("Folders -- Fail in role assignment because : " + ex.Message.ToString() + " for folder " + _oFolder.DisplayName);
                        }
                    }
                }
            }
            LogSummery("=============Lib & Folder -- Completed Working on unique Libraries & folders to make them read only.===============");
        }

    }
}
