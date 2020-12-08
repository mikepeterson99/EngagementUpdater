using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Microsoft.ProjectServer.Client;
using Microsoft.SharePoint.Client;
using System.Configuration;
using System.Data.SqlClient;
using System.Data;


using Microsoft.Office.Project;
using Microsoft.Office.Project.Server.Schema;
using System.Web.Script.Serialization;

namespace EngagementUpdater
{
    class Program
    {

        private static AppSettingsReader asr = new AppSettingsReader();
        private static string pwaPath = "https://milestonecg.sharepoint.com/sites/RPMDev"; // RPMDev //MCGQA  PWADEv
        private static string cn = "Data Source = sqlmcg04az; Initial Catalog = ProjectOnlineMCGRPM; Integrated Security = True"; //ProjectOnlineMCGRPM

        private static ProjectContext projContext;
        private static int timeoutSeconds = 10;

        private static EnterpriseResource res;
        private static EnterpriseResourceCollection resources;
        private static DraftProject draftProj = null;
        private static PublishedProject pubProj;
        private static string m_proj_uid = "";

        static void Main(string[] args)
        {
            LoginPOL();
            SyncEngagments();
        }

        static void LoginPOL()
        {
            pwaPath = (string)asr.GetValue("pwa_url", typeof(string));
            cn = (string)asr.GetValue("pwa_db_cn", typeof(string));

            string Customer_Admin_User_Name = (string)asr.GetValue("Customer_Admin_User_Name", typeof(string));
            string Customer_Admin_Password = (string)asr.GetValue("Customer_Admin_Password", typeof(string));

            projContext = new ProjectContext(pwaPath);
            System.Security.SecureString pass = new System.Security.SecureString();
            Customer_Admin_Password.ToCharArray().ToList().ForEach(x => pass.AppendChar(x));
            projContext.Credentials = new SharePointOnlineCredentials(Customer_Admin_User_Name, pass);
        }

        static void SyncEngagments()
        {
            ProjectEngagement pEng = null;
            string res_uid = "";
            string proj_uid = "";
            string err = "";
            string sql = "exec [MCG_RM_GetPendingUpdates_engagements]";

            DataSet ds = new DataSet();
            SqlDataAdapter da = new SqlDataAdapter(sql, cn);
            da.Fill(ds);

            foreach (DataRow r in ds.Tables[0].Rows)
            {
                res_uid = r["resourceuid"].ToString();
                proj_uid = r["projectuid"].ToString();

                pubProj = projContext.Projects.GetByGuid(new Guid(proj_uid));
                projContext.Load(pubProj);
                projContext.ExecuteQuery();

                draftProj = pubProj.CheckOut();

                projContext.Load(draftProj.Engagements);
                projContext.ExecuteQuery();

                if (!r.IsNull("engagementuid"))
                {
                    pEng = draftProj.Engagements.GetByGuid(new Guid(r["engagementuid"].ToString()));
                }

                if (pEng == null)
                {
                    res = projContext.EnterpriseResources.GetByGuid(new Guid(res_uid));
                    projContext.Load(res);
                    projContext.ExecuteQuery();

                    ProjectEngagementCreationInformation peci = new ProjectEngagementCreationInformation();
                    peci.Id = Guid.NewGuid();
                    peci.Start = Convert.ToDateTime(r["start_dt"].ToString());
                    peci.Finish = Convert.ToDateTime(r["end_dt"].ToString());
                    peci.Resource = res;
                    peci.Work = "0h";
                    peci.Description = "RPM_" + r["planuid"].ToString();

                    draftProj.Engagements.Add(peci).Status = EngagementStatus.Proposed;
                    draftProj.Engagements.Update();
                    pEng = draftProj.Engagements.Last();
                    projContext.Load(pEng);
                    projContext.ExecuteQuery();
                }


                DataRow[] rows = ds.Tables[1].Select("resourceuid='" + res_uid + "' and projectuid='" + proj_uid + "'");
                ProjectEngagementTimephasedCollection petpc = pEng.GetTimephased(Convert.ToDateTime(r["start_dt"]), Convert.ToDateTime(r["end_dt"]), TimeScale.Days, EngagementContourType.Draft);

                projContext.Load(petpc);
                projContext.ExecuteQuery();

                foreach (DataRow row in rows)
                {
                    petpc.GetByStart(Convert.ToDateTime(row["timebyday"].ToString())).Work = row["allocationwork"].ToString();
                }
                pEng.Status = EngagementStatus.Reproposed; //this is needed
                draftProj.Engagements.Update();

                draftProj.CheckIn(false);
                //this updates the last TBD engagement in PWA, Approved will remove Eng from PWA
                QueueJob qJob1 = projContext.Projects.Update();
                JobState jobState = projContext.WaitForQueue(qJob1, timeoutSeconds);

                {
                    //approve proposed request
                    if (res == null)
                    {
                        res = projContext.EnterpriseResources.GetByGuid(new Guid(res_uid));
                        projContext.Load(res);
                        projContext.Load(res.Engagements);
                        projContext.ExecuteQuery();
                    }

                    ResourceEngagement eng = res.Engagements.GetById(pEng.Id.ToString());
                    projContext.Load(eng);
                    projContext.ExecuteQuery();  //Too many resources: 4205. You cannot load dependent objects for more than 1000 resources. Use a filter to restrict your query

                    eng.Status = EngagementStatus.Approved;
                    res.Engagements.Update();

                    QueueJob qJob = projContext.Projects.Update();
                    jobState = projContext.WaitForQueue(qJob, timeoutSeconds);
                }

            }

        }

        static void getEngagements()
        {
            string err = "";
            string sql = "exec [MCG_RM_GetPendingUpdates_engagements]";

            DataSet ds = new DataSet();
            SqlDataAdapter da = new SqlDataAdapter(sql, cn);
            da.Fill(ds);

            resources = projContext.EnterpriseResources;
            projContext.Load(resources);
            projContext.ExecuteQuery();

            testEngagement(ds);

            int x = 0;
            if (ds.Tables[0].Rows.Count > 0)
            {
                foreach (DataRow r in ds.Tables[0].Rows)
                {
                    checkEngagement(r["projectuid"].ToString());

                    //if (x >= 17)
                    {
                        try
                        {
                            createPWAEngagement(r["projectuid"].ToString(), r["resourceuid"].ToString(), r["start_dt"].ToString(), r["end_dt"].ToString(), Convert.ToDouble(r["pct1"].ToString()), 0, "", "");
                        }
                        catch (Exception ex)
                        {
                            err += r["projectuid"].ToString() + "|" + r["resourceuid"].ToString() + "|" + ex.Message + ";";
                        }
                    }
                    x = x + 1;
                }

                draftProj.CheckIn(false);

                QueueJob qJob = projContext.Projects.Update();
                JobState jobState = projContext.WaitForQueue(qJob, timeoutSeconds);
            }


            //createPWAEngagement("ccb09aba-3b48-e611-80ca-00155d784803", "2faccbc2-7cc0-e511-80e2-00155da47821", "3/8/17", "4/1/17", .5, 0, "project", "_test6"); //mcg dev
            //createPWAEngagement("7b0d8c9a-1f51-e611-80db-00155dac312b", "897dacfa-afba-e511-80f6-00155da88b1b", "4/3/17", "4/5/17", 0, 0, "resplan", "");  //tina
            //createPWAEngagement("102c1cf8-15a1-e611-80ce-00155db08b21", "6C8D9F86-775A-E611-80C6-00155DB0931F", "10/1/16", "1/31/17", 0, 0, "resplan", "Team project forecast");

        }

        static void testEngagement(DataSet ds)
        {
            string res_uid = "";
            string proj_uid = "";

            foreach (DataRow r in ds.Tables[0].Rows)
            {
                res_uid = r["resourceuid"].ToString();
                proj_uid = r["projectuid"].ToString();

                pubProj = projContext.Projects.GetByGuid(new Guid(proj_uid));
                projContext.Load(pubProj);
                projContext.ExecuteQuery();

                draftProj = pubProj.CheckOut();

                projContext.Load(draftProj.Engagements);
                projContext.ExecuteQuery();

                ProjectEngagement pEng = null;

                if (!r.IsNull("engagementuid"))
                {
                    pEng = draftProj.Engagements.GetByGuid(new Guid(r["engagementuid"].ToString()));
                }

                if (pEng == null)
                {
                    res = resources.GetByGuid(new Guid(res_uid));
                    projContext.Load(res);
                    projContext.ExecuteQuery();

                    ProjectEngagementCreationInformation peci = new ProjectEngagementCreationInformation();
                    peci.Id = Guid.NewGuid();
                    peci.Start = Convert.ToDateTime(r["start_dt"].ToString());
                    peci.Finish = Convert.ToDateTime(r["end_dt"].ToString());
                    peci.Resource = res;
                    peci.Work = "0h";
                    peci.Description = "RPM_" + r["planuid"].ToString();

                    draftProj.Engagements.Add(peci).Status = EngagementStatus.Proposed;
                    draftProj.Engagements.Update();
                    pEng = draftProj.Engagements.Last();

                    projContext.Load(pEng);
                    projContext.ExecuteQuery();
                }

                DataRow[] rows = ds.Tables[1].Select("resourceuid='" + res_uid + "' and projectuid='" + proj_uid + "'");
                ProjectEngagementTimephasedCollection petpc = pEng.GetTimephased(Convert.ToDateTime(r["start_dt"]), Convert.ToDateTime(r["end_dt"]), TimeScale.Days, EngagementContourType.Draft);

                projContext.Load(petpc);
                projContext.ExecuteQuery();

                foreach (DataRow row in rows)
                {
                    petpc.GetByStart(Convert.ToDateTime(row["timebyday"].ToString())).Work = row["allocationwork"].ToString();
                }
                pEng.Status = EngagementStatus.Reproposed; //this is needed
                draftProj.Engagements.Update();

                draftProj.CheckIn(false);
                //this updates the last TBD engagement in PWA, Approved will remove Eng from PWA
                QueueJob qJob1 = projContext.Projects.Update();
                JobState jobState = projContext.WaitForQueue(qJob1, timeoutSeconds);

                //{
                //    //approve proposed request
                //    projContext.Load(res.Engagements);
                //    projContext.ExecuteQuery();

                //    ResourceEngagement eng = res.Engagements.GetById(pEng.Id.ToString());
                //    projContext.Load(eng);
                //    projContext.ExecuteQuery();  //Too many resources: 4205. You cannot load dependent objects for more than 1000 resources. Use a filter to restrict your query

                //    eng.Status = EngagementStatus.Approved;
                //    res.Engagements.Update();

                //    QueueJob qJob = projContext.Projects.Update();
                //    jobState = projContext.WaitForQueue(qJob, timeoutSeconds);
                //}

            }

        }

        static void testEngagement_old(DataSet ds)
        {
            string res_uid = "";
            string proj_uid = "";

            foreach(DataRow r in ds.Tables[1].Rows)
            {
                res_uid = r["resourceuid"].ToString();
                proj_uid = r["projectuid"].ToString();

                pubProj = projContext.Projects.GetByGuid(new Guid(proj_uid));
                projContext.Load(pubProj);
                projContext.ExecuteQuery();

                draftProj = pubProj.CheckOut();

                projContext.Load(draftProj.Engagements);
                projContext.ExecuteQuery();

                ProjectEngagement pEng = null;
                foreach (ProjectEngagement pe in draftProj.Engagements)
                {
                    projContext.Load(pe);
                    projContext.Load(pe.Resource);
                    projContext.ExecuteQuery();
                    if (pe.Resource.Id.ToString().ToUpper() == res_uid.ToUpper())
                    {
                        pEng = pe;
                        res = pe.Resource;
                        break;
                    }
                }

                if (pEng == null)
                {
                    res = resources.GetByGuid(new Guid(res_uid));
                    projContext.Load(res);
                    projContext.ExecuteQuery();

                    ProjectEngagementCreationInformation peci = new ProjectEngagementCreationInformation();
                    peci.Id = Guid.NewGuid();
                    peci.Start = Convert.ToDateTime("1/1/2019");
                    peci.Finish = Convert.ToDateTime("1/1/2019");
                    peci.Resource = res;
                    peci.Work = "0h";
                    peci.Description = "RPM_" + "alloc_uid";

                    draftProj.Engagements.Add(peci).Status = EngagementStatus.Proposed;
                    draftProj.Engagements.Update();
                    pEng = draftProj.Engagements.Last();

                    projContext.Load(pEng);
                    projContext.Load(pEng.Resource);
                    projContext.ExecuteQuery();
                }


                DataRow[] rows = ds.Tables[2].Select("resourceuid='" + res_uid + "' and projectuid='" + proj_uid + "'");
                ProjectEngagementTimephasedCollection petpc = pEng.GetTimephased(Convert.ToDateTime(r["start_dt"]), Convert.ToDateTime(r["end_dt"]), TimeScale.Days, EngagementContourType.Draft);

                //works
                //ProjectEngagementTimephasedCollection petpc = pEng.GetTimephased(Convert.ToDateTime("10/1/2019"), Convert.ToDateTime("10/31/2019"), TimeScale.Months, EngagementContourType.Draft);

                //ProjectEngagementTimephasedCollection petpc = pEng.GetTimephased(Convert.ToDateTime("10/1/2019"), Convert.ToDateTime("12/31/2019"), TimeScale.Days, EngagementContourType.Draft);

                projContext.Load(petpc);
                projContext.ExecuteQuery();

                //petpc.GetByStart(Convert.ToDateTime("10/1/2019")).Work = "1";
                //petpc.GetByStart(Convert.ToDateTime("10/2/2019")).Work = "2";
                //petpc.GetByStart(Convert.ToDateTime("10/3/2019")).Work = "3";
                //petpc.GetByStart(Convert.ToDateTime("10/4/2019")).Work = "0";
                //petpc[3].Work = "4";
                //DateTime dt = petpc[2].Start;

                //petpc[0].Work = "100h";
                //petpc.GetByStart(Convert.ToDateTime("10/1/2019")).Work = "50h";
                //draftProj.Engagements.Update();
                ProjectEngagementTimephasedPeriod petpP = null;
                foreach (DataRow row in rows)
                {
                    petpc.GetByStart(Convert.ToDateTime(row["timebyday"].ToString())).Work = row["allocationwork"].ToString();

                    //string dt = row["timebyday"].ToString();
                    //petpP = petpc.GetByStart(Convert.ToDateTime(row["timebyday"]));
                    //projContext.Load(petpP);
                    //projContext.ExecuteQuery();
                    //petpP.Work = row["allocationwork"].ToString();

                    //ProjectEngagementTimephasedCollection petpc = pEng.GetTimephased(Convert.ToDateTime(row["timebyday"]), Convert.ToDateTime(row["timebyday"]), TimeScale.Days, EngagementContourType.Draft);
                    //projContext.Load(petpc);
                    //projContext.ExecuteQuery();
                    //petpc[0].Work = row["allocationwork"].ToString();
                    //draftProj.Engagements.Update();
                }

                //draftProj.Engagements.Update();
                pEng.Status = EngagementStatus.Reproposed; //this is needed
                draftProj.Engagements.Update();

                draftProj.CheckIn(false);
                //this updates the last TBD engagement in PWA, Approved will remove Eng from PWA
                QueueJob qJob1 = projContext.Projects.Update();
                JobState jobState = projContext.WaitForQueue(qJob1, timeoutSeconds);

                {
                    //approve proposed request
                    projContext.Load(res.Engagements);
                    projContext.ExecuteQuery();

                    ResourceEngagement eng = res.Engagements.GetById(pEng.Id.ToString());
                    projContext.Load(eng);
                    projContext.ExecuteQuery();  //Too many resources: 4205. You cannot load dependent objects for more than 1000 resources. Use a filter to restrict your query

                    eng.Status = EngagementStatus.Approved;
                    res.Engagements.Update();

                    QueueJob qJob = projContext.Projects.Update();
                    jobState = projContext.WaitForQueue(qJob, timeoutSeconds);
                }

            }

        }

        static void checkEngagement(string proj_uid, string res_uid = "4d6e3553-5ab1-e411-9a07-00155d509515", string alloc_uid = "zd6e3553-5ab1-e411-9a07-00155d509515")
        {
            res_uid = "8a8b380f-10ba-e411-ab5d-00155da4340f"; //bt
            pubProj = projContext.Projects.GetByGuid(new Guid(proj_uid));
            projContext.Load(pubProj);
            projContext.ExecuteQuery();

            draftProj = pubProj.CheckOut();

            projContext.Load(draftProj.Engagements);
            projContext.ExecuteQuery();

            ProjectEngagement pEng = null;
            foreach (ProjectEngagement pe in draftProj.Engagements)
            {
                projContext.Load(pe);
                projContext.Load(pe.Resource);
                projContext.ExecuteQuery();
                if(pe.Resource.Id.ToString() == res_uid)
                {
                    pEng = pe;
                    res = pe.Resource;
                    break;
                }
            }
            //ProjectEngagement pEng = draftProj.Engagements.First(e => e.Resource.Id.ToString() == res_uid );
            //projContext.Load(pEng.Resource);
            //projContext.ExecuteQuery();
            //res = pEng.Resource;

            if(pEng == null)
            {
                res = resources.GetByGuid(new Guid(res_uid));
                projContext.Load(res);
                projContext.ExecuteQuery();

                ProjectEngagementCreationInformation peci = new ProjectEngagementCreationInformation();
                peci.Id = Guid.NewGuid();
                peci.Start = Convert.ToDateTime("1/1/2019");
                peci.Finish = Convert.ToDateTime("1/1/2019");
                peci.Resource = res;
                peci.Work = "0h";
                peci.Description = "RPM_" + alloc_uid ;



                //projContext.Load(draftProj.Engagements);
                //projContext.Load(pubProj.Engagements);
                //projContext.ExecuteQuery();

                //if (calcfrom == "project")
                //    draftProj.UtilizationType = ProjectUtilizationType.ProjectPlan;
                //else
                //    draftProj.UtilizationType = ProjectUtilizationType.ResourceEngagements;

                draftProj.Engagements.Add(peci).Status = EngagementStatus.Proposed;
                draftProj.Engagements.Update();
                pEng = draftProj.Engagements.Last();

                projContext.Load(pEng);
                projContext.Load(pEng.Resource);
                projContext.ExecuteQuery();
            }

            ProjectEngagementTimephasedCollection petpc = pEng.GetTimephased(Convert.ToDateTime("5/1/2020"), Convert.ToDateTime("5/31/2020"), TimeScale.Months, EngagementContourType.Draft);
            projContext.Load(petpc);
            projContext.ExecuteQuery();
            petpc[0].Work = "20h";
            petpc[1].Work = "20h";
            petpc[2].Work = "30h";
            petpc[3].Work = "40h";
            petpc[4].Work = "50h";
            petpc[5].Work = "60h";
            petpc[6].Work = "70h";
            petpc[7].Work = "80h";

            pEng.Status = EngagementStatus.Reproposed;
            draftProj.Engagements.Update();

            draftProj.CheckIn(false);
            QueueJob qJob1 = projContext.Projects.Update();
            JobState jobState = projContext.WaitForQueue(qJob1, timeoutSeconds);


            {
                //approve proposed request
                projContext.Load(res.Engagements);
                projContext.ExecuteQuery();

                ResourceEngagement eng = res.Engagements.GetById(pEng.Id.ToString());
                projContext.Load(eng);
                projContext.ExecuteQuery();  //Too many resources: 4205. You cannot load dependent objects for more than 1000 resources. Use a filter to restrict your query

                eng.Status = EngagementStatus.Approved;
                res.Engagements.Update();

                QueueJob qJob = projContext.Projects.Update();
                jobState = projContext.WaitForQueue(qJob, timeoutSeconds);
            }

        }

        static void createPWAEngagement(string proj_uid, string res_uid, string dtfrom, string dtto, double pct, int isDraft, string calcfrom, string description)
        {
            int x = 0;
            Guid engagement = Guid.NewGuid();

            //EngagementDataSet eds = PJContext.Current.PSI.EngagementWebService.ReadEngagementsForProject(i_ProjectUID);


            if (m_proj_uid != proj_uid)
            {
                if (x != 0)
                {
                    draftProj.CheckIn(false);

                    QueueJob qJob = projContext.Projects.Update();
                    JobState jobState = projContext.WaitForQueue(qJob, timeoutSeconds);
                }

                pubProj = projContext.Projects.GetByGuid(new Guid(proj_uid));
                projContext.Load(pubProj);
                projContext.ExecuteQuery();
                draftProj = pubProj.CheckOut();

                m_proj_uid = proj_uid;
            }


            //setProj(proj_uid,res_uid);
            //res = projContext.EnterpriseResources.GetByGuid(new Guid(res_uid));
            res = resources.GetByGuid(new Guid(res_uid));
            projContext.Load(res);
            projContext.ExecuteQuery();

            ProjectEngagementCreationInformation peci = new ProjectEngagementCreationInformation();
            peci.Id = engagement;
            peci.Start = Convert.ToDateTime(dtfrom);
            peci.Finish = Convert.ToDateTime(dtto);
            peci.Resource = res;
            peci.MaxUnits = pct;
            //peci.Work = "8h"
            //peci.Description = description;



            //projContext.Load(draftProj.Engagements);
            //projContext.Load(pubProj.Engagements);
            //projContext.ExecuteQuery();

            //if (calcfrom == "project")
            //    draftProj.UtilizationType = ProjectUtilizationType.ProjectPlan;
            //else
            //    draftProj.UtilizationType = ProjectUtilizationType.ResourceEngagements;

            if (isDraft == 1)
                draftProj.Engagements.Add(peci).Status = EngagementStatus.Draft;
            else
                draftProj.Engagements.Add(peci).Status = EngagementStatus.Proposed;

            draftProj.Engagements.Update();


            //if (2 == 1)
            //{
            //    //Success: Retrieve all projects
            //    //var projects = projContext.LoadQuery(projContext.Projects);
            //    //projContext.ExecuteQuery();


            //    //approve proposed request
            //    //projContext.Load(res.Engagements.GetById(engagement.ToString()));
            //    //projContext.Load(projContext.EnterpriseResources.GetByGuid(res.Id).Engagements);
            //    projContext.Load(res.Engagements);
            //    projContext.ExecuteQuery();


            //    ResourceEngagement eng = res.Engagements.GetById(engagement.ToString());
            //    //projContext.Load(eng);
            //    //projContext.ExecuteQuery();  //Too many resources: 4205. You cannot load dependent objects for more than 1000 resources. Use a filter to restrict your query

            //    eng.Status = EngagementStatus.Approved;
            //    res.Engagements.Update();

            //    qJob = projContext.Projects.Update();
            //    jobState = projContext.WaitForQueue(qJob, timeoutSeconds);
            //}
            x = x + 1;
        }

        private Guid CreateNewEngagement(Guid i_ResUID, Guid i_ProjUID, string i_AllocationPlanName, Guid i_AllocationPlanUID, DateTime i_StartDate, 
            DateTime i_FinishDate, decimal i_WorkHours)
        {
            Guid res = Guid.NewGuid();

            try
            {
                Eng newEng = new Eng();
                newEng.Res = i_ResUID;
                newEng.Proj = i_ProjUID;
                newEng.Name = string.Format("{0} ( {1} )", i_AllocationPlanName, i_AllocationPlanUID.ToString());
                newEng.Start = i_StartDate;
                newEng.Finish = i_FinishDate;
                newEng.Work = Math.Round(i_WorkHours, 2) * 60000;
                //newEng.Comment = i_Comment;
                string json = newEng.ToJson();
                //res = PJContext.Current.PSI.PWAWebService.EngagementCreateApprovedEngagement(json);
            }
            catch (Exception ex)
            {
                //LogManager.WriteEntry("Error occured in CreateNewEngagement (EngagementsSynchManager.cs) mmethod. Details: " + ex.ToString(), EventLogEntryType.Error);
                res = Guid.Empty;
            }

            return res;
        }

        private void UpdateExistingEngagement(Guid i_EngagementUID, Guid i_ProjectUID, DateTime i_From, DateTime i_To, decimal i_Hours)
        {
            try
            {
                i_To = i_To.AddDays(1);

                //List<EngUpdate> updates = new List<EngUpdate>();
                //EngUpdate update = new EngUpdate(i_EngagementUID, i_From, i_To, i_Hours, ENGAGEMENT_DATE_FORMAT);
                //updates.Add(update);

                //List<EngProperty> props = new List<EngProperty>();
                //EngProperty prop = new EngProperty();
                //prop.Key = i_EngagementUID;
                //prop.Value = new EngPropertyValue(i_ProjectUID, (int)EngagementStatus.Approved);
                //props.Add(prop);

                //JavaScriptSerializer serializer = new JavaScriptSerializer();

                //PJContext.Current.PSI.PWAWebService.EngagementSendResourceRequestsUpdate("", serializer.Serialize(updates), serializer.Serialize(props), 0, "Date(" + i_From.Ticks.ToString() + ")",
                //    "Date(" + i_To.Ticks.ToString() + ")", 5, true, true, "");

            }
            catch (Exception ex)
            {
                //LogManager.WriteEntry("Error occured in UpdateExistingEngagement (EngagementsSynchManager.cs) method. Details: " + ex.ToString(), EventLogEntryType.Error);
            }
        }

    }


    public class EngagementPeriod
    {
        public Guid? EngagementUID { get; set; }
        public Guid ProjectUID { get; set; }
        public Guid ResourceUID { get; set; }
        public Guid AllocationPlanUID { get; set; }
        public DateTime StartDate { get; set; }
        public DateTime FinishDate { get; set; }
        public decimal AllocationHours { get; set; }

        public EngagementPeriod(Guid i_ProjUID, Guid i_ResUID, Guid i_AllocPLanUID)
        {
            ProjectUID = i_ProjUID;
            ResourceUID = i_ResUID;
            AllocationPlanUID = i_AllocPLanUID;
        }

        public EngagementPeriod(Guid? i_EngagementUID, Guid i_ProjUID, Guid i_ResUID, Guid i_AllocPLanUID, DateTime i_StartDate, DateTime i_FinishDate, decimal i_AllocationHours)
        {
            EngagementUID = i_EngagementUID;
            ProjectUID = i_ProjUID;
            ResourceUID = i_ResUID;
            AllocationPlanUID = i_AllocPLanUID;
            StartDate = i_StartDate;
            FinishDate = i_FinishDate;
            AllocationHours = i_AllocationHours;
        }

    }



    public class Eng
    {
        public Guid Res { get; set; }
        public Guid Proj { get; set; }
        public string Name { get; set; }
        public DateTime Start { get; set; }
        public DateTime Finish { get; set; }
        public decimal Work { get; set; }
        public string Comment { get; set; }

        internal string ToJson()
        {
            JavaScriptSerializer serializer = new JavaScriptSerializer();
            return serializer.Serialize(this);
        }
    }

    public class Period
    {
        public DateTime Start { get; set; }
        public DateTime Finish { get; set; }

        public Period(DateTime i_Start, DateTime i_Finish)
        {
            this.Start = i_Start;
            this.Finish = i_Finish;
        }
    }

    public class EngProperty
    {
        public Guid Key { get; set; }
        public EngPropertyValue Value { get; set; }
    }

    public class EngPropertyValue
    {
        public Guid ProjectUid { get; set; }
        public int Status { get; set; }

        public EngPropertyValue(Guid i_ProjectUid, int i_Status)
        {
            this.ProjectUid = i_ProjectUid;
            this.Status = i_Status;
        }
    }

    public class EngUpdate
    {
        public EngUpdateSection[] updates { get; set; }
        public int changeNumber { get; set; }

        public EngUpdate(Guid i_EngagementUID, DateTime i_Start, DateTime i_Finish, decimal i_Hours, string i_EngagementDateFormat)
        {
            List<EngUpdateSection> updatesList = new List<EngUpdateSection>();

            string fieldKey = string.Format("TPD_COMM_{0}#{1}", i_Start.ToString(i_EngagementDateFormat), i_Finish.ToString(i_EngagementDateFormat));

            updatesList.Add(new EngUpdateSection(2, i_EngagementUID, fieldKey, (int)(decimal.Round(i_Hours, 2) * 60000), true));

            updates = updatesList.ToArray();
        }
    }

    public class EngUpdateSection
    {
        public int type { get; set; }
        public Guid recordKey { get; set; }
        public string fieldKey { get; set; }
        public EngNewProp newProp { get; set; }
        public EngUpdateSection(int i_type, Guid i_recordKey, string i_fieldKey, int i_dataValue, bool i_hasDataValue)
        {
            this.type = i_type;
            this.recordKey = i_recordKey;
            this.fieldKey = i_fieldKey;
            this.newProp = new EngNewProp(i_dataValue.ToString(), i_hasDataValue);
        }

    }

    public class EngNewProp
    {
        public string dataValue { get; set; }
        public bool hasDataValue { get; set; }

        public EngNewProp(string i_dataValue, bool i_hasDataValue)
        {
            dataValue = i_dataValue;
            hasDataValue = i_hasDataValue;
        }
    }


    public enum eApprovalODATASource
    {
        Project,
        Resource
    }

    public enum eEngagementInitialMode
    {
        Commited,
        CommitedByWorkflow
    }

    public class EngagementUpdateResult
    {
        public Guid AllocationPlanUID { get; set; }
        public Guid ProjectUID { get; set; }
        public Guid ResourceUID { get; set; }
        public bool UpdateSucceded { get; set; }

        public EngagementUpdateResult(Guid i_AllocationPlanUID, Guid i_ProjectUID, Guid i_ResourceUID, bool i_UpdateSucceded)
        {
            this.AllocationPlanUID = i_AllocationPlanUID;
            this.ProjectUID = i_ProjectUID;
            this.ResourceUID = i_ResourceUID;
            this.UpdateSucceded = i_UpdateSucceded;
        }
    }

    public class EngagementComment
    {
        public Guid AllocationPlanUID { get; set; }
        public Guid ProjectUID { get; set; }
        public Guid ResourceUID { get; set; }
        public string CommentMessage { get; set; }
        public string CommentedBy { get; set; }
        public DateTime CommentedAt { get; set; }

    }

}
