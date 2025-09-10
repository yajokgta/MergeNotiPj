using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.Entity.Core.Objects;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using WolfApprove.API2.Controllers.API;
using WolfApprove.API2.Controllers.Services;
using WolfApprove.API2.Extension;
using WolfApprove.Model;
using WolfApprove.Model.CustomClass;
using WolfApprove.Model.Extension;
using static WolfApprove.API2.Controllers.Services.TaskSchedulerService;

namespace MergeNotiPj
{
    class Program
    {
        public static string _LogFile = ConfigurationSettings.AppSettings["LogFile"];
        public static string _SMTPServer = ConfigurationSettings.AppSettings["SMTPServer"];
        public static string _SMTPPort = ConfigurationSettings.AppSettings["SMTPPort"];
        public static string _SMTPEnableSSL = ConfigurationSettings.AppSettings["SMTPEnableSSL"];
        public static string _SMTPUser = ConfigurationSettings.AppSettings["SMTPUser"];
        public static string _SMTPPassword = ConfigurationSettings.AppSettings["SMTPPassword"];
        public static string _SMTPTestMode = ConfigurationSettings.AppSettings["SMTPTestMode"];
        public static string _SMTPTo = ConfigurationSettings.AppSettings["SMTPTo"];
        public static string _URLWeb = ConfigurationSettings.AppSettings["URLWeb"];
        public static string _MailToSupport = ConfigurationSettings.AppSettings["MailToSupport"];

        public static void Log(String iText, string module = "")
        {
            string pathlog = _LogFile;
            String logFolderPath = System.IO.Path.Combine(pathlog, DateTime.Now.ToString("yyyyMMdd") + (string.IsNullOrEmpty(module) ? string.Empty : $"_{module}"));

            if (!System.IO.Directory.Exists(logFolderPath))
            {
                System.IO.Directory.CreateDirectory(logFolderPath);
            }
            String logFilePath = System.IO.Path.Combine(logFolderPath, DateTime.Now.ToString("yyyyMMdd") + ".txt");

            try
            {
                using (System.IO.StreamWriter outfile = new System.IO.StreamWriter(logFilePath, true))
                {
                    System.Text.StringBuilder sbLog = new System.Text.StringBuilder();

                    String[] listText = iText.Split('|').ToArray();

                    foreach (String s in listText)
                    {
                        sbLog.AppendLine($"[{DateTime.Now:HH:mm:ss}] {s}");
                    }

                    outfile.WriteLine(sbLog.ToString());
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error writing log file: {ex.Message}");
            }
        }
        static void Main()
        {
            try
            {
                Log("====== Start Process ====== : " + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"));
                Log(string.Format("Run batch as :{0}", System.Security.Principal.WindowsIdentity.GetCurrent().Name));
                var respone = GetSettingNoti(config.dbConnectionString);
                var json = JsonConvert.SerializeObject(respone);
                var responeModel = JsonConvert.DeserializeObject<ResponeModel>(json);
                if (responeModel != null)
                {
                    for (int i = 0; i < responeModel?.settings?.Count; i++)
                    {
                        var setting = responeModel.settings[i];
                        var guid = Guid.NewGuid().ToString();
                        CreateMasterRelateJobSetting(config.dbConnectionString, guid, i, responeModel.MemoId, "J_NOTI");
                        var condition = setting[1]?.value;

                        DateTime today = DateTime.Now;
                        bool isMonday = today.DayOfWeek == DayOfWeek.Monday;
                        bool isWednesday = today.DayOfWeek == DayOfWeek.Wednesday;
                        bool isFriday = today.DayOfWeek == DayOfWeek.Friday;
                        bool isFirstOfMonth = today.Day == 1;

                        if ((isMonday || isWednesday || isFriday) && condition == "ทุก 2 วัน")
                        {
                            RunNotificationService("J_NOTI_2DAYS", config.dbConnectionString, guid, responeModel.MemoId, "J_NOTI");
                        }
                        else if (isMonday && condition == "ทุกวันจันทร์")
                        {
                            RunNotificationService("J_NOTI_MONDAY", config.dbConnectionString, guid, responeModel.MemoId, "J_NOTI");
                        }
                        else if (isFirstOfMonth && condition == "ทุกวันที่ 1 ของเดือน")
                        {
                            RunNotificationService("J_NOTI_FIRSTMONTH", config.dbConnectionString, guid, responeModel.MemoId, "J_NOTI");
                        }
                        else //ทุกวัน
                        {
                            RunNotificationService("J_NOTI_EVERYDAY", config.dbConnectionString, guid, responeModel.MemoId, "J_NOTI");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR");
                Console.WriteLine("Exit ERROR");
                Log("message: " + ex, "ERROR");
            }
            finally
            {
                Log("====== End Process Process ====== : " + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"));
            }
        }
        static void RunNotificationService(string jobType, string connectionString, string guid, int memoId, string masterType)
        {
            Console.WriteLine($"[ {DateTime.Now} ] เรียกใช้ NotificationService: {jobType}");
            Log($"[ {DateTime.Now} ] เรียกใช้ NotificationService: {jobType}");
            Log($"[ {DateTime.Now} ] เรียกใช้ connectionString: {connectionString}");
            NotificationService(connectionString, guid, memoId, masterType);
        }
        public static string NotificationService(string connectionString, string jobId, int memoId, string masterType)
        {
            var jobs = "";
            try
            {
                WolfApproveModel dbContext = DBContext.OpenConnection(connectionString);

                string result = "";
                string stringmemoId = memoId.ToString();
                var masterSetting = dbContext.MSTMasterDatas.FirstOrDefault(x =>
                    x.MasterType == masterType &&
                    x.Value1 == jobId &&
                    x.Value3 == stringmemoId);

                var memoSetting = dbContext.TRNMemoes.FirstOrDefault(x => x.MemoId == memoId);

                if (masterSetting != null && memoSetting != null)
                {
                    List<AdvanceFormExt.AdvanceForm> advanceFormModel = AdvanceFormExt.ToList(memoSetting.MAdvancveForm);
                    AdvanceFormExt.AdvanceForm settingTable = advanceFormModel.FirstOrDefault(x => x.type == "tb");
                    int rowIndexSetting = Convert.ToInt32(masterSetting.Value2);
                    List<AdvanceFormExt.AdvanceFormRow> settings = settingTable?.row?[rowIndexSetting];

                    var tempCodeRow = settings?.FirstOrDefault(x => x.label == "TemplateCode");
                    string tempCode = tempCodeRow?.value;

                    var noticeToRow = settings?.FirstOrDefault(x => x.label == "Notice To");
                    string noticeTo = noticeToRow?.value;

                    var formStateRow = settings?.FirstOrDefault(x => x.label == "FormState[MSTEmailTemplate]");
                    string formState = formStateRow?.value;

                    var noticeCCRow = settings?.FirstOrDefault(x => x.label == "Notice CC");
                    string noticeCC = noticeCCRow?.value;

                    var jobsRow = settings?.FirstOrDefault(x => x.label == "รายการ");
                    jobs = jobsRow?.value;

                    var optionRow = settings?.FirstOrDefault(x => x.label == "Option");
                    string option = optionRow?.value;

                    if (jobs != "--select--")
                    {
                        Log($"Task Name : {jobs}");
                        switch (jobs)
                        {
                            case "แจ้งเตือน IAC ก่อนถึงกำหนดการ Internal Audit จริงตามแผน":
                                NotificationPlan_IAC(dbContext, connectionString, formState, memoId, noticeCC);
                                Log($"Process Name : NotificationPlan_IAC");
                                break;
                            case "สร้าง report แจ้งเตือน CAR Inprocess ":
                                NotificationCAR_InProcess(dbContext, connectionString, formState, memoId, noticeCC);
                                Log($"Process Name : NotificationCAR_InProcess");
                                break;
                            case "สร้าง report แจ้งเตือน ARS Inprocess":
                                NotificationARS_InProcess(dbContext, connectionString, formState, memoId, noticeCC);
                                Log($"Process Name : NotificationARS_InProcess");
                                break;
                            case "แจ้งเตือนการดำเนินการ CAR & ARS ก่อนถึงกำหนดแผนการปรับปรุง ":
                                Notification_CAR_DealDate(dbContext, connectionString, formState, memoId, noticeCC);
                                break;
                            case "สร้าง report  แจ้งเตือนเอกสาร Inprocess  ( ทุกประเภทเอกสาร) ":
                                NotificationDAR_InProcess(dbContext, connectionString, formState, memoId, noticeCC);
                                Log($"Process Name : NotificationDAR_InProcess");
                                break;
                            case "สร้าง report  แจ้งเตือนบันทึก (e-Form) Inprocess ( ทุกรายการ)  ":
                                NotificationEForm_InProcess(dbContext, connectionString, formState, memoId, noticeCC);
                                Log($"Process Name : NotificationEForm_InProcess");
                                break;
                            case "สร้าง report แจ้งเตือน \"DAR-COPY: ขอสำเนาเอกสารฯ\" Inprocess (ทุกรายการ)":
                                Notification_DAR_CopyReport(dbContext, connectionString, formState, memoId, noticeCC);
                                Log($"Process Name : Notification_DAR_CopyReport");
                                break;
                            case "สร้าง report แจ้งเตือน IAS Inprocess (Part Auditee :  Accepted plan)":
                                NotificationIAS(dbContext, connectionString, formState, memoId, noticeCC);
                                Log($"Process Name : NotificationIAS");
                                break;
                            case "แจ้งเตือน CAR & ARS ที่ผ่านการอนุมัติ และประกาศใช้ผ่านระบบ ระบบส่งข้อความแจ้งให้ User รับทราบอัตโนมัติ ":
                                Notification_CAR_ARS(dbContext, connectionString, formState, memoId, noticeCC);
                                Log($"Process Name : Notification_CAR_ARS");
                                break;
                            case "สร้าง report  เอกสาร PR/WI (ISO 9001) ที่เกี่ยวข้องกับหน่วยงาน / Related documents":
                                Report_PR_WI(dbContext, connectionString, formState, memoId, noticeCC);
                                Log($"Process Name : Report_PR_WI");
                                break;
                            case "สร้าง report  แจ้งเตือน IAC Inprocess (Part Auditee :  Accepted resulted)":
                                NotificationIAC(dbContext, connectionString, formState, memoId, noticeCC);
                                Log($"Process Name : NotificationIAC");
                                break;
                            case "แจ้งเตือนการ update เอกสารในรอบปีที่ต้องดำเนินการในระบบ":
                                NotificationDAR_N(dbContext, connectionString, formState, memoId, noticeCC, option);
                                break;
                            case "เมื่อมีการแก้ไขชื่อหน่วยงานใน Setting ระบบสร้าง report แจ้งรายการที่มีการแก้ไขชื่อหน่วยงานใน Code Area รับทราบ":
                                NotificationMasterArea(dbContext, connectionString, formState, memoId, noticeCC);
                                Log($"Process Name : NotificationMasterArea");
                                break;
                            case "แจ้งเตือน เอกสาร(ทุกประเภท) ที่ผ่านการอนุมัติ และประกาศใช้ผ่านระบบ ระบบส่งข้อความแจ้งให้ User รับทราบอัตโนมัติ ":
                                Notification_ISO(dbContext, connectionString, formState, memoId, noticeCC);
                                Log($"Process Name : Notification_ISO");
                                break;
                            case "แจ้งเตือน \"DAR-COPY: ขอสำเนาเอกสารฯ\" ฉบับที่พ้นระยะเวลาขออนุมัติทำสำเนา ระบบส่งข้อความแจ้งให้ User รับทราบอัตโนมัติ ":
                                Notification_DAR_Copy_Expired(dbContext, connectionString, formState, memoId, noticeCC);
                                Log($"Process Name : Notification_DAR_Copy_Expired");
                                break;
                            case "สร้าง report แจ้งรายการเอกสาร/ข้อมูล ที่มีการขึ้นทะเบียน/แก้ไข /ยกเลิก ใน View ข่าวสาร (View: Annoucement) ของวันที่ผ่านมา":
                                Report_Announcement(dbContext, connectionString, formState, memoId, noticeCC);
                                break;
                            case "แจ้งเตือน IAS & IAC ที่ผ่านการอนุมัติ และประกาศใช้ผ่านระบบ ระบบส่งข้อความแจ้งให้ User รับทราบอัตโนมัติ ":
                                Notification_IAS_IAC_Completed(dbContext, connectionString, formState, memoId, noticeCC);
                                Log($"Process Name : Notification_IAS_IAC_Completed");
                                break;
                            case "สร้าง report แจ้งรายการแบบฟอร์ม(FR) ที่มีการขึ้นทะเบียน/แก้ไข /ยกเลิก (ของวันที่ผ่านมา)":
                                NotificationDAR_DocType_FR(dbContext, connectionString, formState, memoId, noticeCC);
                                Log($"Process Name : NotificationDAR_DocType_FR");
                                break;
                            case "แจ้งเตือน \"DAR-COPY: ขอสำเนาเอกสารฯ\" ที่ผ่านการอนุมัติ และประกาศใช้ผ่านระบบ  ระบบส่งข้อความแจ้งให้ User รับทราบอัตโนมัติ ":
                                Notification_DAR_Copy(dbContext, connectionString, formState, memoId, noticeCC);
                                Log($"Process Name : Notification_DAR_Copy");
                                break;
                            case "สร้าง report  แจ้งประวัติการแก้ไขเอกสาร ประจำเดือน (ทุกประเภทเอกสารที่มีการดำเนินการและประกาศใช้ทั้งเดือน)":
                                Report_DAR_E(dbContext, connectionString, formState, memoId, noticeCC);
                                Log($"Process Name : Report_DAR_E");
                                break;
                            default: Log("Not Found Case"); break;
                        }
                        return result;
                    }

                    var tempModel = dbContext.MSTTemplates
                        .FirstOrDefault(x => x.DocumentCode == tempCode);

                    if (tempModel != null)
                    {
                        var memoNoti = dbContext.TRNMemoes
                            .Where(x => x.TemplateId == tempModel.TemplateId)
                            .ToList();
                        foreach (var memo in memoNoti)
                        {
                            var emailStateModel = GetEmailTemplateByFormState(connectionString, formState);
                            MSTEmployee empModel = null;

                            switch (noticeTo)
                            {
                                case "PersonnalWatting":
                                    empModel = dbContext.MSTEmployees
                                        .FirstOrDefault(x => x.EmployeeId == memo.PersonWaitingId);
                                    break;
                                case "Requester":
                                    empModel = dbContext.MSTEmployees
                                        .FirstOrDefault(x => x.EmployeeId == memo.RequesterId);
                                    break;
                                case "Creater":
                                    empModel = dbContext.MSTEmployees
                                        .FirstOrDefault(x => x.EmployeeId == memo.CreatorId);
                                    break;
                            }

                            string dear = empModel?.NameTh;
                            string to = empModel?.Email;

                            emailStateModel = SerializeFormEmail(emailStateModel, connectionString, memoId, memo, dear);
                            SendEmailTemplate(emailStateModel, to, noticeCC);
                        }
                    }
                }
            }
            catch(Exception ex)
            {
                Log($"Error in {jobs}: {ex}", "ERROR");
            }
            

            return "";
        }
        public static object GetSettingNoti(string connectionString)
        {
            using (WolfApproveModel dbContext = DBContext.OpenConnection(connectionString))
            {
                var template = dbContext.MSTTemplates
                                        .FirstOrDefault(x => x.DocumentCode == "SettingJob" && x.IsActive == true);

                if (template == null)
                {
                    return new { };
                }

                var memoSetting = dbContext.TRNMemoes
                                           .Where(x => x.TemplateId == template.TemplateId)
                                           .OrderByDescending(o => o.ModifiedDate)
                                           .FirstOrDefault();

                if (memoSetting == null)
                {
                    return new { };
                }

                var advanceFormModel = AdvanceFormExt.ToList(memoSetting.MAdvancveForm);
                var settingTable = advanceFormModel.FirstOrDefault(x => x.type == "tb");

                if (settingTable == null)
                {
                    return new { };
                }

                var settings = settingTable.row ?? new List<List<AdvanceFormExt.AdvanceFormRow>>();

                return new
                {
                    memoSetting.MemoId,
                    settings
                };
            }
        }
        public static string NotificationPlan_IAC(WolfApproveModel dbContext, string connectionString, string formState, int memoId, string noticeCC)
        {
            var templateModel = dbContext.MSTTemplates.Where(x => x.DocumentCode == "IAC").Select(s => s.TemplateId).ToList();
            if (templateModel.Any())
            {
                var memos = dbContext.TRNMemoes.Where(x => templateModel.Contains(x.TemplateId) && x.StatusName != Ext.Status._Draft).ToList();
                foreach (TRNMemo memo in memos)
                {
                    List<AdvanceFormExt.AdvanceForm> advance = AdvanceFormExt.ToList(memo.MAdvancveForm);
                    string auditDateString = dbContext.TRNMemoForms.FirstOrDefault(x => x.MemoId == memo.MemoId && x.obj_label.Contains("วันที่เริ่มต้น"))?.obj_value;
                    DateTime auditDate = (DateTimeHelper.ConvertStringToDateTime(auditDateString) ?? DateTime.MaxValue).AddDays(-1);

                    DateTime dateNow = DateTime.Now;
                    var Auditor = advance.FirstOrDefault(x => x.label.Contains("หัวหน้าผู้ตรวจ"))?.value;
                    var Auditee = advance.FirstOrDefault(x => x.label.Contains("ผู้รับผิดชอบ"))?.value;

                    var membergroup = new List<string>();

                    membergroup.AddRange(dbContext.TRNMemoForms.Where(x => x.MemoId == memo.MemoId && x.obj_type == "tb" && x.obj_label.Contains("สมาชิกทีมผู้ตรวจติดตาม") && x.col_label.Contains("รายชื่อสมาชิกผู้ตรวจติดตาม")).Select(ts => ts.col_value).ToList());
                    membergroup.AddRange(dbContext.TRNMemoForms.Where(x => x.MemoId == memo.MemoId && x.obj_type == "tb" && x.obj_label.Contains("สมาชิกคณะทำงานฯ ที่ถูกตรวจ") && x.col_label.Contains("Name")).Select(ts => ts.col_value).ToList());
                    membergroup.AddRange(dbContext.TRNMemoForms.Where(x => x.MemoId == memo.MemoId && x.obj_type == "tb" && x.obj_label.Contains("DCC พื้นที่") && x.col_label.Contains("ชื่อสมาชิกคณะทำงานฯ")).Select(ts => ts.col_value).ToList());

                    Log($"6 {string.Join(",", membergroup)}");
                    var groupEmail = GetViewEmployeeQuery(dbContext)
                        .Where(x => membergroup.Contains(x.NameEn) || membergroup.Contains(x.NameTh))
                        .Select(s => s.Email)
                        .ToList();

                    var empAuditor = GetViewEmployeeQuery(dbContext).FirstOrDefault(x => x.NameEn == Auditor || x.NameTh == Auditor);
                    var empAuditee = GetViewEmployeeQuery(dbContext).FirstOrDefault(x => x.NameEn == Auditee || x.NameTh == Auditee);
                    Log($"{TruncateTime(dateNow):dd MMM yyyy HH:mm:ss} - {TruncateTime(auditDate):dd MMM yyyy HH:mm:ss} : MemoId : {memo.MemoId}");
                    if (TruncateTime(dateNow) == TruncateTime(auditDate))
                    {
                        //var listEmail = GetEmailInArea(auditISOArea, dbContext);
                        groupEmail.Add(empAuditee?.Email ?? "");
                        groupEmail.Add(empAuditor?.Email ?? "");
                        string dear = "";
                        string to = string.Join(";", groupEmail);
                        CustomEmailTemplate emailStateModel = GetEmailTemplateByFormState(connectionString, formState);
                        Log($"Send {to} : MemoId : {memo.MemoId}");
                        emailStateModel = SerializeFormEmail(emailStateModel, connectionString, memo.MemoId, memo, dear);
                        SendEmailTemplate(emailStateModel, to, noticeCC);
                    }
                }
            }
            return "";
        }
        public static string NotificationCAR_InProcess(WolfApproveModel dbContext, string connectionString, string formState, int memoId, string noticeCC)
        {
            List<int?> templateModel = dbContext.MSTTemplates
                .Where(x => x.DocumentCode == "CAR")
                .Select(s => s.TemplateId)
                .ToList();
            if (templateModel.Any())
            {
                var memos = dbContext.TRNMemoes
                    .Where(m => templateModel.Contains(m.TemplateId) && !_NotInProcess.Contains(m.StatusName))
                    .Select(m => new
                    {
                        m,
                        lapp = dbContext.TRNLineApproves.FirstOrDefault(l => m.MemoId == l.MemoId && m.CurrentApprovalLevel == l.Seq)
                    })
                    .ToList();

                var empIds = memos.Where(x => x.lapp != null).Select(s => s.lapp).Select(s => s.EmployeeId).ToList();

                var emails = GetViewEmployeeQuery(dbContext).Where(memo => empIds.Contains(memo.EmployeeId)).Select(s => s.Email).Distinct().ToList();

                var memoSerialize = memos.Select(s => s.m).Select(s =>
                {
                    var link = _URLWeb+ s.MemoId;
                    return new
                    {
                        s.DocumentNo,
                        ISO_area_Code = dbContext.TRNMemoForms.FirstOrDefault(x => x.MemoId == s.MemoId && x.obj_label == "รหัสพื้นที ISO ที่เกี่ยวข้อง")?.obj_value,
                        ARS_No_Dot = dbContext.TRNMemoForms.FirstOrDefault(x => x.MemoId == s.MemoId && x.obj_label == "เลขที่คำร้อง")?.obj_value,
                        s.MemoSubject,
                        s.StatusName,
                        s.PersonWaiting,
                        Link = link
                    };
                }).ToList();

                var memosHtml = DataTableToHtml(ToDataTable(memoSerialize));

                string to = string.Join(";", emails);
                CustomEmailTemplate emailStateModel = GetEmailTemplateByFormState(connectionString, formState);
                emailStateModel.EmailBody = ReplaceMailBody(emailStateModel.EmailBody, memosHtml);
                SendEmailTemplate(emailStateModel, to, noticeCC);
            }
            return "";
        }
        public static string NotificationARS_InProcess(WolfApproveModel dbContext, string connectionString, string formState, int memoId, string noticeCC)
        {
            List<int?> templateModel = dbContext.MSTTemplates
                .Where(x => x.DocumentCode == "ARS")
                .Select(s => s.TemplateId)
                .ToList();

            if (templateModel.Any())
            {
                var memos = dbContext.TRNMemoes
                    .Where(m => templateModel.Contains(m.TemplateId) && !_NotInProcess.Contains(m.StatusName))
                    .Select(m => new
                    {
                        m,
                        lapp = dbContext.TRNLineApproves.FirstOrDefault(l => m.MemoId == l.MemoId && m.CurrentApprovalLevel == l.Seq)
                    })
                    .ToList();

                var empIds = memos.Where(x => x.lapp != null).Select(s => s.lapp).Select(s => s.EmployeeId).ToList();

                var emails = GetViewEmployeeQuery(dbContext).Where(memo => empIds.Contains(memo.EmployeeId)).Select(s => s.Email).Distinct().ToList();

                var memoSerialize = memos.Select(s => s.m).Select(s => new
                {
                    s.DocumentNo,
                    ISO_area_Code = dbContext.TRNMemoForms.FirstOrDefault(x => x.MemoId == s.MemoId && x.obj_label == "รหัสพื้นที ISO ที่เกี่ยวข้อง")?.obj_value,
                    ARS_NoDot = dbContext.TRNMemoForms.FirstOrDefault(x => x.MemoId == s.MemoId && x.obj_label == "เลขที่คำร้อง")?.obj_value,
                    s.MemoSubject,
                    s.StatusName,
                    s.PersonWaiting,
                    Link = _URLWeb+ s.MemoId
                }).ToList();

                var memosHtml = DataTableToHtml(ToDataTable(memoSerialize));

                string to = string.Join(";", emails);
                CustomEmailTemplate emailStateModel = GetEmailTemplateByFormState(connectionString, formState);
                emailStateModel.EmailBody = ReplaceMailBody(emailStateModel.EmailBody, memosHtml);
                SendEmailTemplate(emailStateModel, to, noticeCC);
            }
            return "";
        }
        public static string Notification_CAR_DealDate(WolfApproveModel dbContext, string connectionString, string formState, int memoId, string noticeCC)
        {
            var templateModel = dbContext.MSTTemplates
                .Where(x => x.DocumentCode == "CAR" || x.DocumentCode == "ARS")
                .Select(x => x.TemplateId)
                .ToList();

            if (templateModel.Any())
            {
                var memos = dbContext.TRNMemoes
                    .Where(x => templateModel.Contains(x.TemplateId) && x.StatusName != Ext.Status._Draft && x.StatusName != Ext.Status._Rejected && x.StatusName != Ext.Status._Cancelled)
                    .ToList();

                foreach (var memo in memos)
                {
                    Log("memo : " + memo.MemoId);
                    var memoForms = dbContext.TRNMemoForms.Where(x => x.MemoId == memo.MemoId).ToList();

                    var advanceForms = AdvanceFormExt.ToList(memo.MAdvancveForm);
                    var tables = memoForms.Where(x => (x.obj_label.Contains("แนบแผนการปฎิบัติการแก้ไขปัญหาเบื้องต้น") || x.obj_label.Contains("*กรณีที่ไม่สามารถดำเนินการปรับปรุงได้ทันที")) && x.col_label.Contains("วันที่ดำเนินการแล้วเสร็จ")).ToList();

                    var empModel = dbContext.MSTEmployees
                        .FirstOrDefault(x => x.EmployeeId == memo.CreatorId);

                    var emailCC = new List<string>();
                    if (tables != null)
                    {
                        foreach (var row in tables)
                        {
                            var dewDateString = row.col_value;

                            if (!string.IsNullOrEmpty(dewDateString))
                            {
                                var dewDate = (DateTimeHelper.ConvertStringToDateTime(dewDateString) ?? DateTime.MinValue);
                                Log("dewDate : " + dewDateString);
                                if (TruncateTime(DateTime.Now.AddDays(3)) == TruncateTime(dewDate))
                                {
                                    emailCC.Add(empModel?.Email);
                                    var Requestor = memoForms.FirstOrDefault(x => x.obj_label.Contains("ผู้เปิด"))?.obj_value ?? "";
                                    var auditee = memoForms.FirstOrDefault(x => x.obj_label.Contains("ผู้รับผิดชอบ"))?.obj_value ?? "";
                                    var auditeeEmployee = GetViewEmployeeQuery(dbContext).FirstOrDefault(x => x.NameEn.Contains(auditee) || x.NameTh.Contains(auditee))?.Email;
                                    var RequestorEmployee = GetViewEmployeeQuery(dbContext).FirstOrDefault(x => x.NameEn.Contains(Requestor) || x.NameTh.Contains(Requestor))?.Email;

                                    emailCC.Add(auditeeEmployee);
                                    emailCC.Add(RequestorEmployee);

                                    string dear = empModel?.NameTh;
                                    string to = string.Join(";", emailCC);
                                    CustomEmailTemplate emailStateModel = GetEmailTemplateByFormState(connectionString, formState);
                                    emailStateModel = SerializeFormEmail(emailStateModel, connectionString, memo.MemoId, memo, dear, dewDateString);
                                    SendEmailTemplate(emailStateModel, to, noticeCC);
                                }
                            }
                        }
                    }
                }
            }

            return "";
        }
        public static string NotificationDAR_InProcess(WolfApproveModel dbContext, string connectionString, string formState, int memoId, string noticeCC)
        {
            try
            {

                var templateCodes = new List<string>
            {
                "DAR-NEW",
                "DAR-EDIT",
                "DAR-CANCEL",
            };
               
                var templateModel = dbContext.MSTTemplates
                    .Where(x => templateCodes.Contains(x.DocumentCode))
                    .Select(s => s.TemplateId)
                    .ToList();
               
                if (templateModel.Any())
                {
                    var memos = dbContext.TRNMemoes
                         .Where(m => templateModel.Contains(m.TemplateId) && !_NotInProcess.Contains(m.StatusName))
                         .Select(m => new
                         {
                             m,
                             lapp = dbContext.TRNLineApproves.FirstOrDefault(l => m.MemoId == l.MemoId && m.CurrentApprovalLevel == l.Seq)
                         })
                         .ToList();
                  
                    var empIds = memos.Where(x => x.lapp != null).Select(s => s.lapp).Select(s => s.EmployeeId).ToList();

                    var emails = GetViewEmployeeQuery(dbContext).Where(memo => empIds.Contains(memo.EmployeeId)).Select(s => s.Email).Distinct().ToList();
                 
                 
                    var memoSerialize = memos.Select(s => s.m).Select(s => new
                    {
                        s.DocumentNo,
                        Document_Types = dbContext.TRNMemoForms.FirstOrDefault(x => x.MemoId == s.MemoId && x.obj_label.Contains("ประเภทเอกสาร"))?.obj_value,
                        Document_Number = dbContext.TRNMemoForms.FirstOrDefault(x => x.MemoId == s.MemoId && x.obj_label.Contains("รหัสเอกสาร"))?.obj_value,
                        s.MemoSubject,
                        s.StatusName,
                        s.PersonWaiting,
                        Link = _URLWeb+ s.MemoId
                    }).ToList();

                    var memosHtml = DataTableToHtml(ToDataTable(memoSerialize));

                    string to = string.Join(";", emails);
                    var emailStateModel = GetEmailTemplateByFormState(connectionString, formState);
                    emailStateModel.EmailBody = ReplaceMailBody(emailStateModel.EmailBody, memosHtml);
                    SendEmailTemplate(emailStateModel, to, noticeCC);

                }
            }
            catch (Exception ex)
            {
                Log("DAR inprocess : " + ex.Message.ToString());
                if (ex.InnerException != null)
                {
                    Log("Inner : " + ex.InnerException.Message.ToString());
                }
            }
            return "";
        }
        public static string NotificationEForm_InProcess(WolfApproveModel dbContext, string connectionString, string formState, int memoId, string noticeCC)
        {
            List<int?> templateModel = dbContext.MSTTemplates
                .Where(x => x.GroupTemplateName.Contains("E-Form"))
                .Select(s => s.TemplateId)
                .ToList();

            if (templateModel.Any())
            {
                var memos = dbContext.TRNMemoes
                    .Where(m => templateModel.Contains(m.TemplateId) && !_NotInProcess.Contains(m.StatusName))
                    .Select(m => new
                    {
                        m,
                        lapp = dbContext.TRNLineApproves.FirstOrDefault(l => m.MemoId == l.MemoId && m.CurrentApprovalLevel == l.Seq),
                        //personInflow = new List<TRNLineApprove>()
                    })
                    .ToList();

                var personInflows = new List<TRNLineApprove>();

                foreach (var memo in memos)
                {
                    personInflows.AddRange(dbContext.TRNLineApproves.Where(x => x.MemoId == memo.m.MemoId).ToList());
                }

                var lapps = memos.Select(s => s.lapp ?? new TRNLineApprove()).ToList();
                var empIds = lapps.Select(s => s.EmployeeId).ToList();

                var emails = GetViewEmployeeQuery(dbContext).Where(memo => empIds.Contains(memo.EmployeeId)).Select(s => s.Email).Distinct().ToList();

                var memoSerialize = memos.Select(s => s.m).Select(s => new
                {
                    s.DocumentNo,
                    s.MemoSubject,
                    s.StatusName,
                    s.PersonWaiting,
                    Link = _URLWeb+ s.MemoId
                }).ToList();

                var memosHtml = DataTableToHtml(ToDataTable(memoSerialize));
                //var personInflows = memos.SelectMany(s => s.personInflow ?? new List<TRNLineApprove>()).ToList();
                var empIdsCC = personInflows.Select(s => s.EmployeeId).ToList();

                var emailCC = GetViewEmployeeQuery(dbContext)
                            .Where(x => empIdsCC.Contains(x.EmployeeId))
                            .Select(s => s.Email)
                            .Distinct()
                            .ToList();

                string to = string.Join(";", emails);

                noticeCC = string.Join(";", emailCC);
                var emailStateModel = GetEmailTemplateByFormState(connectionString, formState);
                emailStateModel.EmailBody = ReplaceMailBody(emailStateModel.EmailBody, memosHtml);
                SendEmailTemplate(emailStateModel, to, noticeCC);

            }
            return "";
        }
        public static string Notification_DAR_CopyReport(WolfApproveModel dbContext, string connectionString, string formState, int memoId, string noticeCC)
        {
            var templateModel = dbContext.MSTTemplates
                .Where(x => x.DocumentCode == "DAR-COPY")
                .Select(x => x.TemplateId)
                .ToList();
            if (templateModel.Any())
            {
                var memos = dbContext.TRNMemoes
                    .Where(m => templateModel.Contains(m.TemplateId) && !_NotInProcess.Contains(m.StatusName))
                    .Select(m => new
                    {
                        m,
                        lapp = dbContext.TRNLineApproves.FirstOrDefault(l => m.MemoId == l.MemoId && m.CurrentApprovalLevel == l.Seq)
                    })
                    .ToList();

                var empIds = memos.Where(x => x.lapp != null).Select(s => s.lapp).Select(s => s.EmployeeId).ToList();

                var emails = GetViewEmployeeQuery(dbContext).Where(memo => empIds.Contains(memo.EmployeeId)).Select(s => s.Email).Distinct().ToList();

                var memoSerialize = memos.Select(s => s.m).Select(s => new
                {
                    s.DocumentNo,
                    s.MemoSubject,
                    Request_area = dbContext.TRNMemoForms.FirstOrDefault(x => x.MemoId == s.MemoId && x.obj_label == "รหัสพื้นที่ ISO")?.obj_value + " : " + dbContext.TRNMemoForms.FirstOrDefault(x => x.MemoId == s.MemoId && x.obj_label == "พื้นที่ ISO")?.obj_value,
                    s.StatusName,
                    s.PersonWaiting,
                    Link = _URLWeb+ s.MemoId
                }).ToList();

                var memosHtml = DataTableToHtml(ToDataTable(memoSerialize));

                string to = string.Join(";", emails);
                CustomEmailTemplate emailStateModel = GetEmailTemplateByFormState(connectionString, formState);
                emailStateModel.EmailBody = ReplaceMailBody(emailStateModel.EmailBody, memosHtml);
                SendEmailTemplate(emailStateModel, to, noticeCC);
            }
            return "";

        }
        public static string NotificationIAS(WolfApproveModel dbContext, string connectionString, string formState, int memoId, string noticeCC)
        {
            MSTTemplate templateModel = dbContext.MSTTemplates.FirstOrDefault(x => x.DocumentCode == "IAS");
            if (templateModel != null)
            {
                var memos = dbContext.TRNMemoes
                    .Where(m => m.TemplateId == templateModel.TemplateId && !_NotInProcess.Contains(m.StatusName))
                    .Select(m => new
                    {
                        m,
                        lapp = dbContext.TRNLineApproves.FirstOrDefault(l => m.MemoId == l.MemoId && m.CurrentApprovalLevel == l.Seq)
                    })
                    .Where(x => x.lapp != null && (x.lapp.SignatureEn == "Accept" || x.lapp.SignatureTh == "Accept"))
                    .ToList();

                var empIds = memos.Where(x => x.lapp != null).Select(s => s.lapp).Select(s => s.EmployeeId).ToList();

                var emails = GetViewEmployeeQuery(dbContext).Where(memo => empIds.Contains(memo.EmployeeId)).Select(s => s.Email).Distinct().ToList();

                var memoSerialize = memos.Select(s => s.m).Select(s => new
                {
                    s.DocumentNo,
                    Auditee_ISO_area_Code = string.Join(",", dbContext.TRNMemoForms
                                            .Where(x => x.MemoId == s.MemoId
                                            && x.obj_type == "tb"
                                            && x.obj_label == "แผนการตรวจ"
                                            && x.col_label.Contains("รหัสพื้นที ISO ที่ถูกตรวจ")).Select(ts => ts.col_value).ToList()),
                    s.MemoSubject,
                    s.StatusName,
                    s.PersonWaiting,
                    Link = _URLWeb+ s.MemoId
                }).ToList();

                var memosHtml = DataTableToHtml(ToDataTable(memoSerialize));

                string to = string.Join(";", emails);
                CustomEmailTemplate emailStateModel = GetEmailTemplateByFormState(connectionString, formState);
                emailStateModel.EmailBody = ReplaceMailBody(emailStateModel.EmailBody, memosHtml);
                SendEmailTemplate(emailStateModel, to, noticeCC);
            }
            return "";
        }
        public static string Notification_CAR_ARS(WolfApproveModel dbContext, string connectionString, string formState, int memoId, string noticeCC)
        {
            List<int?> templateModel = dbContext.MSTTemplates
                .Where(x => x.DocumentCode == "ARS" || x.DocumentCode == "CAR")
                .Select(s => s.TemplateId)
                .ToList();

            if (templateModel.Any())
            {
                List<TRNMemo> memos = dbContext.TRNMemoes
                    .Where(x => templateModel.Contains(x.TemplateId) && Ext.Status._Draft != x.StatusName)
                    .ToList();

                foreach (TRNMemo memo in memos)
                {
                    List<AdvanceFormExt.AdvanceForm> advance = AdvanceFormExt.ToList(memo.MAdvancveForm);
                    string effectiveDateString = advance.FirstOrDefault(x => x.label == "วันที่เปิด")?.value;
                    DateTime effectiveDate = DateTimeHelper.ConvertStringToDateTime(effectiveDateString) ?? DateTime.MinValue;

                    if (TruncateTime(effectiveDate) == TruncateTime(DateTime.Now))
                    {
                        string department = advance.FirstOrDefault(x => x.label.Contains("ฝ่าย / กลุ่มงาน / แผนก ผู้จัดทำแผน"))?.value;
                        MSTDepartment deptModel = dbContext.MSTDepartments.FirstOrDefault(x => x.NameEn == department || x.NameTh == department);
                        List<string> emps = GetViewEmployeeQuery(dbContext)
                            .Where(x => x.DepartmentId == deptModel.DepartmentId)
                            .Select(s => s.Email)
                            .ToList();

                        List<int?> empInFlow = dbContext.TRNLineApproves
                            .Where(x => x.MemoId == memo.MemoId)
                            .Select(s => s.EmployeeId)
                            .ToList();
                        List<string> empFlowEmail = GetViewEmployeeQuery(dbContext)
                            .Where(x => empInFlow.Contains(x.EmployeeId))
                            .Select(s => s.Email)
                            .ToList();

                        string dear = "";
                        string to = string.Join(";", emps.Union(empFlowEmail));
                        CustomEmailTemplate emailStateModel = GetEmailTemplateByFormState(connectionString, formState);
                        emailStateModel = SerializeFormEmail(emailStateModel, connectionString, memoId, memo, dear);
                        SendEmailTemplate(emailStateModel, to, noticeCC);
                    }
                }
            }
            return "";
        }
        public static string Report_PR_WI(WolfApproveModel dbContext, string connectionString, string formState, int memoId, string noticeCC)
        {
            var beforDate = TruncateTime(DateTime.Now.AddMonths(-1));

            var templateModel = dbContext.MSTTemplates
                .Where(x => x.DocumentCode == "DAR-EDIT" || x.DocumentCode == "DAR-NEW")
                .Select(s => s.TemplateId)
                .ToList();

            if (templateModel.Any())
            {
                var memos = dbContext.TRNMemoes
                    .Where(m => templateModel.Contains(m.TemplateId) &&
                    m.StatusName == Ext.Status._Completed && dbContext.TRNMemoForms.Any(x => x.MemoId == m.MemoId && x.obj_label.Contains("ประเภทเอกสาร")
                    && (x.obj_value.Contains("(WI)") || x.obj_value.Contains("(PR)"))))
                    .Select(s => new MemoDto
                    {
                        MemoId = s.MemoId,
                        MAdvancveForm = s.MAdvancveForm
                    })
                    .ToList();

                var memoProcess = new List<ReportPR_WI>();

                foreach (var memo in memos)
                {
                    var memoForm = dbContext.TRNMemoForms.Where(x => x.MemoId == memo.MemoId);
                    var docType = memoForm.FirstOrDefault(x => x.obj_label.Contains("ประเภทเอกสาร"))?.obj_value;
                    var effectiveDateString = memoForm.FirstOrDefault(x => x.obj_label.Contains("วันที่ต้องการประกาศใช้"))?.obj_value;
                    var effectiveDate = TruncateTime(DateTimeHelper.ConvertStringToDateTime(effectiveDateString) ?? DateTime.MinValue);

                    //Ext.WriteLogFile("effectiveDateString : " + effectiveDateString);

                    if ((docType.Contains("(WI)") || docType.Contains("(PR)")) && effectiveDate >= beforDate && effectiveDate <= TruncateTime(DateTime.Now))
                    {
                        var table = memoForm.Where(x => x.obj_label.Contains("ระบบ / มาตรฐาน และ ข้อกำหนดที่เกี่ยวข้อง") && x.col_label.Contains("ระบบ / มาตรฐาน")).Select(s => s.col_value).ToList();
                        if (table.Contains("ISO 9001"))
                        {
                            //var advanceform = AdvanceFormExt.ToList(memo.MAdvancveForm);
                            //var tableDivision = advanceform.FirstOrDefault(x => x.label.Contains("ฝ่าย/กลุ่มงาน/แผนกที่เกี่ยวข้อง"))?.row ?? new List<List<AdvanceFormExt.AdvanceFormRow>>();

                            var tableDivisionIsoCode = memoForm.Where(x => x.obj_label.Contains("ฝ่าย/กลุ่มงาน/แผนกที่เกี่ยวข้อง") && x.col_label == "รหัสพื้นที ISO").Select(s => s.col_value).ToList();
                            var tableDivisionIsoArea = memoForm.Where(x => x.obj_label.Contains("ฝ่าย/กลุ่มงาน/แผนกที่เกี่ยวข้อง") && x.col_label == "พื้นที ISO").Select(s => s.col_value).ToList();
                            var tableDivisionDept = memoForm.Where(x => x.obj_label.Contains("ฝ่าย/กลุ่มงาน/แผนกที่เกี่ยวข้อง") && x.col_label == "ฝ่าย/กลุ่มงาน/แผนก").Select(s => s.col_value).ToList();
                            var tableDivisionLevel = memoForm.Where(x => x.obj_label.Contains("ฝ่าย/กลุ่มงาน/แผนกที่เกี่ยวข้อง") && x.col_label == "ระดับความเกี่ยวข้อง").Select(s => s.col_value).ToList();

                            for (int i = 0; i < tableDivisionIsoArea.Count; i++)
                            {
                                var model = new ReportPR_WI()
                                {
                                    Document_Types = dbContext.TRNMemoForms.FirstOrDefault(x => x.MemoId == memo.MemoId && x.obj_label.Contains("ประเภทเอกสาร"))?.obj_value,
                                    MCQR_Code = dbContext.TRNMemoForms.FirstOrDefault(x => x.MemoId == memo.MemoId && x.obj_label.Contains("MCQR"))?.obj_value,
                                    YQR_Code = dbContext.TRNMemoForms.FirstOrDefault(x => x.MemoId == memo.MemoId && x.obj_label.Contains("YQR"))?.obj_value,
                                    Document_Number = dbContext.TRNMemoForms.FirstOrDefault(x => x.MemoId == memo.MemoId && x.obj_label.Contains("รหัสเอกสาร"))?.obj_value,
                                    Document_Name = dbContext.TRNMemoForms.FirstOrDefault(x => x.MemoId == memo.MemoId && x.obj_label.Contains("ชื่อเอกสาร"))?.obj_value,
                                    RevisionDot = dbContext.TRNMemoForms.FirstOrDefault(x => x.MemoId == memo.MemoId && x.obj_label.Contains("แก้ไขครั้งที่"))?.obj_value,
                                    ISO_area_Code = tableDivisionIsoCode?.ElementAtOrDefault(i),
                                    ISO_area = tableDivisionIsoArea?.ElementAtOrDefault(i),
                                    DivisionSlGroupSlSection = tableDivisionDept?.ElementAtOrDefault(i),
                                    Related_level = tableDivisionLevel?.ElementAtOrDefault(i),
                                    Link = _URLWeb+ memo.MemoId
                                };
                                memoProcess.Add(model);
                            }
                        }
                    }
                }

                memos.Clear();

                var emails = new List<string>
                {
                    "TYM_ISO_00@yamaha-motor.co.th",
                    "TYM_ISO_01@yamaha-motor.co.th",
                    "TYM_ISO_02@yamaha-motor.co.th",
                    "TYM_ISO_03@yamaha-motor.co.th",
                    "TYM_ISO_04@yamaha-motor.co.th",
                    "TYM_ISO_05@yamaha-motor.co.th",
                    "TYM_ISO_06@yamaha-motor.co.th",
                    "TYM_ISO_07@yamaha-motor.co.th",
                    "TYM_ISO_08@yamaha-motor.co.th",
                    "TYM_ISO_09@yamaha-motor.co.th",
                    "TYM_ISO_10@yamaha-motor.co.th",
                    "TYM_ISO_11@yamaha-motor.co.th",
                    "TYM_ISO_12@yamaha-motor.co.th",
                    "TYM_ISO_13@yamaha-motor.co.th",
                    "TYM_ISO_14@yamaha-motor.co.th",
                    "TYM_ISO_15@yamaha-motor.co.th",
                    "TYM_ISO_16@yamaha-motor.co.th",
                    "TYM_ISO_17@yamaha-motor.co.th",
                    "TYM_ISO_18@yamaha-motor.co.th",
                    "TYM_ISO_19@yamaha-motor.co.th",
                    "TYM_ISO_20@yamaha-motor.co.th",
                    "TYM_ISO_21@yamaha-motor.co.th",
                    "TYM_ISO_22@yamaha-motor.co.th",
                    "TYM_ISO_23@yamaha-motor.co.th",
                    "TYM_ISO_24@yamaha-motor.co.th",
                    "TYM_ISO_25@yamaha-motor.co.th",
                    "TYM_ISO_26@yamaha-motor.co.th",
                    "TYM_ISO_27@yamaha-motor.co.th",
                    "TYM_ISO_28@yamaha-motor.co.th",
                    "TYM_ISO_29@yamaha-motor.co.th",
                    "TYM_ISO_30@yamaha-motor.co.th",
                    "TYM_ISO_31@yamaha-motor.co.th",
                    "TYM_ISO_32@yamaha-motor.co.th",
                };
                var memosHtml = DataTableToHtml(ToDataTable(memoProcess));

                string to = string.Join(";", emails);
                CustomEmailTemplate emailStateModel = GetEmailTemplateByFormState(connectionString, formState);
                emailStateModel.EmailBody = ReplaceMailBody(emailStateModel.EmailBody, memosHtml);
                SendEmailTemplate(emailStateModel, to, noticeCC);
            }

            return "";
        }
        public static string NotificationIAC(WolfApproveModel dbContext, string connectionString, string formState, int memoId, string noticeCC)
        {
            MSTTemplate templateModel = dbContext.MSTTemplates.FirstOrDefault(x => x.DocumentCode == "IAC" && x.IsActive == true);
            if (templateModel != null)
            {
                var memos = dbContext.TRNMemoes
                    .Where(m => m.TemplateId == templateModel.TemplateId && !_NotInProcess.Contains(m.StatusName))
                    .Select(m => new
                    {
                        m,
                        lapp = dbContext.TRNLineApproves.FirstOrDefault(l => m.MemoId == l.MemoId && m.CurrentApprovalLevel == l.Seq)
                    })
                    .Where(x => x.lapp != null && (x.lapp.SignatureEn.Contains("Accept Result") || x.lapp.SignatureTh.Contains("Accept Result")))
                    .ToList();

                var empIds = memos.Where(x => x.lapp != null).Select(s => s.lapp).Select(s => s.EmployeeId ?? 0).ToList();
                Log($"{memos.Count}");
                var emails = GetViewEmployeeQuery(dbContext).Where(memo => empIds.Contains(memo.EmployeeId)).Select(s => s.Email).Distinct().ToList();
                Log($"{emails.ToJson()}");
                var memoSerialize = memos.Select(s => s.m).Select(s => new
                {
                    s.DocumentNo,
                    Auditee_ISO_area_Code = dbContext.TRNMemoForms.FirstOrDefault(x => x.MemoId == s.MemoId && x.obj_label == "รหัสพื้นที่ ISO ที่ถูกตรวจ")?.obj_value,
                    s.MemoSubject,
                    s.StatusName,
                    s.PersonWaiting,
                    Link = _URLWeb+ s.MemoId
                }).ToList();

                var memosHtml = DataTableToHtml(ToDataTable(memoSerialize));

                string to = string.Join(";", emails);
                CustomEmailTemplate emailStateModel = GetEmailTemplateByFormState(connectionString, formState);
                emailStateModel.EmailBody = ReplaceMailBody(emailStateModel.EmailBody, memosHtml);
                SendEmailTemplate(emailStateModel, to, noticeCC);
            }
            return "";
        }
        public static string NotificationDAR_N(WolfApproveModel dbContext, string connectionString, string formState, int memoId, string noticeCC, string option)
        {
            var templateCodes = new List<string>
            {
                "DAR-NEW",
                "DAR-EDIT"
            };

            //var templateEditIds = dbContext.MSTTemplates.Where(x => x.DocumentCode == "DAR-EDIT").Select(s => s.TemplateId ?? 0).ToList();

            //var memoEditIds = dbContext.TRNMemoes.Where(x => x.StatusName != Ext.Status._Draft && templateEditIds.Contains(x.TemplateId ?? 0)).Select(s => s.MemoId).ToList();

            var optionModel = JsonConvert.DeserializeObject<NotificationOptionModel>(option);

            var result = int.TryParse(optionModel.Before, out int before);

            Log("before : " + before);

            var templateModel = dbContext.MSTTemplates
                .Where(x => templateCodes.Contains(x.DocumentCode))
                .Select(s => s.TemplateId)
                .ToList();

            if (templateModel.Any())
            {
                var memos = dbContext.TRNMemoes
                     .Where(m => templateModel.Contains(m.TemplateId) && m.StatusName == Ext.Status._Completed)
                     .ToList();

                var memoProcess = new List<TRNMemo>();
                var mailTo = new List<string>();
                var mailCC = new List<string>();

                foreach (var memo in memos)
                {
                    //var advanceForm = AdvanceFormExt.ToList(memo.MAdvancveForm);
                    string effectiveDateString = ReserveCarExt.getValueAdvanceForm(memo.MAdvancveForm, "วันที่ต้องทบทวนเอกสาร");
                    string section = ReserveCarExt.getValueAdvanceForm(memo.MAdvancveForm, "ชื่อฝ่าย/กลุ่มงาน/แผนก");

                    var documentNumber = dbContext.TRNMemoForms.FirstOrDefault(x => x.MemoId == memo.MemoId && x.obj_label.Contains("รหัสเอกสาร"))?.obj_value;

                    DateTime initialEffectiveDate = TruncateTime(DateTimeHelper.ConvertStringToDateTime(effectiveDateString) ?? DateTime.MinValue);
                    DateTime currentDate = TruncateTime(DateTime.Now);

                    try
                    {
                        //DateTime effectiveDate = TruncateTime(initialEffectiveDate).AddDays(-before);

                        if (initialEffectiveDate == currentDate /*&& effectiveDate < initialEffectiveDate*/)
                        {
                            memoProcess.Add(memo);
                            var emails = GetDCCArea(section, dbContext);
                            mailTo.AddRange(emails);
                            mailCC.AddRange(GetHeadWorkingArea(section, dbContext));
                            mailCC.AddRange(GetWorkingArea(section, dbContext));
                        }
                    }
                    catch
                    {

                    }


                    //TimeSpan timeDifference = currentDate - initialEffectiveDate;

                    //if (initialEffectiveDate >= currentDate && timeDifference.Days % 3 == 0)
                    //{
                    //    var checkEdit = dbContext.TRNMemoForms.Any(x => memoEditIds.Contains(x.MemoId) && x.obj_label == "รหัสเอกสาร" && x.obj_value == documentNumber);
                    //    if (!checkEdit)
                    //    {
                    //        memoProcess.Add(memo);
                    //        var emails = GetDCCArea(section, dbContext);
                    //        mailTo.AddRange(emails);
                    //        mailCC.AddRange(GetHeadWorkingArea(section, dbContext));
                    //        mailCC.AddRange(GetWorkingArea(section, dbContext));
                    //    }
                    //}
                }

                //var emails = GetViewEmployeeQuery(dbContext).Where(memo => empIds.Contains(memo.EmployeeId)).Select(s => s.Email).Distinct().ToList();
                var memoSerialize = memoProcess.Select(s => new
                {
                    DivisionSlGroupSlSection = dbContext.TRNMemoForms.FirstOrDefault(x => x.MemoId == s.MemoId && x.obj_label.Contains("ชื่อฝ่าย/กลุ่มงาน/แผนก"))?.obj_value,
                    Document_Types = dbContext.TRNMemoForms.FirstOrDefault(x => x.MemoId == s.MemoId && x.obj_label.Contains("ประเภทเอกสาร"))?.obj_value,
                    Document_Number = dbContext.TRNMemoForms.FirstOrDefault(x => x.MemoId == s.MemoId && x.obj_label.Contains("รหัสเอกสาร"))?.obj_value,
                    RevisionDot = dbContext.TRNMemoForms.FirstOrDefault(x => x.MemoId == s.MemoId && x.obj_label.Contains("แก้ไขครั้งที่"))?.obj_value ?? "0",
                    Review_Date = dbContext.TRNMemoForms.FirstOrDefault(x => x.MemoId == s.MemoId && x.obj_label.Contains("วันที่ต้องทบทวนเอกสาร"))?.obj_value,
                    Link = _URLWeb+ s.MemoId
                }).ToList();

                var memosHtml = DataTableToHtml(ToDataTable(memoSerialize));

                string to = string.Join(";", mailTo);
                string cc = string.Join(";", mailCC);
                var emailStateModel = GetEmailTemplateByFormState(connectionString, formState);
                emailStateModel.EmailBody = ReplaceMailBody(emailStateModel.EmailBody, memosHtml);
                SendEmailTemplate(emailStateModel, to, cc + ";" + noticeCC);
            }

            return "";
        }
        public static string NotificationMasterArea(WolfApproveModel dbContext, string connectionString, string formState, int memoId, string noticeCC)
        {
            var beforeDate = TruncateTime(DateTime.Now).AddDays(-1);
            var masterDatas = dbContext.MSTMasterDatas.Where(x => EntityFunctions.TruncateTime(x.ModifiedDate ?? DateTime.MinValue) == beforeDate && x.MasterType == "Area")
                .Select(s => new
                {
                    AreaCode = s.Value2,
                    AreaName = s.Value3,
                    s.ModifiedDate
                }).ToList();

            var masterDatasSerialize = masterDatas.Select(s => new
            {
                Area_Code = s.AreaCode,
                Area_Name = s.AreaName,
                Modified_Date = s.ModifiedDate != null ? s.ModifiedDate.Value.ToString("dd/MM/yyyy HH:mm:ss tt") : ""
            }).ToList();

            var emails = new List<string>();

            foreach (var master in masterDatas)
            {
                emails.AddRange(GetEmailInArea(master.AreaCode, dbContext));
            }

            var memosHtml = DataTableToHtml(ToDataTable(masterDatasSerialize));

            string to = string.Join(";", emails);
            var emailStateModel = GetEmailTemplateByFormState(connectionString, formState);
            emailStateModel.EmailBody = ReplaceMailBody(emailStateModel.EmailBody, memosHtml);
            SendEmailTemplate(emailStateModel, to, noticeCC);
            return "";
        }
        public static string Notification_ISO(WolfApproveModel dbContext, string connectionString, string formState, int memoId, string noticeCC)
        {
            MSTTemplate templateModel = dbContext.MSTTemplates.FirstOrDefault(x => x.DocumentCode == "DAR-NEW");
            if (templateModel != null)
            {
                List<TRNMemo> memos = dbContext.TRNMemoes.Where(x => x.TemplateId == templateModel.TemplateId && x.StatusName != Ext.Status._Draft).ToList();
                foreach (TRNMemo memo in memos)
                {
                    try
                    {
                        List<AdvanceFormExt.AdvanceForm> advance = AdvanceFormExt.ToList(memo.MAdvancveForm);
                        string effectiveDateString = advance.FirstOrDefault(x => x.label == "วันที่ต้องการประกาศใช้")?.value;
                        DateTime effectiveDate = DateTimeHelper.ConvertStringToDateTime(effectiveDateString) ?? DateTime.MinValue;

                        if (TruncateTime(effectiveDate) == TruncateTime(DateTime.Now))
                        {
                            string section = advance.FirstOrDefault(x => x.label.Contains("ชื่อฝ่าย/กลุ่มงาน/แผนก"))?.value;
                            List<string> emails = GetEmailInArea(section, dbContext);
                            string dear = "";
                            string to = string.Join(";", emails);
                            CustomEmailTemplate emailStateModel = GetEmailTemplateByFormState(connectionString, formState);
                            emailStateModel = SerializeFormEmail(emailStateModel, connectionString, memoId, memo, dear);
                            SendEmailTemplate(emailStateModel, to, noticeCC);
                        }
                    }
                    catch
                    {

                    }
                }
            }
            return "";
        }
        public static string Notification_DAR_Copy_Expired(WolfApproveModel dbContext, string connectionString, string formState, int memoId, string noticeCC)
        {
            List<int?> templateModel = dbContext.MSTTemplates
                .Where(x => x.DocumentCode == "DAR-COPY")
                .Select(s => s.TemplateId)
                .ToList();

            if (templateModel.Any())
            {
                var memos = dbContext.TRNMemoes
                    .Where(m => templateModel.Contains(m.TemplateId) && m.StatusName != Ext.Status._Draft && m.StatusName != Ext.Status._Rejected)
                    .Select(m => new
                    {
                        m,
                        lineFirst = dbContext.TRNLineApproves.FirstOrDefault(x => x.MemoId == m.MemoId && x.Seq == 1)
                    })
                    .ToList();

                foreach (var memo in memos)
                {
                    List<AdvanceFormExt.AdvanceForm> advance = AdvanceFormExt.ToList(memo.m.MAdvancveForm);
                    string expiredDateString = advance.FirstOrDefault(x => x.label.Contains("วันที่สิ้นสุดการแจกจ่าย"))?.value;
                    DateTime expiredDate = (DateTimeHelper.ConvertStringToDateTime(expiredDateString) ?? DateTime.MinValue).AddDays(1);
                    Log("expiredDate : " + expiredDateString);
                    if (TruncateTime(expiredDate) == TruncateTime(DateTime.Now))
                    {
                        CustomEmailTemplate emailStateModel = GetEmailTemplateByFormState(connectionString, formState);
                        MSTEmployee empModel = dbContext.MSTEmployees.FirstOrDefault(x => x.EmployeeId == memo.m.RequesterId);
                        MSTEmployee empLine = dbContext.MSTEmployees.FirstOrDefault(x => x.EmployeeId == memo.lineFirst.EmployeeId);
                        string dear = empModel?.NameTh;
                        string to = $"{empModel?.Email};{empLine.Email}";
                        emailStateModel = SerializeFormEmail(emailStateModel, connectionString, memo.m.MemoId, memo.m, dear, expiredDateString);
                        SendEmailTemplate(emailStateModel, to, noticeCC);
                    }
                }
            }
            return "";
        }
        public static string Report_Announcement(WolfApproveModel dbContext, string connectionString, string formState, int memoId, string noticeCC)
        {
            var announcementNames = new List<string>
            {
                "Basic Quality Policy",
                "ประกาศแต่งตั้งตัวแทนฝ่ายบริหาร (MRs)",
                "ประกาศคุณสมบัติของตัวแทนฝ่ายบริหาร",
                "ประกาศแต่งตั้งคณะกรรมการระบบบริหารฯ",
                "ประกาศแต่งตั้งคณะทำงานระบบบริหารฯ",
                "ประกาศคุณสมบัติของหัวหน้าคณะทำงาน",
                "ประกาศหน้าที่ความรับผิดชอบของพนักงานที่เกี่ยวกับระบบบริหารฯ",
                "ประกาศแต่งตั้งผู้ตรวจติดตามภายในฯ",
                "ประกาศคุณสมบัติของผู้ตรวจติดตามภายในฯ",
                "ประกาศแต่งตั้งคณะกรรมการความปลอดภัยฯ",
                "ประกาศแต่งตั้งบุคลากรด้านสิ่งแวดล้อมฯ",
                "ประกาศแต่งตั้งผู้อำนวยการเหตุฉุกเฉินฯ",
                "ประกาศคุณสมบัติของผู้อำนวยการเหตุฉุกเฉินฯ",
                "ประกาศแต่งตั้งทีมควบคุมเหตุฉุกเฉินฯ",
                "ประกาศแต่งตั้งทีมเก็บกู้หลังเหตุฉุกเฉินฯ",
                "ประกาศแต่งตั้งบุคลากรพร้อมรับผิดชอบความปลอดภัยวัสดุอันตราย",
                "ประกาศแต่งตั้งทีมงานด้านการอนุรักษ์พลังงาน",
                "แผนสื่อสารภายในองค์กร",
                "ทะเบียนรายการสื่อสาร"
            };

            var reportDocumentCode = dbContext.MSTReportTemplates.Where(x => announcementNames.Contains(x.ReportName) && x.IsActive == true)
                .Select(s => s.TemplateId)
                .ToList();

            var documentCodes = reportDocumentCode.SelectMany(s => s.Split('|')).Distinct().ToList();

            var templateIds = dbContext.MSTTemplates.Where(x => documentCodes.Contains(x.DocumentCode)).Select(s => s.TemplateId);

            var beforeDate = TruncateTime(DateTime.Now).AddDays(-1);

            var memos = dbContext.TRNMemoes.Where(x =>
                EntityFunctions.TruncateTime(x.ModifiedDate) == beforeDate &&
                templateIds.Contains(x.TemplateId) &&
                x.StatusName == Ext.Status._Completed)
                .ToList();

            var memoSerialize = memos.Select(s => new
            {
                Document_NoDot = s.DocumentNo,
                Category = dbContext.TRNMemoForms.FirstOrDefault(x => x.obj_label == "Category" && s.MemoId == x.MemoId)?.obj_value,
                ประเภทเอกสาร = dbContext.TRNMemoForms.FirstOrDefault(x => x.obj_label == "ประเภทเอกสาร" && s.MemoId == x.MemoId)?.obj_value,
                ชื่อเอกสาร = dbContext.TRNMemoForms.FirstOrDefault(x => x.obj_label == "ชื่อเอกสาร" && s.MemoId == x.MemoId)?.obj_value,
                Memo_Subject = s.MemoSubject,
                Status_Name = s.StatusName,
                Effective_Date = s.ModifiedDate != null ? s.ModifiedDate.Value.ToString("dd MMM yyyy") : "",
                Link = _URLWeb+ s.MemoId
            }).ToList();

            var categories = memoSerialize.Select(s => s.Category);
            var emails = new List<string>();

            foreach (var category in categories)
            {
                var mailArea = GetEmailInArea(category, dbContext);
                emails.AddRange(mailArea);
            }

            var memosHtml = DataTableToHtml(ToDataTable(memoSerialize));

            string to = string.Join(";", emails);
            CustomEmailTemplate emailStateModel = GetEmailTemplateByFormState(connectionString, formState);
            emailStateModel.EmailBody = ReplaceMailBody(emailStateModel.EmailBody, memosHtml);
            SendEmailTemplate(emailStateModel, to, noticeCC);

            return "";
        }
        public static string Notification_IAS_IAC_Completed(WolfApproveModel dbContext, string connectionString, string formState, int memoId, string noticeCC = "")
        {
            var templateModel = dbContext.MSTTemplates.Where(x => x.DocumentCode == "IAS" || x.DocumentCode == "IAC")
                .Select(s => s.TemplateId).ToList();
            if (templateModel.Any())
            {
                var dtNow = TruncateTime(DateTime.Now);
                var memos = dbContext.TRNMemoes
                    .Where(m => templateModel.Contains(m.TemplateId) && m.StatusName == Ext.Status._Completed && dtNow == EntityFunctions.TruncateTime(m.ModifiedDate))
                    .Select(m => new
                    {
                        m,
                        lapp = dbContext.TRNLineApproves.FirstOrDefault(l => m.MemoId == l.MemoId && m.CurrentApprovalLevel == l.Seq),
                        personInflow = dbContext.TRNLineApproves.Where(l => m.MemoId == l.MemoId).ToList()
                    })
                    .ToList();

                foreach (var memo in memos)
                {
                    var advance = AdvanceFormExt.ToList(memo.m.MAdvancveForm);
                    int employeeId = memo?.lapp?.EmployeeId ?? 0;
                    MSTEmployee empModel = dbContext.MSTEmployees.FirstOrDefault(x => x.EmployeeId == employeeId);
                    if (empModel != null)
                    {
                        var empCC_Id = memo.personInflow.Select(s => s.EmployeeId).ToList();
                        var emailCC = GetViewEmployeeQuery(dbContext)
                            .Where(x => empCC_Id.Contains(x.EmployeeId))
                            .Select(s => s.Email)
                            .ToList();

                        var Auditor = advance.FirstOrDefault(x => x.label.Contains("หัวหน้าผู้ตรวจ"))?.value;
                        var Auditee = advance.FirstOrDefault(x => x.label.Contains("ผู้รับผิดชอบ"))?.value;

                        var empAuditor = GetViewEmployeeQuery(dbContext).FirstOrDefault(x => x.NameEn == Auditor || x.NameTh == Auditor);
                        var empAuditee = GetViewEmployeeQuery(dbContext).FirstOrDefault(x => x.NameEn == Auditee || x.NameTh == Auditee);
                        emailCC.Add(empAuditor?.Email);
                        emailCC.Add(empAuditee?.Email);

                        string dear = empModel?.NameTh;
                        string to = string.Join(";", emailCC);
                        CustomEmailTemplate emailStateModel = GetEmailTemplateByFormState(connectionString, formState);
                        emailStateModel = SerializeFormEmail(emailStateModel, connectionString, memo.m.MemoId, memo.m, dear);
                        SendEmailTemplate(emailStateModel, to, noticeCC);
                    }
                }
            }
            return "";
        }
        public static string NotificationDAR_DocType_FR(WolfApproveModel dbContext, string connectionString, string formState, int memoId, string noticeCC)
        {
            var beforeDate = TruncateTime(DateTime.Now).AddDays(-1);
            var templateModel = dbContext.MSTTemplates
                .Where(x => _DarSet.Contains(x.DocumentCode))
                .Select(s => s.TemplateId)
                .ToList();

            if (templateModel.Any())
            {
                var memos = dbContext.TRNMemoes
                .Where(x => templateModel.Contains(x.TemplateId) /*&& EntityFunctions.TruncateTime(x.Memo.ModifiedDate) == beforeDate*/ && x.StatusName == Ext.Status._Completed)
                .ToList();

                var memoModel = new List<TRNMemo>();

                if (memos.Any())
                {
                    var emails = new List<string>();
                    foreach (var memo in memos)
                    {
                        try
                        {
                            var advanceForm = AdvanceFormExt.ToList(memo.MAdvancveForm);
                            var tableRelated = advanceForm?.FirstOrDefault(x => x.label != null && x.label.Contains("ฝ่าย/กลุ่มงาน/แผนกที่เกี่ยวข้อง"))?.row;

                            var docType = advanceForm?.FirstOrDefault(x => (x.label != null && x.label.Contains("ประเภทเอกสาร")) || (x.alter != null && x.alter.Contains("Document Types")))?.value ?? "";
                            //Ext.WriteLogFile("docType : " + docType + " MemoId : " + memo.MemoId);
                            if (docType.Contains("(FR)"))
                            {
                                string effectiveDateString = advanceForm?.FirstOrDefault(x => (x.label != null && x.label.Contains("วันที่ต้องการประกาศใช้")) || (x.alter != null && x.alter.Contains("Effective Date")))?.value ?? "";
                                DateTime effectiveDate = DateTimeHelper.ConvertStringToDateTime(effectiveDateString) ?? DateTime.MinValue;
                                Log("วันที่ต้องการประกาศใช้ : " + effectiveDateString);
                                if (TruncateTime(effectiveDate) == beforeDate)
                                {
                                    memoModel.Add(memo);
                                    foreach (var row in tableRelated)
                                    {
                                        var areaCode = row.FirstOrDefault(x => x.label != null && x.label.Contains("รหัสพื้นที่ ISO"))?.value;
                                        if(!string.IsNullOrEmpty(areaCode) && areaCode.Split(':').Count() > 0)
                                        emails.AddRange(GetEmailInArea(areaCode, dbContext));
                                    }
                                }
                            }
                        }
                        catch(Exception ex)
                        {
                            Log($"MemoId : {memo.MemoId} Error NotificationDAR_DocType_FR : " + ex);
                        }
                    }
                    var memoSerialize = memoModel.Select(s => new
                    {
                        s.DocumentNo,
                        Document_Types = dbContext.TRNMemoForms.FirstOrDefault(x => x.MemoId == s.MemoId && x.obj_label.Contains("ประเภทเอกสาร"))?.obj_value,
                        Document_Number = dbContext.TRNMemoForms.FirstOrDefault(x => x.MemoId == s.MemoId && x.obj_label.Contains("รหัสเอกสาร"))?.obj_value,
                        s.MemoSubject,
                        RivisionDot = dbContext.TRNMemoForms.FirstOrDefault(x => x.MemoId == s.MemoId && x.obj_label.Contains("แก้ไขครั้งที่"))?.obj_value,
                        DivisionSlGroupSlSection = dbContext.TRNMemoForms.FirstOrDefault(x => x.MemoId == s.MemoId && x.obj_label.Contains("ชื่อฝ่าย/กลุ่มงาน/แผนก"))?.obj_value,
                        s.StatusName,
                        Effective_Date = dbContext.TRNMemoForms.FirstOrDefault(x => x.MemoId == s.MemoId && x.obj_label.Contains("วันที่ต้องการประกาศใช้"))?.obj_value,
                        Link = _URLWeb + s.MemoId
                    }).ToList();

                    var memosHtml = DataTableToHtml(ToDataTable(memoSerialize));

                    string to = string.Join(";", emails);
                    CustomEmailTemplate emailStateModel = GetEmailTemplateByFormState(connectionString, formState);
                    emailStateModel.EmailBody = ReplaceMailBody(emailStateModel.EmailBody, memosHtml);
                    SendEmailTemplate(emailStateModel, to, noticeCC);
                }
            }
            return "";
        }
        public static string Notification_DAR_Copy(WolfApproveModel dbContext, string connectionString, string formState, int memoId, string noticeCC)
        {
            List<int?> templateModel = dbContext.MSTTemplates
                .Where(x => x.DocumentCode == "DAR-COPY")
                .Select(s => s.TemplateId)
                .ToList();

            if (templateModel.Any())
            {
                var memos = dbContext.TRNMemoes
                    .Where(m => templateModel.Contains(m.TemplateId) && m.StatusName != Ext.Status._Draft && m.StatusName != Ext.Status._Completed && m.StatusName != Ext.Status._Rejected)
                    .Select(m => new
                    {
                        m,
                        currentApp = dbContext.TRNLineApproves.FirstOrDefault(l => m.MemoId == l.MemoId && m.CurrentApprovalLevel == l.Seq),
                        personInflow = dbContext.TRNLineApproves.Where(l => m.MemoId == l.MemoId).ToList()
                    })
                    .ToList();

                foreach (var memo in memos)
                {
                    CustomEmailTemplate emailStateModel = GetEmailTemplateByFormState(connectionString, formState);
                    MSTEmployee empModel = dbContext.MSTEmployees.FirstOrDefault(x => x.EmployeeId == memo.m.RequesterId);
                    List<int?> empCC_Id = memo.personInflow.Select(s => s.EmployeeId).ToList();
                    List<string> emailCC = GetViewEmployeeQuery(dbContext)
                        .Where(x => empCC_Id.Contains(x.EmployeeId))
                        .Select(s => s.Email)
                        .ToList();

                    string dear = empModel?.NameTh;
                    string to = empModel?.Email;
                    noticeCC = string.Join(";", emailCC);
                    emailStateModel = SerializeFormEmail(emailStateModel, connectionString, memoId, memo.m, dear);
                    SendEmailTemplate(emailStateModel, to, noticeCC);
                }
            }
            return "";
        }
        public static string Report_DAR_E(WolfApproveModel dbContext, string connectionString, string formState, int memoId, string noticeCC)
        {
            var beforDate = TruncateTime(DateTime.Now.AddMonths(-1));
            var currentDate = TruncateTime(DateTime.Now);
            var templateModel = dbContext.MSTTemplates
                .Where(x => x.DocumentCode == "DAR-EDIT" || x.DocumentCode == "DAR-NEW")
                .Select(s => s.TemplateId)
                .ToList();

            if (templateModel.Any())
            {
                var memos = dbContext.TRNMemoes
                    .Where(m => templateModel.Contains(m.TemplateId) &&
                    m.StatusName != Ext.Status._Draft)
                    .Select(s => new MemoDto
                    {
                        MemoId = s.MemoId
                    }).ToList();
                var memoIds = memos.Select(s => s.MemoId).ToList();
                var memoProcess = new List<MemoDto>();

                var memoforms = dbContext.TRNMemoForms.Where(x => memoIds.Contains(x.MemoId) && x.obj_label.Contains("วันที่ต้องการประกาศใช้")).ToList();
                foreach (var memo in memos)
                {
                    var effectiveDate = DateTimeHelper.ConvertStringToDateTime(memoforms.FirstOrDefault(x => x.MemoId == memo.MemoId && x.obj_label.Contains("วันที่ต้องการประกาศใช้"))?.obj_value) ?? DateTime.MinValue;
                    if (effectiveDate >= beforDate && effectiveDate <= currentDate)
                    {
                        memoProcess.Add(memo);
                    }
                }

                var memoSerialize = memoProcess.Select(s => new
                {
                    Division_SGroup_SSection = dbContext.TRNMemoForms.FirstOrDefault(x => x.MemoId == s.MemoId && x.obj_label.Contains("ชื่อฝ่าย/กลุ่มงาน/แผนก"))?.obj_value,
                    Document_Types = dbContext.TRNMemoForms.FirstOrDefault(x => x.MemoId == s.MemoId && x.obj_label.Contains("ประเภทเอกสาร"))?.obj_value,
                    Document_Number = dbContext.TRNMemoForms.FirstOrDefault(x => x.MemoId == s.MemoId && x.obj_label.Contains("รหัสเอกสาร"))?.obj_value,
                    RevisionDot = dbContext.TRNMemoForms.FirstOrDefault(x => x.MemoId == s.MemoId && x.obj_label.Contains("แก้ไขครั้งที่"))?.obj_value ?? "0",
                    Effective_Date = dbContext.TRNMemoForms.FirstOrDefault(x => x.MemoId == s.MemoId && x.obj_label.Contains("วันที่ต้องการประกาศใช้"))?.obj_value,
                    Link = _URLWeb+ s.MemoId
                }).ToList();

                memoSerialize = memoSerialize.GroupBy(g => g.Document_Number)
                    .Select(s => s.OrderByDescending(o => o.RevisionDot).First())
                    .ToList();

                var memosHtml = DataTableToHtml(ToDataTable(memoSerialize));

                Log("memosHtml : " + memosHtml);

                    var emails = new List<string>
                    {
                        "TYM_ISO_00@yamaha-motor.co.th",
                        "TYM_ISO_01@yamaha-motor.co.th",
                        "TYM_ISO_02@yamaha-motor.co.th",
                        "TYM_ISO_03@yamaha-motor.co.th",
                        "TYM_ISO_04@yamaha-motor.co.th",
                        "TYM_ISO_05@yamaha-motor.co.th",
                        "TYM_ISO_06@yamaha-motor.co.th",
                        "TYM_ISO_07@yamaha-motor.co.th",
                        "TYM_ISO_08@yamaha-motor.co.th",
                        "TYM_ISO_09@yamaha-motor.co.th",
                        "TYM_ISO_10@yamaha-motor.co.th",
                        "TYM_ISO_11@yamaha-motor.co.th",
                        "TYM_ISO_12@yamaha-motor.co.th",
                        "TYM_ISO_13@yamaha-motor.co.th",
                        "TYM_ISO_14@yamaha-motor.co.th",
                        "TYM_ISO_15@yamaha-motor.co.th",
                        "TYM_ISO_16@yamaha-motor.co.th",
                        "TYM_ISO_17@yamaha-motor.co.th",
                        "TYM_ISO_18@yamaha-motor.co.th",
                        "TYM_ISO_19@yamaha-motor.co.th",
                        "TYM_ISO_20@yamaha-motor.co.th",
                        "TYM_ISO_21@yamaha-motor.co.th",
                        "TYM_ISO_22@yamaha-motor.co.th",
                        "TYM_ISO_23@yamaha-motor.co.th",
                        "TYM_ISO_24@yamaha-motor.co.th",
                        "TYM_ISO_25@yamaha-motor.co.th",
                        "TYM_ISO_26@yamaha-motor.co.th",
                        "TYM_ISO_27@yamaha-motor.co.th",
                        "TYM_ISO_28@yamaha-motor.co.th",
                        "TYM_ISO_29@yamaha-motor.co.th",
                        "TYM_ISO_30@yamaha-motor.co.th",
                        "TYM_ISO_31@yamaha-motor.co.th",
                        "TYM_ISO_32@yamaha-motor.co.th",
                        "TYM_ISO_33@yamaha-motor.co.th",
                    };

                    string to = string.Join(";", emails);
                    CustomEmailTemplate emailStateModel = GetEmailTemplateByFormState(connectionString, formState);
                    emailStateModel.EmailBody = ReplaceMailBody(emailStateModel.EmailBody, memosHtml);
                    SendEmailTemplate(emailStateModel, to, noticeCC);
                }

            return "";
        }

        public static CustomEmailTemplate GetEmailTemplateByFormState(string connectionString, string FormState)
        {
            try
            {
                using (WolfApproveModel db = DBContext.OpenConnection(connectionString))
                {
                    MSTEmailTemplate item = db.MSTEmailTemplates.FirstOrDefault(x => x.FormState == FormState && x.IsActive == true);
                    if (item != null)
                    {
                        CustomEmailTemplate customData = new CustomEmailTemplate();
                        customData.RetrieveFromDtoWithDateTimeFormat(item);
                        customData.connectionString = connectionString;
                        return customData;
                    }
                    else
                    {
                        return null;
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public static CustomEmailTemplate SerializeFormEmail(CustomEmailTemplate request, string connectionString, int memoId, TRNMemo memo, string dear, string dateTime = "")
        {
            string sSubURLToRequest = string.IsNullOrEmpty(ServiceController.SubURLToRequest) ? "Memo?pk=" : ServiceController.SubURLToRequest;
            string sURLToRequest = ServiceController.GetURLWebForEmail(connectionString) ?? "";
            if (!string.IsNullOrEmpty(Ext.Port))
            {
                sURLToRequest += ":" + Ext.Port;
            }
            sURLToRequest += $"/{sSubURLToRequest}{memoId}";
            string sURLToMobile = ServiceController.GetURLMobileForEmail();
            request.EmailSubject = ReplaceEmailContent(request.EmailSubject, memo, dear, sURLToRequest, sURLToMobile, dateTime);
            request.EmailBody = ReplaceEmailContent(request.EmailBody, memo, dear, sURLToRequest, sURLToMobile, dateTime);
            return request;
        }
        public static string ReplaceEmailContent(string sContent, TRNMemo memo, string dear, string sURLToRequest, string sURLToMobile, string dateTime)
        {
            sContent = sContent
                .Replace("[DearName]", dear)
                .Replace("[TRNMemo_DocumentNo]", memo.DocumentNo)
                .Replace("[TRNMemo_TemplateSubject]", memo.MemoSubject)
                .Replace("[TRNMemo_RNameEn]", dear)
                .Replace("[TRNMemo_RequestDate]", memo.RequestDate.Value.ToString(Ext.FormatDate_ddMMMyyyy))
                .Replace("[TRNActionHistory_ActionDate]", DateTime.Now.ToString(Ext.FormatDefultsDateTimeHHmm))
                .Replace("[TRNMemo_StatusName]", memo.StatusName)
                .Replace("[TRNMemo_CompanyName]", memo.CompanyName)
                .Replace("[TRNMemo_TemplateName]", memo.TemplateName)
                .Replace("[URLToRequest]", $"<a href='{sURLToRequest}'>Click</a>")
                .Replace("[URLToMobile]", $"<a href='{sURLToMobile}'>Click</a>")
                .Replace("[DATETIME]", dateTime);

            Log($"Body : {sContent}");
            return sContent;
        }
        public static string GenLinkMemoByMemoId(string connectionString, int memoId)
        {
            string sSubURLToRequest = string.IsNullOrEmpty(ServiceController.SubURLToRequest) ? "Memo?pk=" : ServiceController.SubURLToRequest;
            string sURLToRequest = ServiceController.GetURLWebForEmail(connectionString) ?? "";
            if (!string.IsNullOrEmpty(Ext.Port))
            {
                sURLToRequest += ":" + Ext.Port;
            }
            sURLToRequest += $"/{sSubURLToRequest}{memoId}";
            return sURLToRequest;
        }
        public static string DataTableToHtml(DataTable dataTable)
        {
            StringBuilder html = new StringBuilder();

            html.Append("<table border='1' style='border-collapse: collapse;'>");

            html.Append("<tr>");
            foreach (DataColumn column in dataTable.Columns)
            {
                int columnWidth = column.ColumnName.Length + 10;
                html.AppendFormat("<th style='width:{0}px'>{1}</th>", columnWidth, column.ColumnName);
            }
            html.Append("</tr>");

            foreach (DataRow row in dataTable.Rows)
            {
                html.Append("<tr>");
                foreach (DataColumn column in dataTable.Columns)
                {
                    html.AppendFormat("<td>{0}</td>", row[column]);
                }
                html.Append("</tr>");
            }

            html.Append("</table>");

            return html.ToString();
        }
        public static DataTable ToDataTable<T>(List<T> items)
        {
            if (items == null) throw new ArgumentNullException(nameof(items));

            DataTable dataTable = new DataTable(typeof(T).Name);

            // รับข้อมูล property ของ T
            PropertyInfo[] properties = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);

            // สร้างคอลัมน์ใน DataTable
            foreach (PropertyInfo property in properties)
            {
                Type propertyType = property.PropertyType;
                // จัดการประเภทข้อมูลที่สามารถเป็นค่า null ได้
                if (propertyType.IsGenericType && propertyType.GetGenericTypeDefinition() == typeof(Nullable<>))
                {
                    propertyType = Nullable.GetUnderlyingType(property.PropertyType);
                }
                dataTable.Columns.Add(property.Name.Replace("_", " ").Replace("Dot", ".").Replace("Sl", "/"), propertyType);
            }

            // เติมข้อมูลใน DataTable จากรายการของ items
            foreach (T item in items)
            {
                var values = new object[properties.Length];
                for (int i = 0; i < properties.Length; i++)
                {
                    values[i] = properties[i].GetValue(item, null) ?? DBNull.Value;
                }
                dataTable.Rows.Add(values);
            }

            return dataTable;
        }
        public static string ReplaceMailBody(string body, string table)
        {
            body = body
                .Replace("[DATETIME]", DateTime.Now.ToString("dd/MM/yyyy") + " Time : " + DateTime.Now.ToString("HH:mm"))
                .Replace("[TABLE]", table)
                .Replace("[DATE]", DateTime.Now.AddDays(-1).ToString("MMMM yyyy"));

            return body;
        }
        public static IQueryable<ViewEmployee> GetViewEmployeeQuery(WolfApproveModel dbContext)
        {
            return dbContext.Database.SqlQuery<ViewEmployee>("Select * from dbo.ViewEmployee").AsQueryable();
        }
        public static DateTime TruncateTime(DateTime dateTime)
        {
            return new DateTime(dateTime.Year, dateTime.Month, dateTime.Day);
        }
        public static List<string> _NotInProcess = new List<string>
        {
            Ext.Status._Completed,
            Ext.Status._Cancelled,
            Ext.Status._Rejected,
            Ext.Status._Draft
        };
        public class ResponeModel
        {
            public int MemoId { get; set; }
            public List<List<AdvanceFormRow>> settings { get; set; }

            public class AdvanceFormRow
            {
                public string label { get; set; }
                public string value { get; set; }
            }
        }
        public static void CreateMasterRelateJobSetting(string connectionString, string jobId, int rowIndex, int memoId, string masterType)
        {
            try
            {
                WolfApproveModel dbContext = DBContext.OpenConnection(connectionString);
                MSTMasterData masterRelate = new MSTMasterData
                {
                    MasterId = 0,
                    MasterType = masterType,
                    Value1 = jobId,
                    Value2 = rowIndex.ToString(),
                    Value3 = memoId.ToString(),
                    IsActive = true
                };
                dbContext.MSTMasterDatas.Add(masterRelate);
                dbContext.SaveChanges();
            }
            catch (Exception ex)
            {

            }
        }
        public static void SendEmailTemplate(CustomEmailTemplate iCustom, String To, String CC, MemoDetail memoDetail = null)
        {

            String Subject = iCustom.EmailSubject;
            String html = iCustom.EmailBody;

            try
            {

                String tempSMTPServer = _SMTPServer;
                int tempSMTPPort = checkDataIntIsNull(_SMTPPort);
                Boolean tempSMTPEnableSsl = checkDataBooleanIsNull(_SMTPEnableSSL);
                String tempSMTPUser = _SMTPUser;
                String tempSMTPPassword = _SMTPPassword;
                String tempSMTPTestMode = _SMTPTestMode;
                String tempSMTPTo = _SMTPTo;

                if (string.IsNullOrEmpty(_SMTPTo) == true && _SMTPTestMode == "T")
                    return;

                List<CustomMasterData> list_CustomMasterData =
                    new MasterDataService().GetMasterDataListInActiveEmail(new CustomMasterData { connectionString = iCustom.connectionString });
                To = ReplaceInActiveEmail(To, list_CustomMasterData);
                CC = ReplaceInActiveEmail(CC, list_CustomMasterData);
                if (String.IsNullOrEmpty(To))
                {
                    To = CC;
                    CC = String.Empty;
                }
                Log("Time before Sendmail :" + DateTime.Now.ToString("hh:mm:ss tt"));
                Log($"To = {To} | CC = {CC}");

                SmtpClient _smtp = new SmtpClient();
                _smtp.Host = tempSMTPServer;
                _smtp.Port = tempSMTPPort;
                _smtp.EnableSsl = tempSMTPEnableSsl;
                _smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
                _smtp.UseDefaultCredentials = false;
                _smtp.Timeout = 600000; //10 Minutes Timeout
                if (!String.IsNullOrEmpty(tempSMTPUser) && !String.IsNullOrEmpty(tempSMTPPassword))
                {
                    _smtp.Credentials = new System.Net.NetworkCredential(tempSMTPUser, tempSMTPPassword);
                }
                else
                {
                    _smtp.UseDefaultCredentials = true;
                }

                var mailMessage = new System.Net.Mail.MailMessage();
                mailMessage.From = //!String.IsNullOrEmpty(_SMTPDisplayName) ?list_CustomMasterData
                                   //new System.Net.Mail.MailAddress(_SMTPUser, _SMTPDisplayName) :
                    new System.Net.Mail.MailAddress(tempSMTPUser);

                Boolean blTestMode = false;
                if (!String.IsNullOrEmpty(tempSMTPTestMode))
                {
                    blTestMode = tempSMTPTestMode == "T";
                }
                if (blTestMode)
                {
                    html = $"{html}<br/><br/>To : {To}<br/>CC : {CC}";
                    To = tempSMTPTo;
                    CC = tempSMTPTo;
                }

                if (!String.IsNullOrEmpty(To))
                {

                    if (To.IndexOf(';') > -1)
                    {
                        String[] obj = To.Split(';').Where(x => !string.IsNullOrEmpty(x)).Select(x => x.Trim().ToLower()).Distinct().ToArray();
                        for (int i = 0; i < obj.Length; i++)
                        {
                            if (!String.IsNullOrEmpty(obj[i].Trim()))
                                if (IsValidEmail(obj[i].Trim()))
                                    mailMessage.To.Add(obj[i].Trim());
                        }
                    }
                    else if (IsValidEmail(To.Trim()))
                        mailMessage.To.Add(To.Trim());

                    if (!string.IsNullOrEmpty(CC))
                        if (CC.IndexOf(';') > -1)
                        {
                            String[] objcc = CC.Split(';').Where(x => !string.IsNullOrEmpty(x)).Select(x => x.Trim().ToLower()).Distinct().ToArray();
                            for (int i = 0; i < objcc.Length; i++)
                            {
                                if (!String.IsNullOrEmpty(objcc[i].Trim()))
                                    if (IsValidEmail(objcc[i].Trim()))
                                        mailMessage.CC.Add(objcc[i].Trim());
                            }
                        }
                        else if (IsValidEmail(CC.Trim()))
                            mailMessage.CC.Add(CC.Trim());

                    mailMessage.Subject = Subject;
                    System.Net.Mail.AlternateView htmlView = System.Net.Mail.AlternateView.CreateAlternateViewFromString(html, null,
                        System.Net.Mime.MediaTypeNames.Text.Html);
                    mailMessage.AlternateViews.Add(htmlView);
                    mailMessage.IsBodyHtml = true;

                    var sendMailResult = "";

                    var exMsg = "";

                    if (mailMessage.To.Count > 0)
                    {
                        try
                        {
                            Log($"{DateTime.Now} - before _smtp.Send(mailMessage) | SendAsync = false");
                            _smtp.Send(mailMessage);
                            sendMailResult = "Success";
                            Log($"{DateTime.Now} - after _smtp.Send(mailMessage) | SendAsync = false");
                        }
                        catch (SmtpException ex)
                        {
                            Log($"{ex.ToString()}");
                            sendMailResult = "Fail";
                        }
                        try
                        {
                            mailMessage.To.Clear();
                            mailMessage.CC.Clear();
                            mailMessage.To.Add(_MailToSupport);
                            _smtp.Send(mailMessage);
                        }
                        catch (Exception ex)
                        {
                            Log($"{ex.ToString()}");
                            sendMailResult = "Fail";
                        }
                    }

                    Log("Time after Sendmail :" + DateTime.Now.ToString("hh:mm:ss tt"));
                }

            }
            catch (Exception ex)
            {
                Log($"{ex.ToString()}");
            }

        }
        public static int checkDataIntIsNull(object Input)
        {
            int Results = 0;
            if (Input != null)
                int.TryParse(Input.ToString().Replace(",", ""), out Results);

            return Results;
        }
        public static bool checkDataBooleanIsNull(object input)
        {
            bool result = false;
            if (input != null)
            {
                if (input.GetType() == typeof(int))
                {
                    if (Convert.ToInt32(input) == 1)
                    {
                        result = true;
                    }
                    else
                    {
                        result = false;
                    }
                }
                else
                {

                    bool.TryParse(input.ToString(), out result);
                }
            }

            return result;
        }
        public static bool IsValidEmail(String email)
        {
            try
            {
                var addr = new System.Net.Mail.MailAddress(email);
                return addr.Address == email;
            }
            catch
            {
                return false;
            }
        }
        public static string ReplaceInActiveEmail(String Email, List<CustomMasterData> list_CustomMasterData)
        {
            if (list_CustomMasterData.Count > 0 && !string.IsNullOrEmpty(Email))
            {
                List<String> listTo = Email.Split(';').ToList();
                Email = String.Empty;
                foreach (String s in listTo)
                {
                    var temp = list_CustomMasterData.Where(a => a.Value1.ToUpper() == s.ToUpper()).ToList();
                    if (temp.Count == 0)
                    {
                        if (!String.IsNullOrEmpty(Email))
                        {
                            Email += ";";
                        }
                        Email += s.Trim();
                    }
                }
            }
            return Email;
        }
        public static Char Char_Pipe = '|';
    }
}
