using System.Configuration;
using System;
using System.Collections.Generic;
using System.IO;
using Newtonsoft.Json;
using System.Data;
using System.Data.SqlClient;
using System.Net;
using DocuSign.eSign.Api;
using DocuSign.eSign.Model;
using DocuSign.eSign.Client;
using iTextSharp.text;
using iTextSharp.text.pdf;
using Document = DocuSign.eSign.Model.Document;

namespace DocuSign_AUA
{
    class DocuSignTemplate
    {
        //Check CC F(x) is ON for PROD. // Form1=Student Clerkship, Form2 = Student Faculty,Form3=Mid-Clerkship,Form4=Portfolio,Form5=Comprehensive
        static string formID = "5"; 
        //Form ID change for testing purposes 
        static int batchRunID = 20180703;
        static string ProdID = "PROD";
        static string startDate = "";        //is this really needed?
        static string formName = "";
        static string shortRotaName = "";
        static string subjectLine = "";
        //***********************   Query for records  **********************************//
        //static string strSQL = "SELECT * FROM DSStudentCoursesInformation_tbl with (nolock) WHERE CourseDateStart = '10-09-2017' order by StudentID";         //create a stored procedure?

        //static string strSQL = "SELECT * FROM DSStudentCoursesInformation_tbl where CourseDateStart = '06-11-2018' order by StudentID";

        static string strSQL = "SELECT * FROM [DSStudentCoursesInformation_tbl] where StudentID = 100886 and CourseUID = 7932";


        // Integrator Key (aka API key) is needed to authenticate your API calls.  This is an application-wide key
        public string Username { get; } = ConfigurationManager.AppSettings["username"];
      
        static readonly string StudentClerkshipEvaluationForm        = ConfigurationManager.AppSettings["StudentClerkshipEvaluationForm_TemplateID"];                 //Form 1 [Student Only]
        static readonly string StudentFacultyEvaluationForm          = ConfigurationManager.AppSettings["StudentFacultyEvaluationForm_TemplateID"];                   //Form 2 [Student Only]

        static readonly string MidClerkshipAssessmentForm            = ConfigurationManager.AppSettings["MidClerkshipAssessmentForm_TemplateID"];                     //Form 3 [Student & Preceptor]
        //static readonly string MidClerkshipAssessmentFormEditable    = ConfigurationManager.AppSettings["MidClerkshipAssessmentForm_Editable_TemplateID"];            //Form 3 [Student & Preceptor - Editable]

        static readonly string StudentPortfolioForm                  = ConfigurationManager.AppSettings["StudentPortfolio_TemplateID"];                               //Form 4 [Student to Preceptor]
        static readonly string StudentPortfolioFormEditable          = ConfigurationManager.AppSettings["StudentPortfolio_Editable_TemplateID"];                      //Form 4 [Student to Preceptor - Editable]

        static readonly string CompStudentClerkshipAssessmentForm           = ConfigurationManager.AppSettings["ComprehensiveStudentClerkshipAssessmentForm_TemplateID"];               //Form 5 [Preceptor & DME]
        static readonly string CompStudentClerkshipAssessmentFormEditable   = ConfigurationManager.AppSettings["ComprehensiveStudentClerkshipAssessmentForm_Editable_TemplateID"];      //Form 5 [Preceptor & DME Editable]

        private static string templateNameCSA;
        private static string templateNameStudentPF;
        
        //////////////////////////////////////////////////////////
        // Main()
        //////////////////////////////////////////////////////////
        static void Main(string[] args)
        {
            SetHeader();

            //GetBatchRunID();

            GetStudentRotationInfo(batchRunID, ProdID);

            Console.WriteLine();
            Console.WriteLine("***************************************************************");
            Console.WriteLine("***************************************************************");

            switch (formID)
            {
                case "1":
                    formName = " Student Clerkship Evaluation";
                    break;
                case "2":
                    formName = " Student Faculty Evaluation";
                    break;
                case "3":
                    formName = " Mid-Clerkship Formative Student Assessment";
                    break;
                case "4":
                    formName = " Student Portfolio";
                    break;
                case "5":
                    formName = " Comprehensive Student Clerkship Assessment";
                    break;
            }

            Console.WriteLine(iCount + formName + " forms sent.");
            Console.WriteLine("***************************************************************");
            Console.WriteLine("***************************************************************");
            Console.Read();
        }

        public static void SetHeader()
        {
            //ApiClient apiClient = DocuSign.eSign.Client.Configuration.Default.ApiClient;

            string authHeader = "{\"Username\":\"" + ConfigurationManager.AppSettings["username"] + "\"," +
                                " \"Password\":\"" + ConfigurationManager.AppSettings["password"] + "\"," +
                                " \"IntegratorKey\":\"" + ConfigurationManager.AppSettings["INTEGRATOR_KEY"] + "\"}";

            DocuSign.eSign.Client.Configuration.Default.AddDefaultHeader("X-DocuSign-Authentication", authHeader);


        }

        //public static void GetBatchRunID()
        //{
        //    string storedProc = "dbo.sp_GetBatchRunID_Demo";


        //    using (SqlConnection SQLCon = new SqlConnection(ConfigurationManager.AppSettings["SMSDB_CONNSTRING"]))
        //    {
        //        try
        //        {
        //            SqlCommand SQLcmd = new SqlCommand(storedProc, SQLCon);
        //            SQLcmd.CommandType = CommandType.StoredProcedure;

        //            SQLcmd.Parameters.AddWithValue("@startDate", "05/01/17").Direction = System.Data.ParameterDirection.Input;
        //            SQLcmd.Parameters.AddWithValue("@sendDate", Convert.ToString(DateTime.Now)).Direction = System.Data.ParameterDirection.Input;
        //            SQLcmd.Parameters.Add("@runId", SqlDbType.Int).Direction = System.Data.ParameterDirection.Output;

        //            SQLcmd.Connection = SQLCon;
        //            SQLCon.Open();
        //            SQLcmd.ExecuteNonQuery();

        //            batchRunID = (int)SQLcmd.Parameters["@runId"].Value;

        //            GetStudentRotationInfo(batchRunID,"PROD");
        //        }
        //        finally
        //        {
        //            SQLCon.Close();
        //            SQLCon.Dispose();
        //        }
        //    }

        //}

        private static int iCount = 0;

        public static void GetStudentRotationInfo(int batchRunID, string Environment)
        {
            SqlConnection conn = null;
            SqlDataReader rdr = null;
           //string sendDatesProcedure = "sp_DocuSignFormSendDates_Demo";
            Student student = new Student();
            EnvelopeSummary envelopeSumm = new EnvelopeSummary();

            //Get all data required for prepopulating DocuSign templates
            //The table is created from an Excel spreadsheet 

         
            
            try
            {
                conn = new SqlConnection(ConfigurationManager.AppSettings["DocuSign_DB_CONNSTRING"]);
                conn.Open();

                var cmd = new SqlCommand(strSQL, conn)
                {
                    CommandType = CommandType.Text
                };
                cmd.Parameters.AddWithValue("@StartDate", startDate);

                rdr = cmd.ExecuteReader();

                if (Environment == ProdID)
                {
                    while (rdr.Read())
                    {
                        
                        /////////////////////////////////////////////////////////////////////////////
                        student.StudentId           = rdr["StudentID"].ToString();
                        student.StudentFirstName    = rdr["StudentFirstName"].ToString();
                        student.StudentFirstName = student.StudentFirstName.Trim();
                        student.StudentLastName     = rdr["StudentLastName"].ToString();
                        student.StudentLastName = student.StudentLastName.Trim();
                        student.StudentName = student.StudentFirstName + " " + student.StudentLastName;
                        //student.StudentName         = rdr["StudentName"].ToString();
                        student.StudentPicturePath  = rdr["StudentPicturePath"].ToString();
                        student.CourseId            = rdr["CourseUID"].ToString();
                        student.StudentEmailSchool  = rdr["StudentEmailSchool"].ToString();
                       // student.StudentEmailSchool  = "isherchan@auamed.org";  //For testing
                        student.ClinicalRotation    = rdr["CourseTitle"].ToString();

                        if (student.ClinicalRotation == "Family Medicine I/Internal Medicine I")
                        {
                            shortRotaName = student.ClinicalRotation;
                        }
                        else
                        {
                            shortRotaName = student.ClinicalRotation.Remove(student.ClinicalRotation.IndexOf('('));
                        }
                       
                        
                        student.CreditWeeks         = rdr["CourseCreditWeeks"].ToString();
                        student.ClinicalSite        = rdr["HospitalName"].ToString();

                        student.StartDate   = rdr["CourseDateStart"].ToString();
                        student.EndDate     = rdr["CourseDateEnd"].ToString();

                        student.StartDate   = DateTime.Parse(student.StartDate).ToShortDateString();
                        student.EndDate     = DateTime.Parse(student.EndDate).ToShortDateString();

                        /* Assign Preceptor & DME Info Here*/
                        //Send to Hospital Contacts if no DME or Preceptor Info is available
                        //if (formID == "3" || formID == "4" || formID == "5")

                         if (formID == "4" || formID == "5")
                            {
                            //////////////////////////////////////////////////////////////////////////////
                            if (rdr["Preceptor Email"].ToString() != "")
                            {
                                student.StudentPreceptorEmail = rdr["Preceptor Email"].ToString();
                                student.StudentPreceptorName = rdr["Preceptor Name"].ToString();
                            }

                            if (rdr["Preceptor Email"].ToString() == "" || rdr["Preceptor Email"].ToString() == "NULL" )
                            {
                                student.StudentPreceptorEmail = rdr["Preceptor Recipient"].ToString();
                                
                                //***Fill In Name Here***
                                student.StudentPreceptorName = "***PRECEPTOR: Please Fill In Name Here***";       
                            }

                            //////////////////////////////////////////////////////////////////////////////
                            if (rdr["DME Email"].ToString() != "")
                            {
                                student.StudentDMEEmail = rdr["DME Email"].ToString();
                                student.StudentDMEName = rdr["DME Name"].ToString();
                            }

                            if (rdr["DME Email"].ToString() == "" || rdr["DME Email"].ToString() == "NULL")
                            {
                                student.StudentDMEEmail = rdr["DME Recipient"].ToString();
                           
                                //***Fill In Name Here***
                               student.StudentDMEName = "***DME: Please Fill In Name Here***";
                                
                            }

                            //////////////////////////////////////////////////////////////////////////////
                            if (rdr["Preceptor Email"].ToString() == "" || rdr["DME Email"].ToString() == "")
                            {
                                templateNameCSA = CompStudentClerkshipAssessmentFormEditable;
                                templateNameStudentPF = StudentPortfolioFormEditable;
                                
                            }
                            else
                            {
                                templateNameCSA = CompStudentClerkshipAssessmentForm;
                                templateNameStudentPF = StudentPortfolioForm;
                           
                            }

                           
                        }


                        #region DUMMY DATA
                        //Dummy Data *************************************************

                        //templateNameCSA = CompStudentClerkshipAssessmentForm;
                        //templateNameStudentPF = StudentPortfolioForm;
                        //templateNameMidClerkship = MidClerkshipAssessmentForm;


                        student.StudentId = "123456";
                        student.StudentFirstName = "John";
                        student.StudentLastName = "Miller";
                        student.StudentName = "John Miller";
                        //student.StudentPicturePath = @"C:\Docusign\doctor.jpg";
                        student.CourseId = "7216";

                        student.ClinicalRotation = "Critical Care (4w Elective)";
                        student.CreditWeeks = "4";
                        student.ClinicalSite = "Manhattan General";

                        student.StudentPreceptorEmail = "Donna.green@rochesterregional.org"; //tejass@skopiq.com
                        student.StudentPreceptorName = "Dr.Preceptor";

                        student.StudentDMEEmail = "MaryJane.Schwan@rochesterregional.org";
                        student.StudentDMEName = "GME Officer";

                        student.StudentEmailSchool = "cgrenard@auamed.org";
                        #endregion

                        envelopeSumm = RequestStudentSignatureFromTemplate(student);

                        //CALL TO INSERT FORMSENT & ENVELOPE_HISTORY RECORDs
                        InsertFormAndEnvelopeInfo(student, envelopeSumm, batchRunID);

                    } //END WHILE LOOP 
                }

            } //END Try 
            finally
            {
                // close reader
                rdr?.Close();

                // close connection
                conn?.Close();
            }
        }

        public static void InsertFormAndEnvelopeInfo(Student student, EnvelopeSummary envelopeSumm, int batchRunId)
        {
            string insertProcedure = "dbo.sp_InsertFormAndEnvelopeInfo_Demo";

            using (SqlConnection SQLCon = new SqlConnection(ConfigurationManager.AppSettings["DocuSign_DB_CONNSTRING"]))
            {
                try
                {
                    SqlCommand SQLcmd = new SqlCommand(insertProcedure, SQLCon);

                    SQLcmd.CommandType = CommandType.StoredProcedure;

                    //  Create the input paramenters
                    SQLcmd.Parameters.AddWithValue("@batchRunId", batchRunID);
                    SQLcmd.Parameters.AddWithValue("@formId", formID);
                    SQLcmd.Parameters.AddWithValue("@studentId", student.StudentId);
                    SQLcmd.Parameters.AddWithValue("@courseId", student.CourseId);
                    SQLcmd.Parameters.AddWithValue("@envelopeId", envelopeSumm.EnvelopeId);
                    SQLcmd.Parameters.AddWithValue("@envelopeStatus", envelopeSumm.Status);
                    SQLcmd.Parameters.AddWithValue("@statusDate", Convert.ToDateTime(envelopeSumm.StatusDateTime));
                    SQLcmd.Parameters.AddWithValue("@receipientRoleId", 2);

                    SQLcmd.Connection = SQLCon;
                    SQLCon.Open();
                    SQLcmd.ExecuteNonQuery();
                }
                finally
                {
                    SQLCon.Close();
                    SQLCon.Dispose();
                }
            }

        }


        public static EnvelopeSummary RequestStudentSignatureFromTemplate(Student student)
        {

            //DocuSignTemplate dsTemplate = new DocuSignTemplate();

            // instantiate api client with appropriate environment (for production change to www.docusign.net/restapi)
            
            configureApiClient(ConfigurationManager.AppSettings["RestAPI_URL"]);

            //===========================================================
            // Step 1: Login()
            //===========================================================

            // call the Login() API which sets the user's baseUrl and returns their accountId
            //var accountId = loginApi(Username, password);

            //===========================================================
            // Step 2: Signature Request from Template 
            //===========================================================

            var envDef = new EnvelopeDefinition { CompositeTemplates = new List<CompositeTemplate>() };


            //===========================================================
            // Step 3: Create a CompositeTemplate, List<ServerTemplate>,List<InlineTemplate>
            //===========================================================
            CompositeTemplate ct = new CompositeTemplate
            {
                ServerTemplates = new List<ServerTemplate>(),
                InlineTemplates = new List<InlineTemplate>()
            };


            ServerTemplate st = new ServerTemplate();

            InlineTemplate it = new InlineTemplate
            {
                Recipients = new Recipients(),
                Documents = new List<Document>()
            };

            st.Sequence = "2";
            ct.ServerTemplates.Add(st);

            Recipients r = new Recipients();
            r.Signers = new List<Signer>();

            r.Agents = new List<Agent>();

            // studentSigner = CreateSigner(student.StudentEmailSchool, student.StudentName, templateRoleName, routingOrder, recipientId);

            Signer studentSigner = new Signer();
            Signer preceptorSigner = new Signer();
            Signer DMESigner = new Signer();
            Agent  HospContact = new Agent();
            Signer signer = new Signer();

            //string newReminderDelay = "";

            DateTime endDate = Convert.ToDateTime(student.EndDate);
            DateTime today = DateTime.Today;

            var days = (endDate - today).TotalDays;
            days -= 5;

            #region reminders

            //if (formID == "3")
            //{
            //    days /= 2;
            //}

            //newReminderDelay = Convert.ToString(days);

            // student.EndDate = student.CreditWeeks == "6" ? "01/06/2017" : "02/17/2017";

            //switch (student.CreditWeeks)
            //{
            //    case "4":
            //            newReminderDelay = "28";
            //            break;
            //    case "6":
            //            newReminderDelay = "42";
            //            break;
            //    case "8":
            //            newReminderDelay = "56";
            //            break;
            //    case "12":
            //            newReminderDelay = "84";
            //            break;
            //    default:
            //            newReminderDelay = "14";
            //            break;
            //}

            //Set Dynamic Notifications
            envDef.Notification = new Notification
            {
                UseAccountDefaults = "false",
                Reminders = new Reminders
                {
                    ReminderEnabled = "true",
                    ReminderFrequency = "3",
                    ReminderDelay = "12" //newReminderDelay
                },
                Expirations = new Expirations
                {
                    ExpireEnabled = "true",
                    ExpireWarn = "7",
                    ExpireAfter = "999"
                }
            };
            #endregion

            //string formName = "";

            //Assign the correct template

            //Adds all controls to List<Recipients>

            subjectLine = student.StudentLastName + ", " + student.StudentFirstName + "; " + student.StartDate + "-" + student.EndDate + "; " + shortRotaName;

            switch (formID)
            {
                case "1":
                    formName = "Student Clerkship Evaluation Form";
                    st.TemplateId = StudentClerkshipEvaluationForm;
                    envDef.EmailBlurb = "Please complete your Clerkship Evaluation form: " + subjectLine + "; " + student.ClinicalSite + " at the end of this rotation";
                    envDef.EmailSubject = subjectLine + "; Student Clerkship Eval.";
                    studentSigner = CreateSigner(student.StudentEmailSchool, student.StudentName, "Student", "1", "1");
                    r.Signers.Add(studentSigner);
                    signer = studentSigner;
                    break;

                case "2":
                    formName = "Student Faculty Evaluation Form";
                    st.TemplateId = StudentFacultyEvaluationForm;
                    envDef.EmailBlurb = "Please complete your Faculty Evaluation form: " + subjectLine + "; " + student.ClinicalSite + " at the end of this rotation";
                    envDef.EmailSubject = subjectLine + "; Student Faculty Eval.";
                    studentSigner = CreateSigner(student.StudentEmailSchool, student.StudentName, "Student", "1", "1");
                    r.Signers.Add(studentSigner);
                    signer = studentSigner;
                    break;

                case "3":
                    formName = "Mid Clerkship Assessment Form";
                    st.TemplateId = MidClerkshipAssessmentForm;
                    envDef.EmailBlurb = "Please make sure both you and your Supervising Physician (Preceptor) complete this Mid-Clerkship Assessment Form: " + subjectLine + "; " + student.ClinicalSite + " at the middle of this rotation";
                    envDef.EmailSubject = subjectLine + "; Mid-Clerkship Assessment";

                    //June 2017 - Change to send to Student (1), then Preceptor (2) -------Check Document XML Role ID
                    studentSigner = CreateSigner(student.StudentEmailSchool, student.StudentName, "Student", "1", "1");
                    r.Signers.Add(studentSigner);

                    //preceptorSigner = CreateSigner(student.StudentPreceptorEmail, student.StudentPreceptorName, "Preceptor", "2", "1");
                    //r.Signers.Add(preceptorSigner);

                    signer = studentSigner;
                    break;

                case "4":
                    formName = "Student Portfolio Form";
                    st.TemplateId = templateNameStudentPF;
                    envDef.EmailBlurb = "Please complete this Student Portfolio for: " + subjectLine + "; " + student.ClinicalSite;
                    envDef.EmailSubject = subjectLine + "; Student Portfolio";
                    
                    studentSigner = CreateSigner(student.StudentEmailSchool, student.StudentName, "Student", "1", "1");
                    r.Signers.Add(studentSigner);

                    preceptorSigner = CreateSigner(student.StudentPreceptorEmail, student.StudentPreceptorName, "Preceptor", "2", "2");
                    r.Signers.Add(preceptorSigner);

                    signer = studentSigner;

                    break;

                case "5":
                    formName = "Comprehensive Student Clerkship Assessment Form";
                    envDef.EmailBlurb = "Please complete this Comprehensive Clerkship Assessment Form - " + subjectLine + "; " + student.ClinicalSite
                                        + ", " + student.ClinicalSite;
                    envDef.EmailSubject = subjectLine + "; Comprehensive Assessment";
                    //envDef.EmailSubject = "Fixed: " + subjectLine + "; Comprehensive Assessment";
                    st.TemplateId = templateNameCSA;

                    preceptorSigner = CreateSigner(student.StudentPreceptorEmail, student.StudentPreceptorName, "Preceptor", "1", "1");
                    r.Signers.Add(preceptorSigner);

                    //After Preceptor completes and signs it goes to DME for completion.
                    DMESigner = CreateSigner(student.StudentDMEEmail, student.StudentDMEName, "DME", "2", "2");
                    r.Signers.Add(DMESigner);

                    signer = preceptorSigner;

                    break;
            }

            //CC Requested Hospital Contacts
            CCByHospitalAndForm(student, r);



            #region FormTabs
            /*
            The font type used for the information in the tab. Possible values are: 
            Arial, ArialNarrow, Calibri, CourierNew, Garamond, Georgia, Helvetica, LucidaConsole, Tahoma, TimesNewRoman, Trebuchet, and Verdana.

            The font size used for the information in the tab. 
            Possible values are: Size7, Size8, Size9, Size10, Size11, Size12, Size14, Size16, Size18, Size20, Size22, Size24, Size26, Size28, Size36, Size48, or Size72.

            The font color used for the information in the tab. Possible values are: Black, BrightBlue, BrightRed, DarkGreen, DarkRed, Gold, Green, NavyBlue, Purple, or White.
            */

            //Create a List<> for Checkboxes and add them to the collection

            //if (formID != "3")
            //{
            //    Text DEndDate = new Text
            //    {
            //        TabLabel = "DEndDate",
            //        Value = student.EndDate,
            //        DocumentId = "2",
            //    };
            //    signer.Tabs.TextTabs.Add(DEndDate);

            //    Text REndDate = new Text
            //    {
            //        TabLabel = "REndDate",
            //        Value = student.EndDate,
            //        DocumentId = "2",
            //    };

            //    signer.Tabs.TextTabs.Add(REndDate);
            //}


            signer.Tabs.CheckboxTabs = new List<Checkbox>();

            string core = "false";
            string elective = "false";
            string rotationType = "";

            rotationType = student.ClinicalRotation;
            // rotationType = rotationType.ToUpper();

            if (rotationType.ToUpper().Contains("CORE"))
            {
                core = "true";
            }
            else
            {
                elective = "true";
            }

            Checkbox chkCore = new Checkbox
            {
                TabLabel = "chkCore",
                Selected = core,
                DocumentId = "2"                     //DEBUG
            };
            signer.Tabs.CheckboxTabs.Add(chkCore);

            Checkbox chkElective = new Checkbox
            {
                TabLabel = "chkElective",
                Selected = elective,
                DocumentId = "2"                     //DEBUG
            };
            signer.Tabs.CheckboxTabs.Add(chkElective);

            Text StudentName = new Text
            {
                TabLabel = "StudentName",
                Value = student.StudentFirstName + " " + student.StudentLastName,
                DocumentId = "2",
            };
            signer.Tabs.TextTabs.Add(StudentName);

            Text ClinicalRotationName = new Text
            {
                TabLabel = "ClinicalRotationName",
                Value = student.ClinicalRotation,
                DocumentId = "2"
            };

            signer.Tabs.TextTabs.Add(ClinicalRotationName);

            Text ClinicalRotationSite = new Text
            {
                TabLabel = "ClinicalRotationSite",
                Value = student.ClinicalSite,
                DocumentId = "2",
              //  ConcealValueOnDocument = "true"
            };
            signer.Tabs.TextTabs.Add(ClinicalRotationSite);

            Text StartDate = new Text
            {
                TabLabel = "StartDate",
                Value = student.StartDate,
                DocumentId = "2"
            };
            signer.Tabs.TextTabs.Add(StartDate);

            Text EndDate = new Text
            {
                TabLabel = "EndDate",
                Value = student.EndDate,
                DocumentId = "2"
            };
            signer.Tabs.TextTabs.Add(EndDate);

            Text txtStudentID = new Text
            {
                TabLabel = "txtStudentID",
                Value = student.StudentId,
                DocumentId = "2",
            };
            signer.Tabs.TextTabs.Add(txtStudentID);

            //Text txtMSPEComments = new Text
            //{
            //    TabLabel = "txtMSPEComments",
            //    MaxLength = 4000,
            //    DocumentId = "2",
            //};
            //signer.Tabs.TextTabs.Add(txtMSPEComments);


            #endregion FormTabs


            it.Recipients = r;  //might need to change
            it.Sequence = "1";


            Document doc = new Document
            {
                Name = formName,
                DocumentBase64 = CreatePersonalizedForm(student, Convert.ToInt32(formID)),
                DocumentId = "2"
            };
            //why #2?

            it.Documents.Add(doc);

            ct.InlineTemplates.Add(it);
            envDef.CompositeTemplates.Add(ct);
            
            
            //Setting Status to sent sends the email
            envDef.Status = "sent";

            
            // |EnvelopesApi| contains methods related to creating and sending Envelopes (aka signature requests)
            EnvelopesApi envelopesApi = new EnvelopesApi();
            EnvelopeSummary envelopeSummary = envelopesApi.CreateEnvelope(ConfigurationManager.AppSettings["accountId"], envDef);


            //UPDATE Form_Sent & INSERT EnvelopeHistory
            var envId = envelopeSummary.EnvelopeId;
            //Console.WriteLine("EnvelopeId: " + envId + "\n");
            //

            // print the JSON response

            iCount++;

            Console.WriteLine(iCount + ". EnvelopeSummary:\n{0}", JsonConvert.SerializeObject(envelopeSummary));

            //write a log file method or store to DB to show file was sent

            return envelopeSummary;

        }
        // end requestSignatureFromTemplateTest()


        private static void CCByHospitalAndForm(Student student, Recipients r)
        {
            switch (student.ClinicalSite)
            {
                case "Richmond University Medical Center":
                    if (formID == "5")
                    {
                        var ccRecipient = new CarbonCopy
                        {
                            Email = "dmangold@rumsci.org",
                            Note = "You were CC'd on the completed assessment",
                            Name = "Carbon Copy Recipient",
                            RecipientId = "3",
                            RoutingOrder = "3"
                        };
                        var ccRecipient2 = new CarbonCopy
                        {
                            Email = "auastudenteval@rumcsi.org",
                            Note = "You were CC'd on the completed assessment",
                            Name = "Carbon Copy Recipient",
                            RecipientId = "3",
                            RoutingOrder = "3"
                        };
                        r.CarbonCopies = new List<CarbonCopy> { ccRecipient, ccRecipient2 };
                    }
                    if (formID == "5" && student.ClinicalRotation.Contains("Perioperative"))
                    {
                        var ccRecipient = new CarbonCopy
                        {
                            Email = "riso.janice@gmail.com",
                            Note = "You were CC'd",
                            Name = "Carbon Copy Recipient",
                            RecipientId = "3",
                            RoutingOrder = "3"
                        };
                        var ccRecipient2 = new CarbonCopy
                        {
                            Email = "dmangold@rumsci.org",
                            Note = "You were CC'd on the completed assessment",
                            Name = "Carbon Copy Recipient",
                            RecipientId = "3",
                            RoutingOrder = "3"
                        };
                        var ccRecipient3 = new CarbonCopy
                        {
                            Email = "auastudenteval@rumcsi.org",
                            Note = "You were CC'd on the completed assessment",
                            Name = "Carbon Copy Recipient",
                            RecipientId = "3",
                            RoutingOrder = "3"
                        };
                        r.CarbonCopies = new List<CarbonCopy> { ccRecipient, ccRecipient2, ccRecipient3 };
                    }
                    break;


                case "Interfaith Medical Center":
                    if (formID == "5")
                    {
                        var ccRecipient = new CarbonCopy
                        {
                            Email = "ktheodore@interfaithmedical.org",
                            Note = "You were CC'd",
                            Name = "Carbon Copy Recipient",
                            RecipientId = "3",
                            RoutingOrder = "2"
                        };
                        r.CarbonCopies = new List<CarbonCopy> { ccRecipient };
                    }

                    break;

                case "University of Maryland Medical Center-Midtown Campus":
                    if (formID == "5")
                    {
                        var ccRecipient = new CarbonCopy
                        {
                            Email = "PIncoom@umm.edu",
                            Note = "You were CC'd",
                            Name = "Carbon Copy Recipient",
                            RecipientId = "3",
                            RoutingOrder = "3"
                        };
                        var ccRecipient2 = new CarbonCopy
                        {
                            Email = "GBrandon@umm.edu",
                            Note = "You were CC'd",
                            Name = "Carbon Copy Recipient",
                            RecipientId = "3",
                            RoutingOrder = "3"
                        };
                        r.CarbonCopies = new List<CarbonCopy> { ccRecipient, ccRecipient2 };
                    }

                    break;

                case "University of Maryland Medical Center-Mi":
                    if (formID == "5")
                    {
                        var ccRecipient = new CarbonCopy
                        {
                            Email = "PIncoom@umm.edu",
                            Note = "You were CC'd",
                            Name = "Carbon Copy Recipient",
                            RecipientId = "3",
                            RoutingOrder = "3"
                        };
                        var ccRecipient2 = new CarbonCopy
                        {
                            Email = "GBrandon@umm.edu",
                            Note = "You were CC'd",
                            Name = "Carbon Copy Recipient",
                            RecipientId = "3",
                            RoutingOrder = "3"
                        };
                        r.CarbonCopies = new List<CarbonCopy> { ccRecipient, ccRecipient2 };
                    }

                    break;

                case "Georgia Regional Hospital- Atlanta":
                    if (formID == "5")
                    {
                        var ccRecipient = new CarbonCopy
                        {
                            Email = "BStubbs@mdcsa.net",
                            Note = "You were CC'd",
                            Name = "Carbon Copy Recipient",
                            RecipientId = "3",
                            RoutingOrder = "3"
                        };
                        r.CarbonCopies = new List<CarbonCopy> { ccRecipient };
                    }

                    break;

                case "Southside Hospital":
                    
                    if (formID == "5" && student.ClinicalRotation.Contains("Obstetrics"))
                    {
                        var ccRecipient = new CarbonCopy
                        {
                            Email = "cdolce@mdcsa.net",
                            Note = "You were CC'd",
                            Name = "Carbon Copy Recipient",
                            RecipientId = "3",
                            RoutingOrder = "3"
                        };
                        r.CarbonCopies = new List<CarbonCopy> { ccRecipient };
                    }
                    break;

                case "Harbor Hospital Center":
                    if (formID == "4" || formID == "5")
                    {
                        var ccRecipient = new CarbonCopy
                        {
                            Email = "Terry.Kus@medstar.net",
                            Note = "You were CC'd",
                            Name = "Carbon Copy Recipient",
                            RecipientId = "2",
                            RoutingOrder = "2"
                        };

                        var ccRecipient2 = new CarbonCopy
                        {
                            Email = "Denise.McCumbers@medstar.net",
                            Note = "You were CC'd",
                            Name = "Carbon Copy Recipient",
                            RecipientId = "2",
                            RoutingOrder = "2"
                        };
                        r.CarbonCopies = new List<CarbonCopy> { ccRecipient, ccRecipient2 };
                    }
                    break;

                case "Northside Medical Center":
                    if (student.ClinicalRotation.Contains("Surgery") && (formID == "4" || formID == "5"))
                    {
                        var ccRecipient = new CarbonCopy
                        {
                            Email = "MaryAnn.Evanchick@steward.org",
                            Note = "You were CC'd",
                            Name = "Carbon Copy Recipient",
                            RecipientId = "2",
                            RoutingOrder = "2"
                        };

                        r.CarbonCopies = new List<CarbonCopy> { ccRecipient };
                    }
                    if (student.ClinicalRotation.Contains("Infectious") && (formID == "4" || formID == "5"))
                    {
                        var ccRecipient = new CarbonCopy
                        {
                            Email = "Marilyn.Mangino@steward.org",
                            Note = "You were CC'd",
                            Name = "Carbon Copy Recipient",
                            RecipientId = "2",
                            RoutingOrder = "2"
                        };

                        r.CarbonCopies = new List<CarbonCopy> { ccRecipient };
                    }

                    if (student.ClinicalRotation.Contains("MICU") && (formID == "4" || formID == "5"))
                    {
                        var ccRecipient = new CarbonCopy
                        {
                            Email = "Marilyn.Mangino@steward.org",
                            Note = "You were CC'd",
                            Name = "Carbon Copy Recipient",
                            RecipientId = "2",
                            RoutingOrder = "2"
                        };

                        r.CarbonCopies = new List<CarbonCopy> { ccRecipient };
                    }

                    if (student.ClinicalRotation.Contains("Internal Medicine (12w Core)") && (formID == "4" || formID == "5"))
                    {
                        var ccRecipient = new CarbonCopy
                        {
                            Email = "Marilyn.Mangino@steward.org",
                            Note = "You were CC'd",
                            Name = "Carbon Copy Recipient",
                            RecipientId = "2",
                            RoutingOrder = "2"
                        };

                        r.CarbonCopies = new List<CarbonCopy> { ccRecipient };
                    }
                    break;


                case "Wyckoff Heights Medical Center":

                    //if (student.ClinicalRotation.Contains("Surgery") && formID == "4")
                    //{
                    //    var ccRecipient = new CarbonCopy
                    //    {
                    //        Email = "atrzpis@wyckoffhospital.org",
                    //        Note = "You were Carbon Copied on this email, for your record.",
                    //        Name = "Carbon Copy Recipient",
                    //        RecipientId = "3",
                    //        RoutingOrder = "2"
                    //    };

                    //    r.CarbonCopies = new List<CarbonCopy> { ccRecipient };
                    //}

                    if (formID == "5")
                    {
                        var ccRecipient = new CarbonCopy
                        {
                            Email = "CRodriguez@wyckoffhospital.org",
                            Note = "You were Carbon Copied on this email, for your record.",
                            Name = "Carbon Copy Recipient",
                            RecipientId = "3",
                            RoutingOrder = "3"
                        };
                        r.CarbonCopies = new List<CarbonCopy> { ccRecipient };
                    }
                    //else
                    //{
                    //    var ccRecipient = new CarbonCopy
                    //    {
                    //        Email = "CRodriguez@wyckoffhospital.org",
                    //        Note = "You were Carbon Copied on this email, for your record.",
                    //        Name = "Carbon Copy Recipient",
                    //        RecipientId = "3",
                    //        RoutingOrder = "3"
                    //    };

                    //    var ccRecipient2 = new CarbonCopy
                    //    {
                    //        Email = "ACorrea@wyckoffhospital.org",
                    //        Note = "You were Carbon Copied on this email, for your record.",
                    //        Name = "Carbon Copy Recipient",
                    //        RecipientId = "3",
                    //        RoutingOrder = "3"
                    //    };

                    //    r.CarbonCopies = new List<CarbonCopy> { ccRecipient, ccRecipient2 };
                    //}

                    break;

                case "Sinai Hospital - Baltimore":
                    if (formID == "5")
                    {
                        var ccRecipient = new CarbonCopy
                        {
                            Email = "Cdallas@lifebridgehealth.org",
                            Note = "You were CC'd",
                            Name = "Carbon Copy Recipient",
                            RecipientId = "3",
                            RoutingOrder = "3"
                        };
                        r.CarbonCopies = new List<CarbonCopy> { ccRecipient };
                    }

                    break;

                case "University of Miami":

                    if (formID == "4" && student.ClinicalRotation.Contains("ECG and Bedside Skills"))
                    {
                        var ccRecipient = new CarbonCopy
                        {
                            Email = "isherchan@AUAMED.ORG",
                            Note = "You were Carbon Copied on this email, to remove Preceptor Recipient, as agreed upon.",
                            Name = "Carbon Copy Recipient",
                            RecipientId = "3",
                            RoutingOrder = "1"
                        };

                        r.CarbonCopies = new List<CarbonCopy> { ccRecipient };
                    }
                    break;

                case "Kingsbrook Jewish Medical Center":
                    if (formID == "5")
                    {
                        var ccRecipient = new CarbonCopy
                        {
                            Email = "MPaz@kingsbrook.org",
                            Note = "You were CC'd",
                            Name = "Carbon Copy Recipient",
                            RecipientId = "2",
                            RoutingOrder = "2"
                        };
                        r.CarbonCopies = new List<CarbonCopy> { ccRecipient };
                    }

                    break;

                default:
                    break;
            }
        }


        //public CarbonCopy CCRecipient(string email, string name, string message, string recipientID, string routingOrder, Recipients r)
        //{
        //        var ccRecipient = new CarbonCopy
        //        {
        //            Email = email,
        //            Note = message,
        //            Name = name,
        //            RecipientId = recipientID,
        //            RoutingOrder = routingOrder
        //        };
        //    r.CarbonCopies = new List<CarbonCopy> { ccRecipient };

        //    return ccRecipient;

        //}


        public static string CreatePersonalizedForm(Student s, int pdfFormId)
        {
            string currentForm = "";
            string outputForm = "";
            string imagePath = "";


               //imagePath = @"c:\docusign\doctor.jpg"; //--FOR TESTING

               if (DateTime.Parse(s.StartDate) >= DateTime.Parse("08-01-2017"))
               {
                    imagePath = @"w:\data\aua\StudentPhotos\" + s.StudentPicturePath + ".jpg";
               }
               else
               {
                    imagePath = @"w:\data\aua\StudentPhotos\" + s.StudentPicturePath + "_" + s.StudentId + ".jpg";
               }

            //Rotations prior to August
            // imagePath = @"w:\data\aua\StudentPhotos\" + s.StudentPicturePath + "_" + s.StudentId + ".jpg";

            //August rotations and onward
            // imagePath = @"w:\data\aua\StudentPhotos\" + s.StudentPicturePath + ".jpg";

            //imagePath.Replace(" ", "");

            imagePath = @"c:\docusign\doctor.jpg"; //--FOR TESTING


            switch (pdfFormId)
            {
                case 1:
                    currentForm = @"C:\Docusign_2019\AUA_Forms\StudentClerkshipEvaluationForm.pdf";
                    outputForm = @"C:\Docusign_2019\AUA_Forms\AUA_Forms\Output_Forms\StudentClerkshipEvaluationForm_" + s.StudentId + ".pdf";
                    break;
                case 2:
                    currentForm = @"C:\Docusign_2019\AUA_Forms\StudentFacultyEvaluationForm.pdf";
                    outputForm = @"C:\Docusign_2019\AUA_Forms\AUA_Forms\Output_Forms\StudentFacultyEvaluationForm_" + s.StudentId + ".pdf";
                    break;
                case 3:
                    currentForm = @"C:\Docusign_2019\AUA_Forms\MidClerkshipAssessmentForm.pdf";
                    outputForm = @"C:\Docusign_2019\AUA_Forms\AUA_Forms\Output_Forms\MidClerkshipAssessmentForm_" + s.StudentId + ".pdf";
                    break;
                case 4:
                    currentForm = @"C:\Docusign_2019\AUA_Forms\StudentPortfolioForm.pdf";
                    outputForm = @"C:\Docusign_2019\AUA_Forms\AUA_Forms\Output_Forms\StudentPortfolioForm_" + s.StudentId + ".pdf";
                    break;
                case 5:
                    currentForm = @"C:\Docusign_2019\AUA_Forms\ComprehensiveStudentClerkshipAssessmentForm.pdf";
                    outputForm = @"C:\Docusign_2019\AUA_Forms\AUA_Forms\Output_Forms\ComprehensiveStudentClerkshipAssessmentForm_" + s.StudentId + ".pdf";
                    break;

            }
            
            //Step 1: Open the PDF form to be sent
            using (Stream inputPdfStream = new FileStream(currentForm, FileMode.Open, FileAccess.Read, FileShare.Read))

            //Step 2: Identify the Student Image to be inserted
            using (Stream inputImageStream = new FileStream(imagePath, FileMode.Open, FileAccess.Read, FileShare.Read))

            //Step 3: Create a new instance of the form with the Student's Picture Inserted
            using (Stream outputPdfStream = new FileStream(outputForm, FileMode.Create, FileAccess.Write, FileShare.ReadWrite))
            {
                PdfReader reader = new PdfReader(inputPdfStream);
                PdfStamper stamper = new PdfStamper(reader, outputPdfStream);
                PdfContentByte pdfContentByte = stamper.GetOverContent(1);

                Image image = Image.GetInstance(inputImageStream);
                image.ScaleAbsoluteHeight(150);
                image.ScaleAbsoluteWidth(150);

                image.Border = 1;
                image.BorderWidth = 3;
                image.BorderWidthRight = 3;
                image.BorderWidthLeft = 3;
                image.BorderWidthBottom = 3;

                image.SetAbsolutePosition(34, 504);
                pdfContentByte.AddImage(image);
                stamper.Close();

                byte[] bytes = File.ReadAllBytes(outputForm);

                //delete output form
                //outputForm.

                string returnPdf = Convert.ToBase64String(bytes);

                return returnPdf;

            }

        }

        private static Signer CreateSigner(string recipientEmail, string recipientName, string templateRoleName, string routingOrder, string recipientId)
        {
            Signer signer = new Signer
            {
                Email = recipientEmail,  //Preceptor Email will be blank at first
                Name = recipientName,     //Preceptor Name will be blank at first
                RoleName = templateRoleName,
                RoutingOrder = routingOrder,
                RecipientId = recipientId,
                Tabs = new Tabs { TextTabs = new List<Text>() }
            };

            //Only create once for Recipient/Signer
            //Each recip gets their own tabs
            return signer;

        }

        //private static Agent CreateAgent(string recipientEmail, string recipientName, string templateRoleName, string routingOrder, string recipientId)
        //{
        //    Agent agent = new Agent()
        //    {
        //        Email = recipientEmail,  //Preceptor Email will be blank at first
        //        Name = recipientName,     //Preceptor Name will be blank at first
        //        RoleName = templateRoleName,
        //        RoutingOrder = routingOrder,
        //        RecipientId = recipientId,
        //        CanEditRecipientEmails = "false",
        //        CanEditRecipientNames = "false"
        //    };
            
        //    return agent;
        //}




        public static void configureApiClient(string basePath)
        {
            // instantiate a new api client
            ApiClient apiClient = new ApiClient(basePath);

            // set client in global config so we don't need to pass it to each API object.
            DocuSign.eSign.Client.Configuration.Default.ApiClient = apiClient;

            //Update 6/21/2018 - Force TLS > 1.0 - Ideal is TLS1.2
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;
        }

    } // end class

} //end namespace




