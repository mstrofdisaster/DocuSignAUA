using System.Configuration;
using System;
using System.Collections.Generic;
using System.IO;
using Newtonsoft.Json;
using System.Data;
using System.Data.SqlClient;
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
        static string startDate = "10/10/2016";
        static string formID = "1";
        static string formID2 = "2";                                                                                     //Form ID change for testing purposes
        static int batchRunID = 0;

        // Integrator Key (aka API key) is needed to authenticate your API calls.  This is an application-wide key
        readonly string INTEGRATOR_KEY = ConfigurationManager.AppSettings["INTEGRATOR_KEY"];
        public string Username { get; } = ConfigurationManager.AppSettings["username"];
        readonly string password = ConfigurationManager.AppSettings["password"];

        static string StudentClerkshipEvaluationForm = ConfigurationManager.AppSettings["StudentClerkshipEvaluationForm_TemplateID"];                              //Form 1 [Student Only]
        static string StudentClerkshipEvaluationDocumentId = ConfigurationManager.AppSettings["StudentClerkshipEvaluationForm_DocumentId"];                        //Form 1 [Student Only]

        static string StudentFacultyEvaluationForm = ConfigurationManager.AppSettings["StudentFacultyEvaluationForm_TemplateID"];                                  //Form 2 [Student Only]
        static string MidClerkshipAssessmentForm = ConfigurationManager.AppSettings["MidClerkshipAssessmentForm_TemplateID"];                                      //Form 3 [Student & Preceptor]
        static string StudentPortfolioForm = ConfigurationManager.AppSettings["StudentPortfolio_TemplateID"];                                                      //Form 4 [Student Only]
        static string ComprehensiveStudentClerkshipAssessmentForm = ConfigurationManager.AppSettings["ComprehensiveStudentClerkshipAssessmentForm_TemplateID"];    //Form 5 [Preceptor & DME]


        //////////////////////////////////////////////////////////
        // Main()
        //////////////////////////////////////////////////////////

        //CLASS SCOPE
        public static int batchRunId;


        public static void Main(string[] args)
        {
            SetHeader();

            batchRunId = GetBatchRunID();

            GetStudentRotationInfo();

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

        public static void GetStudentRotationInfo()
        {
            SqlConnection conn = null;
            SqlDataReader rdr = null;
            string sendDatesProcedure = "sp_DocuSignFormSendDates_Demo";
            Student student = new Student();
            EnvelopeSummary envelopeSumm = new EnvelopeSummary();

            try
            {
                conn = new SqlConnection(ConfigurationManager.AppSettings["SMSDB_CONNSTRING2"]);
                conn.Open();

                var cmd = new SqlCommand(sendDatesProcedure, conn)
                {
                    CommandType = CommandType.StoredProcedure
                };
                cmd.Parameters.AddWithValue("@StartDate", startDate);

                rdr = cmd.ExecuteReader();

                while (rdr.Read())
                {
                    student.StudentId = rdr["StudentID"].ToString();
                    student.StudentFirstName = rdr["StudentFirstName"].ToString();
                    student.StudentLastName = rdr["StudentLastName"].ToString();
                    student.StudentName = rdr["StudentName"].ToString();
                    student.StudentPicturePath = rdr["StudentPicturePath"].ToString();
                    student.CourseId = rdr["CourseUID"].ToString();

                    //FOR QA - Uncomment for PROD
                    //student.StudentEmailSchool  = rdr["StudentEmailSchool"].ToString();
                    //****************************************************************************************************
                    //FOR Testing - Comment out for PROD
                    student.StudentEmailSchool = "cgrenard@auamed.org";
                    //****************************************************************************************************

                    student.CourseId = rdr["CourseUID"].ToString();
                    student.ClinicalRotation = rdr["CourseTitle"].ToString();
                    student.CreditWeeks = rdr["CourseCreditWeeks"].ToString();
                    student.ClinicalSite = rdr["HospitalName"].ToString();
                    student.StartDate = rdr["CourseDateStart"].ToString();
                    student.EndDate = rdr["CourseDateEnd"].ToString();


                    envelopeSumm = RequestStudentSignatureFromTemplate(student);

                    //CALL TO INSERT FORMSENT & ENVELOPE_HISTORY RECORDs
                    InsertFormAndEnvelopeInfo(student, envelopeSumm, batchRunID);

                } //END WHILE LOOP

            } //END Try 
            finally
            {
                // close reader
                rdr?.Close();

                // close connection
                conn?.Close();
            }
        }

        public static EnvelopeSummary RequestStudentSignatureFromTemplate(Student student)
        {

            // instantiate api client with appropriate environment (for production change to www.docusign.net/restapi)
            //configureApiClient("https://na2.docusign.net/restapi");
            configureApiClient("https://demo.docusign.net/restapi");


            // var envID = envDef.EnvelopeId;
            //AUA --> DME --> ?PRECEPTOR?

            var envDef = new EnvelopeDefinition { CompositeTemplates = new List<CompositeTemplate>() };

            // studentSigner = CreateSigner(student.StudentEmailSchool, student.StudentName, templateRoleName, routingOrder, recipientId);

            Signer studentSigner = new Signer();
            Signer preceptorSigner = new Signer();
            Signer DMESigner = new Signer();
            Signer signer = new Signer();

            Agent agent = new Agent();
            CarbonCopy DME_CC = new CarbonCopy();


            SetEnvelopeReminders(student, envDef);


            string formName = "";

            //Assign the correct template
            //Adds all controls to List<Recipients>

            //Make a list of forms to 
            var formsList = new List<string>();

            // if Student
            formsList.Add("1");
            //formsList.Add("2");
            //formsList.Add("4");

            // if DME/Preceptor
            //formsList.Add("3");
            //formsList.Add("5");

            int i = 1;

            foreach (string form in formsList)
            {
                CompositeTemplate ct = new CompositeTemplate
                {
                    ServerTemplates = new List<ServerTemplate>(),
                    InlineTemplates = new List<InlineTemplate>()
                };


                ServerTemplate st = new ServerTemplate();  //??

                InlineTemplate it = new InlineTemplate
                {
                    Recipients = new Recipients(),
                    Documents = new List<Document>()
                };

                st.Sequence = i.ToString();

                ct.ServerTemplates.Add(st);

                Recipients r = new Recipients(); 
                r.Signers = new List<Signer>();

                string documentId = string.Empty;

                switch (form)
                {
                    case "1":
                        formName = "Student Clerkship Evaluation Form";
                        st.TemplateId = StudentClerkshipEvaluationForm;
                        documentId = StudentClerkshipEvaluationDocumentId;
                        envDef.EmailBlurb = "Please complete your Clerkship Evaluation on the last Friday of this rotation.";
                        envDef.EmailSubject = "AUA - Student Clerkship Evaluation Form";
                        studentSigner = CreateSigner(student.StudentEmailSchool, student.StudentName, "Student", "1", "1");
                        r.Signers.Add(studentSigner);
                        signer = studentSigner;
                        break;

                    case "2":
                        formName = "Student Faculty Evaluation Form";
                        st.TemplateId = StudentFacultyEvaluationForm;
                        documentId = "";
                        envDef.EmailBlurb = "Please complete your Faculty Evaluation on the last Friday of this rotation.";
                        envDef.EmailSubject = "AUA - Student Faculty Evaluation Form";
                        studentSigner = CreateSigner(student.StudentEmailSchool, student.StudentName, "Student", "1", "1");
                        r.Signers.Add(studentSigner);
                        signer = studentSigner;
                        break;

                    case "3":
                        formName = "Mid Clerkship Assessment Form";
                        st.TemplateId = MidClerkshipAssessmentForm;
                        documentId = "";
                        envDef.EmailBlurb = "Please complete this Mid-Clerkship Assessment Form at the middle of the rotation.";
                        envDef.EmailSubject = "AUA - Mid Clerkship Assessment Form";


                        //DME FIRST (1)
                        //  DMESigner = CreateSigner("cgrenard@auamed.org", "Chris Grenard", "DME", "2", "3");
                        //   r.Signers.Add(DMESigner);
                        //  CarbonCopy carbonCopy = new CarbonCopy();
                        DME_CC = CreateCarbonCopy("cgrenard@auamed.org", "Chris Grenard", "DME", "2", "3");
                        if (r.CarbonCopies == null)
                        {
                            r.CarbonCopies = new List<CarbonCopy>();
                        }
                        r.CarbonCopies.Add(DME_CC);


                        //BEGIN PRECEPTOR CHECK
                        //If we don't know the preceptor then create agent.

                        //AGENT SECOND (2)
                        agent = CreateAgent("cgrenard@gmail.com", "Chris G", "DMEAgent", "1", "4");

                        if (r.Agents == null)
                        {
                            r.Agents = new List<Agent>();
                        }
                        r.Agents.Add(agent);



                        //else Create a precpetor using their ACTUAL INFO***
                        //PRECEP (3)
                        preceptorSigner = CreateSigner("", "", "Preceptor", "3", "1");
                        r.Signers.Add(preceptorSigner);

                        //end if


                        //Student (4)
                        studentSigner = CreateSigner(student.StudentEmailSchool, student.StudentName, "Student", "4", "2");
                        r.Signers.Add(studentSigner);

                        // Commented out b/c I don't need to prepopulate
                        signer = preceptorSigner;
                        break;

                    case "4":
                        formName = "Student Portfolio Form";
                        st.TemplateId = StudentPortfolioForm;
                        documentId = "";
                        envDef.EmailBlurb = "Please complete your Student Portfolio prior to the end of this rotation.";
                        envDef.EmailSubject = "AUA - Student Portfolio";
                        studentSigner = CreateSigner(student.StudentEmailSchool, student.StudentName, "Student", "1", "1");
                        r.Signers.Add(studentSigner);

                        signer = studentSigner;

                        break;

                    case "5":
                        formName = "Comprehensive Student Clerkship Assessment Form";
                        st.TemplateId = ComprehensiveStudentClerkshipAssessmentForm;
                        documentId = "";
                        envDef.EmailBlurb = "Please complete this Comprehensive Clerkship Assessment Form for each respective student.";
                        envDef.EmailSubject = "AUA - Comprehensive Student Clerkship Assessment Form";
                        preceptorSigner = CreateSigner("cgrenard@auamed.org", "Dr. Gometi", "Preceptor", "1", "2");
                        DMESigner = CreateSigner("cgrenard@auamed.org", "Janes Rice, PHd.", "DME", "2", "3");
                        r.Signers.Add(preceptorSigner);
                        r.Signers.Add(DMESigner);
                        signer = preceptorSigner;
                        break;

                    default:
                        break;
                }

                it.Recipients = r;  //might need to change

                it.Sequence = i.ToString();


                Document doc = new Document
                {
                    Name = formName,
                    DocumentBase64 = CreatePersonalizedForm(student, Convert.ToInt32(form)),
                    DocumentId = documentId
                };
                //why #2?

                it.Documents.Add(doc);

                ct.InlineTemplates.Add(it);
                envDef.CompositeTemplates.Add(ct);

                #region FormTabs
                /*
                The font type used for the information in the tab. Possible values are: 
                Arial, ArialNarrow, Calibri, CourierNew, Garamond, Georgia, Helvetica, LucidaConsole, Tahoma, TimesNewRoman, Trebuchet, and Verdana.

                The font size used for the information in the tab. 
                Possible values are: Size7, Size8, Size9, Size10, Size11, Size12, Size14, Size16, Size18, Size20, Size22, Size24, Size26, Size28, Size36, Size48, or Size72.

                The font color used for the information in the tab. Possible values are: Black, BrightBlue, BrightRed, DarkGreen, DarkRed, Gold, Green, NavyBlue, Purple, or White.
                */

                //Create a List<> for Checkboxes and add them to the collection

                //REFACTOR SOMEDAY INTO NEW METHOD SO WE CAN DYNAMICALLY ASSIGN TABS PER ROLE TYPE...
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
                    Selected = core
                };
                signer.Tabs.CheckboxTabs.Add(chkCore);

                Checkbox chkElective = new Checkbox
                {
                    TabLabel = "chkElective",
                    Selected = elective
                };
                signer.Tabs.CheckboxTabs.Add(chkElective);

                Text StudentName = new Text
                {
                    TabLabel = "StudentName",
                    Value = student.StudentFirstName + " " + student.StudentLastName
                };
                signer.Tabs.TextTabs.Add(StudentName);


                Text CourseId = new Text
                {
                    TabLabel = "CourseId",
                    Value = student.CourseId,
                    FontColor = "White"

                };
                signer.Tabs.TextTabs.Add(CourseId);

                Text ClinicalRotationName = new Text
                {
                    TabLabel = "ClinicalRotationName",
                    Value = student.ClinicalRotation
                };

                signer.Tabs.TextTabs.Add(ClinicalRotationName);

                Text ClinicalRotationSite = new Text
                {
                    TabLabel = "ClinicalRotationSite",
                    Value = student.ClinicalSite
                };
                signer.Tabs.TextTabs.Add(ClinicalRotationSite);

                Text StartDate = new Text
                {
                    TabLabel = "StartDate",
                    Value = student.StartDate
                };
                signer.Tabs.TextTabs.Add(StartDate);

                Text EndDate = new Text
                {
                    TabLabel = "EndDate",
                    Value = student.EndDate
                };
                signer.Tabs.TextTabs.Add(EndDate);

                Text txtStudentID = new Text
                {
                    TabLabel = "txtStudentID",
                    Value = student.StudentId
                };
                signer.Tabs.TextTabs.Add(txtStudentID);


                #endregion FormTabs


                i++;
            }



           


            //Setting Status to sent sends the email
            envDef.Status = "sent";

            // |EnvelopesApi| contains methods related to creating and sending Envelopes (aka signature requests)
            EnvelopesApi envelopesApi = new EnvelopesApi();
            EnvelopeSummary envelopeSummary = envelopesApi.CreateEnvelope(ConfigurationManager.AppSettings["accountId"], envDef);


            // print the JSON response
            Console.WriteLine("EnvelopeSummary:\n{0}", JsonConvert.SerializeObject(envelopeSummary));

            //write a log file method or store to DB to show file was sent

            return envelopeSummary;

        }

        public static void InsertFormAndEnvelopeInfo(Student student, EnvelopeSummary envelopeSumm, int batchRunId)
        {
            string insertProcedure = "dbo.sp_InsertFormAndEnvelopeInfo_Demo";

            using (SqlConnection SQLCon = new SqlConnection(ConfigurationManager.AppSettings["DocuSign_DB_CONNSTRING"]))
            {
                try
                {
                    SqlCommand SQLcmd = new SqlCommand(insertProcedure, SQLCon)
                    {
                        CommandType = CommandType.StoredProcedure
                    };


                    //  Create the input paramenters
                    SQLcmd.Parameters.AddWithValue("@batchRunId", batchRunID);
                    SQLcmd.Parameters.AddWithValue("@formId", formID);
                    SQLcmd.Parameters.AddWithValue("@studentId", student.StudentId);
                    SQLcmd.Parameters.AddWithValue("@courseId", student.CourseId);
                    SQLcmd.Parameters.AddWithValue("@envelopeId", envelopeSumm.EnvelopeId);
                    SQLcmd.Parameters.AddWithValue("@envelopeStatus", envelopeSumm.Status);
                    SQLcmd.Parameters.AddWithValue("@statusDate", Convert.ToDateTime(envelopeSumm.StatusDateTime));
                    SQLcmd.Parameters.AddWithValue("@receipientRoleId", 1); // Student = 1, DME = 2, Preceptor = 3

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


      

        public static int GetBatchRunID()
        {
            string storedProc = "dbo.sp_GetBatchRunID_Demo";


            using (SqlConnection SQLCon = new SqlConnection(ConfigurationManager.AppSettings["SMSDB_CONNSTRING2"]))
            {
                try
                {
                    var SQLcmd = new SqlCommand(storedProc, SQLCon);
                    SQLcmd.CommandType = CommandType.StoredProcedure;

                    SQLcmd.Parameters.AddWithValue("@startDate", startDate).Direction = ParameterDirection.Input;
                    SQLcmd.Parameters.AddWithValue("@sendDate", Convert.ToString(DateTime.Now)).Direction = ParameterDirection.Input;
                    SQLcmd.Parameters.Add("@runId", SqlDbType.Int).Direction = ParameterDirection.Output;

                    SQLcmd.Connection = SQLCon;
                    SQLCon.Open();
                    SQLcmd.ExecuteNonQuery();

                    batchRunID = (int)SQLcmd.Parameters["@runId"].Value;

                   // GetStudentRotationInfo(batchRunID);
                }
                finally
                {
                    SQLCon.Close();
                    SQLCon.Dispose();
                }
            }

            return batchRunID;

        }

   

       

        private static void SetEnvelopeReminders(Student student, EnvelopeDefinition envDef)
        {
            string newReminderDelay = "";

            switch (student.CreditWeeks)
            {
                case "4":
                    newReminderDelay = "28";
                    break;
                case "6":
                    newReminderDelay = "42";
                    break;
                case "8":
                    newReminderDelay = "56";
                    break;
                case "12":
                    newReminderDelay = "84";
                    break;
                default:
                    newReminderDelay = "40";
                    break;
            }

            //Set Dynamic Notifications
            envDef.Notification = new Notification
            {
                UseAccountDefaults = "false",
                Reminders = new Reminders
                {
                    ReminderEnabled = "true",
                    ReminderFrequency = "2",
                    ReminderDelay = newReminderDelay
                },
                Expirations = new Expirations
                {
                    ExpireEnabled = "true",
                    ExpireWarn = "7",
                    ExpireAfter = "120"
                }
            };
        }

// end requestSignatureFromTemplateTest()


        public static string CreatePersonalizedForm(Student s, int pdfFormId)
        {
            string currentForm = "";
            string outputForm = "";
            string imagePath = "";



            imagePath = @"w:\data\aua\StudentPhotos\" + s.StudentPicturePath + "_" + s.StudentId + ".jpg";
            //imagePath.Replace(" ", "");

            switch (pdfFormId)
            {
                case 1:
                    currentForm = @"C:\Docusign\AUA_Forms\StudentClerkshipEvaluationForm.pdf";
                    outputForm = @"C:\Docusign\AUA_Forms\StudentClerkshipEvaluationForm_" + s.StudentId + ".pdf";
                    break;
                case 2:
                    currentForm = @"C:\Docusign\AUA_Forms\StudentFacultyEvaluationForm.pdf";
                    outputForm = @"C:\Docusign\AUA_Forms\StudentFacultyEvaluationForm_" + s.StudentId + ".pdf";
                    break;
                case 3:
                    currentForm = @"C:\Docusign\AUA_Forms\MidClerkshipAssessmentForm.pdf";
                    outputForm = @"C:\Docusign\AUA_Forms\MidClerkshipAssessmentForm_" + s.StudentId + ".pdf";
                    break;
                case 4:
                    currentForm = @"C:\Docusign\AUA_Forms\StudentPortfolioForm.pdf";
                    outputForm = @"C:\Docusign\AUA_Forms\StudentPortfolioForm_" + s.StudentId + ".pdf";
                    break;
                case 5:
                    currentForm = @"C:\Docusign\AUA_Forms\ComprehensiveStudentClerkshipAssessmentForm.pdf";
                    outputForm = @"C:\Docusign\AUA_Forms\ComprehensiveStudentClerkshipAssessmentForm_" + s.StudentId + ".pdf";
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


        //Create Standard Signer
        private static Signer CreateSigner(string recipientEmail, string recipientName, string templateRoleName, string routingOrder, string recipientId)
        {
            Signer signer = new Signer
            {
                Email = recipientEmail,
                Name = recipientName,
                RoleName = templateRoleName,
                RoutingOrder = routingOrder,
                RecipientId = recipientId,
                Tabs = new Tabs { TextTabs = new List<Text>() }
            };

            //Only create once for Recipient/Signer
            //Each recip gets their own tabs
            return signer;

        }

        //Create an Intermediary
        private static Agent CreateAgent(string recipientEmail, string recipientName, string templateRoleName,
            string routingOrder, string recipientId)
        {
            Agent agent = new Agent()
            {
                Email = recipientEmail,
                Name = recipientName,
                RoleName = templateRoleName,
                RoutingOrder = routingOrder,
                RecipientId = recipientId,

            };
            return agent;
        }

        //Create a CC
        private static CarbonCopy CreateCarbonCopy(string recipientEmail, string recipientName, string templateRoleName, string routingOrder, string recipientId)
        {
            CarbonCopy carbonCopy = new CarbonCopy()
            {
                Email = recipientEmail,
                Name = recipientName,
                RoleName = templateRoleName,
                RoutingOrder = routingOrder,
                RecipientId = recipientId,

            };
            //Only create once for Recipient/Signer
            //Each recip gets their own tabs
            return carbonCopy;

        }
        


        //**********************************************************************************************************************
        //**********************************************************************************************************************
        //*  HELPER FUNCTIONS
        //**********************************************************************************************************************
        //**********************************************************************************************************************
        public static void configureApiClient(string basePath)
        {
            // instantiate a new api client
            ApiClient apiClient = new ApiClient(basePath);

            // set client in global config so we don't need to pass it to each API object.
            DocuSign.eSign.Client.Configuration.Default.ApiClient = apiClient;

        }



    } // end class


} //end namespace




