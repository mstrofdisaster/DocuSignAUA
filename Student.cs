namespace DocuSign_AUA
{
    struct Student
    {
        public string StudentId             { get; set; }
        public string StudentFirstName      { get; set; }
        public string StudentLastName       { get; set; }
        public string StudentName           { get; set; }
        public string StudentPicturePath    { get; set; }
        public string StudentEmailSchool    { get; set; }

        public string CourseId          { get; set; }
        public string ClinicalRotation  { get; set; }
        public string ClinicalSite      { get; set; }
        public string StartDate         { get; set; }
        public string EndDate           { get; set; }
        public string EnvelopeId        { get; set; }
        public string CreditWeeks        { get; set; }

        public string StudentPreceptorName { get; set; }
        public string StudentPreceptorEmail { get; set; }
        public string StudentDMEName { get; set; }
        public string StudentDMEEmail { get; set; }
        public string StudentDMEComments { get; set; }



    }

}
