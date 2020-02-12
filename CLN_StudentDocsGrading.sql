select sc.StudentID,sc.CourseUID,sc.CourseDateStart,sc.CourseDateEnd, sc.Eval1,sc.Eval2,sc.Eval3,sc.Eval4,sc.Eval5,sc.Eval6,sc.StudentEvalRcvdDate,sc.StudentMidClerkEvalRcvdDate
,sc.StudentPortfolioRcvdDate from Student_Courses_tbl sc where
StudentID IN
(59086	,
67309	,
69103	,
69133	,
100093	,
100117	,
100408	,
100507	,
100544	,
100650	,
100693	,
100715	,
101088	,
101140	,
101279	,
101332	,
101470	,
370255	,
390147	,
555456	,
1075599	,
1076813	,
--1078833	,
1079593	
)
and CourseUID = 7116
and sc.CourseDateStart = '2016-11-28'