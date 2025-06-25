SELECT A.Project, 
       B.EMail, 
       B.Name, 
       B.TitleGroup
FROM EGC_ProjectionCoversheet A 
INNER JOIN EGC_JobTeamRoster B 
    ON A.Project = B.Project
WHERE B.TitleGroup IN ('PM', 'PX')
    AND A.Project IN ('24-2902', '24-2933')-- FOR TESTING