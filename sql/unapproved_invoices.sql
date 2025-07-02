SELECT 
    ReviewerName AS 'Reviewer', 
    Job, 
    DateAssigned AS 'Date Assigned', 
    InvDate AS 'Invoice Date', 
    UIMth AS 'Month', 
    Name AS 'Vendor Name', 
    InvoiceLineTotal AS 'Invoice Line Total'
FROM EGC_Unapproved_Invoices
WHERE ABS(DATEDIFF(DAY, DateAssigned, GETDATE())) > 7
order by DateAssigned desc;
