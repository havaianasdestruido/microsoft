--Corporate Benefits Sample
--==========================
--by Ken Bergmann

GO
Set nocount on

PRINT 'INSERTING Version Information...'
GO

pcIns_VERS 'Corporate Benefits Sample', NULL, '1.00.0000'
go

PRINT 'INSERTING Version Information...Completed'
GO

--update tbVERS_Version set appversion = '1.00.0000' where appname = 'Corporate Benefits Sample'
