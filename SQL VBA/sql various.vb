dbcc sqlperf(logspace);
go

SELECT name, recovery_model_desc  
   FROM sys.databases  
GO

use master;
alter database tsql2012 set recovery full;

backup database [tsql2012]
TO DISK = 'C:\Users\v.doynov\Desktop\backup\TSQL2012_db.bak'
go

backup log TSQL2012
TO DISK = 'C:\Users\v.doynov\Desktop\backup\TSQL2012.trn'
go

USE Tempt
GO
SELECT [Current LSN], [Begin Time], SPID, [Database Name], [Transaction Begin], [Transaction ID], [Transaction Name], [Transaction SID], Context, Operation
FROM ::fn_dblog (null, null)
WHERE [Transaction Name] = 'INSERT'
GO

Select * FROM sys.fn_dblog(NULL,NULL)

SELECT [begin time], 
       [rowlog contents 1], 
       [Transaction Name], 
       Operation
  FROM sys.fn_dblog
   (NULL, NULL)
  WHERE operation IN
   ('LOP_DELETE_ROWS');

SELECT *
FROM fn_dump_dblog
(NULL,NULL,N'DISK',1,N'C:\Users\v.doynov\Desktop\backup\TSQL2012.trn', 
DEFAULT,DEFAULT,DEFAULT,DEFAULT,DEFAULT,DEFAULT,DEFAULT,DEFAULT,DEFAULT, 
DEFAULT,DEFAULT,DEFAULT,DEFAULT,DEFAULT,DEFAULT,DEFAULT,DEFAULT,DEFAULT,
DEFAULT,DEFAULT,DEFAULT,DEFAULT,DEFAULT,DEFAULT,DEFAULT,DEFAULT,DEFAULT,
DEFAULT,DEFAULT,DEFAULT,DEFAULT,DEFAULT,DEFAULT,DEFAULT,DEFAULT,DEFAULT,
DEFAULT,DEFAULT,DEFAULT,DEFAULT,DEFAULT,DEFAULT,DEFAULT,DEFAULT,DEFAULT,
DEFAULT,DEFAULT,DEFAULT,DEFAULT,DEFAULT,DEFAULT,DEFAULT,DEFAULT,DEFAULT,
DEFAULT,DEFAULT,DEFAULT,DEFAULT,DEFAULT,DEFAULT,DEFAULT,DEFAULT, 
DEFAULT);
