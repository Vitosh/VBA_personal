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


if  exists (select * from sys.objects) 
	select 'The sys.objects exists'
else
	select 'The sys.objects does not exist'

GO

--BACKUP DB
select * from sys.fn_builtin_permissions(default)

DECLARE @databasename AS NVARCHAR(MAX)
	, @timecomponent AS NVARCHAR(MAX)
	, @sqlcommand AS NVARCHAR(MAX);
	
SET @databasename = (SELECT MIN(name) FROM sys.databases WHERE name 
    NOT IN ('master', 'model', 'msdb', 'tempdb'));
	
WHILE @databasename IS NOT NULL
BEGIN
  SET @timecomponent = REPLACE(REPLACE(REPLACE(CONVERT(NVARCHAR, 
    GETDATE(), 120), ' ', '_'), ':', ''), '-', '');
	
  SET @sqlcommand = 'BACKUP DATABASE ' + @databasename + ' TO DISK =
       ''C:\Backups\' + @databasename + '_' + @timecomponent + '.bak''';
	   
  PRINT @sqlcommand;
  
  EXEC(@sqlcommand);
  
  SET @databasename = (SELECT MIN(name) FROM sys.databases WHERE name 
    NOT IN ('master', 'model', 'msdb', 'tempdb') AND name > @databasename);
END;

GO
-- 
Automatic insert of data
DECLARE @i AS int = 1;
WHILE @i < 1000000
BEGIN
SET @i = @i + 1;
INSERT INTO dbo.TestStructure
(id, filler1, filler2)
VALUES
(@i, 'a', 'b');
END;
GO

select * from dbo.TestStructure
--
--clearing cache
DBCC FREEPROCCACHE
GO
--

