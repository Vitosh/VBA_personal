USE Gladiator
GO

WITH PartitionedCounted AS 
(
    SELECT ROW_NUMBER() OVER(PARTITION BY Objekt, EOMONTH(DATUM) ORDER BY DATUM) [Nr]
	,EOMONTH(DATUM) [EOM]
	,Datum
	,Objekt
	FROM tempt_report AS tbl
)

SELECT pc.EOM [Date],pc.Objekt,tbl.Datum [Initial Date]
FROM PartitionedCounted AS pc
INNER JOIN tempt_report AS tbl ON tbl.Objekt=pc.Objekt AND
EOMONTH(TBL.Datum)=pc.EOM
WHERE pc.Nr > 1
ORDER BY Objekt
--
IF 2=2
SELECT 101
ELSE 
SELECT 222
--
USE Gladiator
GO

WITH PartitionedCounted AS 
(
    SELECT ROW_NUMBER() OVER(PARTITION BY Objekt, EOMONTH(DATUM) ORDER BY DATUM) [Nr]
	,EOMONTH(DATUM) [EOM]
	,Datum
	,Objekt
		FROM 
		(
		SELECT * FROM
			(
				SELECT * FROM tempt_report
				UNION ALL	
				SELECT * FROM tempt_report_test
			)
			AS united_table
		)
		AS tbl
)
SELECT pc.EOM [Date],pc.Objekt ,tbl.Datum [Initial Date], LEFT(tbl.Zeit,5)[Initial Time], tbl.Benutzer[User]
FROM PartitionedCounted AS pc
INNER JOIN tempt_report AS tbl ON tbl.Objekt=pc.Objekt AND
EOMONTH(TBL.Datum)=pc.EOM
WHERE pc.Nr > 1
ORDER BY Objekt, [Initial Date], [Initial Time]
