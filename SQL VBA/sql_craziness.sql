WITH PartitionedCounted AS
(
    SELECT ROW_NUMBER() OVER(PARTITION BY Objekt, EOMONTH(DATUM) ORDER BY DATUM) AS Nr
          ,EOMONTH(DATUM) AS EOM
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
