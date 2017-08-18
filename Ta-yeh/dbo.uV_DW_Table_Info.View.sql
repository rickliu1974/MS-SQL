USE [DW]
GO
/****** Object:  View [dbo].[uV_DW_Table_Info]    Script Date: 07/24/2017 14:43:55 ******/
DROP VIEW [dbo].[uV_DW_Table_Info]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
Create View [dbo].[uV_DW_Table_Info]
as
 SELECT Top 1000
        [DB_Name] = db_name(),
        [Object_Name] = a2.name, 
        [Object_Tyee] = a2.Type_desc,
        [Rec_Counts] = a1.rows, 
        [Data_Size_GB] = Convert(decimal(18,4), a1.data * 8.0/1024/1024), 
        [Index_Size_GB] = Convert(decimal(18,4), (CASE WHEN (a1.used + ISNULL(a4.used,0)) > a1.data THEN (a1.used + ISNULL(a4.used,0)) - a1.data ELSE 0 END) * 8.0/1024/1024), 
        [NonUses_Size_GB] = Convert(decimal(18,4), (CASE WHEN (a1.reserved + ISNULL(a4.reserved,0)) > a1.used THEN (a1.reserved + ISNULL(a4.reserved,0)) - a1.used ELSE 0 END) * 8.0/1024/1024),
        [Total_Size_GB] = Convert(decimal(18,4), (a1.reserved + ISNULL(a4.reserved,0))* 8.0/1024/1024),
        [Object_FullName] = db_name()+'.'+a3.name+'.'+a2.name
   FROM (SELECT ps.object_id, 
                SUM (CASE WHEN (ps.index_id < 2) THEN row_count ELSE 0 END) AS [rows], SUM (ps.reserved_page_count) AS reserved, SUM (ps.used_page_count) AS used, 
                SUM (CASE WHEN (ps.index_id < 2) THEN (ps.in_row_data_page_count + ps.lob_used_page_count + ps.row_overflow_used_page_count) ELSE (ps.lob_used_page_count + ps.row_overflow_used_page_count) END) AS data
           FROM sys.dm_db_partition_stats as ps 
          GROUP BY ps.object_id 
        ) AS a1 LEFT OUTER JOIN 
       (SELECT it.parent_id, 
               SUM(ps.reserved_page_count) AS reserved, SUM(ps.used_page_count) AS used 
          FROM sys.dm_db_partition_stats as ps 
               INNER JOIN sys.internal_tables as it ON (it.object_id = ps.object_id) 
         WHERE it.internal_type IN (202, 204) 
         GROUP BY it.parent_id 
       ) AS a4 ON (a4.parent_id = a1.object_id) 
       INNER JOIN sys.all_objects a2 ON (a1.object_id = a2.object_id) 
       INNER JOIN sys.schemas a3 ON (a2.schema_id = a3.schema_id) 
 WHERE a2.type <> 'S' and a2.type <> 'IT' 
 ORDER BY 4 DESC
GO
