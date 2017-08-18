USE [DW]
GO
/****** Object:  StoredProcedure [dbo].[uSP_ETL_ReIndexs2]    Script Date: 08/18/2017 17:18:56 ******/
DROP PROCEDURE [dbo].[uSP_ETL_ReIndexs2]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
Create Procedure [dbo].[uSP_ETL_ReIndexs2](@SendMail Int = 2)
as
begin
  Declare @Proc Varchar(50) = 'uSP_ETL_ReIndexs2'
  Declare @Cnt Int =0
  Declare @Msg Varchar(Max) =''
  Declare @RowCnt Int =0
  Declare @Errcode int = -1
  Declare @CR Varchar(4) = char(13)+char(10)

  Declare @Rebuild_Value Int = 50
  Declare @Reorganize_Value Int = 10  

  Declare @strSQL Varchar(Max)=''

  Begin Try
    -- 2015/10/21 Rickliu 參考此網址資訊，更改重建索引方式
    -- 重建或重新組織索引(建議透過 sys.dm_db_index_physical_stats DMV 調查)
    -- http://www.dotblogs.com.tw/ricochen/archive/2012/12/05/85419.aspx

    Declare Cur_Rebuild_Index Cursor for
       SELECT 'ALTER INDEX [' + ix.name + '] ON [' + s.name + '].[' + t.name + '] ' +
              CASE
                WHEN ps.avg_fragmentation_in_percent > @Rebuild_Value THEN 'REBUILD'
                ELSE 'REORGANIZE'
              END +
              CASE
                WHEN pc.partition_count > 1
                THEN ' PARTITION = ' + CAST(ps.partition_number AS nvarchar(MAX))
                ELSE ''
              END as Script, 
              
              '因索引 [Table:'+t.name+', Index:'+ix.name +', Columns: '+
              Replace(
                (SELECT CLS.[name] + ', '
                   FROM [sys].[index_columns] INXCLS
                        INNER JOIN [sys].[columns] CLS 
                           ON INXCLS.[object_id] = CLS.[object_id] 
                          AND INXCLS.[column_id] = CLS.[column_id]
                  WHERE IX.[object_id] = INXCLS.[object_id] 
                    AND IX.[index_id] = INXCLS.[index_id]
                    AND INXCLS.[is_included_column] = 0
                    FOR XML PATH('')
              
                )+'#', ', #', '')+'] 碎片值達 ['+
              Convert(Varchar(10), Round(ps.avg_fragmentation_in_percent, 2))+'%], 將進行重新'+
              CASE
                WHEN ps.avg_fragmentation_in_percent > @Rebuild_Value THEN '建立.'
                ELSE '組織.'
              END as Msg
         FROM sys.indexes AS ix
              INNER JOIN sys.tables t
                 ON t.object_id = ix.object_id
              INNER JOIN sys.schemas s
                 ON t.schema_id = s.schema_id
              INNER JOIN
                 (SELECT object_id                   ,
                         index_id                    ,
                         avg_fragmentation_in_percent,
                         partition_number
                    FROM sys.dm_db_index_physical_stats (DB_ID(), NULL, NULL, NULL, NULL)
                   where avg_fragmentation_in_percent > @Reorganize_Value 
                 ) ps
                 ON t.object_id = ps.object_id
                AND ix.index_id = ps.index_id
              INNER JOIN
                (SELECT object_id,
                        index_id ,
                        COUNT(DISTINCT partition_number) AS partition_count
                   FROM sys.partitions
                  GROUP BY object_id, index_id
                ) pc
                ON t.object_id = pc.object_id
               AND ix.index_id = pc.index_id
         WHERE ix.name IS NOT NULL
           and t.name <> 'Trans_Log'
         order by ps.avg_fragmentation_in_percent desc
   
    Open Cur_Rebuild_Index
    Fetch Next From Cur_Rebuild_Index Into @strSQL, @Msg
    
    While @@Fetch_status = 0
    begin
      Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
      Fetch Next From Cur_Rebuild_Index Into @strSQL, @Msg
    end
    Close Cur_Rebuild_Index
    Deallocate Cur_Rebuild_Index
  end try
  begin catch
    set @Cnt = -1
    set @Msg = '(錯誤訊息:'+ERROR_MESSAGE()+')'
  end catch
  Exec uSP_Sys_Write_Log @Proc, @Msg, @Msg, @Cnt
end
GO
