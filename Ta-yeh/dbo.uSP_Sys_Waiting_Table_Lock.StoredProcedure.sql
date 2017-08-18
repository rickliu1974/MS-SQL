USE [DW]
GO
/****** Object:  StoredProcedure [dbo].[uSP_Sys_Waiting_Table_Lock]    Script Date: 08/18/2017 17:18:56 ******/
DROP PROCEDURE [dbo].[uSP_Sys_Waiting_Table_Lock]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE procedure [dbo].[uSP_Sys_Waiting_Table_Lock] 
  @tb_name varchar(500)
as
begin
  Declare @Cnt int
  Declare @Proc Varchar(50) = 'uSP_Sys_Waiting_Table_Lock'
  Declare @Msg Varchar(1000)
  /*
     -- Rickliu 2014/09/26 為了 RealTime 常態性 DeadLock 所以撰寫此程序進行規避
     select object_name(p.object_id) as TableName, 
            resource_type, resource_description
       from sys.dm_tran_locks l
            join sys.partitions p 
              on l.resource_associated_entity_id = p.hobt_id
      where object_name(p.object_id) = 'RealTime_SaleData'
  */
  set @tb_name = LTrim(Rtrim(Upper(@tb_name)))
  if @tb_name <> ''
     While 1=1
     begin
        select @Cnt = Count(1)
          from sys.dm_tran_locks l
               join sys.partitions p 
                 on l.resource_associated_entity_id = p.hobt_id
         where Upper(object_name(p.object_id)) = @tb_name
        
        if @Cnt = 0
           Break
        else
        begin
           Set @Msg = @tb_name+' 資料表仍在鎖定中!!'
           Exec uSP_Sys_Write_Log @Proc, @Msg, '', 0
           Waitfor delay '00:00:01'
        end
     end
end
GO
