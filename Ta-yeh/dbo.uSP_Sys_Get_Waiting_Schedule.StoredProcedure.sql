USE [DW]
GO
/****** Object:  StoredProcedure [dbo].[uSP_Sys_Get_Waiting_Schedule]    Script Date: 07/24/2017 14:44:00 ******/
DROP PROCEDURE [dbo].[uSP_Sys_Get_Waiting_Schedule]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
Create Proc [dbo].[uSP_Sys_Get_Waiting_Schedule](@Self_Schedule_Name Varchar(Max) = '', @Run_Time Int = 1000)
as
begin
  Declare @Proc Varchar(100) ='uSP_Sys_Get_Waiting_Schedule'
  Declare @Errcode Int = -1
  Declare @Result Int = 0
  Declare @CR Varchar(4) =Char(13)+Char(10)
  Declare @Msg Varchar(max) = ''
  Declare @Wait_Sec Int = 0
       
  Declare @JobInfo Table
   (
    [Job_ID] uniqueidentifier,
    [Last_Run_Date] varchar(255),
    [Last_Run_Time] varchar(255),
    [Next_Run_Date] varchar(255),
    [Next_Run_Time] varchar(255),
    [Next_Run_Schedule_ID] varchar(255),
    [Requested_To_Run] varchar(255),
    [Request_Source] varchar(255),
    [Request_Source_ID] varchar(255),
    [Running] varchar(255),
    [Current_Step] varchar(255),
    [Current_Retry_Attempt] varchar(255),
    [State] varchar(255)
   )
  Declare @Run_Cnt Int = 1;
  Declare @Run_List Varchar(max)
  
  While @Run_Cnt <> @Run_Time
  begin
    Begin Try
      Declare @Delay_Time Varchar(20) = (Select Convert(Varchar(20), dateadd(SECOND, Cast(Rand() * 100 as Int), convert(DATETIME, 0)), 114))
      
      Delete @JobInfo
      Print 'Check Time:'+Cast(@Run_Cnt as Varchar)+'/'+Cast(@Run_Time as Varchar)
      
      insert into @JobInfo Exec master.dbo.xp_sqlagent_enum_jobs 1,''

      set @Run_List =''
      select @Run_List = Coalesce(@Run_List+','+Char(13)+Char(10),'')+sj.name
        from @JobInfo as jf
             left join msdb.dbo.sysjobs as sj
               on jf.job_id = sj.job_id
       where jf.State = 1 -- 目前執行中的排程
         and sj.Name Not Like '%uSP_Sys_Kill_LY_Process%'
         and sj.Name <> @Self_Schedule_Name -- 排除自己的排程，判別是否還有其他排程再進行
      
      If Isnull(@Run_List, '') <> '' And @Run_Cnt <> @Run_Time
      begin
         Waitfor Delay @Delay_Time
         set @Msg = '執行 ['+@Self_Schedule_Name+'] 排程但時間衝突，'+
                    '檢查第 '+Cast(@Run_Cnt as Varchar)+' / '+Cast(@Run_Time as Varchar)+' 次，'+
                    '隨機等待 ['+@Delay_Time+'] 再行檢查。正在運行的排程：['+@Run_List+']。'
                    
         Print 'Other schedule at running.'
         Print 'Running schedule :'+Char(13)+Char(10)+@Run_List
         Print 'Waiting Time: '+@Delay_Time
      
         Exec uSP_Sys_Write_Log @Self_Schedule_Name, @Msg, '', @Run_Cnt, 1
         set @Result = @Errcode
      end
      else
      begin
         set @Msg = '檢查第 '+Cast(@Run_Cnt as Varchar)+' / '+Cast(@Run_Time as Varchar)+' 次，'+
                    '目前沒有任何排程正在運行，將開始進行 ['+@Self_Schedule_Name+'] 排程。'
         Exec uSP_Sys_Write_Log @Self_Schedule_Name, @Msg, '', @Run_Cnt, 1
         Print 'No any schedule at running.'
         set @Result = 0
         Break
      end
      
      set @Run_Cnt = @Run_Cnt + 1
    end Try
    Begin Catch
        set @Result = @Errcode
        Set @Msg = @Proc+' 命令失敗!!...(錯誤訊息:'+ERROR_MESSAGE()+')'       
        Exec uSP_Sys_Write_Log @Self_Schedule_Name, @Msg, '', @Run_Cnt, 1
    End Catch
  end
  
  Return @Result 
end
GO
