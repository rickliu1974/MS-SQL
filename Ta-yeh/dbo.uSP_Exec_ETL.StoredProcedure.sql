USE [DW]
GO
/****** Object:  StoredProcedure [dbo].[uSP_Exec_ETL]    Script Date: 08/18/2017 17:43:40 ******/
DROP PROCEDURE [dbo].[uSP_Exec_ETL]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE Procedure [dbo].[uSP_Exec_ETL](@Core_ETL Int = 0)
as
begin
  Declare @Proc Varchar(50) = 'uSP_Exec_ETL'
  Declare @Sender Varchar(100) = ''
  Declare @CR Varchar(4) =Char(13)+Char(10)
  Declare @Mail_1 NVarchar(Max) = '', @Mail_2 NVarchar(Max) = '', @Run_Proc_Msg NVarchar(Max) = '', @Run_Msg NVarchar(1000) = ''
  Declare @Cnt_xls int = 0, @Cnt_txt int = 0, @Run_Err int = 0, @Run_Core_Err int = 0
  Declare @Cnt Int = 0;
  Declare @Proc_SubName NVarchar(Max) = 'ETL ��ƭ��x��Ƥ��ѧ@�~'
  Declare @Start_Time DateTime, @End_Time DateTime
  Declare @strSQL Varchar(Max)
  Declare @sStr Varchar(1000)
  Declare @Msg Varchar(1000) = '', @sProcess Varchar(4)= ''
  Declare @Process Float = 0
  Declare @Tot_Process int = 9
  Declare @Get_Result int = 0
  Declare @Result Int = 0
  Declare @Err_Code int = -1
  Declare @chk_lys int = 0
  Declare @sRebuild Varchar(20) = '���㤤!!..'
  Declare @sSuss Varchar(20) = '���\'
  Declare @sFail Varchar(20) = '����'
  Declare @Run_ETL Int = 0
  Declare @sJob_id NVarchar(100), @sJob_Name NVarchar(100), @enabled int

  Declare @Sch_Jobs Table(
    [job_id] [uniqueidentifier] NOT NULL,
    [name] [sysname] NOT NULL,
    [enabled] [tinyint] NOT NULL
  )

  set @Mail_1 = '���@�~��C��W��06:00, 12:00 �}�l�i��C'+@CR
  set @Mail_1 = @Mail_1+'�ǳƶi�� '+@Proc_SubName+@CR
print '1-----------------------------------------------------------------------------------------------'
print 'Mail Body:'+@Mail_1
print '-----------------------------------------------------------------------------------------------'

  set @Process = 0
  set @sProcess = Convert(Varchar(4), Convert(Int, Round((@Process / @Tot_Process) * 100, 0)))+'%'

  Begin Try
    insert Into @Sch_Jobs
    (job_id, name, enabled)
    select job_id, name, enabled
      from msdb.dbo.sysjobs
     where name not like 
           case
            when 1=@Core_ETL 
             then @Proc+'_By_Core.%'
             else @Proc+'.%'
           end
       and enabled = 1

    Declare Cur_Agent_Jobs Cursor for
      select job_id, name, enabled
        from @Sch_Jobs

    set @Run_Proc_Msg = '�ˬd��V�t�άO�_�i�歫��@�~!!!'
    set @chk_lys = (select l_field2 from TYEIPDBS2.LYTDBTA13.dbo.ssyslb)
    if @chk_lys = 1
    begin
      Set @sStr = '�ثe���b�����V���� ['+@Proc+']�A�ҥH�L�k���楻�{��!!..'+convert(varchar(19), getdate(), 121)
      Set @Mail_1 = @Mail_1 + @sStr+@CR
      Exec uSP_Sys_Write_Log @Proc, @sStr, '', @Err_Code 
print '2-----------------------------------------------------------------------------------------------'
print 'Mail Body:'+@Mail_1
print '-----------------------------------------------------------------------------------------------'
      Goto End_Proc
    end

    set @Run_Proc_Msg = '�N�q����אּ����e��!!!'
    Update TB_Config set Flag = 1, Value = @sRebuild+@sProcess where kind = '0001'
    Set @Run_Msg = @Proc+' ' +@Run_Msg +'...����}�l�ɶ��G[ '+CONVERT(Varchar(100), GETDATE(),120)+' ].'
    Set @Mail_1 = @Mail_1 +@Run_Msg+@sProcess+@CR
print '3-----------------------------------------------------------------------------------------------'
print 'Mail Body:'+@Mail_1
print '-----------------------------------------------------------------------------------------------'
    Exec uSP_Sys_Write_Log @Proc, @Run_Msg, '', 0, 1
    
    set @Start_Time = Getdate()
    set @Run_Proc_Msg = '���� uSP_Sys_Release_Memory�i�����Ʈw�O����Ŷ��j... '+@sProcess
    Exec @Get_Result = DW.dbo.uSP_Sys_Release_Memory 
      
    ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
    -- 2014/11/28 Rickliu �u���o���b�ϥ�(enabled = 0)�A�קK���ϥΪ��Ƶ{��{�ǵ�����S�Q�}�ҡC�S�]��Ƶ��Ƥp�� 100 ���A�ҥH�ĥ��ܼƪ��覡
    Open Cur_Agent_Jobs
    Fetch Next From Cur_Agent_Jobs Into @sJob_id, @sJob_Name, @enabled

    While @@Fetch_Status = 0
    begin
print '4-----------------------------------------------------------------------------------------------'
print 'Mail Body:'+@Mail_1
print '-----------------------------------------------------------------------------------------------'
      set @Start_Time = Getdate()
      set @Run_Proc_Msg = '���αƵ{ ['+Isnull(@sJob_Name, '')+']'
      set @strSQL = 'EXEC msdb.dbo.sp_update_job @job_id=N'''+@sJob_id+''', @enabled=0 '
      Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @sStr, @strSQL, 0

      if @Get_Result = @Err_Code
         set @Run_Msg = @sFail
      else
         set @Run_Msg = @sSuss

      set @End_Time = Getdate()
      set @sStr = @Run_Proc_Msg+'..'+@Run_Msg+'!!..����'+ Convert(Varchar(20), @end_time - @Start_Time, 114) 
      set @Mail_1 = @Mail_1 + @sStr+@CR
print '5-----------------------------------------------------------------------------------------------'
print 'Mail Body:'+@Mail_1
print '-----------------------------------------------------------------------------------------------'
      Exec uSP_Sys_Write_Log @Proc, @sStr, '', 0

      Fetch Next From Cur_Agent_Jobs Into @sJob_id, @sJob_Name, @enabled
    end
    Close Cur_Agent_Jobs
  
    -- �H�U���涶�Ǥ��i�H�A��
    ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
    ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
    set @Process = 1
    set @sProcess = Convert(Varchar(4), Convert(Int, Round((@Process / @Tot_Process) * 100, 0)))+'%'
    set @Start_Time = Getdate()
    set @Run_Proc_Msg = Substring(Convert(Varchar(4), @Process+100), 2, 2)+'.���� uSP_ETL_New_Stock_Lists�i��s�ӫ~���s�~��T�j... '+@sProcess
    set @Run_ETL = 0
    Update TB_Config set Flag = 1, Value = @sRebuild+@sProcess where kind = '0001'
    Exec @Get_Result = DW.dbo.uSP_ETL_New_Stock_Lists @Run_ETL
      
    if @Get_Result = @Err_Code
       set @Run_Msg = @sFail
    else
       set @Run_Msg = @sSuss

    set @End_Time = Getdate()
    set @sStr = @Run_Proc_Msg +'..'+@Run_Msg+'!!..����'+ Convert(Varchar(20), @end_time - @Start_Time, 114) 
    set @Mail_1 = @Mail_1 + @sStr+@CR
print '6-----------------------------------------------------------------------------------------------'
print 'Mail Body:'+@Mail_1
print '-----------------------------------------------------------------------------------------------'
    Exec uSP_Sys_Write_Log @Proc, @sStr, '', @Get_Result, 1

    ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
    ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
    set @Process = 2
    set @sProcess = Convert(Varchar(4), Convert(Int, Round((@Process / @Tot_Process) * 100, 0)))+'%'
    set @Start_Time = Getdate()
    set @Run_Proc_Msg = Substring(Convert(Varchar(4), @Process+100), 2, 2)+'.���� uSP_ETL_PEMPLOY�iPEMPLOY ���u��ƪ�i����ѡj... '+@sProcess
    Update TB_Config set Flag = 1, Value = @sRebuild+@sProcess where kind = '0001'
    Exec @Get_Result = DW.dbo.uSp_ETL_pemploy

    if @Get_Result = @Err_Code
       set @Run_Msg = @sFail
    else
       set @Run_Msg = @sSuss

    set @End_Time = Getdate()
    set @sStr = @Run_Proc_Msg +'..'+@Run_Msg+'!!..����'+ Convert(Varchar(20), @end_time - @Start_Time, 114) 
    set @Mail_1 = @Mail_1 + @sStr+@CR
print '33-----------------------------------------------------------------------------------------------'
print 'Mail Body:'+@Mail_1
print '-----------------------------------------------------------------------------------------------'
    Exec uSP_Sys_Write_Log @Proc, @sStr, '', @Get_Result, 1

    ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
    ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
    set @Process = 3
    set @sProcess = Convert(Varchar(4), Convert(Int, Round((@Process / @Tot_Process) * 100, 0)))+'%'
    set @Start_Time = Getdate()
    set @Run_Proc_Msg = Substring(Convert(Varchar(4), @Process+100), 2, 2)+'.���� uSP_ETL_PCUST�iPCUST �t��/�Ȥ��ƪ�i����ѡj... '+@sProcess
    Update TB_Config set Flag = 1, Value = @sRebuild+@sProcess where kind = '0001'
    Exec @Get_Result = DW.dbo.uSp_ETL_pcust

    if @Get_Result = @Err_Code
       set @Run_Msg = @sFail
    else
       set @Run_Msg = @sSuss

    set @End_Time = Getdate()
    set @sStr = @Run_Proc_Msg +'..'+@Run_Msg+'!!..����'+ Convert(Varchar(20), @end_time - @Start_Time, 114) 
    set @Mail_1 = @Mail_1 + @sStr+@CR
print '34-----------------------------------------------------------------------------------------------'
print 'Mail Body:'+@Mail_1
print '-----------------------------------------------------------------------------------------------'
    Exec uSP_Sys_Write_Log @Proc, @sStr, '', @Get_Result, 1

    ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
    ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---

    set @Process = 4
    set @sProcess = Convert(Varchar(4), Convert(Int, Round((@Process / @Tot_Process) * 100, 0)))+'%'
    set @Start_Time = Getdate()
    set @Run_Proc_Msg = Substring(Convert(Varchar(4), @Process+100), 2, 2)+'.���� uSP_ETL_SWAREDT�iSWAREDT �w�s��ƪ�i����ѡj... '+@sProcess
    Update TB_Config set Flag = 1, Value = @sRebuild+@sProcess where kind = '0001'
    Exec @Get_Result = DW.dbo.uSp_ETL_SWAREDT

    if @Get_Result = @Err_Code
       set @Run_Msg = @sFail
    else
       set @Run_Msg = @sSuss

    set @End_Time = Getdate()
    set @sStr = @Run_Proc_Msg +'..'+@Run_Msg+'!!..����'+ Convert(Varchar(20), @end_time - @Start_Time, 114) 
    set @Mail_1 = @Mail_1 + @sStr+@CR
print '36-----------------------------------------------------------------------------------------------'
print 'Mail Body:'+@Mail_1
print '-----------------------------------------------------------------------------------------------'
    Exec uSP_Sys_Write_Log @Proc, @sStr, '', @Get_Result, 1

    ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
    ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---

    set @Process = 5
    set @sProcess = Convert(Varchar(4), Convert(Int, Round((@Process / @Tot_Process) * 100, 0)))+'%'
    set @Start_Time = Getdate()
    set @Run_Proc_Msg = Substring(Convert(Varchar(4), @Process+100), 2, 2)+'.���� uSP_ETL_SSTOCK�iSSTOCK �ӫ~��ƪ�i����ѡj... '+@sProcess
    Update TB_Config set Flag = 1, Value = @sRebuild+@sProcess where kind = '0001'
    Exec @Get_Result = DW.dbo.uSp_ETL_sstock

    if @Get_Result = @Err_Code
       set @Run_Msg = @sFail
    else
       set @Run_Msg = @sSuss

    set @End_Time = Getdate()
    set @sStr = @Run_Proc_Msg +'..'+@Run_Msg+'!!..����'+ Convert(Varchar(20), @end_time - @Start_Time, 114) 
    set @Mail_1 = @Mail_1 + @sStr+@CR
print '35-----------------------------------------------------------------------------------------------'
print 'Mail Body:'+@Mail_1
print '-----------------------------------------------------------------------------------------------'
    Exec uSP_Sys_Write_Log @Proc, @sStr, '', @Get_Result, 1

    ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
    ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---

    set @Process = 6
    set @sProcess = Convert(Varchar(4), Convert(Int, Round((@Process / @Tot_Process) * 100, 0)))+'%'
    set @Start_Time = Getdate()
    set @Run_Proc_Msg = Substring(Convert(Varchar(4), @Process+100), 2, 2)+'.���� uSP_ETL_SSLIP�iSSLIP �U����ڪ��Y��ƪ�i����ѡj... '+@sProcess
    Update TB_Config set Flag = 1, Value = @sRebuild+@sProcess where kind = '0001'
    Exec @Get_Result = DW.dbo.uSP_ETL_sslip

    if @Get_Result = @Err_Code
       set @Run_Msg = @sFail
    else
       set @Run_Msg = @sSuss

    set @End_Time = Getdate()
    set @sStr = @Run_Proc_Msg +'..'+@Run_Msg+'!!..����'+ Convert(Varchar(20), @end_time - @Start_Time, 114) 
    set @Mail_1 = @Mail_1 + @sStr+@CR
print '37-----------------------------------------------------------------------------------------------'
print 'Mail Body:'+@Mail_1
print '-----------------------------------------------------------------------------------------------'
    Exec uSP_Sys_Write_Log @Proc, @sStr, '', @Get_Result, 1

    ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
    ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
    set @Process = 7
    set @sProcess = Convert(Varchar(4), Convert(Int, Round((@Process / @Tot_Process) * 100, 0)))+'%'
    set @Start_Time = Getdate()
    set @Run_Proc_Msg = Substring(Convert(Varchar(4), @Process+100), 2, 2)+'.���� uSP_ETL_SSLPDT�iSSLPDT �U����ک��Ӹ�ƪ�i����ѡj... '+@sProcess
    Update TB_Config set Flag = 1, Value = @sRebuild+@sProcess where kind = '0001'
    Exec @Get_Result = DW.dbo.uSP_ETL_sslpdt

    if @Get_Result = @Err_Code
       set @Run_Msg = @sFail
    else
       set @Run_Msg = @sSuss

    set @End_Time = Getdate()
    set @sStr = @Run_Proc_Msg +'..'+@Run_Msg+'!!..����'+ Convert(Varchar(20), @end_time - @Start_Time, 114) 
    set @Mail_1 = @Mail_1 + @sStr+@CR
print '38-----------------------------------------------------------------------------------------------'
print 'Mail Body:'+@Mail_1
print '-----------------------------------------------------------------------------------------------'
    Exec uSP_Sys_Write_Log @Proc, @sStr, '', @Get_Result, 1

    ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
    ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
    set @Process = 8
    set @sProcess = Convert(Varchar(4), Convert(Int, Round((@Process / @Tot_Process) * 100, 0)))+'%'
    set @Start_Time = Getdate()
    set @Run_Proc_Msg = Substring(Convert(Varchar(4), @Process+100), 2, 2)+'.���� uSP_ETL_ReIndexs�iETL ���ثإ߯������j... '+@sProcess
    Update TB_Config set Flag = 1, Value = @sRebuild+@sProcess where kind = '0001'
    Exec @Get_Result = DW.dbo.uSP_ETL_ReIndexs2 1

    if @Get_Result = @Err_Code
       set @Run_Msg = @sFail
    else
       set @Run_Msg = @sSuss

    set @End_Time = Getdate()
    set @sStr = @Run_Proc_Msg +'..'+@Run_Msg+'!!..����'+ Convert(Varchar(20), @end_time - @Start_Time, 114) 
    set @Mail_1 = @Mail_1 + @sStr+@CR
print '40-----------------------------------------------------------------------------------------------'
print 'Mail Body:'+@Mail_1
print '-----------------------------------------------------------------------------------------------'
    Exec uSP_Sys_Write_Log @Proc, @sStr, '', @Get_Result, 1

    ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
    ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
    set @Process = 9
    set @sProcess = Convert(Varchar(4), Convert(Int, Round((@Process / @Tot_Process) * 100, 0)))+'%'
    set @Start_Time = Getdate()
    set @Run_Proc_Msg = Substring(Convert(Varchar(4), @Process+100), 2, 2)+'.���� uSP_Realtime_SaleData�iETL RealTime �Y�ɸ�ơj... '+@sProcess
    Update TB_Config set Flag = 1, Value = @sRebuild+@sProcess where kind = '0001'
    --Exec @Get_Result = DW.dbo.uSP_Realtime_SaleData '', 1 -- �}�ҵo�e Log ����
    Exec @Get_Result = DW.dbo.uSP_Realtime_SaleData '', 2
    
    if @Get_Result = @Err_Code
       set @Run_Msg = @sFail
    else
       set @Run_Msg = @sSuss

    set @End_Time = Getdate()
    set @sStr = @Run_Proc_Msg +'..'+@Run_Msg+'!!..����'+ Convert(Varchar(20), @end_time - @Start_Time, 114) 
    set @Mail_1 = @Mail_1 + @sStr+@CR
print '41-----------------------------------------------------------------------------------------------'
print 'Mail Body:'+@Mail_1
print '-----------------------------------------------------------------------------------------------'
    Exec uSP_Sys_Write_Log @Proc, @sStr, '', @Get_Result, 1
/*
    ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
    ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
    set @Process = 10
    set @sProcess = Convert(Varchar(4), Convert(Int, Round((@Process / @Tot_Process) * 100, 0)))+'%'
    set @Start_Time = Getdate()
    set @Run_Proc_Msg = Substring(Convert(Varchar(4), @Process+100), 2, 2)+'.���� uSP_ETL_Cust_NonSale_Stock�iETL �Ҧ��Ȥ᥼�P�ӫ~��ơj... '+@sProcess
    Update TB_Config set Flag = 1, Value = @sRebuild+@sProcess where kind = '0001'
    Exec @Get_Result = DW.dbo.uSP_ETL_Cust_NonSale_Stock

    if @Get_Result = @Err_Code
       set @Run_Msg = @sFail
    else
       set @Run_Msg = @sSuss

    set @End_Time = Getdate()
    set @sStr = @Run_Proc_Msg +'..'+@Run_Msg+'!!..����'+ Convert(Varchar(20), @end_time - @Start_Time, 114) 
    set @Mail_1 = @Mail_1 + @sStr+@CR
print '42-----------------------------------------------------------------------------------------------'
print 'Mail Body:'+@Mail_1
print '-----------------------------------------------------------------------------------------------'
    Exec uSP_Sys_Write_Log @Proc, @sStr, '', @Get_Result, 1
*/
-- 2017/08/07 Rickliu �����ϥ�
/*
    ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
    ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
    set @Process = 11
    set @sProcess = Convert(Varchar(4), Convert(Int, Round((@Process / @Tot_Process) * 100, 0)))+'%'
    set @Start_Time = Getdate()
    set @Run_Proc_Msg = Substring(Convert(Varchar(4), @Process+100), 2, 2)+'.���� uSP_ETL_Hunderd_NonSale_Stock�iETL �ʤj�Ȥ᥼�P�ӫ~��ơj... '+@sProcess
    Update TB_Config set Flag = 1, Value = @sRebuild+@sProcess where kind = '0001'
    Exec @Get_Result = DW.dbo.uSP_ETL_Hunderd_NonSale_Stock

    if @Get_Result = @Err_Code
       set @Run_Msg = @sFail
    else
       set @Run_Msg = @sSuss

    set @End_Time = Getdate()
    set @sStr = @Run_Proc_Msg +'..'+@Run_Msg+'!!..����'+ Convert(Varchar(20), @end_time - @Start_Time, 114) 
    set @Mail_1 = @Mail_1 + @sStr+@CR
print '43-----------------------------------------------------------------------------------------------'
print 'Mail Body:'+@Mail_1
print '-----------------------------------------------------------------------------------------------'
    Exec uSP_Sys_Write_Log @Proc, @sStr, '', @Get_Result, 1
*/
-- 2017/08/07 Rickliu �����ϥ�
/*
    ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
    ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
    set @Process = 12
    set @sProcess = Convert(Varchar(4), Convert(Int, Round((@Process / @Tot_Process) * 100, 0)))+'%'
    set @Start_Time = Getdate()
    set @Run_Proc_Msg = Substring(Convert(Varchar(4), @Process+100), 2, 2)+'.���� uSP_ETL_Cust_WeekSale_Stock�iETL �ʤj�Ȥ᥼�P�ӫ~���(�g)�j... '+@sProcess
    Update TB_Config set Flag = 1, Value = @sRebuild+@sProcess where kind = '0001'
    Exec @Get_Result = DW.dbo.uSP_ETL_Cust_WeekSale_Stock

    if @Get_Result = @Err_Code
       set @Run_Msg = @sFail
    else
       set @Run_Msg = @sSuss

    set @End_Time = Getdate()
    set @sStr = @Run_Proc_Msg +'..'+@Run_Msg+'!!..����'+ Convert(Varchar(20), @end_time - @Start_Time, 114) 
    set @Mail_1 = @Mail_1 + @sStr+@CR
print '44-----------------------------------------------------------------------------------------------'
print 'Mail Body:'+@Mail_1
print '-----------------------------------------------------------------------------------------------'
    Exec uSP_Sys_Write_Log @Proc, @sStr, '', @Get_Result, 1
*/
-- 2017/08/07 Rickliu �����ϥ�
/*
	---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
    ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
    set @Process = 13
    set @sProcess = Convert(Varchar(4), Convert(Int, Round((@Process / @Tot_Process) * 100, 0)))+'%'
    set @Start_Time = Getdate()
    set @Run_Proc_Msg = Substring(Convert(Varchar(4), @Process+100), 2, 2)+'.���� uSP_ETL_Cust_MonthSale_Stock�iETL �ʤj�Ȥ᥼�P�ӫ~���(��)�j... '+@sProcess
    Update TB_Config set Flag = 1, Value = @sRebuild+@sProcess where kind = '0001'
    Exec @Get_Result = DW.dbo.uSP_ETL_Cust_MonthSale_Stock

    if @Get_Result = @Err_Code
       set @Run_Msg = @sFail
    else
       set @Run_Msg = @sSuss

    set @End_Time = Getdate()
    set @sStr = @Run_Proc_Msg +'..'+@Run_Msg+'!!..����'+ Convert(Varchar(20), @end_time - @Start_Time, 114) 
    set @Mail_1 = @Mail_1 + @sStr+@CR
print '45-----------------------------------------------------------------------------------------------'
print 'Mail Body:'+@Mail_1
print '-----------------------------------------------------------------------------------------------'
    Exec uSP_Sys_Write_Log @Proc, @sStr, '', @Get_Result, 1
*/
-- 2017/08/07 Rickliu �����ϥ�
/*
    ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
    ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
    set @Process = 14
    set @sProcess = Convert(Varchar(4), Convert(Int, Round((@Process / @Tot_Process) * 100, 0)))+'%'
    set @Start_Time = Getdate()
    set @Run_Proc_Msg = Substring(Convert(Varchar(4), @Process+100), 2, 2)+'.���� uSP_ETL_Hunderd_NonSale_NewStock�iETL �ʤj�s�~�P���ơj... '+@sProcess
    Update TB_Config set Flag = 1, Value = @sRebuild+@sProcess where kind = '0001'
    Exec @Get_Result = DW.dbo.uSP_ETL_Hunderd_NonSale_NewStock

    if @Get_Result = @Err_Code
       set @Run_Msg = @sFail
    else
       set @Run_Msg = @sSuss

    set @End_Time = Getdate()
    set @sStr = @Run_Proc_Msg +'..'+@Run_Msg+'!!..����'+ Convert(Varchar(20), @end_time - @Start_Time, 114) 
    set @Mail_1 = @Mail_1 + @sStr+@CR
print '46-----------------------------------------------------------------------------------------------'
print 'Mail Body:'+@Mail_1
print '-----------------------------------------------------------------------------------------------'
    Exec uSP_Sys_Write_Log @Proc, @sStr, '', @Get_Result, 1
*/    
  end try
  begin catch
    if (select cursor_status('global', 'Cur_Agent_Jobs')) <> -3 Close Cur_Agent_Jobs
    Set @Run_Msg = '���� '+@Run_Proc_Msg+'...�o���Y�����~!!���~�T���G'+@CR+ERROR_MESSAGE()
    Set @Mail_1 = @Mail_1 + @Run_Msg
print '48-----------------------------------------------------------------------------------------------'
print 'Mail Body:'+@Mail_1
print '-----------------------------------------------------------------------------------------------'
    Exec uSP_Sys_Write_Log @Proc, @Run_Msg, '', @Run_Err, 1
  end catch
  ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
  ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
  -- 2015/11/13 Rickliu ���ޤW�z����O�_���~�A���n���s�ҥαƵ{�P Trigger�A�]���~�|�W�ߤ@�� Try..catch.
  Begin Try
    Set @Run_Proc_Msg = '���s�ҥ� SQL �Ƶ{'
    Open Cur_Agent_Jobs
    Fetch Next From Cur_Agent_Jobs Into @sJob_id, @sJob_Name, @enabled
    While @@Fetch_Status = 0
    begin
      -- set @sStr = '���� ['+@Proc+']�����A�N�ҥ� SQL �Ƶ{ ['+@sJob_Name+']!!..'+convert(varchar(19), getdate(), 121)
      --set @Mail_1 = @Mail_1 + @sStr+@CR
      set @Start_Time = Getdate()
      set @Run_Proc_Msg = '�ҥαƵ{ ['+@sJob_Name+']'
      set @strSQL = 'EXEC msdb.dbo.sp_update_job @job_id=N'''+@sJob_id+''', @enabled=1 '
      Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, Run_Proc_Msg, @strSQL, 0

      if @Get_Result = @Err_Code
         set @Run_Msg = @sFail
      else
         set @Run_Msg = @sSuss
      
      set @End_Time = Getdate()
      set @sStr = @Run_Proc_Msg+'..'+@Run_Msg+'!!..����'+ Convert(Varchar(20), @end_time - @Start_Time, 114) 
      set @Mail_1 = @Mail_1 + @sStr+@CR       
print '49-----------------------------------------------------------------------------------------------'
print 'Mail Body:'+@Mail_1
print '-----------------------------------------------------------------------------------------------'

      Exec uSP_Sys_Write_Log @Proc, @sStr, '', 0

      Fetch Next From Cur_Agent_Jobs Into @sJob_id, @sJob_Name, @enabled
    end
    Close Cur_Agent_Jobs
    Deallocate Cur_Agent_Jobs

    set @sProcess = Convert(Varchar(4), Convert(Int, Round((@Process / @Tot_Process) * 100, 0)))+'%'
    set @Run_Proc_Msg = Substring(Convert(Varchar(4), @Process+100), 2, 2)+'.�t�έ���w����!!... '+@sProcess
    Update TB_Config set Flag = 0, Value = '�t�έ���w����!!..'+@sProcess where kind = '0001'
    Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, '', 0, 1
 
--    exec TYEIPDBS2.lytdbTA13.dbo.SP_Trigger_Control 'lytdbTA13', 'pcust', 'Tri_TA13_UPD_Pcust_Same_Custom', 1 -- �}�� Trigger
  end try
  begin catch
    Set @Run_Msg = '���� '+@Run_Proc_Msg+'...�o���Y�����~!!���~�T���G'+@CR+ERROR_MESSAGE()
    Set @Mail_1 = @Mail_1 + @Run_Msg+@CR
print '50-----------------------------------------------------------------------------------------------'
print 'Mail Body:'+@Mail_1
print '-----------------------------------------------------------------------------------------------'
    Exec uSP_Sys_Write_Log @Proc, @Run_Msg, '', @Err_Code, 1
    set @Get_Result = @Err_Code
  end catch

End_Proc:
  if (select cursor_status('global', 'Cur_Agent_Jobs')) <> -3 
  begin
    Close Cur_Agent_Jobs
    Deallocate Cur_Agent_Jobs
  end

  ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
  ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
  -- ���� uSP_Sys_Release_Memory�i�����Ʈw�O����Ŷ��j

  set @Run_Proc_Msg = '���� uSP_Sys_Release_Memory�i�����Ʈw�O����Ŷ��j... '+@sProcess
  Exec @Get_Result = DW.dbo.uSP_Sys_Release_Memory 

  ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
  ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
  Set @Run_Msg = @Proc+' ' +@Run_Msg +'...���浲���ɶ��G[ '+CONVERT(Varchar(100), GETDATE(),120)+' ].'
  Set @Mail_1 = @Mail_1 +@Run_Msg+@sProcess+@CR
  -- �o�e MAIL �������H��
  Exec uSP_Sys_Write_Log @Proc, @Run_Msg, @Mail_1, 0, 1
  
print '51-----------------------------------------------------------------------------------------------'
print 'Mail Body:'+@Mail_1
print '-----------------------------------------------------------------------------------------------'
  
  if @Get_Result <> 0 
     set @Result = @Err_Code
  Return(@Result)
End
GO
