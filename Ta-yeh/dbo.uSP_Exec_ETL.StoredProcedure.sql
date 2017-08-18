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
  Declare @Proc_SubName NVarchar(Max) = 'ETL 資料倉儲資料分解作業'
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
  Declare @sRebuild Varchar(20) = '重整中!!..'
  Declare @sSuss Varchar(20) = '成功'
  Declare @sFail Varchar(20) = '失敗'
  Declare @Run_ETL Int = 0
  Declare @sJob_id NVarchar(100), @sJob_Name NVarchar(100), @enabled int

  Declare @Sch_Jobs Table(
    [job_id] [uniqueidentifier] NOT NULL,
    [name] [sysname] NOT NULL,
    [enabled] [tinyint] NOT NULL
  )

  set @Mail_1 = '此作業於每日上午06:00, 12:00 開始進行。'+@CR
  set @Mail_1 = @Mail_1+'準備進行 '+@Proc_SubName+@CR
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

    set @Run_Proc_Msg = '檢查凌越系統是否進行重整作業!!!'
    set @chk_lys = (select l_field2 from TYEIPDBS2.LYTDBTA13.dbo.ssyslb)
    if @chk_lys = 1
    begin
      Set @sStr = '目前正在執行凌越重整 ['+@Proc+']，所以無法執行本程序!!..'+convert(varchar(19), getdate(), 121)
      Set @Mail_1 = @Mail_1 + @sStr+@CR
      Exec uSP_Sys_Write_Log @Proc, @sStr, '', @Err_Code 
print '2-----------------------------------------------------------------------------------------------'
print 'Mail Body:'+@Mail_1
print '-----------------------------------------------------------------------------------------------'
      Goto End_Proc
    end

    set @Run_Proc_Msg = '將電視牆改為重整畫面!!!'
    Update TB_Config set Flag = 1, Value = @sRebuild+@sProcess where kind = '0001'
    Set @Run_Msg = @Proc+' ' +@Run_Msg +'...執行開始時間：[ '+CONVERT(Varchar(100), GETDATE(),120)+' ].'
    Set @Mail_1 = @Mail_1 +@Run_Msg+@sProcess+@CR
print '3-----------------------------------------------------------------------------------------------'
print 'Mail Body:'+@Mail_1
print '-----------------------------------------------------------------------------------------------'
    Exec uSP_Sys_Write_Log @Proc, @Run_Msg, '', 0, 1
    
    set @Start_Time = Getdate()
    set @Run_Proc_Msg = '執行 uSP_Sys_Release_Memory【釋放資料庫記憶體空間】... '+@sProcess
    Exec @Get_Result = DW.dbo.uSP_Sys_Release_Memory 
      
    ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
    -- 2014/11/28 Rickliu 只取得正在使用(enabled = 0)，避免未使用的排程於程序結束後又被開啟。又因資料筆數小於 100 筆，所以採用變數表格方式
    Open Cur_Agent_Jobs
    Fetch Next From Cur_Agent_Jobs Into @sJob_id, @sJob_Name, @enabled

    While @@Fetch_Status = 0
    begin
print '4-----------------------------------------------------------------------------------------------'
print 'Mail Body:'+@Mail_1
print '-----------------------------------------------------------------------------------------------'
      set @Start_Time = Getdate()
      set @Run_Proc_Msg = '停用排程 ['+Isnull(@sJob_Name, '')+']'
      set @strSQL = 'EXEC msdb.dbo.sp_update_job @job_id=N'''+@sJob_id+''', @enabled=0 '
      Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @sStr, @strSQL, 0

      if @Get_Result = @Err_Code
         set @Run_Msg = @sFail
      else
         set @Run_Msg = @sSuss

      set @End_Time = Getdate()
      set @sStr = @Run_Proc_Msg+'..'+@Run_Msg+'!!..執行'+ Convert(Varchar(20), @end_time - @Start_Time, 114) 
      set @Mail_1 = @Mail_1 + @sStr+@CR
print '5-----------------------------------------------------------------------------------------------'
print 'Mail Body:'+@Mail_1
print '-----------------------------------------------------------------------------------------------'
      Exec uSP_Sys_Write_Log @Proc, @sStr, '', 0

      Fetch Next From Cur_Agent_Jobs Into @sJob_id, @sJob_Name, @enabled
    end
    Close Cur_Agent_Jobs
  
    -- 以下執行順序不可以顛倒
    ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
    ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
    set @Process = 1
    set @sProcess = Convert(Varchar(4), Convert(Int, Round((@Process / @Tot_Process) * 100, 0)))+'%'
    set @Start_Time = Getdate()
    set @Run_Proc_Msg = Substring(Convert(Varchar(4), @Process+100), 2, 2)+'.執行 uSP_ETL_New_Stock_Lists【更新商品的新品資訊】... '+@sProcess
    set @Run_ETL = 0
    Update TB_Config set Flag = 1, Value = @sRebuild+@sProcess where kind = '0001'
    Exec @Get_Result = DW.dbo.uSP_ETL_New_Stock_Lists @Run_ETL
      
    if @Get_Result = @Err_Code
       set @Run_Msg = @sFail
    else
       set @Run_Msg = @sSuss

    set @End_Time = Getdate()
    set @sStr = @Run_Proc_Msg +'..'+@Run_Msg+'!!..執行'+ Convert(Varchar(20), @end_time - @Start_Time, 114) 
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
    set @Run_Proc_Msg = Substring(Convert(Varchar(4), @Process+100), 2, 2)+'.執行 uSP_ETL_PEMPLOY【PEMPLOY 員工資料表進行分解】... '+@sProcess
    Update TB_Config set Flag = 1, Value = @sRebuild+@sProcess where kind = '0001'
    Exec @Get_Result = DW.dbo.uSp_ETL_pemploy

    if @Get_Result = @Err_Code
       set @Run_Msg = @sFail
    else
       set @Run_Msg = @sSuss

    set @End_Time = Getdate()
    set @sStr = @Run_Proc_Msg +'..'+@Run_Msg+'!!..執行'+ Convert(Varchar(20), @end_time - @Start_Time, 114) 
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
    set @Run_Proc_Msg = Substring(Convert(Varchar(4), @Process+100), 2, 2)+'.執行 uSP_ETL_PCUST【PCUST 廠商/客戶資料表進行分解】... '+@sProcess
    Update TB_Config set Flag = 1, Value = @sRebuild+@sProcess where kind = '0001'
    Exec @Get_Result = DW.dbo.uSp_ETL_pcust

    if @Get_Result = @Err_Code
       set @Run_Msg = @sFail
    else
       set @Run_Msg = @sSuss

    set @End_Time = Getdate()
    set @sStr = @Run_Proc_Msg +'..'+@Run_Msg+'!!..執行'+ Convert(Varchar(20), @end_time - @Start_Time, 114) 
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
    set @Run_Proc_Msg = Substring(Convert(Varchar(4), @Process+100), 2, 2)+'.執行 uSP_ETL_SWAREDT【SWAREDT 庫存資料表進行分解】... '+@sProcess
    Update TB_Config set Flag = 1, Value = @sRebuild+@sProcess where kind = '0001'
    Exec @Get_Result = DW.dbo.uSp_ETL_SWAREDT

    if @Get_Result = @Err_Code
       set @Run_Msg = @sFail
    else
       set @Run_Msg = @sSuss

    set @End_Time = Getdate()
    set @sStr = @Run_Proc_Msg +'..'+@Run_Msg+'!!..執行'+ Convert(Varchar(20), @end_time - @Start_Time, 114) 
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
    set @Run_Proc_Msg = Substring(Convert(Varchar(4), @Process+100), 2, 2)+'.執行 uSP_ETL_SSTOCK【SSTOCK 商品資料表進行分解】... '+@sProcess
    Update TB_Config set Flag = 1, Value = @sRebuild+@sProcess where kind = '0001'
    Exec @Get_Result = DW.dbo.uSp_ETL_sstock

    if @Get_Result = @Err_Code
       set @Run_Msg = @sFail
    else
       set @Run_Msg = @sSuss

    set @End_Time = Getdate()
    set @sStr = @Run_Proc_Msg +'..'+@Run_Msg+'!!..執行'+ Convert(Varchar(20), @end_time - @Start_Time, 114) 
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
    set @Run_Proc_Msg = Substring(Convert(Varchar(4), @Process+100), 2, 2)+'.執行 uSP_ETL_SSLIP【SSLIP 各類單據表頭資料表進行分解】... '+@sProcess
    Update TB_Config set Flag = 1, Value = @sRebuild+@sProcess where kind = '0001'
    Exec @Get_Result = DW.dbo.uSP_ETL_sslip

    if @Get_Result = @Err_Code
       set @Run_Msg = @sFail
    else
       set @Run_Msg = @sSuss

    set @End_Time = Getdate()
    set @sStr = @Run_Proc_Msg +'..'+@Run_Msg+'!!..執行'+ Convert(Varchar(20), @end_time - @Start_Time, 114) 
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
    set @Run_Proc_Msg = Substring(Convert(Varchar(4), @Process+100), 2, 2)+'.執行 uSP_ETL_SSLPDT【SSLPDT 各類單據明細資料表進行分解】... '+@sProcess
    Update TB_Config set Flag = 1, Value = @sRebuild+@sProcess where kind = '0001'
    Exec @Get_Result = DW.dbo.uSP_ETL_sslpdt

    if @Get_Result = @Err_Code
       set @Run_Msg = @sFail
    else
       set @Run_Msg = @sSuss

    set @End_Time = Getdate()
    set @sStr = @Run_Proc_Msg +'..'+@Run_Msg+'!!..執行'+ Convert(Varchar(20), @end_time - @Start_Time, 114) 
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
    set @Run_Proc_Msg = Substring(Convert(Varchar(4), @Process+100), 2, 2)+'.執行 uSP_ETL_ReIndexs【ETL 重建建立索引欄位】... '+@sProcess
    Update TB_Config set Flag = 1, Value = @sRebuild+@sProcess where kind = '0001'
    Exec @Get_Result = DW.dbo.uSP_ETL_ReIndexs2 1

    if @Get_Result = @Err_Code
       set @Run_Msg = @sFail
    else
       set @Run_Msg = @sSuss

    set @End_Time = Getdate()
    set @sStr = @Run_Proc_Msg +'..'+@Run_Msg+'!!..執行'+ Convert(Varchar(20), @end_time - @Start_Time, 114) 
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
    set @Run_Proc_Msg = Substring(Convert(Varchar(4), @Process+100), 2, 2)+'.執行 uSP_Realtime_SaleData【ETL RealTime 即時資料】... '+@sProcess
    Update TB_Config set Flag = 1, Value = @sRebuild+@sProcess where kind = '0001'
    --Exec @Get_Result = DW.dbo.uSP_Realtime_SaleData '', 1 -- 開啟發送 Log 機制
    Exec @Get_Result = DW.dbo.uSP_Realtime_SaleData '', 2
    
    if @Get_Result = @Err_Code
       set @Run_Msg = @sFail
    else
       set @Run_Msg = @sSuss

    set @End_Time = Getdate()
    set @sStr = @Run_Proc_Msg +'..'+@Run_Msg+'!!..執行'+ Convert(Varchar(20), @end_time - @Start_Time, 114) 
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
    set @Run_Proc_Msg = Substring(Convert(Varchar(4), @Process+100), 2, 2)+'.執行 uSP_ETL_Cust_NonSale_Stock【ETL 所有客戶未銷商品資料】... '+@sProcess
    Update TB_Config set Flag = 1, Value = @sRebuild+@sProcess where kind = '0001'
    Exec @Get_Result = DW.dbo.uSP_ETL_Cust_NonSale_Stock

    if @Get_Result = @Err_Code
       set @Run_Msg = @sFail
    else
       set @Run_Msg = @sSuss

    set @End_Time = Getdate()
    set @sStr = @Run_Proc_Msg +'..'+@Run_Msg+'!!..執行'+ Convert(Varchar(20), @end_time - @Start_Time, 114) 
    set @Mail_1 = @Mail_1 + @sStr+@CR
print '42-----------------------------------------------------------------------------------------------'
print 'Mail Body:'+@Mail_1
print '-----------------------------------------------------------------------------------------------'
    Exec uSP_Sys_Write_Log @Proc, @sStr, '', @Get_Result, 1
*/
-- 2017/08/07 Rickliu 取消使用
/*
    ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
    ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
    set @Process = 11
    set @sProcess = Convert(Varchar(4), Convert(Int, Round((@Process / @Tot_Process) * 100, 0)))+'%'
    set @Start_Time = Getdate()
    set @Run_Proc_Msg = Substring(Convert(Varchar(4), @Process+100), 2, 2)+'.執行 uSP_ETL_Hunderd_NonSale_Stock【ETL 百大客戶未銷商品資料】... '+@sProcess
    Update TB_Config set Flag = 1, Value = @sRebuild+@sProcess where kind = '0001'
    Exec @Get_Result = DW.dbo.uSP_ETL_Hunderd_NonSale_Stock

    if @Get_Result = @Err_Code
       set @Run_Msg = @sFail
    else
       set @Run_Msg = @sSuss

    set @End_Time = Getdate()
    set @sStr = @Run_Proc_Msg +'..'+@Run_Msg+'!!..執行'+ Convert(Varchar(20), @end_time - @Start_Time, 114) 
    set @Mail_1 = @Mail_1 + @sStr+@CR
print '43-----------------------------------------------------------------------------------------------'
print 'Mail Body:'+@Mail_1
print '-----------------------------------------------------------------------------------------------'
    Exec uSP_Sys_Write_Log @Proc, @sStr, '', @Get_Result, 1
*/
-- 2017/08/07 Rickliu 取消使用
/*
    ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
    ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
    set @Process = 12
    set @sProcess = Convert(Varchar(4), Convert(Int, Round((@Process / @Tot_Process) * 100, 0)))+'%'
    set @Start_Time = Getdate()
    set @Run_Proc_Msg = Substring(Convert(Varchar(4), @Process+100), 2, 2)+'.執行 uSP_ETL_Cust_WeekSale_Stock【ETL 百大客戶未銷商品資料(週)】... '+@sProcess
    Update TB_Config set Flag = 1, Value = @sRebuild+@sProcess where kind = '0001'
    Exec @Get_Result = DW.dbo.uSP_ETL_Cust_WeekSale_Stock

    if @Get_Result = @Err_Code
       set @Run_Msg = @sFail
    else
       set @Run_Msg = @sSuss

    set @End_Time = Getdate()
    set @sStr = @Run_Proc_Msg +'..'+@Run_Msg+'!!..執行'+ Convert(Varchar(20), @end_time - @Start_Time, 114) 
    set @Mail_1 = @Mail_1 + @sStr+@CR
print '44-----------------------------------------------------------------------------------------------'
print 'Mail Body:'+@Mail_1
print '-----------------------------------------------------------------------------------------------'
    Exec uSP_Sys_Write_Log @Proc, @sStr, '', @Get_Result, 1
*/
-- 2017/08/07 Rickliu 取消使用
/*
	---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
    ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
    set @Process = 13
    set @sProcess = Convert(Varchar(4), Convert(Int, Round((@Process / @Tot_Process) * 100, 0)))+'%'
    set @Start_Time = Getdate()
    set @Run_Proc_Msg = Substring(Convert(Varchar(4), @Process+100), 2, 2)+'.執行 uSP_ETL_Cust_MonthSale_Stock【ETL 百大客戶未銷商品資料(月)】... '+@sProcess
    Update TB_Config set Flag = 1, Value = @sRebuild+@sProcess where kind = '0001'
    Exec @Get_Result = DW.dbo.uSP_ETL_Cust_MonthSale_Stock

    if @Get_Result = @Err_Code
       set @Run_Msg = @sFail
    else
       set @Run_Msg = @sSuss

    set @End_Time = Getdate()
    set @sStr = @Run_Proc_Msg +'..'+@Run_Msg+'!!..執行'+ Convert(Varchar(20), @end_time - @Start_Time, 114) 
    set @Mail_1 = @Mail_1 + @sStr+@CR
print '45-----------------------------------------------------------------------------------------------'
print 'Mail Body:'+@Mail_1
print '-----------------------------------------------------------------------------------------------'
    Exec uSP_Sys_Write_Log @Proc, @sStr, '', @Get_Result, 1
*/
-- 2017/08/07 Rickliu 取消使用
/*
    ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
    ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
    set @Process = 14
    set @sProcess = Convert(Varchar(4), Convert(Int, Round((@Process / @Tot_Process) * 100, 0)))+'%'
    set @Start_Time = Getdate()
    set @Run_Proc_Msg = Substring(Convert(Varchar(4), @Process+100), 2, 2)+'.執行 uSP_ETL_Hunderd_NonSale_NewStock【ETL 百大新品銷售資料】... '+@sProcess
    Update TB_Config set Flag = 1, Value = @sRebuild+@sProcess where kind = '0001'
    Exec @Get_Result = DW.dbo.uSP_ETL_Hunderd_NonSale_NewStock

    if @Get_Result = @Err_Code
       set @Run_Msg = @sFail
    else
       set @Run_Msg = @sSuss

    set @End_Time = Getdate()
    set @sStr = @Run_Proc_Msg +'..'+@Run_Msg+'!!..執行'+ Convert(Varchar(20), @end_time - @Start_Time, 114) 
    set @Mail_1 = @Mail_1 + @sStr+@CR
print '46-----------------------------------------------------------------------------------------------'
print 'Mail Body:'+@Mail_1
print '-----------------------------------------------------------------------------------------------'
    Exec uSP_Sys_Write_Log @Proc, @sStr, '', @Get_Result, 1
*/    
  end try
  begin catch
    if (select cursor_status('global', 'Cur_Agent_Jobs')) <> -3 Close Cur_Agent_Jobs
    Set @Run_Msg = '執行 '+@Run_Proc_Msg+'...發生嚴重錯誤!!錯誤訊息：'+@CR+ERROR_MESSAGE()
    Set @Mail_1 = @Mail_1 + @Run_Msg
print '48-----------------------------------------------------------------------------------------------'
print 'Mail Body:'+@Mail_1
print '-----------------------------------------------------------------------------------------------'
    Exec uSP_Sys_Write_Log @Proc, @Run_Msg, '', @Run_Err, 1
  end catch
  ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
  ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
  -- 2015/11/13 Rickliu 不管上述執行是否錯誤，都要重新啟用排程與 Trigger，因此才會獨立一個 Try..catch.
  Begin Try
    Set @Run_Proc_Msg = '重新啟用 SQL 排程'
    Open Cur_Agent_Jobs
    Fetch Next From Cur_Agent_Jobs Into @sJob_id, @sJob_Name, @enabled
    While @@Fetch_Status = 0
    begin
      -- set @sStr = '執行 ['+@Proc+']完成，將啟用 SQL 排程 ['+@sJob_Name+']!!..'+convert(varchar(19), getdate(), 121)
      --set @Mail_1 = @Mail_1 + @sStr+@CR
      set @Start_Time = Getdate()
      set @Run_Proc_Msg = '啟用排程 ['+@sJob_Name+']'
      set @strSQL = 'EXEC msdb.dbo.sp_update_job @job_id=N'''+@sJob_id+''', @enabled=1 '
      Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, Run_Proc_Msg, @strSQL, 0

      if @Get_Result = @Err_Code
         set @Run_Msg = @sFail
      else
         set @Run_Msg = @sSuss
      
      set @End_Time = Getdate()
      set @sStr = @Run_Proc_Msg+'..'+@Run_Msg+'!!..執行'+ Convert(Varchar(20), @end_time - @Start_Time, 114) 
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
    set @Run_Proc_Msg = Substring(Convert(Varchar(4), @Process+100), 2, 2)+'.系統重整已完成!!... '+@sProcess
    Update TB_Config set Flag = 0, Value = '系統重整已完成!!..'+@sProcess where kind = '0001'
    Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, '', 0, 1
 
--    exec TYEIPDBS2.lytdbTA13.dbo.SP_Trigger_Control 'lytdbTA13', 'pcust', 'Tri_TA13_UPD_Pcust_Same_Custom', 1 -- 開啟 Trigger
  end try
  begin catch
    Set @Run_Msg = '執行 '+@Run_Proc_Msg+'...發生嚴重錯誤!!錯誤訊息：'+@CR+ERROR_MESSAGE()
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
  -- 執行 uSP_Sys_Release_Memory【釋放資料庫記憶體空間】

  set @Run_Proc_Msg = '執行 uSP_Sys_Release_Memory【釋放資料庫記憶體空間】... '+@sProcess
  Exec @Get_Result = DW.dbo.uSP_Sys_Release_Memory 

  ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
  ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
  Set @Run_Msg = @Proc+' ' +@Run_Msg +'...執行結束時間：[ '+CONVERT(Varchar(100), GETDATE(),120)+' ].'
  Set @Mail_1 = @Mail_1 +@Run_Msg+@sProcess+@CR
  -- 發送 MAIL 給相關人等
  Exec uSP_Sys_Write_Log @Proc, @Run_Msg, @Mail_1, 0, 1
  
print '51-----------------------------------------------------------------------------------------------'
print 'Mail Body:'+@Mail_1
print '-----------------------------------------------------------------------------------------------'
  
  if @Get_Result <> 0 
     set @Result = @Err_Code
  Return(@Result)
End
GO
