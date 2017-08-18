USE [DW]
GO
/****** Object:  StoredProcedure [dbo].[uSP_Imp_Absent]    Script Date: 08/18/2017 17:43:40 ******/
DROP PROCEDURE [dbo].[uSP_Imp_Absent]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
/*********************************************************************************
--考勤下班時間
select * from  apsnt_bk where ps_date='2014/1/2'

update apsnt_bk
set ps_tm1e=b.ps_time
from  (
select ps_time,e_no from [dbo].[Ori_Txt#Apsnt_Tmp_1]  where  e_no='T09092' ) b
where ps_no collate Chinese_Taiwan_Stroke_CI_AS =e_no collate Chinese_Taiwan_Stroke_CI_AS
and ps_date='2013/12/13'
**********************************************************************************/

--Exec uSP_Imp_Apsnt
CREATE Procedure [dbo].[uSP_Imp_Absent]
  @ps_date datetime = null
as
begin
  /*************************************************************************************************************
   2013/12/21 與李協理確認，周一至週五及周日都設定為正常班時段，周六則下班提早一小時，出勤僅看上班及下班，
              加班、外出都不透過卡機判別。
              上班僅研判最早出勤時間，下班則判別最晚下班時間。
              正常班：07:00~17:45, 周六班：07:00~16:45
  *************************************************************************************************************/
  -- 自身程序相關設定
  Declare @Proc Varchar(50) = 'uSP_Imp_Absent'
  Declare @Cnt Int = 0
  Declare @RowCnt Int = 0
  Declare @Msg Varchar(4000) =''
  Declare @strSQL Varchar(Max)
  --20131228 add by brian 增加下班時間
  Declare @strSQL1 Varchar(Max)
 
  Declare @CR Varchar(5) = ' '+char(13)+char(10)
  Declare @RowCount Table (cnt int)
  --Declare @rDB Varchar(50) = 'TYEIPDBS2.lytdbTA13.dbo.'
  Declare @rDB Varchar(50) = 'SYNC_TA13.dbo.' -- 2017/03/17 Rickliu 改為以訂閱資料庫為參考資料
  Declare @wDB Varchar(50) = 'TYEIPDBS2.lytdbTA13.dbo.' -- 2017/03/17 Rickliu 新增回寫目的資料庫
  
  --set @rDB = @wDB -- Rickliu 2017/03/24 在訂閱服務還沒好時，暫時使用 @wDB 資料庫的資料
  
  Declare @Errcode int = -1

  -- 考勤相關設定
  Declare @Proce_SubName NVarchar(Max) = '考勤資料匯入程序'
  Declare @Body_Msg NVarchar(Max) = ''
  Declare @Run_Msg NVarchar(Max) = ''
  
  -- 檔案相關設定
  Declare @file_Name varchar(50) = ''
  Declare @file_ext varchar(10) = '.Txt'
  Declare @file_path varchar(255) = 'D:\Transform_Data\Import_Absent\'
  Declare @isExists Int -- 判讀檔案是否存在 1: 存在, 2: 不存在
  
  -- Declare @txt_head varchar(255) = 'Ori_Txt#' 
  Declare @txt_head varchar(255) = '' -- 2017/03/17 Rickliu 改為以訂閱資料庫為參考資料
  Declare @tb_Name varchar(255) ='Apsnt'
  Declare @tb_Ori_Apsnt varchar(255) = @txt_head + @tb_Name -- Ex: Ori_Txt#Apsnt -- Text 出勤轉入介面檔
  Declare @tb_Apsnt_tmp varchar(255) = @tb_Ori_Apsnt + '_Tmp' -- Ex: Ori_Txt#Apsnt_Tmp -- 出勤資料處理暫存檔
  
  -- 考勤時間相關設定
  Declare @Time_Over Varchar(5) = '00:00' -- 考勤跨夜時間
  
  Declare @Time1 Varchar(5) = '07:00' -- 上班起算時間
  Declare @Car1 Varchar(1) = '1' -- 上班編號

  Declare @Time2 Varchar(5) = '18:00' -- 下班結束時間
  Declare @Car2 Varchar(1) = '2' -- 下班編號
  
  Declare @Time3 Varchar(5) = '19:00' -- 加班起算時間
  Declare @Car3 Varchar(1) = '3' -- 加班上班編號

  Declare @Time4 Varchar(5) = '06:29' -- 加班結束時間
  Declare @Car4 Varchar(1) = '4' -- 加班下班編號
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @ps_date = isnull(@ps_date, getdate())
  set @File_Name = @File_path + Convert(varchar(12), @ps_date, 112) + @File_ext -- Ex: D:\考勤轉檔\20131210.Txt
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*

  --Exec uSP_Advanced_Options 1
  
  IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].['+@tb_Ori_Apsnt+']') AND type in (N'U'))
  begin
     set @Msg = '1.清除 出勤轉入介面檔['+@tb_Ori_Apsnt+'].'
     set @Cnt = 0
     set @strSQL = 'Drop Table '+@tb_Ori_Apsnt
     Exec @Cnt = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
     if @Cnt = @Errcode Goto End_Proc
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Msg = '2.建立 出勤轉入介面檔['+@tb_Ori_Apsnt+'].'
  set @Cnt = 0
  set @strSQL = 'create table '+@tb_Ori_Apsnt +@Cr+
                '  (e_no Varchar(10),'+@Cr+
                '   ps_week varchar(10),'+@Cr+
                '   ps_date datetime,'+@Cr+
                '   ps_time varchar(10),'+@Cr+
                '   ps_class int,'+@Cr+
                '   e_name varchar(20), '+@Cr+
                '   Door_No varchar(5),  '+@Cr+ -- Rickliu 2016/12/21 新增卡機編號與名稱
                '   Door_Name varchar(50)) '
  Exec @Cnt = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
  if @Cnt = @Errcode Goto End_Proc
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Msg = '3.判讀考勤檔案['+@File_Name+']是否存在.'
  set @Cnt = 0
  exec master.dbo.xp_fileexist @File_Name, @isExists OUTPUT
  if @isExists <> 1
  begin
     set @Cnt = @Errcode
     set @Msg = @Msg + '...檔案不存在.'
     Exec uSP_Sys_Write_Log @Proc, @Msg, '', 0
     if @Cnt = @Errcode goto End_Proc
  end
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Msg = '4.執行匯入 出勤資料至轉入介面檔 ['+@File_Path+'] ==> ['+@tb_Ori_Apsnt+'].'
  set @Cnt = 0
  set @strSQL = 'BULK INSERT '+@tb_Ori_Apsnt +@Cr+
                '       FROM '''+@File_Name+''' WITH(FIELDTERMINATOR='' '' , ROWTERMINATOR=''\n'', TABLOCK )'
 
  Exec @Cnt = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
  if @Cnt = @Errcode Goto End_Proc
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Msg = '5.判斷 ['+@tb_Ori_Apsnt+'] 是否有資料.'
  set @Cnt = 0
  set @strSQL = 'select Count(1) as Cnt From '+@tb_Ori_Apsnt
  
  delete @RowCount
  insert into @RowCount exec (@strSQL)
  set @Cnt =(select Cnt from @RowCount)
  set @Msg = @Msg + '...匯入筆數 ['+cast(@Cnt as varchar)+']'
  
  If @Cnt = 0 
  begin
     Set @Msg = @Msg + '..無資料，取消執行!!'
     Exec uSP_Sys_Write_Log @Proc, @Msg, '', 0
     Goto End_Exit
  end
  Exec uSP_Sys_Write_Log @Proc, @Msg, '', 0
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].['+@tb_Apsnt_tmp+']') AND type in (N'U'))
  begin
     set @Msg = '6.重建 出勤資料處理暫存檔['+@tb_Apsnt_tmp+'].'
     set @Cnt = 0
     set @strSQL = 'Drop Table '+@tb_Apsnt_tmp
     Exec @Cnt = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
  end
  if @Cnt = @Errcode Goto End_Proc
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Msg = '7.重建 出勤資料處理暫存檔['+@tb_Apsnt_tmp+'].'
  set @Cnt = 0
  -- 2014/07/09 Rickliu 因出勤補正單所以會更改刷卡紀錄，而為了保留原始刷卡紀錄，於以儲存至 ps_tm3s, ps_tm3e, 
  --                    程式若判斷到 ps_tm4s, ps_tm4e 有時間則代表有進行出勤補正動作，所以會取代掉 ps_tm1s, ps_tm2s
  /***********************************************************************************************************************
  -- 2017/02/09 Rickliu 因公司為了因應政府一例一休政策，且主張不鼓勵員工加班，下班時間到了就應準時下班，但又無法透過人員勸
                        導員工即時下班且又深怕勞動部抽查情況。因此馬董事長及阮總經理指示要求，透過系統程式集體修訂應下班人
                        員而未準時下班者一率修訂出勤下班時間為 18:00。
                        資訊課長 RICKLIU 曾透過多次會議反應，公司修改出勤紀錄將會觸犯偽造文書及登載不實文書罪，此舉萬不可
                        行，但高層仍堅定要求此作法，因此，本部也只能無奈配合公司政策。
  ***********************************************************************************************************************/
  
  set @strSQL = ';With CTE_Q1 as ( '+@Cr+
                '   select e_no, e_name, '+@Cr+
                '          ps_date, '+@Cr+
                '          min(ps_time) as ps_time, '+@Cr+
                '          case '+@Cr+
                '            when datediff(mi, min(ps_time), max(ps_time)) < 30 then null '+@Cr+
                '            else max(ps_time) '+@Cr+
                '          end as ps_time2, '+@Cr+
                '          case '+@Cr+
                '            when max(ps_time) between ''00:00'' and ''07:00'' then max(ps_time) '+@Cr+
                '            else null '+@Cr+
                '          end as ps_time3 '+@Cr+
                '    from '+@tb_Ori_Apsnt+@Cr+
                '   group by e_no, e_name, ps_date '+@Cr+
                '), CTE_Q2 as ( '+@Cr+
                '  select distinct e_no, e_name, '+@Cr+
                '         substring(datename(weekday, ps_date), 3, 1) as ps_week, '+@Cr+
                '         ps_date, '+@Cr+
                '         max(ps_time) as ps_time, '+@Cr+
                '         max(ps_time2) as ps_time2, '+@Cr+
                '         max(ps_time3) as ps_time3 '+@Cr+
                '    from CTE_Q1 '+@Cr+
                '   group by e_no, e_name, ps_date '+@Cr+
                '), CTE_Q3 as ('+@Cr+
                '  select distinct e_no, e_dept '+@Cr+
                '    from '+@rDB+'pemploy '+@Cr+
                '), CTE_Q4 as ('+@Cr+
                '  select ps_date, min(ps_time) as ps_min_time, max(ps_time) as ps_max_time '+@Cr+
                '    from '+@tb_Ori_Apsnt+@Cr+
                '   group by ps_date '+@Cr+
                ') '+@Cr+
                ' '+@Cr+
                'select distinct m.ps_date as A_ps_date, m.ps_week, m.e_no as e_no, m.e_name, '+@Cr+
                -- 2017/02/09 Rickliu 員工上班打卡紀錄
                '       case  '+@Cr+
                '         when (Replace(isnull(d1.ps_tm4s, ''''), '':'', '''') <> '''') then Convert(Varchar(5), Convert(Time, Replace(Rtrim(d1.ps_tm4s), '':'', ''''))) '+@Cr+
                '         else isnull(m.ps_time, '''') '+@Cr+
                '       end as ps_time_s, '+@Cr+
                -- 2017/02/09 Rickliu 員工下班打卡紀錄
                '       case '+@Cr+
                -- 2017/02/09 Rickliu 若人事單位有填寫凌越出勤記錄維護之大夜班下班欄位則代表有下班出勤補正動作，系統則回填出勤補正之下班時間。
                '         when isnull(m.ps_time2, '''')= '''' And (Replace(isnull(d1.ps_tm4e, ''''), '':'', '''') <> '''') then Convert(Varchar(5), Convert(Time, Replace(Rtrim(d1.ps_tm4e), '':'', ''''))) '+@Cr+
                -- 2017/02/09 Rickliu 若超過 18:30 未打卡者，系統一率自動設定 18:30 打卡
                '         when (Replace(isnull(isnull(isnull(m.ps_time2, d1.ps_tm3e), d1.ps_tm4e), ''''), '':'', '''') = '''') And (Convert(Varchar(5), Convert(Time, getdate())) >= ''18:30'') then ''18:30'' '+@Cr+
                -- 2017/02/09 Rickliu 員工若已經打完卡且時間超過 18:30 者，系統一率自動設定 18:30 打卡。
                '         when (Convert(Varchar(5), Convert(Time, Replace(Rtrim(d1.ps_tm3e), '':'', ''''))) > ''18:30'') then ''18:30'' '+@Cr+
                -- 2017/02/09 Rickliu 無下班卡且無原始刷卡且無出勤補正，但該日公司確有出勤紀錄，則代表該員忘記打卡，因此直接回填 18:30
                '         when (Replace(isnull(isnull(isnull(m.ps_time2, d1.ps_tm3e), d1.ps_tm4e), ''''), '':'', '''') = '''') and (isnull(ps_max_time, '''') >= ''18:30'') then ''18:30'' '+@Cr+
                -- 2017/02/09 Rickliu 以上條件都不符合則保留原值
                '         else isnull(m.ps_time2, '''') '+@Cr+
                '       end as ps_time_e, '+@Cr+
                -- 2017/02/09 Rickliu 上班時間2
                '       isnull(m.ps_time3, '''') as ps_time2_s, '+@Cr+
                -- 2017/02/09 Rickliu 下班時間2
                '       '''' as ps_time2_e, '+@Cr+
                -- 2017/02/09 Rickliu 仍保留原始上班打卡紀錄
                '       isnull(m.ps_time, '''') as ps_time3_s, '+@Cr+
                -- 2017/02/09 Rickliu 仍保留原始下班打卡紀錄
                '       isnull(m.ps_time2, '''') as ps_time3_e, '+@Cr+
                -- 2017/02/09 Rickliu 上班時間4(出勤補正單之上班打卡紀錄)
                '       case  '+@Cr+
                '         when (Replace(isnull(d1.ps_tm4s, ''''), '':'', '''') <> '''') then Convert(Varchar(5), Convert(Time, Replace(Rtrim(d1.ps_tm4s), '':'', ''''))) '+@Cr+
                '         else '''' '+@Cr+
                '       end as ps_time4_s, '+@Cr+
                -- 2017/02/09 Rickliu 上班時間4(出勤補正單之下班打卡紀錄)
                '       case '+@Cr+
                '         when (Replace(isnull(d1.ps_tm4e, ''''), '':'', '''') <> '''') then Convert(Varchar(5), Convert(Time, Replace(Rtrim(d1.ps_tm4e), '':'', ''''))) '+@Cr+
                '         else '''' '+@Cr+
                '       end as ps_time4_e, '+@Cr+
                '       d.e_dept '+@Cr+
                '       into '+@tb_Apsnt_tmp+@Cr+
                '  from CTE_Q2 m '+@Cr+
                '       left join CTE_Q3 d on m.e_no collate Chinese_Taiwan_Stroke_CI_AS =d.e_no collate Chinese_Taiwan_Stroke_CI_AS '+@Cr+
                '       left join '+@rDB+'Apsnt d1 on m.e_no collate Chinese_Taiwan_Stroke_CI_AS =d1.ps_no collate Chinese_Taiwan_Stroke_CI_AS and m.ps_date = d1.ps_date '+@Cr+
                '       left join CTE_Q4 d2 on m.ps_date = d2.ps_date '

  Exec @Cnt = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
  if @Cnt = @Errcode Goto End_Proc


  Set @Msg = '7.1.出勤資料處理暫存檔['+@tb_Apsnt_tmp+'] 產出筆數為:'
  set @strSQL = 'select Count(1) as Cnt From '+@tb_Apsnt_tmp
  Exec @Cnt = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL

  set @Msg = '8.清除 ['+Convert(varchar(12), @ps_date, 111)+'] 日凌越系統出勤資料.'
  set @Cnt = 0
  set @strSQL = -- 'Delete '+@rDB+@tb_Name+@Cr+
                'Delete '+@wDB+@tb_Name+@Cr+
                '  from '+@tb_Apsnt_tmp+' m ' +@Cr+
                ' where 1=1 '+@Cr+
                '   and ps_date = m.A_ps_date '+@Cr+
                '   and ps_no collate Chinese_Taiwan_Stroke_CI_AS = m.e_no collate Chinese_Taiwan_Stroke_CI_AS '+@Cr+
                '   and (m.ps_time_s collate Chinese_Taiwan_Stroke_CI_AS <> ps_tm1s collate Chinese_Taiwan_Stroke_CI_AS '+@Cr+
                '    or m.ps_time_e collate Chinese_Taiwan_Stroke_CI_AS <> ps_tm1e collate Chinese_Taiwan_Stroke_CI_AS '+@Cr+
                '    or m.ps_time2_s collate Chinese_Taiwan_Stroke_CI_AS <> ps_tm2s collate Chinese_Taiwan_Stroke_CI_AS) '
  
  Exec @Cnt = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
  if @Cnt = @Errcode Goto End_Proc
  

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Msg = '9.匯入 出勤資料至凌越系統['+@tb_Name+'].'
  set @Cnt = 0
  set @strSQL = --'insert into '+@rDB+@tb_Name+@Cr+
                'insert into '+@wDB+@tb_Name+@Cr+
                '(ps_date, ps_wk, ps_no, ps_tm1s, ps_tm1e, ps_tm2s, ps_tm2e, ps_tm3s, ps_tm3e, ps_tm4s, ps_tm4e, ps_dept)'+@Cr+
-- 2014/06/25 Rickliu 管理部指示出勤紀錄僅取第一筆與最後一筆當作上下班紀錄。
                'select A_ps_date, ps_week, e_no, ps_time_s, ps_time_e, ps_time2_s, ps_time2_e, ps_time3_s, ps_time3_e, ps_time4_s, ps_time4_e, e_dept '+@Cr+
                '  from '+@tb_Apsnt_tmp+' m ' +@Cr+
                ' where not exists'+@Cr+
                '       (select 1'+@Cr+
                '          from '+@wDB+@tb_Name+' d'+@Cr+
                '         where m.e_no collate Chinese_Taiwan_Stroke_CI_AS = d.ps_no collate Chinese_Taiwan_Stroke_CI_AS'+@Cr+
                '           and m.A_ps_date = d.ps_date '+@Cr+
                '       )'

  Exec @Cnt = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
  --2014/01/02 brian增加下班時間

  set @strSQL='UPDATE '+@wDB+@tb_Name++@Cr+
              '   SET ps_tm1s = ps_tm4s	'+@Cr+	
              ' where Rtrim(isnull(ps_tm4s, '''')) <> '''' '+@Cr+
              '   and Rtrim(isnull(ps_tm1s, '''')) <> '''' '+@Cr+
              ' '+@Cr+
              'UPDATE '+@wDB+@tb_Name++@Cr+
              '   SET ps_tm1e = ps_tm4e	'+@Cr+	
              ' where Rtrim(isnull(ps_tm4e, '''')) <> '''' '+@Cr+
              '   and Rtrim(isnull(ps_tm1e, '''')) <> '''' '

  Exec @Cnt = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL

  if @Cnt = @Errcode Goto End_Proc
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
End_Proc:
/*
  -- 發送 MAIL 給相關人等
--20140106 brian增加判斷星期日無刷卡資料不寄送錯誤
--   星期一  1  星期日 7

if @weekday <> 7    
	begin
	*/
	  if @Cnt = @Errcode 
	  begin
		 set @Run_Msg = '['+Convert(varchar(12), @ps_date, 111)+'] '+ @Proce_SubName + '...執行失敗!!!'
		 set @Body_Msg = '錯誤訊息：'+@Cr+@Msg + '...執行失敗!!!'
		 
		 Declare @sCmd Varchar(1000)
		 Declare @sMsg Varchar(1000)
         Declare @weekday int
         Exec uSP_Sys_Write_Log @Proc, @Run_Msg, @Body_Msg, -1, 2
         
         select @Cnt = count(1) 
           from trans_log 
          where process=@Proc
            and recordcount = -1
            --and (sqlcmd like '%sp_send_dbmail%' )
            and Trans_date >= Convert(Varchar(20), GETDATE(), 111)
            and Trans_date <= Convert(Varchar(20), GETDATE()+1, 111)
         
         set @weekday=(SELECT DATEPART(WEEKDAY, GETDATE()-1))
         /**************************************************************************************************
         注意：會做這一段是主要因為一旦星期六日沒有上班時仍會每分鐘發MAIL出去，
               在此檢核 Trans_Log 是否有送錯誤訊息 10 次，若有只要發一封MAIL出去即可。-- Rickliu 2014/04/07
         **************************************************************************************************/
         if (@Cnt = 10) and (@weekday in (6, 7))
            Exec uSP_Sys_Write_Log @Proc, @Run_Msg, @Body_Msg, -1

		 
	  end
End_Exit:
  --Exec uSP_Advanced_Options 1
  Return(@Cnt)
end
GO
