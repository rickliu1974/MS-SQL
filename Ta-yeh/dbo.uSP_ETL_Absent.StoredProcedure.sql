USE [DW]
GO
/****** Object:  StoredProcedure [dbo].[uSP_ETL_Absent]    Script Date: 07/24/2017 14:43:59 ******/
DROP PROCEDURE [dbo].[uSP_ETL_Absent]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE Proc [dbo].[uSP_ETL_Absent]
  @Kind Int = 0
as
begin
  /*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
   Rickliu 2017/04/24
   Kind: 0 �w�]�H�t�Τ���i��X�Ը�Ƥ���
         1 �i��Ҧ��X�Ը�Ƥ���
  *--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*/
  Declare @Proc Varchar(50) = 'uSP_ETL_Absent'
  --Declare @Ori_TA13 Varchar(20) = 'Ori_TA13#' -- 2017/03/17 �ק��ϥ� Ori_TA13 �אּ�H SYNC_TA13 �q�\��Ʈw���ѦҸ��
  --Declare @rDB Varchar(30) = 'TYEIPDBS2.lytdbta13.dbo.'
  Declare @rDB Varchar(30) = 'SYNC_TA13.dbo.' 
  Declare @wDB Varchar(30) = 'TYEIPDBS2.lytdbta13.dbo.' -- 2017/03/17 �ק��ϥ� Ori_TA13 �אּ�H SYNC_TA13 �q�\��Ʈw���ѦҸ��
  
  --set @rDB = @wDB -- Rickliu 2017/03/24 �b�q�\�A���٨S�n�ɡA�Ȯɨϥ� @wDB ��Ʈw�����
    
  Declare @TB_Name1 Varchar(50) = 'TB_Absent1'
  Declare @TB_Name2 Varchar(50)= 'TB_Absent2'
  Declare @TB1_Exists Bit = 1 -- ��ƪ�O�_�s�b 0:N, 1:Y
  Declare @TB2_Exists Bit = 1 -- ��ƪ�O�_�s�b 0:N, 1:Y
  Declare @Sender Varchar(100) = 'it@ta-yeh.com.tw'
  Declare @Body_Msg NVarchar(Max) = '', @Run_Proc_Msg NVarchar(Max) = '', @Run_Msg NVarchar(1000) = ''

  Declare @Cnt Int =0
  Declare @Msg Varchar(4000) =''
  Declare @Errcode int = -1
  Declare @Result int = 0

  Declare @strSQL Varchar(Max)=''
  Declare @strSQL_TB1_Insert Varchar(Max)= 'Insert Into '+@TB_Name1
  Declare @strSQL_TB2_Insert Varchar(Max)= 'Insert Into '+@TB_Name2


  Declare @Table_Name Varchar(100)=''
  Declare @CR Varchar(4) = char(13)+char(10)

  -- 2017/04/21 Rickliu �p�G�O Null �h�N��n����������ơA�Y�����w������h���ӤѪ���ơA�_�h�@�v�����t�Τ骺���
  Declare @ps_date Varchar(10) = Case when @Kind = 0 then Convert(Varchar(10), getdate(), 111) else '' end
  
  -- case when (@ps_set_date is not null) and (isdate(@ps_set_date) = 0) then Convert(Varchar(10), getdate(), 111) else '' end

  Begin Try
     IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].['+@TB_Name1+']') AND type in (N'U'))
     begin
        Set @Msg = '�M��'+@ps_date+'�X�Ը�� ['+@TB_Name1+']'
        if @ps_date = ''
           Set @strSQL = 'Drop Table '+@TB_Name1
        else
           Set @strSQL = 'Delete '+@TB_Name1+@Cr+' where abs_day = '''+@ps_date+''' '
   
        Exec @Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
        if @Result <>  -1 Set @Result = 0
     end
     else
        set @TB1_Exists = 0
 
     IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].['+@TB_Name2+']') AND type in (N'U'))
     begin
        Set @Msg = '�M��'+@ps_date+'�X�Բ��`���['+@TB_Name2+']'
        if @ps_date = ''
           Set @strSQL = 'Drop Table '+@TB_Name2
        else
           Set @strSQL = 'Delete '+@TB_Name2+@Cr+' where abs_day = '''+@ps_date+''' '
   
        Exec @Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
        if @Result <>  -1 Set @Result = 0
     end
     else
        set @TB2_Exists = 0
          
     Set @strSQL = ';With CTE_Q1 as ( '+@CR+
                   -- 2017/04/21 Rickliu �ѩ�q�\�^�g�ɤ������G�A�^�g�^�Ӫ��P���W�ٷ|�ܦ��w�ǿ��áA�]����ħ���U�~��覡�C
                   '  select distinct ps_date as Days, substring(datename(weekday, ps_date), 3, 1) as wk '+@CR+
                   '    from '+@rDB+'apsnt m '+@CR+
                   '   where 1=1  '+@CR+
                   Case When @ps_date = '' then '' else '     and ps_date = '''+@ps_date+''' ' end+' '+@CR+
                   --*--*--*--*--*
                   '), CTE_Q2 as (  '+@CR+
                   -- 2016/09/08 Rickliu 
                   --  �o�̥D�n�O���ͨC�H�C�Ѫ��ťեX�Ԫ�
                   '   select e_no, AccountID, e_name, dp_no, DP_NAME, e_rdate, '+@CR+
                   --  2016/09/08 Rickliu �׭q�u���ӤH�b¾��ƾ�A�Y�ӭ��S����¾�A�h�N��¾��]�w���̤j����C
                   --  2017/04/21 Rickliu �Y�P�@���s�ȥu�O�մ������A���ɭ��s���ܡAEIP �b���P EIP ���A�|�Q���ܡA�]���Y�ӱb�����ΡA�]�����ӱb���w�g��¾�C
                   '          case when (Status = ''Y'' and e_ldate=''1900-01-01'') then ''9999-12-31'' else e_ldate end as e_ldate, '+@CR+
                   '          Chg_Master_Dept, Chg_Duty_Level, Chg_Dept_Level, Remark '+@CR+
                   '     from uV_TA13#EMP_LEVEL '+@CR+
                   '    where E_NO <> ''LY'' '+@CR+
                   '      and substring(E_NO, 1, 3) not in (''TAZ'', ''CZZ'') '+@CR+
                   -- and status =''Y'' '+@CR+
                   -- 2016/09/09 Rickliu ��P�O�H����¾��O�_���b�X�Ԥ�A�@���ӭ��O�_���b¾���̾ڡC
                   --*--*--*--*--*
                   '), CTE_Q3 as ( '+@CR+
                   '  select * '+@CR+
                   '    from CTE_Q1 m '+@CR+
                   '         Left Join CTE_Q2 d on days between e_rdate and e_ldate '+@CR+
                   --*--*--*--*--*
                   --2013/12/11 brianhsu �W�[ps_tm2s,ps_tm2e �[�Z�ɶ�
                   --2014/07/09 Rickliu �W�[ ps_tm3s, ps_tm3e ��l��d
                   '), CTE_Q4 as ('+@CR+
                   '  select ps_date ,ps_wk ,ps_no, '+@CR+
                   '         ps_tm1s=RTrim(Replace(ps_tm1s, ''  :  '', '''')), ps_tm1e=RTrim(Replace(ps_tm1e, ''  :  '', '''')), '+@CR+
                   '         ps_tm2s, ps_tm2e, '+@CR+
                   '         ps_caltm= Convert(Varchar(8), '+@CR+
                   '         Convert(DateTime, Case when RTrim(Replace(ps_tm1e, ''  :  '', '''')) = '''' then RTrim(Replace(ps_tm1s, ''  :  '', '''')) else RTrim(Replace(ps_tm1e, ''  :  '', '''')) end)- '+@CR+
                   '         Convert(DateTime, Case when RTrim(Replace(ps_tm1s, ''  :  '', '''')) = '''' then RTrim(Replace(ps_tm1e, ''  :  '', '''')) else RTrim(Replace(ps_tm1s, ''  :  '', '''')) end), 108), '+@CR+
                   '		 ps_tm3s, ps_tm3e '+@CR+
                   '    from '+@rDB+'apsnt '+@CR+
                   '   where 1=1 '+@CR+
                   Case When @ps_date = '' then '' else '     and ps_date = '''+@ps_date+''' ' end+' '+@CR+

                   -- �u�O�d���Ӥ�
                   --'                 where ps_date >= Convert(Varchar(10), DateAdd(mm, -6, '''+@ps_date+'''), 110) '+@CR+
                   --*--*--*--*--*
                   '), CTE_Q5 as ('+@CR+
                   '  select m.days, m.wk, m.dp_no, m.AccountID, m.DP_NAME, m.e_no, m.e_name, '+@CR+
                   '         d.ps_tm1s, d.ps_tm1e, d.ps_tm2s, d.ps_tm2e, d.ps_caltm, d.ps_tm3s, d.ps_tm3e, '+@CR+
                   '         m.e_ldate, Chg_Duty_Level, Chg_Dept_Level, Remark '+@CR+
                   '    from CTE_Q3 m '+@CR+
                   '         left join CTE_Q4 d on d.ps_date = m.days and d.ps_no = m.e_no '+@CR+
                   --*--*--*--*--*
                   -- ��V����O�� 
                   '), CTE_Q6 as ('+@CR+
                   '  select m.hl_no Collate Chinese_Taiwan_Stroke_CI_AS as e_no, '+@CR+ --D2.e_Name, 
                   '         Case when d.sn = 0 then Substring(m.hl_memo Collate Chinese_Taiwan_Stroke_CI_AS, 1, 100) else d.Name Collate Chinese_Taiwan_Stroke_CI_AS end as abs_name, '+@CR+
                   '         m.hl_date , m.hl_sdate, m.hl_edate, m.hl_memo Collate Chinese_Taiwan_Stroke_CI_AS as hl_memo '+@CR+
                   '    from '+@rDB+'ahold m '+@CR+
                   '         inner join '+@wDB+'lyholtype d on m.hl_hfg = d.sn '+@CR+
                   --*--*--*--*--*
                   --Rickliu 2014/11/14 ���s�׭q�u���w�gñ�ֹL�����
                   '), CTE_Q7 as ('+@CR+
                   '  select sFormSerialID, Max(sSignDateTime) as sSignDateTime '+@CR+
                   '    from WebFlow3.dbo.afu_sign_TYG_AD_03_B01 '+@CR+
                   '   group by sFormSerialID '+@CR+
                   --*--*--*--*--*
                   --Rickliu 2014/11/14 ���s�׭q�u���w�gñ�ֹL�����
                   '), CTE_Q8 as ('+@CR+
                   '  select M.* '+@CR+
                   '    from WebFlow3.dbo.afu_sign_TYG_AD_03_B01 m '+@CR+
                   '         inner join CTE_Q7 d on M.sFormSerialID = D.sFormSerialID and M.sSignDateTime = D.sSignDateTime '+@CR+
                   '   where 1=1 '+@CR+
                   '     and M.sResult= ''1'' '+@CR+
                   --*--*--*--*--*
                   -- EIP ���X�O��
                   '), CTE_Q9 as ('+@CR+
                   '  select Distinct '+@CR+
                   '         D1.e_no, '+@CR+
                   '         FormName=Replace(M.FormName Collate Chinese_Taiwan_Stroke_CI_AS, ''�ӽг�'', '''') Collate Chinese_Taiwan_Stroke_CI_AS, '+@CR+
                   '         ApplyDateTime=Convert(DateTime, Convert(Varchar(10), M.applyDateTime, 111)), '+@CR+
                   '         D.ExpectoutDate, D.ExpectBackDate, '+@CR+
                   '         ReMark=RTrim(Isnull(D.OutLocation Collate Chinese_Taiwan_Stroke_CI_AS, '''')+''-''+ '+@CR+
                   '         Isnull(D.OutTarget Collate Chinese_Taiwan_Stroke_CI_AS, '''') +''-''+ '+@CR+
                   '         Isnull(D.OutReason Collate Chinese_Taiwan_Stroke_CI_AS, '''')) '+@CR+
                   '   from webflow3.dbo.afs_flow M '+@CR+
                   '        inner join webflow3.dbo.afu_form_tyg_ad_03_b01 D on M.SerialID=D.SerialID '+@CR+
                   '        inner join uV_TA13#EMP_LEVEL D1 on m.applyID=D1.AccountID '+@CR+
                   --2013/12/12 brianhsu �W�[FLOW ���u��ñ�֪��A 1 �wñ�� 2 ��^
                   --       inner join  [VMS\VMSOP].WebFlow3.dbo.afu_sign_TYG_AD_03_B01 C 
                   --          on D.[SerialID]=C.[sFormSerialID]
                   '        inner join CTE_Q8 D2 on D.[SerialID]=D2.[sFormSerialID] '+@CR+
                   --*--*--*--*--*
                   '), CTE_Q10 as ('+@CR+
                   '  select * from CTE_Q6 '+@CR+
                   '  union '+@CR+
                   '  select * from CTE_Q9 '+@CR+
                   ') '+@CR+
                   ' '+@CR+
                   --*--*--*--*--*
                   Case When (@ps_date = '' Or @TB1_Exists = 0) then '' else @strSQL_TB1_Insert end+' '+@CR+
                   --*--*--*--*--*
                   'select abs_day=m.days, abs_month=month(m.days), abs_week=m.wk, '+@CR+
                   '       abs_dept=Rtrim(m.DP_no), abs_deptname=Rtrim(m.DP_NAME), abs_emp_no=Rtrim(m.e_no), '+@CR+
                   '       abs_emp_name=Rtrim(m.e_name), Chg_Duty_Level, Chg_Dept_Level, '+@CR+
                   '       abs_stime= Case when (m.ps_tm1s <> m.ps_tm3s) or (m.ps_tm3s <> m.ps_tm3s) then m.ps_tm3s else m.ps_tm1s end, '+@CR+
                   '       abs_etime=m.ps_tm1e, m.ps_tm2s, m.ps_tm2e, abs_seTime=m.ps_caltm, '+@CR+
                   '       abs_seTime_sec=isnull(DatePart(hour, cast(m.ps_caltm as datetime))*60*60 + DatePart(mi, cast(m.ps_caltm as datetime))*60 + DatePart(s, cast(m.ps_caltm as datetime)), 0), '+@CR+
                   '       abs_ldate=m.e_ldate, m.ps_tm3s, m.ps_tm3e, '+@CR+
                   '       hold_Name= '+@CR+
                   '         Case '+@CR+
                   --2015/12/30 NanLiao �t�X���q�F���A2016/01/01�}�l�W�Z�ɶ��אּ08:30
                   --2017/01/03 NanLiao �t�X���q�F���A2017/01/01�}�l�W�Z�ɶ��אּ08:45
				   '            when d.abs_name IS NULL AND m.ps_tm1s >= ''08:46'' and year(m.days) <= 2015 then ''���'' '+@CR+
                   '            when d.abs_name IS NULL AND m.ps_tm1s >= ''08:31'' and year(m.days) > 2015 and year(m.days) <= 2016 then ''���'' '+@CR+
                   '            when d.abs_name IS NULL AND m.ps_tm1s >= ''08:46'' and year(m.days) > 2016 then ''���'' '+@CR+
                   '            when d.abs_name <> '''' then Rtrim(d.abs_name) '+@CR+
                   '         end, '+@CR+
                   '       hold_Memo= '+@CR+
                   '         Case '+@CR+
                   '           when Rtrim(isnull(d.abs_name, '''')) = '''' and Rtrim(isnull(m.ps_tm1s, '''')+isnull(m.ps_tm1e, '''')) =''''  then ''�L��d'' '+@CR+
                   '           when rtrim(isnull(d.hl_memo, '''')) = '''' then Rtrim(m.ReMark) '+@CR+
                   '           else substring(hl_memo, 1, 100) '+@CR+
                   '         end + '+@CR+
                   '         Case when (m.ps_tm1s <> m.ps_tm3s) or (m.ps_tm3s <> m.ps_tm3s) then ''(�X�Ըɥ�)'' else '''' end, '+@CR+
                   --2014/07/09 Rickliu �W�[�X�Ըɥ����O
                   '       hold_chg=Case when (m.ps_tm1s <> m.ps_tm3s) or (m.ps_tm3s <> m.ps_tm3s) then ''Y'' else '''' end, '+@CR+
                   '       hold_sdate=d.hl_sdate, hold_edate=d.hl_edate, '+@CR+
                   '       hold_seDay= '+@CR+
                   '         Case '+@CR+
                   '           when Convert(Varchar(5), (d.hl_edate - d.hl_sdate)-1, 108) >=''08:45'' and Convert(Varchar(2), Datepart(dd, d.hl_edate - d.hl_sdate)-1) =0 '+@CR+
                   '           then Convert(Varchar(2), Datepart(dd, d.hl_edate - d.hl_sdate)) '+@CR+
                   '           when Convert(Varchar(5), (d.hl_edate - d.hl_sdate)-1, 108) <> '''' and Convert(Varchar(2), Datepart(dd, d.hl_edate - d.hl_sdate)-1) =0 '+@CR+
                   '           then Convert(Varchar(2), Datepart(dd, d.hl_edate - d.hl_sdate)-1) '+@CR+
                   '           else Convert(Varchar(2), Datepart(dd, d.hl_edate - d.hl_sdate)) '+@CR+
                   '         end, '+@CR+
                   '       hold_seTime=Case when Convert(Varchar(8), (d.hl_edate - d.hl_sdate)-1, 108) =''08:45'' then '''' else Convert(Varchar(8), (d.hl_edate - d.hl_sdate)-1, 108) end, '+@CR+
                   '       hold_seTime_sec= '+@CR+
                   '         Case '+@CR+
                   '           when d.hl_edate is null or d.hl_sdate is null then 0 '+@CR+
                   '           when Convert(Varchar(8), (d.hl_edate - d.hl_sdate)-1, 108) =''08:45'' then 0 '+@CR+
                   '           else DatePart(hour, Convert(Varchar(8), (d.hl_edate - d.hl_sdate)-1, 108))*60*60 + '+@CR+
                   '                DatePart(mi, Convert(Varchar(8), (d.hl_edate - d.hl_sdate)-1, 108))*60 + '+@CR+
                   '                DatePart(s, Convert(Varchar(8), (d.hl_edate - d.hl_sdate)-1, 108)) '+@CR+
                   '         end, '+@CR+
                   '       Rtrim(m.AccountID) as AccountID, '+@CR+
                   '       UPdate_DateTime = Getdate() '+@CR+
                   --2013/12/11 brianhsu �W�[ps_tm2s,ps_tm2e �[�Z�ɶ�
                   Case When (@ps_date = '' Or @TB1_Exists = 0) then Replace(@strSQL_TB1_Insert, 'Insert', ' ') else '' end+' '+@CR+
                   '  From CTE_Q5 m '+@CR+
                   '       Left Join CTE_Q10 d '+@CR+
                   '         on m.e_no = d.e_no '+@CR+
                   '        and CONVERT(varchar(10), m.days, 111) between CONVERT(varchar(10), d.hl_sdate, 111) '+@CR+
                   '        and CONVERT(varchar(10), d.hl_edate, 111) '+@CR+
                   ' where 1=1 '+@CR+
                   --and m.e_no not in ('T08103', 'T88062', 'T88061', 'T90021')
                   '   and m.e_no Like ''T[0123456789]%'' '+@CR+
                   '   and isnull(Rtrim(m.AccountID), '''') not like ''vp%'' '+@CR+
                   '   and m.e_no <> ''T08103'' '
     Exec @Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
     if @Result <>  -1 Set @Result = 0
   
     Set @strSQL = Case When (@ps_date = '' Or @TB2_Exists = 0) then '' else @strSQL_TB2_Insert end+' '+@CR+
                   'Select Distinct '+@CR+
                   '       abs_day, '+@CR+
                   '       abs_week, '+@CR+
                   '       abs_month, '+@CR+
                   '       abs_dept, '+@CR+
                   '       abs_deptname, '+@CR+
                   '       AccountID, '+@CR+
                   '       abs_emp_no, '+@CR+
                   '       abs_emp_name, '+@CR+
                   '       abs_stime, '+@CR+
                   '       abs_etime, '+@CR+
                   '       abs_seTime, '+@CR+
                   '       abs_seTime_sec, '+@CR+
                   '       ps_tm2s, '+@CR+
                   '       ps_tm2e, '+@CR+
                   '       ps_tm3s, '+@CR+
                   '       ps_tm3e, '+@CR+
                   '       hold_name, '+@CR+
                   '       hold_memo, '+@CR+
                   '       hold_chg, '+@CR+
                   '       hold_sdate, '+@CR+
                   '       hold_edate, '+@CR+
                   '       Convert(Varchar(8), hold_sdate, 108) as hold_stime, '+@CR+
                   '       Convert(Varchar(8), hold_edate, 108) as hold_etime, '+@CR+
                   '       case '+@CR+
                   '         when hold_seDay = 0 then hold_seTime '+@CR+
                   '         else Substring(Convert(Varchar(3), hold_seDay+100), 2, 2) + ''��'' '+@CR+
                   '       end as hold_seTime, '+@CR+
                   '       hold_seTime_sec, '+@CR+
                   '       case '+@CR+
                   '         when hold_seDay = 0 then ''*'' '+@CR+
                   '         else '''' '+@CR+
                   '       end as hold_Flag, '+@CR+
                   '       abs_ldate, '+@CR+
                   '       Chg_Duty_Level, '+@CR+
                   '       Chg_Dept_Level, '+@CR+
                   '       UPdate_DateTime = Getdate() '+@CR+
                   Case When (@ps_date = '' Or @TB2_Exists = 0) then Replace(@strSQL_TB2_Insert, 'Insert', ' ') else '' end +' '+@CR+
                   '  From '+@TB_Name1+@CR+
                   ' where 1=1  '+@CR+
                   '   and (hold_name <> '''' or hold_Memo <> '''') '+@CR+
                   Case When (@ps_date = '' Or @TB2_Exists = 0) then '' else '   and abs_day = '''+@ps_date+''' ' end+' '+@CR+
                   ' Order By hold_etime, hold_seTime, abs_emp_no, hold_name, hold_memo '
   
     Exec @Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
     if @Result <> -1 Set @Result = 0

     Exec uSP_Sys_Write_Log @Proc, @Run_Msg, '', @Body_Msg, @Result
  end Try
  begin Catch
     Set @Run_Msg = '���� '+@Proc+'...�o���Y�����~!!'
     Set @Body_Msg = @Proc+' ���~�T���G'+@CR+ERROR_MESSAGE()
     Set @Result = -1
     -- �o�e MAIL �������H��
     Exec uSP_Sys_Write_Log @Proc, @Run_Msg, '', @Body_Msg, @Result
  end Catch
  Return(@Result)
end
GO
