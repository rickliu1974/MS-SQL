USE [DW]
GO
/****** Object:  StoredProcedure [dbo].[zNew_uSP_ETL_pemploy]    Script Date: 07/24/2017 14:44:00 ******/
DROP PROCEDURE [dbo].[zNew_uSP_ETL_pemploy]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE Proc [dbo].[zNew_uSP_ETL_pemploy](@sDB Varchar(10) = 'TA13')
as
begin
  /*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
   Program Name: uSP_ETL_pemploy
   Create Date: 2013/03/01
   Creator: Rickliu
   Updated Date: 2013/09/06 [修改 Trans_Log 訊息內容]
  *-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*/  

  Declare @Proc Varchar(50) = 'uSP_ETL_pemploy'
  Declare @Cnt Int =0
  Declare @RowCount table (cnt int)
  Declare @Err_Code Int = -1
  Declare @Msg Varchar(Max) =''
  Declare @strSQL Varchar(Max)=''
  Declare @CR Varchar(4) = char(13)+char(10)
  Declare @Tb_Name Varchar(20) = 'Fact_pemploy'
  Declare @Result Int = 0

  Begin Try
    IF Not EXISTS (SELECT * FROM sys.databases WHERE name = 'SYNC_'+@sDB)
    begin
       set @Msg = '所選的資料庫不存在 ['+@sDB+']....'
       Raiserror(@Msg, 16, 1)
    end

    If (@sDB = 'AN13' Or @sDB = 'PL13') set @sDB = 'TA13'
    set @Tb_Name = @sDB+'#'+@Tb_Name

    IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].['+@Tb_Name+']') AND type in (N'U'))
    begin
       Set @Msg = '刪除資料表 ['+@Tb_Name+']'
       set @strSQL= 'DROP TABLE [dbo].['+@Tb_Name+']'

       Exec @Result = SP_Exec_SQL @Proc, @Msg, @strSQL
       if @Result <> - 1 set @Result = 0
    end
  
    set @Msg = '建立 ['+@Tb_Name+'] 資料表...'
    set @strSQL = 'select distinct '+@CR+
                  '       Rtrim(M.[E_NO]) as e_no, '+@CR+
                  '       Rtrim(M.[E_NAME]) as e_name, M.[E_SEX], M.[E_MARR], M.[E_BTYPE], '+@CR+
                  '       M.[E_PREV], M.[E_ID], M.[E_BIRTH], M.[E_DEPT], M.[E_DUTY], '+@CR+
                  '       M.[E_RDATE], M.[E_LDATE], '+@CR+
                  '       Rtrim(M.[E_BKNO]) as e_bkno, '+@CR+
                  '       Rtrim(M.[E_CONTACT]) as e_contact, '+@CR+
                  '       Rtrim(M.[E_TEL]) as e_tel, '+@CR+
                  '       [E_ADDR]=Rtrim(Convert(Varchar(255), M.[E_ADDR])), '+@CR+
                  '       [E_BADDR]=Rtrim(Convert(Varchar(255), M.[E_BADDR])), '+@CR+
                  '       [E_CRED]=Rtrim(Convert(Varchar(255), M.[E_CRED])), '+@CR+
                  '       [E_VITA]=Rtrim(Convert(Varchar(255), M.[E_VITA])), '+@CR+
                  '       [E_REM]=Rtrim(Convert(Varchar(255), M.[E_REM])), '+@CR+
                  '       Rtrim(M.[E_FIGNAME]) as e_figname, '+@CR+
                  '       Rtrim(M.[E_SPDAY]) as e_spday, '+@CR+
                  '       Rtrim(M.[E_FLD1]) as e_fld1, '+@CR+
                  '       Rtrim(M.[E_FLD2]) as e_fld2, '+@CR+
                  '       Rtrim(M.[E_FLD3]) as e_fld3, '+@CR+
                  '       Rtrim(M.[E_FLD4]) as e_fld4, '+@CR+
                  '       Rtrim(M.[E_FLD5]) as e_fld5, '+@CR+ 
                  '       Rtrim(M.[E_FLD6]) as e_fld6, '+@CR+
                  '       Rtrim(M.[E_FLD7]) as e_fld7, '+@CR+
                  '       Rtrim(M.[E_FLD8]) as e_fld8, '+@CR+
                  '       Rtrim(M.[E_FLD9]) as e_fld9, '+@CR+
                  '       Rtrim(M.[E_FLD10]) as e_fld10, '+@CR+
                  '       Rtrim(M.[E_FLD11]) as e_fld11, '+@CR+
                  '       Rtrim(M.[E_FLD12]) as e_fld12, '+@CR+
                  '       M.[E_TFG], '+@CR+
                  '       M.[E_INBANK], '+@CR+
                  '       Rtrim(M.[E_MSTNO]) as e_mstno, '+@CR+
                  '       [EM_BIMID], '+@CR+
                  -- 血型
                  '       [Chg_E_BTYPE]= '+@CR+
                  '         Case M.E_BTYPE '+@CR+
                  '           when 1 then ''O'' '+@CR+
                  '           when 2 then ''A'' '+@CR+
                  '           when 3 then ''B'' '+@CR+
                  '           when 4 then ''AB'' '+@CR+
                  '           else ''NA'' '+@CR+
                  '         end, '+@CR+
                  -- 性別
                  '       Chg_e_sex= '+@CR+
                  '         case M.e_sex '+@CR+
                  '           when ''1'' then ''男'' '+@CR+ 
                  '           when ''2'' then ''女'' '+@CR+
                  '           else '''' '+@CR+ 
                  '         end, '+@CR+
                  -- 星座
                  '       Chg_E_BIRTH= '+@CR+
                  '         Case '+@CR+
                  '           when (DatePart(MM, M.E_BIRTH) = 1 and  DatePart(DD, M.E_BIRTH) >= 21) Or (DatePart(MM, M.E_BIRTH) = 2 and  DatePart(DD, M.E_BIRTH) <= 19) then ''水瓶座'' '+@CR+
                  '           when (DatePart(MM, M.E_BIRTH) = 2 and  DatePart(DD, M.E_BIRTH) >= 20) Or (DatePart(MM, M.E_BIRTH) = 3 and  DatePart(DD, M.E_BIRTH) <= 20) then ''雙魚座'' '+@CR+
                  '           when (DatePart(MM, M.E_BIRTH) = 3 and  DatePart(DD, M.E_BIRTH) >= 21) Or (DatePart(MM, M.E_BIRTH) = 4 and  DatePart(DD, M.E_BIRTH) <= 20) then ''牡羊座'' '+@CR+
                  '           when (DatePart(MM, M.E_BIRTH) = 4 and  DatePart(DD, M.E_BIRTH) >= 21) Or (DatePart(MM, M.E_BIRTH) = 5 and  DatePart(DD, M.E_BIRTH) <= 21) then ''金牛座'' '+@CR+
                  '           when (DatePart(MM, M.E_BIRTH) = 5 and  DatePart(DD, M.E_BIRTH) >= 22) Or (DatePart(MM, M.E_BIRTH) = 6 and  DatePart(DD, M.E_BIRTH) <= 21) then ''雙子座'' '+@CR+
                  '           when (DatePart(MM, M.E_BIRTH) = 6 and  DatePart(DD, M.E_BIRTH) >= 22) Or (DatePart(MM, M.E_BIRTH) = 7 and  DatePart(DD, M.E_BIRTH) <= 23) then ''巨蟹座'' '+@CR+
                  '           when (DatePart(MM, M.E_BIRTH) = 7 and  DatePart(DD, M.E_BIRTH) >= 24) Or (DatePart(MM, M.E_BIRTH) = 8 and  DatePart(DD, M.E_BIRTH) <= 23) then ''獅子座'' '+@CR+
                  '           when (DatePart(MM, M.E_BIRTH) = 8 and  DatePart(DD, M.E_BIRTH) >= 24) Or (DatePart(MM, M.E_BIRTH) = 9 and  DatePart(DD, M.E_BIRTH) <= 23) then ''處女座'' '+@CR+
                  '           when (DatePart(MM, M.E_BIRTH) = 9 and  DatePart(DD, M.E_BIRTH) >= 24) Or (DatePart(MM, M.E_BIRTH) = 10 and DatePart(DD, M.E_BIRTH) <= 23) then ''天秤座'' '+@CR+
                  '           when (DatePart(MM, M.E_BIRTH) = 10 and DatePart(DD, M.E_BIRTH) >= 24) Or (DatePart(MM, M.E_BIRTH) = 11 and DatePart(DD, M.E_BIRTH) <= 22) then ''天蝎座'' '+@CR+
                  '           when (DatePart(MM, M.E_BIRTH) = 11 and DatePart(DD, M.E_BIRTH) >= 23) Or (DatePart(MM, M.E_BIRTH) = 12 and DatePart(DD, M.E_BIRTH) <= 22) then ''射手座'' '+@CR+
                  '           when (DatePart(MM, M.E_BIRTH) = 12 and DatePart(DD, M.E_BIRTH) >= 23) Or (DatePart(MM, M.E_BIRTH) = 1 and  DatePart(DD, M.E_BIRTH) <= 20) then ''魔羯座'' '+@CR+
                  '           else ''NA'' '+@CR+
                  '         end, '+@CR+
                  -- 到職年份
                  '       Chg_e_rdate_YY=DatePart(YY, Isnull(case when Convert(Varchar(4), M.e_rdate, 112) = ''1900'' then null else M.e_rdate end, getdate())), '+@CR+
                  -- 到職月份
                  '       Chg_e_rdate_MM=DatePart(MM, Isnull(case when Convert(Varchar(4), M.e_rdate, 112) = ''1900'' then null else M.e_rdate end, getdate())), '+@CR+
                  -- 離職年份
                  '       Chg_e_ldate_YY=DatePart(YY, M.e_ldate), '+@CR+
                  -- 離職月份
                  '       Chg_e_ldate_MM=DatePart(MM, M.e_ldate), '+@CR+
                  -- 到職天數
                  '       Chg_emp_day= '+@CR+
                  '         case '+@CR+
                  '           when (Convert(Varchar(4), M.e_ldate, 112) = ''1900'') then datediff(dd, M.e_rdate, getdate()) '+@CR+
                  '           else datediff(dd, M.e_rdate, M.e_ldate) '+@CR+
                  '         end, '+@CR+
                  -- 年資
                  '       Chg_emp_sumYM= '+@CR+
                  '         case '+@CR+
                  '           when (Convert(Varchar(4), M.e_ldate, 112) = ''1900'') '+@CR+
                  '           then (datediff(dd, M.e_rdate, getdate()) / 365)+ '+@CR+
                  '                 FLOOR(((datediff(dd, M.e_rdate, getdate()) + '+@CR+
                  '                       ((datediff(dd, M.e_rdate, getdate()) * 1.0) / 365) - '+@CR+
                  '                        (datediff(dd, M.e_rdate, getdate()) / 365) * 365) / 30)) * 0.1 '+@CR+
                  '           else '+@CR+
                  '                (datediff(dd, M.e_rdate, M.e_ldate) / 365)+ '+@CR+
                  '                 FLOOR(((datediff(dd, M.e_rdate, M.e_ldate) + '+@CR+
                  '                       ((datediff(dd, M.e_rdate, M.e_ldate) * 1.0) / 365) - '+@CR+
                  '                        (datediff(dd, M.e_rdate, M.e_ldate) / 365) * 365) / 30)) * 0.1 '+@CR+
                  '         end, '+@CR+
                  -- 離職否(Y:離職, N:在職)
                  '       Chg_leave= '+@CR+
                  '         case when (Convert(Varchar(4), M.e_ldate, 112) = ''1900'') then ''N'' else ''Y'' end, '+@CR+
                  -- 部門
                  '       Chg_dp_name=Rtrim(D1.dp_name), '+@CR+
                  '       Chg_dp_ename=Rtrim(D1.dp_ename), '+@CR+
                  '       Chg_dp_whno=Rtrim(D1.dp_whno), '+@CR+
                  '       Chg_dp_rem=Rtrim(Convert(Varchar(255), D1.dp_rem)), '+@CR+
                  -- 組別
                  '       Chg_mst_name=RTrim(D2.mst_name), '+@CR+
                  -- 本年度業績目標(是否為業務可以從此判別)
                  '       [Chg_Year_Tot_SaleAmt]= '+@CR+
                  '         (select Year_Tot_SaleAmt=Isnull(sum(bo_sale), 0) '+@CR+
                  '            from v_sales_base '+@CR+
                  '           where bo_class = ''1'' '+@CR+
                  '             and bo_yr=DatePart(YY, GetDate()) '+@CR+
                  '             and bo_no=M.e_no '+@CR+
                  '         ), '+@CR+
                  '       D3.Chg_Master_Dept, D3.Chg_Duty_Level, D3.Chg_Dept_Level, D3.stock_amt_Level, '+@CR+
                  '       D3.Result_e_no, D3.Result_dept_level, D3.Result_Duty_Level, D3.Result_e_mstno, '+@CR+
                  '       update_datetime = getdate() '+@CR+
                  '       into '+@Tb_Name+' '+@CR+
                  '  from SYNC_'+@sDB+'.dbo.pemploy M '+@CR+
                  '       left join SYNC_'+@sDB+'.dbo.pdept D1 on M.e_dept=D1.dp_no '+@CR+
                  '       left join SYNC_'+@sDB+'.dbo.masterm D2 on M.e_mstno=D2.mst_no '+@CR+
                  '       left join uV_'+@sDB+'#Emp_Level D3 on M.e_no=D3.e_no '

    Exec @Result = SP_Exec_SQL @Proc, @Msg, @strSQL
    if @Result <> -1 Set @Result = 0
    
    
    set @Msg = '計算 ['+@Tb_Name+'] 資料表筆數...'
    set @strSQL = '--'+@Msg+@CR+'select count(1) as cnt from '+@Tb_Name
    print @strSQL
    insert into @RowCount Exec(@strSQL)
    Select @Cnt = Cnt from @RowCount
    
    if @Cnt = 0 Set @Result = -1
          
    Exec SP_Write_Log @Proc, @Msg, @strSQL, @Result
  end try
  begin catch
    set @Result = @Err_Code
    set @Msg = @Msg+'(錯誤訊息:'+ERROR_MESSAGE()+')'

    Exec SP_Write_Log @Proc, @Msg, @strSQL, @Result
  end catch

  Return(@Result)
end
GO
