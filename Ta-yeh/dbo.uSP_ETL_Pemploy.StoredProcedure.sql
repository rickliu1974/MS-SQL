USE [DW]
GO
/****** Object:  StoredProcedure [dbo].[uSP_ETL_Pemploy]    Script Date: 08/18/2017 17:43:39 ******/
DROP PROCEDURE [dbo].[uSP_ETL_Pemploy]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE Proc [dbo].[uSP_ETL_Pemploy]
as
begin
  /*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
   Program Name: uSP_TA13#ETL_pemploy
   Create Date: 2013/03/01
   Creator: Rickliu
   Updated Date: 2013/09/06 [修改 Trans_Log 訊息內容]
  *-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*/  

  Declare @Proc Varchar(50) = 'uSP_ETL_pemploy'
  Declare @Cnt Int =0
  Declare @Cnt_Ori Int =0
  Declare @Err_Code Int = -1
  Declare @Msg Varchar(Max) =''
  Declare @strSQL Varchar(Max)=''
  Declare @CR Varchar(4) = char(13)+char(10)
  Declare @Result Int = 0

  IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[Fact_pemploy]') AND type in (N'U'))
  begin
     Set @Msg = '刪除資料表 [Fact_pemploy]'
     set @strSQL= 'DROP TABLE [dbo].[Fact_pemploy]'

     Exec @Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
     if @Result <> - 1 set @Result = 0
  end
  
  begin try
    set @Msg = 'ETL pemploy to [Fact_pemploy]...'
    Exec uSP_Sys_Write_Log @Proc, @Msg, @Msg, 0

    select distinct 
           Rtrim(M.[E_NO]) as e_no, 
           Rtrim(M.[E_NAME]) as e_name, M.[E_SEX], M.[E_MARR], M.[E_BTYPE],
           M.[E_PREV], M.[E_ID], M.[E_BIRTH], M.[E_DEPT], M.[E_DUTY],
           M.[E_RDATE], M.[E_LDATE], 
           Rtrim(M.[E_BKNO]) as e_bkno, 
           Rtrim(M.[E_CONTACT]) as e_contact, 
           Rtrim(M.[E_TEL]) as e_tel,
           [E_ADDR]=Rtrim(Convert(Varchar(255), M.[E_ADDR])), 
           [E_BADDR]=Rtrim(Convert(Varchar(255), M.[E_BADDR])), 
           [E_CRED]=Rtrim(Convert(Varchar(255), M.[E_CRED])), 
           [E_VITA]=Rtrim(Convert(Varchar(255), M.[E_VITA])), 
           [E_REM]=Rtrim(Convert(Varchar(255), M.[E_REM])),
           Rtrim(M.[E_FIGNAME]) as e_figname, 
           Rtrim(M.[E_SPDAY]) as e_spday, 
           Rtrim(M.[E_FLD1]) as e_fld1, 
           Rtrim(M.[E_FLD2]) as e_fld2, 
           Rtrim(M.[E_FLD3]) as e_fld3,
           Rtrim(M.[E_FLD4]) as e_fld4, 
           Rtrim(M.[E_FLD5]) as e_fld5, 
           Rtrim(M.[E_FLD6]) as e_fld6, 
           Rtrim(M.[E_FLD7]) as e_fld7, 
           Rtrim(M.[E_FLD8]) as e_fld8,
           Rtrim(M.[E_FLD9]) as e_fld9, 
           Rtrim(M.[E_FLD10]) as e_fld10, 
           Rtrim(M.[E_FLD11]) as e_fld11, 
           Rtrim(M.[E_FLD12]) as e_fld12, 
           M.[E_TFG],
           M.[E_INBANK], 
           Rtrim(M.[E_MSTNO]) as e_mstno, 
           [EM_BIMID],
           -- 血型
           [Chg_E_BTYPE]=
             Case M.E_BTYPE
               when 1 then 'O'
               when 2 then 'A'
               when 3 then 'B'
               when 4 then 'AB'
               else 'NA'
             end,
           -- 性別
           Chg_e_sex=
             case M.e_sex
               when '1' then '男' 
               when '2' then '女' 
               else '' 
             end, 
           -- 星座
           Chg_E_BIRTH=
             Case
               when (DatePart(MM, M.E_BIRTH) = 1 and  DatePart(DD, M.E_BIRTH) >= 21) Or (DatePart(MM, M.E_BIRTH) = 2 and  DatePart(DD, M.E_BIRTH) <= 19) then '水瓶座'
               when (DatePart(MM, M.E_BIRTH) = 2 and  DatePart(DD, M.E_BIRTH) >= 20) Or (DatePart(MM, M.E_BIRTH) = 3 and  DatePart(DD, M.E_BIRTH) <= 20) then '雙魚座'
               when (DatePart(MM, M.E_BIRTH) = 3 and  DatePart(DD, M.E_BIRTH) >= 21) Or (DatePart(MM, M.E_BIRTH) = 4 and  DatePart(DD, M.E_BIRTH) <= 20) then '牡羊座'
               when (DatePart(MM, M.E_BIRTH) = 4 and  DatePart(DD, M.E_BIRTH) >= 21) Or (DatePart(MM, M.E_BIRTH) = 5 and  DatePart(DD, M.E_BIRTH) <= 21) then '金牛座'
               when (DatePart(MM, M.E_BIRTH) = 5 and  DatePart(DD, M.E_BIRTH) >= 22) Or (DatePart(MM, M.E_BIRTH) = 6 and  DatePart(DD, M.E_BIRTH) <= 21) then '雙子座'
               when (DatePart(MM, M.E_BIRTH) = 6 and  DatePart(DD, M.E_BIRTH) >= 22) Or (DatePart(MM, M.E_BIRTH) = 7 and  DatePart(DD, M.E_BIRTH) <= 23) then '巨蟹座'
               when (DatePart(MM, M.E_BIRTH) = 7 and  DatePart(DD, M.E_BIRTH) >= 24) Or (DatePart(MM, M.E_BIRTH) = 8 and  DatePart(DD, M.E_BIRTH) <= 23) then '獅子座'
               when (DatePart(MM, M.E_BIRTH) = 8 and  DatePart(DD, M.E_BIRTH) >= 24) Or (DatePart(MM, M.E_BIRTH) = 9 and  DatePart(DD, M.E_BIRTH) <= 23) then '處女座'
               when (DatePart(MM, M.E_BIRTH) = 9 and  DatePart(DD, M.E_BIRTH) >= 24) Or (DatePart(MM, M.E_BIRTH) = 10 and DatePart(DD, M.E_BIRTH) <= 23) then '天秤座'
               when (DatePart(MM, M.E_BIRTH) = 10 and DatePart(DD, M.E_BIRTH) >= 24) Or (DatePart(MM, M.E_BIRTH) = 11 and DatePart(DD, M.E_BIRTH) <= 22) then '天蝎座'
               when (DatePart(MM, M.E_BIRTH) = 11 and DatePart(DD, M.E_BIRTH) >= 23) Or (DatePart(MM, M.E_BIRTH) = 12 and DatePart(DD, M.E_BIRTH) <= 22) then '射手座'
               when (DatePart(MM, M.E_BIRTH) = 12 and DatePart(DD, M.E_BIRTH) >= 23) Or (DatePart(MM, M.E_BIRTH) = 1 and  DatePart(DD, M.E_BIRTH) <= 20) then '魔羯座'
               else 'NA'
             end,
           -- 到職年份
           Chg_e_rdate_YY=DatePart(YY, Isnull(case when Convert(Varchar(4), M.e_rdate, 112) = '1900' then null else M.e_rdate end, getdate())), 
           -- 到職月份
           Chg_e_rdate_MM=DatePart(MM, Isnull(case when Convert(Varchar(4), M.e_rdate, 112) = '1900' then null else M.e_rdate end, getdate())), 
           -- 離職年份
           Chg_e_ldate_YY=DatePart(YY, M.e_ldate), 
           -- 離職月份
           Chg_e_ldate_MM=DatePart(MM, M.e_ldate), 
           -- 到職天數
           Chg_emp_day=
             case 
               when (Convert(Varchar(4), M.e_ldate, 112) = '1900') then datediff(dd, M.e_rdate, getdate())
               else datediff(dd, M.e_rdate, M.e_ldate)
             end,
           -- 年資
           Chg_emp_sumYM=
             case 
               when (Convert(Varchar(4), M.e_ldate, 112) = '1900')
                 then (datediff(dd, M.e_rdate, getdate()) / 365)+
                       FLOOR(((datediff(dd, M.e_rdate, getdate()) +
                             ((datediff(dd, M.e_rdate, getdate()) * 1.0) / 365) -
                              (datediff(dd, M.e_rdate, getdate()) / 365) * 365) / 30))*0.1
               else  
                      (datediff(dd, M.e_rdate, M.e_ldate) / 365)+
                       FLOOR(((datediff(dd, M.e_rdate, M.e_ldate) +
                             ((datediff(dd, M.e_rdate, M.e_ldate) * 1.0) / 365) -
                              (datediff(dd, M.e_rdate, M.e_ldate) / 365) * 365) / 30))*0.1
             end,
           -- 離職否(Y:離職, N:在職)
           Chg_leave=
             case when (Convert(Varchar(4), M.e_ldate, 112) = '1900') then 'N' else 'Y' end,
           -- 部門
           Chg_dp_name=Rtrim(D1.dp_name),
           Chg_dp_ename=Rtrim(D1.dp_ename),
           Chg_dp_whno=Rtrim(D1.dp_whno),
           Chg_dp_rem=Rtrim(Convert(Varchar(255), D1.dp_rem)),
           -- 組別
           Chg_mst_name=RTrim(D2.mst_name),
           -- 本年度業績目標(是否為業務可以從此判別)
           [Chg_Year_Tot_SaleAmt]= 
             (select Year_Tot_SaleAmt=Isnull(sum(bo_sale), 0)
                from uV_sales_base 
               where bo_class = '1'
                 and bo_yr=DatePart(YY, GetDate())
                 and bo_no=M.e_no
             ),
           case when (e_birth = '1900/01/01' or e_birth >= getdate()) then 0 else DateDiff(year, e_birth, getdate()) end as Chg_e_Age,
           D3.Chg_Master_Dept, D3.Chg_Duty_Level, D3.Chg_Dept_Level, D3.stock_amt_Level,
           D3.Result_e_no, D3.Result_dept_level, D3.Result_Duty_Level, D3.Result_e_mstno,
           update_datetime = getdate(),
           pemploy_timestamp = m.Timestamp_column
           into Fact_pemploy
      from SYNC_TA13.dbo.pemploy M
           left join SYNC_TA13.dbo.pdept D1 on M.e_dept=D1.dp_no
           left join SYNC_TA13.dbo.masterm D2 on M.e_mstno=D2.mst_no
           left join uV_TA13#Emp_Level D3 on M.e_no=D3.e_no
           
    /*==============================================================*/
    /* Index: pemploy_Timestamp                                     */
    /*==============================================================*/
    --Set @Msg = '建立索引 [Fact_Pemploy.pemploy_Timestamp]'
    --Print @Msg
    --if exists (select * from sys.indexes where object_id = object_id('[dbo].[fact_pemploy]') and name = 'pk_pemploy')
    --   alter table [dbo].[pk_pemploy] drop constraint [pk_pemploy]

    --alter table [dbo].[fact_pemploy] add  constraint [pk_pemploy] primary key nonclustered ([pemploy_timestamp] asc) with 
    --(pad_index  = off, statistics_norecompute  = off, sort_in_tempdb = off, ignore_dup_key = off, online = off, allow_row_locks  = on, allow_page_locks  = on) on [primary]

    /*==============================================================*/
    /* Index: epno                                                  */
    /*==============================================================*/
    Set @Msg = '建立索引 [Fact_Pemploy.epno]'
    Print @Msg
    if exists (select 1 from sysindexes where id = object_id('dbo.fact_pemploy') and name  = 'epno' and indid > 0 and indid < 255) drop index dbo.fact_pemploy.epno
    create clustered index epno on dbo.fact_pemploy (e_no) with fillfactor= 30 on "PRIMARY"

    /*==============================================================*/
    /* Index: bkno                                                  */
    /*==============================================================*/
    Set @Msg = '建立索引 [Fact_Pemploy.bkno]'
    Print @Msg
    if exists (select 1 from sysindexes where id = object_id('dbo.fact_pemploy') and name  = 'bkno' and indid > 0 and indid < 255) drop index dbo.fact_pemploy.bkno
    create index bkno on dbo.fact_pemploy (e_bkno) with fillfactor= 30 on "PRIMARY"

    /*==============================================================*/
    /* Index: dept                                                  */
    /*==============================================================*/
    Set @Msg = '建立索引 [Fact_Pemploy.dept]'
    Print @Msg
    if exists (select 1 from sysindexes where id = object_id('dbo.fact_pemploy') and name  = 'dept' and indid > 0 and indid < 255) drop index dbo.fact_pemploy.dept
    create index dept on dbo.fact_pemploy (e_dept) with fillfactor= 30 on "PRIMARY"

    /*==============================================================*/
    /* Index: epnm                                                  */
    /*==============================================================*/
    Set @Msg = '建立索引 [Fact_Pemploy.epnm]'
    Print @Msg
    if exists (select 1 from sysindexes where id = object_id('dbo.fact_pemploy') and name  = 'epnm' and indid > 0 and indid < 255) drop index dbo.fact_pemploy.epnm
    create index epnm on dbo.fact_pemploy (e_name) with fillfactor= 30 on "PRIMARY"

    /*==============================================================*/
    /* Index: ldate                                                 */
    /*==============================================================*/
    Set @Msg = '建立索引 [Fact_Pemploy.ldate]'
    Print @Msg
    if exists (select 1 from sysindexes where id = object_id('dbo.fact_pemploy') and name  = 'ldate' and indid > 0 and indid < 255) drop index dbo.fact_pemploy.ldate
    create index ldate on dbo.fact_pemploy (e_ldate) with fillfactor= 30 on "PRIMARY"
    
    /*==============================================================*/
    /* Index: mstno                                                 */
    /*==============================================================*/
    Set @Msg = '建立索引 [Fact_Pemploy.mstno]'
    Print @Msg
    if exists (select 1 from sysindexes where id = object_id('dbo.fact_pemploy') and name  = 'mstno' and indid > 0 and indid < 255) drop index dbo.fact_pemploy.mstno
    create index mstno on dbo.fact_pemploy (e_mstno) with fillfactor= 30 on "PRIMARY"

  end try
  begin catch
    set @Result = @Err_Code
    set @Msg = @Proc+'...(錯誤訊息:'+ERROR_MESSAGE()+', '+@Msg+')...(錯誤列:'+Convert(Varchar(10), ERROR_LINE())+')'

    Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Result
  end catch

  Return(@Result)
end
GO
