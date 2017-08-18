USE [DW]
GO
/****** Object:  StoredProcedure [dbo].[uSP_ETL_swaredt]    Script Date: 07/24/2017 14:43:59 ******/
DROP PROCEDURE [dbo].[uSP_ETL_swaredt]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
Create Proc [dbo].[uSP_ETL_swaredt]
as
begin
  /*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
   Program Name: uSP_ETL_swaredt
   Create Date: 2015/01/20
   Creator: Rickliu
   Updated Date: 
   Desc:
   1. 抓取期初請以 迄年月抓取之前之所有總數
   ex: 
    select wd_skno, wd_no, wd_ym, sum(wd_amt) as Bf_Qty
      from fact_swaredt
     where wd_ym < '2017/07'
       and wd_skno ='AA150088'
     group by wd_skno, wd_no, wd_ym
   
  *-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*/  
  Declare @Proc Varchar(50) = 'uSP_ETL_swaredt'
  Declare @Cnt Int =0
  Declare @Cnt_Ori Int =0
  Declare @Err_Code Int = -1
  Declare @Msg Varchar(Max) =''
  Declare @strSQL Varchar(Max)=''
  Declare @CR Varchar(4) = char(13)+char(10)
  Declare @Result Int = 0
  
  -- 2014/07/02 -- Rickliu 採購美如要求將他所提供的新品列表之到貨年月，更新至凌越的商品基本資料內。
  IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[Fact_swaredt]') AND type in (N'U'))
  begin
     Set @Msg = '刪除資料表 [Fact_swaredt]'
     set @strSQL= 'DROP TABLE [dbo].[Fact_swaredt]'

     Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
  end

  begin try
    set @Msg = '[產出 Fact_swaredt 庫存資料表]...'
    Exec uSP_Sys_Write_Log @Proc, @Msg, @Msg, 0

    ;With CTE_Q1 as (
      select distinct rtrim(wd_no) as wd_no, rtrim(wd_skno) as wd_skno, rtrim(d.wh_name) as wh_name
        from Sync_TA13.dbo.swaredt m
             left join Sync_TA13.dbo.sware d
               on m.wd_no = d.wh_no
              and len(wh_no) > 1
       where m.wd_class='0'
         and m.wd_no <> 'A'
         
    ), CTE_Q2 as (
      select distinct cal_year, convert(varchar(7), cal_date, 111) as cal_ym
        from Calendar
       where cal_year >= '2012'
         and cal_year <= Year(getdate())
    ), CTE_Q3 as (
      select wd_no, wd_skno, wd_yr, 1 as wd_mo, Convert(Varchar(4), wd_yr)+'/01' as wd_ym, wd_amt1 as wd_amt, wd_ave1 as wd_ave
        from Sync_TA13.dbo.swaredt 
       where wd_class = '0'
         and wd_no <> 'A'
       union
      select wd_no, wd_skno, wd_yr, 2 as wd_mo, Convert(Varchar(4), wd_yr)+'/02' as wd_ym, wd_amt2 as wd_amt, wd_ave2 as wd_ave
        from Sync_TA13.dbo.swaredt 
       where wd_class = '0'
         and wd_no <> 'A'
       union
      select wd_no, wd_skno, wd_yr, 3 as wd_mo, Convert(Varchar(4), wd_yr)+'/03' as wd_ym, wd_amt3 as wd_amt, wd_ave3 as wd_ave
        from Sync_TA13.dbo.swaredt 
       where wd_class = '0'
         and wd_no <> 'A'
       union
      select wd_no, wd_skno, wd_yr, 4 as wd_mo, Convert(Varchar(4), wd_yr)+'/04' as wd_ym, wd_amt4 as wd_amt, wd_ave4 as wd_ave
        from Sync_TA13.dbo.swaredt 
       where wd_class = '0'
         and wd_no <> 'A'
       union
      select wd_no, wd_skno, wd_yr, 5 as wd_mo, Convert(Varchar(4), wd_yr)+'/05' as wd_ym, wd_amt5 as wd_amt, wd_ave5 as wd_ave
        from Sync_TA13.dbo.swaredt 
       where wd_class = '0'
         and wd_no <> 'A'
       union
      select wd_no, wd_skno, wd_yr, 6 as wd_mo, Convert(Varchar(4), wd_yr)+'/06' as wd_ym, wd_amt6 as wd_amt, wd_ave6 as wd_ave
        from Sync_TA13.dbo.swaredt 
       where wd_class = '0'
         and wd_no <> 'A'
       union
      select wd_no, wd_skno, wd_yr, 7 as wd_mo, Convert(Varchar(4), wd_yr)+'/07' as wd_ym, wd_amt7 as wd_amt, wd_ave7 as wd_ave
        from Sync_TA13.dbo.swaredt 
       where wd_class = '0'
         and wd_no <> 'A'
       union
      select wd_no, wd_skno, wd_yr, 8 as wd_mo, Convert(Varchar(4), wd_yr)+'/08' as wd_ym, wd_amt8 as wd_amt, wd_ave8 as wd_ave
        from Sync_TA13.dbo.swaredt 
       where wd_class = '0'
         and wd_no <> 'A'
       union
      select wd_no, wd_skno, wd_yr, 9 as wd_mo, Convert(Varchar(4), wd_yr)+'/09' as wd_ym, wd_amt9 as wd_amt, wd_ave9 as wd_ave
        from Sync_TA13.dbo.swaredt 
       where wd_class = '0'
         and wd_no <> 'A'
       union
      select wd_no, wd_skno, wd_yr, 10 as wd_mo, Convert(Varchar(4), wd_yr)+'/10' as wd_ym, wd_amt10 as wd_amt, wd_ave10 as wd_ave
        from Sync_TA13.dbo.swaredt 
       where wd_class = '0'
         and wd_no <> 'A'
       union
      select wd_no, wd_skno, wd_yr, 11 as wd_mo, Convert(Varchar(4), wd_yr)+'/11' as wd_ym, wd_amt11 as wd_amt, wd_ave11 as wd_ave
        from Sync_TA13.dbo.swaredt 
       where wd_class = '0'
         and wd_no <> 'A'
       union
      select wd_no, wd_skno, wd_yr, 12 as wd_mo, Convert(Varchar(4), wd_yr)+'/12' as wd_ym, wd_amt12 as wd_amt, wd_ave12 as wd_ave
        from Sync_TA13.dbo.swaredt 
       where wd_class = '0'
         and wd_no <> 'A'
     )

    select m.wd_no, wh_name, wd_yr, wd_ym, 
           m.wd_skno, rtrim(d2.sk_name) as sk_name,
           isnull(wd_amt, 0) as wd_amt,
           isnull(wd_ave, 0) as wd_ave,
           isnull(sk_save, 0) as sk_save,
           isnull(wd_amt, 0) * isnull(sk_save, 0)as wd_save_tot,
           d.*,
           chg_skno_bkind3, chg_skno_bkind_name3, 
           swaredt_update_datetime = getdate()
           into Fact_swaredt
      from CTE_Q1 m
           cross join CTE_Q2 d
           left join CTE_Q3 d1
            on m.wd_skno = d1.wd_skno
           and m.wd_no = d1.wd_no
           and d.cal_ym = d1.wd_ym
           left join Fact_sstock d2
            on m.wd_skno = d2.sk_no
     where wd_yr >= year(DateAdd(year, -5, getdate()))
    /*==============================================================*/
    /* Index: wdskno                                                */
    /*==============================================================*/
    Set @Msg = '建立索引 [Fact_swaredt.wdskno]'
    Print @Msg
    if exists (select 1 from sysindexes where id = object_id('dbo.fact_swaredt') and name  = 'wdskno' and indid > 0 and indid < 255) drop index dbo.fact_swaredt.no
       create  clustered  index [wdskno] on [dbo].[fact_swaredt]([wd_skno], [wd_no], [wd_yr]) on [primary]

    /*==============================================================*/
    /* Index: wdno                                                  */
    /*==============================================================*/
    Set @Msg = '建立索引 [Fact_swaredt.wdno]'
    Print @Msg
    if exists (select 1 from sysindexes where id = object_id('dbo.fact_swaredt') and name  = 'wdno' and indid > 0 and indid < 255) drop index dbo.fact_swaredt.no
       create  index [wdno] on [dbo].[fact_swaredt]([wd_no], [wd_skno], [wd_yr]) on [primary]

    /*==============================================================*/
    /* Index: wdym                                                  */
    /*==============================================================*/
    Set @Msg = '建立索引 [Fact_swaredt.wdym]'
    Print @Msg
    if exists (select 1 from sysindexes where id = object_id('dbo.fact_swaredt') and name  = 'wdym' and indid > 0 and indid < 255) drop index dbo.fact_swaredt.no
       create  index [wdyr] on [dbo].[fact_swaredt]([wd_no], [wd_yr], [wd_ym], [wd_skno]) on [primary]

    Exec uSP_Sys_Write_Log @Proc, @Msg, @Msg, @Cnt
  end try
  begin catch
    set @Result = @Err_Code
    set @Msg = @Proc+'...(錯誤訊息:'+ERROR_MESSAGE()+', '+@Msg+')...(錯誤列:'+Convert(Varchar(10), ERROR_LINE())+')'

    Exec uSP_Sys_Write_Log @Proc, @Msg, @Msg, @Cnt
  end catch
end
GO
