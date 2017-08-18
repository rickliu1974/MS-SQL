USE [DW]
GO
/****** Object:  View [dbo].[uV_swaredt]    Script Date: 08/18/2017 17:43:39 ******/
DROP VIEW [dbo].[uV_swaredt]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
create view [dbo].[uV_swaredt]
as
    select m.wd_no, d.cal_year as wd_yr, d.cal_ym as wd_ym, 
           m.wd_skno,
           isnull(d1.wd_amt, 0) as wd_amt,
           isnull(d1.wd_ave, 0) as wd_ave,
           isnull(d1.wd_amt, 0) * isnull(sk_save, 0)as wd_save_tot,
           d2.*
      from (select distinct wd_no, wd_skno
              from Sync_TA13.dbo.swaredt
             where wd_class='0'
               and wd_no <> 'A'
           ) m
           cross join
           (select distinct cal_year, convert(varchar(7), cal_date, 111) as cal_ym
              from Calendar
             where cal_year >= '2012'
               and cal_year <= Year(getdate())
           ) d
           left join
           (select wd_no, wd_skno, wd_yr, 1 as wd_mo, Convert(Varchar(4), wd_yr)+'/01' as wd_ym, wd_amt1 as wd_amt, wd_ave1 as wd_ave
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
           ) d1
            on m.wd_skno = d1.wd_skno
           and m.wd_no = d1.wd_no
           and d.cal_ym = d1.wd_ym
          left join Fact_sstock d2
            on m.wd_skno = d2.sk_no
GO
