USE [DW]
GO
/****** Object:  View [dbo].[uV_Current_Hunderd_List]    Script Date: 08/18/2017 17:43:39 ******/
DROP VIEW [dbo].[uV_Current_Hunderd_List]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
Create View [dbo].[uV_Current_Hunderd_List]
as
  with CTE_Q1 as (
     select distinct chg_bu_no, ct_no8, ct_fld3, ct_fld4,
            case 
              when ct_fld2 <> '' then '-關'
              when chg_is_bu = 'Y' then '-總'
              else '' 
            end as ct_close
       from fact_pcust
      where ct_class = '1'
        and chg_hunderd_customer ='Y'
  ), CTE_Q2 as (
     select chg_bu_no, ct_fld3, ct_fld3+'('+chg_bu_no+')' as bu_name,
            case when ct_close not like '%總' and ct_close = '' then count(1) else 0 end as ct_open_cnt,
            case when ct_close not like '%總' and ct_close <> '' then count(1) else 0 end as ct_close_cnt,
            case when ct_close like '%總' and ct_close <> '' then count(1) else 0 end as ct_bu_cnt,
            count(1) as ct_total_cnt
       from CTE_Q1
      group by chg_bu_no, ct_fld3, ct_close
  ), CTE_Q3 as (
     select chg_bu_no, ct_fld3 as bu_name,
            bu_name+' [開:'+
            convert(varchar(10), Sum(ct_open_cnt))+',  關:'+ 
            convert(varchar(10), Sum(ct_close_cnt))+', 總:'+
            convert(varchar(10), Sum(ct_bu_cnt))+', 計:'+
            convert(varchar(10), Sum(ct_total_cnt))+']' as Info,
            Sum(ct_open_cnt) as ct_open_cnt,
            Sum(ct_close_cnt) as ct_close_cnt,
            Sum(ct_bu_cnt) as ct_bu_cnt,
            Sum(ct_total_cnt) as ct_total_cnt
       from CTE_Q2
      group by chg_bu_no, ct_fld3, bu_name
  )
  
  select ROW_NUMBER() Over (order by chg_bu_no) as NO,
         chg_bu_no, bu_name,
         ct_open_cnt, ct_close_cnt, ct_bu_cnt, ct_total_cnt,
         Info,
         Reverse(Substring(Reverse(
           (select ct_fld4+'('+ct_no8+ct_close+'),'
              from CTE_Q1 d
             where m.chg_bu_no = d.chg_bu_no
              for xml path('')
           )), 2, 2000)
         ) as List
    from CTE_Q3 m
   union
  select 999 as NO, 'Total' as chg_bu_no, '總計' as bu_name,
         Sum(ct_open_cnt) as ct_open_cnt,
         Sum(ct_close_cnt) as ct_close_cnt,
         Sum(ct_bu_cnt) as ct_bu_cnt,
         Sum(ct_total_cnt) as ct_total_cnt,
         '總計 [開:'+convert(varchar(10), Sum(ct_open_cnt))+
         ', 關:'+convert(varchar(10), Sum(ct_close_cnt))+
         ', 總公司:'+convert(varchar(10), Sum(ct_bu_cnt))+
         ', 計：'+convert(varchar(10), Sum(ct_total_cnt))+
         ']' as Info,
         '' as List
    from CTE_Q2
GO
