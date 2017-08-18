USE [DW]
GO
/****** Object:  UserDefinedFunction [dbo].[uFn_Get_Custom_List]    Script Date: 08/18/2017 17:43:41 ******/
DROP FUNCTION [dbo].[uFn_Get_Custom_List]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
Create Function [dbo].[uFn_Get_Custom_List](@Close_Date Varchar(10), @Is_hunderd_customer Int = 1)
returns @Result Table
(
  NO Int, 
  Chg_bu_no Varchar(100), 
  bu_name Varchar(100),
  bu_full_name varchar(100),
  ct_open_cnt Int,
  ct_close_cnt Int,
  ct_bu_cnt Int,
  ct_total_cnt Int,
  Info Varchar(2000),
  List Varchar(2000),
  Open_List Varchar(2000),
  Close_List Varchar(2000)
)
as
Begin
Declare @sIs_hunderd_customer Varchar(1)
   set @sIs_hunderd_customer = case when @Is_hunderd_customer = 1 then 'Y' else '' end
   
   if isdate(@Close_Date)=1 
      set @Close_Date = Convert(Varchar(10), Convert(DateTime, @Close_Date), 111)
   else
      set @Close_Date = Convert(Varchar(10), getdate(), 111)

  ;with CTE_Q1 as (
     select distinct chg_bu_no, ct_no8, ct_fld3, ct_fld4,
            case 
              when (rtrim(ct_fld2) <> '') and isdate(rtrim(ct_fld2))=1 then
                case
                  when convert(varchar(10), convert(datetime, rtrim(ct_fld2)), 111) <=@Close_Date then '-關'
                  else ''
                 end
              when chg_is_bu = 'Y' then '-總'
              else '' 
            end as ct_close
       from fact_pcust m
            left join fact_pemploy d
              on m.ct_sales = d.e_no
      where ct_class = '1'
        and Len(ct_no8) <>''
        and case 
              when @Is_hunderd_customer = 1 then chg_hunderd_customer 
              else ''
            end =''+@sIs_hunderd_customer+''
        and e_dept like 'B%'
        and substring(chg_bu_no, 1, 2) not in ('IT', 'IZ', 'ZZ', '')
  ), CTE_Q2 as (
     select chg_bu_no, ct_fld3, '('+chg_bu_no+')'+ct_fld3 as bu_name,
            case when ct_close not like '%總' and ct_close = '' then count(1) else 0 end as ct_open_cnt,
            case when ct_close not like '%總' and ct_close <> '' then count(1) else 0 end as ct_close_cnt,
            case when ct_close like '%總' and ct_close <> '' then count(1) else 0 end as ct_bu_cnt,
            count(1) as ct_total_cnt
       from CTE_Q1
      group by chg_bu_no, ct_fld3, ct_close
  ), CTE_Q3 as (
     select chg_bu_no, ct_fld3 as bu_name, bu_name as bu_full_name, --I0001020
            '[開:'+
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

  Insert Into @Result
  select ROW_NUMBER() Over (order by chg_bu_no) as NO,
         chg_bu_no, bu_name, bu_full_name,
         ct_open_cnt, ct_close_cnt, ct_bu_cnt, ct_total_cnt,
         Info,
         Isnull(Reverse(Substring(Reverse(
           (select ct_fld4+'('+ct_no8+ct_close+'),'
              from CTE_Q1 d
             where m.chg_bu_no = d.chg_bu_no
              for xml path('')
           )), 2, 8000)
         ), '') as List,
         Isnull(Reverse(Substring(Reverse(
           (select ct_fld4+'('+ct_no8+'),'
              from CTE_Q1 d
             where m.chg_bu_no = d.chg_bu_no
               and ct_close not like '%關%'
              for xml path('')
           )), 2, 8000)
         ), '') as Open_List,
         Isnull(Reverse(Substring(Reverse(
           (select ct_fld4+'('+ct_no8+'),'
              from CTE_Q1 d
           where 1=1
	       and m.chg_bu_no = d.chg_bu_no
               and ct_close like '%關%'
              for xml path('')
           )), 2, 8000)
         ), '') as Close_List
    from CTE_Q3 m
   union
  select 0 as NO, '000000' as chg_bu_no, 
         'Total' as bu_name, '總計' as bu_full_name,
         Sum(ct_open_cnt) as ct_open_cnt,
         Sum(ct_close_cnt) as ct_close_cnt,
         Sum(ct_bu_cnt) as ct_bu_cnt,
         Sum(ct_total_cnt) as ct_total_cnt,
         '[開:'+convert(varchar(10), Sum(ct_open_cnt))+
         ', 關:'+convert(varchar(10), Sum(ct_close_cnt))+
         ', 總:'+convert(varchar(10), Sum(ct_bu_cnt))+
         ', 計：'+convert(varchar(10), Sum(ct_total_cnt))+
         ']' as Info,
         '關店日期統計至 ['+isnull(@Close_Date, convert(Varchar(10), getdate(), 111))+']' as List,
         '' as Open_List,
         '' as Close_List
    from CTE_Q2
  return
end
GO
