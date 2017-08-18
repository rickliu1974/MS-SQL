USE [DW]
GO
/****** Object:  View [dbo].[uV_Sale_members]    Script Date: 08/18/2017 17:18:53 ******/
DROP VIEW [dbo].[uV_Sale_members]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
Create View [dbo].[uV_Sale_members]
as
  --select distinct top 1000
  --       e_dept, chg_dp_name, e_no, e_name, e_duty, 
  --       e_mstno, chg_mst_name, e_rdate, chg_e_sex, chg_e_birth, 
  --       chg_emp_day, chg_emp_sumym, chg_leave
  --  from Fact_sslip M inner join RealTime_SaleData d on m.e_no collate Chinese_Taiwan_Stroke_CI_AS = d.area
  -- where 1=1
  --   --and chg_year_tot_saleamt <> 0
  --   and e_no is not null
  --   and E_DEPT like 'B2%'
  --   and E_DEPT <> 'B2910'
     
  -- order by 13, 6, 1, 3

  select distinct top 1000
         e_dept, chg_dp_name, e_no, e_name, e_duty, 
         e_mstno, chg_mst_name, e_rdate, chg_e_sex, chg_e_birth, 
         chg_emp_day, chg_emp_sumym, chg_leave
    from RealTime_SaleData as m 
         left join Fact_pemploy as d 
           on m.area=d.e_no collate Chinese_Taiwan_Stroke_CI_AS 
   where 1=1
     and kind='00005'
     and e_no is not null
     and (chg_year_tot_saleamt <> 0 or Chg_leave='N')
   order by 13, 6, 1, 3
GO
