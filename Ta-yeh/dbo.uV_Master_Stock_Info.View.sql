USE [DW]
GO
/****** Object:  View [dbo].[uV_Master_Stock_Info]    Script Date: 07/24/2017 14:43:55 ******/
DROP VIEW [dbo].[uV_Master_Stock_Info]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
Create View [dbo].[uV_Master_Stock_Info]
as
  With CTE_STOCK as
  (select distinct sk_no, sk_mname
     from ori_xls#master_stock
  )
  select m.sk_no, d.sk_name, m.sk_mname, 
         -- 2014/06/17 Rick 新增 AA 期初安全存量
         Chg_WD_AA_first_sQty = sum(Isnull(D.Chg_WD_AA_sQty, 0)),
         -- 2014/06/17 Rick 新增 AA 期初現有庫存數
         Chg_WD_AA_first_Qty = sum(Isnull(D.Chg_WD_AA_first_Qty, 0)),           
         -- 2014/06/17 Rick 新增 AA 期初庫存差異數
         Chg_WD_AA_first_diff_Qty = sum(Isnull(D.Chg_WD_AA_first_diff_Qty, 0)),
         -- 2014/06/07 Rick 新增 AA 期末現有庫存數
         Chg_WD_AA_last_Qty = sum(Isnull(D.Chg_WD_AA_last_Qty, 0)),
         -- 2014/06/17 Rick 新增 AA 期末庫存差異數
         Chg_WD_AA_last_diff_Qty = sum(Isnull(D.Chg_WD_AA_last_diff_Qty, 0)),
           
         -- 2014/06/17 Rick 新增 AB 期初安全存量
         Chg_WD_AB_first_sQty = sum(Isnull(D.Chg_WD_AB_sQty, 0)),
         -- 2014/06/17 Rick 新增 AB 期初現有庫存數
         Chg_WD_AB_first_Qty = sum(Isnull(D.Chg_WD_AB_first_Qty, 0)),           
         -- 2014/06/17 Rick 新增 AB 期初庫存差異數
         Chg_WD_AB_first_diff_Qty = sum(Isnull(D.Chg_WD_AB_first_diff_Qty, 0)),
         -- 2014/06/07 Rick 新增 AB 期末現有庫存數
         Chg_WD_AB_last_Qty = sum(Isnull(D.Chg_WD_AB_last_Qty, 0)),
         -- 2014/06/17 Rick 新增 AB 期末庫存差異數
         Chg_WD_AB_last_diff_Qty = sum(Isnull(D.Chg_WD_AB_last_diff_Qty, 0))
    from CTE_STOCK m
         inner join fact_sstock d
            on m.sk_no collate Chinese_PRC_BIN =d.sk_no collate Chinese_PRC_BIN
   group by m.sk_no, d.sk_name, sk_mname
GO
