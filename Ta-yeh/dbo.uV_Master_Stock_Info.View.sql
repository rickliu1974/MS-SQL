USE [DW]
GO
/****** Object:  View [dbo].[uV_Master_Stock_Info]    Script Date: 08/18/2017 17:18:53 ******/
DROP VIEW [dbo].[uV_Master_Stock_Info]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
Create View [dbo].[uV_Master_Stock_Info]
as
  With CTE_STOCK as
  (select distinct 
          rtrim(sk_no) as sk_no
     from ori_xls#master_stock
  )
  select m.sk_no, rtrim(d.sk_name) as sk_name, 
         rtrim(isnull(d.sk_color, '')) as sk_color, 
         rtrim(isnull(d.sk_size, '')) as sk_size,
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
         Chg_WD_AB_last_diff_Qty = sum(Isnull(D.Chg_WD_AB_last_diff_Qty, 0)),
         stock_kind_list
    from CTE_STOCK m
         inner join fact_sstock d
            on m.sk_no collate Chinese_PRC_BIN =d.sk_no collate Chinese_PRC_BIN
   group by m.sk_no, d.sk_name, d.sk_color, d.sk_size, 
            d.chg_is_master_stock, d.chg_is_new_stock, d.chg_is_Dead_Stock,
            stock_kind_list
GO
