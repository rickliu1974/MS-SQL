USE [DW]
GO
/****** Object:  StoredProcedure [dbo].[xuSP_ETL_Cust_NonSale_Stock(201708停用)]    Script Date: 08/18/2017 17:18:56 ******/
DROP PROCEDURE [dbo].[xuSP_ETL_Cust_NonSale_Stock(201708停用)]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE Proc [dbo].[xuSP_ETL_Cust_NonSale_Stock(201708停用)]
as
begin
  /*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
   Program Name: uSP_ETL_Cust_NonSale_Stock
   Create Date: 2014/06/05
   Creator: Rickliu
  *-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*/  

  Declare @Proc Varchar(50) = 'uSP_ETL_Cust_NonSale_Stock'
  Declare @Cnt Int =0
  Declare @Msg Varchar(Max) =''
  Declare @strSQL Varchar(Max)=''
  Declare @CR Varchar(4) = char(13)+char(10)
  Declare @Result int = 0
  
  IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[Cust_NonSale_Stock]') AND type in (N'U'))
  begin
     Set @Msg = '刪除資料表 [Cust_NonSale_Stock]'
     set @strSQL= 'DROP TABLE [dbo].[Cust_NonSale_Stock]'

     Exec @Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
     if @Result <>  -1 Set @Result = 0
  end

  set @Msg = 'Create to [Cust_NonSale_Stock]...'
  begin try
    select *,
           0 as dcount_ctno, 
           0 as dcount_ctno8, -- 2017/02/08 Rickliu Add 增加八碼客戶編號
           0 as dcount_skno,
           update_datetime = getdate()
           into Cust_NonSale_Stock
      from (select m.*, d.*,
                   isnull(d1.sd_qty, 0) as sd_qty,
                   isnull(d1.sd_stot, 0) as sd_stot,
                   isnull(d1.sd2_qty, 0) as sd2_qty,   --20150904 Add by Nan 增加託售及回貨計算
                   isnull(d1.sd2_stot, 0) as sd2_stot, --20150904 Add by Nan 增加託售及回貨計算
                   isnull(d1.sd6_qty, 0) as sd6_qty,   --20150904 Add by Nan 增加託售及回貨計算
                   isnull(d1.sd6_stot, 0) as sd6_stot, --20150904 Add by Nan 增加託售及回貨計算
                   isnull(d1.sd7_qty, 0) as sd7_qty,   --20150904 Add by Nan 增加託售及回貨計算
                   isnull(d1.sd7_stot, 0) as sd7_stot, --20150904 Add by Nan 增加託售及回貨計算
                   case
                     when isnull(sd_qty, 0) = 0 then '2'
                     else '1'
                   end sale_flag
              from (select ct_no8, ct_sname8, -- 2017/02/08 Rickliu Add 增加八碼客戶編號
                           ct_sales, chg_ct_sales_name,
                           ct_fld3, 
                           chg_hunderd_customer, Chg_Hunderd_Customer_Name, 
                           chg_is_lan_custom
                      from fact_pcust
                     where 1=1
                       and ct_class = 1
                       --and len(ct_no) =  9
                       and substring(ct_no, 9, 1) =  '1'  --20150428 Modify by Nan 配合宜靜取得客編尾碼為1資料
                       and substring(ct_no, 1, 2) not in ('IT', 'IZ')
                   ) m
                   cross join
                   (select sk_no, sk_name, sk_save
                           Chg_Wd_AA_first_Qty, Chg_Wd_AB_first_Qty,
                           Chg_Wd_AA_last_Qty, Chg_Wd_AB_last_Qty,

                           Chg_Wd_AA_first_Qty * sk_save as Chg_Wd_AA_first_amt, 
                           Chg_Wd_AB_first_Qty * sk_save as Chg_Wd_AB_first_amt, 
                           Chg_Wd_AA_last_Qty * sk_save as Chg_Wd_AA_last_amt, 
                           Chg_Wd_AB_last_Qty * sk_save as Chg_Wd_AB_last_amt,

                           Chg_skno_accno, Chg_skno_accno_Name, Chg_skno_BKind, 
                           Chg_skno_Bkind_Name, Chg_skno_SKind, Chg_skno_SKind_Name, 
                           Chg_kind_Name, Chg_is_dead_stock, chg_dead_stock_ym, chg_new_arrival_ym
                      from fact_sstock
                     where 1=1
                       and substring(sk_no, 1, 1) = 'A'
                       and sk_name <> ''
                       and chg_stock_nonsales = 'Y'
                   ) d
                   left join 
                   (select ct_no8, sd_skno, sum(sd_qty) as sd_qty, sum(sd_stot) as sd_stot
										   , case    --20150904 Add by Nan 增加託售及回貨計算
											   when sd_slip_fg = '2' then SUM(Isnull(sd_qty, 0))
											   else 0
											 end as sd2_qty,
											 case
											   when sd_slip_fg = '6' then SUM(Isnull(sd_qty, 0))
											   else 0
											 end as sd6_qty,
											 case
											   when sd_slip_fg = '7' then SUM(Isnull(sd_qty, 0))
											   else 0
											 end as sd7_qty,
											 case
											   when sd_slip_fg = '2' then SUM(Isnull(sd_stot, 0))
											   else 0
											 end as sd2_stot,
											 case
											   when sd_slip_fg = '6' then SUM(Isnull(sd_stot, 0))
											   else 0
											 end as sd6_stot,
											 case
											   when sd_slip_fg = '7' then SUM(Isnull(sd_stot, 0))
											   else 0
											 end as sd7_stot   --20150904 Add by Nan 增加託售及回貨計算
                      from fact_sslpdt
                     where 1=1
                       --and sd_class = '1'             --20150904 Del by Nan 增加託售及回貨計算
					   --and sd_slip_fg in('2', '6', '7') --20150904 Add by Nan 增加託售及回貨計算
                        and (sd_slip_fg  = '2'
                        or sd_slip_fg  = '6'
                        or sd_slip_fg  = '7')
                       --and len(sd_ctno) =  9
                       and substring(sd_ctno, 9, 1) =  '1'  --20150428 Modify by Nan 配合宜靜取得客編尾碼為1資料
                       --and substring(sd_ctno, 1, 2) not in ('IT', 'IZ')
                       and (sd_ctno not like 'IT%' or sd_ctno not like 'IZ%')
                       and chg_stock_nonsales = 'Y'
                       -- 近兩年資料
                       and sd_date >= 
                           case
                             when Convert(datetime, '2012/12/01') >= dateadd(yy, -2, getdate())
                             then Convert(datetime, '2012/12/01')
                             else dateadd(yy, -2, getdate())
                           end
                     group by ct_no8, sd_skno,sd_slip_fg --20150904 Modify by Nan 增加託售及回貨計算
                   ) d1
                     on m.ct_no8 collate Chinese_Taiwan_Stroke_CI_AS = d1.ct_no8 collate Chinese_Taiwan_Stroke_CI_AS
                    and d.sk_no collate Chinese_Taiwan_Stroke_CI_AS = d1.sd_skno collate Chinese_Taiwan_Stroke_CI_AS
           ) m

     -- 2015/01/26 Rickliu 增加業務的客戶數量及商品數量
     update Cust_NonSale_Stock
        set dcount_ctno = d.dcount_ctno,
            dcount_skno = d.dcount_skno
       from Cust_NonSale_Stock m
            left join 
            (select ct_sales, 
                    Count(Distinct ct_no8) as dcount_ctno, -- 2017/02/08 Rickliu Add 增加八碼客戶編號
                    Count(Distinct sk_no) as dcount_skno
               from Cust_NonSale_Stock
              group by ct_sales
            ) d on m.ct_sales = d.ct_sales
            
     select @Cnt = count(1) from Cust_NonSale_Stock
     
     if @Cnt = 0 
        Set @Result = -1
     else
        Set @Result = @Cnt

     Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt
  end try
  begin catch
    set @Msg = @Msg+'(錯誤訊息:'+ERROR_MESSAGE()+')'
    set @Result = -1
    Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Result
  end catch

  Return(@Result)

end
GO
