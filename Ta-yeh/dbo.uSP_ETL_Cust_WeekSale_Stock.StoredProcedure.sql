USE [DW]
GO
/****** Object:  StoredProcedure [dbo].[uSP_ETL_Cust_WeekSale_Stock]    Script Date: 07/24/2017 14:43:59 ******/
DROP PROCEDURE [dbo].[uSP_ETL_Cust_WeekSale_Stock]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE Proc [dbo].[uSP_ETL_Cust_WeekSale_Stock]
as
begin
  /*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
   Program Name: uSP_ETL_Cust_WeekSale_Stock
   Create Date: 2015/08/21
   Creator: Nanliao
  *-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*/  

  Declare @Proc Varchar(50) = 'uSP_ETL_Cust_WeekSale_Stock'
  Declare @Cnt Int =0
  Declare @Msg Varchar(Max) =''
  Declare @strSQL Varchar(Max)=''
  Declare @CR Varchar(4) = char(13)+char(10)
  Declare @Result int = 0
  
  IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[Cust_WeekSale_Stock]') AND type in (N'U'))
  begin
     Set @Msg = '刪除資料表 [Cust_WeekSale_Stock]'
     set @strSQL= 'DROP TABLE [dbo].[Cust_WeekSale_Stock]'

     Exec @Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
     if @Result <>  -1 Set @Result = 0
  end

  set @Msg = 'Create to [Cust_WeekSale_Stock]...'
  begin try
    select *,
           0 as all_count_ctno, 0 as all_count_skno,
           0 as master_count_ctno, 0 as master_count_skno,
           0 as nonsale_count_ctno, 0 as nonsale_count_skno,
           0 as newstock_count_ctno, 0 as newstock_count_skno,
           0 as deadstock_count_ctno, 0 as deadstock_count_skno,
           update_datetime = getdate()
           into Cust_WeekSale_Stock
      from (select kind,kind_Name, ct_no8, ct_sname8, chg_ct_fld3, -- 2017/02/08 Rickliu Add 增加八碼客戶編號
                   Chg_Hunderd_Customer_Name, Chg_IS_Lan_Custom,
                   sd_skno, sd_name, 
                   Chg_Wd_AA_first_Qty, Chg_Wd_AB_first_Qty,
                   Chg_Wd_AA_last_Qty, Chg_Wd_AB_last_Qty,
                   Chg_skno_accno, Chg_skno_accno_Name, 
                   Chg_skno_BKind, Chg_skno_Bkind_Name, 
                   Chg_skno_SKind, Chg_skno_SKind_Name, 
                   Chg_kind_Name, Chg_sp_date_YM,

                   cal_week_begin, cal_week_end, cal_week_Range,
                   cal_yearweek, cal_quarter,

                   Chg_sp_sales, Chg_sales_Name, 
                   Chg_sale_Month_Master,
                   Chg_sale_Month_MasterName,
                   Chg_Hunderd_Customer,
                   Chg_Stock_NonSales,
                   Chg_is_dead_stock, 
                   chg_dead_stock_ym, 
                   chg_new_arrival_ym,
				   SUM(Isnull(sd2_qty, 0)) as sd2_qty,
				   SUM(Isnull(sd6_qty, 0)) as sd6_qty,
				   SUM(Isnull(sd7_qty, 0)) as sd7_qty,
				   SUM(Isnull(sd2_stot, 0)) as sd2_stot,
				   SUM(Isnull(sd6_stot, 0)) as sd6_stot,
				   SUM(Isnull(sd7_stot, 0)) as sd7_stot
			  from (select '1' as kind, '已銷' as kind_Name, ct_no8, ct_sname8, chg_ct_fld3,-- 2017/02/08 Rickliu Add 增加八碼客戶編號 
                     Chg_Hunderd_Customer_Name, Chg_IS_Lan_Custom,
                     sd_skno, sd_name, 
                     t.Chg_Wd_AA_first_Qty, t.Chg_Wd_AB_first_Qty,
                     t.Chg_Wd_AA_last_Qty, t.Chg_Wd_AB_last_Qty,
                     t.Chg_skno_accno, t.Chg_skno_accno_Name, 
                     t.Chg_skno_BKind, t.Chg_skno_Bkind_Name, 
                     t.Chg_skno_SKind, t.Chg_skno_SKind_Name, 
                     t.Chg_kind_Name, Chg_sp_date_YM,
                     cal_week_begin, cal_week_end, cal_week_Range,
                     cal_yearweek, cal_quarter,
                     Chg_sp_sales, Chg_sales_Name, 
                     Chg_sale_Month_Master,
                     Chg_sale_Month_MasterName,
                     Chg_Hunderd_Customer,
                     t.Chg_Stock_NonSales,
                     t.Chg_is_dead_stock, 
                     t.chg_dead_stock_ym, 
                     t.chg_new_arrival_ym,
                     case
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
                     end as sd7_stot
               from (select *
                       from Fact_sslpdt
                      where 1=1
                        --and sd_slip_fg in('2', '6', '7')
                        and (sd_slip_fg  = '2'
                        or sd_slip_fg  = '6'
                        or sd_slip_fg  = '7')
                    ) t 
                    inner join Calendar d
                       on cal_week ='1'
                      and Convert(varchar(7), cal_week_begin, 111) >= '2012/12'
                      and Convert(varchar(7), cal_week_end, 111) <= Convert(varchar(7), getdate(), 111)
                      and t.sd_date BETWEEN d.cal_week_begin AND cal_week_end 
              where 1=1
                and len(rtrim(t.sd_ctno)) >= 8
              group by sd_slip_fg, Chg_sp_date_YM, ct_no8, ct_sname8, chg_ct_fld3, -- 2017/02/08 Rickliu Add 增加八碼客戶編號
                       Chg_Hunderd_Customer_Name, Chg_IS_Lan_Custom,
                       sd_skno, sd_name, 
                       t.Chg_Wd_AA_first_Qty, t.Chg_Wd_AB_first_Qty,
                       t.Chg_Wd_AA_last_Qty, t.Chg_Wd_AB_last_Qty,
                       t.Chg_skno_accno, t.Chg_skno_accno_Name, 
                       t.Chg_skno_BKind, t.Chg_skno_Bkind_Name, 
                       t.Chg_skno_SKind, t.Chg_skno_SKind_Name, 
                       t.Chg_kind_Name, Chg_sp_sales, Chg_sales_Name, 
                       Chg_sale_Month_Master, Chg_sale_Month_MasterName,
                       Chg_Hunderd_Customer,
                       t.Chg_Stock_NonSales,
                       t.Chg_is_dead_stock, 
                       t.chg_dead_stock_ym, 
                       t.chg_new_arrival_ym,
                       cal_week_begin, cal_week_end, cal_week_Range,
                       cal_yearweek, cal_quarter
					 ) a
			group by kind,kind_Name, ct_no8, ct_sname8, chg_ct_fld3, -- 2017/02/08 Rickliu Add 增加八碼客戶編號
                     Chg_Hunderd_Customer_Name, Chg_IS_Lan_Custom,
                     sd_skno, sd_name, 
                     Chg_Wd_AA_first_Qty, Chg_Wd_AB_first_Qty,
                     Chg_Wd_AA_last_Qty, Chg_Wd_AB_last_Qty,
                     Chg_skno_accno, Chg_skno_accno_Name, 
                     Chg_skno_BKind, Chg_skno_Bkind_Name, 
                     Chg_skno_SKind, Chg_skno_SKind_Name, 
                     Chg_kind_Name, Chg_sp_date_YM,

                     cal_week_begin, cal_week_end, cal_week_Range,
                     cal_yearweek, cal_quarter,

                     Chg_sp_sales, Chg_sales_Name, 
                     Chg_Hunderd_Customer,
                     Chg_Stock_NonSales,
                     Chg_is_dead_stock, 
                     chg_dead_stock_ym, 
                     chg_new_arrival_ym,
                     Chg_sale_Month_Master,
                     Chg_sale_Month_MasterName

           ) m

     -- 2015/01/26 Rickliu 增加業務的客戶數量及商品數量
     update Cust_WeekSale_Stock
        set all_count_ctno = isnull(d.all_count_ctno,0),
            all_count_skno = isnull(d.all_count_skno,0)
       from Cust_WeekSale_Stock m
            left join 
            (select Chg_sp_sales, cal_week_Range,
                    Count(Distinct ct_no8) as all_count_ctno,
                    Count(Distinct sd_skno) as all_count_skno
               from Cust_WeekSale_Stock
              group by Chg_sp_sales,cal_week_Range
            ) d on m.Chg_sp_sales = d.Chg_sp_sales and m.cal_week_Range=d.cal_week_Range
			
			
     update Cust_WeekSale_Stock
        set master_count_ctno = isnull(d.master_count_ctno,0),
            master_count_skno = isnull(d.master_count_skno,0)
       from Cust_WeekSale_Stock m
            left join 
            (select Chg_sp_sales, cal_week_Range,
                    Count(Distinct ct_no8) as master_count_ctno,
                    Count(Distinct sd_skno) as master_count_skno
               from Cust_WeekSale_Stock
			  where Chg_sale_Month_Master = 'Y'
              group by Chg_sp_sales,cal_week_Range
            ) d on m.Chg_sp_sales = d.Chg_sp_sales and m.cal_week_Range=d.cal_week_Range
			

     update Cust_WeekSale_Stock
        set nonsale_count_ctno = isnull(d.nonsale_count_ctno,0),
            nonsale_count_skno = isnull(d.nonsale_count_skno,0)
       from Cust_WeekSale_Stock m
            left join 
            (select Chg_sp_sales, cal_week_Range,
                    Count(Distinct ct_no8) as nonsale_count_ctno,
                    Count(Distinct sd_skno) as nonsale_count_skno
               from Cust_WeekSale_Stock
			  where Chg_Stock_NonSales = 'Y'
              group by Chg_sp_sales,cal_week_Range
            ) d on m.Chg_sp_sales = d.Chg_sp_sales and m.cal_week_Range=d.cal_week_Range
			

     update Cust_WeekSale_Stock
        set newstock_count_ctno = isnull(d.newstock_count_ctno,0),
            newstock_count_skno = isnull(d.newstock_count_skno,0)
       from Cust_WeekSale_Stock m
            left join 
            (select Chg_sp_sales, cal_week_Range,
                    Count(Distinct ct_no8) as newstock_count_ctno,
                    Count(Distinct sd_skno) as newstock_count_skno
               from Cust_WeekSale_Stock
			  where chg_new_arrival_ym > ''
              group by Chg_sp_sales,cal_week_Range
            ) d on m.Chg_sp_sales = d.Chg_sp_sales and m.cal_week_Range=d.cal_week_Range
			

     update Cust_WeekSale_Stock
        set deadstock_count_ctno = isnull(d.deadstock_count_ctno,0),
            deadstock_count_skno = isnull(d.deadstock_count_skno,0)
       from Cust_WeekSale_Stock m
            left join 
            (select Chg_sp_sales, cal_week_Range,
                    Count(Distinct ct_no8) as deadstock_count_ctno,
                    Count(Distinct sd_skno) as deadstock_count_skno
               from Cust_WeekSale_Stock
			  where Chg_is_dead_stock = 'Y'
              group by Chg_sp_sales,cal_week_Range
            ) d on m.Chg_sp_sales = d.Chg_sp_sales and m.cal_week_Range=d.cal_week_Range
            
     select @Cnt = count(1) from Cust_WeekSale_Stock
     
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
