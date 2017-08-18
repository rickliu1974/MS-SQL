USE [DW]
GO
/****** Object:  StoredProcedure [dbo].[xuSP_ETL_Hunderd_NonSale_NewStock(201708停用)]    Script Date: 08/18/2017 17:18:56 ******/
DROP PROCEDURE [dbo].[xuSP_ETL_Hunderd_NonSale_NewStock(201708停用)]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE Proc [dbo].[xuSP_ETL_Hunderd_NonSale_NewStock(201708停用)]
as
begin
  /*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
   Program Name: uSP_ETL_Hunderd_NonSale_NewStock
   Create Date: 2014/06/05
   Creator: Rickliu
  *-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*/  

  Declare @Proc Varchar(50) = 'uSP_ETL_Hunderd_NonSale_NewStock'
  Declare @Cnt Int =0
  Declare @Msg Varchar(Max) =''
  Declare @strSQL Varchar(Max)=''
  Declare @CR Varchar(4) = char(13)+char(10)
  Declare @Result int = 0
  
  IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[Hunderd_NonSale_NewStock]') AND type in (N'U'))
  begin
     Set @Msg = '刪除資料表 [Hunderd_NonSale_NewStock]'
     set @strSQL= 'DROP TABLE [dbo].[Hunderd_NonSale_NewStock]'

     Exec @Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
     if @Result <> -1 set @Result = 0
  end

  set @Msg = 'Create to [Hunderd_NonSale_NewStock]...'
  begin try
    
    
    select *,
           update_datetime = getdate()
           into Hunderd_NonSale_NewStock
      from (select '1' as kind, '已銷' as kind_Name, 
                   t.ct_no8, t.ct_sname8+
                   Case
                     when chg_ct_close = 'Y' then '(關店)'
                     else ''
                   end as ct_sname8, chg_ct_fld3, -- 2017/02/08 Rickliu Add 增加八碼客戶編號
                   Chg_Hunderd_Customer_Name, Chg_IS_Lan_Custom,
                   sd_skno, sd_name, 
                   Chg_Wd_AA_first_Qty, Chg_Wd_AB_first_Qty,
                   Chg_Wd_AA_last_Qty, Chg_Wd_AB_last_Qty,
                   t.Chg_skno_accno, t.Chg_skno_accno_Name, 
                   t.Chg_skno_BKind, t.Chg_skno_Bkind_Name, 
                   t.Chg_skno_SKind, t.Chg_skno_SKind_Name, 
                   t.Chg_kind_Name, Chg_sp_date_YM,
                   Chg_sp_sales, Chg_sales_Name, 
                   -- 新品到貨日、新品期初數量
                   Convert(Varchar(7), Convert(DateTime, Chg_New_Arrival_Date), 111) as Chg_New_Arrival_YM, Chg_New_First_Qty, 
                   Isnull(SUM(sd_qty), 0) as qty, 
                   isnull(sum(sd_stot), 0) as amt
              from Fact_sslpdt t 
             where 1=1
               and Chg_Hunderd_Customer ='Y'
               and Chg_IS_New_Stock ='Y'
               and sd_slip_fg in ('2', '6') -- 銷貨、託售皆算    
               and exists 
                   (select CT_NO8, SK_NO 
                      from Fact_sstock m 
                           Cross Join 
                           (select distinct ct_no8 
                              from Fact_pcust 
                             where CT_CLASS ='1' 
                               and len(ct_no) >= 8 
                               and Chg_Hunderd_Customer = 'Y'
                               --and chg_ct_close = ''
                           )d 
                     where 1=1
                       and Chg_New_Arrival_YM IS Not Null
                       and exists 
                           (select ct_no8, sd_skno 
                              from Fact_sslpdt d1 
                             where sd_slip_fg in ('2', '6') -- 銷貨、託售皆算
                               and m.SK_NO collate Chinese_Taiwan_Stroke_CI_AS = d1.sd_skno collate Chinese_Taiwan_Stroke_CI_AS
                               and d.CT_NO8 collate Chinese_Taiwan_Stroke_CI_AS = d1.ct_no8 collate Chinese_Taiwan_Stroke_CI_AS
                               and Chg_Hunderd_Customer ='Y'
                               and Chg_IS_New_Stock ='Y'
                            ) 
                       and t.ct_no8 collate Chinese_Taiwan_Stroke_CI_AS = CT_NO8 collate Chinese_Taiwan_Stroke_CI_AS
                       and t.sd_skno collate Chinese_Taiwan_Stroke_CI_AS = sk_no collate Chinese_Taiwan_Stroke_CI_AS
                   ) 
             group by Chg_sp_date_YM, 
                      ct_no8, ct_sname8+
                      Case
                        when chg_ct_close = 'Y' then '(關店)'
                        else ''
                      end, chg_ct_fld3, -- 2017/02/08 Rickliu Add 增加八碼客戶編號
                      Chg_Hunderd_Customer_Name, Chg_IS_Lan_Custom,
                      sd_skno, sd_name, 
                      Chg_Wd_AA_first_Qty, Chg_Wd_AB_first_Qty,
                      Chg_Wd_AA_last_Qty, Chg_Wd_AB_last_Qty,
                      t.Chg_skno_accno, t.Chg_skno_accno_Name, 
                      t.Chg_skno_BKind, t.Chg_skno_Bkind_Name, 
                      t.Chg_skno_SKind, t.Chg_skno_SKind_Name, 
                      t.Chg_kind_Name, Chg_sp_sales, Chg_sales_Name, 
                      Convert(Varchar(7), Convert(DateTime, Chg_New_Arrival_Date), 111), Chg_New_First_Qty
            union 
            select '2' as kind, '未銷' as kind_name, 
                   ct_no8, ct_sname8+
                   Case
                     when chg_ct_close = 'Y' then '(關店)'
                     else ''
                   end as ct_sname8, Chg_ct_fld3,  -- 2017/02/08 Rickliu Add 增加八碼客戶編號
                   Chg_Hunderd_Customer_Name, Chg_IS_Lan_Custom,
                   m.sd_skno, m.sk_name,
                   Chg_Wd_AA_first_Qty, Chg_Wd_AB_first_Qty,
                   Chg_Wd_AA_last_Qty, Chg_Wd_AB_last_Qty,
                   Chg_skno_accno, Chg_skno_accno_Name, 
                   Chg_skno_BKind, Chg_skno_Bkind_Name, 
                   Chg_skno_SKind, Chg_skno_SKind_Name, 
                   Chg_kind_Name, Chg_sp_date_YM,
                   'NA' as Chg_sp_sales, 'NA' as Chg_sales_Name, 
                   Chg_New_Arrival_YM, Chg_New_First_Qty,
                   0 as qty, 0 as amt 
              from (select sk_no as sd_skno, sk_name, 
                           Chg_Wd_AA_first_Qty, Chg_Wd_AB_first_Qty,
                           Chg_Wd_AA_last_Qty, Chg_Wd_AB_last_Qty,
                           Chg_skno_accno, Chg_skno_accno_Name, 
                           Chg_skno_BKind, Chg_skno_Bkind_Name, 
                           Chg_skno_SKind, Chg_skno_SKind_Name, 
                           Chg_kind_Name, 
                           -- 新品年月、新品期初
                           Chg_New_Arrival_YM, Chg_New_First_Qty
                      from Fact_sstock m
                     where Chg_New_Arrival_YM Is Not Null
                    ) m
                   Cross Join 
                   (select ct_no8, ct_sname8, -- 2017/02/08 Rickliu Add 增加八碼客戶編號
                           ct_fld3 as Chg_ct_fld3, Chg_Hunderd_Customer_Name, Chg_IS_Lan_Custom, chg_ct_close
                      from Fact_pcust 
                     where CT_CLASS ='1' 
                       and len(ct_no) >= 8 
                       and Chg_Hunderd_Customer ='Y'
                   ) D
                   Cross Join 
                   (select Chg_sp_date_YM
                      from (select distinct CONVERT(varchar(7), cal_date, 111) as Chg_sp_date_YM
                              from Calendar
                             where CONVERT(varchar(7), cal_date, 111) <=CONVERT(varchar(7), GETDATE(), 111)
                               and CONVERT(varchar(7), cal_date, 111) >='2012/01') M
                     where 1=1
                   ) D1
                   left join Ori_XLS#New_Stock_Lists D2
                     ON m.sd_skno collate Chinese_Taiwan_Stroke_CI_AS = D2.sk_no collate Chinese_Taiwan_Stroke_CI_AS
                    and D1.Chg_sp_date_YM = D2.Arrival_Date
             where not exists
                   (select *
                      from Fact_sslpdt
                     where sd_slip_fg in ('2', '6') -- 銷貨、託售皆算
                       and Chg_sp_date_YM=D1.Chg_sp_date_YM
                       and ct_no8 collate Chinese_Taiwan_Stroke_CI_AS = d.ct_no8 collate Chinese_Taiwan_Stroke_CI_AS
                       and sd_skno collate Chinese_Taiwan_Stroke_CI_AS = m.sd_skno collate Chinese_Taiwan_Stroke_CI_AS
                       and Chg_Hunderd_Customer ='Y'
                       and Chg_Stock_NonSales ='Y'
                    )
           )m 
    where 1=1              

    select @Result = count(1) from Hunderd_NonSale_NewStock

    Exec uSP_Sys_Write_Log @Proc, @Msg, '', @Result
  end try
  begin catch
    set @Result = -1
    set @Msg = @Msg+'(錯誤訊息:'+ERROR_MESSAGE()+')'

    Exec uSP_Sys_Write_Log @Proc, @Msg, '', @Result
  end catch
  Return(@Result)
end
GO
