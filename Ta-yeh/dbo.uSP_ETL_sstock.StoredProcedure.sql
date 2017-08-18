USE [DW]
GO
/****** Object:  StoredProcedure [dbo].[uSP_ETL_sstock]    Script Date: 08/18/2017 17:18:56 ******/
DROP PROCEDURE [dbo].[uSP_ETL_sstock]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
Create Proc [dbo].[uSP_ETL_sstock]
as
begin
  /*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
   Program Name: uSP_ETL_sstock
   Create Date: 2013/03/01
   Creator: Rickliu
   Updated Date: 2013/09/06 [修改 Trans_Log 訊息內容]
  *-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*/  
  Declare @Proc Varchar(50) = 'uSP_ETL_sstock'
  Declare @Cnt Int =0
  Declare @Cnt_Ori Int =0
  Declare @Err_Code Int = -1
  Declare @Msg Varchar(Max) =''
  Declare @strSQL Varchar(Max)=''
  Declare @CR Varchar(4) = char(13)+char(10)
  Declare @Result Int = 0
  
  IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[Fact_sstock]') AND type in (N'U'))
  begin
     Set @Msg = '刪除資料表 [Fact_sstock]'
     set @strSQL= 'DROP TABLE [dbo].[Fact_sstock]'

     Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
  end

  begin try
    set @Msg = 'ETL SStock to [Fact_sstock]...'
    Exec uSP_Sys_Write_Log @Proc, @Msg, @Msg, 0

    -- 2017/08/02 Rickliu 滯銷品 SubQuery
    ;With CTE_Q1 as (
      -- 抓取下架商品且計算下架半年後的起迄日期
      select Rtrim(sk_no) as sk_no, Rtrim(sk_name) as sk_name, 
             convert(datetime, sk_color+'01') as Dead_BDate, 
             dateadd(day, -1, dateadd(month, +6, convert(datetime, sk_color+'01'))) as Dead_EDate
        from sync_ta13.dbo.sstock m
       where 1=1
         and isdate(sk_color+'01') = 1
    ), CTE_Q2 as (
      -- 抓取單據為 sd_class = ('[0-3]', '6', '8')且單據日期落於 CTE_Q1 的起迄日期
      -- 若有資料則代表下架後之半年內仍有交易，不列為滯銷品
      select Rtrim(sd_skno) as sd_skno, Rtrim(sd_name) as sd_name, 
             Dead_BDate, Dead_EDate, max(sd_date) as sd_date
        from sync_ta13.dbo.sslpdt m
             left join CTE_Q1 d
               on m.sd_skno = d.sk_no
       where 1=1
         and sd_class in ('[0-3]', '6', '8')
         and sd_date between d.Dead_BDate and d.Dead_EDate
       group by sd_skno, sd_name, Dead_BDate, Dead_EDate
    --  having max(sd_date) > bdate
    ), CTE_Q3 as (
      select distinct sk_no
        from Ori_xls#Master_Stock
    )
    
    select distinct 
           Rtrim(M.[sk_no]) as sk_no, 
           Rtrim(M.[sk_bcode]) as sk_bcode, 
           Rtrim(M.[sk_name]) as sk_name, M.[sk_kind], M.[sk_ivkd],
           Rtrim(M.[sk_spec]) as sk_spec, 
           Rtrim(M.[sk_unit]) as sk_unit, 
           Rtrim(M.[sk_color]) as sk_color, 
           Rtrim(M.[sk_use]) as sk_use, M.[sk_aux_q],
           M.[sk_aux_u], 
           Rtrim(M.[sk_size]) as sk_size, 
           M.[sk_save], M.[sk_bqty], 
           [sk_rem]=Rtrim(Convert(Varchar(255), M.[sk_rem])),
           M.[sk_fign],  
           -- 2017/07/12 Rickliu 為了 SmartQuery 報表能夠正常列印情況下，將原本 sk_fign 內所指的目錄改為 WEB 目錄方式
           REPLACE( REPLACE(m.sk_fign,'W:\5.產品圖','http://192.168.1.51/Smart-Query/pic'),'\','/') as sk_pic,
           
           M.[sk_tfg], 
           [st_cspes]=Rtrim(Convert(Varchar(255), M.[st_cspes])), 
           [st_espes]=Rtrim(Convert(Varchar(255), M.[st_espes])), M.[st_unit],
   
           M.[st_inerqty], M.[st_inerunt], M.[st_outqty], M.[st_outunt], 
           M.[st_appr], M.[st_apprunt], M.[st_nw], M.[st_gw], M.[st_gwunt], 
           M.[st_itclass], M.[st_cccode], M.[st_sespes],  M.[st_20cyqty], M.[st_40cyqty], 
           M.[st_45cyqty], 
          
           Rtrim(M.[sk_fld1]) as sk_fld1, 
           Rtrim(M.[sk_fld2]) as sk_fld2, 
           Rtrim(M.[sk_fld3]) as sk_fld3, 
           Rtrim(M.[sk_fld4]) as sk_fld4, 
           Rtrim(M.[sk_fld5]) as sk_fld5, 
           Rtrim(M.[sk_fld6]) as sk_fld6, 
           Rtrim(M.[sk_fld7]) as sk_fld7, 
           Rtrim(M.[sk_fld8]) as sk_fld8, 
           Rtrim(M.[sk_fld9]) as sk_fld9, 
           Rtrim(M.[sk_fld10]) as sk_fld10, 
           Rtrim(M.[sk_fld11]) as sk_fld11, 
           Rtrim(M.[sk_fld12]) as sk_fld12, 
         
           M.[sk_ltime], M.[sk_ikind], M.[s_supp], M.[s_locate], M.[s_price1], 
           M.[s_price2], M.[s_price3], M.[s_price4], M.[s_price5], M.[s_price6], 
           M.[s_aprice], M.[s_m_ave], M.[s_updat1], M.[s_updat2], M.[s_lprice1], 
           M.[s_lprice2], M.[s_accno1], M.[s_accno2], M.[s_accno3], M.[s_accno4], 
           M.[s_accno5], M.[s_accno6], M.[st_jname], [st_jspes]=Rtrim(Convert(Varchar(255), M.[st_jspes])), M.[sk_nowqty],
           M.[st_proven], M.[st_lenght], M.[st_width], M.[st_height], M.[st_lwhunit], 
           M.[st_unw], M.[st_ugw], M.[st_uappr], M.[st_uarea], M.[st_areaunt], 
           M.[st_freno], M.[st_fign2], M.[st_fign3], M.[st_fign4], M.[sk_poseq], 
           M.[sk_abcode], M.[s_apice2], M.[s_apice3], M.[s_apice4], M.[s_apice5],
    
           M.[s_apice6], 
           M.[sk_wfno], 
           Rtrim(M.[sk_whno]) as sk_whno, M.[sk_mdpfg], M.[sk_lotfg],
           -- [sk_no]
           Chg_skno_accno = isnull(substring(M.[sk_no], 1, 1), '00'),
           Chg_skno_accno_Name = isnull(D1.Code_Name, '未設定'),
           Chg_skno_BKind = isnull(substring(M.[sk_no], 1, 2), '00'),
           Chg_skno_Bkind_Name = isnull(D2.Code_Name, '未設定'),

           -- Rickliu 2014/09/29 協理要求將 AD, AE, AE 納入 AA 百貨計算
           Chg_skno_BKind2 = 
             Case 
               when substring(M.[sk_no], 1, 2) in ('AD', 'AE', 'AZ') then 'AA'
               else isnull(substring(M.[sk_no], 1, 2), '00')
             end,
           Chg_skno_Bkind_Name2 = 
             Case 
               when substring(M.[sk_no], 1, 2) in ('AD', 'AE', 'AZ') then 
                    isnull((select code_name as code_name
                              from Ori_XLS#Sys_stockcode
                             where code_level = 2
                               and code_no = 'AA'
                           ), '未設定')
               else isnull(D2.Code_Name, '未設定')
             end,
           /***********************************************************  
            Rickliu 2015/01/23 美如協理要求計算庫存時將進行以下分類
            商品類別  類別名稱  商品編號前兩碼
            AA        百貨      AA, BA, CA
            AB        3C        AB, BB, CB
            AC        化工      AC, BC, CC
            AD        工具      AD, BD, CD
            AE        電裝      AE, BE, CE
            AF        客訂      AF, BF, CF
            OT        其他      OT
           ***********************************************************/
           Chg_skno_BKind3 = 
             Case 
               when substring(M.[sk_no], 1, 2) in ('AA', 'BA', 'CA') then 'AA'
               when substring(M.[sk_no], 1, 2) in ('AB', 'BB', 'CB') then 'AB'
               when substring(M.[sk_no], 1, 2) in ('AC', 'BC', 'CC') then 'AC'
               when substring(M.[sk_no], 1, 2) in ('AD', 'BD', 'CD') then 'AD'
               when substring(M.[sk_no], 1, 2) in ('AE', 'BE', 'CE') then 'AE'
               when substring(M.[sk_no], 1, 2) in ('AF', 'BF', 'CF') then 'AF'
               when substring(M.[sk_no], 1, 1) in ('E') then 'OT'
               else ''
             end,
           Chg_skno_Bkind_Name3 = 
             Case 
               when substring(M.[sk_no], 1, 2) in ('AA', 'BA', 'CA') then '百貨'
               when substring(M.[sk_no], 1, 2) in ('AB', 'BB', 'CB') then '3C'
               when substring(M.[sk_no], 1, 2) in ('AC', 'BC', 'CC') then '化工'
               when substring(M.[sk_no], 1, 2) in ('AD', 'BD', 'CD') then '工具'
               when substring(M.[sk_no], 1, 2) in ('AE', 'BE', 'CE') then '電裝'
               when substring(M.[sk_no], 1, 2) in ('AF', 'BF', 'CF') then '客訂'
               when substring(M.[sk_no], 1, 1) in ('E') then '其他'
               else ''
             end,
             
           Chg_skno_SKind = substring(M.[sk_no], 1, 4),
           Chg_skno_SKind_Name = D3.Code_Name,
           Chg_kind_Name = D4.kd_name,
           Chg_ivkd_Name = D5.kd_name,
           Chg_StartYear = substring(Replace(M.[sk_size], '無',''), 1, 4),
           Chg_StartMonth = substring(Replace(M.[sk_size], '無',''), 5, 2),
           Chg_EndYear = substring(Replace(M.[sk_color], '無',''), 1, 4),
           Chg_EndMonth = substring(Replace(M.[sk_color], '無',''), 5, 2),
           Chg_supp_Name = D6.ct_name,
           Chg_supp_SName = D6.ct_sname,
           Chg_locate_MArea = substring(M.[s_locate], 1, 1),
           Chg_locate_DArea = substring(M.[s_locate], 1, 2),
           Chg_locate_row = substring(M.[s_locate], 4, 1),
           Chg_locate_col = substring(M.[s_locate], 6, 1),
           Chg_updat1_Year = DATEPART(yy, M.s_updat1),
           Chg_updat1_Month = DATEPART(mm, M.s_updat1),
           Chg_updat2_Year = DATEPART(yy, M.s_updat2),
           Chg_updat2_Month = DATEPART(mm, M.s_updat2),
           -- 2014/06/07 未銷商品註記
           -- 2017/08/07 Rickliu 取消使用商品未銷列表
           --Chg_Stock_NonSales = 
           --  Case 
           --    when D7.sk_no is Not null Then 'Y'
           --    else 'N'
           --  End,
           -- 2014/06/17 Rick 新增 AA 期初安全存量
           --Chg_WD_AA_first_sQty = Isnull(D8.wd_first_sqty, 0),
           Chg_WD_AA_sQty = Isnull(D8.wd_sqty, 0),
           -- 2014/06/17 Rick 新增 AA 期初現有庫存數
           Chg_WD_AA_first_Qty = Isnull(D8.wd_first_qty, 0),           
           -- 2014/06/17 Rick 新增 AA 期初庫存差異數
           Chg_WD_AA_first_diff_Qty = Isnull(D8.wd_first_qty_diff, 0),
           -- 2014/06/07 Rick 新增 AA 期末現有庫存數
           Chg_WD_AA_last_Qty = Isnull(D8.wd_last_qty, 0),
           -- 2014/06/17 Rick 新增 AA 期末庫存差異數
           Chg_WD_AA_last_diff_Qty = Isnull(D8.wd_last_qty_diff, 0),
           
           -- 2014/06/17 Rick 新增 AB 期初安全存量
           --Chg_WD_AB_first_sQty = Isnull(D9.wd_first_sqty, 0),
           Chg_WD_AB_sQty = Isnull(D9.wd_sqty, 0),
           -- 2014/06/17 Rick 新增 AB 期初現有庫存數
           Chg_WD_AB_first_Qty = Isnull(D9.wd_first_qty, 0),           
           -- 2014/06/17 Rick 新增 AB 期初庫存差異數
           Chg_WD_AB_first_diff_Qty = Isnull(D9.wd_first_qty_diff, 0),
           -- 2014/06/07 Rick 新增 AB 期末現有庫存數
           Chg_WD_AB_last_Qty = Isnull(D9.wd_last_qty, 0),
           -- 2014/06/17 Rick 新增 AB 期末庫存差異數
           Chg_WD_AB_last_diff_Qty = Isnull(D9.wd_last_qty_diff, 0),

           -- 2014/06/17 Rick 新增 AC 期初安全存量
           --Chg_WD_AB_first_sQty = Isnull(D10.wd_first_sqty, 0),
           Chg_WD_AC_sQty = Isnull(D10.wd_sqty, 0),
           -- 2014/06/17 Rick 新增 AC 期初現有庫存數
           Chg_WD_AC_first_Qty = Isnull(D10.wd_first_qty, 0),           
           -- 2014/06/17 Rick 新增 AC 期初庫存差異數
           Chg_WD_AC_first_diff_Qty = Isnull(D10.wd_first_qty_diff, 0),
           -- 2014/06/07 Rick 新增 AC 期末現有庫存數
           Chg_WD_AC_last_Qty = Isnull(D10.wd_last_qty, 0),
           -- 2014/06/17 Rick 新增 AC 期末庫存差異數
           Chg_WD_AC_last_diff_Qty = Isnull(D10.wd_last_qty_diff, 0),

           -- 2014/06/17 Rick 新增 AG 期初安全存量
           Chg_WD_AG_sQty = Isnull(D11.wd_sqty, 0),
           --Chg_WD_AG_first_sQty = Isnull(D10.wd_first_sqty, 0),
           -- 2014/06/17 Rick 新增 AG 期初現有庫存數
           Chg_WD_AG_first_Qty = Isnull(D11.wd_first_qty, 0),           
           -- 2014/06/17 Rick 新增 AG 期初庫存差異數
           Chg_WD_AG_first_diff_Qty = Isnull(D11.wd_first_qty_diff, 0),
           -- 2014/06/07 Rick 新增 AG 期末現有庫存數
           Chg_WD_AG_last_Qty = Isnull(D11.wd_last_qty, 0),
           -- 2014/06/17 Rick 新增 AG 期末庫存差異數
           Chg_WD_AG_last_diff_Qty = Isnull(D11.wd_last_qty_diff, 0),

           -- 2014/06/16 Rickliu 新增滯銷品
           Chg_IS_Dead_Stock = Case when D12.sk_no is not null then 'Y' else 'N' end,
           -- 2017/08/02 Rickliu 修訂滯銷品定義，改為以下架日後之半年內無任何交易者(sd_class in ('[0-3]', '6', '8'))
           --Chg_IS_Dead_Stock = 
           --  Case
           --    when D15.sd_skno is null then 'N'
           --    else 'Y'
           --  end,

           Chg_Dead_First_Qty = Round(isnull(D12.First_Qty, 0), 0),
           Chg_Dead_First_Amt = Round(isnull(D12.First_Amt, 0), 0),
           Chg_Dead_Stock_YM = Convert(Varchar(7), Isnull(D12.Dead_YM, ''), 111),
           
           -- 2017/08/01 Rickliu 修訂新品定義，改為以一年內之引進之產品為新品
           Chg_IS_New_Stock = Case when Convert(DateTime, D13.Arrival_Date) >= dateadd(year, -1, getdate()) then 'Y' else 'N' end,
           -- 2014/07/02 Rickliu 新增採購新品資訊
           Chg_New_sm_no = rtrim(D13.sm_no),
           Chg_New_Arrival_Date = Convert(DateTime, D13.Arrival_Date),
           Chg_New_Arrival_YM = Convert(Varchar(7), Convert(DateTime, D13.Arrival_Date), 111),
           Chg_New_First_Qty = Round(isnull(D13.First_Qty, 0), 0),
           -- 2017/08/01 Rickliu 重新定義主銷，取消原本在單據明細內判別是否為主銷
           Chg_IS_Master_Stock = case when rtrim(m.sk_no) = rtrim(D16.sk_no) collate Chinese_Taiwan_Stroke_CI_AS then 'Y' else 'N' end,
		   -- 2015/05/15 Nanliao 新增產品屬性表資訊
		                            
		   Chg_color            = D14.color           ,
		   Chg_size             = D14.size            ,
		   Chg_package          = D14.package         ,
		   Chg_barcode_name     = D14.barcode_name    ,
		   Chg_price            = D14.price           ,
		   Chg_pic_6            = D14.pic_6           ,
		   Chg_pic_4            = D14.pic_4           ,
		   Chg_pic_2            = D14.pic_2           ,
		   Chg_main_pic         = D14.main_pic        ,
		   Chg_pic1             = D14.pic1            ,
		   Chg_pic2             = D14.pic2            ,
		   Chg_pic3             = D14.pic3            ,
		   Chg_avg_price        = D14.avg_price       ,
		   Chg_product_property = D14.product_property,
		   Chg_gross_property   = D14.gross_property  ,

           case when rtrim(m.sk_no) = rtrim(D16.sk_no) collate Chinese_Taiwan_Stroke_CI_AS then '[主]' else '' end+
           Case when Convert(DateTime, D13.Arrival_Date) >= dateadd(year, -1, getdate()) then '[新]' else '' end+
           Case when D12.sk_no is not null then '[滯]' else '' end as stock_kind_list,
           sstock_update_datetime = getdate(),
           sstock_Timestamp = m.Timestamp_Column
           into Fact_sstock
      from SYNC_TA13.dbo.sstock M With(NoLock) 
           left join Ori_XLS#Sys_stockcode as D1 With(NoLock) 
                  ON D1.Code_Level = '1' AND substring([sk_no], 1, 1) Collate Chinese_Taiwan_Stroke_CI_AS = D1.Code_No Collate Chinese_Taiwan_Stroke_CI_AS
           left join Ori_XLS#Sys_stockcode as D2 With(NoLock) 
                  ON D2.Code_Level = '2' AND substring([sk_no], 1, 2) Collate Chinese_Taiwan_Stroke_CI_AS = D2.Code_No Collate Chinese_Taiwan_Stroke_CI_AS
           left join Ori_XLS#Sys_stockcode as D3 With(NoLock) 
                  ON D3.Code_Level = '3' AND substring([sk_no], 1, 4) Collate Chinese_Taiwan_Stroke_CI_AS = D3.Code_No Collate Chinese_Taiwan_Stroke_CI_AS
           left join SYNC_TA13.dbo.skind as D4 With(NoLock) 
                  ON D4.kd_class='1' AND M.sk_kind = D4.kd_no 
           left join SYNC_TA13.dbo.skind as D5 With(NoLock) 
                  ON D5.kd_class='2' AND M.sk_ivkd = D5.kd_no 
           left join SYNC_TA13.dbo.pcust as D6 With(NoLock) 
                  ON D6.ct_class='2' AND M.s_supp = D6.ct_no
           -- 2017/08/07 Rickliu 取消使用商品未銷列表
           --left join Ori_xls#Stock_NonSales as D7 With(NoLock) 
           --       ON M.sk_no Collate Chinese_Taiwan_Stroke_CI_AS = D7.sk_no Collate Chinese_Taiwan_Stroke_CI_AS
           left join SYNC_TA13.dbo.V_Last_Swaredt as D8 With(NoLock) 
                  On M.sk_no = D8.sk_no
                 And D8.Wd_no ='AA'
           left join SYNC_TA13.dbo.V_Last_Swaredt as D9 With(NoLock) 
                  On M.sk_no = D9.sk_no
                 And D9.Wd_no ='AB'
           left join SYNC_TA13.dbo.V_Last_Swaredt as D10 With(NoLock) 
                  On M.sk_no = D10.sk_no
                 And D9.Wd_no ='AC'
           left join SYNC_TA13.dbo.V_Last_Swaredt as D11 With(NoLock) 
                  On M.sk_no = D11.sk_no
                 And D11.Wd_no ='AG'
           left join Ori_xls#Dead_Stock_Lists as D12 With(NoLock) 
                  on M.sk_no Collate Chinese_Taiwan_Stroke_CI_AS = D12.sk_no Collate Chinese_Taiwan_Stroke_CI_AS
           left join Ori_xls#New_Stock_Lists as D13 With(NoLock) 
                  on M.sk_no Collate Chinese_Taiwan_Stroke_CI_AS = D13.sk_no Collate Chinese_Taiwan_Stroke_CI_AS and D13.sk_no > ''
		   -- 2015/05/15 Nanliao 新增產品屬性表資訊
           left join Ori_xls#Stock_Property as D14 With(NoLock) 
                  on M.sk_no Collate Chinese_Taiwan_Stroke_CI_AS = D14.sk_no Collate Chinese_Taiwan_Stroke_CI_AS and D14.kind = 'P'
           left join CTE_Q2 as D15 on m.sk_no = D15.sd_skno
           -- 2017/08/07 Rickliu 重新修訂主銷列表
           left join CTE_Q3 as D16 on m.sk_no = D16.sk_no collate Chinese_Taiwan_Stroke_CI_AS
           

    /*==============================================================*/
    /* Index: sstock_Timestamp                                      */
    /*==============================================================*/
    --Set @Msg = '建立索引 [Fact_sstock.sstock_Timestamp]'
    --Print @Msg
    --if exists (select * from sys.indexes where object_id = object_id('[dbo].[fact_sstock]') and name = 'pk_sstock')
    --   alter table [dbo].[pk_sstock] drop constraint [pk_sstock]

    --alter table [dbo].[fact_sstock] add  constraint [pk_sstock] primary key nonclustered ([sstock_timestamp] asc) with 
    --(pad_index  = off, statistics_norecompute  = off, sort_in_tempdb = off, ignore_dup_key = off, online = off, allow_row_locks  = on, allow_page_locks  = on) on [primary]

    /*==============================================================*/
    /* Index: SKNO                                                  */
    /*==============================================================*/
    Set @Msg = '建立索引 [Fact_sstock.SKNO]'
    Print @Msg
    if exists (select 1 from sysindexes where id = object_id('dbo.fact_sstock') and name  = 'SKNO' and indid > 0 and indid < 255) 
       drop index dbo.fact_sstock.SKNO
    create clustered index SKNO on dbo.fact_sstock (sk_no) with fillfactor= 30 on "PRIMARY"

    /*==============================================================*/
    /* Index: SKABNO                                                */
    /*==============================================================*/
    Set @Msg = '建立索引 [Fact_sstock.SKABNO]'
    Print @Msg
    if exists (select 1 from sysindexes where id = object_id('dbo.fact_sstock') and name  = 'SKABNO' and indid > 0 and indid < 255) drop index dbo.fact_sstock.SKABNO
       create index SKABNO on dbo.fact_sstock (sk_abcode, sk_no) with fillfactor= 30 on "PRIMARY"

    /*==============================================================*/
    /* Index: SKBNO                                                 */
    /*==============================================================*/
    Set @Msg = '建立索引 [Fact_sstock.SKBNO]'
    Print @Msg
    if exists (select 1 from sysindexes where id = object_id('dbo.fact_sstock') and name  = 'SKBNO' and indid > 0 and indid < 255) drop index dbo.fact_sstock.SKBNO
       create index SKBNO on dbo.fact_sstock (sk_bcode, sk_no) with fillfactor= 30 on "PRIMARY"

    /*==============================================================*/
    /* Index: SKKD                                                  */
    /*==============================================================*/
    Set @Msg = '建立索引 [Fact_sstock.SKKD]'
    Print @Msg
    if exists (select 1 from sysindexes where id = object_id('dbo.fact_sstock') and name  = 'SKKD' and indid > 0 and indid < 255) drop index dbo.fact_sstock.SKKD
       create index SKKD on dbo.fact_sstock (sk_kind, sk_no) with fillfactor= 30 on "PRIMARY"

    /*==============================================================*/
    /* Index: SKNM                                                  */
    /*==============================================================*/
    Set @Msg = '建立索引 [Fact_sstock.SKNM]'
    Print @Msg
    if exists (select 1 from sysindexes where id = object_id('dbo.fact_sstock') and name  = 'SKNM' and indid > 0 and indid < 255) drop index dbo.fact_sstock.SKNM
       create index SKNM on dbo.fact_sstock (sk_name) with fillfactor= 30 on "PRIMARY"
    
    /*==============================================================*/
    /* Index: SUPP                                                  */
    /*==============================================================*/
    Set @Msg = '建立索引 [Fact_sstock.SUPP]'
    Print @Msg
    if exists (select 1 from sysindexes where id = object_id('dbo.fact_sstock') and name  = 'SUPP' and indid > 0 and indid < 255) drop index dbo.fact_sstock.SUPP
       create index SUPP on dbo.fact_sstock (s_supp) with fillfactor= 30 on "PRIMARY"
  end try
  begin catch
    set @Result = @Err_Code
    set @Msg = @Proc+'...(錯誤訊息:'+ERROR_MESSAGE()+', '+@Msg+')...(錯誤列:'+Convert(Varchar(10), ERROR_LINE())+')'

    Exec uSP_Sys_Write_Log @Proc, @Msg, @Msg, @Cnt
  end catch
  Return(@Cnt)
end
GO
