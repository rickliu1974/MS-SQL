USE [DW]
GO
/****** Object:  StoredProcedure [dbo].[uSP_ETL_sslpdt]    Script Date: 08/18/2017 17:43:39 ******/
DROP PROCEDURE [dbo].[uSP_ETL_sslpdt]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE Proc [dbo].[uSP_ETL_sslpdt]
as
begin
  /*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
   Program Name: uSP_ETL_sslpdt
   Create Date: 2013/03/01
   Creator: Rickliu
   Updated Date: 2013/09/06 [修改 Trans_Log 訊息內容]
  *-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*/  
  Declare @Proc Varchar(50) = 'uSP_ETL_sslpdt'
  Declare @Cnt Int =0
  Declare @Cnt_Ori Int =0
  Declare @Master_Stock_Cnt Int = 0
  Declare @Msg Varchar(Max) =''
  Declare @strSQL Varchar(Max)=''
  Declare @CR Varchar(4) = char(13)+char(10)
  Declare @slip_fg_Add Varchar(50)
  Declare @slip_fg_Less Varchar(50)
  
  Set @slip_fg_Add  = (select code_end from Ori_Xls#Sys_Code where code_class = '90' and code_begin ='+')
  Set @slip_fg_Less = (select code_end from Ori_Xls#Sys_Code where code_class = '90' and code_begin ='-')
  
  IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[Fact_sslpdt]') AND type in (N'U'))
  begin
     Set @Msg = '刪除資料表 [Fact_sslpdt]'
     set @strSQL= 'DROP TABLE [dbo].[Fact_sslpdt]'

     Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
  end

  
  set @Msg = 'ETL sslpdt to [Fact_sslpdt]...'
  begin try
    -- 2017/04/05 Rickliu 不知道為何，明細資料處理速度變得很慢，研判應該是沒有做所謂的非叢集 Index，另外改用 CTE 撰寫方式，看是否能加快速度。
    ;with CTE_Q1 as (
      Select sd_class,
             sd_slip_fg,
             sd_ctno,
             Max(sd_no) AS Max_SP_NO, 
             Max(CONVERT(varchar(10), sd_date, 111)) AS sp_date, 
             min(sd_price) as sd_price,
             sd_skno,
             'Y' as Flag -- Rickliu 2017/06/06  這裡是指最後一次交易，但不包含贈品且金額小於 1元者
        From SYNC_TA13.dbo.sslpdt With(NoLock)
       where 1=1
         and sd_name is not null
         and sd_date >= '2012/12/01'
         and sd_slip_fg = '2'
         and isnull(sd_csno, '') <> ''
         and sd_sendfg = 0
         and sd_price > 1
         and sd_sendfg = 0
       Group by sd_class, sd_slip_fg, sd_ctno, sd_skno
    ), CTE_Q2 as (
      Select sd_class,
             sd_slip_fg,
             --2016/09/09 Rickliu 因組織擴編，增加機車通路，因此將原本 2~5 改為 2~6
             substring(sd_ctno, 1, 6) as BU_NO,
             Max(sd_no) AS Max_SP_NO, 
             Max(CONVERT(varchar(10), sd_date, 111)) AS sp_date, 
             min(sd_price) as sd_price,
             sd_skno,
             'Y' as Flag -- Rickliu 2017/06/06  這裡是指最後一次交易，但不包含贈品且金額小於 1元者
        From SYNC_TA13.dbo.sslpdt With(NoLock) 
       where 1=1
         and sd_name is not null
         and sd_date >= '2012/12/01'
         and sd_slip_fg = '2'
             --2016/09/09 Rickliu 因組織擴編，增加機車通路，因此將原本 2~5 改為 2~6
         and isnull(sd_csno, '') <> ''
         and sd_sendfg = 0
         and sd_price > 1
         and sd_sendfg = 0
       Group by sd_class, sd_slip_fg, substring(sd_ctno, 1, 6), sd_skno
    )

    select m.[sd_class], m.[sd_slip_fg], m.[sd_date], d9.sp_pdate, 
           Rtrim(m.[sd_no]) as sd_no, 
           Rtrim(m.[sd_ctno]) as sd_ctno, 
           Rtrim(m.[sd_skno]) as sd_skno,
           Rtrim(D4.[sk_name]) as sd_name, m.[sd_whno], m.[sd_whno2], m.[sd_qty], m.[sd_price], m.[sd_dis],
           m.[sd_stot], m.[sd_spec], [sd_rem]=Convert(Varchar(255), m.[sd_rem]), m.[sd_unit], m.[sd_unit_fg], m.[sd_ave_p],
           m.[sd_pave_p], m.[sd_rate], m.[sd_val1], m.[sd_val2], m.[sd_rqty], m.[sd_nqty],
           m.[sd_bmno], m.[sd_nokind], m.[sd_sekind], m.[sd_ordno], m.[sd_postfg], m.[sd_acspseq],
           m.[sd_csno], m.[sd_csrec], m.[sd_sendfg], m.[sd_mafkd], m.[sd_mave_p], m.[sd_mstot],
           m.[sd_adjkd], m.[sd_surefg], m.[sd_seqfld], m.[sd_lotno],

           rtrim(sp_sales) as Chg_sp_sales, -- Rickliu 2014/3/17 新增業務員
           rtrim(Chg_sales_Name) as Chg_sales_Name, -- Rickliu 2014/5/14 新增業務員名稱

           rtrim(ct_sales) as ct_sales, -- Rickliu 2014/07/03 新增客戶所屬業務員
           rtrim(Chg_ct_sales_Name) as Chg_ct_sales_Name, -- Rickliu 2014/07/03 新增客戶所屬業務員名稱
           rtrim(ct_name) as Chg_ct_name, --Rickliu 2014/1/25 新增 公司名稱
           rtrim(ct_sname) as Chg_ct_sname,--Rickliu 2014/1/25 新增 公司簡稱

           [Chg_sd_whno_name]=Rtrim(Isnull(D2.wh_name,  '')),
           [Chg_sd_class]=Rtrim(D5.Code_Name),
           [Chg_sd_slip_fg]=Rtrim(D6.Code_Name), 
           /***[單據種類]******************************************************************************
            =0進貨單(+)     =1進退單(-)     =2出貨單(+)     =3出退單(-)     =4借貨單(-)     =5還貨單(+)
            =6託售單(+)     =7託回單(-)     =8入庫單(+)     =9出庫單(-)     =A調整單        =B調撥單
            =C服務單(+)     =R借入單(+)     =S借還單        =T盤點單
           *******************************************************************************************/
           [Chg_sd_qty]= 
             Case 
               WHEN m.SD_SLIP_FG like '[02568ABCRST]' THEN IsNull(m.SD_QTY, 0)
               WHEN m.SD_SLIP_FG like '[13479]' THEN Isnull(-m.SD_QTY, 0)
               else 0
             END, 
           [Chg_sd_sale_qty]= 
             Case 
               WHEN m.SD_SLIP_FG like '[02568ABCRST]' THEN IsNull(m.SD_QTY, 0)
               else 0
             END, 
           [Chg_sd_ret_qty]= 
             Case 
               WHEN m.SD_SLIP_FG like '[13479]' THEN Isnull(-m.SD_QTY, 0)
               else 0
             END, 
           [Chg_sd_stot]= 
             Case 
               WHEN m.SD_SLIP_FG like '[02568ABCRST]' THEN Isnull(m.SD_STOT, 0)
               WHEN m.SD_SLIP_FG like '[13479]' THEN Isnull(-m.SD_STOT, 0)
               else 0
             END * isnull(d9.sp_rate, 1),
           [Chg_sd_ave_p]=
             Case 
               WHEN m.SD_SLIP_FG like '[02568ABCRST]' THEN Isnull(m.sd_ave_p, 0) 
               WHEN m.SD_SLIP_FG like '[13479]' THEN Isnull(-m.sd_ave_p, 0)
               else 0
             END * isnull(d9.sp_rate, 1),
           
           
           -- 是否為銷售日主銷商品
           -- 2017/08/07 Rickliu 改以商品基本資料取得是否為主銷
           --[Chg_sale_Month_Master] = 
           --   Case 
           --     when D10.SK_NO is not Null and D10.Sale_Year= D9.Chg_sp_date_year and D10.Sale_Month=D9.Chg_sp_date_month
           --      Then 'Y'
           --     else 'N'
           --   end,
           -- 2017/08/07 Rickliu 改以商品基本資料取得是否為主銷
           --[Chg_sale_Month_MasterName] =
           --   Case 
           --     when D10.SK_NO is not Null and D10.Sale_Year= D9.Chg_sp_date_year and D10.Sale_Month=D9.Chg_sp_date_month
           --       Then D10.sk_Mname
           --     else 'NA'
           --   end,
           -- 2017/08/07 Rickliu 改以商品基本資料取得是否為主銷
           -- 是否為請款日主銷商品
           --[Chg_inv_Month_Master] = -- 2013/6/1 新增銷售主銷名稱
           --   Case 
           --     when D8.SK_NO is not Null and D8.Sale_Year= D9.Chg_sp_pdate_year and D8.Sale_Month=D9.Chg_sp_pdate_month
           --       Then 'Y'
           --     else 'N'
           --   end,
           -- 2017/08/07 Rickliu 改以商品基本資料取得是否為主銷
           --[Chg_inv_Month_MasterName] = -- 2013/6/1 新增請款主銷名稱
           --   Case 
           --     when D8.SK_NO is not Null and D8.Sale_Year= D9.Chg_sp_pdate_year and D8.Sale_Month=D9.Chg_sp_pdate_month
           --       Then D8.sk_Mname
           --     else 'NA'
           --   end,
           -- 2013/1/17 協理與副總確認，當貨品單價高於等於中盤價時才認列業績(此部分確認有含小盤價的業績)，日後的業績報表必須依此呈現。
           -- 2013/1/18 協理與財務林課確認，所有銷單都是以單一品項折扣，所以只要拿主檔折扣進行計算即可。
           [Chg_sd_sale_overmid_tot]= --(此金額為未稅金額)
             Case 
               WHEN m.SD_SLIP_FG like '[02568ABCRST]' THEN 
                 Case 
                   when m.[sd_price] >= [s_apice4] then Isnull(m.SD_STOT, 0)
                   else 0
                 end 
               WHEN m.SD_SLIP_FG like '[13479]' THEN 
                 Case 
                   when m.[sd_price] >= [s_apice4] then Isnull(-m.SD_STOT, 0)
                   else 0
                 end
               else 0
             END * isnull(d9.sp_rate, 1),
           [Chg_sd_sale_overmid_dis]= --(此金額為未稅金額)
             Case 
               WHEN m.SD_SLIP_FG like '[02568ABCRST]' THEN 
                 Case 
                   when m.[sd_price] >= [s_apice4] then Isnull(m.sd_dis, 0)
                   else 0
                 end
               WHEN m.SD_SLIP_FG like '[13479]' THEN 
                 Case 
                   when m.[sd_price] >= [s_apice4] then Isnull(-m.sd_dis, 0)
                   else 0
                 end
               else 0
             END * isnull(d9.sp_rate, 1),
           -- 2013/1/17 協理與副總確認，當貨品單價高於等於小盤價時則另外認列激勵業績，日後的業績報表必須依此呈現。
           [Chg_sd_sale_oversmall_tot]= --(此金額為未稅金額)
             Case 
               WHEN m.SD_SLIP_FG like '[02568ABCRST]' THEN 
                 Case 
                   when m.[sd_price] >= [s_apice5] then 0
                   else Isnull(m.SD_STOT, 0)
                 end
               WHEN m.SD_SLIP_FG like '[13479]' THEN 
                 Case 
                   when m.[sd_price] > [s_apice5] then 0
                   else Isnull(-m.SD_STOT, 0)
                 end
               else 0
             END * isnull(d9.sp_rate, 1),
           [Chg_sd_sale_oversmall_dis]= --(此金額為未稅金額)
             Case 
               WHEN m.SD_SLIP_FG like '[02568ABCRST]' THEN 
                 Case 
                   when m.[sd_price] >= [s_apice5] then 0
                   else Isnull(sd_dis, 0)
                 end
               WHEN m.SD_SLIP_FG like '[13479]' THEN 
                 Case 
                   when m.[sd_price] > [s_apice5] then 0
                   else Isnull(-sd_dis, 0)
                 end
               else 0
             END * isnull(d9.sp_rate, 1),
           -- 單價
           [Chg_sd_price] = 
             Case 
               WHEN m.SD_SLIP_FG collate Chinese_Taiwan_Stroke_CI_AS like '[02568ABCRST]' THEN Isnull(m.sd_price, 0)
               WHEN m.SD_SLIP_FG collate Chinese_Taiwan_Stroke_CI_AS like '[13479]' THEN Isnull(-m.sd_price, 0)
               else 0
             END * isnull(d9.sp_rate, 1), 
           -- 含稅單價
           [Chg_sd_price_tax] = 
             Convert(decimal (18,4), 
               Case 
                 WHEN m.SD_SLIP_FG collate Chinese_Taiwan_Stroke_CI_AS like '[02568ABCRST]' THEN Isnull(m.sd_price, 0)
                 WHEN m.SD_SLIP_FG collate Chinese_Taiwan_Stroke_CI_AS like '[13479]' THEN Isnull(-m.sd_price, 0)
                 else 0
               END *
               -- 2015/04/23 Rickliu 會議決議將業績予以含稅，在此則將每筆明細的稅金加以計算, 並取小數點第四位, 以避免明細稅金與表頭稅金有所落差
               Case 
                  when SP_Tax <> 0 then 1.05
                  else 0
               end * 
               isnull(d9.sp_rate, 1)), 
           -- 小計 含稅價
           [Chg_sd_stot_tax] = 
             Convert(decimal (18,4), 
               Case 
                 WHEN m.SD_SLIP_FG collate Chinese_Taiwan_Stroke_CI_AS like '[02568ABCRST]' THEN Isnull(m.SD_STOT, 0)
                 WHEN m.SD_SLIP_FG collate Chinese_Taiwan_Stroke_CI_AS like '[13479]' THEN Isnull(-m.SD_STOT, 0)
                 else 0
               END * 1.05 * isnull(d9.sp_rate, 1)), 
           -- 小計稅金
           [Chg_sd_tax] = 
             Convert(decimal (18,4), 
               Case 
                 WHEN m.SD_SLIP_FG collate Chinese_Taiwan_Stroke_CI_AS like '[02568ABCRST]' THEN Isnull(m.SD_STOT, 0)
                 WHEN m.SD_SLIP_FG collate Chinese_Taiwan_Stroke_CI_AS like '[13479]' THEN Isnull(-m.SD_STOT, 0)
                 else 0
               END  * 0.05 * isnull(d9.sp_rate, 1)), 
           -- 2014/10/02 Rickliu 新增低於中盤價(售價 - 中盤價)
           Chg_Low_sd_price = IsNull(m.sd_price, 0) - IsNull(s_price4, 0),
           
           -- 2015/03/05 Rickliu 新增成本小計 (單位成本 * 數量)
           Chg_Cost_stot = 
              Case 
                WHEN m.SD_SLIP_FG like '[02568ABCRST]' THEN Isnull(m.sd_ave_p, 0) 
                WHEN m.SD_SLIP_FG like '[13479]' THEN Isnull(-m.sd_ave_p, 0)
                else 0
              END * isnull(d9.sp_rate, 1) * sd_Qty,

           -- 2014/10/02 Rickliu 新增毛利 (售價 - 單位成本) * 數量
           Chg_Profit = 
             (Case 
                WHEN m.SD_SLIP_FG like '[02568ABCRST]' THEN Isnull(m.sd_price, 0) 
                WHEN m.SD_SLIP_FG like '[13479]' THEN Isnull(-m.sd_price, 0)
                else 0
              END * isnull(d9.sp_rate, 1)  - 
              Case 
                WHEN m.SD_SLIP_FG like '[02568ABCRST]' THEN Isnull(m.sd_ave_p, 0)
                WHEN m.SD_SLIP_FG like '[13479]' THEN Isnull(-m.sd_ave_p, 0)
                else 0
              END * isnull(d9.sp_rate, 1)) * Isnull(sd_Qty, 0),

           -- 2014/10/02 Rickliu 新增毛利率 (毛利 / 銷貨淨額)
           Chg_Profit_Rate =
              Case 
                WHEN Isnull(m.sd_price, 0) * Isnull(sd_Qty, 0) = 0 THEN 0
                WHEN m.SD_SLIP_FG like '[02568ABCRST]' THEN 
                     ((Isnull(m.sd_price, 0)   * isnull(d9.sp_rate, 1) * Isnull(sd_Qty, 0))   - 
                      (Isnull(m.sd_ave_p, 0) * isnull(d9.sp_rate, 1) * Isnull(sd_Qty, 0)))  / 
                      (Isnull(m.sd_price, 0)   * isnull(d9.sp_rate, 1) * Isnull(sd_Qty, 0))
                WHEN m.SD_SLIP_FG like '[13479]' THEN 
                     ((Isnull(m.sd_price, 0)   * isnull(d9.sp_rate, 1) * Isnull(-sd_Qty, 0))  - 
                      (Isnull(m.sd_ave_p, 0) * isnull(d9.sp_rate, 1) * Isnull(-sd_Qty, 0))) / 
                      (Isnull(m.sd_price, 0)   * isnull(d9.sp_rate, 1) * Isnull(-sd_Qty, 0))
                else 0
              end,
              -- 2015/01/19 Rickliu 新增銷售成本 (標準成本 * 數量)
           Chg_save_stot =
             Case 
               WHEN m.SD_SLIP_FG like '[02568ABCRST]' THEN IsNull(sk_save, 0) * Isnull(sd_Qty, 0)
               WHEN m.SD_SLIP_FG like '[13479]' THEN IsNull(sk_save, 0) * Isnull(-sd_Qty, 0)
               else 0
             end,
           -- 平均攤提折讓(不依客戶收款方式及類別進行攤提) Chg_SP_Dis 已經內含匯率
           Chg_sd_dis_avg =
             Round(
               Case
                 when Chg_SP_Dis = 0 then 0
                 else Chg_SP_Dis / (select count(1) from SYNC_TA13.dbo.sslpdt where sd_slip_fg = m.sd_slip_fg and sd_no = m.sd_no) 
               end, 4),

           Chg_sd_dis_rate =
            Convert(decimal (18,4), 
              Case
                 when Chg_SP_Dis = 0 then 0
                 else (Case 
                         WHEN m.SD_SLIP_FG like '[02568ABCRST]' THEN Isnull(m.SD_STOT, 0)
                         WHEN m.SD_SLIP_FG like '[13479]' THEN Isnull(-m.SD_STOT, 0)
                         else 0
                       END * 
                       isnull(d9.sp_rate, 1) *
                       Case
                         when sp_tax <> 0 then 1.05
                         else 1
                       end
                      ) / Chg_SP_Dis  
                 --select count(*) from SYNC_TA13.dbo.sslpdt where sd_slip_fg = m.sd_slip_fg and sd_no = m.sd_no) 
               end),
           -- 單據收退款業績達成率
           Chg_SP_Pay_Rate,
           -- 商品收退款業績金額(含稅、含折讓)
           Chg_sd_Pay_Sale =
             Convert(decimal (18,4), 
               Case
                 when (SP_TOT = 0) Then 0 -- 已收金額
                 else (m.sd_stot * (case when sp_tax <> 0  then 1.05 else 1 end) * isnull(sp_rate, 1)) * Chg_SP_Pay_Rate
               end
               ),
           -- 2015/10/08 Rickliu  商品收退款業績毛利(含稅、含折讓)
           Chg_sd_Pay_Profit =
             Convert(decimal (18,4), 
               Case
                 when (SP_TOT = 0) Then 0 -- 已收金額
                 else ((Case 
                          WHEN m.SD_SLIP_FG like '[02568ABCRST]' THEN Isnull(m.sd_price, 0) 
                          WHEN m.SD_SLIP_FG like '[13479]' THEN Isnull(-m.sd_price, 0)
                          else 0
                        END * isnull(d9.sp_rate, 1)  - 
                        Case 
                          WHEN m.SD_SLIP_FG like '[02568ABCRST]' THEN Isnull(m.sd_ave_p, 0)
                          WHEN m.SD_SLIP_FG like '[13479]' THEN Isnull(-m.sd_ave_p, 0)
                          else 0
                        END * isnull(d9.sp_rate, 1)) * Isnull(sd_Qty, 0)
                        * (case when sp_tax <> 0  then 1.05 else 1 end) * isnull(sp_rate, 1)) 
                        * Chg_SP_Pay_Rate
               end
               ),
/**********************************************************************************************************************************************************
  表頭資料區
 **********************************************************************************************************************************************************/
           [###sslip###]='### sslip ###',
           -- 2015/03/02 Rickliu 新增本張成本
           [Chg_sp_ave_p],
           -- 2015/03/02 Rickliu 整張月平均成本
           [Chg_sp_mave_p],
           --2014/05/15 Rickliu -- 僅在明細的第一筆呈現銷貨金額
           [Chg_SP_Sale_Tot]=
             case
               when m.sd_slip_fg = '2' and sd_seqfld=1 then sp_tot * isnull(sp_rate, 1)
             else 0
             end,
           --2014/05/15 Rickliu -- 僅在明細的第一筆呈現退貨金額
           [Chg_SP_Ret_Tot]=
             case
               when m.sd_slip_fg = '3' and sd_seqfld=1 then sp_tot * isnull(sp_rate, 1)
               else 0
             end,
           --2014/05/15 Rickliu-- 僅在明細的第一筆呈現稅金金額(此為單頭資料，已經有計算過幣別)
           [Chg_sp_tax]=
             Case
               when sd_seqfld=1 then Chg_sp_tax * isnull(sp_rate, 1)
               else 0
             end,
           --2014/05/15 Rickliu-- 僅在明細的第一筆呈現折讓金額(此為單頭資料，已經有計算過幣別)
           [Chg_sp_dis]=
             Case
               when sd_seqfld=1 then Chg_sp_dis * isnull(sp_rate, 1)
               else 0
             end,
------------------------------------------------------------------------------------------------------------------------           
------------------------------------------------------------------------------------------------------------------------
------------------------------------------------------------------------------------------------------------------------           
           --2016/03/31 NanLiao-- 僅在明細的第一筆呈現折讓金額含稅(此為單頭資料，已經有計算過幣別)
           [Chg_sp_dis_tax]=
             Case
                when m.sd_seqfld=1 
                then Chg_sp_dis * isnull(sp_rate, 1) * 1.05
               else 0
             end,

           --2016/03/31 NanLiao-- 在明細的呈現折讓Tag(此為單頭資料，已經有計算過幣別)
           [Chg_sp_dis_flg] =
             Case
               when SUBSTRING(m.sd_ctno, 9, 1) = '2' then 'AB'
               else 'AA'
             end,
------------------------------------------------------------------------------------------------------------------------           
------------------------------------------------------------------------------------------------------------------------
------------------------------------------------------------------------------------------------------------------------     
           --2016/09/13 NanLiao-- 僅在明細的第一筆呈現折讓金額(此為折讓單資料，已經有計算過幣別)
           [Chg_sp_dis2]=
             Case
               when sd_seqfld=1 then Chg_sp_dis2
               else 0
             end,

           --2016/09/13 NanLiao-- 僅在明細的第一筆呈現折讓稅金(此為折讓單資料，已經有計算過幣別)
           [Chg_sp_dis_tax2]=
             Case
               when sd_seqfld=1 then Chg_sp_dis_tax2
               else 0
             end,
             
           --2016/09/13 NanLiao-- 僅在明細的第一筆呈現折讓金額含稅(此為折讓單資料，已經有計算過幣別)
           [Chg_sp_dis_tot2]=
             Case
               when sd_seqfld=1 then Chg_sp_dis_tot2
               else 0
             end,

------------------------------------------------------------------------------------------------------------------------           
------------------------------------------------------------------------------------------------------------------------
------------------------------------------------------------------------------------------------------------------------           
           --2014/05/15 Rickliu-- 僅在明細的第一筆呈現合計金額(此為單頭資料，已經有計算過幣別)
           [Chg_sp_sum_tot]=
             Case
               when sd_seqfld=1 then Chg_sp_sum_tot
               else 0
             end,
           --2014/05/15 Rickliu-- 僅在明細的第一筆呈現應收付金額(此為單頭資料，已經有計算過幣別)
           [Chg_Non_SP_Pay]=
             Case
               when sd_seqfld=1 then Chg_sp_sum_tot - Chg_SP_PAY
               else 0
             end,
           -- 2015/02/26 Rickliu 增加 加減項金額欄位
           -- 加項金額
           [Chg_SP_PAMT]=
             Case
               when sd_seqfld=1 then [Chg_SP_PAMT]
               else 0
             end,
           -- 減項金額
           [Chg_SP_MAMT]=
             Case
               when sd_seqfld=1 then [Chg_SP_MAMT]
               else 0
             end,
           -- 2014/06/24 Rickliu 新增單據完整名稱
           [Chg_sp_slip_name],
          
           --主檔的銷售日期
           [Chg_sp_date_Year],
           [Chg_sp_date_Quarter],
           [Chg_sp_date_Month],
           [Chg_sp_date_YearWeek],
           [Chg_sp_date_YM],
           [Chg_sp_date_MD],
           -- 2014/1/27 Rickliu 新增 單據日期的星期起訖
           [Chg_sp_date_Weekno],
           [Chg_sp_date_Weekname],
           [Chg_sp_date_Week_Range],

           --主檔的請款日期
           Chg_sp_pdate_Year, 
           Chg_sp_pdate_Quarter, 
           Chg_sp_pdate_Month, 
           Chg_sp_pdate_YearWeek, 
           [Chg_sp_pdate_YM],
           [Chg_sp_pdate_MD],
           -- 2014/1/27 Rickliu 新增 單據請款日期的星期起訖
           [Chg_sp_pdate_Weekno],
           [Chg_sp_pdate_Weekname],
           [Chg_sp_pdate_Week_Range],
           -- 2015/03/07 Rickliu 新增 表頭的單據附註
           [sp_maker],
           [sp_rem],
           -- 2015/03/07 Rickliu 物流目前是以調撥單據作為紀錄退貨原因所在，並在表頭的附註填寫單據編號，而在表身的附註填寫退貨原因，因此為了統計退貨原因，
           -- 在此新增特殊欄位
           [Chg_Logistics_Rej_flg] = 
             case
               when Isnumeric(Substring(sp_rem, 5, 10)) = 1 and Len(RTrim(Substring(sp_rem, 5, 10))) = 10 and Isnull(D13.Code_Begin, '') like '[0-7RSC]'
               then Isnull(D13.Code_Begin, '')
               else ''
             end,               
           [Chg_Logistics_Rej_sp_name] = 
             case
               when Isnumeric(Substring(sp_rem, 5, 10)) = 1 and Len(RTrim(Substring(sp_rem, 5, 10))) = 10 and Isnull(D13.Code_Begin, '') like '[0-7RSC]'
               then RTrim(Substring(Sp_rem, 1, 3))
               else ''
             end,               
           [Chg_Logistics_Rej_sp_no] = 
             case
               when Isnumeric(Substring(sp_rem, 5, 10)) = 1 and Len(RTrim(Substring(sp_rem, 5, 10))) = 10 and Isnull(D13.Code_Begin, '') like '[0-7RSC]'
               then Isnull(Substring(sp_rem, 5, 10), '')
               else ''
             end ,
           [Chg_Logistics_Rej_sp_date] = 
             case
               when Isnumeric(Substring(sp_rem, 5, 10)) = 1 and Len(RTrim(Substring(sp_rem, 5, 10))) = 10 and Isnull(D13.Code_Begin, '') like '[0-7RSC]'
               then 
                   (select sp_date
                      from fact_sslip 
                     where sp_slip_fg = Isnull(D13.Code_Begin, '') Collate Chinese_Taiwan_Stroke_CI_AS 
                       and sp_no = Substring(sp_rem, 5, 10) Collate Chinese_Taiwan_Stroke_CI_AS)
               else ''
             end,
           [Chg_Logistics_Rej_sp_ctno] = 
             case
               when Isnumeric(Substring(sp_rem, 5, 10)) = 1 and Len(RTrim(Substring(sp_rem, 5, 10))) = 10 and Isnull(D13.Code_Begin, '') like '[0-7RSC]' 
               then 
                   (select sp_ctno 
                      from fact_sslip 
                     where sp_slip_fg = Isnull(D13.Code_Begin, '') Collate Chinese_Taiwan_Stroke_CI_AS 
                       and sp_no = Substring(sp_rem, 5, 10) Collate Chinese_Taiwan_Stroke_CI_AS)
               else ''
             end,
           [Chg_Logistics_Rej_sp_ctname] = 
             case
               when Isnumeric(Substring(sp_rem, 5, 10)) = 1 and Len(RTrim(Substring(sp_rem, 5, 10))) = 10 and Isnull(D13.Code_Begin, '') like '[0-7RSC]'
               then
                    (select sp_ctname 
                       from fact_sslip 
                      where sp_slip_fg = Isnull(D13.Code_Begin, '') Collate Chinese_Taiwan_Stroke_CI_AS 
                        and sp_no = Substring(sp_rem, 5, 10)) Collate Chinese_Taiwan_Stroke_CI_AS
               else ''
             end,
           -- 2017/08/04 rickliu 新增單據客戶之業務是否離職
           ct_sale_leave,
/**********************************************************************************************************************************************************
  員工基本資料區
 **********************************************************************************************************************************************************/
           [###pemploy###]='### pemploy ###',
           --2014/05/15 Rickliu -- 將單據上的重要欄位也一併帶入，如此便可不需要額外對主檔作 JOIN，可以加速查詢速度。
           sp_sales, -- 業務員編
           E_NAME as sp_sales_name, -- 業務名稱
           ct_class, -- 客戶或廠商類別
           sp_ctno, -- 客戶或廠商編號
           ct_no8, -- 客戶八碼編號
           ct_sname8, -- 客戶八碼名稱
           sp_ctname, -- 客戶或廠商名稱
           ct_sname, -- 客戶或廠商簡稱
           ct_fld3, 
           e_dept, -- 部門編號
           Chg_sp_dp_Name,
           Chg_dp_name, -- 部門名稱
           ct_loc, -- 區域編號
           Chg_loc_Name, -- 區域名稱
           Chg_ctno_CustKind_CustCity, -- 客戶類型|縣市別名稱
           Chg_ctno_CustChain, -- 通路別名稱
           Chg_busine_Name, -- 行業別名稱
           ct_sour, -- 客戶來源編號
           Chg_sour_Name, -- 客戶來源名稱
           Chg_ctno_Port_Office, -- 內外銷|公司別名稱
           ct_kind, -- 客戶類別
           Chg_ctclass, --客戶類別名稱
           sp_mstno, -- 業務組別編號
           Chg_sp_mst_name, --業務組別名稱
           chg_leave, -- 離職否
/**********************************************************************************************************************************************************
  客戶基本資料區
 **********************************************************************************************************************************************************/
           [###pcust###]='### pcust ###',
           -- 2015/02/26 Rickliu 增加 客戶 銷貨類型, 取編號第九碼 1: 百貨, 2: 3C, 3, 寄賣
           Chg_Cust_Sale_Class,
           Chg_Cust_Sale_Class_Name,
           Chg_Cust_Sale_Class_sName,
           D9.chg_ct_close,
           -- 2014/01/27 Rickliu新增 網通客戶
           D9.Chg_IS_Lan_Custom,
           Chg_BU_NO, --總公司編號
           rtrim(ct_fld3) as Chg_ct_fld3, --Rickliu 2014/1/24 新增 總公司
           
           Chg_Hunderd_Customer, --Rickliu 2014/3/17 新增 是否為百大
           Chg_Hunderd_Customer_Name, --Rickliu 2014/1/24 百大客戶名稱
           -- 2014/10/02 Rickliu 新增分店最後一次交易旗標
           Chg_CTNO_Last_Order_Flag = IsNull(D11.Flag, ''),

           -- 2014/10/02 Rickliu 新增總公司最後一次交易旗標
           Chg_BUNO_Last_Order_Flag = IsNull(D12.Flag, ''),
/**********************************************************************************************************************************************************
  貨品基本資料區
 **********************************************************************************************************************************************************/
           -- 由於一個資料表的行儲存不能超過 8,060 字元，所以必須挑必要維度進來
           [###sstock###]='### sstock ###',
           [sk_bcode], --條碼代號
           [sk_kind], --類別
           [sk_ivkd], --發票類別
           [sk_spec], --規格
           [sk_unit], --基本單位
           [sk_color], --顏色 改 產品下架日期(西元年月)
           [sk_size], --尺寸 改 新品到貨日期(西元年月)
           [sk_use], --用途
           [sk_aux_q], --輔助單位數量
           [sk_aux_u], --輔助單位
           [sk_save], --標準成本
           [sk_bqty], --經濟批量
           [sk_fign], --圖形檔名
           [sk_pic], -- Web 方式呈現圖形檔名, 2017/07/12 Rickliu
           [sk_tfg], --傳輸旗標
           [st_cspes], --貨品規格 改 材積
           [st_espes], --貨品規格(英)
           [st_unit], --單位
           [st_inerqty], --內包裝單位包裝數
           [st_inerunt], --內包裝單位
           [st_outqty], --外包裝單位數
           [st_outunt], --外包裝單位
           [st_appr], --每箱材積
           [st_apprunt], --材積單位(CBM或CFT)
           [st_nw], --淨重
           [st_gw], --毛重
           [st_gwunt], --重量單位
           [st_itclass], --貨品分類
           [st_cccode], --CCC碼
           [st_sespes], --品名(英)
           [st_20cyqty], --20呎貨櫃可容納數量
           [st_40cyqty], --40呎貨櫃可容納數量
           [st_45cyqty], --45呎貨櫃可容納數量
           [sk_fld1], --附註一 改 貨品長寬高
           [sk_fld2], --附註二 改 貨品規格
           [sk_fld3], --附註三 改 容量
           [sk_fld4], --附註四 改 品牌-ABT/M&M
           [sk_fld5], --附註五 改 季節/節日
           [sk_fld6], --附註六 改 味道
           [sk_fld11], --附註十一-滯銷確認日
           [sk_fld12], --附註十二-(舊編號 2012開帳使用)
           [sk_ikind], --物料類別
           [s_supp], --供應商
           [s_locate], --存放地點
           [s_price1], --期初成本
           [s_price2], --市價
           [s_price3], --大盤價
           [s_price4], --中盤價
           [s_price5], --小盤價
           [s_price6], --一般價
           [s_aprice], --目前平均成本
           [s_m_ave], --月平均成本
           [s_updat1], --最近進貨日
           [s_updat2], --最近銷貨日
           [s_lprice1], --最近進價
           [s_lprice2], --最近銷價
           [s_accno1], --進貨 科目代號
           [sk_nowqty], --目前存量
           [st_proven], --貨品來源
           [st_lenght], --體積規格 長
           [st_width], --體積規格 寬
           [st_height], --體積規格 高
           [st_lwhunit], --體積規格長寬高單位
           [st_unw], --貨品單位淨重
           [st_ugw], --貨品單位毛重
           [st_uappr], --貨品單位材積
           [st_uarea], --貨品單位面積
           [st_areaunt], --貨品面積單位
           [st_fign2], --圖檔名2
           [st_fign3], --圖檔名3
           [st_fign4], --圖檔名4
           [sk_poseq], --主供應商序號
           [sk_abcode], --包裝條碼代號
           [s_apice2], --市  價(輔)
           [s_apice3], --大盤價(輔)
           [s_apice4], --中盤價(輔)
           [s_apice5], --小盤價(輔)
           [s_apice6], --一般價(輔)
           [sk_whno], --入庫倉庫
           [Chg_skno_accno],
           [Chg_skno_accno_Name],
           [Chg_skno_BKind],
           [Chg_skno_Bkind_Name],
           [Chg_skno_BKind2],
           [Chg_skno_Bkind_Name2],
           [Chg_skno_SKind],
           [Chg_skno_SKind_Name],
           [Chg_kind_Name],
           [Chg_ivkd_Name],
           [Chg_StartYear],
           [Chg_StartMonth],
           [Chg_EndYear],
           [Chg_EndMonth],
           [Chg_supp_Name],
           [Chg_supp_SName],
           [Chg_locate_MArea],
           [Chg_locate_DArea],
           [Chg_locate_row],
           [Chg_locate_col],
           [Chg_updat1_Year],
           [Chg_updat1_Month],
           [Chg_updat2_Year],
           [Chg_updat2_Month],
           [Chg_New_Arrival_Date], -- 2015/03/06 Rickliu 新增新品到貨日
           -- 2014/07/02 Rickliu 新增採購新品
           --Chg_IS_New_Stock =
           --  Case
           --    when (D4.[sk_no] is not null) And ([Chg_sp_date_YM] >= D4.[Chg_New_Arrival_YM]) then 'Y'
           --    else 'N'
           --  end,
           -- 2017/08/01 Rickliu 修訂新品定義，改為以一年內之引進之產品為新品
           [Chg_IS_New_Stock],
           D4.[Chg_New_First_Qty],
           chg_new_arrival_ym,
           
           -- 2017/08/07 Rickliu 取消使用商品未銷列表
           -- Chg_Stock_NonSales, 
           -- 2017/08/07 Rickliu 改以商品基本資料取得是否為主銷
           [Chg_IS_Master_Stock],
           stock_kind_list,
/**********************************************************************************************************************************************************
  庫存資料區
 **********************************************************************************************************************************************************/
           [###sware###]='### sware ###',
           -- 2014/06/17 Rickliu新增 AA 期初安全存量
           [Chg_WD_AA_sQty],
           -- 2014/06/17 Rickliu新增 AA 期初現有庫存數
           [Chg_WD_AA_first_Qty],           
           -- 2014/06/17 Rickliu新增 AA 期初庫存差異數
           [Chg_WD_AA_first_diff_Qty],
           -- 2014/06/07 Rickliu新增 AA 期末現有庫存數
           [Chg_WD_AA_last_Qty],
           -- 2014/06/17 Rickliu新增 AA 期末庫存差異數
           [Chg_WD_AA_last_diff_Qty],
             
           -- 2014/06/17 Rickliu新增 AB 期初安全存量
           [Chg_WD_AB_sQty],
           -- 2014/06/17 Rickliu新增 AB 期初現有庫存數
           [Chg_WD_AB_first_Qty],           
           -- 2014/06/17 Rickliu新增 AB 期初庫存差異數
           [Chg_WD_AB_first_diff_Qty],
           -- 2014/06/07 Rickliu新增 AB 期末現有庫存數
           [Chg_WD_AB_last_Qty],
           -- 2014/06/17 Rickliu新增 AB 期末庫存差異數
           [Chg_WD_AB_last_diff_Qty],
           
           -- 2014/06/17 Rickliu新增 AC 期初安全存量
           [Chg_WD_AC_sQty],
           -- 2014/06/17 Rickliu新增 AC 期初現有庫存數
           [Chg_WD_AC_first_Qty],           
           -- 2014/06/17 Rickliu新增 AC 期初庫存差異數
           [Chg_WD_AC_first_diff_Qty],
           -- 2014/06/07 Rickliu新增 AC 期末現有庫存數
           [Chg_WD_AC_last_Qty],
           -- 2014/06/17 Rickliu新增 AC 期末庫存差異數
           [Chg_WD_AC_last_diff_Qty],

           -- 2014/06/17 Rickliu新增 AG 期初安全存量
           [Chg_WD_AG_sQty],
           -- 2014/06/17 Rickliu新增 AG 期初現有庫存數
           [Chg_WD_AG_first_Qty],           
           -- 2014/06/17 Rickliu新增 AG 期初庫存差異數
           [Chg_WD_AG_first_diff_Qty],
           -- 2014/06/07 Rickliu新增 AG 期末現有庫存數
           [Chg_WD_AG_last_Qty],
           -- 2014/06/17 Rickliu新增 AG 期末庫存差異數
           [Chg_WD_AG_last_diff_Qty],

           -- 2014/06/16 Rickliu 新增滯銷品
           [Chg_IS_Dead_Stock],
           [Chg_Dead_First_Qty],
           [Chg_Dead_First_Amt],
           [Chg_Dead_Stock_YM],
           
           -- 2015/02/27 Rickliu 新增部門銷售通路類型, 取編號第九碼 1: 百貨, 2: 3C, 3: 寄賣, 4:代工(OEM)
           [Chg_Dept_Cust_Chain_No],
           [Chg_Dept_Cust_Chain_Name],

           sslpdt_update_datetime = getdate(),
           sslpdt_timestamp = m.timestamp_column
           into Fact_sslpdt 
      from SYNC_TA13.dbo.sslpdt M With(NoLock)
           Left join SYNC_TA13.dbo.sware D2 With(NoLock) On M.sd_whno=D2.wh_no 
           Left join Fact_sstock D4 With(NoLock) On M.sd_skno=D4.sk_no
           Left join Ori_Xls#Sys_Code D5 With(NoLock) On  D5.code_class ='7' and M.sd_class = D5.Code_Begin Collate Chinese_Taiwan_Stroke_CI_AS 
           Left join Ori_Xls#Sys_Code D6 With(NoLock) On  D6.code_class ='8' and M.sd_class = D6.Code_Begin Collate Chinese_Taiwan_Stroke_CI_AS and M.sd_slip_fg = D6.Code_End Collate Chinese_Taiwan_Stroke_CI_AS
           Left join Fact_sslip D9 With(NoLock) On sd_class = D9.sp_class and sd_slip_fg = D9.sp_slip_fg and sd_no = D9.sp_no 
           -- 2017/08/07 Rickliu 改以商品基本資料取得是否為主銷
           --left join Ori_XLS#Master_Stock D8 With(NoLock)
           --  On Year(D9.sp_pdate) = D8.Sale_Year and Month(D9.sp_pdate) = D8.Sale_month
           -- and M.sd_skno collate Chinese_Taiwan_Stroke_CI_AS = D8.sk_no collate Chinese_Taiwan_Stroke_CI_AS 
           -- 2017/08/07 Rickliu 改以商品基本資料取得是否為主銷
           --left join Ori_XLS#Master_Stock D10 With(NoLock) 
           --  On Year(m.sd_date) = D10.Sale_Year and Month(m.sd_date) = D10.Sale_month
           -- 2013/6/1 變更主銷結構，取消 MasterKey 欄位，主銷的 SK_NAME 則改為主銷商品名稱
           -- and M.sd_skno collate Chinese_Taiwan_Stroke_CI_AS = D10.Masterkey collate Chinese_Taiwan_Stroke_CI_AS
           -- and M.sd_skno collate Chinese_Taiwan_Stroke_CI_AS = D10.sk_no collate Chinese_Taiwan_Stroke_CI_AS
            -- 2014/10/02 Rickliu 增加最後分店一次交易旗標
           left join CTE_Q1 D11
            On m.sd_class = d11.sd_class
           and m.sd_slip_fg = d11.sd_slip_fg
           and m.sd_ctno = d11.sd_ctno
           and m.sd_no = d11.Max_SP_NO
           and m.sd_skno = d11.sd_skno
           and m.sd_price = d11.sd_price -- Rickliu 2017/06/08 CTE_Q1 增加判斷過濾贈品註記、金額小於1 也算贈品

           -- 2014/10/02 Rickliu 增加最後總公司一次交易旗標
           left join CTE_Q2 D12
            On m.sd_class = d12.sd_class
           and m.sd_slip_fg = d12.sd_slip_fg
           --2016/09/09 Rickliu 因組織擴編，增加機車通路，因此將原本 2~5 改為 2~6
           --and Substring(m.sd_ctno, 1, 5) = d12.BU_NO
           and Substring(m.sd_ctno, 1, 6) = d12.BU_NO
           and m.sd_no = d12.Max_SP_NO
           and m.sd_skno = d12.sd_skno
           and m.sd_price = d12.sd_price -- Rickliu 2017/06/08 CTE_Q1 增加判斷過濾贈品註記、金額小於1 也算贈品
           ---
           left join Ori_Xls#Sys_Code D13 With(NoLock) 
             On M.sd_date >= '2015/03/01' 
            and M.sd_slip_fg = 'B'
            and D13.code_class ='8' 
            and D13.Code_Name = RTrim(Substring(Sp_rem, 1, 3)) Collate Chinese_Taiwan_Stroke_CI_AS
    where 1=1
      -- 2017/08/08 Rickliu 從 2013 年起算，保留近五年資料
      and m.sd_date > '2013/01/01'
      and m.sd_date >= Convert(Varchar(5), year(DateAdd(year, -5, getdate())))+ '/01/01' 
    
  /*==============================================================*/
  /* Index: sslpdt_Timestamp                                      */
  /*==============================================================*/
  --set @Msg = '建立索引 [Fact_sslpdt.sslpdt_timestamp]'
  --Print @Msg
  --if exists (select * from sys.indexes where object_id = object_id('[dbo].[fact_sslpdt]') and name = 'pk_sslpdt')
  --   alter table [dbo].[pk_sslpdt] drop constraint [pk_sslpdt]

  --alter table [dbo].[fact_sslpdt] add  constraint [pk_sslpdt] primary key nonclustered ([sslpdt_timestamp] asc) with 
  --(pad_index  = off, statistics_norecompute  = off, sort_in_tempdb = off, ignore_dup_key = off, online = off, allow_row_locks  = on, allow_page_locks  = on) on [primary]


  /*==============================================================*/
  /* Index: sdno                                                  */
  /*==============================================================*/
  Set @Msg = '建立索引 [fact_sslpdt.sdno]'
  Print @Msg
  if exists (select 1 from sysindexes where id = object_id('dbo.fact_sslpdt') and name  = 'sdno' and indid > 0 and indid < 255) 
     drop index dbo.fact_sslpdt.sdno
  create clustered index sdno on dbo.fact_sslpdt (sd_no, sd_slip_fg) with fillfactor= 30  on "PRIMARY"
  /*==============================================================*/
  /* Index: class                                                 */
  /*==============================================================*/
  Set @Msg = '建立索引 [fact_sslpdt.class]'
  Print @Msg
  if exists (select 1 from sysindexes where id = object_id('dbo.fact_sslpdt') and name  = 'class' and indid > 0 and indid < 255) 
     drop index dbo.fact_sslpdt.class
  create index class on dbo.fact_sslpdt (sd_class) with fillfactor= 30 on "PRIMARY"
  /*==============================================================*/
  /* Index: CSNO                                                  */
  /*==============================================================*/
  Set @Msg = '建立索引 [fact_sslpdt (sd_csno)...'
  Print @Msg
  if exists (select 1 from sysindexes where id = object_id('dbo.fact_sslpdt') and name  = 'CSNO' and indid > 0 and indid < 255) 
     drop index dbo.fact_sslpdt.CSNO
  create index CSNO on dbo.fact_sslpdt (sd_csno) with fillfactor= 30 on "PRIMARY"
  /*==============================================================*/
  /* Index: ctno                                                  */
  /*==============================================================*/
  Set @Msg = '建立索引 [fact_sslpdt (sd_ctno, sd_class)...'
  Print @Msg
  if exists (select 1 from sysindexes where id = object_id('dbo.fact_sslpdt') and name  = 'ctno' and indid > 0 and indid < 255) 
     drop index dbo.fact_sslpdt.ctno
  create index ctno on dbo.fact_sslpdt (sd_ctno, sd_class) with fillfactor= 30 on "PRIMARY"
  /*==============================================================*/
  /* Index: ordno                                                 */
  /*==============================================================*/
  Set @Msg = '建立索引 [fact_sslpdt.ordno]'
  Print @Msg
  if exists (select 1 from sysindexes where id = object_id('dbo.fact_sslpdt') and name  = 'ordno' and indid > 0 and indid < 255) 
     drop index dbo.fact_sslpdt.ordno
  create index ordno on dbo.fact_sslpdt (sd_ordno, sd_class) with fillfactor= 30 on "PRIMARY"
  /*==============================================================*/
  /* Index: sdda                                                  */
  /*==============================================================*/
  Set @Msg = '建立索引 [fact_sslpdt.sdda]'
  Print @Msg
  if exists (select 1 from sysindexes where id = object_id('dbo.fact_sslpdt') and name  = 'sdda' and indid > 0 and indid < 255) 
     drop index dbo.fact_sslpdt.sdda
  create index sdda on dbo.fact_sslpdt (sd_date, sd_class) with fillfactor= 30 on "PRIMARY"
  /*==============================================================*/
  /* Index: skda                                                  */
  /*==============================================================*/
  Set @Msg = '建立索引 [fact_sslpdt.skda]'
  Print @Msg
  if exists (select 1 from sysindexes where id = object_id('dbo.fact_sslpdt') and name  = 'skda' and indid > 0 and indid < 255) 
     drop index dbo.fact_sslpdt.skda
  create index skda on dbo.fact_sslpdt (sd_skno, sd_date) with fillfactor= 30 on "PRIMARY"
  /*==============================================================*/
  /* Index: skno                                                  */
  /*==============================================================*/
  Set @Msg = '建立索引 [fact_sslpdt.skno]'
  Print @Msg
  if exists (select 1 from sysindexes where id = object_id('dbo.fact_sslpdt') and name  = 'skno' and indid > 0 and indid < 255) 
     drop index dbo.fact_sslpdt.skno
  create index skno on dbo.fact_sslpdt (sd_skno, sd_ctno, sd_date, sd_class) with fillfactor= 30 on "PRIMARY"
  /*==============================================================*/
  /* Index: slip_fg                                               */
  /*==============================================================*/
  Set @Msg = '建立索引 [Fact_sslpdt.slip_fg]'
  Print @Msg
  if exists (select 1 from sysindexes where id = object_id('dbo.fact_sslpdt') and name  = 'slip_fg' and indid > 0 and indid < 255) 
     drop index dbo.fact_sslpdt.slip_fg
  create index slip_fg on dbo.fact_sslpdt (sd_slip_fg) with fillfactor= 30 on "PRIMARY"
  /*==============================================================*/
  /* Index: IX_sslpdt_N1                                          */
  /*==============================================================*/
  --  Print 'Create index [IX_sslpdt_N1] on [dbo].[Fact_sslpdt] ([Chg_skno_BKind], [sd_slip_fg], [Chg_sp_pdate_Year])...'
  --  if exists (select 1 from sysindexes where id = object_id('dbo.fact_sslpdt') and name  = 'IX_sslpdt_N1' and indid > 0 and indid < 255) 
  --     drop index dbo.fact_sslpdt.slip_fg
  --  create index [IX_sslpdt_N1] on [dbo].[Fact_sslpdt] ([Chg_skno_BKind], [sd_slip_fg], [Chg_sp_pdate_Year])
  --  INCLUDE ([Chg_sales_Name],[Chg_sd_Pay_Sale],[Chg_sp_pdate_Month],[sp_sales])


/*********************************************************************************************************************************************************
  2017/08/15 Ricliu 以下此表是專給 主銷、新品、滯銷之百大及非百大 鋪貨率 報表所用
********************************************************************************************************************************************************/
  IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[fact_sslpdt_Near_Year]') AND type in (N'U'))
  begin
     Set @Msg = '刪除資料表 [fact_sslpdt_Near_Year]'
     set @strSQL= 'DROP TABLE [dbo].[fact_sslpdt_Near_Year]'

     Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
  end

  select Chg_sp_pdate_YM,
         Chg_bu_no, ct_fld3 as Chg_bu_Name,
         ct_sales, chg_ct_sales_name,
         ct_no8, ct_sname8+
         case
           when chg_ct_close  = 'Y' then '[關]'
           else ''
         end as ct_sname8,
         chg_ct_close,
         sd_skno, sd_name, stock_kind_list,
         chg_is_master_stock,
         chg_is_New_Stock,
         chg_is_Dead_Stock,
         Chg_Hunderd_Customer,
         Sum(Chg_sd_qty) as Chg_sd_qty,
         Sum(Chg_sd_stot) as Chg_sd_stot
         into fact_sslpdt_Near_Year
    from fact_sslpdt with(NoLock)
   where 1=1
     and sd_class in ('1', '8')
     and Chg_sp_pdate_YM >='2013/01'
     and substring(ct_no8, 1, 2) Not in ('IT', 'IZ', 'ZZ')
     and substring(sd_skno, 1, 1) = 'A'
     and Chg_sp_pdate_YM >= Convert(Varchar(7), DateAdd(mm, -12,getdate()), 111)
     and ct_fld3 <> ''
   group by Chg_sp_pdate_YM, Chg_bu_no, ct_fld3,
            ct_sales, chg_ct_sales_name,
            ct_no8, ct_sname8, chg_ct_close,
            sd_skno, sd_name, stock_kind_list,
            chg_is_master_stock,
            chg_is_New_Stock,
            chg_is_Dead_Stock,
             Chg_Hunderd_Customer
   order by Chg_sp_pdate_YM, Chg_bu_no, ct_fld3,
            ct_sales, chg_ct_sales_name,
            ct_no8, ct_sname8, chg_ct_close,
            sd_skno, sd_name, stock_kind_list,
            chg_is_master_stock,
            chg_is_New_Stock,
            chg_is_Dead_Stock,
            Chg_Hunderd_Customer
   
  end try
  begin catch
    set @Cnt = -1
    set @Msg = @Proc+'...(錯誤訊息:'+ERROR_MESSAGE()+', '+@Msg+')...(錯誤列:'+Convert(Varchar(10), ERROR_LINE())+')'

    Exec uSP_Sys_Write_Log @Proc, @Msg, @Msg, @Cnt
  end catch
  Return(@Cnt)
end
GO
