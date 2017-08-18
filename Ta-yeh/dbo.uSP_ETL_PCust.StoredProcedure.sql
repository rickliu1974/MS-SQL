USE [DW]
GO
/****** Object:  StoredProcedure [dbo].[uSP_ETL_PCust]    Script Date: 08/18/2017 17:18:56 ******/
DROP PROCEDURE [dbo].[uSP_ETL_PCust]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE Proc [dbo].[uSP_ETL_PCust]
as
begin
  /*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
   Program Name: uSP_TA13#ETL_PCust
   Create Date: 2013/03/01
   Creator: Rickliu
   Updated Date: 2013/09/06 [修改 Trans_Log 訊息內容]
  *-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*/  

  Declare @Proc Varchar(50) = 'uSP_ETL_PCust'
  Declare @Cnt Int =0
  Declare @Cnt_Ori Int =0
  Declare @Msg Varchar(Max) =''
  Declare @strSQL Varchar(Max)=''
  Declare @CR Varchar(4) = char(13)+char(10)
  Declare @subject Varchar(500)= 'Exec '+@Proc+' Error!!'
  Declare @Result int = 0
  
  IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[Fact_PCust]') AND type in (N'U'))
  begin
     Set @Msg = '刪除資料表 [Fact_PCust]'
     set @strSQL= 'DROP TABLE [dbo].[Fact_PCust]'

     Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
  end

  begin try
    set @Msg = '建立資料表 [Fact_PCust]'
    Exec uSP_Sys_Write_Log @Proc, @Msg, @Msg, 0

    -- 2017/04/21 Rickliu 幣別最後一次日期
    ;With CTE_Q1 as (
       select r_name, max(r_date) as r_date
         from SYNC_TA13.dbo.prate
        group by r_name
    -- 2017/04/21 Rickliu 最新幣別
    ), CTE_Q2 as (
       select a.r_name, a.r_rate, a.r_date as r_date
         from SYNC_TA13.dbo.prate as a
         left join CTE_Q1 as b
           on a.r_name=b.r_name and a.r_date=b.r_date
        where a.r_name=b.r_name and a.r_date=b.r_date
    )

    select -- Distinct
           M.[ct_class], --類別
           Rtrim(M.[ct_no]) as ct_no, --編號
           Rtrim(M.[ct_name]) as ct_name, --名稱
           Rtrim(M.[ct_sname]) as ct_sname, --簡稱
           -- 2017/08/07 Rickliu 只要是 IT業務個人、IZ現收客戶、ZZ其他客戶 一律都將編成 尾碼為 000001
           Case
             when m.ct_class = '1' and Len(Rtrim(M.[ct_no])) > 8 and Substring(M.ct_no, 1, 2) in ('IT', 'IZ', 'ZZ')
             then Substring(M.ct_no, 1, 2)+'000001'
             else Substring(Rtrim(M.[ct_no]), 1, 8) 
           end as ct_no8, --客戶八碼編號(單一分店編號，六碼則為總公司編號)
           -- 20170428 Rickliu 原抓取 ct_sname 方式，仍會有錯誤情況，將改採抓取 ct_fld3+ct_fld4 方式呈現
           --LTrim(RTrim(Replace(Replace(Replace(Replace(Replace(Replace(m.ct_sname, '#', ''), '@', ''), '-', ''), 'T', ''), 'P', ''), '/', ''))) as ct_sname8, -- 客戶八碼名稱
           -- 2017/08/07 Rickliu 只要是 IT業務個人、IZ現收客戶、ZZ其他客戶 一律都歸單一名稱
           Case
             when m.ct_class = '1' and Len(Rtrim(M.[ct_no])) > 8 and Substring(M.ct_no, 1, 2) = 'IT' then '業務個人'
             when m.ct_class = '1' and Len(Rtrim(M.[ct_no])) > 8 and Substring(M.ct_no, 1, 2) = 'IZ' then '現收客戶'
             when m.ct_class = '1' and Len(Rtrim(M.[ct_no])) > 8 and Substring(M.ct_no, 1, 2) = 'ZZ' then '其他客戶'
             else rtrim(m.ct_fld3)+case when rtrim(m.ct_fld4) <> '' then '-' else '' end+rtrim(m.ct_fld4) 
           end as ct_sname8,

           [ct_addr1]=Rtrim(Convert(Varchar(255), M.ct_addr1)), --公司地址
           [ct_addr2]=Rtrim(Convert(Varchar(255), M.ct_addr2)), --發票地址
           [ct_addr3]=Rtrim(Convert(Varchar(255), M.ct_addr3)), --送貨地址
           Rtrim(M.[ct_tel]) as ct_tel, --電話
           Rtrim(M.[ct_fax]) as ct_fax, --傳真
           Rtrim(M.[ct_unino]) as ct_unino, --統一編號
           Rtrim(M.[ct_presid]) as ct_presid, --負責人
           Rtrim(M.[ct_contact]) as ct_contact, --聯絡人
           M.[ct_payfg], --售價種類
           Rtrim(M.[ct_sales]) as ct_sales, --業務員
           M.[ct_p_limit], --帳款額度
           M.[ct_b_limit], --信用額度
           Rtrim(M.[ct_bkno]) as ct_bkno, --銀行帳號
           Rtrim(M.[ct_bknm]) as ct_bknm, --銀行名稱
           [ct_rem]=Rtrim(Convert(Varchar(255), M.ct_rem)), --備註
           Rtrim(M.[ct_dept]) as ct_dept, --部門編號
           M.[ct_payrate], --售價折數
           Rtrim(M.[ct_ivtitle]) as ct_ivtitle, --發票抬頭
           Rtrim(M.[ct_porter]) as ct_porter, --貨運公司
           Rtrim(M.[ct_credit]) as ct_credit, --請款|付款對象
           M.[ct_pmode], --付款方式
           M.[ct_pdate], --付款天數
           M.[ct_prenpay], --未清款
           M.[ct_prepay], --預收款
           M.[ct_last_dt], --最近交易日
           M.[ct_flg], --客戶兼廠商
           M.[ct_t_fg], --傳輸旗標
           M.[ct_grade], --等級
           M.[ct_area], --地區,國別代碼
           -- 2015/02/03 交易幣別
           case
             when RTrim(isnull(M.ct_curt_id, '')) = '' then 'NT'
             else RTrim(isnull(M.ct_curt_id, ''))
           end as ct_curt_id,

           Rtrim(M.[ct_cont_sp]) as ct_cont_sp, --聯絡人職稱
           M.[ct_pay], --貿易付款方式
           Rtrim(M.[ct_regist]) as ct_regist, --工廠登記証
           M.[ct_worker], --員工數
           M.[ct_capital], --資本額
           M.[ct_skpay], --貿易預收款
           M.[ct_sknpay], --貿易未清款
           M.[ct_accno2], --貿易應付帳科目
           M.[ct_chkno2], --貿易應付票科目
           M.[ct_cdate], -- 建檔日期
           M.[ct_payer], --進銷存請款對象
           M.[ct_advance], --進銷存預收款
           M.[ct_debt], --進銷存未清款
           M.[ct_abroad], --國內.國外
           Rtrim(M.[ct_fld1]) as ct_fld1, --客戶|廠商附註一(客：開發日期、廠：配合起始日)
           Rtrim(M.[ct_fld2]) as ct_fld2, --客戶|廠商附註二(客：關店日期、廠：結束配合日)
           -- 2017/08/07 Rickliu 只要是 IT業務個人、IZ現收客戶、ZZ其他客戶 一律都歸單一名稱
           --Rtrim(M.[ct_fld3]) as ct_fld3, --客戶|廠商附註三(客：總公司別、廠：未使用)
           Case
             when m.ct_class = '1' and Len(Rtrim(M.[ct_no])) > 8 and Substring(M.ct_no, 1, 2) = 'IT' then '業務個人'
             when m.ct_class = '1' and Len(Rtrim(M.[ct_no])) > 8 and Substring(M.ct_no, 1, 2) = 'IZ' then '現收客戶'
             when m.ct_class = '1' and Len(Rtrim(M.[ct_no])) > 8 and Substring(M.ct_no, 1, 2) = 'ZZ' then '其他客戶'
             else rtrim(m.ct_fld3)
           end as ct_fld3,
           Rtrim(M.[ct_fld4]) as ct_fld4, --客戶|廠商附註四(客：分店別、廠：廠商評比)
           Rtrim(M.[ct_fld5]) as ct_fld5, --客戶|廠商附註五(客：客服專用1、廠：簽約到期日)
           Rtrim(M.[ct_fld6]) as ct_fld6, --客戶|廠商附註六(客：客服專用2、廠：未使用)
           Rtrim(M.[ct_fld7]) as ct_fld7, --客戶|廠商附註七(客：客服專用3、廠：未使用)
           Rtrim(M.[ct_fld8]) as ct_fld8, --客戶|廠商附註八(客：客服專用4、廠：未使用)
           Rtrim(M.[ct_fld9]) as ct_fld9, --客戶|廠商附註九(客：收款方式、廠：未使用)
           Rtrim(M.[ct_fld10]) as ct_fld10, --客戶|廠商附註十(客：發票名稱、廠：未使用)
           Rtrim(M.[ct_fld11]) as ct_fld11, --客戶|廠商附註十一(客：付款條件、廠：未使用)
           Rtrim(M.[ct_fld12]) as ct_fld12, --客戶|廠商附註十二(客：舊編號 2012開帳使用)
           M.[ct_sea], --貿易運送方式
           M.[ct_invofg], --開立發票種類
           M.[ct_udec], --交易單價的小數位數
           M.[ct_tdec], --交易金額的小數位數
           M.[ct_busine], --行業別
           Rtrim(M.[ct_banpay]) as ct_banpay, --押匯銀行
           M.[ct_loc], --區域編號
           M.[ct_sour], --客戶來源
           M.[ct_kind], --客戶類別
           M.[ct_tkday], --拜訪週期
           Chg_ctclass=
             Case 
               when M.ct_class = '1' and Len(m.ct_no) = 9 then '客戶'
               when M.ct_class = '2' and Len(m.ct_no) = 5 then '廠商'
             end,
           -- 客戶
           --2016/09/09 Rickliu 因組織擴編，增加機車通路，因此將原本 2~5 改為 2~6
           Case
             when m.ct_class = '1' and Len(Rtrim(M.[ct_no])) > 6 and Substring(M.ct_no, 1, 2) = 'IT' then 'IT0000'
             when m.ct_class = '1' and Len(Rtrim(M.[ct_no])) > 6 and Substring(M.ct_no, 1, 2) = 'IZ' then 'IZ0000'
             when m.ct_class = '1' and Len(Rtrim(M.[ct_no])) > 6 and Substring(M.ct_no, 1, 2) = 'ZZ' then 'ZZ0000'
             when m.ct_class = '1' and Len(Rtrim(M.[ct_no])) > 6 then Substring(Rtrim(M.[ct_no]), 1, 6) 
             else ''
           end as Chg_BU_NO, -- 客戶總公司編號
           --2017/06/21 Rickliu 因百大客戶需過濾總公司客編，因此新增此欄位
           Case
             when m.ct_class = '1' and Len(Rtrim(M.[ct_no])) > 6 and 
                  Substring(M.ct_no, 1, 2) Not in ('IT', 'IZ', 'ZZ') and
                  (m.ct_fld3 like '%總公司%' or m.ct_fld4 like '%總公司%')
             then 'Y'
             else 'N'
           end as Chg_Is_BU, -- 是否為總公司
           [Chg_ctno_Port_Office] = Rtrim(V1.Code_Name),
           [Chg_ctno_CustKind_CustCity] = Rtrim(Isnull(V2.Code_Name, V6.Code_Name)),
           [Chg_ctno_CustChain] = Rtrim(V3.Code_Name),
           [Chg_ct_dept_Name] = Rtrim(D1.DP_NAME), 
           [Chg_ct_sales_Name] = Rtrim(D2.E_NAME),
           [Chg_credit_Name] =Rtrim(D3.ct_sname),
           [Chg_busine_Name] = Rtrim(P4.Tn_Contact),
           [Chg_loc_Name] = Rtrim(P5.Tn_Contact),
           [Chg_sour_Name] = Rtrim(P6.Tn_Contact),
           [Chg_Customer_kind_Name] = Rtrim(P7.Tn_Contact),
           [Chg_payfg_Name] = Rtrim(V4.code_name),
           [Chg_porter_Name] = Rtrim(P8.tr_name),
           [Chg_pmode_Name] = Rtrim(V5.code_name),
           Chg_fld1_Year = substring(M.[ct_fld1], 1, 4),
           Chg_fld1_Month = substring(M.[ct_fld1], 6, 2), 
           Chg_fld2_Year = substring(M.[ct_fld2], 1, 4),
           Chg_fld2_Month = substring(M.[ct_fld2], 6, 2),
           Chg_Hunderd_Customer = -- 業務客戶百大名單
             Case
               when isnull(P9.Customer, '') <> '' then 'Y'
               else 'N'
             end,
           Chg_Hunderd_Customer_Name =
             Case
               when isnull(P9.Customer, '') <> '' then Rtrim(Customer)
               else ''
             end,
           -- 2014/1/27 Rick 增加判斷是否為網通客戶
           Case
             when ((Upper(substring(m.ct_no, 1, 2)) = 'I9' Or
                    Upper(substring(m.ct_no, 1, 5)) = 'I1826' Or
                    Upper(substring(m.ct_no, 1, 5)) = 'IZ000') and 
                    (m.ct_class = '1')
                  ) 
             then 'Y'
             else 'N'
           end as Chg_IS_Lan_Custom,
           -- 2014/10/09 Rickliu 增加判別有效客戶及廠商
           Case
             when (m.ct_name like '關店%' or m.ct_name like '不做%' or m.ct_name like '重覆不用%' or 
                   m.ct_name like '不用%' or m.ct_name like '停用%' or (LTrim(RTrim(m.ct_name)) = '') or
                   m.ct_name like '關帳%' or m.ct_name like '倒店%' or
                   Rtrim(replace(m.ct_fld2, '無', '')) <> '' or Rtrim(m.ct_name) = '' or Rtrim(replace(m.ct_fld3, '無', '')) = '' or
                   m.ct_no like 'IT%' or m.ct_no like 'IZ%' or m.ct_no like 'ZZ%'
                  ) -- 2017/04/28 無須判斷是否為客戶或廠商
             then 'Y'
             else ''
           end as Chg_ct_close,
           -- 2015/02/03 Rickliu 新增匯率
           Chg_rate_date = P10.r_date,
           Chg_rate = P10.r_rate,
           -- 2015/02/26 Rickliu 增加 客戶 收款類型, 取編號第九碼 1: 百貨, 2: 3C, 3: 寄賣, 4:代工(OEM)
           Case
             when m.ct_class = '1' and Len(m.ct_no) = 9 then Substring(RTrim(m.ct_no), 9, 1)
             when m.ct_class = '1' and Len(m.ct_no) = 5 then '0'
             else ''
           end as Chg_Cust_Sale_Class,
           Case
             when m.ct_class = '1' and Len(m.ct_no) = 5 then '廠商'
             when m.ct_class = '1' and Substring(RTrim(m.ct_no), 9, 1) = '1' then '百貨'
             when m.ct_class = '1' and Substring(RTrim(m.ct_no), 9, 1) = '2' then '3C'
             when m.ct_class = '1' and Substring(RTrim(m.ct_no), 9, 1) = '3' then '寄賣'
             when m.ct_class = '1' and Substring(RTrim(m.ct_no), 9, 1) = '4' then '代工(OEM)'
             else ''
           end as Chg_Cust_Sale_Class_Name,
           Case
             when m.ct_class = '1' and Len(m.ct_no) = 5 then RTrim(m.ct_sname)+'-廠商'
             when m.ct_class = '1' and Substring(RTrim(m.ct_no), 9, 1) = '1' then Replace(Replace(RTrim(m.ct_fld3), '@', ''), '#', '')+'-百貨'
             when m.ct_class = '1' and Substring(RTrim(m.ct_no), 9, 1) = '2' then Replace(Replace(RTrim(m.ct_fld3), '@', ''), '#', '')+'-3C'
             when m.ct_class = '1' and Substring(RTrim(m.ct_no), 9, 1) = '3' then Replace(Replace(RTrim(m.ct_fld3), '@', ''), '#', '')+'-寄賣'
             when m.ct_class = '1' and Substring(RTrim(m.ct_no), 9, 1) = '4' then Replace(Replace(RTrim(m.ct_fld3), '@', ''), '#', '')+'-代工(OEM)'
             else ''
           end as Chg_Cust_Sale_Class_sName,
           -- 2015/03/05 Rickliu 增加 客戶收款類型對應貨品折讓類別
           Case
             when m.ct_class = '1' and Substring(RTrim(m.ct_no), 9, 1) = '1' then 'A'
             when m.ct_class = '1' and Substring(RTrim(m.ct_no), 9, 1) = '2' then 'B'
             when m.ct_class = '1' and Substring(RTrim(m.ct_no), 9, 1) = '3' then 'C'
             when m.ct_class = '1' and Substring(RTrim(m.ct_no), 9, 1) = '4' then 'F'
             else ''
           end as Chg_Cust_Dis_Mapping,
           -- 2017/08/04 Rickliu 新增 廠商採購員或客戶業務 是否離職
           case when (Convert(Varchar(4), D2.e_ldate, 112) = '1900') then 'N' else 'Y' end as ct_sale_leave,
           pcust_update_datetime = getdate(),
           pcust_timestamp = m.timestamp_column
           into Fact_pcust
      from SYNC_TA13.dbo.PCust M
           left join SYNC_TA13.dbo.pdept D1 On M.ct_dept = D1.DP_NO -- 部門資料
           left join SYNC_TA13.dbo.Pemploy D2 on M.ct_sales = D2.E_NO -- 員工資料
           left join SYNC_TA13.dbo.PCust D3 On M.ct_class = D3.ct_class and M.ct_credit = D3.ct_no
           left join SYNC_TA13.dbo.pattn P4 On P4.tn_class='4' and M.ct_busine = P4.TN_NO  -- 客戶類別
           left join SYNC_TA13.dbo.pattn P5 On P5.tn_class='5' and M.ct_loc = P5.TN_NO -- 區域別
           left join SYNC_TA13.dbo.pattn P6 On P6.tn_class='6' and M.ct_sour = P6.TN_NO -- 客戶來源別
           left join SYNC_TA13.dbo.pattn P7 On P7.tn_class='7' and M.ct_kind = P7.TN_NO-- 客戶類別
           left join SYNC_TA13.dbo.struc P8 On M.ct_porter = P8.tr_no-- 貨運公司

           --2015/03/06 Rickliu 變更採用 Ori_Xls#Sys_Code 廠商第一碼公司別 及 客戶第一碼內外銷
           --left join V_Ctno_Port_Office V1 On M.ct_class = V1.ct_class and Substring(M.[ct_no], 1, 1)= V1.ctno_Port_Office_Kind-- 
           Left join Ori_Xls#Sys_Code V1 
                  On (V1.code_class ='6' and M.ct_class = '1' And Substring(M.[ct_no], 1, 1)= V1.Code_End Collate Chinese_Taiwan_Stroke_CI_AS)
                  Or (V1.code_class ='6' and M.ct_class = '2' And Substring(M.[ct_no], 1, 1)= V1.Code_End Collate Chinese_Taiwan_Stroke_CI_AS)


           Left join Ori_Xls#Sys_Code V2 
                  On (V2.Code_class = '1' And Len(m.ct_no) = 9 And 
                      Case
                        when Substring(M.[ct_no], 2, 4) like '[0-9][0-9][0-9][0-9]' then Convert(Int, Substring(M.[ct_no], 2, 4)) 
                        else null
                      end
                      between Convert(Int, V2.Code_Begin) and Convert(Int, V2.Code_End)
                     )
                  
                  
           --2015/03/06 Rickliu 變更採用 Ori_Xls#Sys_Code 客戶第九碼收款類別，商品折讓對應大分類別碼
           --left join V_Ctno_CustChain V3 On V3.ct_class = '1' and Substring(M.[ct_no], 9, 1) COLLATE Chinese_Taiwan_Stroke_CI_AS = V3.ctno_CustChain_Kind COLLATE Chinese_Taiwan_Stroke_CI_AS-- 
           Left join Ori_Xls#Sys_Code V3 On V3.code_class ='3' and Substring(M.[ct_no], 9, 1) = V3.Code_Begin Collate Chinese_Taiwan_Stroke_CI_AS
           
           --2015/03/06 Rickliu 變更採用 Ori_Xls#Sys_Code 售價及進價種類
           --left join V_Ctno_Payfg_Name V4 On M.ct_payfg = V4.ct_payfg
           left join Ori_Xls#Sys_Code V4 On V3.code_class ='4' and M.ct_payfg = V4.Code_Begin Collate Chinese_Taiwan_Stroke_CI_AS
            
           --2015/03/06 Rickliu 變更採用 Ori_Xls#Sys_Code 商品 請/付款天數方式
           --left join V_Ctno_Pmode_Name V5 On M.ct_pmode = V5.ct_pmode
           left join Ori_Xls#Sys_Code V5 On V3.code_class ='5' and M.ct_pmode = V5.Code_Begin Collate Chinese_Taiwan_Stroke_CI_AS
           left join ori_xls#Hunderd_Customer P9 on m.ct_class='1' -- 2013/2/2 與雅玉訪談後確認是使用簡稱去對應
                 and m.ct_name collate Chinese_PRC_BIN like '%'+P9.customer+'%' collate Chinese_PRC_BIN

           Left join Ori_Xls#Sys_Code V6 
                  On (V6.Code_class = '24' And Len(m.ct_no) = 9 And 
                      Case
                        when Substring(M.[ct_no], 2, 1) like '[A-Z]' then Substring(M.[ct_no], 2, 1) Collate Chinese_Taiwan_Stroke_CI_AS
                        else null
                      end
                      between V6.Code_Begin Collate Chinese_Taiwan_Stroke_CI_AS and V6.Code_End Collate Chinese_Taiwan_Stroke_CI_AS
                     )

		   --2015/07/01 NanLiao改寫邏輯 取得唯一值
           --left join (select r_name, r_rate, max(r_date) as r_date
           --             from SYNC_TA13.dbo.prate
           --            group by r_name, r_rate
           left join CTE_Q2 P10 
             on Case
                  when RTrim(Isnull(M.ct_curt_id, '')) = '' then 'NT'
                  else M.ct_curt_id 
                end = P10.r_name


    /*==============================================================*/
    /* Index: pcust_Timestamp                                       */
    /*==============================================================*/
    --set @Msg = '建立索引 [Fact_pcust.pcust_Timestamp]'
    --Print @Msg
    --if exists (select * from sys.indexes where object_id = object_id('[dbo].[fact_pcust]') and name = 'pk_pcust')
    --   alter table [dbo].[pk_pcust] drop constraint [pk_pcust]

    --alter table [dbo].[fact_pcust] add  constraint [pk_pcust] primary key nonclustered ([pcust_timestamp] asc) with 
    --(pad_index  = off, statistics_norecompute  = off, sort_in_tempdb = off, ignore_dup_key = off, online = off, allow_row_locks  = on, allow_page_locks  = on) on [primary]

    /*==============================================================*/
    /* Index: no                                                    */
    /*==============================================================*/
    set @Msg = '建立索引 [Fact_pcust.no]'
    Print @Msg
    if exists (select 1 from sysindexes where id = object_id('dbo.fact_pcust') and name  = 'no' and indid > 0 and indid < 255) drop index dbo.fact_pcust.no
    create clustered index no on dbo.fact_pcust (ct_no, ct_class) with fillfactor= 30 on "PRIMARY"

    /*==============================================================*/
    /* Index: class                                                 */
    /*==============================================================*/
    set @Msg = '建立索引 [Fact_pcust.Class]'
    Print @Msg
    if exists (select 1 from sysindexes where id = object_id('dbo.fact_pcust') and name  = 'class' and indid > 0 and indid < 255) drop index dbo.fact_pcust.class
    create index class on dbo.fact_pcust (ct_class) with fillfactor= 30 on "PRIMARY"

    /*==============================================================*/
    /* Index: ct_payrate                                            */
    /*==============================================================*/
    set @Msg = '建立索引 [Fact_pcust.ct_payrate]'
    Print @Msg
    if exists (select 1 from sysindexes where id = object_id('dbo.fact_pcust') and name  = 'ct_payrate' and indid > 0 and indid < 255) drop index dbo.fact_pcust.ct_payrate
    create index ct_payrate on dbo.fact_pcust (ct_payer, ct_class) with fillfactor= 30 on "PRIMARY"

    /*==============================================================*/
    /* Index: ctbus                                                 */
    /*==============================================================*/
    set @Msg = '建立索引 [Fact_pcust.ctbus]'
    Print @Msg
    if exists (select 1 from sysindexes where id = object_id('dbo.fact_pcust') and name  = 'ctbus' and indid > 0  and indid < 255) drop index dbo.fact_pcust.ctbus
    create index ctbus on dbo.fact_pcust (ct_busine, ct_class) with fillfactor= 30 on "PRIMARY"

    /*==============================================================*/
    /* Index: ctkind                                                */
    /*==============================================================*/
    set @Msg = '建立索引 [Fact_pcust.ctkind]'
    Print @Msg
    if exists (select 1 from sysindexes where id = object_id('dbo.fact_pcust') and name  = 'ctkind' and indid > 0 and indid < 255) drop index dbo.fact_pcust.ctkind
    create index ctkind on dbo.fact_pcust (ct_kind, ct_class) with fillfactor= 30 on "PRIMARY"

    /*==============================================================*/
    /* Index: ctloc                                                 */
    /*==============================================================*/
    set @Msg = '建立索引 [Fact_pcust.ctloc]'
    Print @Msg
    if exists (select 1 from sysindexes where id = object_id('dbo.fact_pcust') and name  = 'ctloc' and indid > 0 and indid < 255) drop index dbo.fact_pcust.ctloc
    create index ctloc on dbo.fact_pcust (ct_loc, ct_class) with fillfactor= 30 on "PRIMARY"

    /*==============================================================*/
    /* Index: ctname                                                */
    /*==============================================================*/
    set @Msg = '建立索引 [Fact_pcust.ctname]'
    Print @Msg
    if exists (select 1 from sysindexes where id = object_id('dbo.fact_pcust') and name = 'ctname' and indid > 0 and indid < 255) drop index dbo.fact_pcust.ctname
    create index ctname on dbo.fact_pcust (ct_name) with fillfactor= 30 on "PRIMARY"

    /*==============================================================*/
    /* Index: sname                                                 */
    /*==============================================================*/
    set @Msg = '建立索引 [Fact_pcust.sname]'
    Print @Msg
    if exists (select 1 from sysindexes where id = object_id('dbo.fact_pcust') and name  = 'sname' and indid > 0 and indid < 255) drop index dbo.fact_pcust.sname
    create index sname on dbo.fact_pcust (ct_sname, ct_class) with fillfactor= 30 on "PRIMARY"

    /*==============================================================*/
    /*  資料檢核                                                    */
    /*==============================================================*/
    set @Msg =(select ct_no+'('+ct_fld3+'),'
                 from (select chg_bu_no, count(distinct ct_fld3) as cnt
                         from fact_pcust  
                        where chg_bu_no <> ''
                        group by chg_bu_no
                       having count(distinct ct_fld3) > 1
                      ) m
                      inner join fact_pcust d
                         on m.chg_bu_no = d.chg_bu_no
                  for xml path('')
              )
    set @Msg = @Proc+'客戶總公司不一致：'+Reverse(Substring(Reverse(@Msg), 2, Len(@Msg)))
    if @Msg is not null
       Exec uSP_Sys_Write_Log @Proc, @Msg, @Msg, 0, 1      
      
  end try
  begin catch
    set @Result = -1
    set @Msg = @Msg+'...(錯誤訊息:'+ERROR_MESSAGE()+')...(錯誤列:'+Convert(Varchar(10), ERROR_LINE())+')'

    Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Result
  end catch

  Return(@Result)
end
GO
