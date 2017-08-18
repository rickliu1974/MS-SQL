USE [DW]
GO
/****** Object:  StoredProcedure [dbo].[xuSP_ETL_PCust_Test(201708停用)]    Script Date: 08/18/2017 17:18:56 ******/
DROP PROCEDURE [dbo].[xuSP_ETL_PCust_Test(201708停用)]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE Proc [dbo].[xuSP_ETL_PCust_Test(201708停用)] (@Kind Varchar(1)='', @Value Timestamp = Null)
as
begin
  /*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
   Program Name: uSP_TA13#ETL_PCust
   Create Date: 2013/03/01
   Creator: Rickliu
   Updated Date: 2013/09/06 [修改 Trans_Log 訊息內容]
  *-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*/  

  Declare @Proc Varchar(50) = 'uSP_ETL_PCust_Test'
  Declare @Cnt Int =0
  Declare @Cnt_Ori Int =0
  Declare @Msg Varchar(Max) =''
  Declare @strSQL Varchar(Max)=''
  Declare @strSQL_Inx Varchar(2000) = ''

  Declare @CR Varchar(4) = char(13)+char(10)
  Declare @subject Varchar(500)= 'Exec '+@Proc+' Error!!'
  Declare @Result int = 0
  
  Declare @Chk_Tb_Exists Varchar(1) = 'N'
  Declare @Owner_Name Varchar(50) = 'dbo.'
  Declare @Tb_Name Varchar(50) = 'Fact_PCust_Test'
  Declare @Tb_tmp_Name Varchar(50) = @Tb_Name+'_tmp'

  Declare @Inx_Name Varchar(50) = ''
  Declare @Inx_Columns Varchar(1000) = ''
  Declare @Inx_clustered Varchar(50) = ''
  
  -- Timestamp 可以先轉換成 Int 後再轉回 Timestamp，文字雖可轉 Timestamp，但數值會不正確。
  Declare @sTimestamp Varchar(1000) = Isnull(Convert(Varchar(1000), Convert(Int, @Value)), '')

  Begin try
 print '1'
    set @strSQL_Inx = 
        'if exists (select 1 from sysindexes where id = object_id(''@Tb_Name'') and name  = ''@Inx_Name'' and indid > 0 and indid < 255) drop index @Tb_Name.@Inx_Name '+@CR+
        '   create @Inx_clustered index @Inx_Name on @Tb_Name (@Inx_Columns) with fillfactor= 30 on "PRIMARY" '
  
 print '2'
    If @Kind Not In ('I', 'U', 'D', '') raiserror ('@Kind 參數必須為 I, U, D, 空白', 50005, 10, 1)

 print '3'
    IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].['+@Tb_name+']') AND type in (N'U'))
       Set @Chk_Tb_Exists = 'Y'

    IF (@Kind = '' And @Value = 0 )
    begin
       Set @Msg = '清空資料表 ['+@Tb_name+']'
       set @strSQL= 'Truncate Table '+@Tb_name

       Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
    end
  
    If (@Kind In ('U', 'D') And @Value Is Not Null)
    begin
       set @strSQL= 'Delete '+@Tb_Name+' where pcust_Timestamp = Convert(Timestamp, '+@sTimestamp+')'
       Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
    end

    set @Msg = 'ETL PCust to [Fact_PCust]...'
    Exec uSP_Sys_Write_Log @Proc, @Msg, @Msg, 0

    -- CTE Query Begin
    set @strSQL = ';With CTE_Q1 as ( '+@CR+
                  '  select r_name, max(r_date) as r_date '+@CR+
                  '    from SYNC_TA13.dbo.prate '+@CR+
                  '   group by r_name '+@CR+
                  '), CTE_Q2 as ( '+@CR+
                  '  select a.r_name, a.r_rate, a.r_date as r_date '+@CR+
                  '    from SYNC_TA13.dbo.prate as a '+@CR+
                  '         left join CTE_Q1 as b '+@CR+
                  '           on a.r_name=b.r_name and a.r_date=b.r_date '+@CR+
                  '   where a.r_name=b.r_name and a.r_date=b.r_date '+@CR+
                  '), CTE_Q3 as ( '+@CR+
                  '  Select * '+@CR+
                  '    from Sync_TA13.dbo.'+@Tb_Name

    If (@Kind In ('I', 'U', 'D') And @Value Is Not Null)
       set @strSQL = @strSQL +@CR+
                     '   where Timestamp_colum = Convert(Timestamp, '+@sTimestamp+') '
    set @strSQL = @strSQL+@CR+')'
    -- CTE Query End

    -- 如果資料表存在，則就直接 Insert            
    If (@Chk_Tb_Exists = 'Y')
       set @strSQL = @strSQL +@CR+
                  'Insert into '+@Tb_tmp_name
   
    set @strSQL = @strSQL +@CR+
                  -- 類別, 編號, 名稱, 簡稱
                  'select M.[ct_class], Rtrim(M.[ct_no]) as ct_no, Rtrim(M.[ct_name]) as ct_name, Rtrim(M.[ct_sname]) as ct_sname, '+@CR+
                  -- 客戶八碼編號
                  '       Case when m.ct_class = ''1'' and Len(Rtrim(M.[ct_no])) > 8 and Substring(M.ct_no, 1, 2) Not in (''IT'', ''IZ'') then Substring(Rtrim(M.[ct_no]), 1, 8) else '''' end as ct_no8, '+@CR+ 
                  -- 客戶八碼名稱
                  '       LTrim(RTrim(Replace(Replace(Replace(Replace(Replace(Replace(m.ct_sname, ''#'', ''''), ''@'', ''''), ''-'', ''''), ''T'', ''''), ''P'', ''''), ''/'', ''''))) as ct_sname8, '+@CR+
                  -- 公司地址, 發票地址, 送貨地址
                  '       [ct_addr1]=Rtrim(Convert(Varchar(255), M.ct_addr1)), [ct_addr2]=Rtrim(Convert(Varchar(255), M.ct_addr2)),  [ct_addr3]=Rtrim(Convert(Varchar(255), M.ct_addr3)), '+@CR+
                  -- 電話, 傳真, 統一編號, 負責人, 聯絡人
                  '       Rtrim(M.[ct_tel]) as ct_tel, Rtrim(M.[ct_fax]) as ct_fax, Rtrim(M.[ct_unino]) as ct_unino, Rtrim(M.[ct_presid]) as ct_presid, Rtrim(M.[ct_contact]) as ct_contact, '+@CR+
                  -- 售價種類, 業務員, 帳款額度, 信用額度, 銀行帳號, 銀行名稱
                  '       M.[ct_payfg], Rtrim(M.[ct_sales]) as ct_sales, M.[ct_p_limit], M.[ct_b_limit], Rtrim(M.[ct_bkno]) as ct_bkno, Rtrim(M.[ct_bknm]) as ct_bknm, '+@CR+
                  -- 備註, 部門編號, 售價折數, 發票抬頭, 貨運公司
                  '       [ct_rem]=Rtrim(Convert(Varchar(255), M.ct_rem)), Rtrim(M.[ct_dept]) as ct_dept,  M.[ct_payrate], Rtrim(M.[ct_ivtitle]) as ct_ivtitle, Rtrim(M.[ct_porter]) as ct_porter, '+@CR+
                  -- 請款|付款對象, 付款方式, 付款天數, 未清款, 預收款, 最近交易日, 客戶兼廠商, 傳輸旗標, 等級, 地區,國別代碼
                  '       Rtrim(M.[ct_credit]) as ct_credit, M.[ct_pmode], M.[ct_pdate], M.[ct_prenpay], M.[ct_prepay], M.[ct_last_dt], M.[ct_flg], M.[ct_t_fg], M.[ct_grade], M.[ct_area], '+@CR+
                  -- 2015/02/03 交易幣別
                  '       case when RTrim(isnull(M.ct_curt_id, '''')) = '''' then ''NT'' else RTrim(isnull(M.ct_curt_id, '''')) end as ct_curt_id, '+@CR+
                  -- 聯絡人職稱, 貿易付款方式, 工廠登記証, 員工數, 資本額, 貿易預收款, 貿易未清款, 貿易應付帳科目, 貿易應付票科目
                  '       Rtrim(M.[ct_cont_sp]) as ct_cont_sp, M.[ct_pay], Rtrim(M.[ct_regist]) as ct_regist, M.[ct_worker],  M.[ct_capital], M.[ct_skpay], M.[ct_sknpay], M.[ct_accno2], M.[ct_chkno2], '+@CR+
                  -- 建檔日期, 進銷存請款對象, 進銷存預收款, 進銷存未清款, 國內.國外
                  '       M.[ct_cdate], M.[ct_payer], M.[ct_advance], M.[ct_debt], M.[ct_abroad], '+@CR+
                  -- 客戶|廠商附註一(客：開發日期、廠：配合起始日), 客戶|廠商附註二(客：關店日期、廠：結束配合日), 客戶|廠商附註三(客：總公司別、廠：未使用), 客戶|廠商附註四(客：分店別、廠：廠商評比), 客戶|廠商附註五(客：客服專用1、廠：簽約到期日), 客戶|廠商附註六(客：客服專用2、廠：未使用)
                  '       Rtrim(M.[ct_fld1]) as ct_fld1, Rtrim(M.[ct_fld2]) as ct_fld2, Rtrim(M.[ct_fld3]) as ct_fld3, Rtrim(M.[ct_fld4]) as ct_fld4, Rtrim(M.[ct_fld5]) as ct_fld5, Rtrim(M.[ct_fld6]) as ct_fld6, '+@CR+
                  -- 客戶|廠商附註七(客：客服專用3、廠：未使用), 客戶|廠商附註八(客：客服專用4、廠：未使用), 客戶|廠商附註九(客：收款方式、廠：未使用), 客戶|廠商附註十(客：發票名稱、廠：未使用), 客戶|廠商附註十一(客：付款條件、廠：未使用), 客戶|廠商附註十二(客：舊編號 2012開帳使用)
                  '       Rtrim(M.[ct_fld7]) as ct_fld7, Rtrim(M.[ct_fld8]) as ct_fld8, Rtrim(M.[ct_fld9]) as ct_fld9, Rtrim(M.[ct_fld10]) as ct_fld10, Rtrim(M.[ct_fld11]) as ct_fld11, Rtrim(M.[ct_fld12]) as ct_fld12, '+@CR+
                  --貿易運送方式, 開立發票種類, 交易單價的小數位數, 交易金額的小數位數, 行業別, 押匯銀行, 區域編號, 客戶來源, 客戶類別, 拜訪週期
                  '       M.[ct_sea], M.[ct_invofg], M.[ct_udec], M.[ct_tdec], M.[ct_busine], Rtrim(M.[ct_banpay]) as ct_banpay, M.[ct_loc], M.[ct_sour], M.[ct_kind], M.[ct_tkday], '+@CR+
                  '       Chg_ctclass = Case when M.ct_class = ''1'' and Len(m.ct_no) = 9 then ''客戶'' when M.ct_class = ''2'' and Len(m.ct_no) = 5 then ''廠商'' end, '+@CR+
                  -- 客戶
                  -- 2016/09/09 Rickliu 因組織擴編，增加機車通路，因此將原本 2~5 改為 2~6
                  -- Chg_BU_NO = Substring(Rtrim(M.[ct_no]), 1, 5), -- 總公司編號
                  -- 總公司編號
                  '       Chg_BU_NO = Substring(Rtrim(M.[ct_no]), 1, 6), Chg_ctno_Port_Office = Rtrim(V1.Code_Name), Chg_ctno_CustKind_CustCity = Rtrim(V2.Code_Name), Chg_ctno_CustChain = Rtrim(V3.Code_Name), '+@CR+
                  '       Chg_ct_dept_Name = Rtrim(D1.DP_NAME), Chg_ct_sales_Name = Rtrim(D2.E_NAME), Chg_credit_Name = Rtrim(D3.ct_sname), Chg_busine_Name = Rtrim(P4.Tn_Contact), Chg_loc_Name = Rtrim(P5.Tn_Contact), '+@CR+
                  '       Chg_sour_Name = Rtrim(P6.Tn_Contact), Chg_Customer_kind_Name = Rtrim(P7.Tn_Contact), Chg_payfg_Name = Rtrim(V4.code_name), Chg_porter_Name = Rtrim(P8.tr_name), Chg_pmode_Name = Rtrim(V5.code_name),  '+@CR+
                  '       Chg_fld1_Year = substring(M.[ct_fld1], 1, 4), Chg_fld1_Month = substring(M.[ct_fld1], 6, 2), Chg_fld2_Year = substring(M.[ct_fld2], 1, 4), Chg_fld2_Month = substring(M.[ct_fld2], 6, 2),  '+@CR+
                  -- 業務客戶百大名單
                  '       Chg_Hunderd_Customer = Case when isnull(P9.Customer, '''') <> '''' then ''Y'' else ''N'' end, Chg_Hunderd_Customer_Name = Case when isnull(P9.Customer, '''') <> '''' then Rtrim(Customer) else '''' end, '+@CR+
                  -- 2014/1/27 Rick 增加判斷是否為網通客戶
                  '       Case when ((Upper(substring(m.ct_no, 1, 2)) = ''I9'' Or Upper(substring(m.ct_no, 1, 5)) = ''I1826'' Or  Upper(substring(m.ct_no, 1, 5)) = ''IZ000'') and (m.ct_class = ''1'')) then ''Y'' else ''N'' end as Chg_IS_Lan_Custom, '+@CR+
                  -- 2014/10/09 Rickliu 增加判別有效客戶及廠商
                  '       Case when (LTrim(RTrim(m.ct_name)) = '''' or m.ct_name like ''%關店%'' or m.ct_name like ''不做%'' or m.ct_name like ''重覆不用%'' or m.ct_name like ''不用%'' or m.ct_name like ''停用%'' or '+@CR+
                  '            (m.ct_class = ''1'' and Rtrim(m.ct_fld2) <> '''' ) Or (m.ct_class = ''1'' and substring(m.ct_no, 1, 1) Not in (''I'', ''E'', ''B'')) Or '+@CR+
                  '            (m.ct_class = ''1'' and substring(m.ct_no, 2, 1) in (''P'', ''A'', ''T'', ''Z''))) '+@CR+
                  '            then ''Y'' else '''' end as Chg_ct_close, '+@CR+
                  -- 2015/02/03 Rickliu 新增匯率
                  '       Chg_rate_date = P10.r_date, Chg_rate = P10.r_rate, '+@CR+
                  -- 2015/02/26 Rickliu 增加 客戶 收款類型, 取編號第九碼 1: 百貨, 2: 3C, 3: 寄賣, 4:代工(OEM)
                  '       Case when m.ct_class = ''1'' and Len(m.ct_no) = 9 then Substring(RTrim(m.ct_no), 9, 1) when m.ct_class = ''1'' and Len(m.ct_no) = 5 then ''0'' else '''' end as Chg_Cust_Sale_Class, '+@CR+
                  '       Case when m.ct_class = ''1'' and Len(m.ct_no) = 5 then ''廠商'' when m.ct_class = ''1'' and Substring(RTrim(m.ct_no), 9, 1) = ''1'' then ''百貨'' '+@CR+
                  '            when m.ct_class = ''1'' and Substring(RTrim(m.ct_no), 9, 1) = ''2'' then ''3C'' when m.ct_class = ''1'' and Substring(RTrim(m.ct_no), 9, 1) = ''3'' then ''寄賣'' '+@CR+
                  '            when m.ct_class = ''1'' and Substring(RTrim(m.ct_no), 9, 1) = ''4'' then ''代工(OEM)'' else '''' end as Chg_Cust_Sale_Class_Name, '+@CR+
                  '       Case when m.ct_class = ''1'' and Len(m.ct_no) = 5 then RTrim(m.ct_sname)+''-廠商'' '+@CR+
                  '            when m.ct_class = ''1'' and Substring(RTrim(m.ct_no), 9, 1) = ''1'' then Replace(Replace(RTrim(m.ct_fld3), ''@'', ''), ''#'', '''')+''-百貨'' '+@CR+
                  '            when m.ct_class = ''1'' and Substring(RTrim(m.ct_no), 9, 1) = ''2'' then Replace(Replace(RTrim(m.ct_fld3), ''@'', ''), ''#'', '''')+''-3C'' '+@CR+
                  '            when m.ct_class = ''1'' and Substring(RTrim(m.ct_no), 9, 1) = ''3'' then Replace(Replace(RTrim(m.ct_fld3), ''@'', ''), ''#'', '''')+''-寄賣'' '+@CR+
                  '            when m.ct_class = ''1'' and Substring(RTrim(m.ct_no), 9, 1) = ''4'' then Replace(Replace(RTrim(m.ct_fld3), ''@'', ''), ''#'', '''')+''-代工(OEM)'' '+@CR+
                  '            else '''' end as Chg_Cust_Sale_Class_sName, '+@CR+
                  -- 2015/03/05 Rickliu 增加 客戶收款類型對應貨品折讓類別
                  '       Case when m.ct_class = ''1'' and Substring(RTrim(m.ct_no), 9, 1) = ''1'' then ''A'' when m.ct_class = ''1'' and Substring(RTrim(m.ct_no), 9, 1) = ''2'' then ''B'' '+@CR+
                  '            when m.ct_class = ''1'' and Substring(RTrim(m.ct_no), 9, 1) = ''3'' then ''C'' when m.ct_class = ''1'' and Substring(RTrim(m.ct_no), 9, 1) = ''4'' then ''F'' '+@CR+
                  '            else '''' end as Chg_Cust_Dis_Mapping, '+@CR+
                  '       pcust_update_datetime = getdate(), pcust_timestamp = m.timestamp_column '
    If (@Chk_Tb_Exists = 'N')
       set @strSQL = @strSQL +@CR+
                  '       Into '+@Tb_name

    set @strSQL = @strSQL +@CR+
                  '  From CTE_Q3 m '+@CR+
                  -- 部門資料
                  '       left join SYNC_TA13.dbo.pdept D1 On M.ct_dept = D1.DP_NO '+@CR+
                  -- 員工資料
                  '       left join SYNC_TA13.dbo.Pemploy D2 on M.ct_sales = D2.E_NO '+@CR+
                  -- 客戶類別
                  '       left join SYNC_TA13.dbo.PCust D3 On M.ct_class = D3.ct_class and M.ct_credit = D3.ct_no '+@CR+
                  -- 區域別
                  '       left join SYNC_TA13.dbo.pattn P4 On P4.tn_class=''4'' and M.ct_busine = P4.TN_NO '+@CR+
                  -- 客戶來源別
                  '       left join SYNC_TA13.dbo.pattn P5 On P5.tn_class=''5'' and M.ct_loc = P5.TN_NO '+@CR+
                  -- 客戶類別 
                  '       left join SYNC_TA13.dbo.pattn P6 On P6.tn_class=''6'' and M.ct_sour = P6.TN_NO '+@CR+
                  -- 貨運公司
                  '       left join SYNC_TA13.dbo.pattn P7 On P7.tn_class=''7'' and M.ct_kind = P7.TN_NO '+@CR+
                  '       left join SYNC_TA13.dbo.struc P8 On M.ct_porter = P8.tr_no '+@CR+
                  -- 2015/03/06 Rickliu 變更採用 Ori_Xls#Sys_Code 廠商第一碼公司別 及 客戶第一碼內外銷
                  '       left join Ori_Xls#Sys_Code V1 '+@CR+
                  '         On (V1.code_class =''6'' and M.ct_class = ''1'' and V1.Code_Class = ''1'' And Substring(M.[ct_no], 1, 1)= V1.Code_End Collate Chinese_Taiwan_Stroke_CI_AS) '+@CR+
                  '         Or (V1.code_class =''6'' and M.ct_class = ''2'' and V1.Code_Class = ''2'' And Substring(M.[ct_no], 1, 1)= V1.Code_End Collate Chinese_Taiwan_Stroke_CI_AS) '+@CR+
                  '       left join Ori_Xls#Sys_Code V2 '+@CR+
                  -- 2016/09/09 Rickliu 因組織擴編，增加機車通路，因此將原本 2~5 改為 2~6
                  '        On (V2.Code_class = ''1'' and M.ct_class = ''1'' AND Substring(M.[ct_no], 2, 6) COLLATE Chinese_Taiwan_Stroke_CI_AS '+@CR+ 
                  '           between V2.Code_Begin COLLATE Chinese_Taiwan_Stroke_CI_AS and V2.Code_End COLLATE Chinese_Taiwan_Stroke_CI_AS) Or '+@CR+
                  '           (V2.Code_class = ''1'' and M.ct_class = ''2'' And Substring(M.[ct_no], 2, 1) COLLATE Chinese_Taiwan_Stroke_CI_AS '+@CR+
                  '           between V2.Code_Begin COLLATE Chinese_Taiwan_Stroke_CI_AS and V2.Code_End COLLATE Chinese_Taiwan_Stroke_CI_AS) '+@CR+
                  -- 2015/03/06 Rickliu 變更採用 Ori_Xls#Sys_Code 客戶第九碼收款類別，商品折讓對應大分類別碼
                  '       left join Ori_Xls#Sys_Code V3 On V3.code_class =''3'' and Substring(M.[ct_no], 9, 1) = V3.Code_Begin Collate Chinese_Taiwan_Stroke_CI_AS'+@CR+
                  -- 2015/03/06 Rickliu 變更採用 Ori_Xls#Sys_Code 售價及進價種類
                  '       left join Ori_Xls#Sys_Code V4 On V3.code_class =''4'' and M.ct_payfg = V4.Code_Begin Collate Chinese_Taiwan_Stroke_CI_AS'+@CR+
                  -- 2015/03/06 Rickliu 變更採用 Ori_Xls#Sys_Code 商品 請/付款天數方式
                  '       left join Ori_Xls#Sys_Code V5 On V3.code_class =''5'' and M.ct_pmode = V5.Code_Begin Collate Chinese_Taiwan_Stroke_CI_AS'+@CR+
                  -- 2013/2/2 與雅玉訪談後確認是使用簡稱去對應
                  '       left join ori_xls#Hunderd_Customer P9 on m.ct_class=''1'' and m.ct_name collate Chinese_PRC_BIN like ''%''+P9.customer+''%'' collate Chinese_PRC_BIN '+@CR+
                  -- 2015/07/01 NanLiao改寫邏輯 取得唯一值
                  '       left join CTE_Q2 P10 on Case when RTrim(Isnull(M.ct_curt_id, '')) = '''' then ''NT'' else M.ct_curt_id end = P10.r_name '
     Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL

    -- 2017/04/06 Rickliu 索引檔不在則重建
    -- Index: no
    set @Inx_Name = 'no'
    set @Inx_Columns = 'ct_no, ct_class'
    set @Inx_clustered = 'clustered'
    Set @Msg = '檢核索引檔 ['+@Tb_Name+'.'+@Inx_Name+'('+@Inx_Columns+')]'
    set @strSQL= Replace(Replace(Replace(Replace(@strSQL_Inx, '@Tb_Name', @Tb_Name), '@Inx_Name', @Inx_Name), '@Inx_Columns', @Inx_Columns), '@Inx_clustered', @Inx_clustered)
    Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
    
    -- Index: class
    set @Inx_Name = 'class'
    set @Inx_Columns = 'ct_class'
    set @Inx_clustered = ''
    Set @Msg = '檢核索引檔 ['+@Tb_Name+'.'+@Inx_Name+'('+@Inx_Columns+')]'
    set @strSQL= Replace(Replace(Replace(Replace(@strSQL_Inx, '@Tb_Name', @Tb_Name), '@Inx_Name', @Inx_Name), '@Inx_Columns', @Inx_Columns), '@Inx_clustered', @Inx_clustered)
    Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
  
    -- Index: ct_payrate
    set @Inx_Name = 'ct_payrate'
    set @Inx_Columns = 'ct_payer, ct_class'
    set @Inx_clustered = ''
    Set @Msg = '檢核索引檔 ['+@Tb_Name+'.'+@Inx_Name+'('+@Inx_Columns+')]'
    set @strSQL= Replace(Replace(Replace(Replace(@strSQL_Inx, '@Tb_Name', @Tb_Name), '@Inx_Name', @Inx_Name), '@Inx_Columns', @Inx_Columns), '@Inx_clustered', @Inx_clustered)
    Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
  
    -- Index: ctbus
    set @Inx_Name = 'ctbus'
    set @Inx_Columns = 'ct_busine, ct_class'
    set @Inx_clustered = ''
    Set @Msg = '檢核索引檔 ['+@Tb_Name+'.'+@Inx_Name+'('+@Inx_Columns+')]'
    set @strSQL= Replace(Replace(Replace(Replace(@strSQL_Inx, '@Tb_Name', @Tb_Name), '@Inx_Name', @Inx_Name), '@Inx_Columns', @Inx_Columns), '@Inx_clustered', @Inx_clustered)
    Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
  
    -- Index: ctkind
    set @Inx_Name = 'ctkind'
    set @Inx_Columns = 'ct_kind, ct_class'
    set @Inx_clustered = ''
    Set @Msg = '檢核索引檔 ['+@Tb_Name+'.'+@Inx_Name+'('+@Inx_Columns+')]'
    set @strSQL= Replace(Replace(Replace(Replace(@strSQL_Inx, '@Tb_Name', @Tb_Name), '@Inx_Name', @Inx_Name), '@Inx_Columns', @Inx_Columns), '@Inx_clustered', @Inx_clustered)
    Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
  
    -- Index: ctloc
    set @Inx_Name = 'ctloc'
    set @Inx_Columns = 'ct_loc, ct_class'
    set @Inx_clustered = ''
    Set @Msg = '檢核索引檔 ['+@Tb_Name+'.'+@Inx_Name+'('+@Inx_Columns+')]'
    set @strSQL= Replace(Replace(Replace(Replace(@strSQL_Inx, '@Tb_Name', @Tb_Name), '@Inx_Name', @Inx_Name), '@Inx_Columns', @Inx_Columns), '@clustered', @Inx_clustered)
    Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
  
    -- Index: ctname
    set @Inx_Name = 'ctname'
    set @Inx_Columns = 'ct_name'
    set @Inx_clustered = ''
    Set @Msg = '檢核索引檔 ['+@Tb_Name+'.'+@Inx_Name+'('+@Inx_Columns+')]'
    set @strSQL= Replace(Replace(Replace(Replace(@strSQL_Inx, '@Tb_Name', @Tb_Name), '@Inx_Name', @Inx_Name), '@Inx_Columns', @Inx_Columns), '@Inx_clustered', @Inx_clustered)
    Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
  
    -- Index: sname
    set @Inx_Name = 'sname'
    set @Inx_Columns = 'ct_sname, ct_class'
    set @Inx_clustered = ''
    Set @Msg = '檢核索引檔 ['+@Tb_Name+'.'+@Inx_Name+'('+@Inx_Columns+')]'
    set @strSQL= Replace(Replace(Replace(Replace(@strSQL_Inx, '@Tb_Name', @Tb_Name), '@Inx_Name', @Inx_Name), '@Inx_Columns', @Inx_Columns), '@Inx_clustered', @Inx_clustered)
    Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
  
  end try
  begin catch
    set @Result = -1
    set @Msg = @Msg+'(錯誤訊息:'+ERROR_MESSAGE()+')'

    Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Result
  end catch

  Return(@Result)
end
GO
