USE [DW]
GO
/****** Object:  StoredProcedure [dbo].[zNew_uSP_ETL_sstock]    Script Date: 07/24/2017 14:44:00 ******/
DROP PROCEDURE [dbo].[zNew_uSP_ETL_sstock]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
Create Proc [dbo].[zNew_uSP_ETL_sstock](@sDB Varchar(10) = 'TA13')
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
  Declare @RowCount table (cnt int)
  Declare @Cnt_Ori Int =0
  Declare @Msg Varchar(Max) =''
  Declare @strSQL Varchar(Max)=''
  Declare @subject Varchar(500)=''
  Declare @CR Varchar(4) = char(13)+char(10)
  Declare @Tb_Name Varchar(20) = 'Fact_sstock'
  Declare @Result Int = 0
  
  Begin Try
    IF Not EXISTS (SELECT * FROM sys.databases WHERE name = 'SYNC_'+@sDB)
    begin
       set @Msg = '所選的資料庫不存在 ['+@sDB+']....'
       Raiserror(@Msg, 16, 1)
    end

    set @Tb_Name = @sDB+'#'+@Tb_Name

    IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].['+@Tb_Name+']') AND type in (N'U'))
    begin
       Set @Msg = '刪除資料表 ['+@Tb_Name+']'
       set @strSQL= 'DROP TABLE [dbo].['+@Tb_Name+']'

       Exec SP_Exec_SQL @Proc, @Msg, @strSQL
    end

    set @Msg = '建立 ['+@Tb_Name+'] 資料表...'

    set @strSQL = 'select distinct '+@CR+
                  '       Rtrim(M.[sk_no]) as sk_no, '+@CR+
                  '       Rtrim(M.[sk_bcode]) as sk_bcode, '+@CR+
                  '       Rtrim(M.[sk_name]) as sk_name, M.[sk_kind], M.[sk_ivkd], '+@CR+
                  '       Rtrim(M.[sk_spec]) as sk_spec, '+@CR+
                  '       Rtrim(M.[sk_unit]) as sk_unit, '+@CR+
                  '       Rtrim(M.[sk_color]) as sk_color, '+@CR+
                  '       Rtrim(M.[sk_use]) as sk_use, M.[sk_aux_q], '+@CR+
                  '       M.[sk_aux_u], '+@CR+
                  '       Rtrim(M.[sk_size]) as sk_size, '+@CR+
                  '       M.[sk_save], M.[sk_bqty], '+@CR+
                  '       [sk_rem]=Rtrim(Convert(Varchar(255), M.[sk_rem])), '+@CR+
                  '       M.[sk_fign],  M.[sk_tfg], '+@CR+
                  '       [st_cspes]=Rtrim(Convert(Varchar(255), M.[st_cspes])), '+@CR+
                  '       [st_espes]=Rtrim(Convert(Varchar(255), M.[st_espes])), M.[st_unit], '+@CR+
                  '       M.[st_inerqty], M.[st_inerunt], M.[st_outqty], M.[st_outunt], '+@CR+
                  '       M.[st_appr], M.[st_apprunt], M.[st_nw], M.[st_gw], M.[st_gwunt], '+@CR+
                  '       M.[st_itclass], M.[st_cccode], M.[st_sespes],  M.[st_20cyqty], M.[st_40cyqty], '+@CR+
                  '       M.[st_45cyqty], '+@CR+
                  '       Rtrim(M.[sk_fld1]) as sk_fld1, '+@CR+
                  '       Rtrim(M.[sk_fld2]) as sk_fld2, '+@CR+
                  '       Rtrim(M.[sk_fld3]) as sk_fld3, '+@CR+
                  '       Rtrim(M.[sk_fld4]) as sk_fld4, '+@CR+
                  '       Rtrim(M.[sk_fld5]) as sk_fld5, '+@CR+
                  '       Rtrim(M.[sk_fld6]) as sk_fld6, '+@CR+
                  '       Rtrim(M.[sk_fld7]) as sk_fld7, '+@CR+
                  '       Rtrim(M.[sk_fld8]) as sk_fld8, '+@CR+
                  '       Rtrim(M.[sk_fld9]) as sk_fld9, '+@CR+
                  '       Rtrim(M.[sk_fld10]) as sk_fld10, '+@CR+
                  '       Rtrim(M.[sk_fld11]) as sk_fld11, '+@CR+
                  '       Rtrim(M.[sk_fld12]) as sk_fld12, '+@CR+
                  '       M.[sk_ltime], M.[sk_ikind], M.[s_supp], M.[s_locate], M.[s_price1], '+@CR+
                  '       M.[s_price2], M.[s_price3], M.[s_price4], M.[s_price5], M.[s_price6], '+@CR+
                  '       M.[s_aprice], M.[s_m_ave], M.[s_updat1], M.[s_updat2], M.[s_lprice1], '+@CR+
                  '       M.[s_lprice2], M.[s_accno1], M.[s_accno2], M.[s_accno3], M.[s_accno4], '+@CR+
                  '       M.[s_accno5], M.[s_accno6], M.[st_jname], [st_jspes]=Rtrim(Convert(Varchar(255), M.[st_jspes])), M.[sk_nowqty], '+@CR+
                  '       M.[st_proven], M.[st_lenght], M.[st_width], M.[st_height], M.[st_lwhunit], '+@CR+
                  '       M.[st_unw], M.[st_ugw], M.[st_uappr], M.[st_uarea], M.[st_areaunt], '+@CR+ 
                  '       M.[st_freno], M.[st_fign2], M.[st_fign3], M.[st_fign4], M.[sk_poseq], '+@CR+
                  '       M.[sk_abcode], M.[s_apice2], M.[s_apice3], M.[s_apice4], M.[s_apice5], '+@CR+
                  '       M.[s_apice6], '+@CR+
                  '       M.[sk_wfno], '+@CR+
                  '       Rtrim(M.[sk_whno]) as sk_whno, M.[sk_mdpfg], M.[sk_lotfg], '+@CR+
                  -- [sk_no]
                  '       Chg_skno_accno = isnull(substring(M.[sk_no], 1, 1), ''00''), '+@CR+
                  '       Chg_skno_accno_Name = isnull(D1.Code_Name, ''未設定''), '+@CR+
                  '       Chg_skno_BKind = isnull(substring(M.[sk_no], 1, 2), ''00''), '+@CR+
                  '       Chg_skno_Bkind_Name = isnull(D2.Code_Name, ''未設定''),'+@CR+
                  -- Rickliu 2014/09/29 協理要求將 AD, AE, AE 納入 AA 百貨計算
                  '       Chg_skno_BKind2 = '+@CR+
                  '         Case '+@CR+
                  '           when substring(M.[sk_no], 1, 2) in (''AD'', ''AE'', ''AZ'') then ''AA'' '+@CR+
                  '           else isnull(substring(M.[sk_no], 1, 2), ''00'') '+@CR+
                  '         end, '+@CR+
                  '       Chg_skno_Bkind_Name2 = '+@CR+
                  '         Case '+@CR+
                  '           when substring(M.[sk_no], 1, 2) in (''AD'', ''AE'', ''AZ'') then '+@CR+
                  '                isnull((select code_name as code_name '+@CR+
                  '                          from sstock_code '+@CR+
                  '                         where code_level = 2 '+@CR+
                  '                           and code_no = ''AA'' '+@CR+
                  '                       ), ''未設定'') '+@CR+
                  '           else isnull(D2.Code_Name, ''未設定'') '+@CR+
                  '         end, '+@CR+
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
                  '       Chg_skno_BKind3 = '+@CR+
                  '         Case '+@CR+
                  '           when substring(M.[sk_no], 1, 2) in (''AA'', ''BA'', ''CA'') then ''AA'' '+@CR+
                  '           when substring(M.[sk_no], 1, 2) in (''AB'', ''BB'', ''CB'') then ''AB'' '+@CR+
                  '           when substring(M.[sk_no], 1, 2) in (''AC'', ''BC'', ''CC'') then ''AC'' '+@CR+
                  '           when substring(M.[sk_no], 1, 2) in (''AD'', ''BD'', ''CD'') then ''AD'' '+@CR+
                  '           when substring(M.[sk_no], 1, 2) in (''AE'', ''BE'', ''CE'') then ''AE'' '+@CR+
                  '           when substring(M.[sk_no], 1, 2) in (''AF'', ''BF'', ''CF'') then ''AF'' '+@CR+
                  '           when substring(M.[sk_no], 1, 1) in (''E'') then ''OT'' '+@CR+
                  '           else '''' '+@CR+
                  '         end, '+@CR+
                  '       Chg_skno_Bkind_Name3 = '+@CR+
                  '         Case '+@CR+
                  '           when substring(M.[sk_no], 1, 2) in (''AA'', ''BA'', ''CA'') then ''百貨'' '+@CR+
                  '           when substring(M.[sk_no], 1, 2) in (''AB'', ''BB'', ''CB'') then ''3C'' '+@CR+
                  '           when substring(M.[sk_no], 1, 2) in (''AC'', ''BC'', ''CC'') then ''化工'' '+@CR+
                  '           when substring(M.[sk_no], 1, 2) in (''AD'', ''BD'', ''CD'') then ''工具'' '+@CR+
                  '           when substring(M.[sk_no], 1, 2) in (''AE'', ''BE'', ''CE'') then ''電裝'' '+@CR+
                  '           when substring(M.[sk_no], 1, 2) in (''AF'', ''BF'', ''CF'') then ''客訂'' '+@CR+
                  '           when substring(M.[sk_no], 1, 1) in (''E'') then ''其他'' '+@CR+
                  '           else '''' '+@CR+
                  '         end, '+@CR+
                  '       Chg_skno_SKind = substring(M.[sk_no], 1, 4), '+@CR+
                  '       Chg_skno_SKind_Name = D3.Code_Name, '+@CR+
                  '       Chg_kind_Name = D4.kd_name, '+@CR+
                  '       Chg_ivkd_Name = D5.kd_name, '+@CR+
                  '       Chg_StartYear = substring(Replace(M.[sk_size], ''無'',''''), 1, 4), '+@CR+
                  '       Chg_StartMonth = substring(Replace(M.[sk_size], ''無'',''''), 5, 2), '+@CR+
                  '       Chg_EndYear = substring(Replace(M.[sk_color], ''無'',''''), 1, 4), '+@CR+
                  '       Chg_EndMonth = substring(Replace(M.[sk_color], ''無'',''), 5, 2), '+@CR+
                  '       Chg_supp_Name = D6.ct_name, '+@CR+
                  '       Chg_supp_SName = D6.ct_sname, '+@CR+
                  '       Chg_locate_MArea = substring(M.[s_locate], 1, 1), '+@CR+
                  '       Chg_locate_DArea = substring(M.[s_locate], 1, 2), '+@CR+
                  '       Chg_locate_row = substring(M.[s_locate], 4, 1), '+@CR+
                  '       Chg_locate_col = substring(M.[s_locate], 6, 1), '+@CR+
                  '       Chg_updat1_Year = DATEPART(yy, M.s_updat1), '+@CR+
                  '       Chg_updat1_Month = DATEPART(mm, M.s_updat1), '+@CR+
                  '       Chg_updat2_Year = DATEPART(yy, M.s_updat2), '+@CR+
                  '       Chg_updat2_Month = DATEPART(mm, M.s_updat2), '+@CR+
                  -- 2014/06/07 未銷商品註記
                  '       Chg_Stock_NonSales = '+@CR+
                  '         Case '+@CR+
                  '           when D7.sk_no is Not null Then ''Y'' '+@CR+
                  '           else ''N'' '+@CR+
                  '         End, '+@CR+
                  -- 2014/06/17 Rick 新增 AA 期初安全存量
                  -- Chg_WD_AA_first_sQty = Isnull(D8.wd_first_sqty, 0),
                  '       Chg_WD_AA_sQty = Isnull(D8.wd_sqty, 0), '+@CR+
                  -- 2014/06/17 Rick 新增 AA 期初現有庫存數
                  '       Chg_WD_AA_first_Qty = Isnull(D8.wd_first_qty, 0), '+@CR+
                  -- 2014/06/17 Rick 新增 AA 期初庫存差異數
                  '       Chg_WD_AA_first_diff_Qty = Isnull(D8.wd_first_qty_diff, 0), '+@CR+
                  -- 2014/06/07 Rick 新增 AA 期末現有庫存數
                  '       Chg_WD_AA_last_Qty = Isnull(D8.wd_last_qty, 0), '+@CR+
                  -- 2014/06/17 Rick 新增 AA 期末庫存差異數
                  '       Chg_WD_AA_last_diff_Qty = Isnull(D8.wd_last_qty_diff, 0), '+@CR+
                  -- 2014/06/17 Rick 新增 AB 期初安全存量
                  -- Chg_WD_AB_first_sQty = Isnull(D9.wd_first_sqty, 0),
                  '       Chg_WD_AB_sQty = Isnull(D9.wd_sqty, 0), '+@CR+
                  -- 2014/06/17 Rick 新增 AB 期初現有庫存數
                  '       Chg_WD_AB_first_Qty = Isnull(D9.wd_first_qty, 0), '+@CR+
                  -- 2014/06/17 Rick 新增 AB 期初庫存差異數
                  '       Chg_WD_AB_first_diff_Qty = Isnull(D9.wd_first_qty_diff, 0), '+@CR+
                  -- 2014/06/07 Rick 新增 AB 期末現有庫存數
                  '       Chg_WD_AB_last_Qty = Isnull(D9.wd_last_qty, 0), '+@CR+
                  -- 2014/06/17 Rick 新增 AB 期末庫存差異數
                  '       Chg_WD_AB_last_diff_Qty = Isnull(D9.wd_last_qty_diff, 0), '+@CR+
                  -- 2014/06/17 Rick 新增 AC 期初安全存量
                  -- Chg_WD_AB_first_sQty = Isnull(D10.wd_first_sqty, 0),
                  '       Chg_WD_AC_sQty = Isnull(D10.wd_sqty, 0), '+@CR+
                  -- 2014/06/17 Rick 新增 AC 期初現有庫存數
                  '       Chg_WD_AC_first_Qty = Isnull(D10.wd_first_qty, 0), '+@CR+
                  -- 2014/06/17 Rick 新增 AC 期初庫存差異數
                  '       Chg_WD_AC_first_diff_Qty = Isnull(D10.wd_first_qty_diff, 0), '+@CR+
                  -- 2014/06/07 Rick 新增 AC 期末現有庫存數
                  '       Chg_WD_AC_last_Qty = Isnull(D10.wd_last_qty, 0), '+@CR+
                  -- 2014/06/17 Rick 新增 AC 期末庫存差異數
                  '       Chg_WD_AC_last_diff_Qty = Isnull(D10.wd_last_qty_diff, 0), '+@CR+
                  -- 2014/06/17 Rick 新增 AG 期初安全存量
                  '       Chg_WD_AG_sQty = Isnull(D11.wd_sqty, 0), '+@CR+
                  -- Chg_WD_AG_first_sQty = Isnull(D10.wd_first_sqty, 0),
                  -- 2014/06/17 Rick 新增 AG 期初現有庫存數
                  '       Chg_WD_AG_first_Qty = Isnull(D11.wd_first_qty, 0), '+@CR+
                  -- 2014/06/17 Rick 新增 AG 期初庫存差異數
                  '       Chg_WD_AG_first_diff_Qty = Isnull(D11.wd_first_qty_diff, 0), '+@CR+
                  -- 2014/06/07 Rick 新增 AG 期末現有庫存數
                  '       Chg_WD_AG_last_Qty = Isnull(D11.wd_last_qty, 0), '+@CR+
                  -- 2014/06/17 Rick 新增 AG 期末庫存差異數
                  '       Chg_WD_AG_last_diff_Qty = Isnull(D11.wd_last_qty_diff, 0), '+@CR+
                  -- 2014/06/16 Rickliu 新增滯銷品
                  '       Chg_IS_Dead_Stock =  '+@CR+
                  '         Case '+@CR+
                  '           when D12.sk_no is not null then ''Y'' '+@CR+
                  '           else ''N'' '+@CR+
                  '         end, '+@CR+
                  '       Chg_Dead_First_Qty = Round(isnull(D12.First_Qty, 0), 0), '+@CR+
                  '       Chg_Dead_First_Amt = Round(isnull(D12.First_Amt, 0), 0), '+@CR+
                  '       Chg_Dead_Stock_YM = Convert(Varchar(7), Isnull(D12.Dead_YM, ''''), 111), '+@CR+
                  -- 2014/07/02 Rickliu 新增採購新品資訊
                  '       Chg_New_Arrival_Date = Convert(DateTime, D13.Arrival_Date), '+@CR+
                  '       Chg_New_Arrival_YM = Convert(Varchar(7), Convert(DateTime, D13.Arrival_Date), 111), '+@CR+
                  '       Chg_New_First_Qty = Round(isnull(D13.First_Qty, 0), 0), '+@CR+

                  -- 2015/05/15 Nanliao 新增產品屬性表資訊
                  '       Chg_color = D14.color, '+@CR+
                  '       Chg_size = D14.size, '+@CR+
                  '       Chg_package = D14.package, '+@CR+
                  '       Chg_barcode_name = D14.barcode_name, '+@CR+
                  '       Chg_price = D14.price, '+@CR+
                  '       Chg_pic_6 = D14.pic_6, '+@CR+
                  '       Chg_pic_4 = D14.pic_4, '+@CR+
                  '       Chg_pic_2 = D14.pic_2, '+@CR+
                  '       Chg_main_pic = D14.main_pic, '+@CR+
                  '       Chg_pic1 = D14.pic1, '+@CR+
                  '       Chg_pic2 = D14.pic2, '+@CR+
                  '       Chg_pic3 = D14.pic3, '+@CR+
                  '       Chg_avg_price = D14.avg_price, '+@CR+
                  '       Chg_product_property = D14.product_property, '+@CR+
                  '       Chg_gross_property = D14.gross_property, '+@CR+
                  '       sstock_update_datetime = getdate() '+@CR+
                  '       into Fact_sstock '+@CR+
                  '  from SYNC_'+@sDB+'.dbo.sstock M With(NoLock) '+@CR+
                  '       left join Ori_XLS#stock_code as D1 With(NoLock) '+@CR+
                  '              ON D1.Code_Level = ''1'' AND substring([sk_no], 1, 1) Collate Chinese_Taiwan_Stroke_CI_AS = D1.Code_No Collate Chinese_Taiwan_Stroke_CI_AS '+@CR+
                  '       left join Ori_XLS#stock_code as D2 With(NoLock) '+@CR+
                  '              ON D2.Code_Level = ''2'' AND substring([sk_no], 1, 2) Collate Chinese_Taiwan_Stroke_CI_AS = D2.Code_No Collate Chinese_Taiwan_Stroke_CI_AS '+@CR+
                  '       left join Ori_XLS#stock_code as D3 With(NoLock) '+@CR+
                  '              ON D3.Code_Level = ''3'' AND substring([sk_no], 1, 4) Collate Chinese_Taiwan_Stroke_CI_AS = D3.Code_No Collate Chinese_Taiwan_Stroke_CI_AS '+@CR+
                  '       left join SYNC_'+@sDB+'.dbo.skind as D4 With(NoLock) '+@CR+
                  '              ON D4.kd_class=''1'' AND M.sk_kind = D4.kd_no '+@CR+ 
                  '       left join SYNC_'+@sDB+'.dbo.skind as D5 With(NoLock) '+@CR+
                  '              ON D5.kd_class=''2'' AND M.sk_ivkd = D5.kd_no '+@CR+ 
                  '       left join SYNC_'+@sDB+'.pcust as D6 With(NoLock) '+@CR+
                  '              ON D6.ct_class=''2'' AND M.s_supp = D6.ct_no '+@CR+
                  '       left join Ori_xls#Stock_NonSales as D7 With(NoLock) '+@CR+
                  '              ON M.sk_no Collate Chinese_Taiwan_Stroke_CI_AS = D7.sk_no Collate Chinese_Taiwan_Stroke_CI_AS '+@CR+
                  '       left join SYNC_'+@sDB+'.dbo.V_Last_Swaredt as D8 With(NoLock) '+@CR+
                  '              On M.sk_no = D8.sk_no '+@CR+
                  '             And D8.Wd_no =''AA'' '+@CR+
                  '       left join SYNC_'+@sDB+'.dbo.V_Last_Swaredt as D9 With(NoLock) '+@CR+
                  '              On M.sk_no = D9.sk_no '+@CR+
                  '             And D9.Wd_no =''AB'' '+@CR+
                  '       left join SYNC_'+@sDB+'.dbo.V_Last_Swaredt as D10 With(NoLock) '+@CR+
                  '              On M.sk_no = D10.sk_no '+@CR+
                  '             And D9.Wd_no =''AC'' '+@CR+
                  '       left join SYNC_'+@sDB+'.dbo.V_Last_Swaredt as D11 With(NoLock) '+@CR+
                  '              On M.sk_no = D11.sk_no '+@CR+
                  '             And D11.Wd_no =''AG'' '+@CR+
                  '       left join Ori_xls#Dead_Stock_Lists as D12 With(NoLock) '+@CR+
                  '              on M.sk_no Collate Chinese_Taiwan_Stroke_CI_AS = D12.sk_no Collate Chinese_Taiwan_Stroke_CI_AS '+@CR+
                  '       left join Ori_xls#New_Stock_Lists as D13 With(NoLock) '+@CR+
                  '              on M.sk_no Collate Chinese_Taiwan_Stroke_CI_AS = D13.sk_no Collate Chinese_Taiwan_Stroke_CI_AS and D13.sk_no > '''' '+@CR+
                  -- 2015/05/15 Nanliao 新增產品屬性表資訊
                  '       left join Ori_xls#Stock_Property as D14 With(NoLock) '+@CR+
                  '              on M.sk_no Collate Chinese_Taiwan_Stroke_CI_AS = D14.sk_no Collate Chinese_Taiwan_Stroke_CI_AS and D14.kind = ''P'' '
    Exec @Result = SP_Exec_SQL @Proc, @Msg, @strSQL
    if @Result <> -1 Set @Result = 0
    
    
    set @Msg = '計算 ['+@Tb_Name+'] 資料表筆數...'
    set @strSQL = '--'+@Msg+@CR+'select count(1) as cnt from '+@Tb_Name
    print @strSQL
    insert into @RowCount Exec(@strSQL)
    Select @Cnt = Cnt from @RowCount
    
    if @Cnt = 0 Set @Result = -1
          
    Exec SP_Write_Log @Proc, @Msg, @strSQL, @Cnt
  end try
  begin catch
    set @Cnt = -1
    set @Msg = '(錯誤訊息:'+ERROR_MESSAGE()+')'

    Exec SP_Write_Log @Proc, @Msg, @Msg, @Cnt
  end catch
  Return(@Cnt)
end
GO
