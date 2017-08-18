USE [DW]
GO
/****** Object:  StoredProcedure [dbo].[xNew_uSP_ETL_Cust_WeekSale_Stock(201708停用)]    Script Date: 08/18/2017 17:18:56 ******/
DROP PROCEDURE [dbo].[xNew_uSP_ETL_Cust_WeekSale_Stock(201708停用)]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE Proc [dbo].[xNew_uSP_ETL_Cust_WeekSale_Stock(201708停用)](@sDB Varchar(10) = 'TA13')
as
begin
  /*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
   Program Name: uSP_ETL_Cust_WeekSale_Stock
   Create Date: 2015/08/21
   Creator: Nanliao
  *-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*/  

  Declare @Proc Varchar(50) = 'uSP_ETL_Cust_WeekSale_Stock'
  Declare @Cnt Int =0
  Declare @RowCount table (cnt int)
  Declare @Msg Varchar(Max) =''
  Declare @strSQL Varchar(Max)=''
  Declare @CR Varchar(4) = char(13)+char(10)
  Declare @Tb_Name Varchar(20) = 'Cust_WeekSale_Stock'
  Declare @Result int = 0
  
  begin try
    IF Not EXISTS (SELECT * FROM sys.databases WHERE name = 'SYNC_'+@Tb_Name)
    begin
       set @Msg = '所選的資料庫不存在 ['+@sDB+']....'
       Raiserror(@Msg, 16, 1)
    end

    set @sDB = @sDB+'#'
    set @Tb_Name = @sDB+'#'+@Tb_Name

    IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].['+@Tb_Name+']') AND type in (N'U'))
    begin
       Set @Msg = '刪除資料表 ['+@Tb_Name+']'
       set @strSQL= 'DROP TABLE [dbo].['+@Tb_Name+']'
       Exec @Result = SP_Exec_SQL @Proc, @Msg, @strSQL

       if @Result <> -1 Set @Result = 0
    end

    set @Msg = '建立 ['+@Tb_Name+'] 資料表...'
    set @strSQL = ';With CTE_Qry1 as '+@CR+
                  '(select * '+@CR+
                  '   from '+@sDB+'Fact_sslpdt '+@CR+
                  '  where 1=1 '+@CR+
                  '    and (sd_slip_fg  = ''2'' or sd_slip_fg  = ''6'' or sd_slip_fg  = ''7'') '+@CR+
                  '), CTE_Qry2 '+@CR+
                  '(select ''1'' as kind, ''已銷'' as kind_Name, ct_no8, ct_sname8, chg_ct_fld3, '+@CR+ -- 2017/02/08 Rickliu Add 增加八碼客戶編號 
                  '        Chg_Hunderd_Customer_Name, Chg_IS_Lan_Custom, '+@CR+
                  '        sd_skno, sd_name, '+@CR+
                  '        Chg_Wd_AA_first_Qty, Chg_Wd_AB_first_Qty, '+@CR+
                  '        Chg_Wd_AA_last_Qty, Chg_Wd_AB_last_Qty, '+@CR+
                  '        Chg_skno_accno, Chg_skno_accno_Name, '+@CR+
                  '        Chg_skno_BKind, Chg_skno_Bkind_Name, '+@CR+
                  '        Chg_skno_SKind, Chg_skno_SKind_Name, '+@CR+
                  '        Chg_kind_Name, Chg_sp_date_YM, '+@CR+
                  '        cal_week_begin, cal_week_end, cal_week_Range, '+@CR+
                  '        cal_yearweek, cal_quarter, '+@CR+
                  '        Chg_sp_sales, Chg_sales_Name, '+@CR+
                  '        Chg_sale_Month_Master, '+@CR+
                  '        Chg_sale_Month_MasterName, '+@CR+
                  '        Chg_Hunderd_Customer, '+@CR+
                  '        Chg_Stock_NonSales, '+@CR+
                  '        s.Chg_is_dead_stock, '+@CR+
                  '        s.chg_dead_stock_ym, '+@CR+
                  '        chg_new_arrival_ym, '+@CR+
                  '        case when sd_slip_fg = ''2'' then SUM(Isnull(sd_qty, 0)) else 0 end as sd2_qty, '+@CR+
                  '        case when sd_slip_fg = ''6'' then SUM(Isnull(sd_qty, 0)) else 0 end as sd6_qty, '+@CR+
                  '        case when sd_slip_fg = ''7'' then SUM(Isnull(sd_qty, 0)) else 0 end as sd7_qty, '+@CR+
                  '        case when sd_slip_fg = ''2'' then SUM(Isnull(sd_stot, 0)) else 0 end as sd2_stot, '+@CR+
                  '        case when sd_slip_fg = ''6'' then SUM(Isnull(sd_stot, 0)) else 0 end as sd6_stot, '+@CR+
                  '        case when sd_slip_fg = ''7'' then SUM(Isnull(sd_stot, 0)) else 0 end as sd7_stot '+@CR+
                  '   from CTE_Qry1 m '+@CR+
                  '        inner join Calendar d '+@CR+
                  '           on cal_week =''1'' '+@CR+
                  '          and Convert(varchar(7), cal_week_begin, 111) >= ''2012/12'' '+@CR+
                  '          and Convert(varchar(7), cal_week_end, 111) <= Convert(varchar(7), getdate(), 111) '+@CR+
                  '          and sd_date BETWEEN d.cal_week_begin AND cal_week_end '+@CR+
                  '         left join '+@sDB+'Fact_sstock as s '+@CR+
                  '           on sd_skno=s.sk_no '+@CR+
                  '  where 1=1 '+@CR+
                  '    and len(rtrim(sd_ctno)) >= 8 '+@CR+
                  '  group by sd_slip_fg, Chg_sp_date_YM, ct_no8, ct_sname8, chg_ct_fld3, '+@CR+ -- 2017/02/08 Rickliu Add 增加八碼客戶編號
                  '        Chg_Hunderd_Customer_Name, Chg_IS_Lan_Custom, '+@CR+
                  '        sd_skno, sd_name, '+@CR+
                  '        Chg_Wd_AA_first_Qty, Chg_Wd_AB_first_Qty, '+@CR+
                  '        Chg_Wd_AA_last_Qty, Chg_Wd_AB_last_Qty, '+@CR+
                  '        Chg_skno_accno, Chg_skno_accno_Name, '+@CR+
                  '        Chg_skno_BKind, Chg_skno_Bkind_Name, '+@CR+
                  '        Chg_skno_SKind, Chg_skno_SKind_Name, '+@CR+
                  '        Chg_kind_Name, Chg_sp_sales, Chg_sales_Name, '+@CR+
                  '        Chg_sale_Month_Master, Chg_sale_Month_MasterName, '+@CR+
                  '        Chg_Hunderd_Customer, '+@CR+
                  '        Chg_Stock_NonSales, '+@CR+
                  '        s.Chg_is_dead_stock, '+@CR+ 
                  '        s.chg_dead_stock_ym, '+@CR+
                  '        chg_new_arrival_ym, '+@CR+
                  '        cal_week_begin, cal_week_end, cal_week_Range, '+@CR+
                  '        cal_yearweek, cal_quarter '+@CR+
                  ') '+@CR+
                  'select kind, kind_Name, ct_no8, ct_sname8, chg_ct_fld3, '+@CR+ -- 2017/02/08 Rickliu Add 增加八碼客戶編號
                  '       Chg_Hunderd_Customer_Name, Chg_IS_Lan_Custom, '+@CR+
                  '       sd_skno, sd_name, '+@CR+
                  '       Chg_Wd_AA_first_Qty, Chg_Wd_AB_first_Qty, '+@CR+
                  '       Chg_Wd_AA_last_Qty, Chg_Wd_AB_last_Qty, '+@CR+
                  '       Chg_skno_accno, Chg_skno_accno_Name, '+@CR+
                  '       Chg_skno_BKind, Chg_skno_Bkind_Name, '+@CR+
                  '       Chg_skno_SKind, Chg_skno_SKind_Name, '+@CR+
                  '       Chg_kind_Name, Chg_sp_date_YM, '+@CR+
                  '       cal_week_begin, cal_week_end, cal_week_Range, '+@CR+
                  '       cal_yearweek, cal_quarter, '+@CR+
                  '       Chg_sp_sales, Chg_sales_Name, '+@CR+
                  '       Chg_sale_Month_Master, '+@CR+
                  '       Chg_sale_Month_MasterName, '+@CR+
                  '       Chg_Hunderd_Customer, '+@CR+
                  '       Chg_Stock_NonSales, '+@CR+
                  '       Chg_is_dead_stock, '+@CR+
                  '       chg_dead_stock_ym, '+@CR+
                  '       chg_new_arrival_ym, '+@CR+
                  '       SUM(Isnull(sd2_qty, 0)) as sd2_qty, '+@CR+
                  '       SUM(Isnull(sd6_qty, 0)) as sd6_qty, '+@CR+
                  '       SUM(Isnull(sd7_qty, 0)) as sd7_qty, '+@CR+
                  '       SUM(Isnull(sd2_stot, 0)) as sd2_stot, '+@CR+
                  '       SUM(Isnull(sd6_stot, 0)) as sd6_stot, '+@CR+
                  '       SUM(Isnull(sd7_stot, 0)) as sd7_stot, '+@CR+
                  '       0 as all_count_ctno, 0 as all_count_skno, '+@CR+
                  '       0 as master_count_ctno, 0 as master_count_skno, '+@CR+
                  '       0 as nonsale_count_ctno, 0 as nonsale_count_skno, '+@CR+
                  '       0 as newstock_count_ctno, 0 as newstock_count_skno, '+@CR+
                  '       0 as deadstock_count_ctno, 0 as deadstock_count_skno, '+@CR+
                  '       update_datetime = getdate() '+@CR+
                  '       into '+@Tb_Name+@CR+
                  '  from CTE_Qry2 a '+@CR+
                  ' group by kind,kind_Name, ct_no8, ct_sname8, chg_ct_fld3, '+@CR+ -- 2017/02/08 Rickliu Add 增加八碼客戶編號
                  '       Chg_Hunderd_Customer_Name, Chg_IS_Lan_Custom, '+@CR+
                  '       sd_skno, sd_name, '+@CR+
                  '       Chg_Wd_AA_first_Qty, Chg_Wd_AB_first_Qty, '+@CR+
                  '       Chg_Wd_AA_last_Qty, Chg_Wd_AB_last_Qty, '+@CR+
                  '       Chg_skno_accno, Chg_skno_accno_Name, '+@CR+
                  '       Chg_skno_BKind, Chg_skno_Bkind_Name, '+@CR+ 
                  '       Chg_skno_SKind, Chg_skno_SKind_Name, '+@CR+
                  '       Chg_kind_Name, Chg_sp_date_YM, '+@CR+
                  '       cal_week_begin, cal_week_end, cal_week_Range, '+@CR+
                  '       cal_yearweek, cal_quarter, '+@CR+
                  '       Chg_sp_sales, Chg_sales_Name, '+@CR+
                  '       Chg_Hunderd_Customer, '+@CR+
                  '       Chg_Stock_NonSales, '+@CR+
                  '       Chg_is_dead_stock, '+@CR+
                  '       chg_dead_stock_ym, '+@CR+ 
                  '       chg_new_arrival_ym, '+@CR+
                  '       Chg_sale_Month_Master, '+@CR+
                  '       Chg_sale_Month_MasterName '
    Exec @Result = SP_Exec_SQL @Proc, @Msg, @strSQL
    if @Result <> -1 Set @Result = 0

    -- 2015/01/26 Rickliu 增加業務的客戶數量及商品數量
    set @Msg = '更新 ['+@Tb_Name+'] 資料表 .1/5...'
    set @strSQL = ';With CTE_Qry as '+@CR+
                  '(select Chg_sp_sales, cal_week_Range, '+@CR+
                  '        Count(Distinct ct_no8) as all_count_ctno, '+@CR+
                  '        Count(Distinct sd_skno) as all_count_skno '+@CR+
                  '   from '+@Tb_Name+@CR+
                  ' group by Chg_sp_sales, cal_week_Range '+@CR+
                  ') '+@CR+
                  'update '+@Tb_Name+@CR+
                  '   set all_count_ctno = isnull(d.all_count_ctno,0), '+@CR+
                  '       all_count_skno = isnull(d.all_count_skno,0) '+@CR+
                  '  from '+@Tb_Name+' m '+@CR+
                  '       left join CTE_Qry d '+@CR+
                  '         on m.Chg_sp_sales = d.Chg_sp_sales '+@CR+
                  '        and m.cal_week_Range = d.cal_week_Range '

    Exec @Result = SP_Exec_SQL @Proc, @Msg, @strSQL
    if @Result <> -1 Set @Result = 0


    set @Msg = '更新 ['+@Tb_Name+'] 資料表 2/5...'
    set @strSQL = ';With CTE_Qry as '+@CR+
                  '(select Chg_sp_sales, cal_week_Range, '+@CR+
                  '        Count(Distinct ct_no8) as master_count_ctno, '+@CR+
                  '        Count(Distinct sd_skno) as master_count_skno '+@CR+
                  '   from '+@Tb_Name+@CR+
                  '  where Chg_sale_Month_Master = ''Y'' '+@CR+
                  '  group by Chg_sp_sales, cal_week_Range '+@CR+
                  ') '+@CR+
                  'update '+@Tb_Name+@CR+
                  '   set master_count_ctno = isnull(d.master_count_ctno,0), '+@CR+
                  '       master_count_skno = isnull(d.master_count_skno,0) '+@CR+
                  '  from '+@Tb_Name+' m '+@CR+
                  '       left join CTE_Qry d '+@CR+
                  '         on m.Chg_sp_sales = d.Chg_sp_sales '+@CR+
                  '        and m.cal_week_Range = d.cal_week_Range '
    Exec @Result = SP_Exec_SQL @Proc, @Msg, @strSQL
    if @Result <> -1 Set @Result = 0

    set @Msg = '更新 ['+@Tb_Name+'] 資料表 3/5...'
    set @strSQL = ';With CTE_Qry as '+@CR+
                  '(select Chg_sp_sales, cal_week_Range, '+@CR+
                  '        Count(Distinct ct_no8) as nonsale_count_ctno, '+@CR+
                  '        Count(Distinct sd_skno) as nonsale_count_skno '+@CR+
                  '   from '+@Tb_Name+@CR+
                  '  where Chg_Stock_NonSales = ''Y'' '+@CR+
                  '  group by Chg_sp_sales, cal_week_Range '+@CR+
                  ') '+@CR+
                  'update '+@Tb_Name+@CR+
                  '   set nonsale_count_ctno = isnull(d.nonsale_count_ctno,0), '+@CR+
                  '       nonsale_count_skno = isnull(d.nonsale_count_skno,0) '+@CR+
                  '  from '+@Tb_Name+' m '+@CR+
                  '       left join CTE_Qry d '+@CR+
                  '         on m.Chg_sp_sales = d.Chg_sp_sales '+@CR+
                  '        and m.cal_week_Range = d.cal_week_Range '
    Exec @Result = SP_Exec_SQL @Proc, @Msg, @strSQL
    if @Result <> -1 Set @Result = 0

    set @Msg = '更新 ['+@Tb_Name+'] 資料表 4/5...'
    set @strSQL = ';With CTE_Qry as '+@CR+
                  '(select Chg_sp_sales, cal_week_Range, '+@CR+
                  '        Count(Distinct ct_no8) as newstock_count_ctno, '+@CR+
                  '        Count(Distinct sd_skno) as newstock_count_skno '+@CR+
                  '   from '+@Tb_Name+@CR+
                  '  where chg_new_arrival_ym > '''' '+@CR+
                  '  group by Chg_sp_sales, cal_week_Range '+@CR+
                  ') '+@CR+
                  'update '+@Tb_Name+@CR+
                  '   set newstock_count_ctno = isnull(d.newstock_count_ctno,0), '+@CR+
                  '       newstock_count_skno = isnull(d.newstock_count_skno,0) '+@CR+
                  '  from '+@Tb_Name+' m '+@CR+
                  '       left join CTE_Qry d '+@CR+
                  '         on m.Chg_sp_sales = d.Chg_sp_sales '+@CR+
                  '        and m.cal_week_Range = d.cal_week_Range '
    Exec @Result = SP_Exec_SQL @Proc, @Msg, @strSQL
    if @Result <> -1 Set @Result = 0

    set @Msg = '更新 ['+@Tb_Name+'] 資料表 5/5...'
    set @strSQL = ';With CTE_Qry as '+@CR+
                  '(select Chg_sp_sales, cal_week_Range, '+@CR+
                  '        Count(Distinct ct_no8) as deadstock_count_ctno, '+@CR+
                  '        Count(Distinct sd_skno) as deadstock_count_skno '+@CR+
                  '   from '+@Tb_Name+@CR+
                  '  where Chg_is_dead_stock = ''Y'' '+@CR+
                  '  group by Chg_sp_sales,cal_week_Range '+@CR+
                  ') '+@CR+
                  'update '+@Tb_Name+@CR+
                  '   set deadstock_count_skno = isnull(d.deadstock_count_ctno,0), '+@CR+
                  '       deadstock_count_skno = isnull(d.deadstock_count_skno,0) '+@CR+
                  '  from '+@Tb_Name+' m '+@CR+
                  '       left join CTE_Qry d '+@CR+
                  '         on m.Chg_sp_sales = d.Chg_sp_sales '+@CR+
                  '        and m.cal_week_Range = d.cal_week_Range '
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
    set @Msg = @Msg+'(錯誤訊息:'+ERROR_MESSAGE()+')'
    set @Result = -1

    Exec SP_Write_Log @Proc, @Msg, @strSQL, @Result
  end catch

  Return(@Result)

end
GO
