USE [DW]
GO
/****** Object:  StoredProcedure [dbo].[xNew_uSP_ETL_Cust_NonSale_Stock(201708停用)]    Script Date: 08/18/2017 17:43:41 ******/
DROP PROCEDURE [dbo].[xNew_uSP_ETL_Cust_NonSale_Stock(201708停用)]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE Proc [dbo].[xNew_uSP_ETL_Cust_NonSale_Stock(201708停用)](@sDB Varchar(10) = 'TA13')
as
begin
  /*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
   Program Name: uSP_ETL_Cust_NonSale_Stock
   Create Date: 2014/06/05
   Creator: Rickliu
  *-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*/  

  Declare @Proc Varchar(50) = 'uSP_ETL_Cust_NonSale_Stock'
  Declare @Cnt Int =0
  Declare @RowCount table (cnt int)
  Declare @Msg Varchar(Max) =''
  Declare @strSQL Varchar(Max)=''
  Declare @CR Varchar(4) = char(13)+char(10)
  Declare @Tb_Name Varchar(20) = 'Cust_NonSale_Stock'
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
       Set @Msg = '刪除資料表 ['+@Tb_Name+']'+@CR+
                  'DROP TABLE [dbo].['+@Tb_Name+']'
       Exec @Result = SP_Exec_SQL @Proc, @Msg, @strSQL

       if @Result <> -1 Set @Result = 0
    end

    set @Msg = '建立 ['+@Tb_Name+'] 資料表...'
    set @strSQL = ';With CTE_Qry1 as '+@CR+
                  '(select ct_no8, ct_sname8, '+@CR+ -- 2017/02/08 Rickliu Add 增加八碼客戶編號
                  '         ct_sales, Chg_ct_sales_name, '+@CR+
                  '         ct_fld3, '+@CR+
                  '         Chg_hunderd_customer, Chg_Hunderd_Customer_Name, '+@CR+
                  '         Chg_is_lan_custom '+@CR+
                  '    from '+@sDB+'fact_pcust '+@CR+
                  '   where 1=1 '+@CR+
                  '     and ct_class = 1 '+@CR+
                  '     and substring(ct_no, 9, 1) =  ''1'' '+@CR+  --20150428 Modify by Nan 配合宜靜取得客編尾碼為1資料
                  '     and (substring(ct_no, 1, 2) <> (''IT'' Or substring(ct_no, 1, 2) <> ''IZ'') '+@CR+
                  '), CTE_Qry2 as '+@CR+
                  '(select sk_no, sk_name, sk_save '+@CR+
                  '        Chg_Wd_AA_first_Qty, Chg_Wd_AB_first_Qty, '+@CR+
                  '        Chg_Wd_AA_last_Qty, Chg_Wd_AB_last_Qty, '+@CR+
                  '        Chg_Wd_AA_first_Qty * sk_save as Chg_Wd_AA_first_amt, '+@CR+
                  '        Chg_Wd_AB_first_Qty * sk_save as Chg_Wd_AB_first_amt, '+@CR+ 
                  '        Chg_Wd_AA_last_Qty * sk_save as Chg_Wd_AA_last_amt, '+@CR+
                  '        Chg_Wd_AB_last_Qty * sk_save as Chg_Wd_AB_last_amt, '+@CR+
                  '         Chg_skno_accno, Chg_skno_accno_Name, Chg_skno_BKind, '+@CR+
                  '         Chg_skno_Bkind_Name, Chg_skno_SKind, Chg_skno_SKind_Name, '+@CR+
                  '         Chg_kind_Name, Chg_is_dead_stock, Chg_dead_stock_ym, Chg_new_arrival_ym '+@CR+
                  '    from '+@sDB+'fact_sstock '+@CR+
                  '   where 1=1 '+@CR+
                  '     and substring(sk_no, 1, 1) = ''A'' '+@CR+
                  '     and sk_name <> '''' '+@CR+
                  '     and Chg_stock_nonsales = ''Y'' '+@CR+
                  ' ), CTE_Qry3 as '+@CR+
                  '(select ct_no8, sd_skno, sum(sd_qty) as sd_qty, sum(sd_stot) as sd_stot, '+@CR+
                  '        case when sd_slip_fg = ''2'' then SUM(Isnull(sd_qty, 0)) else 0 end as sd2_qty, '+@CR+
                  '        case when sd_slip_fg = ''6'' then SUM(Isnull(sd_qty, 0)) else 0 end as sd6_qty, '+@CR+
                  '        case when sd_slip_fg = ''7'' then SUM(Isnull(sd_qty, 0)) else 0 end as sd7_qty, '+@CR+ --20150904 Add by Nan 增加託售及回貨計算
                  '        case when sd_slip_fg = ''2'' then SUM(Isnull(sd_stot, 0)) else 0 end as sd2_stot, '+@CR+
                  '        case when sd_slip_fg = ''6'' then SUM(Isnull(sd_stot, 0)) else 0 end as sd6_stot, '+@CR+
                  '        case when sd_slip_fg = ''7'' then SUM(Isnull(sd_stot, 0)) else 0 end as sd7_stot '+@CR+  --20150904 Add by Nan 增加託售及回貨計算
                  '   from '+@sDB+'fact_sslpdt '+@CR+
                  '  where 1=1 '+@CR+
                  '    and (sd_slip_fg  = ''2'' or sd_slip_fg  = ''6'' or sd_slip_fg  = ''7'') '+@CR+
                  '    and substring(sd_ctno, 9, 1) =  ''1'' '+@CR+  --20150428 Modify by Nan 配合宜靜取得客編尾碼為1資料
                  '    and (sd_ctno not like ''IT%'' or sd_ctno not like ''IZ%'') '+@CR+
                  '    and Chg_stock_nonsales = ''Y'' '+@CR+
                  -- 近兩年資料
                  '    and sd_date >= '+@CR+
                  '        case '+@CR+
                  '          when Convert(datetime, ''2012/12/01'') >= dateadd(yy, -2, getdate()) '+@CR+
                  '          then Convert(datetime, ''2012/12/01'') '+@CR+
                  '          else dateadd(yy, -2, getdate()) '+@CR+
                  '        end '+@CR+
                  '   group by ct_no8, sd_skno,sd_slip_fg '+@CR+--20150904 Modify by Nan 增加託售及回貨計算
                  ') '+@CR+
                  'select m.*, d.*, '+@CR+
                  '       isnull(d1.sd_qty, 0) as sd_qty, '+@CR+
                  '       isnull(d1.sd_stot, 0) as sd_stot, '+@CR+
                  '       isnull(d1.sd2_qty, 0) as sd2_qty, '+@CR+   --20150904 Add by Nan 增加託售及回貨計算
                  '       isnull(d1.sd2_stot, 0) as sd2_stot, '+@CR+ --20150904 Add by Nan 增加託售及回貨計算
                  '       isnull(d1.sd6_qty, 0) as sd6_qty, '+@CR+   --20150904 Add by Nan 增加託售及回貨計算
                  '       isnull(d1.sd6_stot, 0) as sd6_stot, '+@CR+ --20150904 Add by Nan 增加託售及回貨計算
                  '       isnull(d1.sd7_qty, 0) as sd7_qty, '+@CR+   --20150904 Add by Nan 增加託售及回貨計算
                  '       isnull(d1.sd7_stot, 0) as sd7_stot, '+@CR+ --20150904 Add by Nan 增加託售及回貨計算
                  '       case when isnull(sd_qty, 0) = 0 then ''2'' else ''1'' end sale_flag '+@CR+
                  '       0 as dcount_ctno, '+@CR+ 
                  '       0 as dcount_ctno8, '+@CR+ -- 2017/02/08 Rickliu Add 增加八碼客戶編號
                  '       0 as dcount_skno, '+@CR+
                  '       update_datetime = getdate() '+@CR+
                  '  into '+@Tb_Name+@CR+
                  '  from CTE_Qry1 m '+@CR+
                  '       Cross join CTE_Qry2 d '+@CR+
                  '       left join CTE_Qry3 d1 '+@CR+
                  '         on m.ct_no8 = d1.ct_no8 and d.sk_no = d1.sd_skno '
    Exec @Result = SP_Exec_SQL @Proc, @Msg, @strSQL
    if @Result <> -1 Set @Result = 0
    
    -- 2015/01/26 Rickliu 增加業務的客戶數量及商品數量
    set @Msg = '更新 ['+@Tb_Name+'] 資料表 ...'
    set @strSQL = ';With CTE_Qry as '+@CR+
                  '(select ct_sales, '+@CR+
                  '        Count(Distinct ct_no8) as dcount_ctno, '+@CR+ -- 2017/02/08 Rickliu Add 增加八碼客戶編號
                  '        Count(Distinct sk_no) as dcount_skno '+@CR+
                  '   from '+@Tb_Name+@CR+
                  '  group by ct_sales '+@CR+
                  ') '+@CR+
                  'update '+@Tb_Name+' '+@CR+
                  '   set dcount_ctno = d.dcount_ctno, '+@CR+
                  '       dcount_skno = d.dcount_skno '+@CR+
                  '  from '+@Tb_Name+' m '+@CR+
                  '       left join CTE_Qry d '+@CR+
                  '         on m.ct_sales = d.ct_sales '
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
