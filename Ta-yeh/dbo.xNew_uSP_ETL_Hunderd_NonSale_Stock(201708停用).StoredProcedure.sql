USE [DW]
GO
/****** Object:  StoredProcedure [dbo].[xNew_uSP_ETL_Hunderd_NonSale_Stock(201708停用)]    Script Date: 08/18/2017 17:18:56 ******/
DROP PROCEDURE [dbo].[xNew_uSP_ETL_Hunderd_NonSale_Stock(201708停用)]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE Proc [dbo].[xNew_uSP_ETL_Hunderd_NonSale_Stock(201708停用)](@sDB Varchar(10) = 'TA13')
as
begin
  /*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
   Program Name: uSP_ETL_Hunderd_NonSale_Stock
   Create Date: 2014/06/05
   Creator: Rickliu
  *-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*/  

  Declare @Proc Varchar(50) = 'uSP_ETL_Hunderd_NonSale_Stock'
  Declare @Cnt Int =0
  Declare @RowCount table (cnt int)
  Declare @Msg Varchar(Max) =''
  Declare @strSQL Varchar(Max)=''
  Declare @CR Varchar(4) = char(13)+char(10)
  Declare @Tb_Name Varchar(20) = 'Hunderd_NonSale_NewStock'
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
       if @Result <> -1 set @Result = 0
    end

    set @Msg = '建立 ['+@Tb_Name+'] 資料表...'
    set @strSQL = ';With CTE_Qry1 as '+@CR+
                  '(select distinct ct_no8, ct_sname8, '+@CR+ -- 2017/02/08 Rickliu Add 增加八碼客戶編號
                  '        ct_fld3 as Chg_ct_fld3, Chg_Hunderd_Customer_Name, Chg_IS_Lan_Custom '+@CR+
                  '   from '+@sDB+'Fact_pcust '+@CR+
                  '  where CT_CLASS = ''1'' '+@CR+
                  '    and len(ct_no) >= 8 '+@CR+
                  '    and Chg_Hunderd_Customer = ''Y'''+@CR+
                  '), CTE_Qry2 '+@CR+
                  '(select ct_no8, sd_skno '+@CR+
                  '   from '+@sDB+'Fact_sslpdt '+@CR+ 
                  '  where (sd_slip_fg = ''2'' Or sd_slip_fg = ''6'') '+@CR+ -- 銷貨、託售皆算
                  '    and Chg_Hunderd_Customer = ''Y'' '+@CR+
                  -- 將百大與商品做 Cross Join
                  '), CTE_Qry3 '+@CR+
                  '(select CT_NO8, SK_NO, Chg_Stock_NonSales '+@CR+
                  '   from Fact_sstock '+@CR+
                  '        Cross Join CTE_Qry1 '+@CR+
                  '), CTE_Qry4 '+@CR+
                  -- 商品未銷列表
                  '(select sk_no as sd_skno, sk_name, '+@CR+
                  '        Chg_Wd_AA_first_Qty, Chg_Wd_AB_first_Qty, '+@CR+
                  '        Chg_Wd_AA_last_Qty, Chg_Wd_AB_last_Qty, '+@CR+
                  '        Chg_skno_accno, Chg_skno_accno_Name, '+@CR+
                  '        Chg_skno_BKind, Chg_skno_Bkind_Name, '+@CR+
                  '        Chg_skno_SKind, Chg_skno_SKind_Name, '+@CR+
                  '        Chg_kind_Name, Chg_Stock_NonSales '+@CR+
                  '   from Fact_sstock m '+@CR+
                  '  where Chg_Stock_NonSales = ''Y'' '+@CR+
                  ')'+@CR+
                  'select ''1'' as kind, ''已銷'' as kind_Name, ct_no8, ct_sname8, chg_ct_fld3, '+@CR+ -- 2017/02/08 Rickliu Add 增加八碼客戶編號
                  '       Chg_Hunderd_Customer_Name, Chg_IS_Lan_Custom, '+@CR+
                  '       sd_skno, sd_name, '+@CR+
                  '       Chg_Wd_AA_first_Qty, Chg_Wd_AB_first_Qty, '+@CR+
                  '       Chg_Wd_AA_last_Qty, Chg_Wd_AB_last_Qty, '+@CR+
                  '       m.Chg_skno_accno, m.Chg_skno_accno_Name, '+@CR+
                  '       m.Chg_skno_BKind, m.Chg_skno_Bkind_Name, '+@CR+
                  '       m.Chg_skno_SKind, m.Chg_skno_SKind_Name, '+@CR+
                  '       m.Chg_kind_Name, Chg_sp_date_YM, '+@CR+
                  '       Chg_sp_sales, Chg_sales_Name, '+@CR+
                  '       Chg_sale_Month_Master, '+@CR+
                  '       Chg_sale_Month_MasterName, '+@CR+
                  '       Isnull(SUM(sd_qty), 0) as qty, '+@CR+
                  '       isnull(sum(sd_stot), 0) as amt '+@CR+
                  '       update_datetime = getdate() '+@CR+
                  '  into Hunderd_NonSale_Stock '+@CR+
                  '  from Fact_sslpdt m '+@CR+
                  ' where 1=1 '+@CR+
                  '   and Chg_Hunderd_Customer = ''Y'' '+@CR+
                  '   and Chg_Stock_NonSales = ''Y'' '+@CR+
                  '   and (sd_slip_fg = ''2'' Or sd_slip_fg = ''6'') '+@CR+ -- 銷貨、託售皆算
                  '   and exists '+@CR+
                  '       (select * '+@CR+
                  '          from CTE_Qry3 d'+@CR+
                  '         where 1=1 '+@CR+
                  '           and d.Chg_Stock_NonSales =''Y'' '+@CR+
                  '           and m.ct_no8 = d.CT_NO8 '+@CR+
                  '           and m.sd_skno = d.sk_no '+@CR+
                  '         ) '+@CR+
                  ' group by Chg_sp_date_YM, ct_no8, ct_sname8, chg_ct_fld3, '+@CR+ -- 2017/02/08 Rickliu Add 增加八碼客戶編號
                  '       Chg_Hunderd_Customer_Name, Chg_IS_Lan_Custom, '+@CR+
                  '       sd_skno, sd_name, '+@CR+
                  '       Chg_Wd_AA_first_Qty, Chg_Wd_AB_first_Qty, '+@CR+
                  '       Chg_Wd_AA_last_Qty, Chg_Wd_AB_last_Qty, '+@CR+
                  '       m.Chg_skno_accno, m.Chg_skno_accno_Name, '+@CR+
                  '       m.Chg_skno_BKind, m.Chg_skno_Bkind_Name, '+@CR+
                  '       m.Chg_skno_SKind, m.Chg_skno_SKind_Name, '+@CR+
                  '       m.Chg_kind_Name, Chg_sp_sales, Chg_sales_Name, '+@CR+
                  '       Chg_sale_Month_Master, Chg_sale_Month_MasterName '+@CR+
                  '  union '+@CR+
                  ' select ''2'' as kind, ''未銷'' as kind_name, ct_no8, ct_sname8, Chg_ct_fld3, '+@CR+ -- 2017/02/08 Rickliu Add 增加八碼客戶編號
                  '        Chg_Hunderd_Customer_Name, Chg_IS_Lan_Custom, '+@CR+
                  '        m.sd_skno, m.sk_name, '+@CR+
                  '        Chg_Wd_AA_first_Qty, Chg_Wd_AB_first_Qty, '+@CR+
                  '        Chg_Wd_AA_last_Qty, Chg_Wd_AB_last_Qty, '+@CR+
                  '        Chg_skno_accno, Chg_skno_accno_Name, '+@CR+
                  '        Chg_skno_BKind, Chg_skno_Bkind_Name, '+@CR+
                  '        Chg_skno_SKind, Chg_skno_SKind_Name, '+@CR+
                  '        Chg_kind_Name, Chg_sp_date_YM, '+@CR+
                  '        ''NA'' as Chg_sp_sales, ''NA'' as Chg_sales_Name, '+@CR+
                  '        case when SK_MName is null then ''N'' '+@CR+
                  '          when SK_MName = ''NA'' then ''N'' '+@CR+
                  '          else ''Y'' '+@CR+
                  '        end as Chg_sale_Month_Master, '+@CR+
                  '        isnull(sk_Mname, ''NA'') as sk_Mname, '+@CR+
                  '        0 as qty, 0 as amt '+@CR+
                  '   from CTE_Qry4 m '+@CR+
                  '        Cross Join '+@CR+
                  -- 2017/02/08 Rickliu Add 增加八碼客戶編號 
                  '        CTE_Qry1 D '+@CR+
                  '        Cross Join '+@CR+ 
                  '        (select Chg_sp_date_YM '+@CR+
                  '           from (select distinct CONVERT(varchar(7), cal_date, 111) as Chg_sp_date_YM '+@CR+
                  '                   from Calendar '+@CR+
                  '                  where CONVERT(varchar(7), cal_date, 111) <= CONVERT(varchar(7), GETDATE(), 111) '+@CR+
                  '                    and CONVERT(varchar(7), cal_date, 111) >= ''2012/01'') M '+@CR+
                  '          where 1=1 '+@CR+
                  '        ) D1 '+@CR+
                  '        left join Ori_XLS#Master_Stock D2 '+@CR+
                  '          ON m.sd_skno collate Chinese_Taiwan_Stroke_CI_AS = D2.sk_no collate Chinese_Taiwan_Stroke_CI_AS '+@CR+
                  '         and Substring(D1.Chg_sp_date_YM, 1, 4) = D2.Sale_Year '+@CR+
                  '         and Substring(D1.Chg_sp_date_YM, 5, 2) = D2.Sale_month '+@CR+
                  '  where not exists '+@CR+
                  '        (select * '+@CR+
                  '           from Fact_sslpdt '+@CR+
                  '          where sd_slip_fg in (''2'', ''6'') '+@CR+ -- 銷貨、託售皆算
                  '            and Chg_sp_date_YM=D1.Chg_sp_date_YM '+@CR+
                  '            and ct_no8 = d.ct_no8 '+@CR+
                  '            and sd_skno = m.sd_skno '+@CR+
                  '            and Chg_Hunderd_Customer = ''Y'' '+@CR+
                  '            and Chg_Stock_NonSales = ''Y'' '+@CR+
                  '            and chg_ct_close = '''' '+@CR+
                  '         ) '+@CR+
                  ')m '
         

    select @Result = count(1) from Hunderd_NonSale_Stock

    Exec SP_Write_Log @Proc, @Msg, '', @Result
  end try
  begin catch
    set @Result = -1
    set @Msg = @Msg+'(錯誤訊息:'+ERROR_MESSAGE()+')'

    Exec SP_Write_Log @Proc, @Msg, '', @Result
  end catch
  Return(@Result)
end
GO
