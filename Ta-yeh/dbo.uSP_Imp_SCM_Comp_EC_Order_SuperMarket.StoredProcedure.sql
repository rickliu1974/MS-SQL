USE [DW]
GO
/****** Object:  StoredProcedure [dbo].[uSP_Imp_SCM_Comp_EC_Order_SuperMarket]    Script Date: 07/24/2017 14:43:59 ******/
DROP PROCEDURE [dbo].[uSP_Imp_SCM_Comp_EC_Order_SuperMarket]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE Proc [dbo].[uSP_Imp_SCM_Comp_EC_Order_SuperMarket]
as
begin
  /***********************************************************************************************************
     Tw_Mobile 後台拉單對帳單並進行凌越比對
  ***********************************************************************************************************/
  Declare @Proc Varchar(50) = 'uSP_Imp_SCM_Comp_EC_Order_SuperMarket'
  Declare @Cnt Int =0
  Declare @RowCnt Int =0
  Declare @Msg Varchar(4000) =''
  Declare @Errcode int = -1
  
  Declare @Kind Varchar(1), @KindName Varchar(10), @sd_class Varchar(1), @sd_slip_fg Varchar(1), @OD_Name Varchar(20), @OD_CName Varchar(20), @oper Varchar(1), @F1_Tag Varchar(100)
  Declare @xls_head Varchar(100), @TB_head Varchar(200), @TB_head_tmp Varchar(200), @TB_xls_Name Varchar(200), @TB_OD_Name Varchar(200), @TB_tmp_name Varchar(200)

  Declare @strSQL Varchar(Max)
  Declare @CR Varchar(4) = char(13)+char(10)
  Declare @RowCount Table (cnt int)
  Declare @RowData Table (aData Varchar(255))
  
  Declare @new_sdno varchar(10)
  Declare @sd_date varchar(10)
  Declare @sp_date varchar(10)
  Declare @Print_Date varchar(10)
  Declare @Last_Date Varchar(2)
  Declare @sp_cnt int
  Declare @xls_year varchar(4)
  Declare @xls_month varchar(2)
  Declare @Collate Varchar(50)
  Declare @ct_no Varchar(20)
  declare @b_pdate varchar(10)
  declare @e_pdate varchar(10)
       
  --產生貨單日期，2014/6/23 小嫻表示若 XLS 的往來期間為 4/1~4/30 則是 3月份的帳，而要對凌越 3/21~4/20 對帳日期的銷退貨單
  set @Last_Date = '20'
  set @sp_date = Substring(Convert(Varchar(10), getdate(), 111), 1, 8)+@Last_Date
  set @sd_date = @sp_date

  Declare @CompanyName Varchar(100), @CompanyLikeName Varchar(100), @Rm Varchar(100), @sp_maker Varchar(20), @str Varchar(200)
  Declare @Pos Int
  
  Set @CompanyName = '超級商城' --> 請勿亂變動
  Set @CompanyLikeName = '%'+@CompanyName+'%'
  Set @Rm = '系統匯入'+@CompanyName
  Set @sp_maker = 'Admin'
  Set @Collate = 'Collate Chinese_Taiwan_Stroke_CI_AS'
  set @ct_no ='I90140011'

  Set @xls_head = 'Ori_Xls#'
  Set @TB_head = 'Comp_EC_Order_SuperMarket'
  Set @TB_head_tmp = @TB_head+'_tmp'
  Set @TB_xls_Name = @xls_head+@TB_Head -- Ex: Ori_Xls#Comp_EC_Order_SuperMarket -- Excel 原檔資料

  -- Check 匯入檔案是否存在
  IF Not EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].['+@TB_xls_Name+']') AND type in (N'U'))
  begin
print '1'
     --Send Message
     set @Cnt = @Errcode
     set @strSQL = ''
     Set @Msg = '外部 Excel 匯入資料表 ['+@TB_xls_Name+']不存在，終止進行轉檔作業。'
     Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt
     -- 2013/11/28 增加失敗回傳值
     close Cur_Tw_Mobile_DataKind
     deallocate Cur_Tw_Mobile_DataKind
     Return(@Errcode)
  end
    
  IF Exists(SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].['+@TB_head_tmp+']') AND type in (N'U'))
  begin
print '2'
    Set @Msg = '清除 '+@CompanyName+' 對帳單暫存資料表 ['+@TB_head_tmp+']。'

    Set @strSQL = 'DROP TABLE [dbo].['+@TB_head_tmp+']'
    Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
  end
  
  IF Exists(SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].['+@TB_head+']') AND type in (N'U'))
  begin
print '3'
    Set @Msg = '清除 '+@CompanyName+' 對帳單介面資料表 ['+@TB_head+']。'

    Set @strSQL = 'DROP TABLE [dbo].['+@TB_head+']'
    Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
  end

print '4'
  Set @Msg = '取出出退貨資料。'
  Set @strSQL = 'Declare @YM Varchar(7) '+@CR+
                'Declare @B_Pdate Varchar(10) '+@CR+
                'Declare @E_Pdate Varchar(10) '+@CR+
                ' '+@CR+
                'select top 1 @YM = Convert(Varchar(7), F10, 111) '+@CR+
                '  from '+@TB_xls_name+@CR+
                ' where F1 like ''%付款區間%'' '+@CR+
                ' order by F10 '+@CR+
                ' '+@CR+
                'Exec uSP_Get_Cust_Pdate_Range '''+@ct_no+''', @YM, @b_pdate output, @e_pdate output '+@CR+
                ' '+@CR+
                ' '+@CR+
                'select case '+@CR+
                '         when F9 = '''' then ''2'' '+@CR+
                '         else ''3'' '+@CR+
                '       end slip_fg, '+@CR+ -- 出退貨註記
                '       Rtrim(isnull(F4, '''')) as Order_No, '+@CR+
                '       Convert(Int, isnull(replace(F6, '','', ''''), 0))+Convert(Int, isnull(replace(F7, '','', ''''), 0)) as xls_tot, '+@CR+
                '       xlsfileName, rowid, '+@CR+
                '       @YM as XLS_YM, @B_pdate as b_sp_pdate, @E_pdate as e_sp_pdate, Getdate() as Create_datetime'+@CR+
                '       into '+@TB_head_tmp+@CR+
                '  from '+@TB_xls_name+@CR+
                ' where 1 = 1 '+@CR+
                '   and F4 <> '''' '+@CR+
                '   and Isnumeric(F6) = 1 '+@CR+
                '   and Substring(Rtrim(isnull(F4, '''')), 1, 2) =''YM'' '+@CR+
                ' order by xlsFileName,  rowid '
                
  Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
                
print '5'
  Set @Msg = '將退貨金額變負值。'
  Set @strSQL = 'update '+@TB_head_tmp+' '+@CR+
                '   set xls_tot = xls_tot * -1'+@CR+
                ' where slip_fg =''3'' '
  Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
  
  
  --在此進行比對商品編號，會拿輔助編號以及商品基本資料檔進行比對作業
print '6'
  Set @Msg = '進行 XLS 比對 凌越 商品編號並寫入 ['+@TB_OD_Name+']資料表。'
  set @strSQL = 'select ''XLS'' as master, '+@CR+
                '       XLS_YM,'+@CR+
                '       Date_Range, '+@CR+
                '       rtrim(slip_fg) as slip_fg, '+@CR+
                '       case when rtrim(slip_fg) = ''2'' then ''出'' else ''退'' end as xls_slip_fg_name, '+@CR+
                '       Convert(VarChar(20), rtrim(order_no)) as order_no, '+@CR+
                '       Convert(Varchar(255), isnull(Chg_sp_slip_name, '''')) as sd_spec,'+@CR+
                '       Round(isnull(xls_tot, 0), 0) as xls_tot, Round(isnull(chg_sd_stot, 0), 0) as chg_sd_stot, '+@CR+
                '       0 as order_total, '+@CR+
                '       isnull(cnt, 0) as order_cnt, '+@CR+
                '       case '+@CR+
                '         when xls_tot = chg_sd_stot then ''N'' '+@CR+
                '         else ''Y'' '+@CR+
                '       end as Equal '+@CR+
                '       into '+@TB_head+@CR+
                '  from (select Order_no, Convert(Varchar(7), E_sp_pdate, 111) as XLS_YM, B_sp_pdate, E_sp_pdate,'+@CR+
                '               slip_fg, sum(xls_tot) as xls_tot, count(1) as cnt  '+@CR+
                '          from '+@TB_head_tmp+' '+@CR+
                '         group by slip_fg, Order_no, Convert(Varchar(7), E_sp_pdate, 111), B_sp_pdate, E_sp_pdate '+@CR+
                '       ) m '+@CR+
                '       left join '+@CR+
                '       (select sd_slip_fg, chg_sd_slip_fg, '+@CR+
                '               sd_spec, Chg_sp_slip_name, '+@CR+
                '               Round(sum(chg_sd_stot * 1.05), 0) as chg_sd_stot '+@CR+
                '          from Fact_sslpdt m '+@CR+
                '         where sd_class = ''1'' '+@CR+
                '           and sd_ctno ='''+@ct_no+''' '+@CR+

                '           and sp_pdate >= '+@CR+
                '               (select top 1 B_sp_pdate from '+@TB_head_tmp+') '+@CR+
                '           and sp_pdate <= '+@CR+
                '               (select top 1 E_sp_pdate from '+@TB_head_tmp+') '+@CR+

                '         group by sd_slip_fg, chg_sd_slip_fg, sd_spec, Chg_sp_slip_name '+@CR+
                '       ) d '+@CR+
                '         on Order_no '+@Collate+' = Rtrim(sd_spec) '+@Collate+' '+@CR+
                '        and slip_fg '+@Collate+' = sd_slip_fg '+@Collate+''+@CR+
                '       Cross join '+@CR+
                '       (select Top 1 B_SP_PDate+'' ~ ''+ E_SP_PDate as Date_Range '+@CR+
                '          from '+@TB_head_tmp+@CR+
                '       ) D1 '+@CR+
                ' where 1=1 '+@CR+
                ' order by Order_no '
  
  Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL

print '7'
  Set @Msg = '進行 凌越 比對 XLS 商品編號並寫入 ['+@TB_OD_Name+']資料表。'
  set @strSQL = 'insert into '+@TB_head+@CR+
                'select ''LYS'' as master, '+@CR+
                '       isnull(Chg_sp_pdate_ym, xls_ym) as XLS_YM,  '+@CR+
                '       Date_Range, '+@CR+
                '       isnull(sd_slip_fg, slip_fg) as sd_slip_fg, '+@CR+
                '       case when rtrim(isnull(sd_slip_fg, slip_fg)) = ''2'' then ''出'' else ''退'' end as xls_slip_fg_name, '+@CR+
                '       Case '+@CR+
                '         when order_no '+@Collate+' is null then rtrim(substring(sd_spec, 1, 20)) '+@Collate+' '+@CR+
                '         else rtrim(substring(order_no, 1, 20)) '+@Collate+' '+@CR+
                '       end as order_no, '+@CR+
                '       Convert(Varchar(255), isnull(Chg_sp_slip_name, '''')) as sd_spec, '+@CR+
                '       Round(isnull(xls_tot, 0), 0) as xls_tot, Round(isnull(chg_sd_stot, 0), 0) as chg_sd_stot, '+@CR+
                '       0 as order_total, '+@CR+
                '       isnull(cnt, 0) as order_cnt, '+@CR+
                '       case '+@CR+
                '         when xls_tot = chg_sd_stot then ''N'' '+@CR+
                '         else ''Y'' '+@CR+
                '       end as Equal '+@CR+
                '  from (select chg_sp_pdate_YM, sd_slip_fg, chg_sd_slip_fg, '+@CR+
                '               sd_spec, Chg_sp_slip_name, '+@CR+
                '               Round(sum(chg_sd_stot * 1.05), 0) as chg_sd_stot '+@CR+
                '          from Fact_sslpdt m '+@CR+
                '         where sd_class = ''1'' '+@CR+
                '           and sd_ctno ='''+@ct_no+''' '+@CR+

                '           and sp_pdate >= '+@CR+
                '               (select top 1 B_sp_pdate from '+@TB_head_tmp+') '+@CR+
                '           and sp_pdate <= '+@CR+
                '               (select top 1 E_sp_pdate from '+@TB_head_tmp+') '+@CR+

                '         group by chg_sp_pdate_YM, sd_slip_fg, chg_sd_slip_fg, sd_spec, Chg_sp_slip_name '+@CR+
                '       ) m '+@CR+
                '       Left join '+@CR+
                '       (select Order_no, Convert(Varchar(7), E_sp_pdate, 111) as xls_YM, B_sp_pdate, E_sp_pdate, '+@CR+
                '               slip_fg, sum(xls_tot) as xls_tot, count(1) as cnt  '+@CR+
                '          from '+@TB_head_tmp+' '+@CR+
                '         group by slip_fg, Order_no, Convert(Varchar(7), E_sp_pdate, 111),B_sp_pdate, E_sp_pdate '+@CR+
                '       ) d '+@CR+
                '         on Rtrim(sd_spec) '+@Collate+' = Order_no '+@Collate+' '+@CR+
                '        and sd_slip_fg '+@Collate+' = slip_fg '+@Collate+' '+@CR+
                '       Cross join '+@CR+
                '       (select Top 1 B_SP_PDate+'' ~ ''+ E_SP_PDate as Date_Range '+@CR+
                '          from '+@TB_head_tmp+@CR+
                '       ) D1 '+@CR+                ' where 1=1 '+@CR+
                '   and not exists '+@CR+
                '      (select * '+@CR+
                '         from '+@TB_head+' D1 '+@CR+
                '        where 1=1 '+@CR+
                '          and Rtrim(m.sd_spec) '+@Collate+' = D1.Order_no '+@Collate+' '+@CR+
                '          and m.sd_slip_fg '+@Collate+' = D1.slip_fg '+@Collate+' '+@CR+
                '      ) '+@CR+
                ' order by Order_no '
  Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL

print '8'
  Set @Msg = '依網通單號更新最後結餘金額。'
  set @strSQL =  -- 不限期間對應網通編號，並將單據編號予以填入
                'update '+@TB_head+' '+@CR+
                '   set sd_spec =d.sd_spec, '+@CR+
                '       Chg_sd_stot=d.chg_sd_stot, '+@CR+ 
                '       order_total=d.order_total '+@CR+
                '  from '+@TB_head+' m '+@CR+
                '       inner join '+@CR+
                '       (select m.order_no, slip_fg, '+@CR+
                '               sd_spec =''*''+d.Chg_sp_slip_name, '+@CR+
                '               Chg_sd_stot=Round(sum(d.chg_sd_stot * 1.05), 0), '+@CR+
                '               order_total=Round(sum(m.xls_tot) - Round(sum(d.chg_sd_stot * 1.05), 0), 0) '+@CR+
                '          from '+@TB_head+' m '+@CR+
                '               left join Fact_sslpdt d '+@CR+
                '                 on m.slip_fg '+@Collate+' = d.sd_slip_fg '+@Collate+' '+@CR+
                '                and m.order_no '+@Collate+'= d.sd_spec '+@Collate+' '+@CR+
                '                and d.sd_ctno ='''+@ct_no+''' '+@CR+
                '                and d.sd_class =''1'' '+@CR+
                '         where 1=1 '+@CR+
                '         group by m.order_no, slip_fg, d.Chg_sp_slip_name '+@CR+
                '       ) d '+@CR+
                '        on m.slip_fg '+@Collate+' = d.slip_fg '+@Collate+' '+@CR+
                '       and m.order_no '+@Collate+'= d.order_no '+@Collate+' '+@CR+
                '       and isnull(m.sd_spec, '''')= '''' '+@CR+
                '       and isnull(d.sd_spec, '''')<> '''' '+@CR+
                ' '+@CR+
                
                -- 依網通單號更新最後結餘金額
                'update '+@TB_head+' '+@CR+
                '   set order_total=isnull(m.total, 0), '+@CR+
                '       order_cnt=m.order_cnt '+@CR+
                '  from (select order_no, '+@CR+
                '               sum(isnull(xls_tot, 0) - isnull(chg_sd_stot, 0)) as total, '+@CR+
                '               order_cnt=Count(order_no) '+@CR+
                '          from '+@TB_head+' '+@CR+
                '         group by order_no '+@CR+
                '       ) m '+@CR+
                ' where '+@TB_head+'.order_no=m.order_no '+@CR+
                --'   and '+@TB_head+'.order_cnt = 0'+@CR+
                ' '+@CR+
               
                -- 2015/09/02 Rickliu 變更錯誤註記(PS, * 則代表跨月份異常資料)
                'update '+@TB_head+' '+@CR+
                '   set Equal = '+@CR+
                '         case '+@CR+
                '           when (order_total <> 0) or '+@CR+
                '                Rtrim(isnull(order_no, ''''))='''' or '+@CR+
                '                Rtrim(isnull(sd_spec, ''''))='''' or '+@CR+
                '                rtrim(isnull(sd_spec, '''')) like ''%*%'' '+@CR+
                '           then ''Y'' '+@CR+
                '           else ''N'' '+@CR+
                '         end '+@CR+
                ' '
                
  Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL

end
GO
