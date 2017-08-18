USE [DW]
GO
/****** Object:  StoredProcedure [dbo].[uSP_Realtime_SaleData]    Script Date: 08/18/2017 17:18:56 ******/
DROP PROCEDURE [dbo].[uSP_Realtime_SaleData]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE procedure [dbo].[uSP_Realtime_SaleData](@in_Kind Varchar(2000)='', @SendMail Int = 2)
as
begin

  Declare @Proc Varchar(100) = 'uSP_Realtime_SaleData'
  Declare @Cnt Int =0, @RowCnt Int = 0
  Declare @Err_Code int = -1, @Get_Result Int = 0, @Result Int = 0
  
  Declare @rDB Varchar(100)= 'SYNC_TA13.dbo.'

  Declare @TB_RT_Data Varchar(100) = 'Realtime_SaleData'
  Declare @TB_Target Varchar(100) = 'Target_Sales_Cust_Bkind_Amt'
  Declare @TB_RT_Data_tmp Varchar(100) = @TB_RT_Data+'_tmp'
  Declare @PDay_CT_Salse_Summary Varchar(100) = 'Sale_PDay_CT_Salse_Summary'

  Declare @SP_WTL_TB_RT_Data Varchar(100) = '' -- 'Exec uSP_Sys_Waiting_Table_Lock '''+@TB_RT_Data+''''
  Declare @SP_WTL_TB_RT_Data_tmp Varchar(100) = '' -- 'Exec uSP_Sys_Waiting_Table_Lock '''+@TB_RT_Data_tmp+''''

  Declare @Collate Varchar(100) = 'COLLATE Chinese_Taiwan_Stroke_CI_AS'
  -- 暫時還沒想到好的方法可以解決可能遇到 Table Lock 問題，因此先將 Table Hint 拿除不使用。
  -- 請勿使用 NoLock，會使得 Tempdb 不斷變大，且整個 I/O 讀取效能會降低很多。
  --Declare @TB_Hint Varchar(50) = ' WITH (NoLock)' 
  --Declare @TB_Hint Varchar(50) = ' With(NoWait)' --< 若要使用則前面空白不可拿掉
  Declare @TB_Hint Varchar(50) = ' ' --< 若要使用則前面空白不可拿掉
  
  Declare @NonSet Varchar(100) = '00000', @CNonSet Varchar(10) = '未設定'

  Declare @Kind Varchar(100) = '', @Kind_Name Varchar(1000) = ''
  Declare @Msg Varchar(4000) = '', @Remark Varchar(Max) = '', @strSQL Varchar(Max) = '', @strWhere Varchar(Max) = ''
  Declare @CR Varchar(4) = ' '+char(13)+char(10)

  Declare @RowCount Table (cnt int)
  Declare @BTime DateTime = Getdate(), @ETime DateTime = Getdate()


  -- 20151008 Rickliu Add Check Trans_Log Script
  Print 'Check Trans_Log Begin Scripts...'
  Print 'select * from trans_log where process ='''+@Proc+''' and trans_date >= '''+Convert(Varchar(100), @BTime, 120)+''' order by trans_date'

  -- 2015/10/21 Rickliu 根據此篇文章，進行初步優化 Where 條件，以加快作業速度
  -- http://fecbob.pixnet.net/blog/post/39076287-%E9%97%9C%E6%96%BCsqlserver%E5%BB%BA%E7%AB%8B%E7%B4%A2%E5%BC%95%E9%9C%80%E8%A6%81%E6%B3%A8%E6%84%8F%E7%9A%84%E5%95%8F%E9%A1%8C
  
  -- 2015/12/10 Rickliu Not in = 另一個寫法 and (欄位 <> Value and 欄位 <> Value), In = 另一個寫法 and (欄位 = Value Or 欄位 = Value)
  
  -- 20140930 Rickliu 為了避免 Lock 所以將 V_Sale_PDay_Summary, V_Sale_PDay_CT_Salse_Summary 導為實體 Table。
  -- Check 匯入檔案是否存在
  -- 20151008 Rickliu @SendMail = 2 當要寫入的LOG @Cnt 為 -1且又不要發送郵件時，只要將 @SendMail 設定不為 0, 1 的值即可。

  -- Check 匯入檔案是否存在

  set @Kind = '00011'
  if (@in_Kind in (@Kind) Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     Set @Msg = '重建資料表 ['+@PDay_CT_Salse_Summary+']'
     set @strSQL = 'IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N''[dbo].['+@PDay_CT_Salse_Summary+']'') AND type in (N''U''))'+@CR+
                   '   Drop Table '+@PDay_CT_Salse_Summary+@CR+
                   
                   'select * into '+@PDay_CT_Salse_Summary+@CR+
                   '  from '+@rDB+'V_'+@PDay_CT_Salse_Summary

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  -- Check 匯入檔案是否存在
  if (@in_Kind ='')
  begin
     Set @Msg = '清除臨時資料表 ['+@TB_RT_Data_tmp+']'
     set @strSQL = 'IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N''[dbo].['+@TB_RT_Data_tmp+']'') AND type in (N''U''))'+@CR+
                   '   Drop Table '+@TB_RT_Data_tmp+@CR

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     --Set @Msg = '清除資料表 ['+@TB_RT_Data+']'
     --set @strSQL = 'IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N''[dbo].['+@TB_RT_Data+']'') AND type in (N''U''))'+@CR+
     --              '   Truncate Table '+@TB_RT_Data+@CR

     --Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     --if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  Set @Msg = '重建 ['+@TB_RT_Data_tmp+'] 資料表'
  set @strSQL = 'IF Not EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N''[dbo].['+@TB_RT_Data_tmp+']'') AND type in (N''U''))'+@CR+
                '   CREATE TABLE [dbo].['+@TB_RT_Data_tmp+']('+@CR+
                '     [kind] [varchar](10) NOT NULL,'+@CR+
                '     [kind_name] [varchar](255) NULL,'+@CR+
                '     [Area_year] [int] NOT NULL,'+@CR+
                '     [Area_month] [int] NOT NULL,'+@CR+
                '     [Area] [Varchar](20) NOT NULL,'+@CR+
                '     [Area_name] [Varchar](50) NULL,'+@CR+
                '     [Amt] [float] NULL,'+@CR+
                '     [Not_Accumulate] [Varchar] (1) NULL,'+@CR+
                '     [Data_Type] [int] NULL,'+@CR+ -- 0:金額, 1:成本資料, 2:日期資料, 3:% 資料
                '     [Remark] [Varchar] (255) NULL'+@CR+
                '     Primary Key (Kind, area_year, area_month, area)'+@CR+
                '   )'

  Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
  if @Get_Result = @Err_Code Set @Result = @Err_Code
  
  Set @Msg = '重建 ['+@TB_RT_Data+'] 資料表 '
  set @strSQL = 'IF Not EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N''[dbo].['+@TB_RT_Data+']'') AND type in (N''U''))'+@CR+
                '   CREATE TABLE [dbo].['+@TB_RT_Data+']('+@CR+
                '     [kind] [varchar](10) NOT NULL,'+@CR+
                '     [kind_name] [varchar](255) NULL,'+@CR+
                '     [Area_year] [int] NOT NULL,'+@CR+
                '     [Area_month] [int] NOT NULL,'+@CR+
                '     [Area] [Varchar](20) NOT NULL,'+@CR+
                '     [Area_name] [Varchar](50) NULL,'+@CR+
                '     [Amt] [float] NULL,'+@CR+
                '     [Not_Accumulate] [Varchar] (1) NULL,'+@CR+
                '     [Data_Type] [int] NULL,'+@CR+ -- 0:金額, 1:成本資料, 2:日期資料, 3:% 資料
                '     [Remark] [Varchar] (255) NULL'+@CR+
                '     Primary Key (Kind, area_year, area_month, area)'+@CR+
                '   )'

  Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
  if @Get_Result = @Err_Code Set @Result = @Err_Code

--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--
  -- 產生 99999 個未設定資料，以利日後知道哪比資料沒產生
  if (@in_Kind ='')
  begin
     set @Msg ='產生 99999 個'+@CNonSet+'資料!!'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select Top 99999'+@CR+
                   '       Substring(Convert(Varchar(10), Row_Number() over(order by ct_no)+100000), 2, 5) as kind,'+@CR+
                   '       Substring(Convert(Varchar(10), Row_Number() over(order by ct_no)+100000), 2, 5) + ''.'+@CNonSet+''' '+@Collate+' as kind_name,'+@CR+
                   '       '''+@NonSet+''' as area_year,'+@CR+
                   '       ''00'' as area_month,'+@CR+
                   '       '''+@NonSet+''' as area,'+@CR+
                   '       '''+@CNonSet+''' as area_name,'+@CR+
                   '       0 as amt,'+@CR+
                   '       ''Y'' as Not_Accumulate,'+@CR+ 
                   '       0 as Data_Type,'+@CR+
                   '       '''' as Remark'+@CR+
                   '  from fact_pcust'+@TB_Hint

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00001'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月區域目標金額'
     set @Remark= '區域目標由 '+@TB_Target+' 而來'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       Convert(int, m.Year) as Area_year,'+@CR+
                   '       Convert(int, m.month) as Area_month,'+@CR+
                   '       Rtrim(IsNull(d2.TN_NO '+@Collate+', '''+@NonSet+''')) as area,'+@CR+
                   '       Rtrim(IsNull(d2.Tn_Contact, '''+@CNonSet+''')) as Area_name,'+@CR+
                   '       Sum(IsNull(m.amt, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_Target+' m'+@TB_Hint+@CR+
                   '       left join Fact_pcust d1 '+@TB_Hint+@CR+
                   '         on m.ct_no '+@Collate+' = d1.ct_no '+@Collate+@CR+
                   '       left join '+@rDB+'pattn d2 '+@TB_Hint+@CR+
                   '         on tn_class=''5'''+@CR+
                   '        and d1.ct_loc '+@Collate+' = d2.TN_NO '+@Collate+@CR+
                   ' where 1=1'+@CR+
                   -- 2017/08/08 Rickliu 副總指示客編 IT、ZZ 開頭不認列業績
                   '   and Chg_BU_NO Not In (''IT0000'', ''ZZ0000'') '+@CR+
                   ' group by m.Year, m.month, d2.TN_NO, d2.Tn_Contact  '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00002'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月客戶區域實銷金額'
     set @Remark= '依客戶檔內的區域別進行加總(請款日, 表頭, 含稅, 扣折讓, 無加減項)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_pdate_year,'+@CR+
                   '       chg_sp_pdate_month,'+@CR+
                   '       IsNull(ct_loc, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(chg_loc_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(chg_sp_stot_tax, 0))+Sum(IsNull(Chg_sp_dis_tot2, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslip M'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and chg_sp_pdate_year >= ''2013'''+@CR+
                   '   and (sp_slip_fg =''2'' Or sp_slip_fg =''C'' Or sp_slip_fg = ''3'')'+@CR+
                   '   and chg_sp_pdate_year is not null'+@CR+
                   -- 2017/08/08 Rickliu 副總指示客編 IT、ZZ 開頭不認列業績
                   '   and Chg_BU_NO Not In (''IT0000'', ''ZZ0000'') '+@CR+
                   ' group by chg_sp_pdate_year, chg_sp_pdate_month, ct_loc, chg_loc_name '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00003'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月部門業績目標金額'
     set @Remark= '依業務個人目標累計加總至單位部門'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       year '+@Collate+' as area_year,'+@CR+
                   '       month '+@Collate+' as area_month,'+@CR+
                   '       IsNull(e_dept, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(D.chg_dp_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(amt, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_Target+' m'+@TB_Hint+@CR+
                   '       left join Fact_pemploy D'+@TB_Hint+@CR+
                   '         on M.ct_sales '+@Collate+' =D.e_no'+@CR+
                   ' where 1=1'+@CR+
                   '   and isnull(e_dept, '''+@NonSet+''') is not null'+@CR+
                   ' group by m.year, m.month, D.e_dept, D.chg_dp_name'

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00004'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月部門客戶實銷金額'
     set @Remark= '依客戶檔內的業務部門進行加總(請款日, 表頭, 含稅, 扣折讓, 無加減項)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_pdate_year,'+@CR+
                   '       chg_sp_pdate_month,'+@CR+
                   '       IsNull(e_dept, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(chg_dp_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(chg_sp_stot_tax, 0))+Sum(IsNull(Chg_sp_dis_tot2, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslip M'+@TB_Hint+@CR+ --不能直接 sum 部門方式，因為 sp_slip_fg 類別不同，所以會導致資料產出後變成多個部門情況
                   ' where 1=1'+@CR+
                   '   and chg_sp_pdate_year >= ''2013'''+@CR+
                   '   and (sp_slip_fg =''2'' Or sp_slip_fg =''C'' Or sp_slip_fg = ''3'')'+@CR+
                   '   and chg_sp_pdate_year is not null'+@CR+
                   -- 2017/08/08 Rickliu 副總指示客編 IT、ZZ 開頭不認列業績
                   '   and Chg_BU_NO Not In (''IT0000'', ''ZZ0000'') '+@CR+
                   ' group by chg_sp_pdate_year, chg_sp_pdate_month, e_dept, chg_dp_name '
                   
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00005'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月業務業績目標金額'
     set @Remark= '業務目標從凌越的個人業績而來'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       year '+@Collate+' as area_year,'+@CR+
                   '       month '+@Collate+' as area_month,'+@CR+
                   '       IsNull(e_no, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(e_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(amt, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_Target+' m'+@TB_Hint+@CR+
                   '       left join Fact_pemploy D'+@TB_Hint+@CR+
                   '         on M.ct_sales '+@Collate+' =D.e_no'+@CR+
                   ' where 1=1'+@CR+
                   '   and isnull(e_name, '''+@NonSet+''') is not null'+@CR+
                   ' group by m.year, m.month, D.e_no, D.e_name'

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00006'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月業務實銷金額(個人業績)'
     set @Remark= '(請款日, 表身, 含稅, 含服務單, 扣折讓, 無加減項)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_pdate_year,'+@CR+
                   '       chg_sp_pdate_month,'+@CR+
                   '       IsNull(sp_sales, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(chg_sales_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(chg_sd_stot_tax, 0))+Sum(IsNull(Chg_sp_dis_tot2, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslpdt M'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and chg_sp_pdate_year >= ''2013'''+@CR+
                   '   and (sd_slip_fg =''2'' Or sd_slip_fg =''C'' Or sd_slip_fg = ''3'')'+@CR+
                   '   and chg_sp_pdate_year is not null'+@CR+
                   -- 2017/08/08 Rickliu 副總指示客編 IT、ZZ 開頭不認列業績
                   '   and Chg_BU_NO Not In (''IT0000'', ''ZZ0000'') '+@CR+
                   ' group by chg_sp_pdate_year, chg_sp_pdate_month, sp_sales, chg_sales_name '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00007'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月業務銷售金額(個人業績)'
     set @Remark= '(依請款日, 表頭, 含稅, 含服務單, 無折讓)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_pdate_year,'+@CR+
                   '       chg_sp_pdate_month,'+@CR+
                   '       IsNull(sp_sales, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(chg_sales_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(chg_sp_stot_tax, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslip M'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and chg_sp_pdate_year >= ''2013'''+@CR+
                   '   and (sp_slip_fg =''2'' Or sp_slip_fg =''C'')'+@CR+
                   '   and chg_sp_pdate_year is not null'+@CR+
                   -- 2017/08/08 Rickliu 副總指示客編 IT、ZZ 開頭不認列業績
                   '   and Chg_BU_NO Not In (''IT0000'', ''ZZ0000'') '+@CR+
                   ' group by chg_sp_pdate_year, chg_sp_pdate_month, sp_sales, chg_sales_name '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00008'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月業務退貨金額(個人業績)'
     set @Remark= '(請款日, 表頭, 含稅, 無折讓, 無加減項)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_pdate_year,'+@CR+
                   '       chg_sp_pdate_month,'+@CR+
                   '       IsNull(sp_sales, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(chg_sales_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(chg_sp_stot_tax, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslip M'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and chg_sp_pdate_year >= ''2013'''+@CR+
                   '   and sp_slip_fg = ''3'''+@CR+
                   '   and chg_sp_pdate_year is not null'+@CR+
                   -- 2017/08/08 Rickliu 副總指示客編 IT、ZZ 開頭不認列業績
                   '   and Chg_BU_NO Not In (''IT0000'', ''ZZ0000'') '+@CR+
                   ' group by chg_sp_pdate_year, chg_sp_pdate_month, sp_sales, chg_sales_name '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00009'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月業務新品銷貨金額(個人業績)'
     set @Remark= '(請款日, 表身, 含稅, 無折讓, 無加減項)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_pdate_year,'+@CR+
                   '       chg_sp_pdate_month,'+@CR+
                   '       IsNull(sp_sales, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(chg_sales_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(chg_sd_stot_tax, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslpdt'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and chg_sp_pdate_year >= ''2013'''+@CR+
                   '   and chg_is_new_stock=''Y'''+@CR+
                   '   and (sd_slip_fg =''2'' Or sd_slip_fg =''C'')'+@CR+
                   '   and chg_sp_pdate_year is not null'+@CR+
                   -- 2017/08/08 Rickliu 副總指示客編 IT、ZZ 開頭不認列業績
                   '   and Chg_BU_NO Not In (''IT0000'', ''ZZ0000'') '+@CR+
                   ' group by chg_sp_pdate_year, chg_sp_pdate_month, sp_sales, chg_sales_name '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00010'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月業務新品退貨金額(個人業績)'
     set @Remark= '(請款日, 表身, 含稅, 無折讓, 無加減項)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_pdate_year,'+@CR+
                   '       chg_sp_pdate_month,'+@CR+
                   '       IsNull(sp_sales, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(chg_sales_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(chg_sd_stot_tax, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslpdt'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and chg_sp_pdate_year >= ''2013'''+@CR+
                   '   and chg_is_new_stock=''Y'''+@CR+
                   '   and sd_slip_fg =''3'''+@CR+
                   '   and chg_sp_pdate_year is not null'+@CR+
                   -- 2017/08/08 Rickliu 副總指示客編 IT、ZZ 開頭不認列業績
                   '   and Chg_BU_NO Not In (''IT0000'', ''ZZ0000'') '+@CR+
                   ' group by chg_sp_pdate_year, chg_sp_pdate_month, sp_sales, chg_sales_name '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00011'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月業務客戶應收款(含退貨)(客戶業績)'
     set @Remark= '(依請款日, 表頭, 含稅, 含加減項, 含退貨, 含服務單, 應收款不為零)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       sp_year,'+@CR+
                   '       sp_month,'+@CR+
                   '       IsNull(e_no, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(e_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(sum_recamt, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from (select ct_no8, sp_ctno, ct_sname, m.e_no, m.e_name, e_dept, sp_slip_fg,'+@CR+
                   '               sp_year, sp_month,'+@CR+
                   '               sum_tot, sum_tax, sum_pamt, sum_mamt, sum_dis,'+@CR+ 
                   '               sum_all, sum_recamt, sum_payamt, ct_rem'+@CR+
                   '         from '+@PDay_CT_Salse_Summary+' m'+@TB_Hint+@CR+                   
                   '        where 1=1'+@CR+
                   '          and (sp_slip_fg =''2'' Or sp_slip_fg =''C'' Or sp_slip_fg = ''3'')'+@CR+
                   '          and sum_recamt <> 0'+@CR+
                   -- 2017/08/08 Rickliu 副總指示客編 IT、ZZ 開頭不認列業績
                   '          and Chg_BU_NO Not In (''IT0000'', ''ZZ0000'') '+@CR+
                   '       ) m'+@CR+
                   ' group by sp_year, sp_month, e_no, e_name '

     
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00012'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月客戶總公司實銷金額'
     set @Remark= '(請款日, 表頭, 含稅, 扣折讓, 無加減項)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_pdate_year,'+@CR+
                   '       chg_sp_pdate_month,'+@CR+
                   '       Chg_BU_NO as area,'+@CR+
                   '       IsNull(ct_fld3, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(chg_sp_stot_tax, 0))+Sum(IsNull(Chg_sp_dis_tot2, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslip'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and (sp_slip_fg =''2'' Or sp_slip_fg =''C'' Or sp_slip_fg = ''3'')'+@CR+
                   '   and chg_sp_pdate_year is not null'+@CR+
                   -- 2017/08/08 Rickliu 副總指示客編 IT、ZZ 開頭不認列業績
                   '   and Chg_BU_NO Not In (''IT0000'', ''ZZ0000'') '+@CR+
                   ' group by chg_sp_pdate_year, chg_sp_pdate_month, Chg_BU_NO, IsNull(ct_fld3, '''+@CNonSet+''')'+@CR+
                   'having Sum(chg_sp_stot_tax) <> 0 '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00013'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月業績金額達成率'
     set @Remark= '每月達成率 Formula: 00004 / 00003'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     --Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     --set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
     --              'insert into '+@TB_RT_Data_tmp+@CR+
     --              'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
     --              '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
     --              '       M.Area_year,'+@CR+
     --              '       M.area_month,'+@CR+
     --              '       '''+@NonSet+''' as area,'+@CR+
     --              '       '''+@CNonSet+''' as area_name,'+@CR+
     --              '       Case'+@CR+
     --              '         when Sum(Isnull(M.Amt, 0)) = 0 then 0'+@CR+
     --              '         else Round(Sum(Isnull(D.Amt, 0)) /Sum(Isnull(M.Amt, 0)), 2)'+@CR+
     --              '       end as Amt,'+@CR+
     --              '       '''' as Not_Accumulate,'+@CR+
     --              '       3 as Data_Type,'+@CR+
     --              '       '''+@ReMark+''' as Remark'+@CR+
     --              '  from (select area_year, area_month, Sum(Isnull(amt, 0)) as amt'+@CR+
     --              '          from '+@TB_RT_Data_tmp+@TB_Hint+@CR+
     --              '         where 1=1'+@CR+
     --              '           and kind =''00003'''+@CR+
     --              '         group by area_year, area_month'+@CR+
     --              '       ) M'+@CR+
     --              '       left join'+@CR+ 
     --              '       (select area_year, area_month, Sum(Isnull(amt, 0)) as amt'+@CR+
     --              '          from '+@TB_RT_Data_tmp+@TB_Hint+@CR+
     --              '         where 1=1'+@CR+
     --              '           and kind =''00004'''+@CR+
     --              '         group by area_year, area_month'+@CR+
     --              '       ) D'+@CR+
     --              '         on M.Area_year=D.Area_year'+@CR+
     --              '        and M.Area_month=D.Area_month'+@CR+
     --              ' group by m.Area_year, m.Area_month'

     --Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     --if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00014'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月目標金額差額'
     set @Remark= '每月目標差額 Formula: 00004 - 00003'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       M.Area_year,'+@CR+
                   '       M.area_month,'+@CR+
                   '       '''+@NonSet+''' as area,'+@CR+
                   '       '''+@CNonSet+''' as area_name,'+@CR+
                   '       Sum(Isnull(D.Amt, 0))-Sum(Isnull(M.Amt, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from (select area_year, area_month, Sum(Isnull(amt, 0)) as amt'+@CR+
                   '          from '+@TB_RT_Data_tmp+@TB_Hint+@CR+
                   '         where 1=1'+@CR+
                   '           and kind =''00003'''+@CR+
                   '         group by area_year, area_month'+@CR+
                   '       ) M'+@CR+
                   '       left join'+@CR+ 
                   '       (select area_year, area_month, Sum(Isnull(amt, 0)) as amt'+@CR+
                   '          from '+@TB_RT_Data_tmp+@TB_Hint+@CR+
                   '         where 1=1'+@CR+
                   '           and kind =''00004'''+@CR+
                   '         group by area_year, area_month'+@CR+
                   '       ) D'+@CR+
                   '         on M.Area_year=D.Area_year'+@CR+
                   '        and M.Area_month=D.Area_month'+@CR+
                   ' group by m.Area_year, m.Area_month'

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00015'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月客戶總公司貢獻實銷金額'
     set @Remark= '由 0012 加總而來'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       M.Area_year,'+@CR+
                   '       M.area_month,'+@CR+
                   '       IsNull(m.area, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(m.area_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(Isnull(m.amt, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and m.kind = ''00012'''+@CR+
                   ' group by m.area_year, m.area_month,'+@CR+
                   '       m.area, m.area_name '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00016'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月客戶總公司貢獻實銷金額佔比'
     set @Remark= '每月客戶總公司貢獻實銷金額佔比 Formula: 00012 / Sum(00012)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       M.Area_year,'+@CR+
                   '       M.area_month,'+@CR+
                   '       IsNull(m.area, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(m.area_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Round((Sum(Isnull(m.amt, 0)) / Isnull(tot_amt, 0)), 4) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       3 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '       inner join'+@CR+
                   '       (select area_year, Sum(Isnull(amt, 0)) as tot_amt'+@CR+
                   '          from '+@TB_RT_Data_tmp+@TB_Hint+@CR+
                   '         where 1=1'+@CR+
                   '           and kind = ''00012'''+@CR+
                   '         group by area_year'+@CR+
                   '       ) d'+@CR+
                   '        on m.area_year = d.area_year'+@CR+
                   ' where 1=1'+@CR+
                   '   and m.kind = ''00012'''+@CR+
                   ' group by m.area_year, m.area_month,'+@CR+
                   '       m.area, m.area_name, d.tot_amt '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00019'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月業務客戶新品銷售數量(個人業績)'
     set @Remark= '(請款日, 表身, 新品註記, 含服務單)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_pdate_year,'+@CR+
                   '       chg_sp_pdate_month,'+@CR+
                   '       IsNull(sp_sales, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(chg_sales_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(chg_sd_qty, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslpdt'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and chg_sp_pdate_year >= ''2013'''+@CR+
                   '   and chg_is_new_stock=''Y'''+@CR+
                   '   and (sd_slip_fg =''2'' Or sd_slip_fg =''C'')'+@CR+
                   '   and chg_sp_pdate_year is not null'+@CR+
                   -- 2017/08/08 Rickliu 副總指示客編 IT、ZZ 開頭不認列業績
                   '   and Chg_BU_NO Not In (''IT0000'', ''ZZ0000'') '+@CR+
                   ' group by chg_sp_pdate_year, chg_sp_pdate_month, sp_sales, chg_sales_name '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00020'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月業務客戶新品退貨數量(個人業績)'
     set @Remark= '(請款日, 表身, 新品註記)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_pdate_year,'+@CR+
                   '       chg_sp_pdate_month,'+@CR+
                   '       IsNull(sp_sales, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(chg_sales_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(chg_sd_qty, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslpdt'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and chg_sp_pdate_year >= ''2013'''+@CR+
                   '   and chg_is_new_stock=''Y'''+@CR+
                   '   and sd_slip_fg =''3'''+@CR+
                   '   and chg_sp_pdate_year is not null'+@CR+
                   -- 2017/08/08 Rickliu 副總指示客編 IT、ZZ 開頭不認列業績
                   '   and Chg_BU_NO Not In (''IT0000'', ''ZZ0000'') '+@CR+
                   ' group by chg_sp_pdate_year, chg_sp_pdate_month, sp_sales, chg_sales_name '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00021'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月新品實銷數量'
     set @Remark= '(依請款日, 表身, 含退貨, 含服務單, 新品註記)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_pdate_year as area_year,'+@CR+
                   '       chg_sp_pdate_month as area_month,'+@CR+
                   '       IsNull(sd_skno, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(sd_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(chg_sd_qty, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslpdt'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and chg_is_new_stock=''Y'''+@CR+
                   '   and (sd_slip_fg =''2'' Or sd_slip_fg =''C'' Or sd_slip_fg = ''3'')'+@CR+
                   '   and chg_sp_pdate_year is not null'+@CR+
                   ' group by chg_sp_pdate_year, chg_sp_pdate_month, sd_skno, sd_name '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00022'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月新品銷售數量'
     set @Remark= '(依請款日, 表身, 含服務單, 新品註記)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_pdate_year as area_year,'+@CR+
                   '       chg_sp_pdate_month as area_month,'+@CR+
                   '       IsNull(sd_skno, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(sd_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(chg_sd_qty, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslpdt'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and chg_is_new_stock=''Y'''+@CR+
                   '   and (sd_slip_fg =''2'' Or sd_slip_fg =''C'')'+@CR+
                   '   and chg_sp_pdate_year is not null'+@CR+
                   ' group by chg_sp_pdate_year, chg_sp_pdate_month, sd_skno, sd_name '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00023'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月新品退貨數量'
     set @Remark= '(依請款日, 表身, 新品註記)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_pdate_year as area_year,'+@CR+
                   '       chg_sp_pdate_month as area_month,'+@CR+
                   '       IsNull(sd_skno, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(sd_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(chg_sd_qty, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslpdt'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and chg_is_new_stock=''Y'''+@CR+
                   '   and sd_slip_fg =''3'''+@CR+
                   '   and chg_sp_pdate_year is not null'+@CR+
                   ' group by chg_sp_pdate_year, chg_sp_pdate_month, sd_skno, sd_name '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00024'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '近三個月新品實銷數量'
     set @Remark= '(抓新品日期往前推三個月, 表身, 含退貨, 含服務單, 新品註記)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       Year(DateAdd(mm, -3, getdate())) as area_year,'+@CR+
                   '       Month(DateAdd(mm, -3, getdate())) as area_month,'+@CR+
                   '       IsNull(sd_skno, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(sd_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(chg_sd_qty, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslpdt'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and sk_size >= Convert(Varchar(6), DateAdd(mm, -3, getdate()), 112)'+@CR+
                   '   and chg_is_new_stock=''Y'''+@CR+
                   '   and (sd_slip_fg =''2'' Or sd_slip_fg =''C'' Or sd_slip_fg = ''3'')'+@CR+
                   ' group by sd_skno, sd_name '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00025'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '近三個月新品銷售數量'
     set @Remark= '(抓新品日期往前推三個月, 表身, 含服務單, 新品註記)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       Year(DateAdd(mm, -3, getdate())) as area_year,'+@CR+
                   '       Month(DateAdd(mm, -3, getdate())) as area_month,'+@CR+
                   '       IsNull(sd_skno, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(sd_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(chg_sd_qty, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslpdt'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and sk_size >= Convert(Varchar(6), DateAdd(mm, -3, getdate()), 112)'+@CR+
                   '   and chg_is_new_stock=''Y'''+@CR+
                   '   and (sd_slip_fg =''2'' Or sd_slip_fg =''C'')'+@CR+
                   ' group by sd_skno, sd_name '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00026'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '近三個月新品退貨數量'
     set @Remark= '(抓新品日期往前推三個月, 表身, 新品註記)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       Year(DateAdd(mm, -3, getdate())) as area_year,'+@CR+
                   '       Month(DateAdd(mm, -3, getdate())) as area_month,'+@CR+
                   '       IsNull(sd_skno, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(sd_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(chg_sd_qty, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslpdt'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and sk_size >= Convert(Varchar(6), DateAdd(mm, -3, getdate()), 112)'+@CR+
                   '   and chg_is_new_stock=''Y'''+@CR+
                   '   and sd_slip_fg =''3'''+@CR+
                   ' group by sd_skno, sd_name '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00030'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月業務客戶銷售金額(個人業績)'
     set @Remark= '(依請款日, 表頭, 含稅, 含退貨單, 服務單, 扣折讓)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_pdate_year,'+@CR+
                   '       chg_sp_pdate_month,'+@CR+
                   '       IsNull(sp_sales, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(chg_sales_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(chg_sp_stot_tax, 0))+Sum(IsNull(Chg_sp_dis_tot2, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslip M'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and chg_sp_pdate_year >= ''2013'''+@CR+
                   '   and (sp_slip_fg =''2'' Or sp_slip_fg =''3'' Or sp_slip_fg =''C'')'+@CR+
                   '   and chg_sp_pdate_year is not null'+@CR+
                   -- 2017/08/08 Rickliu 副總指示客編 IT、ZZ 開頭不認列業績
                   '   and Chg_BU_NO Not In (''IT0000'', ''ZZ0000'') '+@CR+
                   ' group by chg_sp_pdate_year, chg_sp_pdate_month, sp_sales, chg_sales_name '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00031' --< 原 00117
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月網通客戶實銷金額'
     set @Remark= '(依請款日, 表頭, 含稅, 扣折讓, 無加減項, 含退貨, 含服務單, 網通客戶註記)且限定網通客戶'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_pdate_year,'+@CR+
                   '       chg_sp_pdate_month,'+@CR+
                   '       IsNull(sp_ctno, ''0000000'') as area,'+@CR+
                   '       IsNull(ct_sname, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(chg_sp_stot_tax, 0))+Sum(IsNull(Chg_sp_dis_tot2, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslip'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and (sp_slip_fg =''2'' Or sp_slip_fg =''C'' Or sp_slip_fg = ''3'')'+@CR+
                   '   and chg_is_lan_custom =''Y'''+@CR+
                   '   and chg_sp_pdate_year is not null'+@CR+
                   -- 2017/08/08 Rickliu 副總指示客編 IT、ZZ 開頭不認列業績
                   '   and Chg_BU_NO Not In (''IT0000'', ''ZZ0000'') '+@CR+
                   ' group by chg_sp_pdate_year, chg_sp_pdate_month, IsNull(sp_ctno, ''0000000''), IsNull(ct_sname, '''+@CNonSet+''')'+@CR+
                   'having Sum(chg_sp_stot_tax) <> 0 '
                   
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00032'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月網通客戶實銷金額(總公司)'
     set @Remark= '(請款日, 表頭, 含稅, 扣折讓, 無加減項)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_pdate_year,'+@CR+
                   '       chg_sp_pdate_month,'+@CR+
                   '       Chg_BU_NO as area,'+@CR+
                   '       IsNull(ct_fld3, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(chg_sp_stot_tax, 0))+Sum(IsNull(Chg_sp_dis_tot2, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslip'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and (sp_slip_fg =''2'' Or sp_slip_fg =''C'' Or sp_slip_fg = ''3'')'+@CR+
                   '   and chg_is_lan_custom =''Y'''+@CR+
                   '   and chg_sp_pdate_year is not null'+@CR+
                   -- 2017/08/08 Rickliu 副總指示客編 IT、ZZ 開頭不認列業績
                   '   and Chg_BU_NO Not In (''IT0000'', ''ZZ0000'') '+@CR+
                   ' group by chg_sp_pdate_year, chg_sp_pdate_month, Chg_BU_NO, IsNull(ct_fld3, '''+@CNonSet+''')'+@CR+
                   'having Sum(chg_sp_stot_tax) <> 0 '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00033' --< 原 00104
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月網通客戶實銷金額達成率(總公司)'
     set @Remark=  '(依請款日) 僅限網通客戶 Sum(00032) / 網通客戶月目標'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       M.Area_year,'+@CR+
                   '       M.area_month,'+@CR+
                   '       IsNull(m.area, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(m.area_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       case'+@CR+
                   '         when Isnull(d.amt, 0) = 0 then 0'+@CR+
                   '         else Round(Sum(Isnull(m.amt, 0)) / d.amt, 4)'+@CR+
                   '       end as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       3 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '       inner join'+@CR+
                   '       (select YEAR, MONTH, substring(ct_no, 1, 6) as ct_no, sum(amt) as amt'+@CR+
                   '          from '+@TB_Target+' as m'+@TB_Hint+@CR+
                   '               left join fact_pemploy as d on m.ct_sales  '+@Collate+'  = d.e_no  '+@Collate+@CR+
                   '         where d.e_dept=''B2300'''+@CR+
                   '         group by substring(ct_no, 1, 6), YEAR, MONTH'+@CR+
                   '       ) as d'+@CR+
                   '         on m.area_year = d.year'+@CR+
                   '        and m.area_month = d.month'+@CR+
                   '        and m.area = d.ct_no'+@CR+
                   '  where 1=1'+@CR+
                   '    and m.kind = ''00032'''+@CR+
                   '  group by m.area_year, m.area_month, m.area, m.area_name, d.amt '
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00034' --< 原 00105
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每年網通客戶實銷金額達成率(總公司)'
     set @Remark= '實銷達成率 Formula: Sum(00032) / 網通客戶年目標'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       M.Area_year,'+@CR+
                   '       ''00'' as area_month,'+@CR+
                   '       IsNull(m.area, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(m.area_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       case'+@CR+
                   '         when Isnull(d.Target_Amt, 0) = 0 then 0'+@CR+
                   '         else Round(Sum(Isnull(m.amt, 0)) / d.Target_Amt, 4)'+@CR+
                   '       end as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       3 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '       inner join'+@CR+
                   '       (select YEAR as area_year, substring(ct_no, 1, 6) as area, sum(amt) as Target_Amt'+@CR+
                   '          from '+@TB_Target+' as m'+@TB_Hint+@CR+
                   '               left join fact_pemploy as d on m.ct_sales  '+@Collate+'  = d.e_no  '+@Collate+@CR+
                   '         where d.e_dept=''B2300'''+@CR+
                   '         group by substring(ct_no, 1, 6), YEAR'+@CR+
                   '       ) d'+@CR+
                   '        on m.area_year = d.area_year'+@CR+
                   '       and m.area = d.area'+@CR+
                   ' where 1=1'+@CR+
                   '   and m.kind = ''00032'''+@CR+
                   ' group by m.area_year, m.area, m.area_name, d.Target_Amt '
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00048'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '年度客戶分公司數量(以總公司做群組)'
     set @Remark= '客戶總公司編號則以客戶編號前 6 碼做統計, 排除 已關店, 公司名稱空白'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       Year(Getdate()) as area_Year,'+@CR+
                   '       ''00'' as area_month,'+@CR+
                   '       IsNull(Rtrim(Ltrim(Chg_BU_NO)), ''0000000'') as area,'+@CR+
                   '       Rtrim(Ltrim(ct_fld3)) as area_name,'+@CR+
                   '       count(distinct Isnull(ct_no8, '''')) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_pcust'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and chg_ct_close = '''' '+@CR+
                   '   and Rtrim(Ltrim(Chg_BU_NO)) <> '''' '+@CR+
                   ' group by Rtrim(Ltrim(Chg_BU_NO)), Rtrim(Ltrim(ct_fld3)) '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end 

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00049'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '年度業務客戶分公司數量(以客戶的業務做統計)'
     set @Remark= '客戶總公司編號則以客戶編號前 8 碼做統計，業務客戶數量則取客戶基本資料的業務為準，排除關店'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       Year(Getdate()) as area_Year,'+@CR+
                   '       ''00'' as area_month,'+@CR+
                   '       IsNull(Rtrim(Ltrim(ct_sales)), '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(Rtrim(Ltrim(chg_ct_sales_name)), '''+@CNonSet+''') as area_name,'+@CR+
                   '       count(distinct Isnull(ct_no8, '''')) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_pcust'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and chg_ct_close = '''' '+@CR+
                   -- 2017/08/08 Rickliu 副總指示客編 IT、ZZ 開頭不認列業績
                   '   and Chg_BU_NO Not In (''IT0000'', ''ZZ0000'') '+@CR+
                   ' group by Rtrim(Ltrim(ct_sales)), Rtrim(Ltrim(chg_ct_sales_name)) '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end 

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00050'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '年度區域別客戶分公司數量'
     set @Remark= '客戶總公司編號則以客戶編號前 8 碼做統計，客戶區域則以客戶基本資料為準，排除關店'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       Year(Getdate()) as area_Year,'+@CR+
                   '       ''00'' as area_month,'+@CR+
                   '       IsNull(Rtrim(Ltrim(ct_loc)), '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(Rtrim(Ltrim(chg_loc_name)), '''+@CNonSet+''') as area_name,'+@CR+
                   '       count(distinct Isnull(ct_no8, '''')) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_pcust'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and chg_ct_close = '''' '+@CR+
                   -- 2017/08/08 Rickliu 副總指示客編 IT、ZZ 開頭不認列業績
                   '   and Chg_BU_NO Not In (''IT0000'', ''ZZ0000'') '+@CR+
                   ' group by Rtrim(Ltrim(ct_loc)), Rtrim(Ltrim(chg_loc_name)) '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end 
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00051'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月商品實銷數量'
     set @Remark= '(依請款日, 表身, 含退貨, 含服務單)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_pdate_year as area_year,'+@CR+
                   '       chg_sp_pdate_month as area_month,'+@CR+
                   '       IsNull(sd_skno, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(sd_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(chg_sd_qty, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslpdt'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and (sd_slip_fg =''2'' Or sd_slip_fg =''C'' Or sd_slip_fg = ''3'')'+@CR+
                   '   and chg_sp_pdate_year is not null'+@CR+
                   ' group by chg_sp_pdate_year, chg_sp_pdate_month, sd_skno, sd_name '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00052'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月商品銷售數量'
     set @Remark= '(依請款日, 表身, 含服務單)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_pdate_year as area_year,'+@CR+
                   '       chg_sp_pdate_month as area_month,'+@CR+
                   '       IsNull(sd_skno, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(sd_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(chg_sd_qty, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslpdt'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and (sd_slip_fg =''2'' Or sd_slip_fg =''C'')'+@CR+
                   '   and chg_sp_pdate_year is not null'+@CR+
                   ' group by chg_sp_pdate_year, chg_sp_pdate_month, sd_skno, sd_name '
                   
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00053'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月商品退貨數量'
     set @Remark= '(依請款日, 表身)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_pdate_year as area_year,'+@CR+
                   '       chg_sp_pdate_month as area_month,'+@CR+
                   '       IsNull(sd_skno, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(sd_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(chg_sd_qty, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslpdt'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and sd_slip_fg =''3'''+@CR+
                   '   and chg_sp_pdate_year is not null'+@CR+
                   ' group by chg_sp_pdate_year, chg_sp_pdate_month, sd_skno, sd_name '
                   
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00054'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月商品實銷金額'
     set @Remark= '(依請款日, 表身, 含稅, 無折讓, 無加減項, 含退貨, 含服務單)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_pdate_year as area_year,'+@CR+
                   '       chg_sp_pdate_month as area_month,'+@CR+
                   '       IsNull(sd_skno, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(sd_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(chg_sd_stot_tax, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslpdt'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and (sd_slip_fg =''2'' Or sd_slip_fg =''C'' Or sd_slip_fg = ''3'')'+@CR+
                   '   and chg_sp_pdate_year is not null'+@CR+
                   ' group by chg_sp_pdate_year, chg_sp_pdate_month, sd_skno, sd_name '
                   
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00055'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月商品銷售金額'
     set @Remark= '(依請款日, 表身, 含稅, 無折讓, 無加減項, 含退貨)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_pdate_year as area_year,'+@CR+
                   '       chg_sp_pdate_month as area_month,'+@CR+
                   '       IsNull(sd_skno, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(sd_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(chg_sd_stot_tax, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslpdt'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and (sd_slip_fg =''2'' Or sd_slip_fg =''C'')'+@CR+
                   '   and chg_sp_pdate_year is not null'+@CR+
                   ' group by chg_sp_pdate_year, chg_sp_pdate_month, sd_skno, sd_name '
                   
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00056'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月商品退貨金額'
     set @Remark= '(依請款日, 表身, 含稅, 無折讓, 無加減項)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_pdate_year as area_year,'+@CR+
                   '       chg_sp_pdate_month as area_month,'+@CR+
                   '       IsNull(sd_skno, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(sd_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(chg_sd_stot_tax, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslpdt'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and sd_slip_fg =''3'''+@CR+
                   '   and chg_sp_pdate_year is not null'+@CR+
                   ' group by chg_sp_pdate_year, chg_sp_pdate_month, sd_skno, sd_name '
                   
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00057'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '年度業績目標金額'
     set @Remark= '年度業績目標是由 Kind: 0003 加總而來, 2013 年目標是由 Ori_XLS#Target_Before_Sale_By_2013 而來'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       area_year,'+@CR+
                   '       ''00'' as area_month,'+@CR+
                   '       '''+@NonSet+''' as area,'+@CR+
                   '       '''+@CNonSet+''' as area_name,'+@CR+
                   '       Sum(IsNull(amt, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and kind =''00003'''+@CR+
                   ' group by area_year' --+@CR+
                   /*
                   ' union'+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       year as area_year,'+@CR+
                   '       ''00'' as area_month,'+@CR+
                   '       '''+@NonSet+''' as area,'+@CR+
                   '       '''+@CNonSet+''' as area_name,'+@CR+
                   '       Sum(IsNull(Target_amt, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from Ori_XLS#Target_Before_Sale_By_2012'+@TB_Hint+@CR+
                   ' group by year '
                   */
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00058'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '年度實銷業績金額'
     set @Remark= '年度實銷業績是從 0006 加總而來，2012年的銷售金額則由 Ori_XLS#Target_Before_Sale_By_2012 加總而來'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       area_year,'+@CR+
                   '       ''00'' as area_month,'+@CR+
                   '       '''+@NonSet+''' as area,'+@CR+
                   '       '''+@CNonSet+''' as area_name,'+@CR+
                   '       Sum(IsNull(amt, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and kind =''00006'''+@CR+
                   ' group by area_year' --+@CR+
                   /*
                   ' union'+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       year as area_year,'+@CR+
                   '       ''00'' as area_month,'+@CR+
                   '       '''+@NonSet+''' as area,'+@CR+
                   '       '''+@CNonSet+''' as area_name,'+@CR+
                   '       Sum(IsNull(Sale_amt, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from Ori_XLS#Target_Before_Sale_By_2012'+@TB_Hint+@CR+
                   ' group by year '
                   */
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00059'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '年度實銷業績達成率'
     set @Remark= '達成率 Formula: 00058 / 00057'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       m.area_year,'+@CR+
                   '       ''00'' as area_month,'+@CR+
                   '       '''+@NonSet+''' as area,'+@CR+
                   '       '''+@CNonSet+''' as area_name,'+@CR+
                   '       case'+@CR+
                   '         when Isnull(d.tot_amt, 0) = 0 then 0'+@CR+
                   '         else Round(Sum(Isnull(m.amt, 0)) / d.tot_amt, 4)'+@CR+
                   '       end as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       3 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '       inner join'+@CR+
                   '       (select area_year, Isnull(amt, 0) as tot_amt'+@CR+
                   '          from '+@TB_RT_Data_tmp+@TB_Hint+@CR+
                   '         where kind =''00057'''+@CR+
                   '       ) d'+@CR+
                   '        on m.area_year=d.area_year'+@CR+
                   ' where 1=1'+@CR+
                   '   and kind=''00058'''+@CR+
                   ' group by m.area_year, tot_amt '
                   
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00060'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月公司業績目標金額'
     set @Remark= '每月公司業績目標由 0003 加總而來, 2012 年以前的目標從 Ori_XLS#Target_Before_Sale_By_2012 加總而來'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       area_year,'+@CR+
                   '       area_month,'+@CR+
                   '       '''+@NonSet+''' as area,'+@CR+
                   '       '''+@CNonSet+''' as area_name,'+@CR+
                   '       Sum(IsNull(amt, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and kind =''00003'''+@CR+
                   ' group by area_year, area_month' --+@CR+
                   /******
                   ' union'+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       year as area_year,'+@CR+
                   '       month as area_month,'+@CR+
                   '       '''+@NonSet+''' as area,'+@CR+
                   '       '''+@CNonSet+''' as area_name,'+@CR+
                   '       IsNull(Target_amt, 0) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from Ori_XLS#Target_Before_Sale_By_2012 '+@TB_Hint+' '
                   ******/
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00061'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月公司實銷業績金額'
     set @Remark= '每月公司實銷業績由 00006 加總而來, 2012 年以前則 Ori_XLS#Target_Before_Sale_By_2012 加總而來'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       area_year,'+@CR+
                   '       area_month,'+@CR+
                   '       '''+@NonSet+''' as area,'+@CR+
                   '       '''+@CNonSet+''' as area_name,'+@CR+
                   '       Sum(IsNull(amt, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and kind =''00006'''+@CR+
                   ' group by area_year, area_month' --+@CR+
                   /****
                   ' union'+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       year as area_year,'+@CR+
                   '       month as area_month,'+@CR+
                   '       '''+@NonSet+''' as area,'+@CR+
                   '       '''+@CNonSet+''' as area_name,'+@CR+
                   '       Sum(IsNull(Target_amt, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from Ori_XLS#Target_Before_Sale_By_2012'+@TB_Hint+@CR+
                   ' group by year, month '
                   ****/
                   
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00062'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月公司實銷業績金額佔比(個人業績)'
     set @Remark= '每月實銷業績佔比 Formula: 00061 / 00058'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       m.area_year,'+@CR+
                   '       m.area_month,'+@CR+
                   '       '''+@NonSet+''' as area,'+@CR+
                   '       '''+@CNonSet+''' as area_name,'+@CR+
                   '       case'+@CR+
                   '         when Isnull(d.tot_amt, 0) = 0 then 0'+@CR+
                   '         else Round(Sum(Isnull(m.amt, 0)) / d.tot_amt, 4)'+@CR+
                   '       end as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       3 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '       inner join'+@CR+
                   '       (select area_year, Isnull(amt, 0) as tot_amt'+@CR+
                   '          from '+@TB_RT_Data_tmp+@TB_Hint+@CR+
                   '         where 1=1'+@CR+
                   '           and kind =''00058'''+@CR+
                   '       ) d'+@CR+
                   '        on m.area_year=d.area_year'+@CR+
                   ' where 1=1'+@CR+
                   '   and kind=''00061'''+@CR+
                   ' group by m.area_year, m.area_month, tot_amt '
                   
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
  set @Kind = '00063'
     set @Kind_Name = '每月公司實銷業績金額達成率(個人業績)'
     set @Remark= '每月實銷業績達成率, Formula: 00061 / 00060'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       m.area_year,'+@CR+
                   '       m.area_month,'+@CR+
                   '       '''+@NonSet+''' as area,'+@CR+
                   '       '''+@CNonSet+''' as area_name,'+@CR+
                   '       case'+@CR+
                   '         when Isnull(d.tot_amt, 0) = 0 then 0'+@CR+
                   '         else Round(Sum(Isnull(m.amt, 0)) / d.tot_amt, 4)'+@CR+
                   '       end as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       3 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '       inner join'+@CR+
                   '       (select area_year, area_month, Isnull(amt, 0) as tot_amt'+@CR+
                   '          from '+@TB_RT_Data_tmp+@TB_Hint+@CR+
                   '         where 1=1'+@CR+
                   '           and kind =''00060'''+@CR+
                   '       ) d'+@CR+
                   '        on m.area_year=d.area_year'+@CR+
                   '       and m.area_month=d.area_month'+@CR+
                   ' where 1=1'+@CR+
                   '   and kind=''00061'''+@CR+
                   ' group by m.area_year, m.area_month, tot_amt '
                   
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  -- Kind 64 保留給 百大目標
  set @Kind = '00064'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = ''+@CNonSet+'保留給年度百大目標金額'
     set @Remark= @Kind_Name
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       '''+@NonSet+''' as area_year,'+@CR+
                   '       ''00'' as area_month,'+@CR+
                   '       '''+@NonSet+''' as area,'+@CR+
                   '       '''+@CNonSet+''' as area_name,'+@CR+
                   '       0 as amt,'+@CR+
                   '       ''Y'' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark '
                   --'  from fact_sslpdt m'+@TB_Hint+@CR+
                   --' where 1=1'+@CR+
                   
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00065'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '年度各百大實銷業績金額(依總公司總和)'
     set @Remark= '(依請款日, 表頭, 含稅, 扣折讓, 無加減項, 含退貨, 含服務單, 百大註記)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       m.chg_sp_pdate_year as area_year,'+@CR+
                   '       ''00'' as area_month,'+@CR+
                   '       chg_bu_no as area,'+@CR+
                   '       IsNull(ct_fld3, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(chg_sp_stot_tax, 0))+Sum(IsNull(Chg_sp_dis_tot2, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslip m'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and chg_sp_pdate_year >= ''2013'''+@CR+
                   '   and (sp_slip_fg =''2'' Or sp_slip_fg =''C'' Or sp_slip_fg = ''3'')'+@CR+
                   '   and chg_hunderd_customer=''Y'''+@CR+
                   -- 2017/08/08 Rickliu 副總指示客編 IT、ZZ 開頭不認列業績
                   '   and Chg_BU_NO Not In (''IT0000'', ''ZZ0000'') '+@CR+
                   ' group by m.chg_sp_pdate_year, chg_bu_no, IsNull(ct_fld3, '''+@CNonSet+''') '
                   
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00066'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '年度百大實銷總業績金額(依年度總和)'
     set @Remark= '由 0065 加總而來'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       area_year,'+@CR+
                   '       ''00'' as area_month,'+@CR+
                   '       '''+@NonSet+''' as area,'+@CR+
                   '       '''+@CNonSet+''' as area_name,'+@CR+
                   '       Round(Sum(amt), 0) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and kind=''00065'''+@CR+
                   ' group by m.area_year '
                   
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00067'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '年度百大實銷總業績金額(依月份總和)'
     set @Remark= '(依請款日, 表頭, 含稅, 扣折讓, 無加減項, 含退貨, 含服務單, 百大註記)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       m.chg_sp_pdate_year as area_year,'+@CR+
                   '       m.chg_sp_pdate_month as area_month,'+@CR+
                   '       ''00'' as area,'+@CR+
                   '       '''+@CNonSet+''' as area_name,'+@CR+
                   '       Sum(IsNull(chg_sp_stot_tax, 0))+Sum(IsNull(Chg_sp_dis_tot2, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslip m'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and chg_sp_pdate_year >= ''2013'''+@CR+
                   '   and (sp_slip_fg =''2'' Or sp_slip_fg =''C'' Or sp_slip_fg = ''3'')'+@CR+
                   '   and chg_hunderd_customer=''Y'''+@CR+
                   '   and m.chg_sp_pdate_year is not null'+@CR+
                   -- 2017/08/08 Rickliu 副總指示客編 IT、ZZ 開頭不認列業績
                   '   and Chg_BU_NO Not In (''IT0000'', ''ZZ0000'') '+@CR+
                   ' group by m.chg_sp_pdate_year, m.chg_sp_pdate_month '
                   
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00068'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月各百大實銷業績金額(依總公司)'
     set @Remark= '(依請款日, 表頭, 含稅, 扣折讓, 無加減項, 含退貨, 含服務單, 百大註記)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       m.chg_sp_pdate_year as area_year,'+@CR+
                   '       m.chg_sp_pdate_month as area_month,'+@CR+
                   '       chg_bu_no as area,'+@CR+
                   '       IsNull(ct_fld3, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(chg_sp_stot_tax, 0))+Sum(IsNull(Chg_sp_dis_tot2, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslip m'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and chg_sp_pdate_year >= ''2013'''+@CR+
                   '   and (sp_slip_fg =''2'' Or sp_slip_fg =''C'' Or sp_slip_fg = ''3'')'+@CR+
                   '   and chg_hunderd_customer=''Y'''+@CR+
                   '   and m.chg_sp_pdate_year is not null'+@CR+
                   -- 2017/08/08 Rickliu 副總指示客編 IT、ZZ 開頭不認列業績
                   '   and Chg_BU_NO Not In (''IT0000'', ''ZZ0000'') '+@CR+
                   ' group by m.chg_sp_pdate_year, m.chg_sp_pdate_month, chg_bu_no, IsNull(ct_fld3, '''+@CNonSet+''') '
                   
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00069'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月各百大實銷業績金額佔百大實銷總業績金額佔比'
     set @Remark= '總業績佔比 Formula: 00068 / Sum(00068)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       m.area_year,'+@CR+
                   '       m.area_month,'+@CR+
                   '       IsNull(m.area, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(m.area_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       case'+@CR+
                   '         when Isnull(d.tot_amt, 0) = 0 then 0'+@CR+
                   '         else Round(Isnull(m.amt, 0) / d.tot_amt, 4)'+@CR+
                   '       end as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       3 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '       inner join'+@CR+
                   '       (select area_year, area_month, Sum(Isnull(amt, 0)) as tot_amt'+@CR+
                   '          from '+@TB_RT_Data_tmp+@TB_Hint+@CR+
                   '         where 1=1'+@CR+
                   '           and kind =''00068'''+@CR+
                   '         group by area_year, area_month'+@CR+
                   '       ) d'+@CR+
                   '        on m.area_year=d.area_year'+@CR+
                   '       and m.area_month=d.area_month'+@CR+
                   ' where 1=1'+@CR+
                   '   and m.kind=''00068'' '
                   
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00070'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月滯銷品實銷數量'
     set @Remark= '(依請款日, 表身, 含退貨, 含服務單, 滯銷品註記, 銷貨年月 >= 滯銷年月)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_pdate_year as area_year,'+@CR+
                   '       chg_sp_pdate_month as area_month,'+@CR+
                   '       IsNull(sd_skno, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(sd_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(chg_sd_qty, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslpdt'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and chg_sp_pdate_year >= ''2013'''+@CR+
                   '   and Chg_IS_Dead_Stock = ''Y'''+@CR+
                   '   and chg_sp_date_YM >= Chg_Dead_Stock_YM'+@CR+
                   '   and (sd_slip_fg =''2'' Or sd_slip_fg =''C'' Or sd_slip_fg = ''3'')'+@CR+
                   '   and chg_sp_pdate_year is not null'+@CR+
                   ' group by chg_sp_pdate_year, chg_sp_pdate_month, sd_skno, sd_name '
                   
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00071'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月滯銷品實銷金額'
     set @Remark= '(依請款日, 表身, 含稅, 無折讓, 無加減項, 含退貨, 含服務單, 滯銷品註記, 銷貨年月 >= 滯銷年月)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_pdate_year as area_year,'+@CR+
                   '       chg_sp_pdate_month as area_month,'+@CR+
                   '       IsNull(sd_skno, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(sd_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(chg_sd_stot_tax, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslpdt'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and chg_sp_pdate_year >= ''2013'''+@CR+
                   '   and Chg_IS_Dead_Stock = ''Y'''+@CR+
                   '   and chg_sp_date_YM >= Chg_Dead_Stock_YM'+@CR+
                   '   and (sd_slip_fg =''2'' Or sd_slip_fg =''C'' Or sd_slip_fg = ''3'')'+@CR+
                   '   and chg_sp_pdate_year is not null'+@CR+
                   ' group by chg_sp_pdate_year, chg_sp_pdate_month, sd_skno, sd_name '
                   
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00072'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月滯銷品實銷數量達成率'
     set @Remark= '達成率 Formula: 00070 / 滯銷商品數量'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       m.area_year,'+@CR+
                   '       '''+@NonSet+''' as area_month,'+@CR+
                   '       IsNull(m.area, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(area_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       case'+@CR+
                   '         when Isnull(d.Chg_Dead_First_Qty, 0) = 0 then 0'+@CR+
                   '         else Round(Sum(IsNull(amt, 0)) / d.Chg_Dead_First_Qty, 4)'+@CR+
                   '       end as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       3 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '       inner join Fact_sstock d'+@TB_Hint+@CR+
                   '          on m.area_year = Substring(d.Chg_Dead_Stock_YM, 1, 4)'+@CR+
                   '         and m.area = d.sk_no '+@Collate+''+@CR+
                   ' where 1=1'+@CR+
                   '   and kind = ''00070'''+@CR+
                   ' group by m.area_year, m.area, m.area_name, d.Chg_Dead_First_Qty '
                   
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00083'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月業務實銷金額達成率(個人業績)'
     set @Remark= '每月業務實銷金額達成率 Formula: 00006 / 00005'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       m.area_year,'+@CR+
                   '       m.area_month,'+@CR+
                   '       IsNull(m.area, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(m.area_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Case'+@CR+
                   '         when (Sum(Isnull(M.Amt, 0)) = 0) then 0'+@CR+
                   '         else IsNull(Round(Sum(Isnull(D.Amt, 0)) / Sum(M.Amt), 4), 0)'+@CR+
                   '       end as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       3 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from (select area_year, area_month, area, area_name, Sum(IsNull(amt, 0)) as amt'+@CR+
                   '          from '+@TB_RT_Data_tmp+@TB_Hint+@CR+
                   '         where 1=1'+@CR+
                   '           and kind =''00005'''+@CR+
                   '         group by area_year, area_month, area, area_name'+@CR+
                   '       ) M'+@CR+
                   '       inner join'+@CR+
                   '       (select area_year, area_month, area, area_name, Sum(IsNull(amt, 0)) as amt'+@CR+
                   '          from '+@TB_RT_Data_tmp+@TB_Hint+@CR+
                   '         where 1=1'+@CR+
                   '           and kind =''00006'''+@CR+
                   '         group by area_year, area_month, area, area_name'+@CR+
                   '       ) D'+@CR+
                   '         on M.Area_year=D.Area_year'+@CR+
                   '        and M.Area_month=D.Area_month'+@CR+
                   '        and M.Area=D.Area'+@CR+
                   ' group by m.Area_year, m.Area_month, m.Area, m.area_name '
                   
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00084'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月業務退貨金額佔比(個人業績)'
     set @Remark= '每月業務退貨金額佔比 Formula: 00008 / Sum(00008)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       m.area_year,'+@CR+
                   '       m.area_month,'+@CR+
                   '       IsNull(m.area, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(m.area_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Case'+@CR+
                   '         when (IsNull(D.Tot_amt, 0) = 0) then 0'+@CR+
                   '         else IsNull(Round(Sum(Isnull(M.Amt, 0)) / d.Tot_amt, 4), 0)'+@CR+
                   '       end as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       3 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from (select area_year, area_month, area, area_name, Sum(IsNull(amt, 0)) as amt'+@CR+
                   '          from '+@TB_RT_Data_tmp+@TB_Hint+@CR+
                   '         where 1=1'+@CR+
                   '           and kind =''00008'''+@CR+
                   '         group by area_year, area_month, area, area_name'+@CR+
                   '       ) M'+@CR+
                   '       inner join'+@CR+
                   '       (select area_year, area_month, Sum(IsNull(amt, 0)) as Tot_amt'+@CR+
                   '          from '+@TB_RT_Data_tmp+@TB_Hint+@CR+
                   '         where 1=1'+@CR+
                   '           and kind =''00008'''+@CR+
                   '         group by area_year, area_month'+@CR+
                   '       ) D'+@CR+
                   '         on M.Area_year=D.Area_year'+@CR+
                   '        and M.Area_month=D.Area_month'+@CR+
                   ' group by m.Area_year, m.Area_month, m.Area, m.area_name, d.Tot_amt '
                   
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00106'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月新品銷售金額佔比'
     set @Remark= '每月新品銷售佔比 Formula: Sum(00009) / Sum(00007)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       m.area_year,'+@CR+
                   '       m.area_month,'+@CR+
                   '       '''+@NonSet+''' as area,'+@CR+
                   '       '''+@CNonSet+''' as area_name,'+@CR+
                   '       case'+@CR+
                   '         when IsNull(m.amt, 0) = 0 then 0'+@CR+
                   '         else Round(Isnull(d.amt, 0) / m.amt, 4)'+@CR+
                   '       end as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       3 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  From (select area_year, area_month, Sum(IsNull(amt, 0)) as amt'+@CR+
                   '          from '+@TB_RT_Data_tmp+@TB_Hint+@CR+
                   '         where kind=''00007'''+@CR+
                   '         group by area_year, area_month'+@CR+
                   '       )m'+@CR+
                   '       inner join'+@CR+
                   '       (select area_year, area_month, Sum(IsNull(amt, 0)) as amt'+@CR+
                   '          from '+@TB_RT_Data_tmp+@TB_Hint+@CR+
                   '         where kind=''00009'''+@CR+
                   '         group by area_year, area_month'+@CR+
                   '      )d'+@CR+
                   '      on m.area_year = d.area_year'+@CR+
                   '     and m.area_month = d.area_month '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00107'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月新品退貨金額佔比'
     set @Remark= '每月新品退貨佔比 Formula: Sum(00010) / Sum(00007)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       m.area_year,'+@CR+
                   '       m.area_month,'+@CR+
                   '       '''+@NonSet+''' as area,'+@CR+
                   '       '''+@CNonSet+''' as area_name,'+@CR+
                   '       case'+@CR+
                   '         when IsNull(m.amt, 0) = 0 then 0'+@CR+
                   '         else Round(Isnull(d.amt, 0) / m.amt, 4)'+@CR+
                   '       end as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       3 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  From (select area_year, area_month, Sum(IsNull(amt, 0)) as amt'+@CR+
                   '          from '+@TB_RT_Data_tmp+@TB_Hint+@CR+
                   '         where kind=''00007'''+@CR+
                   '         group by area_year, area_month'+@CR+
                   '       )m'+@CR+
                   '       inner join'+@CR+
                   '       (select area_year, area_month, Sum(IsNull(amt, 0)) as amt'+@CR+
                   '          from '+@TB_RT_Data_tmp+@TB_Hint+@CR+
                   '         where kind=''00010'''+@CR+
                   '         group by area_year, area_month'+@CR+
                   '      )d'+@CR+
                   '      on m.area_year = d.area_year'+@CR+
                   '     and m.area_month = d.area_month '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00108'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每年新品銷貨金額佔比'
     set @Remark= '每年新品銷貨佔比 Formula: Sum(00009) / Sum(00007)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       m.area_year,'+@CR+
                   '       ''00'' as area_month,'+@CR+
                   '       '''+@NonSet+''' as area,'+@CR+
                   '       '''+@CNonSet+''' as area_name,'+@CR+
                   '       case'+@CR+
                   '         when IsNull(m.amt, 0) = 0 then 0'+@CR+
                   '         else Round(Isnull(d.amt, 0) / m.amt, 4)'+@CR+
                   '       end as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       3 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  From (select area_year, Sum(IsNull(amt, 0)) as amt'+@CR+
                   '          from '+@TB_RT_Data_tmp+@TB_Hint+@CR+
                   '         where kind=''00007'''+@CR+
                   '         group by area_year'+@CR+
                   '       )m'+@CR+
                   '       inner join'+@CR+
                   '       (select area_year, Sum(IsNull(amt, 0)) as amt'+@CR+
                   '          from '+@TB_RT_Data_tmp+@TB_Hint+@CR+
                   '         where kind=''00009'''+@CR+
                   '         group by area_year'+@CR+
                   '      )d'+@CR+
                   '      on m.area_year = d.area_year '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end  

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00109'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每年新品退貨金額佔比'
     set @Remark= '每年新品退貨佔比 Formula: Sum(00010) / Sum(00007)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       m.area_year,'+@CR+
                   '       ''00'' as area_month,'+@CR+
                   '       '''+@NonSet+''' as area,'+@CR+
                   '       '''+@CNonSet+''' as area_name,'+@CR+
                   '       case'+@CR+
                   '         when IsNull(m.amt, 0) = 0 then 0'+@CR+
                   '         else Round(Isnull(d.amt, 0) / m.amt, 4)'+@CR+
                   '       end as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       3 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  From (select area_year, Sum(IsNull(amt, 0)) as amt'+@CR+
                   '          from '+@TB_RT_Data_tmp+@TB_Hint+@CR+
                   '         where kind=''00007'''+@CR+
                   '         group by area_year'+@CR+
                   '       )m'+@CR+
                   '       inner join'+@CR+
                   '       (select area_year, Sum(IsNull(amt, 0)) as amt'+@CR+
                   '          from '+@TB_RT_Data_tmp+@TB_Hint+@CR+
                   '         where kind=''00010'''+@CR+
                   '         group by area_year'+@CR+
                   '      )d'+@CR+
                   '      on m.area_year = d.area_year '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end  

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00110'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '業務超過三個月每月未收款金額(客戶業績)'
     set @Remark= '由 0011 超過三個月未收款金額加總而來'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       area_year,'+@CR+
                   '       area_month,'+@CR+
                   '       IsNull(area, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(area_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Round(Sum(IsNull(amt, 0)), 0) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+@TB_Hint+@CR+
                   ' where kind =''00011'''+@CR+
                   '   and amt > 0'+@CR+
                   '   and area_year <= Year(Dateadd(mm, -3, getdate()))'+@CR+
                   '   and area_month <= month(Dateadd(mm, -3, getdate()))'+@CR+
                   -- 2017/08/08 Rickliu 副總指示客編 IT、ZZ 開頭不認列業績，雖不認列業績但應收扔得收
                   --'   and Chg_BU_NO Not In (''IT0000'', ''ZZ0000'') '+@CR+
                   ' group by area_year, area_month, area, area_name'+@CR+
                   'having Round(Sum(IsNull(amt, 0)), 0) <> 0 '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end  

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00111'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '超過三個月每月未收款金額(客戶業績)'
     set @Remark= '由 0110 加總而來'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       area_year,'+@CR+
                   '       area_month,'+@CR+
                   '       '''+@NonSet+''' as area,'+@CR+
                   '       '''+@CNonSet+''' as area_name,'+@CR+
                   '       Round(Sum(IsNull(amt, 0)), 0) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+@TB_Hint+@CR+
                   ' where kind =''00110'''+@CR+
                   '   and amt > 0'+@CR+
                   '   and area_year <= Year(Dateadd(mm, -3, getdate()))'+@CR+
                   '   and area_month <= month(Dateadd(mm, -3, getdate()))'+@CR+
                   ' group by area_year, area_month'+@CR+
                   'having Round(Sum(IsNull(amt, 0)), 0) <> 0 '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end  
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00112'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '業務超過三個月總未收款金額(客戶業績)'
     set @Remark= '由 0110 加總而來'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       year(getdate()) as area_year,'+@CR+
                   '       month(Dateadd(mm, -3, getdate())) as area_month,'+@CR+
                   '       IsNull(area, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(area_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(amt, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+@TB_Hint+@CR+
                   ' where kind =''00110'''+@CR+
                   ' group by area, area_name'+@CR+
                   'having Round(Sum(IsNull(amt, 0)), 0) <> 0 '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00113'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '超過三個月總未收款金額(客戶業績)'
     set @Remark= '由 0112 加總而來'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       year(getdate()) as area_year,'+@CR+
                   '       month(Dateadd(mm, -3, getdate())) as area_month,'+@CR+
                   '       '''+@NonSet+''' as area,'+@CR+
                   '       '''+@CNonSet+''' as area_name,'+@CR+
                   '       Round(Sum(IsNull(amt, 0)), 0) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+@TB_Hint+@CR+
                   ' where kind =''00112'''+@CR+
                   '   and amt > 0'+@CR+
                   'having Round(Sum(IsNull(amt, 0)), 0) <> 0 '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end  

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00114'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '業務超過三個月每月未收款佔比(客戶業績)'
     set @Remark= '業務超過三個月每月未收款佔比 Formula: 00110 / 00111'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       m.area_year,'+@CR+
                   '       m.area_month,'+@CR+
                   '       IsNull(m.area, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(m.area_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       case'+@CR+
                   '         when IsNull(d.amt, 0) = 0 then 0'+@CR+
                   '         else Round(Isnull(m.amt, 0) / d.amt, 4)'+@CR+
                   '       end as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       3 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '       inner join '+@TB_RT_Data_tmp+' d'+@TB_Hint+@CR+
                   '          on m.area_year = d.area_year'+@CR+
                   '         and m.area_month = d.area_month'+@CR+
                   '         and d.kind =''00111'''+@CR+
                   ' where m.kind =''00110'' '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end  
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00115'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '全部總應收款(客戶業績)'
     set @Remark= '由 0011 加總而來'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       Year(getdate()) as area_year,'+@CR+
                   '       month(getdate()) as area_month,'+@CR+
                   '       '''+@NonSet+''' as area,'+@CR+
                   '       '''+@CNonSet+''' as area_name,'+@CR+
                   '       Sum(IsNull(amt, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   ' where kind =''00011'' '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end 

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00116'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '超過三個月總未收款佔比(客戶業績)'
     set @Remark= '超過三個月總未收款佔比 Formula: 00113 / 00115'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       m.area_year,'+@CR+
                   '       m.area_month,'+@CR+
                   '       '''+@NonSet+''' as area,'+@CR+
                   '       '''+@CNonSet+''' as area_name,'+@CR+
                   '       case'+@CR+
                   '         when IsNull(d.amt, 0) = 0 then 0'+@CR+
                   '         else Round(Isnull(m.amt, 0) / d.amt, 4)'+@CR+
                   '       end as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       3 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '       inner join '+@TB_RT_Data_tmp+' d'+@TB_Hint+@CR+
                   '          on m.area_year = d.area_year'+@CR+
                   '         and d.kind =''00115'''+@CR+
                   ' where m.kind =''00113'' '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end 

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00118'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月客戶分店目標金額(業一、二課)'
     set @Remark= '網通目標由 Ori_XLS#Target_Cust_For_LAN 而來'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select distinct'+@CR+
                   '       '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       d1.cal_year as area_year,'+@CR+
                   '       d1.cal_month as area_month,'+@CR+
                   '       d1.Ct_no as area,'+@CR+
                   '       d1.ct_sName as area_name,'+@CR+
                   '       Isnull(Target_Amt, 0) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from (select distinct a.ct_no, a.ct_name, a.ct_sname, b.e_no, b.e_name, a.chg_bu_no, cal_month, cal_year'+@CR+
                   '          from Fact_pcust as a'+@TB_Hint+@CR+
                   '               left join Fact_pemploy as b '+@TB_Hint+@CR+
                   '                 on a.ct_sales = b.e_no and a.ct_class = ''1'''+@CR+
                   '               left join calendar as m'+@TB_Hint+@CR+
                   '                 on cal_year <= Year(getdate()) and cal_year >= ''2013'''+@CR+
                   '         where 1=1 and b.e_no > '''' and a.ct_no > '''' and len(a.ct_no) > 8 and a.Chg_IS_Lan_Custom =''Y'''+@CR+
                   '       ) d1 left join'+@CR+
                   '       (select Year,'+@CR+
                   '               Month,'+@CR+
                   '               Ltrim(Rtrim(ct_no)) as ct_no,'+@CR+
                   '               sum(Isnull(amt, 0)) as Target_Amt'+@CR+
                   '          from '+@TB_Target+@TB_Hint+@CR+
                   '         group by Year, Month, ct_no'+@CR+
                   '       ) d on D.ct_no = D1.ct_no '+@Collate+' and d1.cal_year=d.Year  and d1.cal_month = d.Month'+@CR+
                   ' where 1=1'+@CR+
                   '   and D1.chg_bu_no like ''I%'''+@CR+
                   '   and substring(D1.chg_bu_no, 2, 1) not like ''[ATZ]''' --+@CR+

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end 

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00119'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每年網通客戶目標金額'
     set @Remark= '由 0118 加總而來'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       area_Year,'+@CR+
                   '       ''00'' as area_month,'+@CR+
                   '       area,'+@CR+
                   '       area_name,'+@CR+
                   '       Sum(IsNull(amt, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and kind =''00118'''+@CR+
                   ' group by area_year, area, area_name '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end 

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00120'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月網通目標金額'
     set @Remark= '由 0118 加總而來'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       area_Year,'+@CR+
                   '       area_month,'+@CR+
                   '       '''+@NonSet+''' as area,'+@CR+
                   '       '''+@CNonSet+''' area_name,'+@CR+
                   '       Sum(IsNull(amt, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and kind =''00118'''+@CR+
                   ' group by area_year, area_month '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end 

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00121'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每年網通目標金額'
     set @Remark= '由 0118 加總而來'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       area_Year,'+@CR+
                   '       ''00'' as area_month,'+@CR+
                   '       '''+@NonSet+''' as area,'+@CR+
                   '       '''+@CNonSet+''' area_name,'+@CR+
                   '       Sum(IsNull(amt, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and kind =''00118'''+@CR+
                   ' group by area_year '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end 

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00122'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月到職人數'
     set @Remark= '(依到職日)員工編號第一碼為 (C, Z)不列入計算'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_e_rdate_YY as area_year,'+@CR+
                   '       chg_e_rdate_MM as area_month,'+@CR+
                   '       '''+@NonSet+''' as area,'+@CR+
                   '       '''+@CNonSet+''' area_name,'+@CR+
                   '       count(*) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_pemploy'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and (Substring(e_no, 1, 1) <> ''C'' and Substring(e_no, 1, 1) <> ''Z'')'+@CR+
                   ' group by chg_e_rdate_YY, chg_e_rdate_MM '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end 

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00123'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月離職人數'
     set @Remark= '(依離職日)員工編號第一碼為 (C, Z)不列入計算'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_e_ldate_YY as area_year,'+@CR+
                   '       chg_e_ldate_MM as area_month,'+@CR+
                   '       '''+@NonSet+''' as area,'+@CR+
                   '       '''+@CNonSet+''' area_name,'+@CR+
                   '       count(*)*-1 as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_pemploy'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and (Substring(e_no, 1, 1) <> ''C'' and Substring(e_no, 1, 1) <> ''Z'')'+@CR+
                   '   and chg_leave = ''Y'''+@CR+
                   ' group by chg_e_ldate_YY, chg_e_ldate_MM '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end 

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00124'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月異動人數'
     set @Remark= '由 0122, 0123 加總而來'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       area_year,'+@CR+
                   '       area_month,'+@CR+
                   '       '''+@NonSet+''' as area,'+@CR+
                   '       '''+@CNonSet+''' area_name,'+@CR+
                   '       Sum(IsNull(amt, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and (m.kind =''00122'' or m.kind =''00123'')'+@CR+
                   ' group by area_year, area_month '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end 
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00125'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月現職人數'
     set @Remark= '由 0124 加總而來且以累計方式計算'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       m.area_year,'+@CR+
                   '       m.area_month,'+@CR+
                   '       '''+@NonSet+''' as area,'+@CR+
                   '       '''+@CNonSet+''' area_name,'+@CR+
                   '       Sum(IsNull(d.amt, 0)) as amt,'+@CR+
                   '       ''Y'' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '       inner join '+@TB_RT_Data_tmp+' d'+@TB_Hint+@CR+
                   '          on m.kind = d.kind and m.kind =''00124'''+@CR+
                   ' where 1=1'+@CR+
                   '   and Convert(Varchar(4), d.area_year) + Substring(Convert(Varchar(3), d.area_month+100), 2, 2) <='+@CR+
                   '       Convert(Varchar(4), m.area_year) + Substring(Convert(Varchar(3), m.area_month+100), 2, 2)'+@CR+
                   ' group by m.area_year, m.area_month '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end 

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00126'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月客戶分店目標金額(業一、二課)'
     set @Remark= '目標值由 Ori_Xls#Target_Cust_For_B2 而來，總目標值 / 12 個月'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select distinct'+@CR+
                   '       '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       m.year as area_year,'+@CR+
                   '       m.month as area_month,'+@CR+
                   '       m.ct_no as area,'+@CR+
                   '       m.ct_Name as area_name,'+@CR+
                   '       sum(Isnull(amt, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from Target_Sales_Cust_Bkind_Amt m'+@TB_Hint+@CR+
                   '       left join Fact_pcust d '+@CR+
                   '         on m.ct_no collate Chinese_Taiwan_Stroke_CI_AS = d.ct_no collate Chinese_Taiwan_Stroke_CI_AS '+@CR+
                   ' where m.year >= ''2013'' and year <= Year(getdate()) '+@CR+
                   '   and d.Chg_IS_Lan_Custom =''N'' '+@CR+
                   ' group by m.Year, m.Month, m.ct_no, m.ct_name '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end 
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00127'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月客戶總公司目標金額(業一、二課)'
     set @Remark=  '每月客戶總公司目標金額(業一、二課)由 0126 而來'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       area_year,'+@CR+
                   '       area_month,'+@CR+
                   '       Substring(m.area, 1, 6) as area,'+@CR+
                   '       Rtrim(d.ct_fld3) as area_name,'+@CR+
                   '       Sum(isnull(amt, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '       left join'+@CR+
                   '       (select distinct chg_bu_no, ct_fld3'+@CR+
                   '          from fact_pcust'+@TB_Hint+@CR+
                   '         where chg_bu_no <> '''' '+@CR+
                   '       ) d'+@CR+
                   '        on Substring(m.area, 1, 6)=d.chg_bu_no '+@Collate+@CR+
                   ' where kind =''00126'''+@CR+
                   ' group by area_year, area_month, Substring(m.area, 1, 6), d.ct_fld3'

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end 

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00128'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每年客戶總公司目標金額(業一、二課)'
     set @Remark= '每年客戶總公司目標金額(業一、二課)由 0127 而來'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       area_year,'+@CR+
                   '       ''00'' as area_month,'+@CR+
                   '       area,'+@CR+
                   '       area_name,'+@CR+
                   '       Sum(Isnull(amt, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   ' where kind =''00127'''+@CR+
                   ' group by area_year, area, area_name '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end 

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00129'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月客戶分店實銷金額'
     set @Remark= '(請款日, 表頭, 含稅, 扣折讓, 無加減項, 含退貨, 含服務單)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_pdate_year,'+@CR+
                   '       chg_sp_pdate_month,'+@CR+
                   '       IsNull(sp_ctno, ''0000000'') as area,'+@CR+
                   '       IsNull(ct_sname, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(chg_sp_stot_tax, 0))+Sum(IsNull(Chg_sp_dis_tot2, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslip'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and (sp_slip_fg =''2'' Or sp_slip_fg =''C'' Or sp_slip_fg = ''3'')'+@CR+
                   '   and chg_sp_pdate_year is not null'+@CR+
                   ' group by chg_sp_pdate_year, chg_sp_pdate_month, IsNull(sp_ctno, ''0000000''), IsNull(ct_sname, '''+@CNonSet+''')'+@CR+
                   'having Sum(Isnull(chg_sp_stot_tax, 0)) <> 0  '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end 

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00130'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月客戶分店實銷金額達成率'
     set @Remark= '每月客戶分店實銷金額達成率 Formula: 00129 / 00126'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       m.area_year,'+@CR+
                   '       m.area_month,'+@CR+
                   '       m.area,'+@CR+
                   '       m.area_name,'+@CR+
                   '       case'+@CR+
                   '         when IsNull(m.amt, 0) = 0 then 0'+@CR+
                   '         else Round(Isnull(d.amt, 0) / m.amt, 4)'+@CR+
                   '       end as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       3 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '       inner join '+@TB_RT_Data_tmp+' d'+@TB_Hint+@CR+
                   '          on m.area_year = d.area_year'+@CR+
                   '         and m.area_month = d.area_month'+@CR+
                   '         and m.area = d.area'+@CR+
                   '         and d.kind =''00129'''+@CR+
                   ' where m.kind =''00126'' '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end 

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00131'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月客戶總公司實銷金額達成率'
     set @Remark= '每月客戶總公司實銷金額達成率 Formula: 00127 / 00012'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       m.area_year,'+@CR+
                   '       m.area_month,'+@CR+
                   '       m.area,'+@CR+
                   '       m.area_name,'+@CR+
                   '       case'+@CR+
                   '         when IsNull(m.amt, 0) = 0 then 0'+@CR+
                   '         else Round(Isnull(d.amt, 0) / m.amt, 4)'+@CR+
                   '       end as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       3 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '       inner join '+@TB_RT_Data_tmp+' d'+@TB_Hint+@CR+
                   '          on m.area_year = d.area_year'+@CR+
                   '         and m.area_month = d.area_month'+@CR+
                   '         and m.area = d.area'+@CR+
                   '         and d.kind =''00127'''+@CR+
                   ' where m.kind =''00012'' '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end 

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00132'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月客戶分店實銷金額累計平均達成率'
     set @Remark= '每月客戶分店實銷平均達成率 Formula: 00130 / 00130 且以累計方式計算'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       m.area_year,'+@CR+
                   '       m.area_month,'+@CR+
                   '       m.area,'+@CR+
                   '       m.area_name,'+@CR+
                   '       case'+@CR+
                   '         when Isnull(m.area_month, 0) = 0 then 0'+@CR+
                   '         else Round(Sum(Isnull(d.amt, 0)) / m.area_month, 4)'+@CR+
                   '       end as amt,'+@CR+
                   '       ''Y'' as Not_Accumulate,'+@CR+
                   '       3 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '       inner join '+@TB_RT_Data_tmp+' d'+@TB_Hint+@CR+
                   '          on m.kind = d.kind and m.area = d.area'+@CR+
                   ' where 1=1'+@CR+
                   '   and m.area_year = d.area_year'+@CR+
                   '   and d.area_month <= m.area_month'+@CR+
                   '   and m.kind =''00130'''+@CR+
                   ' group by m.area_year, m.area_month, m.area, m.area_name'+@CR+
                   ' order by 1, 2, 3 '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end 

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00133'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月客戶總公司實銷金額累計平均達成率'
     set @Remark= '每月客戶總公司實銷平均達成率 Formula: 00131 / 00131 且以累計方式計算'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       m.area_year,'+@CR+
                   '       m.area_month,'+@CR+
                   '       m.area,'+@CR+
                   '       m.area_name,'+@CR+
                   '       case'+@CR+
                   '         when Isnull(m.area_month, 0) = 0 then 0'+@CR+
                   '         else Round(Sum(Isnull(d.amt, 0)) / m.area_month, 4)'+@CR+
                   '       end as amt,'+@CR+
                   '       ''Y'' as Not_Accumulate,'+@CR+
                   '       3 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '       inner join '+@TB_RT_Data_tmp+' d'+@TB_Hint+@CR+
                   '          on m.kind = d.kind and m.area = d.area'+@CR+
                   ' where 1=1'+@CR+
                   '   and m.area_year = d.area_year'+@CR+
                   '   and d.area_month <= m.area_month'+@CR+
                   '   and m.kind =''00131'''+@CR+
                   ' group by m.area_year, m.area_month, m.area, m.area_name'+@CR+
                   ' order by 1, 2, 3 '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end 
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00134'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月廠商最後一次進貨日期'
     set @Remark= '(依銷貨日, 表頭)(值量欄位請以自行轉換成日期型態)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_date_year as area_year,'+@CR+
                   '       chg_sp_date_month as area_month,'+@CR+
                   '       sp_ctno as area,'+@CR+
                   '       ct_sname as area_name,'+@CR+
                   '       Convert(float, Max(sp_date)) as amt,'+@CR+
                   '       ''Y'' as Not_Accumulate,'+@CR+
                   '       2 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslip'+@TB_Hint+@CR+
                   ' where sp_slip_fg = ''0'''+@CR+
                   ' group by chg_sp_date_year, chg_sp_date_month, sp_ctno, ct_sname '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end 

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00135'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月廠商最後一次進貨退出日期'
     set @Remark= '(依銷貨日, 表頭)(值量欄位請以自行轉換成日期型態)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_date_year as area_year,'+@CR+
                   '       chg_sp_date_month as area_month,'+@CR+
                   '       sp_ctno as area,'+@CR+
                   '       ct_sname as area_name,'+@CR+
                   '       Convert(float, Max(sp_date)) as amt,'+@CR+
                   '       ''Y'' as Not_Accumulate,'+@CR+
                   '       2 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslip'+@TB_Hint+@CR+
                   ' where sp_slip_fg = ''1'''+@CR+
                   ' group by chg_sp_date_year, chg_sp_date_month, sp_ctno, ct_sname '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end 

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00136'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '廠商最後一次進貨日期'
     set @Remark= '(依銷貨日, 表頭)(值量欄位請以自行轉換成日期型態)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       max(chg_sp_date_year) as area_year,'+@CR+
                   '       ''00'' as area_month,'+@CR+
                   '       sp_ctno as area,'+@CR+
                   '       ct_sname as area_name,'+@CR+
                   '       Convert(float, Max(sp_date)) as amt,'+@CR+
                   '       ''Y'' as Not_Accumulate,'+@CR+
                   '       2 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslip'+@TB_Hint+@CR+
                   ' where sp_slip_fg = ''0'''+@CR+
                   ' group by sp_ctno, ct_sname '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end 

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00137'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '廠商最後一次進貨退出日期'
     set @Remark= '(依銷貨日, 表頭)(值量欄位請以自行轉換成日期型態)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       max(chg_sp_date_year) as area_year,'+@CR+
                   '       ''00'' as area_month,'+@CR+
                   '       sp_ctno as area,'+@CR+
                   '       ct_sname as area_name,'+@CR+
                   '       Convert(float, Max(sp_date)) as amt,'+@CR+
                   '       ''Y'' as Not_Accumulate,'+@CR+
                   '       2 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslip'+@TB_Hint+@CR+
                   ' where sp_slip_fg = ''1'''+@CR+
                   ' group by sp_ctno, ct_sname '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end 

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00138'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月客戶最後一次銷貨日期含服務單(以客戶八碼統計)'
     set @Remark= '(依銷貨日, 表頭, 含服務單)(值量欄位請以自行轉換成日期型態)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_date_year as area_year,'+@CR+
                   '       chg_sp_date_month as area_month,'+@CR+
                   '       ct_no8 as area,'+@CR+
                   '       ct_sname8 as area_name,'+@CR+
                   '       Convert(float, Max(sp_date)) as amt,'+@CR+
                   '       ''Y'' as Not_Accumulate,'+@CR+
                   '       2 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslip'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and (sp_slip_fg =''2'' Or sp_slip_fg =''C'')'+@CR+
                   '   and Len(RTrim(Isnull(ct_no8, ''''))) = 8'+@CR+
                   ' group by chg_sp_date_year, chg_sp_date_month, ct_no8, ct_sname8 '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end 

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00139'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月客戶最後一次銷貨退回日期(以客戶八碼統計)'
     set @Remark= '(依銷貨日, 表頭)(值量欄位請以自行轉換成日期型態)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_date_year as area_year,'+@CR+
                   '       chg_sp_date_month as area_month,'+@CR+
                   '       ct_no8 as area,'+@CR+
                   '       ct_sname8 as area_name,'+@CR+
                   '       Convert(float, Max(sp_date)) as amt,'+@CR+
                   '       ''Y'' as Not_Accumulate,'+@CR+
                   '       2 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslip'+@TB_Hint+@CR+
                   ' where sp_slip_fg = ''3'''+@CR+
                   '   and Len(RTrim(Isnull(ct_no8, ''''))) = 8'+@CR+
                   ' group by chg_sp_date_year, chg_sp_date_month, ct_no8, ct_sname8 '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end 

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00140'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '客戶最後一次銷貨日期含服務單(以客戶八碼統計)'
     set @Remark= '(依銷貨日, 表頭, 含服務單)(值量欄位請以自行轉換成日期型態)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       max(chg_sp_date_year) as area_year,'+@CR+
                   '       ''00'' as area_month,'+@CR+
                   '       ct_no8 as area,'+@CR+
                   '       ct_sname8 as area_name,'+@CR+
                   '       Convert(float, Max(sp_date)) as amt,'+@CR+
                   '       ''Y'' as Not_Accumulate,'+@CR+
                   '       2 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslip'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and (sp_slip_fg =''2'' Or sp_slip_fg =''C'')'+@CR+
                   '   and Len(RTrim(Isnull(ct_no8, ''''))) = 8'+@CR+
                   ' group by ct_no8, ct_sname8 '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end 
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00141'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '客戶最後一次銷貨退回日期含服務單(以客戶八碼統計)'
     set @Remark= '(依銷貨日, 表頭, 含服務單)(值量欄位請以自行轉換成日期型態)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       max(chg_sp_date_year) as area_year,'+@CR+
                   '       ''00'' as area_month,'+@CR+
                   '       ct_no8 as area,'+@CR+
                   '       ct_sname8 as area_name,'+@CR+
                   '       Convert(float, Max(sp_date)) as amt,'+@CR+
                   '       ''Y'' as Not_Accumulate,'+@CR+
                   '       2 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslip'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and sp_slip_fg =''3'''+@CR+
                   '   and Len(RTrim(Isnull(ct_no8, ''''))) = 8'+@CR+
                   ' group by ct_no8, ct_sname8 '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end 

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00142'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月各百大實銷業績金額佔當月總實銷業績金額佔比'
     set @Remark= '每月各百大實銷業績金額佔當月總實銷業績金額佔比 Formula: 00068 / Sum(00061)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       m.area_year,'+@CR+
                   '       m.area_month,'+@CR+
                   '       IsNull(m.area, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(m.area_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       case'+@CR+
                   '         when (IsNull(D.tot_amt, 0) <= 0) Or'+@CR+
                   '              (IsNull(M.Amt, 0) <= 0 And IsNull(D.tot_amt, 0) <= 0) then 0'+@CR+
                   '         else Round(Isnull(M.Amt, 0) / D.tot_amt, 4)'+@CR+
                   '       end as Amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       3 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '       inner join'+@CR+
                   '       (select area_year, area_month, Sum(amt) as tot_amt'+@CR+
                   '          from '+@TB_RT_Data_tmp+@TB_Hint+@CR+
                   '         where 1=1'+@CR+
                   '           and kind =''00061'''+@CR+
                   '         group by area_year, area_month'+@CR+
                   '       ) d'+@CR+
                   '        on m.area_year=d.area_year'+@CR+
                   '       and m.area_month=d.area_month'+@CR+
                   ' where 1=1'+@CR+
                   '   and kind=''00068'' '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00158'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月銷售實績金額平均達成率'
     set @Remark= '每月銷售實績金額平均達成率 Formula: Sum(00013) / 00013 並以累計方式計算'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       m.area_year,'+@CR+
                   '       m.area_month,'+@CR+
                   '       m.area,'+@CR+
                   '       m.area_name,'+@CR+
                   '       case'+@CR+
                   '         when IsNull(m.area_month, 0) = 0 then 0'+@CR+
                   '         else Round(Sum(Isnull(d.amt, 0)) / m.area_month, 4)'+@CR+
                   '       end as amt,'+@CR+
                   '       ''Y'' as Not_Accumulate,'+@CR+
                   '       3 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '       inner join '+@TB_RT_Data_tmp+' d'+@TB_Hint+@CR+
                   '          on m.kind = d.kind and m.area = d.area'+@CR+
                   ' where 1=1'+@CR+
                   '   and m.area_year = d.area_year'+@CR+
                   '   and d.area_month <= m.area_month'+@CR+
                   '   and m.kind =''00013'''+@CR+
                   ' group by m.area_year, m.area_month, m.area, m.area_name '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00159'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月業務實銷金額平均達成率'
     set @Remark= '每月業務實銷金額平均達成率 Formula: Sum(00083) / 00083 並以累計方式計算'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       m.area_year,'+@CR+
                   '       m.area_month,'+@CR+
                   '       m.area,'+@CR+
                   '       m.area_name,'+@CR+
                   '       case'+@CR+
                   '         when IsNull(m.area_month, 0) = 0 then 0'+@CR+
                   '         else Round(Sum(Isnull(d.amt, 0)) / m.area_month, 4)'+@CR+
                   '       end as amt,'+@CR+
                   '       ''Y'' as Not_Accumulate,'+@CR+
                   '       3 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '       inner join '+@TB_RT_Data_tmp+' d'+@TB_Hint+@CR+
                   '          on m.kind = d.kind and m.area = d.area'+@CR+
                   ' where 1=1'+@CR+
                   '   and m.area_year = d.area_year'+@CR+
                   '   and d.area_month <= m.area_month'+@CR+
                   '   and m.kind =''00083'''+@CR+
                   ' group by m.area_year, m.area_month, m.area, m.area_name '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end  
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00160'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月客戶總公司最後一次銷貨日期含服務單(以客戶六碼統計)'
     set @Remark= '(依銷貨日, 表頭, 含服務單, 排除 已關店)值量欄位請以自行轉換成日期型態'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_date_year as area_year,'+@CR+
                   '       chg_sp_date_month as area_month,'+@CR+
                   '       chg_bu_no as area,'+@CR+
                   '       ct_fld3 as area_name,'+@CR+
                   '       Convert(float, Max(sp_date)) as amt,'+@CR+
                   '       ''Y'' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslip'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and (sp_slip_fg =''2'' Or sp_slip_fg =''C'')'+@CR+
                   '   and chg_bu_no <> '''' '+@CR+
                   ' group by chg_sp_date_year, chg_sp_date_month, chg_bu_no, ct_fld3 '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end 

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00161'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月客戶總公司最後一次銷貨退回日期(以客戶六碼統計)'
     set @Remark=  '(依銷貨日, 表頭, 含服務單, 排除 已關店)值量欄位請以自行轉換成日期型態'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_date_year as area_year,'+@CR+
                   '       chg_sp_date_month as area_month,'+@CR+
                   '       chg_bu_no as area,'+@CR+
                   '       ct_fld3 as area_name,'+@CR+
                   '       Convert(float, Max(sp_date)) as amt,'+@CR+
                   '       ''Y'' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslip'+@TB_Hint+@CR+
                   ' where sp_slip_fg = ''3'''+@CR+
                   '   and chg_bu_no <> '''' '+@CR+
                   ' group by chg_sp_date_year, chg_sp_date_month, chg_bu_no, ct_fld3 '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end 

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00162'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '客戶總公司最後一次銷貨日期(以客戶六碼統計)'
     set @Remark=  '(依銷貨日, 表頭, 含服務單, 排除 已關店)值量欄位請以自行轉換成日期型態'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_date_year as area_year,'+@CR+
                   '       ''00'' as area_month,'+@CR+
                   '       chg_bu_no as area,'+@CR+
                   '       ct_fld3 as area_name,'+@CR+
                   '       Convert(float, Max(sp_date)) as amt,'+@CR+
                   '       ''Y'' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslip'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and (sp_slip_fg =''2'' Or sp_slip_fg =''C'')'+@CR+
                   '   and chg_bu_no <> '''' '+@CR+
                   ' group by chg_sp_date_year, chg_bu_no, ct_fld3 '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end 
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00163'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '客戶總公司最後一次銷貨退回日期(以客戶六碼統計)'
     set @Remark=  '(依銷貨日, 表頭, 含服務單, 排除 已關店)值量欄位請以自行轉換成日期型態'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_date_year as area_year,'+@CR+
                   '       ''00'' as area_month,'+@CR+
                   '       chg_bu_no as area,'+@CR+
                   '       ct_fld3 as area_name,'+@CR+
                   '       Convert(float, Max(sp_date)) as amt,'+@CR+
                   '       ''Y'' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslip'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and sp_slip_fg =''3'''+@CR+
                   '   and chg_bu_no <> '''' '+@CR+
                   ' group by chg_sp_date_year, chg_bu_no, ct_fld3 '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end   
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00164'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月百大總實銷金額佔每月總實銷金額佔比'
     set @Remark=  '每月百大總實銷金額佔每月總實銷金額佔比 Formula: 00067 / 00061'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       m.area_year,'+@CR+
                   '       m.area_month,'+@CR+
                   '       IsNull(m.area, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(m.area_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       case'+@CR+
                   '         when (IsNull(D.Amt, 0) <= 0) Or'+@CR+
                   '              (IsNull(M.Amt, 0) <= 0 And IsNull(D.amt, 0) <= 0) then 0'+@CR+
                   '         else Round(Isnull(M.Amt, 0) / D.Amt, 4)'+@CR+
                   '       end as Amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       3 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '       inner join '+@TB_RT_Data_tmp+' d'+@TB_Hint+@CR+
                   '          on m.area_year=d.area_year and m.area_month=d.area_month'+@CR+
                   ' where 1=1'+@CR+
                   '   and m.kind=''00067'''+@CR+
                   '   and d.kind =''00061'''+@CR

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00165'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月業務實銷金額佔比'
     set @Remark= '每月業務實銷金額佔比 Formula: Sum(00006) / 00061'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       m.area_year,'+@CR+
                   '       m.area_month,'+@CR+
                   '       IsNull(m.area, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(m.area_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Case'+@CR+
                   '         when IsNull(D.Tot_amt, 0) = 0 then 0'+@CR+
                   '         else IsNull(Round(Sum(Isnull(M.Amt, 0)) / D.Tot_amt, 4), 0)'+@CR+
                   '       end as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       3 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from (select area_year, area_month, area, area_name, IsNull(amt, 0) as amt'+@CR+
                   '          from '+@TB_RT_Data_tmp+@TB_Hint+@CR+
                   '         where 1=1'+@CR+
                   '           and kind =''00006'''+@CR+
                   '       ) M'+@CR+
                   '       inner join'+@CR+
                   '       (select area_year, area_month, IsNull(amt, 0) as Tot_amt'+@CR+
                   '          from '+@TB_RT_Data_tmp+@TB_Hint+@CR+
                   '         where 1=1'+@CR+
                   '           and kind =''00061'''+@CR+
                   '       ) D'+@CR+
                   '         on M.Area_year=D.Area_year'+@CR+
                   '        and M.Area_month=D.Area_month'+@CR+
                   ' group by m.Area_year, m.Area_month, m.Area, m.area_name, D.Tot_amt '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00166'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月部門實銷金額達成率'
     set @Remark= '每月部門實銷金額達成率 Formula: Sum(00004) / 00003'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       m.area_year,'+@CR+
                   '       m.area_month,'+@CR+
                   '       m.area,'+@CR+
                   '       m.area_name,'+@CR+
                   '       case'+@CR+
                   '         when IsNull(d.amt, 0) = 0 then 0'+@CR+
                   '         else Round(Sum(Isnull(m.amt, 0)) / d.amt, 4)'+@CR+
                   '       end as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       3 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '       inner join '+@TB_RT_Data_tmp+' d'+@TB_Hint+@CR+
                   '          on m.area_year=d.area_year'+@CR+
                   '         and m.area_month=d.area_month'+@CR+
                   '         and m.area=d.area'+@CR+
                   '         and m.kind = ''00004'''+@CR+
                   '         and d.kind =''00003'''+@CR+
                   '         and d.amt <> 0'+@CR+
                   ' group by m.area_year, m.area_month, m.area, m.area_name, d.amt '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00167'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月部門實銷金額佔比'
     set @Remark= '每月部門實銷金額佔比 Formula: Sum(00003) / 00004'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       m.area_year,'+@CR+
                   '       m.area_month,'+@CR+
                   '       m.area,'+@CR+
                   '       m.area_name,'+@CR+
                   '       case'+@CR+
                   '         when IsNull(d.tot_amt, 0) = 0 then 0'+@CR+
                   '         else Round(Sum(Isnull(m.amt, 0)) / d.tot_amt, 4)'+@CR+
                   '       end as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       3 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '       inner join'+@CR+
                   '       (select area_year,'+@CR+
                   '               area_month,'+@CR+
                   '               Sum(IsNull(amt, 0)) as tot_amt'+@CR+
                   '          from '+@TB_RT_Data_tmp+@TB_Hint+@CR+
                   '         where 1=1'+@CR+
                   '           and kind = ''00004'''+@CR+
                   '         group by area_year, area_month'+@CR+
                   '       ) d'+@CR+
                   '        on m.area_year=d.area_year'+@CR+
                   '       and m.area_month=d.area_month'+@CR+
                   '       and m.kind =''00003'''+@CR+
                   '       and d.tot_amt <> 0'+@CR+
                   ' group by m.area_year, m.area_month, m.area, m.area_name, d.tot_amt '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00168'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月部門實銷金額平均達成率'
     set @Remark= '每月部門實銷金額平均達成率 Formula: Sum(00166) / 00166 並以累計方式計算'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       m.area_year,'+@CR+
                   '       m.area_month,'+@CR+
                   '       m.area,'+@CR+
                   '       m.area_name,'+@CR+
                   '       case'+@CR+
                   '         when IsNull(m.area_month, 0) = 0 then 0'+@CR+
                   '         else Round(Sum(Isnull(d.amt, 0)) / m.area_month, 4)'+@CR+
                   '       end as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       3 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '       inner join '+@TB_RT_Data_tmp+' d'+@TB_Hint+@CR+
                   '          on m.kind = d.kind and m.area = d.area'+@CR+
                   ' where 1=1'+@CR+
                   '   and m.area_year = d.area_year'+@CR+
                   '   and d.area_month <= m.area_month'+@CR+
                   '   and m.kind =''00166'''+@CR+
                   ' group by m.area_year, m.area_month, m.area, m.area_name'+@CR+
                   ' order by 1, 2, 3 '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end 

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00170'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月系列商品實銷金額'
     set @Remark= '(依請款日, 表身, 含稅, 無折讓, 無加減項, 含退貨, 含服務單, 系列類別不含 Z 類別)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_pdate_year as area_year,'+@CR+
                   '       chg_sp_pdate_month as area_month,'+@CR+
                   '       IsNull(sk_kind, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(chg_kind_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(chg_sd_stot_tax, 0)) as amt,'+@CR+
                   '       ''Y'' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslpdt'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and (sd_slip_fg =''2'' Or sd_slip_fg =''C'' Or sd_slip_fg = ''3'')'+@CR+
                   '   and IsNull(chg_kind_name, '''') <> '''''+@CR+
                   '   and sk_kind <> ''Z'''+@CR+
                   '   and chg_sp_pdate_year >= ''2013'''+@CR+
                   ' group by chg_sp_pdate_year, chg_sp_pdate_month, sk_kind, chg_kind_name '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end  
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00171'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月系列商品銷售金額'
     set @Remark= '(依請款日, 表身, 含稅, 無折讓, 無加減項, 含服務單, 系列類別不含 Z 類別)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_pdate_year as area_year,'+@CR+
                   '       chg_sp_pdate_month as area_month,'+@CR+
                   '       IsNull(sk_kind, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(chg_kind_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(chg_sd_stot_tax, 0)) as amt,'+@CR+
                   '       ''Y'' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslpdt'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and (sd_slip_fg =''2'' Or sd_slip_fg =''C'')'+@CR+
                   '   and IsNull(chg_kind_name, '''') <> '''''+@CR+
                   '   and sk_kind <> ''Z'''+@CR+
                   '   and chg_sp_pdate_year >= ''2013'''+@CR+
                   ' group by chg_sp_pdate_year, chg_sp_pdate_month, sk_kind, chg_kind_name '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end  

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00172'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月系列商品退貨金額'
     set @Remark= '(依請款日, 表身, 含稅, 無折讓, 無加減項, 系列類別不含 Z 類別)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_pdate_year as area_year,'+@CR+
                   '       chg_sp_pdate_month as area_month,'+@CR+
                   '       IsNull(sk_kind, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(chg_kind_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(chg_sd_stot_tax, 0)) as amt,'+@CR+
                   '       ''Y'' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslpdt'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and sd_slip_fg =''3'''+@CR+
                   '   and IsNull(chg_kind_name, '''') <> '''''+@CR+
                   '   and sk_kind <> ''Z'''+@CR+
                   '   and chg_sp_pdate_year >= ''2013'''+@CR+
                   ' group by chg_sp_pdate_year, chg_sp_pdate_month, sk_kind, chg_kind_name '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end  

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00173'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月系列商品實銷金額佔比'
     set @Remark= '每月系列商品實銷金額佔比 Formula: Sum(00170) / 00061'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       m.area_year,'+@CR+
                   '       m.area_month,'+@CR+
                   '       area,'+@CR+
                   '       area_name,'+@CR+
                   '       case'+@CR+
                   '         when IsNull(d.Tot_amt, 0) = 0 then 0'+@CR+
                   '         else Round(Sum(Isnull(m.amt, 0)) / d.Tot_amt, 4)'+@CR+
                   '       end as amt,'+@CR+
                   '       ''Y'' as Not_Accumulate,'+@CR+
                   '       3 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '       inner join'+@CR+
                   '       (select area_year,'+@CR+
                   '               area_month,'+@CR+
                   '               Isnull(amt, 0) as tot_amt'+@CR+
                   '          from '+@TB_RT_Data_tmp+@TB_Hint+@CR+
                   '         where 1=1'+@CR+
                   '           and kind =''00061'''+@CR+
                   '       ) d'+@CR+
                   '        on m.area_year=d.area_year'+@CR+
                   '       and m.area_month=d.area_month'+@CR+
                   ' where 1=1'+@CR+
                   '   and kind=''00170'''+@CR+
                   ' group by m.area_year, m.area_month, area, area_name, tot_amt '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end  

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00174'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月系列商品銷售金額佔比'
     set @Remark= '每月系列商品銷售金額佔比 Formula: Sum(00171) / 00061'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       m.area_year,'+@CR+
                   '       m.area_month,'+@CR+
                   '       area,'+@CR+
                   '       area_name,'+@CR+
                   '       case'+@CR+
                   '         when IsNull(d.Tot_amt, 0) = 0 then 0'+@CR+
                   '         else Round(Sum(Isnull(m.amt, 0)) / d.Tot_amt, 4)'+@CR+
                   '       end as amt,'+@CR+
                   '       ''Y'' as Not_Accumulate,'+@CR+
                   '       3 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '       inner join'+@CR+
                   '       (select area_year,'+@CR+
                   '               area_month,'+@CR+
                   '               amt as tot_amt'+@CR+
                   '          from '+@TB_RT_Data_tmp+@TB_Hint+@CR+
                   '         where 1=1'+@CR+
                   '           and kind =''00061'''+@CR+
                   '       ) d'+@CR+
                   '        on m.area_year=d.area_year'+@CR+
                   '       and m.area_month=d.area_month'+@CR+
                   ' where 1=1'+@CR+
                   '   and kind=''00171'''+@CR+
                   ' group by m.area_year, m.area_month, area, area_name, tot_amt '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end  

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00175'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月系列商品退貨金額佔比'
     set @Remark= '每月系列商品退貨金額佔比 Formula: Sum(00172) / 00061'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       m.area_year,'+@CR+
                   '       m.area_month,'+@CR+
                   '       area,'+@CR+
                   '       area_name,'+@CR+
                   '       case'+@CR+
                   '         when IsNull(d.Tot_amt, 0) = 0 then 0'+@CR+
                   '         else Round(Sum(Isnull(m.amt, 0)) / d.Tot_amt, 4)'+@CR+
                   '       end as amt,'+@CR+
                   '       ''Y'' as Not_Accumulate,'+@CR+
                   '       3 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '       inner join'+@CR+
                   '       (select area_year,'+@CR+
                   '               area_month,'+@CR+
                   '               Isnull(amt, 0) as tot_amt'+@CR+
                   '          from '+@TB_RT_Data_tmp+@TB_Hint+@CR+
                   '         where 1=1'+@CR+
                   '           and kind =''00061'''+@CR+
                   '       ) d'+@CR+
                   '        on m.area_year=d.area_year'+@CR+
                   '       and m.area_month=d.area_month'+@CR+
                   ' where 1=1'+@CR+
                   '   and kind=''00172'''+@CR+
                   ' group by m.area_year, m.area_month, area, area_name, tot_amt '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end  
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00176'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月部門銷貨總金額(依客戶檔內的業務部門進行加總)'
     set @Remark= '(依請款日, 表頭, 含稅, 扣折讓, 無加減項, 含服務單)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_pdate_year,'+@CR+
                   '       chg_sp_pdate_month,'+@CR+
                   '       IsNull(e_dept, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(chg_dp_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(chg_sp_stot_tax, 0))+Sum(IsNull(Chg_sp_dis_tot2, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslip M'+@TB_Hint+@CR+ --不能直接 sum 部門方式，因為 sp_slip_fg 類別不同，所以會導致資料產出後變成多個部門情況
                   ' where 1=1'+@CR+
                   '   and chg_sp_pdate_year >= ''2013'''+@CR+
                   '   and (sp_slip_fg =''2'' Or sp_slip_fg =''C'')'+@CR+
                   '   and chg_sp_pdate_year is not null'+@CR+
                   -- 2017/08/08 Rickliu 副總指示客編 IT、ZZ 開頭不認列業績，雖不認列業績但應收扔得收
                   '   and Chg_BU_NO Not In (''IT0000'', ''ZZ0000'') '+@CR+
                   ' group by chg_sp_pdate_year, chg_sp_pdate_month, e_dept, chg_dp_name '
                   
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00177'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月部門退貨總金額(依客戶檔內的業務部門進行加總)'
     set @Remark= '(依請款日, 表頭, 含稅, 扣折讓, 無加減項)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_pdate_year,'+@CR+
                   '       chg_sp_pdate_month,'+@CR+
                   '       IsNull(e_dept, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(chg_dp_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(chg_sp_stot_tax, 0))+Sum(IsNull(Chg_sp_dis_tot2, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslip M'+@TB_Hint+@CR+ --不能直接 sum 部門方式，因為 sp_slip_fg 類別不同，所以會導致資料產出後變成多個部門情況
                   ' where 1=1'+@CR+
                   '   and chg_sp_pdate_year >= ''2013'''+@CR+
                   '   and sp_slip_fg = ''3'''+@CR+
                   '   and chg_sp_pdate_year is not null'+@CR+
                   -- 2017/08/08 Rickliu 副總指示客編 IT、ZZ 開頭不認列業績，雖不認列業績但應收扔得收
                   '   and Chg_BU_NO Not In (''IT0000'', ''ZZ0000'') '+@CR+
                   ' group by chg_sp_pdate_year, chg_sp_pdate_month, e_dept, chg_dp_name '
                   
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00300'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = 'TA13-AA 倉各商品最後庫存數'
     set @Remark= '僅抓取系統年度，也就是當下年度'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       year(getdate()) as area_year,'+@CR+
                   '       ''00'' as area_month,'+@CR+
                   '       wd_skno as area,'+@CR+
                   '       sk_name as area_name,'+@CR+
                   '       Sum(IsNull(wd_amt, 0)) as amt,'+@CR+
                   '       ''Y'' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from Fact_swaredt'+@TB_Hint+@CR+
                   '  where 1=1'+@CR+
                   '        and wd_yr <= year(getdate()) '+@CR+
                   '    and wd_no = ''AA'''+@CR+
                   '  group by wd_skno, sk_name'

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00301'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = 'TA13-AB 倉各商品最後庫存數'
     set @Remark= '僅抓取系統年度，也就是當下年度'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       year(getdate()) as area_year,'+@CR+
                   '       ''00'' as area_month,'+@CR+
                   '       wd_skno as area,'+@CR+
                   '       sk_name as area_name,'+@CR+
                   '       Sum(IsNull(wd_amt, 0)) as amt,'+@CR+
                   '       ''Y'' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from Fact_swaredt'+@TB_Hint+@CR+
                   '  where 1=1'+@CR+
                   '        and wd_yr <= year(getdate()) '+@CR+
                   '    and wd_no = ''AB'''+@CR+
                   '  group by wd_skno, sk_name'

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00302'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = 'TA13-AC 倉各商品最後庫存數'
     set @Remark= '僅抓取系統年度，也就是當下年度'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       year(getdate()) as area_year,'+@CR+
                   '       ''00'' as area_month,'+@CR+
                   '       wd_skno as area,'+@CR+
                   '       sk_name as area_name,'+@CR+
                   '       Sum(IsNull(wd_amt, 0)) as amt,'+@CR+
                   '       ''Y'' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from Fact_swaredt'+@TB_Hint+@CR+
                   '  where 1=1'+@CR+
                   '        and wd_yr <= year(getdate()) '+@CR+
                   '    and wd_no = ''AC'''+@CR+
                   '  group by wd_skno, sk_name'

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00303'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = 'TA13-AD 倉各商品最後庫存數'
     set @Remark= '僅抓取系統年度，也就是當下年度'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       year(getdate()) as area_year,'+@CR+
                   '       ''00'' as area_month,'+@CR+
                   '       wd_skno as area,'+@CR+
                   '       sk_name as area_name,'+@CR+
                   '       Sum(IsNull(wd_amt, 0)) as amt,'+@CR+
                   '       ''Y'' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from Fact_swaredt'+@TB_Hint+@CR+
                   '  where 1=1'+@CR+
                   '        and wd_yr <= year(getdate()) '+@CR+
                   '    and wd_no = ''AD'''+@CR+
                   '  group by wd_skno, sk_name'

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00304'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = 'TA13-AE 倉各商品最後庫存數'
     set @Remark= '僅抓取系統年度，也就是當下年度'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       year(getdate()) as area_year,'+@CR+
                   '       ''00'' as area_month,'+@CR+
                   '       wd_skno as area,'+@CR+
                   '       sk_name as area_name,'+@CR+
                   '       Sum(IsNull(wd_amt, 0)) as amt,'+@CR+
                   '       ''Y'' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from Fact_swaredt'+@TB_Hint+@CR+
                   '  where 1=1'+@CR+
                   '        and wd_yr <= year(getdate()) '+@CR+
                   '    and wd_no = ''AE'''+@CR+
                   '  group by wd_skno, sk_name'

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00305'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = 'TA13-AF 倉各商品最後庫存數'
     set @Remark= '僅抓取系統年度，也就是當下年度'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       year(getdate()) as area_year,'+@CR+
                   '       ''00'' as area_month,'+@CR+
                   '       wd_skno as area,'+@CR+
                   '       sk_name as area_name,'+@CR+
                   '       Sum(IsNull(wd_amt, 0)) as amt,'+@CR+
                   '       ''Y'' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from Fact_swaredt'+@TB_Hint+@CR+
                   '  where 1=1'+@CR+
                   '        and wd_yr <= year(getdate()) '+@CR+
                   '    and wd_no = ''AF'''+@CR+
                   '  group by wd_skno, sk_name'

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00306'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = 'TA13-AG 倉各商品最後庫存數'
     set @Remark= '僅抓取系統年度，也就是當下年度'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       year(getdate()) as area_year,'+@CR+
                   '       ''00'' as area_month,'+@CR+
                   '       wd_skno as area,'+@CR+
                   '       sk_name as area_name,'+@CR+
                   '       Sum(IsNull(wd_amt, 0)) as amt,'+@CR+
                   '       ''Y'' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from Fact_swaredt'+@TB_Hint+@CR+
                   '  where 1=1'+@CR+
                   '        and wd_yr <= year(getdate()) '+@CR+
                   '    and wd_no = ''AG'''+@CR+
                   '  group by wd_skno, sk_name'

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00307'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = 'TA13-AH 倉各商品最後庫存數'
     set @Remark= '僅抓取系統年度，也就是當下年度'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       year(getdate()) as area_year,'+@CR+
                   '       ''00'' as area_month,'+@CR+
                   '       wd_skno as area,'+@CR+
                   '       sk_name as area_name,'+@CR+
                   '       Sum(IsNull(wd_amt, 0)) as amt,'+@CR+
                   '       ''Y'' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from Fact_swaredt'+@TB_Hint+@CR+
                   '  where 1=1'+@CR+
                   '        and wd_yr <= year(getdate()) '+@CR+
                   '    and wd_no = ''AH'''+@CR+
                   '  group by wd_skno, sk_name'

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00308'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = 'TA13-託售剩餘存量(結算至系統年月往前推一個月)'
     set @Remark= '(依請款日, 含託售單, 含託回單)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       year(getdate()) as area_year,'+@CR+
                   '       ''00'' as area_month,'+@CR+
                   '       sd_skno as area,'+@CR+
                   '       sd_name as area_name,'+@CR+
                   '       Sum(Isnull(chg_sd_qty, 0)) as amt,'+@CR+
                   '       ''Y'' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslpdt'+@TB_Hint+@CR+
                   ' where sd_class =''3'''+@CR+
                   '   and Chg_sp_pdate_YM <='+@CR+
                   '       Convert(Varchar(7), DateAdd(mm, -1, getdate()), 111)'+@CR+
                   ' group by sd_skno, sd_name'+@CR+
                   'having Sum(chg_sd_qty) <> 0 '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00309'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = 'TA13-AZ 倉各商品最後庫存數'
     set @Remark= '僅抓取系統年度，也就是當下年度'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       year(getdate()) as area_year,'+@CR+
                   '       ''00'' as area_month,'+@CR+
                   '       wd_skno as area,'+@CR+
                   '       sk_name as area_name,'+@CR+
                   '       Sum(IsNull(wd_amt, 0)) as amt,'+@CR+
                   '       ''Y'' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from Fact_swaredt'+@TB_Hint+@CR+
                   '  where 1=1'+@CR+
                   '        and wd_yr <= year(getdate()) '+@CR+
                   '    and wd_no = ''AZ'''+@CR+
                   '  group by wd_skno, sk_name'

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00310'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = 'TA13-LA 倉各商品最後庫存數'
     set @Remark= '僅抓取系統年度，也就是當下年度'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       year(getdate()) as area_year,'+@CR+
                   '       ''00'' as area_month,'+@CR+
                   '       wd_skno as area,'+@CR+
                   '       sk_name as area_name,'+@CR+
                   '       Sum(IsNull(wd_amt, 0)) as amt,'+@CR+
                   '       ''Y'' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from Fact_swaredt'+@TB_Hint+@CR+
                   '  where 1=1'+@CR+
                   '        and wd_yr <= year(getdate()) '+@CR+
                   '    and wd_no = ''LA'''+@CR+
                   '  group by wd_skno, sk_name'

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00311'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = 'TA13-LB 倉各商品最後庫存數'
     set @Remark= '僅抓取系統年度，也就是當下年度'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       year(getdate()) as area_year,'+@CR+
                   '       ''00'' as area_month,'+@CR+
                   '       wd_skno as area,'+@CR+
                   '       sk_name as area_name,'+@CR+
                   '       Sum(IsNull(wd_amt, 0)) as amt,'+@CR+
                   '       ''Y'' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from Fact_swaredt'+@TB_Hint+@CR+
                   '  where 1=1'+@CR+
                   '        and wd_yr <= year(getdate()) '+@CR+
                   '    and wd_no = ''LB'''+@CR+
                   '  group by wd_skno, sk_name'

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00312'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = 'TA13-LC 倉各商品最後庫存數'
     set @Remark= '僅抓取系統年度，也就是當下年度'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       year(getdate()) as area_year,'+@CR+
                   '       ''00'' as area_month,'+@CR+
                   '       wd_skno as area,'+@CR+
                   '       sk_name as area_name,'+@CR+
                   '       Sum(IsNull(wd_amt, 0)) as amt,'+@CR+
                   '       ''Y'' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from Fact_swaredt'+@TB_Hint+@CR+
                   '  where 1=1'+@CR+
                   '        and wd_yr <= year(getdate()) '+@CR+
                   '    and wd_no = ''LC'''+@CR+
                   '  group by wd_skno, sk_name'

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00313'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = 'TA13-LD 倉各商品最後庫存數'
     set @Remark= '僅抓取系統年度，也就是當下年度'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       year(getdate()) as area_year,'+@CR+
                   '       ''00'' as area_month,'+@CR+
                   '       wd_skno as area,'+@CR+
                   '       sk_name as area_name,'+@CR+
                   '       Sum(IsNull(wd_amt, 0)) as amt,'+@CR+
                   '       ''Y'' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from Fact_swaredt'+@TB_Hint+@CR+
                   '  where 1=1'+@CR+
                   '        and wd_yr <= year(getdate()) '+@CR+
                   '    and wd_no = ''LD'''+@CR+
                   '  group by wd_skno, sk_name'

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00314'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = 'TA13-Z99 倉各商品最後庫存數'
     set @Remark= '僅抓取系統年度，也就是當下年度'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       year(getdate()) as area_year,'+@CR+
                   '       ''00'' as area_month,'+@CR+
                   '       wd_skno as area,'+@CR+
                   '       sk_name as area_name,'+@CR+
                   '       Sum(IsNull(wd_amt, 0)) as amt,'+@CR+
                   '       ''Y'' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from Fact_swaredt'+@TB_Hint+@CR+
                   '  where 1=1'+@CR+
                   '        and wd_yr <= year(getdate()) '+@CR+
                   '    and wd_no = ''Z99'''+@CR+
                   '  group by wd_skno, sk_name'

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00315'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = 'TA13-Z9999 倉各商品最後庫存數'
     set @Remark= '僅抓取系統年度，也就是當下年度'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       year(getdate()) as area_year,'+@CR+
                   '       ''00'' as area_month,'+@CR+
                   '       wd_skno as area,'+@CR+
                   '       sk_name as area_name,'+@CR+
                   '       Sum(IsNull(wd_amt, 0)) as amt,'+@CR+
                   '       ''Y'' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from Fact_swaredt'+@TB_Hint+@CR+
                   '  where 1=1'+@CR+
                   '        and wd_yr <= year(getdate()) '+@CR+
                   '    and wd_no = ''Z9999'''+@CR+
                   '  group by wd_skno, sk_name'

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00316'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = 'TA13-Z999999999 倉各商品最後庫存數'
     set @Remark= '僅抓取系統年度，也就是當下年度'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       year(getdate()) as area_year,'+@CR+
                   '       ''00'' as area_month,'+@CR+
                   '       wd_skno as area,'+@CR+
                   '       sk_name as area_name,'+@CR+
                   '       Sum(IsNull(wd_amt, 0)) as amt,'+@CR+
                   '       ''Y'' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from Fact_swaredt'+@TB_Hint+@CR+
                   '  where 1=1'+@CR+
                   '        and wd_yr <= year(getdate()) '+@CR+
                   '    and wd_no = ''Z999999999'''+@CR+
                   '  group by wd_skno, sk_name'

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00320'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '各商品最後庫存總數(TA13-所有倉庫)'
     set @Remark= '僅抓取系統年度，也就是當下年度, Formula: Sum(00300..00319)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       area_year,'+@CR+
                   '       area_month,'+@CR+
                   '       area,'+@CR+
                   '       area_name,'+@CR+
                   '       Sum(Isnull(amt, 0)) as amt,'+@CR+
                   '       ''Y'' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and kind >= ''00300'''+@CR+
                   '   and kind <= ''00319'''+@CR+
                   ' group by area_year, area_month, area, area_name '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00321'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '各商品最後庫存總數(TA13-僅限 A 開頭倉庫)'
     set @Remark= '(AA + AB + AC + AD + AE + AF + AG + AH + 託售庫存), Formula: Sum(00300..00308)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       area_year,'+@CR+
                   '       area_month,'+@CR+
                   '       area,'+@CR+
                   '       area_name,'+@CR+
                   '       Sum(Isnull(amt, 0)) as amt,'+@CR+
                   '       ''Y'' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and (kind =''00300'' Or kind =''00301'' Or kind =''00302'' Or kind =''00303'' Or kind =''00304'' Or'+@CR+
                   '        kind =''00305'' Or kind =''00306'' Or kind =''00307'' Or kind = ''00308'')'+@CR+
                   ' group by area_year, area_month, area, area_name '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00330'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = 'TA13-AA 倉各商品最後庫存成本總金額'
     set @Remark= '僅抓取系統年度，也就是當下年度(商品標準成本*庫存數量)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       year(getdate()) as area_year,'+@CR+
                   '       ''00'' as area_month,'+@CR+
                   '       wd_skno as area,'+@CR+
                   '       sk_name as area_name,'+@CR+
                   '       Sum(isnull(wd_save_tot, 0)) as amt,'+@CR+
                   '       ''Y'' as Not_Accumulate,'+@CR+
                   '       1 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_swaredt m'+@TB_Hint+@CR+
                   '  where 1=1'+@CR+
                   '    and wd_no = ''AA'''+@CR+
                   '    and wd_yr <= year(getdate()) '+@CR+
                   '  group by wd_skno, sk_name'

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00331'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = 'TA13-AB 倉各商品最後庫存成本總金額'
     set @Remark= '僅抓取系統年度，也就是當下年度(商品標準成本*庫存數量)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       year(getdate()) as area_year,'+@CR+
                   '       ''00'' as area_month,'+@CR+
                   '       wd_skno as area,'+@CR+
                   '       sk_name as area_name,'+@CR+
                   '       Sum(isnull(wd_save_tot, 0)) as amt,'+@CR+
                   '       ''Y'' as Not_Accumulate,'+@CR+
                   '       1 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_swaredt m'+@TB_Hint+@CR+
                   '  where 1=1'+@CR+
                   '    and wd_no = ''AB'''+@CR+
                   '    and wd_yr <= year(getdate()) '+@CR+
                   '  group by wd_skno, sk_name'
                   
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00332'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = 'TA13-AC 倉各商品最後庫存成本總金額'
     set @Remark= '僅抓取系統年度，也就是當下年度(商品標準成本*庫存數量)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       year(getdate()) as area_year,'+@CR+
                   '       ''00'' as area_month,'+@CR+
                   '       wd_skno as area,'+@CR+
                   '       sk_name as area_name,'+@CR+
                   '       Sum(isnull(wd_save_tot, 0)) as amt,'+@CR+
                   '       ''Y'' as Not_Accumulate,'+@CR+
                   '       1 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_swaredt m'+@TB_Hint+@CR+
                   '  where 1=1'+@CR+
                   '    and wd_no = ''AC'''+@CR+
                   '    and wd_yr <= year(getdate()) '+@CR+
                   '  group by wd_skno, sk_name'

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00333'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = 'TA13-AD 倉各商品最後庫存成本總金額'
     set @Remark= '僅抓取系統年度，也就是當下年度(商品標準成本*庫存數量)來'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       year(getdate()) as area_year,'+@CR+
                   '       ''00'' as area_month,'+@CR+
                   '       wd_skno as area,'+@CR+
                   '       sk_name as area_name,'+@CR+
                   '       Sum(isnull(wd_save_tot, 0)) as amt,'+@CR+
                   '       ''Y'' as Not_Accumulate,'+@CR+
                   '       1 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_swaredt m'+@TB_Hint+@CR+
                   '  where 1=1'+@CR+
                   '    and wd_no = ''AD'''+@CR+
                   '    and wd_yr <= year(getdate()) '+@CR+
                   '  group by wd_skno, sk_name'

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00334'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = 'TA13-AE 倉各商品最後庫存成本總金額'
     set @Remark= '僅抓取系統年度，也就是當下年度(商品標準成本*庫存數量)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       year(getdate()) as area_year,'+@CR+
                   '       ''00'' as area_month,'+@CR+
                   '       wd_skno as area,'+@CR+
                   '       sk_name as area_name,'+@CR+
                   '       Sum(isnull(wd_save_tot, 0)) as amt,'+@CR+
                   '       ''Y'' as Not_Accumulate,'+@CR+
                   '       1 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_swaredt m'+@TB_Hint+@CR+
                   '  where 1=1'+@CR+
                   '    and wd_no = ''AE'''+@CR+
                   '    and wd_yr <= year(getdate()) '+@CR+
                   '  group by wd_skno, sk_name'

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00335'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = 'TA13-AF 倉各商品最後庫存成本總金額'
     set @Remark= '僅抓取系統年度，也就是當下年度(商品標準成本*庫存數量)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       year(getdate()) as area_year,'+@CR+
                   '       ''00'' as area_month,'+@CR+
                   '       wd_skno as area,'+@CR+
                   '       sk_name as area_name,'+@CR+
                   '       Sum(isnull(wd_save_tot, 0)) as amt,'+@CR+
                   '       ''Y'' as Not_Accumulate,'+@CR+
                   '       1 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_swaredt m'+@TB_Hint+@CR+
                   '  where 1=1'+@CR+
                   '    and wd_no = ''AF'''+@CR+
                   '    and wd_yr <= year(getdate()) '+@CR+
                   '  group by wd_skno, sk_name'

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00336'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = 'TA13-AG 倉各商品最後庫存成本總金額'
     set @Remark= '僅抓取系統年度，也就是當下年度(商品標準成本*庫存數量)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       year(getdate()) as area_year,'+@CR+
                   '       ''00'' as area_month,'+@CR+
                   '       wd_skno as area,'+@CR+
                   '       sk_name as area_name,'+@CR+
                   '       Sum(isnull(wd_save_tot, 0)) as amt,'+@CR+
                   '       ''Y'' as Not_Accumulate,'+@CR+
                   '       1 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_swaredt m'+@TB_Hint+@CR+
                   '  where 1=1'+@CR+
                   '    and wd_no = ''AG'''+@CR+
                   '    and wd_yr <= year(getdate()) '+@CR+
                   '  group by wd_skno, sk_name'

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00337'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = 'TA13-AH 倉各商品最後庫存成本總金額'
     set @Remark= '僅抓取系統年度，也就是當下年度(商品標準成本*庫存數量)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       year(getdate()) as area_year,'+@CR+
                   '       ''00'' as area_month,'+@CR+
                   '       wd_skno as area,'+@CR+
                   '       sk_name as area_name,'+@CR+
                   '       Sum(isnull(wd_save_tot, 0)) as amt,'+@CR+
                   '       ''Y'' as Not_Accumulate,'+@CR+
                   '       1 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_swaredt m'+@TB_Hint+@CR+
                   '  where 1=1'+@CR+
                   '    and wd_no = ''AH'''+@CR+
                   '    and wd_yr <= year(getdate()) '+@CR+
                   '  group by wd_skno, sk_name'

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00338'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = 'TA13-託售回貨剩餘成本總金額(結算至上一個系統月份)'
     set @Remark= '僅抓取系統年度，也就是當下年度(商品標準成本*庫存數量)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       year(getdate()) as area_year,'+@CR+
                   '       ''00'' as area_month,'+@CR+
                   '       sd_skno as area,'+@CR+
                   '       sd_name as area_name,'+@CR+
                   '       Sum(Isnull(chg_sd_qty * sk_save, 0)) as amt,'+@CR+
                   '       ''Y'' as Not_Accumulate,'+@CR+
                   '       1 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslpdt'+@TB_Hint+@CR+
                   ' where sd_class =''3'''+@CR+
                   '   and Chg_sp_pdate_YM <='+@CR+
                   '       Convert(Varchar(7), DateAdd(mm, -1, getdate()), 111)'+@CR+
                   ' group by sd_skno, sd_name'+@CR+
                   'having Sum(Isnull(chg_sd_qty * sk_save, 0)) <> 0 '
                   
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00339'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = 'TA13-AZ 倉各商品最後庫存成本總金額'
     set @Remark= '僅抓取系統年度，也就是當下年度(商品標準成本*庫存數量)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       year(getdate()) as area_year,'+@CR+
                   '       ''00'' as area_month,'+@CR+
                   '       wd_skno as area,'+@CR+
                   '       sk_name as area_name,'+@CR+
                   '       Sum(isnull(wd_save_tot, 0)) as amt,'+@CR+
                   '       ''Y'' as Not_Accumulate,'+@CR+
                   '       1 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_swaredt m'+@TB_Hint+@CR+
                   '  where 1=1'+@CR+
                   '    and wd_no = ''AZ'''+@CR+
                   '    and wd_yr <= year(getdate()) '+@CR+
                   '  group by wd_skno, sk_name'

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00340'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = 'TA13-LA 倉各商品最後庫存成本總金額'
     set @Remark= '僅抓取系統年度，也就是當下年度(商品標準成本*庫存數量)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       year(getdate()) as area_year,'+@CR+
                   '       ''00'' as area_month,'+@CR+
                   '       wd_skno as area,'+@CR+
                   '       sk_name as area_name,'+@CR+
                   '       Sum(isnull(wd_save_tot, 0)) as amt,'+@CR+
                   '       ''Y'' as Not_Accumulate,'+@CR+
                   '       1 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_swaredt m'+@TB_Hint+@CR+
                   '  where 1=1'+@CR+
                   '    and wd_no = ''LA'''+@CR+
                   '    and wd_yr <= year(getdate()) '+@CR+
                   '  group by wd_skno, sk_name'

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end  
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00341'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = 'TA13-LB 倉各商品最後庫存成本總金額'
     set @Remark= '僅抓取系統年度，也就是當下年度(商品標準成本*庫存數量)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       year(getdate()) as area_year,'+@CR+
                   '       ''00'' as area_month,'+@CR+
                   '       wd_skno as area,'+@CR+
                   '       sk_name as area_name,'+@CR+
                   '       Sum(isnull(wd_save_tot, 0)) as amt,'+@CR+
                   '       ''Y'' as Not_Accumulate,'+@CR+
                   '       1 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_swaredt m'+@TB_Hint+@CR+
                   '  where 1=1'+@CR+
                   '    and wd_no = ''LB'''+@CR+
                   '    and wd_yr <= year(getdate()) '+@CR+
                   '  group by wd_skno, sk_name'

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end  
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00342'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = 'TA13-LC 倉各商品最後庫存成本總金額'
     set @Remark= '僅抓取系統年度，也就是當下年度(商品標準成本*庫存數量)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       year(getdate()) as area_year,'+@CR+
                   '       ''00'' as area_month,'+@CR+
                   '       wd_skno as area,'+@CR+
                   '       sk_name as area_name,'+@CR+
                   '       Sum(isnull(wd_save_tot, 0)) as amt,'+@CR+
                   '       ''Y'' as Not_Accumulate,'+@CR+
                   '       1 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_swaredt m'+@TB_Hint+@CR+
                   '  where 1=1'+@CR+
                   '    and wd_no = ''LC'''+@CR+
                   '    and wd_yr <= year(getdate()) '+@CR+
                   '  group by wd_skno, sk_name'

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end  

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00343'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = 'TA13-LD 倉各商品最後庫存成本總金額'
     set @Remark= '僅抓取系統年度，也就是當下年度(商品標準成本*庫存數量)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       year(getdate()) as area_year,'+@CR+
                   '       ''00'' as area_month,'+@CR+
                   '       wd_skno as area,'+@CR+
                   '       sk_name as area_name,'+@CR+
                   '       Sum(isnull(wd_save_tot, 0)) as amt,'+@CR+
                   '       ''Y'' as Not_Accumulate,'+@CR+
                   '       1 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_swaredt m'+@TB_Hint+@CR+
                   '  where 1=1'+@CR+
                   '    and wd_no = ''LD'''+@CR+
                   '    and wd_yr <= year(getdate()) '+@CR+
                   '  group by wd_skno, sk_name'


     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end  

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00344'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = 'TA13-Z99 倉各商品最後庫存成本總金額'
     set @Remark= '僅抓取系統年度，也就是當下年度(商品標準成本*庫存數量)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       year(getdate()) as area_year,'+@CR+
                   '       ''00'' as area_month,'+@CR+
                   '       wd_skno as area,'+@CR+
                   '       sk_name as area_name,'+@CR+
                   '       Sum(isnull(wd_save_tot, 0)) as amt,'+@CR+
                   '       ''Y'' as Not_Accumulate,'+@CR+
                   '       1 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_swaredt m'+@TB_Hint+@CR+
                   '  where 1=1'+@CR+
                   '    and wd_no = ''Z99'''+@CR+
                   '    and wd_yr <= year(getdate()) '+@CR+
                   '  group by wd_skno, sk_name'

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end  

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00345'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = 'TA13-Z9999 倉各商品最後庫存成本總金額'
     set @Remark= '僅抓取系統年度，也就是當下年度(商品標準成本*庫存數量)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       year(getdate()) as area_year,'+@CR+
                   '       ''00'' as area_month,'+@CR+
                   '       wd_skno as area,'+@CR+
                   '       sk_name as area_name,'+@CR+
                   '       Sum(isnull(wd_save_tot, 0)) as amt,'+@CR+
                   '       ''Y'' as Not_Accumulate,'+@CR+
                   '       1 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_swaredt m'+@TB_Hint+@CR+
                   '  where 1=1'+@CR+
                   '    and wd_no = ''Z9999'''+@CR+
                   '    and wd_yr <= year(getdate()) '+@CR+
                   '  group by wd_skno, sk_name'

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end  

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00346'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = 'TA13-Z999999999 倉各商品最後庫存成本總金額'
     set @Remark= '僅抓取系統年度，也就是當下年度(商品標準成本*庫存數量)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       year(getdate()) as area_year,'+@CR+
                   '       ''00'' as area_month,'+@CR+
                   '       wd_skno as area,'+@CR+
                   '       sk_name as area_name,'+@CR+
                   '       Sum(isnull(wd_save_tot, 0)) as amt,'+@CR+
                   '       ''Y'' as Not_Accumulate,'+@CR+
                   '       1 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_swaredt m'+@TB_Hint+@CR+
                   '  where 1=1'+@CR+
                   '    and wd_no = ''Z999999999'''+@CR+
                   '    and wd_yr <= year(getdate()) '+@CR+
                   '  group by wd_skno, sk_name'
                   
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00350'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '各商品最後庫存成本總金額(TA13-所有倉庫)'
     set @Remark= '僅抓取系統年度，也就是當下年度, Formula: Sum(00330..00345)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       area_year,'+@CR+
                   '       area_month,'+@CR+
                   '       area,'+@CR+
                   '       area_name,'+@CR+
                   '       Sum(Isnull(amt, 0)) as amt,'+@CR+
                   '       ''Y'' as Not_Accumulate,'+@CR+
                   '       1 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and kind >= ''00330'''+@CR+
                   '   and kind <= ''00345'''+@CR+
                   ' group by area_year, area_month, area, area_name'

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00351'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '各商品最後庫存成本總金額(TA13-僅限 A 開頭倉庫)'
     set @Remark= '(AA + AB + AC + AD + AE + AF + AG + AH + 託售庫存), Formula: Sum(00330..00338)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       area_year,'+@CR+
                   '       area_month,'+@CR+
                   '       area,'+@CR+
                   '       area_name,'+@CR+
                   '       Sum(Isnull(amt, 0)) as amt,'+@CR+
                   '       ''Y'' as Not_Accumulate,'+@CR+
                   '       1 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and (kind=''00330'' Or Kind=''00331'' Or Kind=''00332'' Or Kind=''00333'' Or Kind=''00334'' Or'+@CR+
                   '        Kind=''00335'' Or Kind=''00336'' Or Kind=''00337'' Or Kind=''00338'')'+@CR+
                   ' group by area_year, area_month, area, area_name'

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00352'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '庫存成本總金額(TA13-所有倉庫)'
     set @Remark= '僅抓取系統年度，也就是當下年度, 由 0350 計算而來'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       year(getdate()) as area_year,'+@CR+
                   '       ''00'' as area_month,'+@CR+
                   '       '''+@NonSet+''' as area,'+@CR+
                   '       '''+@CNonSet+''' as area_name,'+@CR+
                   '       Sum(Isnull(amt, 0)) as amt,'+@CR+
                   '       ''Y'' as Not_Accumulate,'+@CR+
                   '       1 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and m.kind = ''00350'' '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00353'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '庫存成本總金額(TA13-僅限 A 開頭倉庫)'
     set @Remark= '僅抓取系統年度，也就是當下年度, 由 0351 計算而來'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       year(getdate()) as area_year,'+@CR+
                   '       ''00'' as area_month,'+@CR+
                   '       '''+@NonSet+''' as area,'+@CR+
                   '       '''+@CNonSet+''' as area_name,'+@CR+
                   '       Sum(Isnull(amt, 0)) as amt,'+@CR+
                   '       ''Y'' as Not_Accumulate,'+@CR+
                   '       1 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and m.kind = ''00351'' '
                   
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00360'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '滯銷各商品最後庫存總數(TA13-所有倉庫)'
     set @Remark= '僅抓取系統年度，也就是當下年度, 由 0320 計算而來'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       year(getdate()) as area_year,'+@CR+
                   '       ''00'' as area_month,'+@CR+
                   '       area,'+@CR+
                   '       area_name,'+@CR+
                   '       Isnull(amt, 0) as amt,'+@CR+
                   '       ''Y'' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '       inner join Fact_sstock d'+@TB_Hint+@CR+
                   '          on m.area '+@Collate+' ='+@CR+
                   '             d.sk_no '+@Collate+@CR+
                   '         and d.chg_is_dead_stock = ''Y'''+@CR+
                   ' where 1=1'+@CR+
                   '   and m.kind = ''00320'' '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00361'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '滯銷各商品最後庫存總數(TA13-僅限 A 開頭倉庫)'
     set @Remark= '僅抓取系統年度，也就是當下年度, 由 0321 計算而來'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       year(getdate()) as area_year,'+@CR+
                   '       ''00'' as area_month,'+@CR+
                   '       area,'+@CR+
                   '       area_name,'+@CR+
                   '       Isnull(amt, 0) as amt,'+@CR+
                   '       ''Y'' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '       inner join Fact_sstock d'+@TB_Hint+@CR+
                   '          on m.area '+@Collate+' ='+@CR+
                   '             d.sk_no '+@Collate+@CR+
                   '         and d.chg_is_dead_stock = ''Y'''+@CR+
                   ' where 1=1'+@CR+
                   '   and m.kind = ''00321'' '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00362'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '滯銷各商品最後庫存成本總金額(TA13-所有倉庫)'
     set @Remark= '僅抓取系統年度，也就是當下年度, 由 0350 計算而來'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       year(getdate()) as area_year,'+@CR+
                   '       ''00'' as area_month,'+@CR+
                   '       area,'+@CR+
                   '       area_name,'+@CR+
                   '       Isnull(amt, 0) as amt,'+@CR+
                   '       ''Y'' as Not_Accumulate,'+@CR+
                   '       1 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '       inner join Fact_sstock d'+@TB_Hint+@CR+
                   '          on m.area '+@Collate+' ='+@CR+
                   '             d.sk_no '+@Collate+@CR+
                   '         and d.chg_is_dead_stock = ''Y'''+@CR+
                   ' where 1=1'+@CR+
                   '   and m.kind = ''00350'' '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00363'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '滯銷各商品最後庫存成本總金額(TA13-僅限 A 開頭倉庫)'
     set @Remark= '僅抓取系統年度，也就是當下年度, 由 0351 計算而來'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       year(getdate()) as area_year,'+@CR+
                   '       ''00'' as area_month,'+@CR+
                   '       area,'+@CR+
                   '       area_name,'+@CR+
                   '       Isnull(amt, 0) as amt,'+@CR+
                   '       ''Y'' as Not_Accumulate,'+@CR+
                   '       1 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '       inner join Fact_sstock d'+@TB_Hint+@CR+
                   '          on m.area '+@Collate+' ='+@CR+
                   '             d.sk_no '+@Collate+@CR+
                   '         and d.chg_is_dead_stock = ''Y'''+@CR+
                   ' where 1=1'+@CR+
                   '   and m.kind = ''00351'' '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00364'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '滯銷品最後庫存成本總金額(TA13-所有倉庫)'
     set @Remark= '僅抓取系統年度，也就是當下年度, 由 0362 計算而來'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       year(getdate()) as area_year,'+@CR+
                   '       ''00'' as area_month,'+@CR+
                   '       '''+@NonSet+''' as area,'+@CR+
                   '       '''+@CNonSet+''' as area_name,'+@CR+
                   '       Sum(Isnull(amt, 0)) as amt,'+@CR+
                   '       ''Y'' as Not_Accumulate,'+@CR+
                   '       1 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and m.kind = ''00362'' '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00365'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '滯銷品最後庫存成本總金額(TA13-僅限 A 開頭倉庫)'
     set @Remark= '僅抓取系統年度，由 0363 計算而來'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       year(getdate()) as area_year,'+@CR+
                   '       ''00'' as area_month,'+@CR+
                   '       '''+@NonSet+''' as area,'+@CR+
                   '       '''+@CNonSet+''' as area_name,'+@CR+
                   '       Sum(Isnull(amt, 0)) as amt,'+@CR+
                   '       ''Y'' as Not_Accumulate,'+@CR+
                   '       1 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and m.kind = ''00363'' '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00366'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '滯銷品成本總金額佔庫存成本總金額比例(TA13-所有倉庫)'
     set @Remark= '僅抓取系統年度，也就是當下年度 Formula: 00364 / 00352'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       m.area_year,'+@CR+
                   '       m.area_month,'+@CR+
                   '       m.area,'+@CR+
                   '       m.area_name,'+@CR+
                   '       case'+@CR+
                   '         when Isnull(d.amt, 0) = 0 then 0'+@CR+
                   '         else Round(Isnull(m.amt, 0) / d.amt, 4)'+@CR+
                   '       end as amt,'+@CR+
                   '       ''Y'' as Not_Accumulate,'+@CR+
                   '       3 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '       inner join '+@TB_RT_Data_tmp+' d'+@TB_Hint+@CR+
                   '          on m.area_year = d.area_year'+@CR+
                   '         and m.area_month = d.area_month'+@CR+
                   '         and d.kind = ''00352'''+@CR+
                   ' where 1=1'+@CR+
                   '   and m.kind = ''00364'' '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00367'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '滯銷品成本總金額佔庫存成本總金額比例(TA13-僅限 A 開頭倉庫)'
     set @Remark= '(AA + AB + AC + AD + AE + AF + AG + AH + 託售庫存) Formula: 00365 / 00353'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       m.area_year,'+@CR+
                   '       m.area_month,'+@CR+
                   '       m.area,'+@CR+
                   '       m.area_name,'+@CR+
                   '       case'+@CR+
                   '         when Isnull(d.amt, 0) = 0 then 0'+@CR+
                   '         else Round(Isnull(m.amt, 0) / d.amt, 4)'+@CR+
                   '       end as amt,'+@CR+
                   '       ''Y'' as Not_Accumulate,'+@CR+
                   '       3 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '       inner join '+@TB_RT_Data_tmp+' d'+@TB_Hint+@CR+
                   '          on m.area_year = d.area_year'+@CR+
                   '         and m.area_month = d.area_month'+@CR+
                   '         and d.kind = ''00353'''+@CR+
                   ' where 1=1'+@CR+
                   '   and m.kind = ''00365'' '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00371'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '新品各商品最後庫存總數(TA13-所有倉庫)'
     set @Remark= '僅抓取系統年度，也就是當下年度, 由 0320 計算而來'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       year(getdate()) as area_year,'+@CR+
                   '       ''00'' as area_month,'+@CR+
                   '       area,'+@CR+
                   '       area_name,'+@CR+
                   '       Isnull(amt, 0) as amt,'+@CR+
                   '       ''Y'' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '       inner join Fact_sstock d'+@TB_Hint+@CR+
                   '          on m.area '+@Collate+' ='+@CR+
                   '             d.sk_no '+@Collate+@CR+
                   '         and IsNull(d.Chg_New_Arrival_YM, '''') <> '''''+@CR+
                   ' where 1=1'+@CR+
                   '   and m.kind = ''00320'' '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00372'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '新品各商品最後庫存總數(TA13-僅限 A 開頭倉庫)'
     set @Remark= '(AA + AB + AC + AD + AE + AF + AG + AH + 託售庫存), 由 0321 計算而來'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       year(getdate()) as area_year,'+@CR+
                   '       ''00'' as area_month,'+@CR+
                   '       area,'+@CR+
                   '       area_name,'+@CR+
                   '       Isnull(amt, 0) as amt,'+@CR+
                   '       ''Y'' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '       inner join Fact_sstock d'+@TB_Hint+@CR+
                   '          on m.area '+@Collate+' ='+@CR+
                   '             d.sk_no '+@Collate+@CR+
                   '         and IsNull(d.Chg_New_Arrival_YM, '''') <> '''''+@CR+
                   ' where 1=1'+@CR+
                   '   and m.kind = ''00321'' '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00373'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '新品各商品最後庫存成本總金額(TA13-所有倉庫)'
     set @Remark= '僅抓取系統年度，也就是當下年度, 由 0350 計算而來'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       year(getdate()) as area_year,'+@CR+
                   '       ''00'' as area_month,'+@CR+
                   '       area,'+@CR+
                   '       area_name,'+@CR+
                   '       Isnull(amt, 0) as amt,'+@CR+
                   '       ''Y'' as Not_Accumulate,'+@CR+
                   '       1 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '       inner join Fact_sstock d'+@TB_Hint+@CR+
                   '          on m.area '+@Collate+' ='+@CR+
                   '             d.sk_no '+@Collate+@CR+
                   '         and IsNull(d.Chg_New_Arrival_YM, '''') <> '''''+@CR+
                   ' where 1=1'+@CR+
                   '   and m.kind = ''00350'' '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00374'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '新品各商品最後庫存成本總金額(TA13-僅限 A 開頭倉庫)'
     set @Remark= '(AA + AB + AC + AD + AE + AF + AG + AH + 託售庫存), 由 0351 計算而來'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       year(getdate()) as area_year,'+@CR+
                   '       ''00'' as area_month,'+@CR+
                   '       area,'+@CR+
                   '       area_name,'+@CR+
                   '       Isnull(amt, 0) as amt,'+@CR+
                   '       ''Y'' as Not_Accumulate,'+@CR+
                   '       1 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '       inner join Fact_sstock d'+@TB_Hint+@CR+
                   '          on m.area '+@Collate+' ='+@CR+
                   '             d.sk_no '+@Collate+@CR+
                   '         and IsNull(d.Chg_New_Arrival_YM, '''') <> '''''+@CR+
                   ' where 1=1'+@CR+
                   '   and m.kind = ''00351'' '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00375'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '新品最後庫存成本總金額(TA13-所有倉庫)'
     set @Remark= '僅抓取系統年度，也就是當下年度, 由 0373 計算而來'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       year(getdate()) as area_year,'+@CR+
                   '       ''00'' as area_month,'+@CR+
                   '       '''+@NonSet+''' as area,'+@CR+
                   '       '''+@CNonSet+''' as area_name,'+@CR+
                   '       Sum(Isnull(amt, 0)) as amt,'+@CR+
                   '       ''Y'' as Not_Accumulate,'+@CR+
                   '       1 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and m.kind = ''00373'' '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00376'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '新品最後庫存成本總金額(TA13-僅限 A 開頭倉庫)'
     set @Remark= '(AA + AB + AC + AD + AE + AF + AG + AH + 託售庫存), 由 0374 計算而來'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       year(getdate()) as area_year,'+@CR+
                   '       ''00'' as area_month,'+@CR+
                   '       '''+@NonSet+''' as area,'+@CR+
                   '       '''+@CNonSet+''' as area_name,'+@CR+
                   '       Sum(Isnull(amt, 0)) as amt,'+@CR+
                   '       ''Y'' as Not_Accumulate,'+@CR+
                   '       1 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and m.kind = ''00374'' '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00377'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '新品最後庫存成本總金額佔庫存總金額比例(TA13-所有倉庫)'
     set @Remark= '新品最後庫存成本總金額佔庫存總金額比例, Formula: 00375 / 00352'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       m.area_year,'+@CR+
                   '       m.area_month,'+@CR+
                   '       m.area,'+@CR+
                   '       m.area_name,'+@CR+
                   '       case'+@CR+
                   '         when Isnull(d.amt, 0) = 0 then 0'+@CR+
                   '         else Round(Isnull(m.amt, 0) / d.amt, 4)'+@CR+
                   '       end as amt,'+@CR+
                   '       ''Y'' as Not_Accumulate,'+@CR+
                   '       3 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '       inner join '+@TB_RT_Data_tmp+' d'+@TB_Hint+@CR+
                   '          on m.area_year = d.area_year'+@CR+
                   '         and m.area_month = d.area_month'+@CR+
                   '         and d.kind = ''00352'''+@CR+
                   ' where 1=1'+@CR+
                   '   and m.kind = ''00375'' '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00378'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '新品最後庫存成本總金額佔庫存成本總金額比例(TA13-僅限 A 開頭倉庫)'
     set @Remark= '(AA + AB + AC + AD + AE + AF + AG + AH + 託售庫存), Formula: 00376 / 00353'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       m.area_year,'+@CR+
                   '       m.area_month,'+@CR+
                   '       m.area,'+@CR+
                   '       m.area_name,'+@CR+
                   '       case'+@CR+
                   '         when Isnull(d.amt, 0) = 0 then 0'+@CR+
                   '         else Round(Isnull(m.amt, 0) / d.amt, 4)'+@CR+
                   '       end as amt,'+@CR+
                   '       ''Y'' as Not_Accumulate,'+@CR+
                   '       3 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '       inner join '+@TB_RT_Data_tmp+' d'+@TB_Hint+@CR+
                   '          on m.area_year = d.area_year'+@CR+
                   '         and m.area_month = d.area_month'+@CR+
                   '         and d.kind = ''00353'''+@CR+
                   ' where 1=1'+@CR+
                   '   and m.kind = ''00376'' '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00380'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '各大分類最後庫存成本總金額(TA13-所有倉庫)'
     set @Remark= '僅抓取系統年度，也就是當下年度, Formula: Sum(00350)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       m.area_year,'+@CR+
                   '       m.area_month,'+@CR+
                   '       d.Chg_skno_BKind,'+@CR+
                   '       d.Chg_skno_Bkind_Name,'+@CR+
                   '       Sum(isnull(amt, 0)) as amt,'+@CR+
                   '       ''Y'' as Not_Accumulate,'+@CR+
                   '       1 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '       left join fact_sstock d '+@TB_Hint+@CR+
                   '         on m.area = d.sk_no Collate Chinese_Taiwan_Stroke_CI_AS '+@CR+
                   ' where 1=1'+@CR+
                   '   and m.kind = ''00350'' '+@CR+
                   '   and d.Chg_skno_Bkind is not null '+@CR+
                   ' group by m.area_year, m.area_month, d.Chg_skno_Bkind, d.Chg_skno_Bkind_Name'

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00381'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '各大分類最後庫存成本總金額(TA13-僅限 A 開頭倉庫)'
     set @Remark= '(AA + AB + AC + AD + AE + AF + AG + AH + 託售庫存), Formula: Sum(00351)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       m.area_year,'+@CR+
                   '       m.area_month,'+@CR+
                   '       d.Chg_skno_BKind,'+@CR+
                   '       d.Chg_skno_Bkind_Name,'+@CR+
                   '       Sum(isnull(amt, 0)) as amt,'+@CR+
                   '       ''Y'' as Not_Accumulate,'+@CR+
                   '       1 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '       left join fact_sstock d '+@TB_Hint+@CR+
                   '         on m.area = d.sk_no Collate Chinese_Taiwan_Stroke_CI_AS '+@CR+
                   ' where 1=1'+@CR+
                   '   and m.kind = ''00351'' '+@CR+
                   '   and d.Chg_skno_Bkind is not null '+@CR+
                   ' group by m.area_year, m.area_month, d.Chg_skno_Bkind, d.Chg_skno_Bkind_Name'

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00382'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = 'TA13-AA 倉各大分類最後庫存成本總金額'
     set @Remark= '僅抓取系統年度，也就是當下年度(商品標準成本*庫存數量), Formula: Sum(00330)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       m.area_year,'+@CR+
                   '       m.area_month,'+@CR+
                   '       d.Chg_skno_BKind,'+@CR+
                   '       d.Chg_skno_Bkind_Name,'+@CR+
                   '       Sum(isnull(amt, 0)) as amt,'+@CR+
                   '       ''Y'' as Not_Accumulate,'+@CR+
                   '       1 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '       left join fact_sstock d '+@TB_Hint+@CR+
                   '         on m.area = d.sk_no Collate Chinese_Taiwan_Stroke_CI_AS '+@CR+
                   ' where 1=1'+@CR+
                   '   and m.kind = ''00330'' '+@CR+
                   '   and d.Chg_skno_Bkind is not null '+@CR+
                   ' group by m.area_year, m.area_month, d.Chg_skno_Bkind, d.Chg_skno_Bkind_Name'

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00383'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = 'TA13-AB 倉各大分類最後庫存成本總金額'
     set @Remark= '僅抓取系統年度，也就是當下年度(商品標準成本*庫存數量), Formula: Sum(00331)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       m.area_year,'+@CR+
                   '       m.area_month,'+@CR+
                   '       d.Chg_skno_BKind,'+@CR+
                   '       d.Chg_skno_Bkind_Name,'+@CR+
                   '       Sum(isnull(amt, 0)) as amt,'+@CR+
                   '       ''Y'' as Not_Accumulate,'+@CR+
                   '       1 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '       left join fact_sstock d '+@TB_Hint+@CR+
                   '         on m.area = d.sk_no Collate Chinese_Taiwan_Stroke_CI_AS '+@CR+
                   ' where 1=1'+@CR+
                   '   and m.kind = ''00331'' '+@CR+
                   '   and d.Chg_skno_Bkind is not null '+@CR+
                   ' group by m.area_year, m.area_month, d.Chg_skno_Bkind, d.Chg_skno_Bkind_Name'

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00384'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = 'TA13-AC 倉各大分類最後庫存成本總金額'
     set @Remark= '僅抓取系統年度，也就是當下年度(商品標準成本*庫存數量), Formula: Sum(00332)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       m.area_year,'+@CR+
                   '       m.area_month,'+@CR+
                   '       d.Chg_skno_BKind,'+@CR+
                   '       d.Chg_skno_Bkind_Name,'+@CR+
                   '       Sum(isnull(amt, 0)) as amt,'+@CR+
                   '       ''Y'' as Not_Accumulate,'+@CR+
                   '       1 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '       left join fact_sstock d '+@TB_Hint+@CR+
                   '         on m.area = d.sk_no Collate Chinese_Taiwan_Stroke_CI_AS '+@CR+
                   ' where 1=1'+@CR+
                   '   and m.kind = ''00332'' '+@CR+
                   '   and d.Chg_skno_Bkind is not null '+@CR+
                   ' group by m.area_year, m.area_month, d.Chg_skno_Bkind, d.Chg_skno_Bkind_Name'

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00384'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = 'TA13-AD 倉各大分類最後庫存成本總金額'
     set @Remark= '僅抓取系統年度，也就是當下年度(商品標準成本*庫存數量), Formula: Sum(00333)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       m.area_year,'+@CR+
                   '       m.area_month,'+@CR+
                   '       d.Chg_skno_BKind,'+@CR+
                   '       d.Chg_skno_Bkind_Name,'+@CR+
                   '       Sum(isnull(amt, 0)) as amt,'+@CR+
                   '       ''Y'' as Not_Accumulate,'+@CR+
                   '       1 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '       left join fact_sstock d '+@TB_Hint+@CR+
                   '         on m.area = d.sk_no Collate Chinese_Taiwan_Stroke_CI_AS '+@CR+
                   ' where 1=1'+@CR+
                   '   and m.kind = ''00333'' '+@CR+
                   '   and d.Chg_skno_Bkind is not null '+@CR+
                   ' group by m.area_year, m.area_month, d.Chg_skno_Bkind, d.Chg_skno_Bkind_Name'

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00385'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = 'TA13-AE 倉各大分類最後庫存成本總金額'
     set @Remark= '僅抓取系統年度，也就是當下年度(商品標準成本*庫存數量), Formula: Sum(00334)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       m.area_year,'+@CR+
                   '       m.area_month,'+@CR+
                   '       d.Chg_skno_BKind,'+@CR+
                   '       d.Chg_skno_Bkind_Name,'+@CR+
                   '       Sum(isnull(amt, 0)) as amt,'+@CR+
                   '       ''Y'' as Not_Accumulate,'+@CR+
                   '       1 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '       left join fact_sstock d '+@TB_Hint+@CR+
                   '         on m.area = d.sk_no Collate Chinese_Taiwan_Stroke_CI_AS '+@CR+
                   ' where 1=1'+@CR+
                   '   and m.kind = ''00334'' '+@CR+
                   '   and d.Chg_skno_Bkind is not null '+@CR+
                   ' group by m.area_year, m.area_month, d.Chg_skno_Bkind, d.Chg_skno_Bkind_Name'

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00386'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = 'TA13-AF 倉各大分類最後庫存成本總金額'
     set @Remark= '僅抓取系統年度，也就是當下年度(商品標準成本*庫存數量), Formula: Sum(00335)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       m.area_year,'+@CR+
                   '       m.area_month,'+@CR+
                   '       d.Chg_skno_BKind,'+@CR+
                   '       d.Chg_skno_Bkind_Name,'+@CR+
                   '       Sum(isnull(amt, 0)) as amt,'+@CR+
                   '       ''Y'' as Not_Accumulate,'+@CR+
                   '       1 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '       left join fact_sstock d '+@TB_Hint+@CR+
                   '         on m.area = d.sk_no Collate Chinese_Taiwan_Stroke_CI_AS '+@CR+
                   ' where 1=1'+@CR+
                   '   and m.kind = ''00335'' '+@CR+
                   '   and d.Chg_skno_Bkind is not null '+@CR+
                   ' group by m.area_year, m.area_month, d.Chg_skno_Bkind, d.Chg_skno_Bkind_Name'

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00387'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = 'TA13-AG 倉各大分類最後庫存成本總金額'
     set @Remark= '僅抓取系統年度，也就是當下年度(商品標準成本*庫存數量), Formula: Sum(00336)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       m.area_year,'+@CR+
                   '       m.area_month,'+@CR+
                   '       d.Chg_skno_BKind,'+@CR+
                   '       d.Chg_skno_Bkind_Name,'+@CR+
                   '       Sum(isnull(amt, 0)) as amt,'+@CR+
                   '       ''Y'' as Not_Accumulate,'+@CR+
                   '       1 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '       left join fact_sstock d '+@TB_Hint+@CR+
                   '         on m.area = d.sk_no Collate Chinese_Taiwan_Stroke_CI_AS '+@CR+
                   ' where 1=1'+@CR+
                   '   and m.kind = ''00336'' '+@CR+
                   '   and d.Chg_skno_Bkind is not null '+@CR+
                   ' group by m.area_year, m.area_month, d.Chg_skno_Bkind, d.Chg_skno_Bkind_Name'

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00388'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = 'TA13-AH 倉各大分類最後庫存成本總金額'
     set @Remark= '僅抓取系統年度，也就是當下年度(商品標準成本*庫存數量), Formula: Sum(00337)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       m.area_year,'+@CR+
                   '       m.area_month,'+@CR+
                   '       d.Chg_skno_BKind,'+@CR+
                   '       d.Chg_skno_Bkind_Name,'+@CR+
                   '       Sum(isnull(amt, 0)) as amt,'+@CR+
                   '       ''Y'' as Not_Accumulate,'+@CR+
                   '       1 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '       left join fact_sstock d '+@TB_Hint+@CR+
                   '         on m.area = d.sk_no Collate Chinese_Taiwan_Stroke_CI_AS '+@CR+
                   ' where 1=1'+@CR+
                   '   and m.kind = ''00337'' '+@CR+
                   '   and d.Chg_skno_Bkind is not null '+@CR+
                   ' group by m.area_year, m.area_month, d.Chg_skno_Bkind, d.Chg_skno_Bkind_Name'

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00389'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = 'TA13-託售回貨各大分類剩餘成本總金額(結算至上一個系統月份)'
     set @Remark= '僅抓取系統年度，也就是當下年度(商品標準成本*庫存數量), Formula: Sum(00338)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       m.area_year,'+@CR+
                   '       m.area_month,'+@CR+
                   '       d.Chg_skno_BKind,'+@CR+
                   '       d.Chg_skno_Bkind_Name,'+@CR+
                   '       Sum(isnull(amt, 0)) as amt,'+@CR+
                   '       ''Y'' as Not_Accumulate,'+@CR+
                   '       1 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '       left join fact_sstock d '+@TB_Hint+@CR+
                   '         on m.area = d.sk_no Collate Chinese_Taiwan_Stroke_CI_AS '+@CR+
                   ' where 1=1'+@CR+
                   '   and m.kind = ''00338'' '+@CR+
                   '   and d.Chg_skno_Bkind is not null '+@CR+
                   ' group by m.area_year, m.area_month, d.Chg_skno_Bkind, d.Chg_skno_Bkind_Name'

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00390'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = 'TA13-AZ 倉各大分類最後庫存成本總金額'
     set @Remark= '僅抓取系統年度，也就是當下年度(商品標準成本*庫存數量), Formula: Sum(00339)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       m.area_year,'+@CR+
                   '       m.area_month,'+@CR+
                   '       d.Chg_skno_BKind,'+@CR+
                   '       d.Chg_skno_Bkind_Name,'+@CR+
                   '       Sum(isnull(amt, 0)) as amt,'+@CR+
                   '       ''Y'' as Not_Accumulate,'+@CR+
                   '       1 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '       left join fact_sstock d '+@TB_Hint+@CR+
                   '         on m.area = d.sk_no Collate Chinese_Taiwan_Stroke_CI_AS '+@CR+
                   ' where 1=1'+@CR+
                   '   and m.kind = ''00339'' '+@CR+
                   '   and d.Chg_skno_Bkind is not null '+@CR+
                   ' group by m.area_year, m.area_month, d.Chg_skno_Bkind, d.Chg_skno_Bkind_Name'

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00391'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = 'TA13-LA 倉各大分類最後庫存成本總金額'
     set @Remark= '僅抓取系統年度，也就是當下年度(商品標準成本*庫存數量), Formula: Sum(00340)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       m.area_year,'+@CR+
                   '       m.area_month,'+@CR+
                   '       d.Chg_skno_BKind,'+@CR+
                   '       d.Chg_skno_Bkind_Name,'+@CR+
                   '       Sum(isnull(amt, 0)) as amt,'+@CR+
                   '       ''Y'' as Not_Accumulate,'+@CR+
                   '       1 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '       left join fact_sstock d '+@TB_Hint+@CR+
                   '         on m.area = d.sk_no Collate Chinese_Taiwan_Stroke_CI_AS '+@CR+
                   ' where 1=1'+@CR+
                   '   and m.kind = ''00340'' '+@CR+
                   '   and d.Chg_skno_Bkind is not null '+@CR+
                   ' group by m.area_year, m.area_month, d.Chg_skno_Bkind, d.Chg_skno_Bkind_Name'

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00392'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = 'TA13-LB 倉各大分類最後庫存成本總金額'
     set @Remark= '僅抓取系統年度，也就是當下年度(商品標準成本*庫存數量), Formula: Sum(00341)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       m.area_year,'+@CR+
                   '       m.area_month,'+@CR+
                   '       d.Chg_skno_BKind,'+@CR+
                   '       d.Chg_skno_Bkind_Name,'+@CR+
                   '       Sum(isnull(amt, 0)) as amt,'+@CR+
                   '       ''Y'' as Not_Accumulate,'+@CR+
                   '       1 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '       left join fact_sstock d '+@TB_Hint+@CR+
                   '         on m.area = d.sk_no Collate Chinese_Taiwan_Stroke_CI_AS '+@CR+
                   ' where 1=1'+@CR+
                   '   and m.kind = ''00341'' '+@CR+
                   '   and d.Chg_skno_Bkind is not null '+@CR+
                   ' group by m.area_year, m.area_month, d.Chg_skno_Bkind, d.Chg_skno_Bkind_Name'

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00393'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = 'TA13-LC 倉各大分類最後庫存成本總金額'
     set @Remark= '僅抓取系統年度，也就是當下年度(商品標準成本*庫存數量), Formula: Sum(00342)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       m.area_year,'+@CR+
                   '       m.area_month,'+@CR+
                   '       d.Chg_skno_BKind,'+@CR+
                   '       d.Chg_skno_Bkind_Name,'+@CR+
                   '       Sum(isnull(amt, 0)) as amt,'+@CR+
                   '       ''Y'' as Not_Accumulate,'+@CR+
                   '       1 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '       left join fact_sstock d '+@TB_Hint+@CR+
                   '         on m.area = d.sk_no Collate Chinese_Taiwan_Stroke_CI_AS '+@CR+
                   ' where 1=1'+@CR+
                   '   and m.kind = ''00342'' '+@CR+
                   '   and d.Chg_skno_Bkind is not null '+@CR+
                   ' group by m.area_year, m.area_month, d.Chg_skno_Bkind, d.Chg_skno_Bkind_Name'

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00394'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = 'TA13-LD 倉各大分類最後庫存成本總金額'
     set @Remark= '僅抓取系統年度，也就是當下年度(商品標準成本*庫存數量), Formula: Sum(00343)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       m.area_year,'+@CR+
                   '       m.area_month,'+@CR+
                   '       d.Chg_skno_BKind,'+@CR+
                   '       d.Chg_skno_Bkind_Name,'+@CR+
                   '       Sum(isnull(amt, 0)) as amt,'+@CR+
                   '       ''Y'' as Not_Accumulate,'+@CR+
                   '       1 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '       left join fact_sstock d '+@TB_Hint+@CR+
                   '         on m.area = d.sk_no Collate Chinese_Taiwan_Stroke_CI_AS '+@CR+
                   ' where 1=1'+@CR+
                   '   and m.kind = ''00343'' '+@CR+
                   '   and d.Chg_skno_Bkind is not null '+@CR+
                   ' group by m.area_year, m.area_month, d.Chg_skno_Bkind, d.Chg_skno_Bkind_Name'

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00395'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = 'TA13-Z99 倉各大分類最後庫存成本總金額'
     set @Remark= '僅抓取系統年度，也就是當下年度(商品標準成本*庫存數量), Formula: Sum(00344)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       m.area_year,'+@CR+
                   '       m.area_month,'+@CR+
                   '       d.Chg_skno_BKind,'+@CR+
                   '       d.Chg_skno_Bkind_Name,'+@CR+
                   '       Sum(isnull(amt, 0)) as amt,'+@CR+
                   '       ''Y'' as Not_Accumulate,'+@CR+
                   '       1 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '       left join fact_sstock d '+@TB_Hint+@CR+
                   '         on m.area = d.sk_no Collate Chinese_Taiwan_Stroke_CI_AS '+@CR+
                   ' where 1=1'+@CR+
                   '   and m.kind = ''00344'' '+@CR+
                   '   and d.Chg_skno_Bkind is not null '+@CR+
                   ' group by m.area_year, m.area_month, d.Chg_skno_Bkind, d.Chg_skno_Bkind_Name'

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00396'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = 'TA13-Z9999 倉各大分類最後庫存成本總金額'
     set @Remark= '僅抓取系統年度，也就是當下年度(商品標準成本*庫存數量), Formula: Sum(00345)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       m.area_year,'+@CR+
                   '       m.area_month,'+@CR+
                   '       d.Chg_skno_BKind,'+@CR+
                   '       d.Chg_skno_Bkind_Name,'+@CR+
                   '       Sum(isnull(amt, 0)) as amt,'+@CR+
                   '       ''Y'' as Not_Accumulate,'+@CR+
                   '       1 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '       left join fact_sstock d '+@TB_Hint+@CR+
                   '         on m.area = d.sk_no Collate Chinese_Taiwan_Stroke_CI_AS '+@CR+
                   ' where 1=1'+@CR+
                   '   and m.kind = ''00345'' '+@CR+
                   '   and d.Chg_skno_Bkind is not null '+@CR+
                   ' group by m.area_year, m.area_month, d.Chg_skno_Bkind, d.Chg_skno_Bkind_Name'

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00397'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = 'TA13-Z999999999 倉各大分類最後庫存成本總金額'
     set @Remark= '僅抓取系統年度，也就是當下年度(商品標準成本*庫存數量), Formula: Sum(00346)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       m.area_year,'+@CR+
                   '       m.area_month,'+@CR+
                   '       d.Chg_skno_BKind,'+@CR+
                   '       d.Chg_skno_Bkind_Name,'+@CR+
                   '       Sum(isnull(amt, 0)) as amt,'+@CR+
                   '       ''Y'' as Not_Accumulate,'+@CR+
                   '       1 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '       left join fact_sstock d '+@TB_Hint+@CR+
                   '         on m.area = d.sk_no Collate Chinese_Taiwan_Stroke_CI_AS '+@CR+
                   ' where 1=1'+@CR+
                   '   and m.kind = ''00346'' '+@CR+
                   '   and d.Chg_skno_Bkind is not null '+@CR+
                   ' group by m.area_year, m.area_month, d.Chg_skno_Bkind, d.Chg_skno_Bkind_Name'

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00410'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = 'TA13-各(客戶八碼)寄倉標準成本金額(結算至系統年月往前推一個月)'
     set @Remark= '(依請款日, 含託售單, 含託回單)(商品標準成本*庫存數量)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       year(getdate()) as area_year,'+@CR+
                   '       ''00'' as area_month,'+@CR+
                   '       ct_no8 as area,'+@CR+
                   '       ct_sname8 as area_name,'+@CR+
                   '       Sum(Isnull(chg_sd_qty * sk_save, 0)) as amt,'+@CR+
                   '       ''Y'' as Not_Accumulate,'+@CR+
                   '       1 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslpdt'+@TB_Hint+@CR+
                   ' where sd_class =''3'''+@CR+
                   '   and Chg_sp_pdate_YM <= Convert(Varchar(7), DateAdd(mm, -1, getdate()), 111)'+@CR+
                   -- 2017/08/08 Rickliu 副總指示客編 IT、ZZ 開頭不認列業績
                   '   and Chg_BU_NO Not In (''IT0000'', ''ZZ0000'') '+@CR+
                   '  group by ct_no8, ct_sname8 '+@CR+
                   ' having Sum(Isnull(chg_sd_qty * sk_save, 0)) <> 0 '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00411'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = 'TA13-各客戶總公司寄倉標準成本金額(結算至系統年月往前推一個月)'
     set @Remark= '(依請款日, 含託售單, 含託回單)(商品標準成本*庫存數量)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       year(getdate()) as area_year,'+@CR+
                   '       ''00'' as area_month,'+@CR+
                   '       chg_bu_no as area,'+@CR+
                   '       ct_fld3 as area_name,'+@CR+
                   '       Sum(Isnull(chg_sd_qty * sk_save, 0)) as amt,'+@CR+
                   '       ''Y'' as Not_Accumulate,'+@CR+
                   '       1 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslpdt'+@TB_Hint+@CR+
                   ' where sd_class =''3'''+@CR+
                   '   and Chg_sp_pdate_YM <='+@CR+
                   '       Convert(Varchar(7), DateAdd(mm, -1, getdate()), 111)'+@CR+
                   '  group by chg_bu_no, ct_fld3 '+@CR+
                   ' having Sum(Isnull(chg_sd_qty * sk_save, 0)) <> 0 '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00440'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月總公司單據通路類別實銷金額'
     set @Remark= '(依銷貨日, 表頭, 含稅, 扣除加減項, 扣除折讓, 含退貨, 含服務單), Formula: 銷售金額(含稅) - 加減項金額'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_date_year as area_year,'+@CR+
                   '       chg_sp_date_month as area_month,'+@CR+
                   '       IsNull(Chg_BU_NO+''-''+Chg_Cust_Sale_Class, '''+@CNonSet+''') as area,'+@CR+
                   '       Chg_Cust_Sale_Class_sName as area_name,'+@CR+
                   '       Isnull(Round(Sum(IsNull(chg_sp_stot_tax, 0)), 4) +'+@CR+
                   '       sum(Isnull(Chg_sp_dis_tot2, 0)) + sum(Chg_SP_PAMT+Chg_SP_MAMT) * -1, 0) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       1 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslip'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and (sp_slip_fg=''2'' Or sp_slip_fg=''3'' Or sp_slip_fg=''C'')'+@CR+
                   '   and Chg_BU_NO <> '''' '+@CR+
                   ' group by chg_sp_date_year, chg_sp_date_month, IsNull(Chg_BU_NO+''-''+Chg_Cust_Sale_Class, '''+@CNonSet+'''), Chg_Cust_Sale_Class_sName '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00441'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月總公司單據通路類別實銷成本'
     set @Remark= '(依請款日, 表頭, 未稅, 含加減項, 含折讓, 含退貨, 含服務單) Formula: 單據成本'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_date_year as area_year,'+@CR+
                   '       chg_sp_date_month as area_month,'+@CR+
                   '       IsNull(Chg_BU_NO+''-''+Chg_Cust_Sale_Class, '''+@CNonSet+''') as area,'+@CR+
                   '       Chg_Cust_Sale_Class_sName as area_name,'+@CR+
                   '       sum(Isnull(chg_sp_ave_p, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       1 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslip'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and (sp_slip_fg=''2'' Or sp_slip_fg=''3'' Or sp_slip_fg=''C'')'+@CR+
                   '   and Chg_BU_NO <> '''' '+@CR+
                   ' group by chg_sp_date_year, chg_sp_date_month, IsNull(Chg_BU_NO+''-''+Chg_Cust_Sale_Class, '''+@CNonSet+'''), Chg_Cust_Sale_Class_sName '
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00442'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月分公司單據通路類別實銷利潤'
     set @Remark= '(依請款日, 表頭, 未稅) Formula: 00440 - 00441'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       m.area_year,'+@CR+
                   '       m.area_month,'+@CR+
                   '       m.area,'+@CR+
                   '       m.area_name,'+@CR+
                   '       Round(Isnull(m.amt, 0) - isnull(d.amt, 0), 4) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       1 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '       inner join '+@TB_RT_Data_tmp+' d'+@TB_Hint+@CR+
                   '          on m.area_year = d.area_year'+@CR+
                   '         and m.area_month = d.area_month'+@CR+
                   '         and m.area = d.area'+@CR+
                   ' where 1=1'+@CR+
                   '   and m.kind = ''00440'''+@CR+
                   '   and d.kind = ''00441'' '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00443'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月分公司單據通路類別實銷毛利率'
     set @Remark= '每月總公司實銷毛利率 Formula: 00442 / 00440'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       m.area_year,'+@CR+
                   '       m.area_month,'+@CR+
                   '       m.area,'+@CR+
                   '       m.area_name,'+@CR+
                   '       case'+@CR+
                   '         when IsNull(d.amt, 0) = 0 then 0'+@CR+
                   '         else Round(Isnull(m.amt, 0) / isnull(d.amt, 0), 4)'+@CR+
                   '       end as amt,'+@CR+
                   '       ''Y'' as Not_Accumulate,'+@CR+
                   '       3 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '       inner join '+@TB_RT_Data_tmp+' d'+@TB_Hint+@CR+
                   '          on m.area_year = d.area_year'+@CR+
                   '         and m.area_month = d.area_month'+@CR+
                   '         and m.area = d.area'+@CR+
                   ' where 1=1'+@CR+
                   '   and m.kind = ''00442'''+@CR+
                   '   and d.kind = ''00440'' '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00450'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月部門單據通路類別實銷金額'
     set @Remark= '(依請款日, 表頭, 含稅, 扣除加減項, 扣除折讓, 含退貨, 含服務單), Formula: 實銷金額(含稅) - 加減項金額'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select distinct '+@CR+
                   '       '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_date_year as area_year,'+@CR+
                   '       chg_sp_date_month area_month,'+@CR+
                   '       Chg_Dept_Cust_Chain_No as area,'+@CR+
                   '       Chg_Dept_Cust_Chain_Name as area_name,'+@CR+
                   '       Round(Sum(IsNull(chg_sp_stot_tax, 0)), 4) +'+@CR+
                   '       sum(Isnull(Chg_sp_dis_tot2, 0)) + sum(Isnull(Chg_SP_PAMT,0)+Isnull(Chg_SP_MAMT,0)) * -1 as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       1 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslip '+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and (sp_slip_fg =''2'' Or sp_slip_fg =''C'' Or sp_slip_fg = ''3'')'+@CR+
                   -- 2017/08/08 Rickliu 副總指示客編 IT、ZZ 開頭不認列業績，雖不認列業績但應收扔得收
                   '   and Chg_BU_NO Not In (''IT0000'', ''ZZ0000'') '+@CR+
                   ' group by chg_sp_date_year, chg_sp_date_month, Chg_Dept_Cust_Chain_No, Chg_Dept_Cust_Chain_Name '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00451'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月部門單據通路類別實銷成本'
     set @Remark= '(依銷貨日, 表頭, 未稅, 含加減項, 含折讓, 含退貨, 含服務單) Formula: 單據成本'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select distinct '+@CR+
                   '       '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_date_year as area_year,'+@CR+
                   '       chg_sp_date_month area_month,'+@CR+
                   '       Chg_Dept_Cust_Chain_No as area,'+@CR+
                   '       Chg_Dept_Cust_Chain_Name as area_name,'+@CR+
                   '       Round(Sum(IsNull(chg_sp_ave_p, 0)), 4) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       1 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslip '+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and (sp_slip_fg =''2'' Or sp_slip_fg =''C'' Or sp_slip_fg = ''3'')'+@CR+
                   -- 2017/08/08 Rickliu 副總指示客編 IT、ZZ 開頭不認列業績，雖不認列業績但應收扔得收
                   '   and Chg_BU_NO Not In (''IT0000'', ''ZZ0000'') '+@CR+
                   ' group by chg_sp_date_year, chg_sp_date_month, Chg_Dept_Cust_Chain_No, Chg_Dept_Cust_Chain_Name '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00452'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月部門單據通路類別實銷利潤'
     set @Remark= '(依請款日, 表頭, 未稅) Formula: 00450 - 00451'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       m.area_year,'+@CR+
                   '       m.area_month,'+@CR+
                   '       m.area,'+@CR+
                   '       m.area_name,'+@CR+
                   '       Round(Isnull(m.amt, 0) - isnull(d.amt, 0), 4) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       1 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '       inner join '+@TB_RT_Data_tmp+' d'+@TB_Hint+@CR+
                   '          on m.area_year = d.area_year'+@CR+
                   '         and m.area_month = d.area_month'+@CR+
                   '         and m.area = d.area'+@CR+
                   ' where 1=1'+@CR+
                   '   and m.kind = ''00450'''+@CR+
                   '   and d.kind = ''00451'' '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00453'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月部門單據通路類別實銷毛利率'
     set @Remark= '每月部門實銷毛利率 Formula: 00452 / 00450'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       m.area_year,'+@CR+
                   '       m.area_month,'+@CR+
                   '       m.area,'+@CR+
                   '       m.area_name,'+@CR+
                   '       case'+@CR+
                   '         when IsNull(d.amt, 0) = 0 then 0'+@CR+
                   '         else Round(Isnull(m.amt, 0) / isnull(d.amt, 0), 4)'+@CR+
                   '       end as amt,'+@CR+
                   '       ''Y'' as Not_Accumulate,'+@CR+
                   '       3 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '       inner join '+@TB_RT_Data_tmp+' d'+@TB_Hint+@CR+
                   '          on m.area_year = d.area_year'+@CR+
                   '         and m.area_month = d.area_month'+@CR+
                   '         and m.area = d.area'+@CR+
                   ' where 1=1'+@CR+
                   '   and m.kind = ''00452'''+@CR+
                   '   and d.kind = ''00450'' '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00500'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月客戶分店-業務實銷金額(客戶業績)'
     set @Remark= '(請款日, 表頭, 含稅, 無加減項, 無折讓, 含退貨, 含服務單)客戶編號開頭+業務員分群'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_pdate_year,'+@CR+
                   '       chg_sp_pdate_month,'+@CR+
                   '       IsNull(d.sp_ctno, ''000000000'')+''-''+IsNull(d.sp_sales, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(d.ct_sname, '''+@CNonSet+''')+''-''+IsNull(chg_sales_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Round(Sum(IsNull(chg_sp_stot_tax, 0)), 4) + Round(Sum(IsNull(Chg_sp_dis_tot2, 0)), 4) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from Fact_sslip as d'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and (sp_slip_fg =''2'' Or sp_slip_fg =''C'' Or sp_slip_fg = ''3'')'+@CR+
                   '   and chg_sp_pdate_year is not null'+@CR+
                   -- 2017/08/08 Rickliu 副總指示客編 IT、ZZ 開頭不認列業績，雖不認列業績但應收扔得收
                   '   and Chg_BU_NO Not In (''IT0000'', ''ZZ0000'') '+@CR+
                   ' group by chg_sp_pdate_year, chg_sp_pdate_month, IsNull(sp_ctno, ''000000000'')+''-''+IsNull(sp_sales, '''+@NonSet+'''), IsNull(ct_sname, '''+@CNonSet+''')+''-''+IsNull(chg_sales_name, '''+@CNonSet+''')'+@CR+
                   'having Sum(chg_sp_stot_tax) <> 0  '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end 

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00501'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月客戶分店-業務實銷金額達成率(客戶業績)'
     set @Remark= '每月客戶分店-業務實銷金額達成率 Formula: 00500 / 00126'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       m.area_year,'+@CR+
                   '       m.area_month,'+@CR+
                   '       d.area,'+@CR+
                   '       d.area_name,'+@CR+
                   '       case'+@CR+
                   '         when IsNull(m.amt, 0) = 0 then 0'+@CR+
                   '         else Round(IsNull(d.amt, 0) / m.amt, 4)'+@CR+
                   '       end as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       3 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '       inner join '+@TB_RT_Data_tmp+' d'+@TB_Hint+@CR+
                   '          on m.area_year = d.area_year'+@CR+
                   '         and m.area_month = d.area_month'+@CR+
                   '         and m.area = Substring(d.area, 1, 9)'+@CR+
                   '         and d.kind =''00500'''+@CR+
                   ' where m.kind =''00126'' '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end 
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00502'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月客戶分店-業務實銷金額達成率 For 網通'
     set @Remark= '每月客戶分店-業務實銷金額達成率 For 網通 Formula: 00500 / 00118'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       m.area_year,'+@CR+
                   '       m.area_month,'+@CR+
                   '       d.area,'+@CR+
                   '       d.area_name,'+@CR+
                   '       case'+@CR+
                   '         when IsNull(m.amt, 0) = 0 then 0'+@CR+
                   '         else Round(IsNull(d.amt, 0) / m.amt, 4)'+@CR+
                   '       end as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       3 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '       inner join '+@TB_RT_Data_tmp+' d'+@TB_Hint+@CR+
                   '          on m.area_year = d.area_year'+@CR+
                   '         and m.area_month = d.area_month'+@CR+
                   '         and m.area = Substring(d.area, 1, 9)'+@CR+
                   '         and d.kind =''00500'''+@CR+
                   ' where m.kind =''00118'' '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end 

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00800'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月銷貨成本'
     set @Remark= '由 00052 * 商品成本而來'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_pdate_year as area_year,'+@CR+
                   '       chg_sp_pdate_month as area_month,'+@CR+
                   '       '''+@NonSet+''' as area,'+@CR+
                   '       '''+@CNonSet+''' as area_name,'+@CR+
                   '       Sum(IsNull(chg_sd_qty * sd_ave_p, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       1 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslpdt'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and (sd_slip_fg =''2'' Or sd_slip_fg =''C'')'+@CR+
                   '   and chg_sp_pdate_year is not null'+@CR+
                   ' group by chg_sp_pdate_year, chg_sp_pdate_month '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00801'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每年銷貨成本'
     set @Remark= '由 00800 加總而來'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       m.area_year,'+@CR+
                   '       ''00'' as area_month,'+@CR+ 
                   '       '''+@NonSet+''' as area,'+@CR+ 
                   '       '''+@CNonSet+''' as area_name,'+@CR+ 
                   '       Round(Sum(Isnull(m.amt, 0)), 4) as amt,'+@CR+
                   '       ''Y'' as Not_Accumulate,'+@CR+
                   '       1 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   ' where m.kind = ''00800'''+@CR+
                   ' group by m.area_year '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00802'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每年平均成本'
     set @Remark= '每年平均成本 Formula: 00801 / 00052'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       m.area_year,'+@CR+
                   '       ''00'' as area_month,'+@CR+ 
                   '       '''+@NonSet+''' as area,'+@CR+ 
                   '       '''+@CNonSet+''' as area_name,'+@CR+ 
                   '       case'+@CR+
                   '         when Isnull(d.area_month, 0) = 0 then 0'+@CR+
                   '         else Round(IsNull(m.amt, 0) / d.area_month, 4)'+@CR+
                   '       end as amt,'+@CR+
                   '       ''Y'' as Not_Accumulate,'+@CR+
                   '       1 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '       inner join'+@CR+
                   '       (select area_year, Max(area_month) as area_month'+@CR+
                   '          from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '         where m.kind = ''00052'''+@CR+
                   '         group by area_year'+@CR+
                   '       ) d'+@CR+
                   '        on m.area_year = d.area_year'+@CR+
                   ' where m.kind = ''00801'' '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00803'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '存貨周轉率(銷貨成本/庫存成本)'
     set @Remark= '測度企業存貨週轉速度、銷貨能力高低及存貨量適當與否，就企業資金運用觀點而言，本比率愈高愈好, Formula: 00801 / 00353'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       m.area_year,'+@CR+
                   '       ''00'' as area_month,'+@CR+ 
                   '       '''+@NonSet+''' as area,'+@CR+ 
                   '       '''+@CNonSet+''' as area_name,'+@CR+ 
                   '       case'+@CR+
                   '         when Isnull(d.amt, 0) = 0 then 0'+@CR+
                   '         else Round(IsNull(m.amt, 0) / d.amt, 4)'+@CR+
                   '       end as amt,'+@CR+
                   '       ''Y'' as Not_Accumulate,'+@CR+
                   '       3 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '       inner join '+@TB_RT_Data_tmp+' d'+@TB_Hint+@CR+
                   '          on m.area_year = d.area_year'+@CR+
                   '         and d.kind = ''00353'''+@CR+
                   ' where m.kind = ''00801'' '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '00804'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '平均售或天數(365 / 週轉率)'
     set @Remark= '測度企業存貨週轉速度、銷貨能力高低及存貨量適當與否，銷貨天數愈低愈佳, Formula: 365天 / 00803'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       m.area_year,'+@CR+
                   '       ''00'' as area_month,'+@CR+ 
                   '       '''+@NonSet+''' as area,'+@CR+ 
                   '       '''+@CNonSet+''' as area_name,'+@CR+ 
                   '       case when m.amt=0'+@CR+
                   '            then 0'+@CR+
                   '            else Round(365 / m.amt, 4)'+@CR+
                   '        end as amt,'+@CR+
                   '       ''Y'' as Not_Accumulate,'+@CR+
                   '       3 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   ' where m.kind = ''00803'' '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '01001'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月各大分類實銷金額'
     set @Remark= '(請款日, 表身, 含稅, 扣折讓, 無加減項, 含退貨, 含服務單, 原大分類)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       year,'+@CR+
                   '       month,'+@CR+
                   '       area,'+@CR+
                   '       area_name,'+@CR+
                   '       sum(amt) as amt,'+@CR+
                   '       Not_Accumulate,'+@CR+
                   '       Data_Type,'+@CR+
                   '       Remark '+@CR+
                   '  from ('+@CR+
                   '       select chg_sp_pdate_year as year,'+@CR+
                   '              chg_sp_pdate_month as month,'+@CR+
                   '              IsNull(chg_skno_bkind, '''+@NonSet+''') as area,'+@CR+
                   '              IsNull(chg_skno_bkind_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '              Sum(IsNull(chg_sd_stot_tax, 0)) as amt,'+@CR+
                   '              '''' as Not_Accumulate,'+@CR+
                   '              0 as Data_Type,'+@CR+
                   '              '''+@ReMark+''' as Remark'+@CR+
                   '         from fact_sslpdt'+@TB_Hint+@CR+
                   '        where 1=1'+@CR+
                   '          and chg_sp_pdate_year >= ''2013'''+@CR+
                   '          and (sd_slip_fg =''2'' Or sd_slip_fg =''C'' Or sd_slip_fg = ''3'')'+@CR+
                   '          and chg_sp_pdate_year is not null'+@CR+
                   -- 2017/08/08 Rickliu 副總指示客編 IT、ZZ 開頭不認列業績，雖不認列業績但應收扔得收
                   '          and Chg_BU_NO Not In (''IT0000'', ''ZZ0000'') '+@CR+
                   '        group by chg_sp_pdate_year, chg_sp_pdate_month, chg_skno_bkind, chg_skno_bkind_name'+@CR+
                   '        union'+@CR+
                   '       select chg_sp_pdate_year as year,'+@CR+
                   '              chg_sp_pdate_month as month,'+@CR+
                   '              IsNull(Chg_sp_dis_flg '+@Collate+', '''+@NonSet+''') as area,'+@CR+
                   '              isnull(D2.Code_Name '+@Collate+', '''+@CNonSet+''') as area_name,'+@CR+
                   '              Sum(IsNull(Chg_sp_dis_tot2, 0)) as amt,'+@CR+
                   '              '''' as Not_Accumulate,'+@CR+
                   '              0 as Data_Type,'+@CR+
                   '              '''+@ReMark+''' as Remark'+@CR+
                   '         from fact_sslpdt'+@TB_Hint+@CR+
                   '              left join Ori_XLS#Sys_stockcode as D2'+@TB_Hint+@CR+
                   '                     ON D2.Code_Level = ''2'' AND Chg_sp_dis_flg '+@Collate+' = D2.Code_No '+@Collate+@CR+
                   '        where 1=1'+@CR+
                   '          and chg_sp_pdate_year >= ''2013'''+@CR+
                   '          and (sd_slip_fg =''2'' Or sd_slip_fg =''C'' Or sd_slip_fg = ''3'')'+@CR+
                   '          and chg_sp_pdate_year is not null'+@CR+
                   -- 2017/08/08 Rickliu 副總指示客編 IT、ZZ 開頭不認列業績，雖不認列業績但應收扔得收
                   '          and Chg_BU_NO Not In (''IT0000'', ''ZZ0000'') '+@CR+
                   '        group by chg_sp_pdate_year, chg_sp_pdate_month, Chg_sp_dis_flg, D2.Code_Name'+@CR+
                   ') as m'+@CR+
                   ' group by year, month, area, area_name, Not_Accumulate, Data_Type, Remark' 

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '01002'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月各大分類銷售金額'
     set @Remark= '(請款日, 表身, 含稅, 無折讓, 無加減項, 含服務單, 原大分類)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_pdate_year,'+@CR+
                   '       chg_sp_pdate_month,'+@CR+
                   '       IsNull(chg_skno_bkind, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(chg_skno_bkind_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(chg_sd_stot_tax, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslpdt'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and chg_sp_pdate_year >= ''2013'''+@CR+
                   '   and (sd_slip_fg =''2'' Or sd_slip_fg =''C'')'+@CR+
                   '   and chg_sp_pdate_year is not null'+@CR+
                   -- 2017/08/08 Rickliu 副總指示客編 IT、ZZ 開頭不認列業績，雖不認列業績但應收扔得收
                   '   and Chg_BU_NO Not In (''IT0000'', ''ZZ0000'') '+@CR+
                   ' group by chg_sp_pdate_year, chg_sp_pdate_month, chg_skno_bkind, chg_skno_bkind_name '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '01003'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月各大分類退貨金額'
     set @Remark= '(請款日, 表身, 含稅, 無折讓, 無加減項, 原大分類)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_pdate_year,'+@CR+
                   '       chg_sp_pdate_month,'+@CR+
                   '       IsNull(chg_skno_bkind, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(chg_skno_bkind_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(chg_sd_stot_tax, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslpdt'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and chg_sp_pdate_year >= ''2013'''+@CR+
                   '   and sd_slip_fg =''3'''+@CR+
                   '   and chg_sp_pdate_year is not null'+@CR+
                   -- 2017/08/08 Rickliu 副總指示客編 IT、ZZ 開頭不認列業績，雖不認列業績但應收扔得收
                   '   and Chg_BU_NO Not In (''IT0000'', ''ZZ0000'') '+@CR+
                   ' group by chg_sp_pdate_year, chg_sp_pdate_month, chg_skno_bkind, chg_skno_bkind_name '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '01004'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月大分類目標(僅百貨, 3C, 化工, 工具, 電裝及其他)'
     set @Remark= '分類目標由 ori_xls#Target_Stock_BKind_Personal 而來'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       year as area_year,'+@CR+
                   '       month as area_month,'+@CR+
                   '       BKind as area,'+@CR+
                   '       ''年度 ''+BKind+''目標'' as area_name,'+@CR+
                   '       Sum(Isnull(amt, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_Target+@TB_Hint+@CR+
                   ' group by year, month, BKind'+@CR+
                   ' order by year, month, bKind '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '01005'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '年度大分類目標(僅百貨, 3C, 化工, 工具, 電裝及其他)'
     set @Remark= '年度分類目標由 01004 加總而來'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       area_year,'+@CR+
                   '       '''+@NonSet+''' as area_month,'+@CR+
                   '       IsNull(area, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(area_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(Isnull(amt, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and kind =''01004'''+@CR+
                   ' group by area_year, area, area_name '
                   
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '01011'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月大分類實銷總額(僅百貨, 3C, 化工, 工具, 電裝及其他)'
     set @Remark= '每月大分類實銷總額由 01001 加總而來'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       area_year,'+@CR+
                   '       area_month,'+@CR+
                   '       IsNull(area, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(area_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Isnull(amt, 0) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and kind=''01001'''+@CR+
                   '   and (area =''AA'' Or area =''AB'' Or area = ''AC'' Or area =''AD'' Or area = ''AE'')'+@CR+
                   ' union'+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       area_year,'+@CR+
                   '       area_month,'+@CR+
                   '       ''OT'' as area,'+@CR+
                   '       ''其他'' as area_name,'+@CR+
                   '       Sum(Isnull(amt, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and kind=''01001'''+@CR+
                   '   and (area <>''AA'' and area <>''AB'' and area <> ''AC'' and area <>''AD'' and area <> ''AE'')'+@CR+
                   ' group by area_year, area_month '
                   
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '01012'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月大分類銷售總額(僅百貨, 3C, 化工, 工具, 電裝及其他)'
     set @Remark= '每月大分類銷售總額由 01002 加總而來'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       area_year,'+@CR+
                   '       area_month,'+@CR+
                   '       IsNull(area, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(area_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       IsNull(amt, 0) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and kind=''01002'''+@CR+
                   '   and (area =''AA'' Or area =''AB'' Or area = ''AC'' Or area =''AD'' Or area = ''AE'')'+@CR+
                   ' union'+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       area_year,'+@CR+
                   '       area_month,'+@CR+
                   '       ''OT'' as area,'+@CR+
                   '       ''其他'' as area_name,'+@CR+
                   '       Sum(Isnull(amt, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and kind=''01002'''+@CR+
                   '   and (area <>''AA'' and area <>''AB'' and area <> ''AC'' and area <>''AD'' and area <> ''AE'')'+@CR+
                   ' group by area_year, area_month '
                   
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '01013'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月大分類退貨總額(僅百貨, 3C, 化工, 工具, 電裝及其他)'
     set @Remark= '每月大分類退貨總額由 01003 加總而來'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       area_year,'+@CR+
                   '       area_month,'+@CR+
                   '       IsNull(area, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(area_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       IsNull(amt, 0) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and kind=''01003'''+@CR+
                   '   and (area =''AA'' Or area =''AB'' Or area = ''AC'' Or area =''AD'' Or area = ''AE'')'+@CR+
                   ' union'+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       area_year,'+@CR+
                   '       area_month,'+@CR+
                   '       ''OT'' as area,'+@CR+
                   '       ''其他'' as area_name,'+@CR+
                   '       Sum(IsNull(amt, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and kind=''01003'''+@CR+
                   '   and (area <>''AA'' and area <>''AB'' and area <> ''AC'' and area <>''AD'' and area <> ''AE'')'+@CR+
                   ' group by area_year, area_month '
                   
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '01014'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每年大分類實銷總額(僅百貨, 3C, 化工, 工具, 電裝及其他)'
     set @Remark= '每年大分類實銷總額由 01001 加總而來'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       area_year,'+@CR+
                   '       ''00'' as area_month,'+@CR+
                   '       IsNull(area, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(area_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(amt, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and kind=''01001'''+@CR+
                   '   and (area =''AA'' Or area =''AB'' Or area = ''AC'' Or area =''AD'' Or area = ''AE'')'+@CR+
                   ' group by area_year, area, area_name'+@CR+
                   ' union'+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       area_year,'+@CR+
                   '       ''00'' as area_month,'+@CR+
                   '       ''OT'' as area,'+@CR+
                   '       ''其他'' as area_name,'+@CR+
                   '       Sum(IsNull(amt, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and kind=''01001'''+@CR+
                   '   and (area <>''AA'' and area <>''AB'' and area <> ''AC'' and area <>''AD'' and area <> ''AE'')'+@CR+
                   ' group by area_year '
                   
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '01015'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     --set @Kind_Name = '每月大分類實銷佔全年度大分類實銷佔比(僅百貨, 3C, 化工, 工具, 電裝及其他)'
     set @Kind_Name = '每月大分類實銷佔比(僅百貨, 3C, 化工, 工具, 電裝及其他)'
     --set @Remark= '每月大分類實銷佔比 Formula: 01011 / 01014'
     set @Remark= '每月大分類實銷佔比 Formula: 01011 / 00061'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       m.area_year,'+@CR+
                   '       m.area_month,'+@CR+
                   '       IsNull(m.area, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(m.area_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       case'+@CR+
                   '         when Isnull(d.amt, 0) = 0 then 0'+@CR+
                   '         else Round(Isnull(m.amt, 0) / d.amt, 4)'+@CR+
                   '       end as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       3 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '       inner join '+@TB_RT_Data_tmp+' d'+@TB_Hint+@CR+
                   '          on m.area_year=d.area_year'+@CR+
                   '         and m.area_month=d.area_month'+@CR+
                   --'         and m.area=d.area'+@CR+
                   '         and d.kind =''00061'''+@CR+
                   ' where 1=1'+@CR+
                   '   and m.kind=''01011'' '
                   
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '01016'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每年大分類實銷佔比(僅百貨, 3C, 化工, 工具, 電裝及其他)'
     set @Remark= '每年大分類實銷佔比 Formula: 01014 / Sum(01014)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       m.area_year,'+@CR+
                   '       m.area_month,'+@CR+
                   '       IsNull(m.area, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(m.area_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       case'+@CR+
                   '         when Isnull(d.tot_amt, 0) = 0 then 0'+@CR+
                   '         else Round(Isnull(m.amt, 0) / d.tot_amt, 4)'+@CR+
                   '       end as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       3 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '       inner join'+@CR+
                   '       (select area_year, Sum(IsNull(amt, 0)) as tot_amt'+@CR+
                   '          from '+@TB_RT_Data_tmp+@TB_Hint+@CR+
                   '         where 1=1'+@CR+
                   '           and kind =''01014'''+@CR+
                   '         group by area_year'+@CR+
                   '       ) d'+@CR+
                   '        on m.area_year=d.area_year'+@CR+
                   ' where 1=1'+@CR+
                   '   and m.kind=''01014'' '
                   
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '01017'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月大分類實銷達成率(僅百貨, 3C, 化工, 工具, 電裝及其他)'
     set @Remark= '每月大分類實銷達成率 Formula: 01011 / 01004 加總而來'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       m.area_year,'+@CR+
                   '       m.area_month,'+@CR+
                   '       IsNull(m.area, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(m.area_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       case'+@CR+
                   '         when Isnull(d.amt, 0) = 0 then 0'+@CR+
                   '         else Round(Isnull(m.amt, 0) / d.amt, 4)'+@CR+
                   '       end as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       3 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '       inner join '+@TB_RT_Data_tmp+' d'+@TB_Hint+@CR+
                   '          on m.area_year=d.area_year'+@CR+
                   '         and m.area_month=d.area_month'+@CR+
                   '         and m.area=d.area'+@CR+
                   '         and d.kind =''01004'''+@CR+
                   ' where 1=1'+@CR+
                   '   and m.kind=''01011'' '
                   
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '01018'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每年大分類實銷達成率(僅百貨, 3C, 化工, 工具, 電裝及其他)'
     set @Remark= '每年大分類實銷達成率 Formula: 01014 / 01005 加總而來'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       m.area_year,'+@CR+
                   '       m.area_month,'+@CR+
                   '       IsNull(m.area, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(m.area_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       case'+@CR+
                   '         when Isnull(d.amt, 0) = 0 then 0'+@CR+
                   '         else Round(Isnull(m.amt, 0) / d.amt, 4)'+@CR+
                   '       end as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       3 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '       inner join '+@TB_RT_Data_tmp+' d'+@TB_Hint+@CR+
                   '          on m.area_year=d.area_year'+@CR+
                   '         and m.area_month=d.area_month'+@CR+
                   '         and m.area=d.area'+@CR+
                   '         and d.kind =''01005'''+@CR+
                   ' where 1=1'+@CR+
                   '   and m.kind=''01014'' '
                   
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '01050'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月各小分類實銷金額'
     set @Remark= '(依請款日, 表身, 含稅, 無折讓, 無加減項, 含退貨, 含服務單, 原始分類)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_pdate_year,'+@CR+
                   '       chg_sp_pdate_month,'+@CR+
                   '       IsNull(chg_skno_skind, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(chg_skno_skind_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(chg_sd_stot_tax, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslpdt'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and chg_sp_pdate_year >= ''2013'''+@CR+
                   '   and (sd_slip_fg =''2'' Or sd_slip_fg =''C'' Or sd_slip_fg = ''3'')'+@CR+
                   '   and chg_sp_pdate_year is not null'+@CR+
                   -- 2017/08/08 Rickliu 副總指示客編 IT、ZZ 開頭不認列業績，雖不認列業績但應收扔得收
                   '   and Chg_BU_NO Not In (''IT0000'', ''ZZ0000'') '+@CR+
                   ' group by chg_sp_pdate_year, chg_sp_pdate_month, chg_skno_skind, chg_skno_skind_name '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '01051'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月各小分類實銷數量'
     set @Remark= '(依請款日, 表身, 含退貨, 含服務單, 原始分類)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_pdate_year,'+@CR+
                   '       chg_sp_pdate_month,'+@CR+
                   '       IsNull(chg_skno_skind, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(chg_skno_skind_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(chg_sd_qty, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslpdt'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and chg_sp_pdate_year >= ''2013'''+@CR+
                   '   and (sd_slip_fg =''2'' Or sd_slip_fg =''C'' Or sd_slip_fg = ''3'')'+@CR+
                   '   and chg_sp_pdate_year is not null'+@CR+
                   -- 2017/08/08 Rickliu 副總指示客編 IT、ZZ 開頭不認列業績，雖不認列業績但應收扔得收
                   '   and Chg_BU_NO Not In (''IT0000'', ''ZZ0000'') '+@CR+
                   ' group by chg_sp_pdate_year, chg_sp_pdate_month, chg_skno_skind, chg_skno_skind_name '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '01052'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月小分類實銷金額佔比'
     set @Remark= '每月小分類實銷金額佔比 Formula: 01050 / Sum(01050)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       m.area_year,'+@CR+
                   '       m.area_month,'+@CR+
                   '       IsNull(m.area, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(m.area_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       case'+@CR+
                   '         when Isnull(d.tot_amt, 0) = 0 then 0'+@CR+
                   '         else Round(Isnull(m.amt, 0) / d.tot_amt, 4)'+@CR+
                   '       end as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       3 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '       inner join'+@CR+
                   '       (select area_year, Sum(IsNull(amt, 0)) as tot_amt'+@CR+
                   '          from '+@TB_RT_Data_tmp+@TB_Hint+@CR+
                   '         where 1=1'+@CR+
                   '           and kind =''01050'''+@CR+
                   '         group by area_year'+@CR+
                   '       ) d'+@CR+
                   '        on m.area_year=d.area_year'+@CR+
                   ' where 1=1'+@CR+
                   '   and m.kind=''01050'' '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '01060'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月各小分類銷售金額'
     set @Remark= '(依請款日, 表身, 含稅, 無折讓, 無加減項, 含服務單, 原始分類)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_pdate_year,'+@CR+
                   '       chg_sp_pdate_month,'+@CR+
                   '       IsNull(chg_skno_skind, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(chg_skno_skind_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(chg_sd_stot_tax, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslpdt'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and chg_sp_pdate_year >= ''2013'''+@CR+
                   '   and (sd_slip_fg =''2'' Or sd_slip_fg =''C'')'+@CR+
                   '   and chg_sp_pdate_year is not null'+@CR+
                   -- 2017/08/08 Rickliu 副總指示客編 IT、ZZ 開頭不認列業績，雖不認列業績但應收扔得收
                   '   and Chg_BU_NO Not In (''IT0000'', ''ZZ0000'') '+@CR+
                   ' group by chg_sp_pdate_year, chg_sp_pdate_month, chg_skno_skind, chg_skno_skind_name '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '01061'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月各小分類銷售數量'
     set @Remark= '(依請款日, 表身, 含服務單, 原始分類)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_pdate_year,'+@CR+
                   '       chg_sp_pdate_month,'+@CR+
                   '       IsNull(chg_skno_skind, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(chg_skno_skind_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(chg_sd_qty, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslpdt'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and chg_sp_pdate_year >= ''2013'''+@CR+
                   '   and (sd_slip_fg =''2'' Or sd_slip_fg =''C'')'+@CR+
                   '   and chg_sp_pdate_year is not null'+@CR+
                   -- 2017/08/08 Rickliu 副總指示客編 IT、ZZ 開頭不認列業績，雖不認列業績但應收扔得收
                   '   and Chg_BU_NO Not In (''IT0000'', ''ZZ0000'') '+@CR+
                   ' group by chg_sp_pdate_year, chg_sp_pdate_month, chg_skno_skind, chg_skno_skind_name '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '01062'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月小分類銷售金額佔比'
     set @Remark= '每月小分類銷售金額佔比 Formula: 01060 / Sum(01060)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       m.area_year,'+@CR+
                   '       m.area_month,'+@CR+
                   '       IsNull(m.area, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(m.area_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       case'+@CR+
                   '         when Isnull(d.tot_amt, 0) = 0 then 0'+@CR+
                   '         else Round(Isnull(m.amt, 0) / d.tot_amt, 4)'+@CR+
                   '       end as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       3 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '       inner join'+@CR+
                   '       (select area_year, Sum(IsNull(amt, 0)) as tot_amt'+@CR+
                   '          from '+@TB_RT_Data_tmp+@TB_Hint+@CR+
                   '         where 1=1'+@CR+
                   '           and kind =''01060'''+@CR+
                   '         group by area_year'+@CR+
                   '       ) d'+@CR+
                   '        on m.area_year=d.area_year'+@CR+
                   ' where 1=1'+@CR+
                   '   and m.kind=''01060'' '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '01070'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月各小分類退貨金額'
     set @Remark= '(依請款日, 表身, 含稅, 原始分類)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_pdate_year,'+@CR+
                   '       chg_sp_pdate_month,'+@CR+
                   '       IsNull(chg_skno_skind, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(chg_skno_skind_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(chg_sd_stot_tax, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslpdt'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and chg_sp_pdate_year >= ''2013'''+@CR+
                   '   and sd_slip_fg =''3'''+@CR+
                   '   and chg_sp_pdate_year is not null'+@CR+
                   -- 2017/08/08 Rickliu 副總指示客編 IT、ZZ 開頭不認列業績，雖不認列業績但應收扔得收
                   '   and Chg_BU_NO Not In (''IT0000'', ''ZZ0000'') '+@CR+
                   ' group by chg_sp_pdate_year, chg_sp_pdate_month, chg_skno_skind, chg_skno_skind_name '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '01071'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月各小分類退貨數量'
     set @Remark= '(依請款日, 表身, 原始分類)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_pdate_year,'+@CR+
                   '       chg_sp_pdate_month,'+@CR+
                   '       IsNull(chg_skno_skind, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(chg_skno_skind_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(chg_sd_qty, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslpdt'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and chg_sp_pdate_year >= ''2013'''+@CR+
                   '   and sd_slip_fg =''3'''+@CR+
                   '   and chg_sp_pdate_year is not null'+@CR+
                   -- 2017/08/08 Rickliu 副總指示客編 IT、ZZ 開頭不認列業績，雖不認列業績但應收扔得收
                   '   and Chg_BU_NO Not In (''IT0000'', ''ZZ0000'') '+@CR+
                   ' group by chg_sp_pdate_year, chg_sp_pdate_month, chg_skno_skind, chg_skno_skind_name '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '01072'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月小分類退貨金額佔比'
     set @Remark= '每月小分類退貨金額佔比 Formula: 01070 / Sum(01072)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       m.area_year,'+@CR+
                   '       m.area_month,'+@CR+
                   '       IsNull(m.area, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(m.area_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       case'+@CR+
                   '         when IsNull(d.tot_amt, 0) = 0 then 0'+@CR+
                   '         else Round(Isnull(m.amt, 0) / d.tot_amt, 4)'+@CR+
                   '       end as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       3 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '       inner join'+@CR+
                   '       (select area_year, Sum(IsNull(amt, 0)) as tot_amt'+@CR+
                   '          from '+@TB_RT_Data_tmp+@TB_Hint+@CR+
                   '         where 1=1'+@CR+
                   '           and kind =''01070'''+@CR+
                   '         group by area_year'+@CR+
                   '       ) d'+@CR+
                   '        on m.area_year=d.area_year'+@CR+
                   ' where 1=1'+@CR+
                   '   and m.kind=''01070'' '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '01100'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月業務大分類目標-百貨(個人業績)'
     set @Remark= '(大分類為 AA)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       Year as area_year,'+@CR+
                   '       Month as area_month,'+@CR+
                   '       IsNull(ct_sales, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(sales_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(amt, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_Target+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and Year >= ''2013'''+@CR+
                   '   and BKind =''AA'''+@CR+
                   ' group by Year, Month, ct_sales, sales_name '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '01101'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月業務大分類實銷金額-百貨(個人業績)'
     set @Remark= '(依請款日, 表身, 含稅, 無加減項, 扣折讓, 含退貨, 含服務單, 大分類為 AA)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       area_year,'+@CR+
                   '       area_month,'+@CR+
                   '       area,'+@CR+
                   '       area_name,'+@CR+
                   '       Sum(IsNull(amt1+amt2, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from ('+@CR+
                   '        select chg_sp_pdate_year as area_year,'+@CR+
                   '               chg_sp_pdate_month as area_month,'+@CR+
                   '               IsNull(sp_sales, '''+@NonSet+''') as area,'+@CR+
                   '               IsNull(chg_sales_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '               case when chg_skno_bkind =''AA'' then IsNull(chg_sd_stot_tax, 0) else 0 end as Amt1,'+@CR+
                   '               case when Chg_sp_dis_flg= ''AA'' then IsNull(Chg_sp_dis_tot2, 0) else 0 end as Amt2,'+@CR+
                   '               case when (chg_skno_bkind =''AA'' or Chg_sp_dis_flg = ''AA'') then ''AA'' else '''' end as chg_skno_bkind '+@CR+
                   '          from fact_sslpdt'+@TB_Hint+@CR+
                   '         where 1=1'+@CR+
                   '           and chg_sp_pdate_year >= ''2013'''+@CR+
                   '           and (sd_slip_fg =''2'' Or sd_slip_fg =''C'' Or sd_slip_fg = ''3'')'+@CR+
                   '           and chg_sp_pdate_year is not null'+@CR+
                   '           and (chg_skno_bkind = ''AA'' or Chg_sp_dis_flg = ''AA'') '+@CR+
                   -- 2017/08/08 Rickliu 副總指示客編 IT、ZZ 開頭不認列業績，雖不認列業績但應收扔得收
                   '           and Chg_BU_NO Not In (''IT0000'', ''ZZ0000'') '+@CR+
                   --'         group by chg_sp_pdate_year, chg_sp_pdate_month, sp_sales, chg_sales_name, chg_skno_bkind'+@CR+
                   --'        union'+@CR+
                   --'        select chg_sp_pdate_year as area_year,'+@CR+
                   --'               chg_sp_pdate_month as area_month,'+@CR+
                   --'               IsNull(sp_sales, '''+@NonSet+''') as area,'+@CR+
                   --'               IsNull(chg_sales_name, '''+@CNonSet+''') as area_name,'+@CR+
                   --'               Sum(IsNull(Chg_sp_dis_tot2, 0)) as amt,'+@CR+
                   --'               Chg_sp_dis_flg  '+@Collate+' as chg_skno_bkind'+@CR+
                   --'          from fact_sslpdt'+@TB_Hint+@CR+
                   --'         where 1=1'+@CR+
                   --'           and chg_sp_pdate_year >= ''2013'''+@CR+
                   --'           and (sd_slip_fg =''2'' Or sd_slip_fg =''C'' Or sd_slip_fg = ''3'')'+@CR+
                   --'           and chg_sp_pdate_year is not null'+@CR+
                   --'           and chg_skno_bkind = ''AA'''+@CR+
                   --'         group by chg_sp_pdate_year, chg_sp_pdate_month, sp_sales, chg_sales_name, Chg_sp_dis_flg'+@CR+
                   '       ) as a '+@CR+
                   ' where 1=1 '+@CR+
                   '   and chg_skno_bkind = ''AA'' '+@CR+
                   ' group by area_year, area_month, area, area_name'+@CR+ 
                   ' order by 3, 4, 5'

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '01102'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月業務大分類銷售金額-百貨(個人業績)'
     set @Remark='(依請款日, 表身, 含稅, 無加減項, 無折讓, 含服務單, 大分類為 AA)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_pdate_year as area_year,'+@CR+
                   '       chg_sp_pdate_month as area_month,'+@CR+
                   '       IsNull(sp_sales, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(chg_sales_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(chg_sd_stot_tax, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslpdt'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and chg_sp_pdate_year >= ''2013'''+@CR+
                   '   and chg_skno_bkind = ''AA'''+@CR+
                   '   and (sd_slip_fg =''2'' Or sd_slip_fg =''C'')'+@CR+
                   '   and chg_sp_pdate_year is not null'+@CR+
                   -- 2017/08/08 Rickliu 副總指示客編 IT、ZZ 開頭不認列業績，雖不認列業績但應收扔得收
                   '   and Chg_BU_NO Not In (''IT0000'', ''ZZ0000'') '+@CR+
                   ' group by chg_sp_pdate_year, chg_sp_pdate_month, sp_sales, chg_sales_name '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
    
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '01103'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月業務大分類退貨金額-百貨(個人業績)'
     set @Remark= '(依請款日, 表身, 含稅, 無加減項, 無折讓, 大分類為 AA)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_pdate_year as area_year,'+@CR+
                   '       chg_sp_pdate_month as area_month,'+@CR+
                   '       IsNull(sp_sales, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(chg_sales_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(chg_sd_stot_tax, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslpdt'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and chg_sp_pdate_year >= ''2013'''+@CR+
                   '   and chg_skno_bkind = ''AA'''+@CR+
                   '   and sd_slip_fg =''3'''+@CR+
                   '   and chg_sp_pdate_year is not null'+@CR+
                   -- 2017/08/08 Rickliu 副總指示客編 IT、ZZ 開頭不認列業績，雖不認列業績但應收扔得收
                   '   and Chg_BU_NO Not In (''IT0000'', ''ZZ0000'') '+@CR+
                   ' group by chg_sp_pdate_year, chg_sp_pdate_month, sp_sales, chg_sales_name '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '01104'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月業務大分類實銷金額達成率-百貨(個人業績)'
     set @Remark= '(依請款日)(Formula: 01101 / 百貨目標, 目標資料: ori_xls#Target_Stock_BKind_Personal.Target_AA)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       m.area_year,'+@CR+
                   '       m.area_month,'+@CR+
                   '       IsNull(m.area, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(m.area_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       case'+@CR+
                   '         when IsNull(d.amt, 0) = 0 then 0'+@CR+
                   '         else Round(isnull(m.amt, 0) / isnull(d.amt, 0), 4)'+@CR+
                   '       end as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       3 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '       inner join '+@TB_RT_Data_tmp+' d'+@TB_Hint+@CR+
                   '          on m.area_year = d.area_year'+@CR+
                   '         and m.area_month = d.area_month'+@CR+
                   '         and m.area '+@Collate+'=d.area '+@Collate+@CR+
                   ' where 1=1'+@CR+
                   '   and m.area_year >= ''2013'''+@CR+
                   '   and m.kind =''01101'''+@CR+
                   '   and d.kind =''01100'' '
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '01105'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '年度業務大分類實銷金額達成率-百貨(個人業績)'
     set @Remark= '(依請款日)(Formula: 01101 / 百貨目標, 目標資料: ori_xls#Target_Stock_BKind_Personal.Target_AA)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       m.area_year,'+@CR+
                   '       ''00'' as area_month,'+@CR+
                   '       IsNull(m.area, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(m.area_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       case'+@CR+
                   '         when IsNull(d.Target_Amt, 0) = 0 then 0'+@CR+
                   '         else Round(Sum(Isnull(m.amt, 0)) / d.Target_Amt, 4)'+@CR+
                   '       end as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       3 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '       inner join'+@CR+
                   '         (select area_year, area, Sum(Isnull(amt, 0)) as Target_Amt'+@CR+
                   '            from '+@TB_RT_Data_tmp+' d'+@TB_Hint+@CR+
                   '           where kind=''01100'''+@CR+
                   '           group by area_year, area'+@CR+
                   '         ) D'+@CR+
                   '          on m.area_year=d.area_year'+@CR+
                   '         and m.area '+@Collate+'=d.area '+@Collate+@CR+
                   ' where 1=1'+@CR+
                   '   and m.area_year >= ''2013'''+@CR+
                   '   and m.kind =''01101'''+@CR+
                   ' group by m.area_year, m.area, area_name, Target_Amt'

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end 
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '01106'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月業務大分類實銷金額-百貨(個人業績毛利)'
     set @Remark= '(依請款日, 表身, 未稅, 無加減項, 無折讓, 含退貨, 含服務單, 大分類為 AA)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_pdate_year as area_year,'+@CR+
                   '       chg_sp_pdate_month as area_month,'+@CR+
                   '       IsNull(sp_sales, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(chg_sales_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(Chg_Profit, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslpdt'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and chg_sp_pdate_year >= ''2013'''+@CR+
                   '   and chg_skno_bkind = ''AA'''+@CR+
                   '   and (sd_slip_fg =''2'' Or sd_slip_fg =''C'' Or sd_slip_fg = ''3'')'+@CR+
                   '   and chg_sp_pdate_year is not null'+@CR+
                   -- 2017/08/08 Rickliu 副總指示客編 IT、ZZ 開頭不認列業績，雖不認列業績但應收扔得收
                   '   and Chg_BU_NO Not In (''IT0000'', ''ZZ0000'') '+@CR+
                   ' group by chg_sp_pdate_year, chg_sp_pdate_month, sp_sales, chg_sales_name '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '01120'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月客戶大分類目標-百貨'
     set @Remark= '(大分類為 AA)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       Year as area_year,'+@CR+
                   '       Month as area_month,'+@CR+
                   '       IsNull(ct_no, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(ct_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(amt, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_Target+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and Year >= ''2013'''+@CR+
                   '   and BKind =''AA'''+@CR+
                   ' group by Year, Month, ct_no, ct_name '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '01121'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月客戶大分類實銷金額-百貨'
     set @Remark= '(依請款日, 表身, 含稅, 無加減項, 扣折讓, 含退貨, 含服務單, 大分類為 AA)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       area_year,'+@CR+
                   '       area_month,'+@CR+
                   '       area,'+@CR+
                   '       area_name,'+@CR+
                   '       Sum(IsNull(amt1+amt2, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from ('+@CR+
                   '        select chg_sp_pdate_year as area_year,'+@CR+
                   '               chg_sp_pdate_month as area_month,'+@CR+
                   '               IsNull(sd_ctno, '''+@NonSet+''') as area,'+@CR+
                   '               IsNull(chg_ct_sname, '''+@CNonSet+''') as area_name,'+@CR+
                   '               case when chg_skno_bkind =''AA'' then IsNull(chg_sd_stot_tax, 0) else 0 end as Amt1,'+@CR+
                   '               case when Chg_sp_dis_flg= ''AA'' then IsNull(Chg_sp_dis_tot2, 0) else 0 end as Amt2,'+@CR+
                   '               case when (chg_skno_bkind = ''AA'' or Chg_sp_dis_flg = ''AA'') then ''AA'' else '''' end as chg_skno_bkind '+@CR+
                   '          from fact_sslpdt'+@TB_Hint+@CR+
                   '         where 1=1'+@CR+
                   '           and chg_sp_pdate_year >= ''2013'''+@CR+
                   '           and (sd_slip_fg =''2'' Or sd_slip_fg =''C'' Or sd_slip_fg = ''3'')'+@CR+
                   '           and chg_sp_pdate_year is not null'+@CR+
                   '           and (chg_skno_bkind = ''AA'' or Chg_sp_dis_flg = ''AA'') '+@CR+
                   -- 2017/08/08 Rickliu 副總指示客編 IT、ZZ 開頭不認列業績，雖不認列業績但應收扔得收
                   '           and Chg_BU_NO Not In (''IT0000'', ''ZZ0000'') '+@CR+
                   --'         group by chg_sp_pdate_year, chg_sp_pdate_month, sd_ctno, chg_ct_sname, chg_skno_bkind'+@CR+
                   --'        union'+@CR+
                   --'        select chg_sp_pdate_year as area_year,'+@CR+
                   --'               chg_sp_pdate_month as area_month,'+@CR+
                   --'               IsNull(sd_ctno, '''+@NonSet+''') as area,'+@CR+
                   --'               IsNull(chg_ct_sname, '''+@CNonSet+''') as area_name,'+@CR+
                   --'               Sum(IsNull(Chg_sp_dis_tot2, 0)) as amt,'+@CR+
                   --'               Chg_sp_dis_flg  '+@Collate+' as chg_skno_bkind'+@CR+
                   --'          from fact_sslpdt'+@TB_Hint+@CR+
                   --'         where 1=1'+@CR+
                   --'           and chg_sp_pdate_year >= ''2013'''+@CR+
                   --'           and (sd_slip_fg =''2'' Or sd_slip_fg =''C'' Or sd_slip_fg = ''3'')'+@CR+
                   --'           and chg_sp_pdate_year is not null'+@CR+
                   --'           and chg_skno_bkind = ''AA'''+@CR+
                   --'         group by chg_sp_pdate_year, chg_sp_pdate_month, sd_ctno, chg_ct_sname, Chg_sp_dis_flg'+@CR+
                   '       ) as a'+@CR+
                   ' where 1=1 '+@CR+
                   '   and chg_skno_bkind = ''AA'' '+@CR+
                   ' group by area_year, area_month, area, area_name'+@CR+ 
                   ' order by 3, 4, 5'

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '01122'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月客戶大分類銷售金額-百貨'
     set @Remark='(依請款日, 表身, 含稅, 無加減項, 無折讓, 含服務單, 大分類為 AA)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_pdate_year as area_year,'+@CR+
                   '       chg_sp_pdate_month as area_month,'+@CR+
                   '       IsNull(sd_ctno, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(chg_ct_sname, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(chg_sd_stot_tax, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslpdt'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and chg_sp_pdate_year >= ''2013'''+@CR+
                   '   and chg_skno_bkind = ''AA'''+@CR+
                   '   and (sd_slip_fg =''2'' Or sd_slip_fg =''C'')'+@CR+
                   '   and chg_sp_pdate_year is not null'+@CR+
                   -- 2017/08/08 Rickliu 副總指示客編 IT、ZZ 開頭不認列業績，雖不認列業績但應收扔得收
                   '   and Chg_BU_NO Not In (''IT0000'', ''ZZ0000'') '+@CR+
                   ' group by chg_sp_pdate_year, chg_sp_pdate_month, sd_ctno, chg_ct_sname '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
    
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '01123'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月客戶大分類退貨金額-百貨'
     set @Remark= '(依請款日, 表身, 含稅, 無加減項, 無折讓, 大分類為 AA)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_pdate_year as area_year,'+@CR+
                   '       chg_sp_pdate_month as area_month,'+@CR+
                   '       IsNull(sd_ctno, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(chg_ct_sname, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(chg_sd_stot_tax, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslpdt'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and chg_sp_pdate_year >= ''2013'''+@CR+
                   '   and chg_skno_bkind = ''AA'''+@CR+
                   '   and sd_slip_fg =''3'''+@CR+
                   '   and chg_sp_pdate_year is not null'+@CR+
                   -- 2017/08/08 Rickliu 副總指示客編 IT、ZZ 開頭不認列業績，雖不認列業績但應收扔得收
                   '   and Chg_BU_NO Not In (''IT0000'', ''ZZ0000'') '+@CR+
                   ' group by chg_sp_pdate_year, chg_sp_pdate_month, sd_ctno, chg_ct_sname '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '01124'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月客戶大分類實銷金額達成率-百貨'
     set @Remark= '(依請款日)(Formula: 01101 / 百貨目標, 目標資料: ori_xls#Target_Stock_BKind_Personal.Target_AA)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       m.area_year,'+@CR+
                   '       m.area_month,'+@CR+
                   '       IsNull(m.area, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(m.area_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       case'+@CR+
                   '         when IsNull(d.amt, 0) = 0 then 0'+@CR+
                   '         else Round(isnull(m.amt, 0) / isnull(d.amt, 0), 4)'+@CR+
                   '       end as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       3 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '       inner join '+@TB_RT_Data_tmp+' d'+@TB_Hint+@CR+
                   '          on m.area_year = d.area_year'+@CR+
                   '         and m.area_month = d.area_month'+@CR+
                   '         and m.area '+@Collate+'=d.area '+@Collate+@CR+
                   ' where 1=1'+@CR+
                   '   and m.area_year >= ''2013'''+@CR+
                   '   and m.kind =''01121'''+@CR+
                   '   and d.kind =''01120'' '
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '01125'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '年度業務大分類實銷金額達成率-百貨'
     set @Remark= '(依請款日)(Formula: 01101 / 百貨目標, 目標資料: ori_xls#Target_Stock_BKind_Personal.Target_AA)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       m.area_year,'+@CR+
                   '       ''00'' as area_month,'+@CR+
                   '       IsNull(m.area, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(m.area_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       case'+@CR+
                   '         when IsNull(d.Target_Amt, 0) = 0 then 0'+@CR+
                   '         else Round(Sum(Isnull(m.amt, 0)) / d.Target_Amt, 4)'+@CR+
                   '       end as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       3 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '       inner join'+@CR+
                   '         (select area_year, area, Sum(Isnull(amt, 0)) as Target_Amt'+@CR+
                   '            from '+@TB_RT_Data_tmp+' d'+@TB_Hint+@CR+
                   '           where kind=''01120'''+@CR+
                   '           group by area_year, area'+@CR+
                   '         ) D'+@CR+
                   '          on m.area_year=d.area_year'+@CR+
                   '         and m.area '+@Collate+'=d.area '+@Collate+@CR+
                   ' where 1=1'+@CR+
                   '   and m.area_year >= ''2013'''+@CR+
                   '   and m.kind =''01121'''+@CR+
                   ' group by m.area_year, m.area, area_name, Target_Amt'

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end 
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '01126'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月客戶大分類實銷金額-百貨(客戶業績毛利)'
     set @Remark= '(依請款日, 表身, 未稅, 無加減項, 無折讓, 含退貨, 含服務單, 大分類為 AA)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_pdate_year as area_year,'+@CR+
                   '       chg_sp_pdate_month as area_month,'+@CR+
                   '       IsNull(sd_ctno, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(chg_ct_sname, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(Chg_Profit, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslpdt'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and chg_sp_pdate_year >= ''2013'''+@CR+
                   '   and chg_skno_bkind = ''AA'''+@CR+
                   '   and (sd_slip_fg =''2'' Or sd_slip_fg =''C'' Or sd_slip_fg = ''3'')'+@CR+
                   '   and chg_sp_pdate_year is not null'+@CR+
                   -- 2017/08/08 Rickliu 副總指示客編 IT、ZZ 開頭不認列業績，雖不認列業績但應收扔得收
                   '   and Chg_BU_NO Not In (''IT0000'', ''ZZ0000'') '+@CR+
                   ' group by chg_sp_pdate_year, chg_sp_pdate_month, sd_ctno, chg_ct_sname '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '01200'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月業務大分類目標-3C(個人業績)'
     set @Remark= '(大分類為 AB)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       Year as area_year,'+@CR+
                   '       Month as area_month,'+@CR+
                   '       IsNull(ct_sales, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(sales_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(amt, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_Target+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and Year >= ''2013'''+@CR+
                   '   and BKind =''AB'''+@CR+
                   ' group by Year, Month, ct_sales, sales_name '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '01201'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月業務大分類實銷金額-3C(個人業績)'
     set @Remark= '(依請款日, 表身, 含稅, 無加減項, 扣折讓, 含退貨, 含服務單, 大分類為 AB)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       area_year,'+@CR+
                   '       area_month,'+@CR+
                   '       area,'+@CR+
                   '       area_name,'+@CR+
                   '       Sum(IsNull(amt1+amt2, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from ('+@CR+
                   '        select chg_sp_pdate_year as area_year,'+@CR+
                   '               chg_sp_pdate_month as area_month,'+@CR+
                   '               IsNull(sp_sales, '''+@NonSet+''') as area,'+@CR+
                   '               IsNull(chg_sales_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '               case when chg_skno_bkind =''AB'' then IsNull(chg_sd_stot_tax, 0) else 0 end as Amt1,'+@CR+
                   '               case when Chg_sp_dis_flg= ''AB'' then IsNull(Chg_sp_dis_tot2, 0) else 0 end as Amt2,'+@CR+
                   '               case when (chg_skno_bkind = ''AB'' or Chg_sp_dis_flg = ''AB'') then ''AB'' else '''' end as chg_skno_bkind '+@CR+
                   '          from fact_sslpdt'+@TB_Hint+@CR+
                   '         where 1=1'+@CR+
                   '           and chg_sp_pdate_year >= ''2013'''+@CR+
                   '           and (sd_slip_fg =''2'' Or sd_slip_fg =''C'' Or sd_slip_fg = ''3'')'+@CR+
                   '           and chg_sp_pdate_year is not null'+@CR+
                   '           and (chg_skno_bkind = ''AB'' or Chg_sp_dis_flg = ''AB'') '+@CR+
                   -- 2017/08/08 Rickliu 副總指示客編 IT、ZZ 開頭不認列業績，雖不認列業績但應收扔得收
                   '           and Chg_BU_NO Not In (''IT0000'', ''ZZ0000'') '+@CR+
                   --'         group by chg_sp_pdate_year, chg_sp_pdate_month, sp_sales, chg_sales_name, chg_skno_bkind'+@CR+
                   --'        union'+@CR+
                   --'        select chg_sp_pdate_year as area_year,'+@CR+
                   --'               chg_sp_pdate_month as area_month,'+@CR+
                   --'               IsNull(sp_sales, '''+@NonSet+''') as area,'+@CR+
                   --'               IsNull(chg_sales_name, '''+@CNonSet+''') as area_name,'+@CR+
                   --'               Sum(IsNull(Chg_sp_dis_tot2, 0)) as amt,'+@CR+
                   --'               Chg_sp_dis_flg  '+@Collate+' as chg_skno_bkind'+@CR+
                   --'          from fact_sslpdt'+@TB_Hint+@CR+
                   --'         where 1=1'+@CR+
                   --'           and chg_sp_pdate_year >= ''2013'''+@CR+
                   --'           and (sd_slip_fg =''2'' Or sd_slip_fg =''C'' Or sd_slip_fg = ''3'')'+@CR+
                   --'           and chg_sp_pdate_year is not null'+@CR+
                   --'           and chg_skno_bkind = ''AB'''+@CR+
                   --'         group by chg_sp_pdate_year, chg_sp_pdate_month, sp_sales, chg_sales_name, Chg_sp_dis_flg'+@CR+
                   '       ) as a'+@CR+
                   ' where 1=1'+@CR+
                   ' group by area_year, area_month, area, area_name'+@CR+ 
                   ' order by 3, 4, 5'

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '01202'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月業務大分類銷售金額-3C(個人業績)'
     set @Remark= '(依請款日, 表身, 含稅, 無加減項, 無折讓, 含服務單, 大分類為 AB)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_pdate_year as area_year,'+@CR+
                   '       chg_sp_pdate_month as area_month,'+@CR+
                   '       IsNull(sp_sales, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(chg_sales_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(chg_sd_stot_tax, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslpdt'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and chg_sp_pdate_year >= ''2013'''+@CR+
                   '   and chg_skno_bkind = ''AB'''+@CR+
                   '   and (sd_slip_fg =''2'' Or sd_slip_fg =''C'')'+@CR+
                   '   and chg_sp_pdate_year is not null'+@CR+
                   -- 2017/08/08 Rickliu 副總指示客編 IT、ZZ 開頭不認列業績，雖不認列業績但應收扔得收
                   '   and Chg_BU_NO Not In (''IT0000'', ''ZZ0000'') '+@CR+
                   ' group by chg_sp_pdate_year, chg_sp_pdate_month, sp_sales, chg_sales_name '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '01203'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月業務大分類退貨金額-3C(個人業績)'
     set @Remark= '(依請款日, 表身, 含稅, 無加減項, 無折讓, 大分類為 AB)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_pdate_year as area_year,'+@CR+
                   '       chg_sp_pdate_month as area_month,'+@CR+
                   '       IsNull(sp_sales, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(chg_sales_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(chg_sd_stot_tax, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslpdt'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and chg_sp_pdate_year >= ''2013'''+@CR+
                   '   and chg_skno_bkind = ''AB'''+@CR+
                   '   and sd_slip_fg =''3'''+@CR+
                   '   and chg_sp_pdate_year is not null'+@CR+
                   -- 2017/08/08 Rickliu 副總指示客編 IT、ZZ 開頭不認列業績，雖不認列業績但應收扔得收
                   '   and Chg_BU_NO Not In (''IT0000'', ''ZZ0000'') '+@CR+
                   ' group by chg_sp_pdate_year, chg_sp_pdate_month, sp_sales, chg_sales_name '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '01204'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月業務大分類實銷金額達成率-3C(個人業績)'
     set @Remark= '(依請款日)(Formula: 01201 / 3C目標, 目標資料: ori_xls#Target_Stock_BKind_Personal.Target_AB)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       m.area_year,'+@CR+
                   '       m.area_month,'+@CR+
                   '       IsNull(m.area, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(m.area_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       case'+@CR+
                   '         when IsNull(d.amt, 0) = 0 then 0'+@CR+
                   '         else Round(isnull(m.amt, 0) / isnull(d.amt, 0), 4)'+@CR+
                   '       end as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       3 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '       inner join '+@TB_RT_Data_tmp+' d'+@TB_Hint+@CR+
                   '          on m.area_year = d.area_year'+@CR+
                   '         and m.area_month = d.area_month'+@CR+
                   '         and m.area '+@Collate+'=d.area '+@Collate+@CR+
                   ' where 1=1'+@CR+
                   '   and m.area_year >= ''2013'''+@CR+
                   '   and m.kind =''01201'''+@CR+
                   '   and d.kind =''01200'' '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '01205'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '年度業務大分類實銷金額達成率-3C(個人業績)'
     set @Remark= '(依請款日)(Formula: 01201 / 3C目標, 目標資料: ori_xls#Target_Stock_BKind_Personal.Target_AB)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       m.area_year,'+@CR+
                   '       ''00'' as area_month,'+@CR+
                   '       IsNull(m.area, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(m.area_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       case'+@CR+
                   '         when IsNull(d.Target_Amt, 0) = 0 then 0'+@CR+
                   '         else Round(Sum(Isnull(m.amt, 0)) / d.Target_Amt, 4)'+@CR+
                   '       end as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       3 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '       inner join'+@CR+
                   '         (select area_year, area, Sum(Isnull(amt, 0)) as Target_Amt'+@CR+
                   '            from '+@TB_RT_Data_tmp+' d'+@TB_Hint+@CR+
                   '           where kind=''01200'''+@CR+
                   '           group by area_year, area'+@CR+
                   '         ) D'+@CR+
                   '          on m.area_year=d.area_year'+@CR+
                   '         and m.area '+@Collate+'=d.area '+@Collate+@CR+
                   ' where 1=1'+@CR+
                   '   and m.area_year >= ''2013'''+@CR+
                   '   and m.kind =''01201'''+@CR+
                   ' group by m.area_year, m.area, area_name, Target_Amt'

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end 

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '01206'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月業務大分類實銷金額-3C(個人業績毛利)'
     set @Remark= '(依請款日, 表身, 未稅, 無加減項, 無折讓, 含退貨, 含服務單, 大分類為 AB)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_pdate_year as area_year,'+@CR+
                   '       chg_sp_pdate_month as area_month,'+@CR+
                   '       IsNull(sp_sales, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(chg_sales_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(Chg_Profit, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslpdt'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and chg_sp_pdate_year >= ''2013'''+@CR+
                   '   and chg_skno_bkind = ''AB'''+@CR+
                   '   and (sd_slip_fg =''2'' Or sd_slip_fg =''C'' Or sd_slip_fg = ''3'')'+@CR+
                   '   and chg_sp_pdate_year is not null'+@CR+
                   -- 2017/08/08 Rickliu 副總指示客編 IT、ZZ 開頭不認列業績，雖不認列業績但應收扔得收
                   '   and Chg_BU_NO Not In (''IT0000'', ''ZZ0000'') '+@CR+
                   ' group by chg_sp_pdate_year, chg_sp_pdate_month, sp_sales, chg_sales_name '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
    
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '01220'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月客戶大分類目標-3C'
     set @Remark= '(大分類為 AB)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       Year as area_year,'+@CR+
                   '       Month as area_month,'+@CR+
                   '       IsNull(ct_no, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(ct_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(amt, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_Target+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and Year >= ''2013'''+@CR+
                   '   and BKind =''AB'''+@CR+
                   ' group by Year, Month, ct_no, ct_name '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '01221'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月客戶大分類實銷金額-3C'
     set @Remark= '(依請款日, 表身, 含稅, 無加減項, 扣折讓, 含退貨, 含服務單, 大分類為 AB)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       area_year,'+@CR+
                   '       area_month,'+@CR+
                   '       area,'+@CR+
                   '       area_name,'+@CR+
                   '       Sum(IsNull(amt1+amt2, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from ('+@CR+
                   '        select chg_sp_pdate_year as area_year,'+@CR+
                   '               chg_sp_pdate_month as area_month,'+@CR+
                   '               IsNull(sd_ctno, '''+@NonSet+''') as area,'+@CR+
                   '               IsNull(chg_ct_sname, '''+@CNonSet+''') as area_name,'+@CR+
                   '               case when chg_skno_bkind =''AB'' then IsNull(chg_sd_stot_tax, 0) else 0 end as Amt1,'+@CR+
                   '               case when Chg_sp_dis_flg= ''AB'' then IsNull(Chg_sp_dis_tot2, 0) else 0 end as Amt2,'+@CR+
                   '               case when (chg_skno_bkind = ''AB'' or Chg_sp_dis_flg = ''AB'') then ''AB'' else '''' end as chg_skno_bkind '+@CR+
                   '          from fact_sslpdt'+@TB_Hint+@CR+
                   '         where 1=1'+@CR+
                   '           and chg_sp_pdate_year >= ''2013'''+@CR+
                   '           and (sd_slip_fg =''2'' Or sd_slip_fg =''C'' Or sd_slip_fg = ''3'')'+@CR+
                   '           and chg_sp_pdate_year is not null'+@CR+
                   '           and (chg_skno_bkind = ''AB'' or Chg_sp_dis_flg = ''AB'') '+@CR+
                   -- 2017/08/08 Rickliu 副總指示客編 IT、ZZ 開頭不認列業績，雖不認列業績但應收扔得收
                   '           and Chg_BU_NO Not In (''IT0000'', ''ZZ0000'') '+@CR+
                   --'         group by chg_sp_pdate_year, chg_sp_pdate_month, sd_ctno, chg_ct_sname, chg_skno_bkind'+@CR+
                   --'        union'+@CR+
                   --'        select chg_sp_pdate_year as area_year,'+@CR+
                   --'               chg_sp_pdate_month as area_month,'+@CR+
                   --'               IsNull(sd_ctno, '''+@NonSet+''') as area,'+@CR+
                   --'               IsNull(chg_ct_sname, '''+@CNonSet+''') as area_name,'+@CR+
                   --'               Sum(IsNull(Chg_sp_dis_tot2, 0)) as amt,'+@CR+
                   --'               Chg_sp_dis_flg  '+@Collate+' as chg_skno_bkind'+@CR+
                   --'          from fact_sslpdt'+@TB_Hint+@CR+
                   --'         where 1=1'+@CR+
                   --'           and chg_sp_pdate_year >= ''2013'''+@CR+
                   --'           and (sd_slip_fg =''2'' Or sd_slip_fg =''C'' Or sd_slip_fg = ''3'')'+@CR+
                   --'           and chg_sp_pdate_year is not null'+@CR+
                   --'           and chg_skno_bkind = ''AB'''+@CR+
                   --'         group by chg_sp_pdate_year, chg_sp_pdate_month, sd_ctno, chg_ct_sname, Chg_sp_dis_flg'+@CR+
                   '       ) as a'+@CR+
                   ' where 1=1 '+@CR+
                   '   and chg_skno_bkind = ''AB'' '+@CR+
                   ' group by area_year, area_month, area, area_name'+@CR+ 
                   ' order by 3, 4, 5'

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '01222'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月客戶大分類銷售金額-3C'
     set @Remark='(依請款日, 表身, 含稅, 無加減項, 無折讓, 含服務單, 大分類為 AB)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_pdate_year as area_year,'+@CR+
                   '       chg_sp_pdate_month as area_month,'+@CR+
                   '       IsNull(sd_ctno, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(chg_ct_sname, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(chg_sd_stot_tax, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslpdt'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and chg_sp_pdate_year >= ''2013'''+@CR+
                   '   and chg_skno_bkind = ''AB'''+@CR+
                   '   and (sd_slip_fg =''2'' Or sd_slip_fg =''C'')'+@CR+
                   '   and chg_sp_pdate_year is not null'+@CR+
                   -- 2017/08/08 Rickliu 副總指示客編 IT、ZZ 開頭不認列業績，雖不認列業績但應收扔得收
                   '   and Chg_BU_NO Not In (''IT0000'', ''ZZ0000'') '+@CR+
                   ' group by chg_sp_pdate_year, chg_sp_pdate_month, sd_ctno, chg_ct_sname '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
    
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '01223'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月客戶大分類退貨金額-3C'
     set @Remark= '(依請款日, 表身, 含稅, 無加減項, 無折讓, 大分類為 AB)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_pdate_year as area_year,'+@CR+
                   '       chg_sp_pdate_month as area_month,'+@CR+
                   '       IsNull(sd_ctno, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(chg_ct_sname, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(chg_sd_stot_tax, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslpdt'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and chg_sp_pdate_year >= ''2013'''+@CR+
                   '   and chg_skno_bkind = ''AB'''+@CR+
                   '   and sd_slip_fg =''3'''+@CR+
                   '   and chg_sp_pdate_year is not null'+@CR+
                   -- 2017/08/08 Rickliu 副總指示客編 IT、ZZ 開頭不認列業績，雖不認列業績但應收扔得收
                   '   and Chg_BU_NO Not In (''IT0000'', ''ZZ0000'') '+@CR+
                   ' group by chg_sp_pdate_year, chg_sp_pdate_month, sd_ctno, chg_ct_sname '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '01224'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月客戶大分類實銷金額達成率-3C'
     set @Remark= '(依請款日)(Formula: 01201 / 3C目標, 目標資料: ori_xls#Target_Stock_BKind_Personal.Target_AB)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       m.area_year,'+@CR+
                   '       m.area_month,'+@CR+
                   '       IsNull(m.area, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(m.area_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       case'+@CR+
                   '         when IsNull(d.amt, 0) = 0 then 0'+@CR+
                   '         else Round(isnull(m.amt, 0) / isnull(d.amt, 0), 4)'+@CR+
                   '       end as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       3 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '       inner join '+@TB_RT_Data_tmp+' d'+@TB_Hint+@CR+
                   '          on m.area_year = d.area_year'+@CR+
                   '         and m.area_month = d.area_month'+@CR+
                   '         and m.area '+@Collate+'=d.area '+@Collate+@CR+
                   ' where 1=1'+@CR+
                   '   and m.area_year >= ''2013'''+@CR+
                   '   and m.kind =''01221'''+@CR+
                   '   and d.kind =''01220'' '
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '01225'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '年度業務大分類實銷金額達成率-3C'
     set @Remark= '(依請款日)(Formula: 01201 / 3C目標, 目標資料: ori_xls#Target_Stock_BKind_Personal.Target_AB)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       m.area_year,'+@CR+
                   '       ''00'' as area_month,'+@CR+
                   '       IsNull(m.area, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(m.area_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       case'+@CR+
                   '         when IsNull(d.Target_Amt, 0) = 0 then 0'+@CR+
                   '         else Round(Sum(Isnull(m.amt, 0)) / d.Target_Amt, 4)'+@CR+
                   '       end as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       3 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '       inner join'+@CR+
                   '         (select area_year, area, Sum(Isnull(amt, 0)) as Target_Amt'+@CR+
                   '            from '+@TB_RT_Data_tmp+' d'+@TB_Hint+@CR+
                   '           where kind=''01220'''+@CR+
                   '           group by area_year, area'+@CR+
                   '         ) D'+@CR+
                   '          on m.area_year=d.area_year'+@CR+
                   '         and m.area '+@Collate+'=d.area '+@Collate+@CR+
                   ' where 1=1'+@CR+
                   '   and m.area_year >= ''2013'''+@CR+
                   '   and m.kind =''01221'''+@CR+
                   ' group by m.area_year, m.area, area_name, Target_Amt'

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end 
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '01226'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月客戶大分類實銷金額-3C(客戶業績毛利)'
     set @Remark= '(依請款日, 表身, 未稅, 無加減項, 無折讓, 含退貨, 含服務單, 大分類為 AB)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_pdate_year as area_year,'+@CR+
                   '       chg_sp_pdate_month as area_month,'+@CR+
                   '       IsNull(sd_ctno, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(chg_ct_sname, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(Chg_Profit, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslpdt'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and chg_sp_pdate_year >= ''2013'''+@CR+
                   '   and chg_skno_bkind = ''AB'''+@CR+
                   '   and (sd_slip_fg =''2'' Or sd_slip_fg =''C'' Or sd_slip_fg = ''3'')'+@CR+
                   '   and chg_sp_pdate_year is not null'+@CR+
                   -- 2017/08/08 Rickliu 副總指示客編 IT、ZZ 開頭不認列業績，雖不認列業績但應收扔得收
                   '   and Chg_BU_NO Not In (''IT0000'', ''ZZ0000'') '+@CR+
                   ' group by chg_sp_pdate_year, chg_sp_pdate_month, sd_ctno, chg_ct_sname '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '01300'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月業務大分類目標-化工(個人業績)'
     set @Remark= '(大分類為 AC)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       Year as area_year,'+@CR+
                   '       Month as area_month,'+@CR+
                   '       IsNull(ct_sales, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(sales_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(amt, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_Target+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and Year >= ''2013'''+@CR+
                   '   and BKind =''AC'''+@CR+
                   ' group by Year, Month, ct_sales, sales_name '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '01301'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月業務大分類實銷金額-化工(個人業績)'
     set @Remark= '(依請款日, 表身, 含稅, 無加減項, 無折讓, 含退貨, 含服務單, 大分類為 AC)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_pdate_year as area_year,'+@CR+
                   '       chg_sp_pdate_month as area_month,'+@CR+
                   '       IsNull(sp_sales, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(chg_sales_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(chg_sd_stot_tax, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslpdt'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and chg_sp_pdate_year >= ''2013'''+@CR+
                   '   and chg_skno_bkind = ''AC'''+@CR+
                   '   and (sd_slip_fg =''2'' Or sd_slip_fg =''C'' Or sd_slip_fg = ''3'')'+@CR+
                   '   and chg_sp_pdate_year is not null'+@CR+
                   -- 2017/08/08 Rickliu 副總指示客編 IT、ZZ 開頭不認列業績，雖不認列業績但應收扔得收
                   '   and Chg_BU_NO Not In (''IT0000'', ''ZZ0000'') '+@CR+
                   ' group by chg_sp_pdate_year, chg_sp_pdate_month, sp_sales, chg_sales_name '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '01302'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月業務大分類銷售金額-化工(個人業績)'
     set @Remark= '(依請款日, 表身, 含稅, 無加減項, 無折讓, 含服務單, 大分類為 AC)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_pdate_year as area_year,'+@CR+
                   '       chg_sp_pdate_month as area_month,'+@CR+
                   '       IsNull(sp_sales, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(chg_sales_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(chg_sd_stot_tax, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslpdt'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and chg_sp_pdate_year >= ''2013'''+@CR+
                   '   and chg_skno_bkind = ''AC'''+@CR+
                   '   and (sd_slip_fg =''2'' Or sd_slip_fg =''C'')'+@CR+
                   '   and chg_sp_pdate_year is not null'+@CR+
                   -- 2017/08/08 Rickliu 副總指示客編 IT、ZZ 開頭不認列業績，雖不認列業績但應收扔得收
                   '   and Chg_BU_NO Not In (''IT0000'', ''ZZ0000'') '+@CR+
                   ' group by chg_sp_pdate_year, chg_sp_pdate_month, sp_sales, chg_sales_name '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
    
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '01303'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月業務大分類退貨金額-化工(個人業績)'
     set @Remark= '(依請款日, 表身, 含稅, 無加減項, 無折讓, 大分類為 AC)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_pdate_year as area_year,'+@CR+
                   '       chg_sp_pdate_month as area_month,'+@CR+
                   '       IsNull(sp_sales, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(chg_sales_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(chg_sd_stot_tax, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslpdt'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and chg_sp_pdate_year >= ''2013'''+@CR+
                   '   and chg_skno_bkind = ''AC'''+@CR+
                   '   and sd_slip_fg =''3'''+@CR+
                   '   and chg_sp_pdate_year is not null'+@CR+
                   -- 2017/08/08 Rickliu 副總指示客編 IT、ZZ 開頭不認列業績，雖不認列業績但應收扔得收
                   '   and Chg_BU_NO Not In (''IT0000'', ''ZZ0000'') '+@CR+
                   ' group by chg_sp_pdate_year, chg_sp_pdate_month, sp_sales, chg_sales_name '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '01304'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月業務大分類實銷金額達成率-化工(個人業績)'
     set @Remark= '(依請款日)(Formula: 01301 / 化工目標, 目標資料: ori_xls#Target_Stock_BKind_Personal.Target_AC)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       m.area_year,'+@CR+
                   '       m.area_month,'+@CR+
                   '       IsNull(m.area, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(m.area_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       case'+@CR+
                   '         when IsNull(d.amt, 0) = 0 then 0'+@CR+
                   '         else Round(isnull(m.amt, 0) / isnull(d.amt, 0), 4)'+@CR+
                   '       end as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       3 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '       inner join '+@TB_RT_Data_tmp+' d'+@TB_Hint+@CR+
                   '          on m.area_year = d.area_year'+@CR+
                   '         and m.area_month = d.area_month'+@CR+
                   '         and m.area '+@Collate+'=d.area '+@Collate+@CR+
                   ' where 1=1'+@CR+
                   '   and m.area_year >= ''2013'''+@CR+
                   '   and m.kind =''01301'''+@CR+
                   '   and d.kind =''01300'' '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '01305'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '年度業務大分類實銷金額達成率-化工(個人業績)'
     set @Remark= '(依請款日)(Formula: 01301 / 化工目標, 目標資料: ori_xls#Target_Stock_BKind_Personal.AC)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       m.area_year,'+@CR+
                   '       ''00'' as area_month,'+@CR+
                   '       IsNull(m.area, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(m.area_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       case'+@CR+
                   '         when IsNull(d.Target_Amt, 0) = 0 then 0'+@CR+
                   '         else Round(Sum(Isnull(m.amt, 0)) / d.Target_Amt, 4)'+@CR+
                   '       end as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       3 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '       inner join'+@CR+
                   '         (select area_year, area, Sum(Isnull(amt, 0)) as Target_Amt'+@CR+
                   '            from '+@TB_RT_Data_tmp+' d'+@TB_Hint+@CR+
                   '           where kind=''01300'''+@CR+
                   '           group by area_year, area'+@CR+
                   '         ) D'+@CR+
                   '          on m.area_year=d.area_year'+@CR+
                   '         and m.area '+@Collate+'=d.area '+@Collate+@CR+
                   ' where 1=1'+@CR+
                   '   and m.area_year >= ''2013'''+@CR+
                   '   and m.kind =''01301'''+@CR+
                   ' group by m.area_year, m.area, area_name, Target_Amt'

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end 

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '01306'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月業務大分類實銷金額-化工(個人業績毛利)'
     set @Remark= '(依請款日, 表身, 未稅, 無加減項, 無折讓, 含退貨, 含服務單, 大分類為 AC)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_pdate_year as area_year,'+@CR+
                   '       chg_sp_pdate_month as area_month,'+@CR+
                   '       IsNull(sp_sales, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(chg_sales_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(Chg_Profit, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslpdt'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and chg_sp_pdate_year >= ''2013'''+@CR+
                   '   and chg_skno_bkind = ''AC'''+@CR+
                   '   and (sd_slip_fg =''2'' Or sd_slip_fg =''C'' Or sd_slip_fg = ''3'')'+@CR+
                   '   and chg_sp_pdate_year is not null'+@CR+
                   -- 2017/08/08 Rickliu 副總指示客編 IT、ZZ 開頭不認列業績，雖不認列業績但應收扔得收
                   '   and Chg_BU_NO Not In (''IT0000'', ''ZZ0000'') '+@CR+
                   ' group by chg_sp_pdate_year, chg_sp_pdate_month, sp_sales, chg_sales_name '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
    
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '01320'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月客戶大分類目標-化工'
     set @Remark= '(大分類為 AC)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       Year as area_year,'+@CR+
                   '       Month as area_month,'+@CR+
                   '       IsNull(ct_no, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(ct_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(amt, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_Target+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and Year >= ''2013'''+@CR+
                   '   and BKind =''AC'''+@CR+
                   ' group by Year, Month, ct_no, ct_name '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '01321'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月客戶大分類實銷金額-化工'
     set @Remark= '(依請款日, 表身, 含稅, 無加減項, 扣折讓, 含退貨, 含服務單, 大分類為 AC)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_pdate_year as area_year,'+@CR+
                   '       chg_sp_pdate_month as area_month,'+@CR+
                   '       IsNull(sd_ctno, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(chg_ct_sname, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(chg_sd_stot_tax, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslpdt'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and chg_sp_pdate_year >= ''2013'''+@CR+
                   '   and chg_skno_bkind = ''AC'''+@CR+
                   '   and (sd_slip_fg =''2'' Or sd_slip_fg =''C'' Or sd_slip_fg = ''3'')'+@CR+
                   '   and chg_sp_pdate_year is not null'+@CR+
                   -- 2017/08/08 Rickliu 副總指示客編 IT、ZZ 開頭不認列業績，雖不認列業績但應收扔得收
                   '   and Chg_BU_NO Not In (''IT0000'', ''ZZ0000'') '+@CR+
                   ' group by chg_sp_pdate_year, chg_sp_pdate_month, sd_ctno, chg_ct_sname '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '01322'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月客戶大分類銷售金額-化工'
     set @Remark='(依請款日, 表身, 含稅, 無加減項, 無折讓, 含服務單, 大分類為 AC)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_pdate_year as area_year,'+@CR+
                   '       chg_sp_pdate_month as area_month,'+@CR+
                   '       IsNull(sd_ctno, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(chg_ct_sname, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(chg_sd_stot_tax, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslpdt'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and chg_sp_pdate_year >= ''2013'''+@CR+
                   '   and chg_skno_bkind = ''AC'''+@CR+
                   '   and (sd_slip_fg =''2'' Or sd_slip_fg =''C'')'+@CR+
                   '   and chg_sp_pdate_year is not null'+@CR+
                   -- 2017/08/08 Rickliu 副總指示客編 IT、ZZ 開頭不認列業績，雖不認列業績但應收扔得收
                   '   and Chg_BU_NO Not In (''IT0000'', ''ZZ0000'') '+@CR+
                   ' group by chg_sp_pdate_year, chg_sp_pdate_month, sd_ctno, chg_ct_sname '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
    
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '01323'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月客戶大分類退貨金額-化工'
     set @Remark= '(依請款日, 表身, 含稅, 無加減項, 無折讓, 大分類為 AC)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_pdate_year as area_year,'+@CR+
                   '       chg_sp_pdate_month as area_month,'+@CR+
                   '       IsNull(sd_ctno, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(chg_ct_sname, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(chg_sd_stot_tax, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslpdt'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and chg_sp_pdate_year >= ''2013'''+@CR+
                   '   and chg_skno_bkind = ''AC'''+@CR+
                   '   and sd_slip_fg =''3'''+@CR+
                   '   and chg_sp_pdate_year is not null'+@CR+
                   -- 2017/08/08 Rickliu 副總指示客編 IT、ZZ 開頭不認列業績，雖不認列業績但應收扔得收
                   '   and Chg_BU_NO Not In (''IT0000'', ''ZZ0000'') '+@CR+
                   ' group by chg_sp_pdate_year, chg_sp_pdate_month, sd_ctno, chg_ct_sname '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '01324'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月客戶大分類實銷金額達成率-化工'
     set @Remark= '(依請款日)(Formula: 01301 / 化工目標, 目標資料: ori_xls#Target_Stock_BKind_Personal.Target_AC)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       m.area_year,'+@CR+
                   '       m.area_month,'+@CR+
                   '       IsNull(m.area, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(m.area_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       case'+@CR+
                   '         when IsNull(d.amt, 0) = 0 then 0'+@CR+
                   '         else Round(isnull(m.amt, 0) / isnull(d.amt, 0), 4)'+@CR+
                   '       end as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       3 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '       inner join '+@TB_RT_Data_tmp+' d'+@TB_Hint+@CR+
                   '          on m.area_year = d.area_year'+@CR+
                   '         and m.area_month = d.area_month'+@CR+
                   '         and m.area '+@Collate+'=d.area '+@Collate+@CR+
                   ' where 1=1'+@CR+
                   '   and m.area_year >= ''2013'''+@CR+
                   '   and m.kind =''01321'''+@CR+
                   '   and d.kind =''01320'' '
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '01325'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '年度業務大分類實銷金額達成率-化工'
     set @Remark= '(依請款日)(Formula: 01301 / 化工目標, 目標資料: ori_xls#Target_Stock_BKind_Personal.Target_AC)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       m.area_year,'+@CR+
                   '       ''00'' as area_month,'+@CR+
                   '       IsNull(m.area, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(m.area_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       case'+@CR+
                   '         when IsNull(d.Target_Amt, 0) = 0 then 0'+@CR+
                   '         else Round(Sum(Isnull(m.amt, 0)) / d.Target_Amt, 4)'+@CR+
                   '       end as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       3 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '       inner join'+@CR+
                   '         (select area_year, area, Sum(Isnull(amt, 0)) as Target_Amt'+@CR+
                   '            from '+@TB_RT_Data_tmp+' d'+@TB_Hint+@CR+
                   '           where kind=''01320'''+@CR+
                   '           group by area_year, area'+@CR+
                   '         ) D'+@CR+
                   '          on m.area_year=d.area_year'+@CR+
                   '         and m.area '+@Collate+'=d.area '+@Collate+@CR+
                   ' where 1=1'+@CR+
                   '   and m.area_year >= ''2013'''+@CR+
                   '   and m.kind =''01321'''+@CR+
                   ' group by m.area_year, m.area, area_name, Target_Amt'

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end 
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '01326'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月客戶大分類實銷金額-化工(客戶業績毛利)'
     set @Remark= '(依請款日, 表身, 未稅, 無加減項, 無折讓, 含退貨, 含服務單, 大分類為 AC)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_pdate_year as area_year,'+@CR+
                   '       chg_sp_pdate_month as area_month,'+@CR+
                   '       IsNull(sd_ctno, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(chg_ct_sname, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(Chg_Profit, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslpdt'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and chg_sp_pdate_year >= ''2013'''+@CR+
                   '   and chg_skno_bkind = ''AC'''+@CR+
                   '   and (sd_slip_fg =''2'' Or sd_slip_fg =''C'' Or sd_slip_fg = ''3'')'+@CR+
                   '   and chg_sp_pdate_year is not null'+@CR+
                   -- 2017/08/08 Rickliu 副總指示客編 IT、ZZ 開頭不認列業績，雖不認列業績但應收扔得收
                   '   and Chg_BU_NO Not In (''IT0000'', ''ZZ0000'') '+@CR+
                   ' group by chg_sp_pdate_year, chg_sp_pdate_month, sd_ctno, chg_ct_sname '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '01400'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月業務大分類目標-工具(個人業績)'
     set @Remark= '(大分類為 AD)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       Year as area_year,'+@CR+
                   '       Month as area_month,'+@CR+
                   '       IsNull(ct_sales, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(sales_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(amt, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_Target+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and Year >= ''2013'''+@CR+
                   '   and BKind =''AD'''+@CR+
                   ' group by Year, Month, ct_sales, sales_name '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '01401'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月業務大分類實銷金額-工具(個人業績)'
     set @Remark= '(依請款日, 表身, 含稅, 無加減項, 無折讓, 含退貨, 含服務單, 大分類為 AD)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_pdate_year as area_year,'+@CR+
                   '       chg_sp_pdate_month as area_month,'+@CR+
                   '       IsNull(sp_sales, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(chg_sales_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(chg_sd_stot_tax, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslpdt'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and chg_sp_pdate_year >= ''2013'''+@CR+
                   '   and chg_skno_bkind = ''AD'''+@CR+
                   '   and (sd_slip_fg =''2'' Or sd_slip_fg =''C'' Or sd_slip_fg = ''3'')'+@CR+
                   '   and chg_sp_pdate_year is not null'+@CR+
                   -- 2017/08/08 Rickliu 副總指示客編 IT、ZZ 開頭不認列業績，雖不認列業績但應收扔得收
                   '   and Chg_BU_NO Not In (''IT0000'', ''ZZ0000'') '+@CR+
                   ' group by chg_sp_pdate_year, chg_sp_pdate_month, sp_sales, chg_sales_name '
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '01402'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月業務大分類銷售金額-工具(個人業績)'
     set @Remark= '(依請款日, 表身, 含稅, 無加減項, 無折讓, 含服務單, 大分類為 AD)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_pdate_year as area_year,'+@CR+
                   '       chg_sp_pdate_month as area_month,'+@CR+
                   '       IsNull(sp_sales, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(chg_sales_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(chg_sd_stot_tax, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslpdt'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and chg_sp_pdate_year >= ''2013'''+@CR+
                   '   and chg_skno_bkind = ''AD'''+@CR+
                   '   and (sd_slip_fg =''2'' Or sd_slip_fg =''C'')'+@CR+
                   '   and chg_sp_pdate_year is not null'+@CR+
                   -- 2017/08/08 Rickliu 副總指示客編 IT、ZZ 開頭不認列業績，雖不認列業績但應收扔得收
                   '   and Chg_BU_NO Not In (''IT0000'', ''ZZ0000'') '+@CR+
                   ' group by chg_sp_pdate_year, chg_sp_pdate_month, sp_sales, chg_sales_name '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
    
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '01403'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月業務大分類退貨金額-工具(個人業績)'
     set @Remark= '(依請款日, 表身, 含稅, 無加減項, 無折讓, 大分類為 AD)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_pdate_year as area_year,'+@CR+
                   '       chg_sp_pdate_month as area_month,'+@CR+
                   '       IsNull(sp_sales, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(chg_sales_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(chg_sd_stot_tax, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslpdt'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and chg_sp_pdate_year >= ''2013'''+@CR+
                   '   and chg_skno_bkind = ''AD'''+@CR+
                   '   and sd_slip_fg =''3'''+@CR+
                   '   and chg_sp_pdate_year is not null'+@CR+
                   -- 2017/08/08 Rickliu 副總指示客編 IT、ZZ 開頭不認列業績，雖不認列業績但應收扔得收
                   '   and Chg_BU_NO Not In (''IT0000'', ''ZZ0000'') '+@CR+
                   ' group by chg_sp_pdate_year, chg_sp_pdate_month, sp_sales, chg_sales_name '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '01404'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月業務大分類實銷金額達成率-工具(個人業績)'
     set @Remark= '(依請款日)(Formula: 01401 / 工具目標, 目標資料: ori_xls#Target_Stock_BKind_Personal.Target_AD'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       m.area_year,'+@CR+
                   '       m.area_month,'+@CR+
                   '       IsNull(m.area, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(m.area_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       case'+@CR+
                   '         when IsNull(d.amt, 0) = 0 then 0'+@CR+
                   '         else Round(isnull(m.amt, 0) / isnull(d.amt, 0), 4)'+@CR+
                   '       end as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       3 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '       inner join '+@TB_RT_Data_tmp+' d'+@TB_Hint+@CR+
                   '          on m.area_year = d.area_year'+@CR+
                   '         and m.area_month = d.area_month'+@CR+
                   '         and m.area '+@Collate+'=d.area '+@Collate+@CR+
                   ' where 1=1'+@CR+
                   '   and m.area_year >= ''2013'''+@CR+
                   '   and m.kind =''01401'''+@CR+
                   '   and d.kind =''01400'' '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '01405'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '年度業務大分類實銷金額達成率-工具(個人業績)'
     set @Remark= '(依請款日)(Formula: 01401 / 工具目標, 目標資料: ori_xls#Target_Stock_BKind_Personal.Target_AD)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       m.area_year,'+@CR+
                   '       ''00'' as area_month,'+@CR+
                   '       IsNull(m.area, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(m.area_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       case'+@CR+
                   '         when IsNull(d.Target_Amt, 0) = 0 then 0'+@CR+
                   '         else Round(Sum(Isnull(m.amt, 0)) / d.Target_Amt, 4)'+@CR+
                   '       end as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       3 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '       inner join'+@CR+
                   '         (select area_year, area, Sum(Isnull(amt, 0)) as Target_Amt'+@CR+
                   '            from '+@TB_RT_Data_tmp+' d'+@TB_Hint+@CR+
                   '           where kind=''01400'''+@CR+
                   '           group by area_year, area'+@CR+
                   '         ) D'+@CR+
                   '          on m.area_year=d.area_year'+@CR+
                   '         and m.area '+@Collate+'=d.area '+@Collate+@CR+
                   ' where 1=1'+@CR+
                   '   and m.area_year >= ''2013'''+@CR+
                   '   and m.kind =''01401'''+@CR+
                   ' group by m.area_year, m.area, area_name, Target_Amt'

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end 
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '01406'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月業務大分類實銷金額-工具(個人業績毛利)'
     set @Remark= '(依請款日, 表身, 未稅, 無加減項, 無折讓, 含退貨, 含服務單, 大分類為 AD)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_pdate_year as area_year,'+@CR+
                   '       chg_sp_pdate_month as area_month,'+@CR+
                   '       IsNull(sp_sales, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(chg_sales_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(Chg_Profit, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslpdt'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and chg_sp_pdate_year >= ''2013'''+@CR+
                   '   and chg_skno_bkind = ''AD'''+@CR+
                   '   and (sd_slip_fg =''2'' Or sd_slip_fg =''C'' Or sd_slip_fg = ''3'')'+@CR+           
                   '   and chg_sp_pdate_year is not null'+@CR+
                   -- 2017/08/08 Rickliu 副總指示客編 IT、ZZ 開頭不認列業績，雖不認列業績但應收扔得收
                   '   and Chg_BU_NO Not In (''IT0000'', ''ZZ0000'') '+@CR+
                   ' group by chg_sp_pdate_year, chg_sp_pdate_month, sp_sales, chg_sales_name '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
    
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '01420'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月客戶大分類目標-工具'
     set @Remark= '(大分類為 AD)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       Year as area_year,'+@CR+
                   '       Month as area_month,'+@CR+
                   '       IsNull(ct_no, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(ct_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(amt, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_Target+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and Year >= ''2014'''+@CR+
                   '   and BKind =''AD'''+@CR+
                   ' group by Year, Month, ct_no, ct_name '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '01421'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月客戶大分類實銷金額-工具'
     set @Remark= '(依請款日, 表身, 含稅, 無加減項, 扣折讓, 含退貨, 含服務單, 大分類為 AD)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_pdate_year as area_year,'+@CR+
                   '       chg_sp_pdate_month as area_month,'+@CR+
                   '       IsNull(sd_ctno, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(chg_ct_sname, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(chg_sd_stot_tax, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslpdt'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and chg_sp_pdate_year >= ''2013'''+@CR+
                   '   and chg_skno_bkind = ''AD'''+@CR+
                   '   and (sd_slip_fg =''2'' Or sd_slip_fg =''C'' Or sd_slip_fg = ''3'')'+@CR+
                   '   and chg_sp_pdate_year is not null'+@CR+
                   -- 2017/08/08 Rickliu 副總指示客編 IT、ZZ 開頭不認列業績，雖不認列業績但應收扔得收
                   '   and Chg_BU_NO Not In (''IT0000'', ''ZZ0000'') '+@CR+
                   ' group by chg_sp_pdate_year, chg_sp_pdate_month, sd_ctno, chg_ct_sname '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '01422'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月客戶大分類銷售金額-工具'
     set @Remark='(依請款日, 表身, 含稅, 無加減項, 無折讓, 含服務單, 大分類為 AD)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_pdate_year as area_year,'+@CR+
                   '       chg_sp_pdate_month as area_month,'+@CR+
                   '       IsNull(sd_ctno, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(chg_ct_sname, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(chg_sd_stot_tax, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslpdt'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and chg_sp_pdate_year >= ''2014'''+@CR+
                   '   and chg_skno_bkind = ''AD'''+@CR+
                   '   and (sd_slip_fg =''2'' Or sd_slip_fg =''C'')'+@CR+
                   '   and chg_sp_pdate_year is not null'+@CR+
                   -- 2017/08/08 Rickliu 副總指示客編 IT、ZZ 開頭不認列業績，雖不認列業績但應收扔得收
                   '   and Chg_BU_NO Not In (''IT0000'', ''ZZ0000'') '+@CR+
                   ' group by chg_sp_pdate_year, chg_sp_pdate_month, sd_ctno, chg_ct_sname '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
    
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '01423'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月客戶大分類退貨金額-工具'
     set @Remark= '(依請款日, 表身, 含稅, 無加減項, 無折讓, 大分類為 AD)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_pdate_year as area_year,'+@CR+
                   '       chg_sp_pdate_month as area_month,'+@CR+
                   '       IsNull(sd_ctno, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(chg_ct_sname, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(chg_sd_stot_tax, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslpdt'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and chg_sp_pdate_year >= ''2014'''+@CR+
                   '   and chg_skno_bkind = ''AD'''+@CR+
                   '   and sd_slip_fg =''3'''+@CR+
                   '   and chg_sp_pdate_year is not null'+@CR+
                   -- 2017/08/08 Rickliu 副總指示客編 IT、ZZ 開頭不認列業績，雖不認列業績但應收扔得收
                   '   and Chg_BU_NO Not In (''IT0000'', ''ZZ0000'') '+@CR+
                   ' group by chg_sp_pdate_year, chg_sp_pdate_month, sd_ctno, chg_ct_sname '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '01424'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月客戶大分類實銷金額達成率-工具'
     set @Remark= '(依請款日)(Formula: 01401 / 工具目標, 目標資料: ori_xls#Target_Stock_BKind_Personal.Target_AD)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       m.area_year,'+@CR+
                   '       m.area_month,'+@CR+
                   '       IsNull(m.area, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(m.area_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       case'+@CR+
                   '         when IsNull(d.amt, 0) = 0 then 0'+@CR+
                   '         else Round(isnull(m.amt, 0) / isnull(d.amt, 0), 4)'+@CR+
                   '       end as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       3 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '       inner join '+@TB_RT_Data_tmp+' d'+@TB_Hint+@CR+
                   '          on m.area_year = d.area_year'+@CR+
                   '         and m.area_month = d.area_month'+@CR+
                   '         and m.area '+@Collate+'=d.area '+@Collate+@CR+
                   ' where 1=1'+@CR+
                   '   and m.area_year >= ''2014'''+@CR+
                   '   and m.kind =''01421'''+@CR+
                   '   and d.kind =''01420'' '
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '01425'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '年度業務大分類實銷金額達成率-工具'
     set @Remark= '(依請款日)(Formula: 01401 / 工具目標, 目標資料: ori_xls#Target_Stock_BKind_Personal.Target_AD)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       m.area_year,'+@CR+
                   '       ''00'' as area_month,'+@CR+
                   '       IsNull(m.area, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(m.area_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       case'+@CR+
                   '         when IsNull(d.Target_Amt, 0) = 0 then 0'+@CR+
                   '         else Round(Sum(Isnull(m.amt, 0)) / d.Target_Amt, 4)'+@CR+
                   '       end as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       3 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '       inner join'+@CR+
                   '         (select area_year, area, Sum(Isnull(amt, 0)) as Target_Amt'+@CR+
                   '            from '+@TB_RT_Data_tmp+' d'+@TB_Hint+@CR+
                   '           where kind=''01420'''+@CR+
                   '           group by area_year, area'+@CR+
                   '         ) D'+@CR+
                   '          on m.area_year=d.area_year'+@CR+
                   '         and m.area '+@Collate+'=d.area '+@Collate+@CR+
                   ' where 1=1'+@CR+
                   '   and m.area_year >= ''2014'''+@CR+
                   '   and m.kind =''01421'''+@CR+
                   ' group by m.area_year, m.area, area_name, Target_Amt'

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end 
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '01426'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月客戶大分類實銷金額-工具(客戶業績毛利)'
     set @Remark= '(依請款日, 表身, 未稅, 無加減項, 無折讓, 含退貨, 含服務單, 大分類為 AD)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_pdate_year as area_year,'+@CR+
                   '       chg_sp_pdate_month as area_month,'+@CR+
                   '       IsNull(sd_ctno, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(chg_ct_sname, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(Chg_Profit, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslpdt'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and chg_sp_pdate_year >= ''2014'''+@CR+
                   '   and chg_skno_bkind = ''AD'''+@CR+
                   '   and (sd_slip_fg =''2'' Or sd_slip_fg =''C'' Or sd_slip_fg = ''3'')'+@CR+
                   '   and chg_sp_pdate_year is not null'+@CR+
                   -- 2017/08/08 Rickliu 副總指示客編 IT、ZZ 開頭不認列業績，雖不認列業績但應收扔得收
                   '   and Chg_BU_NO Not In (''IT0000'', ''ZZ0000'') '+@CR+
                   ' group by chg_sp_pdate_year, chg_sp_pdate_month, sd_ctno, chg_ct_sname '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '01500'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月業務大分類目標-電裝(個人業績)'
     set @Remark= '(大分類為 AE)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       Year as area_year,'+@CR+
                   '       Month as area_month,'+@CR+
                   '       IsNull(ct_sales, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(sales_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(amt, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_Target+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and Year >= ''2013'''+@CR+
                   '   and BKind =''AE'''+@CR+
                   ' group by Year, Month, ct_sales, sales_name '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '01501'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月業務大分類實銷金額-電裝(個人業績)'
     set @Remark= '(依請款日, 表身, 含稅, 無加減項, 無折讓, 含退貨, 含服務單, 大分類為 AE)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_pdate_year as area_year,'+@CR+
                   '       chg_sp_pdate_month as area_month,'+@CR+
                   '       IsNull(sp_sales, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(chg_sales_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(chg_sd_stot_tax, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslpdt'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and chg_sp_pdate_year >= ''2013'''+@CR+
                   '   and chg_skno_bkind = ''AE'''+@CR+
                   '   and (sd_slip_fg =''2'' Or sd_slip_fg =''C'' Or sd_slip_fg = ''3'')'+@CR+
                   '   and chg_sp_pdate_year is not null'+@CR+
                   -- 2017/08/08 Rickliu 副總指示客編 IT、ZZ 開頭不認列業績，雖不認列業績但應收扔得收
                   '   and Chg_BU_NO Not In (''IT0000'', ''ZZ0000'') '+@CR+
                   ' group by chg_sp_pdate_year, chg_sp_pdate_month, sp_sales, chg_sales_name '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '01502'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月業務大分類銷售金額-電裝(個人業績)'
     set @Remark= '(依請款日, 表身, 含稅, 無加減項, 無折讓, 含服務單, 大分類為 AE)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_pdate_year as area_year,'+@CR+
                   '       chg_sp_pdate_month as area_month,'+@CR+
                   '       IsNull(sp_sales, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(chg_sales_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(chg_sd_stot_tax, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslpdt'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and chg_sp_pdate_year >= ''2013'''+@CR+
                   '   and chg_skno_bkind = ''AE'''+@CR+
                   '   and (sd_slip_fg =''2'' Or sd_slip_fg =''C'')'+@CR+
                   '   and chg_sp_pdate_year is not null'+@CR+
                   -- 2017/08/08 Rickliu 副總指示客編 IT、ZZ 開頭不認列業績，雖不認列業績但應收扔得收
                   '   and Chg_BU_NO Not In (''IT0000'', ''ZZ0000'') '+@CR+
                   ' group by chg_sp_pdate_year, chg_sp_pdate_month, sp_sales, chg_sales_name '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
    
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '01503'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月業務大分類退貨金額-電裝(個人業績)'
     set @Remark= '(依請款日, 表身, 含稅, 無加減項, 無折讓, 大分類為 AE)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_pdate_year as area_year,'+@CR+
                   '       chg_sp_pdate_month as area_month,'+@CR+
                   '       IsNull(sp_sales, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(chg_sales_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(chg_sd_stot_tax, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslpdt'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and chg_sp_pdate_year >= ''2013'''+@CR+
                   '   and chg_skno_bkind = ''AE'''+@CR+
                   '   and sd_slip_fg =''3'''+@CR+
                   '   and chg_sp_pdate_year is not null'+@CR+
                   -- 2017/08/08 Rickliu 副總指示客編 IT、ZZ 開頭不認列業績，雖不認列業績但應收扔得收
                   '   and Chg_BU_NO Not In (''IT0000'', ''ZZ0000'') '+@CR+
                   ' group by chg_sp_pdate_year, chg_sp_pdate_month, sp_sales, chg_sales_name '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '01504'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月業務大分類實銷金額達成率-電裝(個人業績)'
     set @Remark= '(依請款日)(Formula: 01501 / 工具目標, 目標資料: ori_xls#Target_Stock_BKind_Personal.Target_AE)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       m.area_year,'+@CR+
                   '       m.area_month,'+@CR+
                   '       IsNull(m.area, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(m.area_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       case'+@CR+
                   '         when IsNull(d.amt, 0) = 0 then 0'+@CR+
                   '         else Round(isnull(m.amt, 0) / isnull(d.amt, 0), 4)'+@CR+
                   '       end as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       3 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '       inner join '+@TB_RT_Data_tmp+' d'+@TB_Hint+@CR+
                   '          on m.area_year = d.area_year'+@CR+
                   '         and m.area_month = d.area_month'+@CR+
                   '         and m.area '+@Collate+'=d.area '+@Collate+@CR+
                   ' where 1=1'+@CR+
                   '   and m.area_year >= ''2013'''+@CR+
                   '   and m.kind =''01501'''+@CR+
                   '   and d.kind =''01500'' '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '01505'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '年度業務大分類實銷金額達成率-電裝(個人業績)'
     set @Remark= '(依請款日)(Formula: 01501 / 工具目標, 目標資料: ori_xls#Target_Stock_BKind_Personal.Target_AE)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       m.area_year,'+@CR+
                   '       ''00'' as area_month,'+@CR+
                   '       IsNull(m.area, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(m.area_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       case'+@CR+
                   '         when IsNull(d.Target_Amt, 0) = 0 then 0'+@CR+
                   '         else Round(Sum(Isnull(m.amt, 0)) / d.Target_Amt, 4)'+@CR+
                   '       end as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       3 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '       inner join'+@CR+
                   '         (select area_year, area, Sum(Isnull(amt, 0)) as Target_Amt'+@CR+
                   '            from '+@TB_RT_Data_tmp+' d'+@TB_Hint+@CR+
                   '           where kind=''01500'''+@CR+
                   '           group by area_year, area'+@CR+
                   '         ) D'+@CR+
                   '          on m.area_year=d.area_year'+@CR+
                   '         and m.area '+@Collate+'=d.area '+@Collate+@CR+
                   ' where 1=1'+@CR+
                   '   and m.area_year >= ''2013'''+@CR+
                   '   and m.kind =''01501'''+@CR+
                   ' group by m.area_year, m.area, area_name, Target_Amt'

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end 
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '01506'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月業務大分類實銷金額-電裝(個人業績毛利)'
     set @Remark= '(依請款日, 表身, 未稅, 無加減項, 無折讓, 含退貨, 含服務單, 大分類為 AE)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_pdate_year as area_year,'+@CR+
                   '       chg_sp_pdate_month as area_month,'+@CR+
                   '       IsNull(sp_sales, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(chg_sales_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(Chg_Profit, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslpdt'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and chg_sp_pdate_year >= ''2013'''+@CR+
                   '   and chg_skno_bkind = ''AE'''+@CR+
                   '   and (sd_slip_fg =''2'' Or sd_slip_fg =''C'' Or sd_slip_fg = ''3'')'+@CR+
                   '   and chg_sp_pdate_year is not null'+@CR+
                   -- 2017/08/08 Rickliu 副總指示客編 IT、ZZ 開頭不認列業績，雖不認列業績但應收扔得收
                   '   and Chg_BU_NO Not In (''IT0000'', ''ZZ0000'') '+@CR+
                   ' group by chg_sp_pdate_year, chg_sp_pdate_month, sp_sales, chg_sales_name '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '01520'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月客戶大分類目標-電裝'
     set @Remark= '(大分類為 AE)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       Year as area_year,'+@CR+
                   '       Month as area_month,'+@CR+
                   '       IsNull(ct_no, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(ct_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(amt, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_Target+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and Year >= ''2015'''+@CR+
                   '   and BKind =''AE'''+@CR+
                   ' group by Year, Month, ct_no, ct_name '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '01521'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月客戶大分類實銷金額-電裝'
     set @Remark= '(依請款日, 表身, 含稅, 無加減項, 扣折讓, 含退貨, 含服務單, 大分類為 AE)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_pdate_year as area_year,'+@CR+
                   '       chg_sp_pdate_month as area_month,'+@CR+
                   '       IsNull(sd_ctno, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(chg_ct_sname, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(chg_sd_stot_tax, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslpdt'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and chg_sp_pdate_year >= ''2013'''+@CR+
                   '   and chg_skno_bkind = ''AE'''+@CR+
                   '   and (sd_slip_fg =''2'' Or sd_slip_fg =''C'' Or sd_slip_fg = ''3'')'+@CR+
                   '   and chg_sp_pdate_year is not null'+@CR+
                   -- 2017/08/08 Rickliu 副總指示客編 IT、ZZ 開頭不認列業績，雖不認列業績但應收扔得收
                   '   and Chg_BU_NO Not In (''IT0000'', ''ZZ0000'') '+@CR+
                   ' group by chg_sp_pdate_year, chg_sp_pdate_month, sd_ctno, chg_ct_sname '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '01522'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月客戶大分類銷售金額-電裝'
     set @Remark='(依請款日, 表身, 含稅, 無加減項, 無折讓, 含服務單, 大分類為 AE)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_pdate_year as area_year,'+@CR+
                   '       chg_sp_pdate_month as area_month,'+@CR+
                   '       IsNull(sd_ctno, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(chg_ct_sname, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(chg_sd_stot_tax, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslpdt'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and chg_sp_pdate_year >= ''2015'''+@CR+
                   '   and chg_skno_bkind = ''AE'''+@CR+
                   '   and (sd_slip_fg =''2'' Or sd_slip_fg =''C'')'+@CR+
                   '   and chg_sp_pdate_year is not null'+@CR+
                   -- 2017/08/08 Rickliu 副總指示客編 IT、ZZ 開頭不認列業績，雖不認列業績但應收扔得收
                   '   and Chg_BU_NO Not In (''IT0000'', ''ZZ0000'') '+@CR+
                   ' group by chg_sp_pdate_year, chg_sp_pdate_month, sd_ctno, chg_ct_sname '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
    
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '01523'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月客戶大分類退貨金額-電裝'
     set @Remark= '(依請款日, 表身, 含稅, 無加減項, 無折讓, 大分類為 AE)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_pdate_year as area_year,'+@CR+
                   '       chg_sp_pdate_month as area_month,'+@CR+
                   '       IsNull(sd_ctno, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(chg_ct_sname, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(chg_sd_stot_tax, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslpdt'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and chg_sp_pdate_year >= ''2015'''+@CR+
                   '   and chg_skno_bkind = ''AE'''+@CR+
                   '   and sd_slip_fg =''3'''+@CR+
                   '   and chg_sp_pdate_year is not null'+@CR+
                   -- 2017/08/08 Rickliu 副總指示客編 IT、ZZ 開頭不認列業績，雖不認列業績但應收扔得收
                   '   and Chg_BU_NO Not In (''IT0000'', ''ZZ0000'') '+@CR+
                   ' group by chg_sp_pdate_year, chg_sp_pdate_month, sd_ctno, chg_ct_sname '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '01524'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月客戶大分類實銷金額達成率-電裝'
     set @Remark= '(依請款日)(Formula: 01501 / 電裝目標, 目標資料: ori_xls#Target_Stock_BKind_Personal.Target_AE)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       m.area_year,'+@CR+
                   '       m.area_month,'+@CR+
                   '       IsNull(m.area, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(m.area_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       case'+@CR+
                   '         when IsNull(d.amt, 0) = 0 then 0'+@CR+
                   '         else Round(isnull(m.amt, 0) / isnull(d.amt, 0), 4)'+@CR+
                   '       end as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       3 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '       inner join '+@TB_RT_Data_tmp+' d'+@TB_Hint+@CR+
                   '          on m.area_year = d.area_year'+@CR+
                   '         and m.area_month = d.area_month'+@CR+
                   '         and m.area '+@Collate+'=d.area '+@Collate+@CR+
                   ' where 1=1'+@CR+
                   '   and m.area_year >= ''2015'''+@CR+
                   '   and m.kind =''01521'''+@CR+
                   '   and d.kind =''01520'' '
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '01525'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '年度業務大分類實銷金額達成率-電裝'
     set @Remark= '(依請款日)(Formula: 01501 / 電裝目標, 目標資料: ori_xls#Target_Stock_BKind_Personal.Target_AE)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       m.area_year,'+@CR+
                   '       ''00'' as area_month,'+@CR+
                   '       IsNull(m.area, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(m.area_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       case'+@CR+
                   '         when IsNull(d.Target_Amt, 0) = 0 then 0'+@CR+
                   '         else Round(Sum(Isnull(m.amt, 0)) / d.Target_Amt, 4)'+@CR+
                   '       end as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       3 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '       inner join'+@CR+
                   '         (select area_year, area, Sum(Isnull(amt, 0)) as Target_Amt'+@CR+
                   '            from '+@TB_RT_Data_tmp+' d'+@TB_Hint+@CR+
                   '           where kind=''01520'''+@CR+
                   '           group by area_year, area'+@CR+
                   '         ) D'+@CR+
                   '          on m.area_year=d.area_year'+@CR+
                   '         and m.area '+@Collate+'=d.area '+@Collate+@CR+
                   ' where 1=1'+@CR+
                   '   and m.area_year >= ''2015'''+@CR+
                   '   and m.kind =''01521'''+@CR+
                   ' group by m.area_year, m.area, area_name, Target_Amt'

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end 
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '01526'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月客戶大分類實銷金額-電裝(客戶業績毛利)'
     set @Remark= '(依請款日, 表身, 未稅, 無加減項, 無折讓, 含退貨, 含服務單, 大分類為 AE)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_pdate_year as area_year,'+@CR+
                   '       chg_sp_pdate_month as area_month,'+@CR+
                   '       IsNull(sd_ctno, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(chg_ct_sname, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(Chg_Profit, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslpdt'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and chg_sp_pdate_year >= ''2015'''+@CR+
                   '   and chg_skno_bkind = ''AE'''+@CR+
                   '   and (sd_slip_fg =''2'' Or sd_slip_fg =''C'' Or sd_slip_fg = ''3'')'+@CR+
                   '   and chg_sp_pdate_year is not null'+@CR+
                   -- 2017/08/08 Rickliu 副總指示客編 IT、ZZ 開頭不認列業績，雖不認列業績但應收扔得收
                   '   and Chg_BU_NO Not In (''IT0000'', ''ZZ0000'') '+@CR+
                   ' group by chg_sp_pdate_year, chg_sp_pdate_month, sd_ctno, chg_ct_sname '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '01901'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月業務大分類實銷金額-其他(個人業績)'
     set @Remark= '(依請款日, 表身, 含稅, 無加減項, 無折讓, 含退貨, 含服務單, 大分類扣除 百貨, 3C, 化工, 工具, 電裝及其他 以外的)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_pdate_year as area_year,'+@CR+
                   '       chg_sp_pdate_month as area_month,'+@CR+
                   '       IsNull(sp_sales, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(chg_sales_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(chg_sd_stot_tax, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslpdt'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and chg_sp_pdate_year >= ''2013'''+@CR+
                   '   and (chg_skno_bkind <> ''AA'' And chg_skno_bkind <> ''AB'' And chg_skno_bkind <> ''AC'' And chg_skno_bkind <> ''AD'' And chg_skno_bkind <> ''AE'')'+@CR+
                   '   and (sd_slip_fg =''2'' Or sd_slip_fg =''C'' Or sd_slip_fg = ''3'')'+@CR+
                   '   and chg_sp_pdate_year is not null'+@CR+
                   -- 2017/08/08 Rickliu 副總指示客編 IT、ZZ 開頭不認列業績，雖不認列業績但應收扔得收
                   '   and Chg_BU_NO Not In (''IT0000'', ''ZZ0000'') '+@CR+
                   ' group by chg_sp_pdate_year, chg_sp_pdate_month, sp_sales, chg_sales_name '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '01902'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月業務大分類銷售金額-其他(個人業績)'
     set @Remark= '(依請款日, 表身, 含稅, 無加減項, 無折讓, 含服務單, 大分類扣除 百貨, 3C, 化工, 工具, 電裝及其他 以外的)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_pdate_year as area_year,'+@CR+
                   '       chg_sp_pdate_month as area_month,'+@CR+
                   '       IsNull(sp_sales, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(chg_sales_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(chg_sd_stot_tax, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslpdt'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and chg_sp_pdate_year >= ''2013'''+@CR+
                   '   and (chg_skno_bkind <> ''AA'' And chg_skno_bkind <> ''AB'' And chg_skno_bkind <> ''AC'' And chg_skno_bkind <> ''AD'' And chg_skno_bkind <> ''AE'')'+@CR+
                   '   and (sd_slip_fg =''2'' Or sd_slip_fg =''C'')'+@CR+
                   '   and chg_sp_pdate_year is not null'+@CR+
                   -- 2017/08/08 Rickliu 副總指示客編 IT、ZZ 開頭不認列業績，雖不認列業績但應收扔得收
                   '   and Chg_BU_NO Not In (''IT0000'', ''ZZ0000'') '+@CR+
                   ' group by chg_sp_pdate_year, chg_sp_pdate_month, sp_sales, chg_sales_name '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '01903'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月業務大分類退貨金額-其他(個人業績)'
     set @Remark= '(依請款日, 表身, 含稅, 無加減項, 無折讓, 大分類扣除 百貨, 3C, 化工, 工具, 電裝及其他 以外的)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_pdate_year as area_year,'+@CR+
                   '       chg_sp_pdate_month as area_month,'+@CR+
                   '       IsNull(sp_sales, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(chg_sales_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(chg_sd_stot_tax, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslpdt'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and chg_sp_pdate_year >= ''2013'''+@CR+
                   '   and (chg_skno_bkind <> ''AA'' And chg_skno_bkind <> ''AB'' And chg_skno_bkind <> ''AC'' And chg_skno_bkind <> ''AD'' And chg_skno_bkind <> ''AE'')'+@CR+
                   '   and sd_slip_fg =''3'''+@CR+
                   '   and chg_sp_pdate_year is not null'+@CR+
                   -- 2017/08/08 Rickliu 副總指示客編 IT、ZZ 開頭不認列業績，雖不認列業績但應收扔得收
                   '   and Chg_BU_NO Not In (''IT0000'', ''ZZ0000'') '+@CR+
                   ' group by chg_sp_pdate_year, chg_sp_pdate_month, sp_sales, chg_sales_name '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '01906'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月業務大分類實銷金額-其他(個人業績毛利)'
     set @Remark= '(依請款日, 表身, 未稅, 無加減項, 無折讓, 含退貨, 含服務單, 大分類扣除 百貨, 3C, 化工, 工具, 電裝及其他 以外的)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_pdate_year as area_year,'+@CR+
                   '       chg_sp_pdate_month as area_month,'+@CR+
                   '       IsNull(sp_sales, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(chg_sales_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(Chg_Profit, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslpdt'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and chg_sp_pdate_year >= ''2013'''+@CR+
                   '   and (chg_skno_bkind <> ''AA'' And chg_skno_bkind <> ''AB'' And chg_skno_bkind <> ''AC'' And chg_skno_bkind <> ''AD'' And chg_skno_bkind <> ''AE'')'+@CR+
                   '   and (sd_slip_fg =''2'' Or sd_slip_fg =''C'' Or sd_slip_fg = ''3'')'+@CR+
                   '   and chg_sp_pdate_year is not null'+@CR+
                   -- 2017/08/08 Rickliu 副總指示客編 IT、ZZ 開頭不認列業績，雖不認列業績但應收扔得收
                   '   and Chg_BU_NO Not In (''IT0000'', ''ZZ0000'') '+@CR+
                   ' group by chg_sp_pdate_year, chg_sp_pdate_month, sp_sales, chg_sales_name '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '01921'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月客戶大分類實銷金額-其他'
     set @Remark= '(依請款日, 表身, 含稅, 無加減項, 扣折讓, 含退貨, 含服務單, 大分類為 OT)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_pdate_year as area_year,'+@CR+
                   '       chg_sp_pdate_month as area_month,'+@CR+
                   '       IsNull(sd_ctno, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(chg_ct_sname, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(chg_sd_stot_tax, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslpdt'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and chg_sp_pdate_year >= ''2013'''+@CR+
                   '   and (chg_skno_bkind <> ''AA'' And chg_skno_bkind <> ''AB'' And chg_skno_bkind <> ''AC'' And chg_skno_bkind <> ''AD'' And chg_skno_bkind <> ''AE'')'+@CR+
                   '   and (sd_slip_fg =''2'' Or sd_slip_fg =''C'' Or sd_slip_fg = ''3'')'+@CR+
                   '   and chg_sp_pdate_year is not null'+@CR+
                   -- 2017/08/08 Rickliu 副總指示客編 IT、ZZ 開頭不認列業績，雖不認列業績但應收扔得收
                   '   and Chg_BU_NO Not In (''IT0000'', ''ZZ0000'') '+@CR+
                   ' group by chg_sp_pdate_year, chg_sp_pdate_month, sd_ctno, chg_ct_sname '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '01922'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月客戶大分類銷售金額-其他'
     set @Remark='(依請款日, 表身, 含稅, 無加減項, 無折讓, 含服務單, 大分類為 OT)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_pdate_year as area_year,'+@CR+
                   '       chg_sp_pdate_month as area_month,'+@CR+
                   '       IsNull(sd_ctno, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(chg_ct_sname, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(chg_sd_stot_tax, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslpdt'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and chg_sp_pdate_year >= ''2015'''+@CR+
                   '   and (chg_skno_bkind <> ''AA'' And chg_skno_bkind <> ''AB'' And chg_skno_bkind <> ''AC'' And chg_skno_bkind <> ''AD'' And chg_skno_bkind <> ''AE'')'+@CR+
                   '   and (sd_slip_fg =''2'' Or sd_slip_fg =''C'')'+@CR+
                   '   and chg_sp_pdate_year is not null'+@CR+
                   -- 2017/08/08 Rickliu 副總指示客編 IT、ZZ 開頭不認列業績，雖不認列業績但應收扔得收
                   '   and Chg_BU_NO Not In (''IT0000'', ''ZZ0000'') '+@CR+
                   ' group by chg_sp_pdate_year, chg_sp_pdate_month, sd_ctno, chg_ct_sname '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
    
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '01923'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月客戶大分類退貨金額-其他'
     set @Remark= '(依請款日, 表身, 含稅, 無加減項, 無折讓, 大分類為 OT)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_pdate_year as area_year,'+@CR+
                   '       chg_sp_pdate_month as area_month,'+@CR+
                   '       IsNull(sd_ctno, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(chg_ct_sname, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(chg_sd_stot_tax, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslpdt'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and chg_sp_pdate_year >= ''2015'''+@CR+
                   '   and (chg_skno_bkind <> ''AA'' And chg_skno_bkind <> ''AB'' And chg_skno_bkind <> ''AC'' And chg_skno_bkind <> ''AD'' And chg_skno_bkind <> ''AE'')'+@CR+
                   '   and sd_slip_fg =''3'''+@CR+
                   '   and chg_sp_pdate_year is not null'+@CR+
                   -- 2017/08/08 Rickliu 副總指示客編 IT、ZZ 開頭不認列業績，雖不認列業績但應收扔得收
                   '   and Chg_BU_NO Not In (''IT0000'', ''ZZ0000'') '+@CR+
                   ' group by chg_sp_pdate_year, chg_sp_pdate_month, sd_ctno, chg_ct_sname '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '01924'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月客戶大分類實銷金額達成率-其他'
     set @Remark= '(依請款日)(Formula: 01501 / 其他目標, 目標資料: ori_xls#Target_Stock_BKind_Personal.Target_OT)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       m.area_year,'+@CR+
                   '       m.area_month,'+@CR+
                   '       IsNull(m.area, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(m.area_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       case'+@CR+
                   '         when IsNull(d.amt, 0) = 0 then 0'+@CR+
                   '         else Round(isnull(m.amt, 0) / isnull(d.amt, 0), 4)'+@CR+
                   '       end as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       3 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '       inner join '+@TB_RT_Data_tmp+' d'+@TB_Hint+@CR+
                   '          on m.area_year = d.area_year'+@CR+
                   '         and m.area_month = d.area_month'+@CR+
                   '         and m.area '+@Collate+'=d.area '+@Collate+@CR+
                   ' where 1=1'+@CR+
                   '   and m.area_year >= ''2015'''+@CR+
                   '   and m.kind =''01921'''+@CR+
                   '   and d.kind =''01920'' '
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '01925'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '年度業務大分類實銷金額達成率-其他'
     set @Remark= '(依請款日)(Formula: 01501 / 其他目標, 目標資料: ori_xls#Target_Stock_BKind_Personal.Target_OT)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       m.area_year,'+@CR+
                   '       ''00'' as area_month,'+@CR+
                   '       IsNull(m.area, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(m.area_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       case'+@CR+
                   '         when IsNull(d.Target_Amt, 0) = 0 then 0'+@CR+
                   '         else Round(Sum(Isnull(m.amt, 0)) / d.Target_Amt, 4)'+@CR+
                   '       end as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       3 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '       inner join'+@CR+
                   '         (select area_year, area, Sum(Isnull(amt, 0)) as Target_Amt'+@CR+
                   '            from '+@TB_RT_Data_tmp+' d'+@TB_Hint+@CR+
                   '           where kind=''01920'''+@CR+
                   '           group by area_year, area'+@CR+
                   '         ) D'+@CR+
                   '          on m.area_year=d.area_year'+@CR+
                   '         and m.area '+@Collate+'=d.area '+@Collate+@CR+
                   ' where 1=1'+@CR+
                   '   and m.area_year >= ''2015'''+@CR+
                   '   and m.kind =''01921'''+@CR+
                   ' group by m.area_year, m.area, area_name, Target_Amt'

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end 
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '01926'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月客戶大分類實銷金額-其他(客戶業績毛利)'
     set @Remark= '(依請款日, 表身, 未稅, 無加減項, 無折讓, 含退貨, 含服務單, 大分類為 OT)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_pdate_year as area_year,'+@CR+
                   '       chg_sp_pdate_month as area_month,'+@CR+
                   '       IsNull(sd_ctno, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(chg_ct_sname, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(Chg_Profit, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslpdt'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and chg_sp_pdate_year >= ''2015'''+@CR+
                   '   and (chg_skno_bkind <> ''AA'' And chg_skno_bkind <> ''AB'' And chg_skno_bkind <> ''AC'' And chg_skno_bkind <> ''AD'' And chg_skno_bkind <> ''AE'')'+@CR+
                   '   and (sd_slip_fg =''2'' Or sd_slip_fg =''C'' Or sd_slip_fg = ''3'')'+@CR+
                   '   and chg_sp_pdate_year is not null'+@CR+
                   -- 2017/08/08 Rickliu 副總指示客編 IT、ZZ 開頭不認列業績，雖不認列業績但應收扔得收
                   '   and Chg_BU_NO Not In (''IT0000'', ''ZZ0000'') '+@CR+
                   ' group by chg_sp_pdate_year, chg_sp_pdate_month, sd_ctno, chg_ct_sname '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  -- 此類型資料比較特別無法使用KIND 方式產生資料，會依系列數量產生對應的資料出來。
  -- 資料範圍為 02000~02099
  set @Kind = '02000'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月業務系列商品實銷金額'
     set @Remark= '動態資料範圍為 02000~02099, 依請款日, 表身, 含稅, 無折讓, 無加減項, 含退貨, 含服務單, 每月業務系列商品不含 Z 類別'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind_name like ''%'+@Kind_name+'%'' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere+@CR+
                   'Delete '+@TB_RT_Data+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select Substring(Convert(Varchar(6), Convert(Int, '''+@Kind+''')-1 + DENSE_RANK() Over(Order By sk_kind) + 100000), 2, 5) '+@Collate+' as kind,'+@CR+
                   '       Substring(Convert(Varchar(6), Convert(Int, '''+@Kind+''')-1 + DENSE_RANK() Over(Order By sk_kind) + 100000), 2, 5)+'+@CR+
                   '       ''.'+@Kind_Name+'(''+Rtrim(IsNull(sk_kind, ''''))+'' ''+Rtrim(IsNull(chg_kind_name, ''''))+'')'' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_pdate_year as area_year,'+@CR+
                   '       chg_sp_pdate_month as area_month,'+@CR+
                   '       IsNull(sp_sales, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(sp_sales_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(chg_sd_stot_tax, 0)) as amt,'+@CR+
                   '       ''Y'' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslpdt'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and (sd_slip_fg =''2'' Or sd_slip_fg =''C'' Or sd_slip_fg = ''3'')'+@CR+
                   '   and IsNull(chg_kind_name, '''') <> '''''+@CR+
                   '   and sk_kind <> ''Z'''+@CR+
                   '   and chg_sp_pdate_year >= ''2013'''+@CR+
                   -- 2017/08/08 Rickliu 副總指示客編 IT、ZZ 開頭不認列業績，雖不認列業績但應收扔得收
                   '   and Chg_BU_NO Not In (''IT0000'', ''ZZ0000'') '+@CR+
                   ' group by chg_sp_pdate_year, chg_sp_pdate_month, sk_kind, chg_kind_name, sp_sales, sp_sales_name '
                   
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.刪除 '+@Kind_Name+@CNonSet+'的資料 ['+@TB_RT_Data_tmp+']'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+@CR+
                   ' where 1=1'+@CR+
                   '   and kind_name like ''%'+@CNonSet+'%'''+@CR+
                   '   and kind in'+@CR+
                   '       (select kind'+@CR+
                   '          from '+@TB_RT_Data_tmp+@TB_Hint+@CR+
                   '         '+@strWhere+@CR+
                   '       )'+@CR+
                   'Delete '+@TB_RT_Data+@CR+
                   ' where 1=1'+@CR+
                   '   and kind_name like ''%'+@CNonSet+'%'''+@CR+
                   '   and kind in'+@CR+
                   '       (select kind'+@CR+
                   '          from '+@TB_RT_Data_tmp+@TB_Hint+@CR+
                   '         '+@strWhere+@CR+
                   '       ) '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end  

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  -- 資料範圍為 02100~02199
  set @Kind = '02100'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月業務系列商品銷售金額'
     set @Remark= '動態資料範圍為 02100~02199, 依請款日, 表身, 含稅, 無折讓, 無加減項, 含服務單, 每月業務系列商品不含 Z 類別'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind_name like ''%'+@Kind_name+'%'' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere+@CR+
                   'Delete '+@TB_RT_Data+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select Substring(Convert(Varchar(6), Convert(Int, '''+@Kind+''')-1 + DENSE_RANK() Over(Order By sk_kind) + 100000), 2, 5) '+@Collate+' as kind,'+@CR+
                   '       Substring(Convert(Varchar(6), Convert(Int, '''+@Kind+''')-1 + DENSE_RANK() Over(Order By sk_kind) + 100000), 2, 5)+'+@CR+
                   '       ''.'+@Kind_Name+'(''+Rtrim(IsNull(sk_kind, ''''))+'' ''+Rtrim(IsNull(chg_kind_name, ''''))+'')'' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_pdate_year as area_year,'+@CR+
                   '       chg_sp_pdate_month as area_month,'+@CR+
                   '       IsNull(sp_sales, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(sp_sales_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(chg_sd_stot_tax, 0)) as amt,'+@CR+
                   '       ''Y'' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslpdt'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and (sd_slip_fg =''2'' Or sd_slip_fg =''C'')'+@CR+
                   '   and IsNull(chg_kind_name, '''') <> '''''+@CR+
                   '   and sk_kind <> ''Z'''+@CR+
                   '   and chg_sp_pdate_year >= ''2013'''+@CR+
                   -- 2017/08/08 Rickliu 副總指示客編 IT、ZZ 開頭不認列業績，雖不認列業績但應收扔得收
                   '   and Chg_BU_NO Not In (''IT0000'', ''ZZ0000'') '+@CR+
                   ' group by chg_sp_pdate_year, chg_sp_pdate_month, sk_kind, chg_kind_name, sp_sales, sp_sales_name '
                   
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.刪除 '+@Kind_Name+@CNonSet+'的資料 ['+@TB_RT_Data_tmp+']'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+@CR+
                   ' where 1=1'+@CR+
                   '   and kind_name like ''%'+@CNonSet+'%'''+@CR+
                   '   and kind in'+@CR+
                   '       (select kind'+@CR+
                   '          from '+@TB_RT_Data_tmp+@TB_Hint+@CR+
                   '         '+@strWhere+@CR+
                   '       )'+@CR+
                   'Delete '+@TB_RT_Data+@CR+
                   ' where 1=1'+@CR+
                   '   and kind_name like ''%'+@CNonSet+'%'''+@CR+
                   '   and kind in'+@CR+
                   '       (select kind'+@CR+
                   '          from '+@TB_RT_Data_tmp+@TB_Hint+@CR+
                   '         '+@strWhere+@CR+
                   '       ) '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end  

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  -- 資料範圍為 02200~02299
  set @Kind = '02200'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月業務系列商品退貨金額'
     set @Remark= '動態資料範圍為 02200~02299, 依請款日, 表身, 含稅, 無折讓, 無加減項, 每月業務系列商品不含 Z 類別'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind_name like ''%'+@Kind_name+'%'' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere+@CR+
                   'Delete '+@TB_RT_Data+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select Substring(Convert(Varchar(6), Convert(Int, '''+@Kind+''')-1 + DENSE_RANK() Over(Order By sk_kind) + 100000), 2, 5) '+@Collate+' as kind,'+@CR+
                   '       Substring(Convert(Varchar(6), Convert(Int, '''+@Kind+''')-1 + DENSE_RANK() Over(Order By sk_kind) + 100000), 2, 5)+'+@CR+
                   '       ''.'+@Kind_Name+'(''+Rtrim(IsNull(sk_kind, ''''))+'' ''+Rtrim(IsNull(chg_kind_name, ''''))+'')'' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_pdate_year as area_year,'+@CR+
                   '       chg_sp_pdate_month as area_month,'+@CR+
                   '       IsNull(sp_sales, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(sp_sales_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(chg_sd_stot_tax, 0)) as amt,'+@CR+
                   '       ''Y'' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslpdt'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and sd_slip_fg =''3'''+@CR+
                   '   and IsNull(chg_kind_name, '''') <> '''''+@CR+
                   '   and sk_kind <> ''Z'''+@CR+
                   '   and chg_sp_pdate_year >= ''2013'''+@CR+
                   -- 2017/08/08 Rickliu 副總指示客編 IT、ZZ 開頭不認列業績，雖不認列業績但應收扔得收
                   '   and Chg_BU_NO Not In (''IT0000'', ''ZZ0000'') '+@CR+
                   ' group by chg_sp_pdate_year, chg_sp_pdate_month, sk_kind, chg_kind_name, sp_sales, sp_sales_name '
                   
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.刪除 '+@Kind_Name+@CNonSet+'的資料 ['+@TB_RT_Data_tmp+']'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+@CR+
                   ' where 1=1'+@CR+
                   '   and kind_name like ''%'+@CNonSet+'%'''+@CR+
                   '   and kind in'+@CR+
                   '       (select kind'+@CR+
                   '          from '+@TB_RT_Data_tmp+@TB_Hint+@CR+
                   '         '+@strWhere+@CR+
                   '       )'+@CR+
                   'Delete '+@TB_RT_Data+@CR+
                   ' where 1=1'+@CR+
                   '   and kind_name like ''%'+@CNonSet+'%'''+@CR+
                   '   and kind in'+@CR+
                   '       (select kind'+@CR+
                   '          from '+@TB_RT_Data_tmp+@TB_Hint+@CR+
                   '         '+@strWhere+@CR+
                   '       ) '
                   
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end  

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  -- 資料範圍為 02300~02399
  set @Kind = '02300'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月業務系列商品實銷金額佔比'
     set @Remark= '動態資料範圍為 02300~02399, 每月業務系列商品不含 Z 類別, Formula: (00180..00209) / Sum(00180..00209)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind_name like ''%'+@Kind_name+'%'' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere+@CR+
                   'Delete '+@TB_RT_Data+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select Substring(Convert(Varchar(5), Convert(Int, '''+@Kind+''')-1 + DENSE_RANK() Over(Order By m.kind) + 100000), 2, 5) '+@Collate+' as kind,'+@CR+
                   '       Substring(Convert(Varchar(5), Convert(Int, '''+@Kind+''')-1 + DENSE_RANK() Over(Order By m.kind) + 100000), 2, 5)+'+@CR+
                   '       ''.'+@Kind_Name+'(由 ''+Rtrim(IsNull(m.kind_name, ''''))+'' 產生而來)'' '+@Collate+' as kind_name,'+@CR+
                   '       m.area_year,'+@CR+
                   '       m.area_month,'+@CR+
                   '       area,'+@CR+
                   '       area_name,'+@CR+
                   '       case'+@CR+
                   '         when IsNull(d.tot_amt, 0) = 0'+@CR+
                   '         then 0'+@CR+
                   '         else Round(Sum(Isnull(m.amt, 0)) / d.tot_amt, 4)'+@CR+
                   '       end as amt,'+@CR+
                   '       ''Y'' as Not_Accumulate,'+@CR+
                   '       3 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '       inner join'+@CR+
                   '       (select kind,'+@CR+
                   '               area_year,'+@CR+
                   '               area_month,'+@CR+
                   '               Sum(IsNull(amt, 0)) as tot_amt'+@CR+
                   '          from '+@TB_RT_Data_tmp+@TB_Hint+@CR+
                   '         where 1=1'+@CR+
                   '           and kind >= ''00180'''+@CR+
                   '           and kind <= ''00209'''+@CR+
                   '         group by kind, area_year, area_month'+@CR+
                   '       ) d'+@CR+
                   '        on m.kind = d.kind'+@CR+
                   '       and m.area_year = d.area_year'+@CR+
                   '       and m.area_month = d.area_month'+@CR+
                   ' where 1=1'+@CR+
                   '   and m.kind >= ''00180'''+@CR+
                   '   and m.kind <= ''00209'''+@CR+
                   '   and m.area_year <> 0'+@CR+
                   '   and m.area_month <> 0'+@CR+
                   ' group by m.kind, m.kind_name, m.area_year, m.area_month, m.area, m.area_name, d.tot_amt '
                   
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.刪除 '+@Kind_Name+@CNonSet+'的資料 ['+@TB_RT_Data_tmp+']'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+@CR+
                   ' where 1=1'+@CR+
                   '   and kind_name like ''%'+@CNonSet+'%'''+@CR+
                   '   and kind in'+@CR+
                   '       (select kind'+@CR+
                   '          from '+@TB_RT_Data_tmp+@TB_Hint+@CR+
                   '         '+@strWhere+@CR+
                   '       )'+@CR+
                   'Delete '+@TB_RT_Data+@CR+
                   ' where 1=1'+@CR+
                   '   and kind_name like ''%'+@CNonSet+'%'''+@CR+
                   '   and kind in'+@CR+
                   '       (select kind'+@CR+
                   '          from '+@TB_RT_Data_tmp+@TB_Hint+@CR+
                   '         '+@strWhere+@CR+
                   '       ) '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end  

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '03001'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月各大分類實銷金額(已收退金額)'
     set @Remark= '(請款日, 表身, 含稅, 含折讓, 含加減項, 含退貨, 含服務單, 原大分類)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_pdate_year,'+@CR+
                   '       chg_sp_pdate_month,'+@CR+
                   '       IsNull(chg_skno_bkind, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(chg_skno_bkind_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(Chg_sd_pay_Sale, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslpdt'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and chg_sp_pdate_year >= ''2013'''+@CR+
                   '   and (sd_slip_fg =''2'' Or sd_slip_fg =''C'' Or sd_slip_fg = ''3'')'+@CR+
                   '   and chg_sp_pdate_year is not null'+@CR+
                   -- 2017/08/08 Rickliu 副總指示客編 IT、ZZ 開頭不認列業績，雖不認列業績但應收扔得收
                   '   and Chg_BU_NO Not In (''IT0000'', ''ZZ0000'') '+@CR+
                   ' group by chg_sp_pdate_year, chg_sp_pdate_month, chg_skno_bkind, chg_skno_bkind_name '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '03002'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月各大分類已收金額(銷售金額)'
     set @Remark= '(請款日, 表身, 含稅, 含折讓, 含加減項, 含服務單, 原大分類)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_pdate_year,'+@CR+
                   '       chg_sp_pdate_month,'+@CR+
                   '       IsNull(chg_skno_bkind, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(chg_skno_bkind_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(Chg_sd_pay_Sale, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslpdt'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and chg_sp_pdate_year >= ''2013'''+@CR+
                   '   and (sd_slip_fg =''2'' Or sd_slip_fg =''C'')'+@CR+
                   '   and chg_sp_pdate_year is not null'+@CR+
                   -- 2017/08/08 Rickliu 副總指示客編 IT、ZZ 開頭不認列業績，雖不認列業績但應收扔得收
                   '   and Chg_BU_NO Not In (''IT0000'', ''ZZ0000'') '+@CR+
                   ' group by chg_sp_pdate_year, chg_sp_pdate_month, chg_skno_bkind, chg_skno_bkind_name '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '03003'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月各大分類已退金額(退貨金額)'
     set @Remark= '(請款日, 表身, 含稅, 含折讓, 含加減項, 原大分類)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_pdate_year,'+@CR+
                   '       chg_sp_pdate_month,'+@CR+
                   '       IsNull(chg_skno_bkind, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(chg_skno_bkind_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(Chg_sd_pay_Sale, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslpdt'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and chg_sp_pdate_year >= ''2013'''+@CR+
                   '   and sd_slip_fg =''3'''+@CR+
                   '   and chg_sp_pdate_year is not null'+@CR+
                   -- 2017/08/08 Rickliu 副總指示客編 IT、ZZ 開頭不認列業績，雖不認列業績但應收扔得收
                   '   and Chg_BU_NO Not In (''IT0000'', ''ZZ0000'') '+@CR+
                   ' group by chg_sp_pdate_year, chg_sp_pdate_month, chg_skno_bkind, chg_skno_bkind_name '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '03011'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月大分類實銷總額(已收退金額)(僅百貨, 3C, 化工, 工具, 電裝及其他)'
     set @Remark= '每月大分類實銷總額由 03001 加總而來'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       area_year,'+@CR+
                   '       area_month,'+@CR+
                   '       IsNull(area, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(area_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Isnull(amt, 0) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and kind=''03001'''+@CR+
                   '   and (area = ''AA'' Or area = ''AB'' Or area = ''AC'' Or area = ''AD'' Or area = ''AE'')'+@CR+
                   ' union'+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       area_year,'+@CR+
                   '       area_month,'+@CR+
                   '       ''OT'' as area,'+@CR+
                   '       ''其他'' as area_name,'+@CR+
                   '       Sum(Isnull(amt, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and kind=''03001'''+@CR+
                   '   and (area <> ''AA'' and area <> ''AB'' and area = ''AC'' and area = ''AD'' and area = ''AE'')'+@CR+
                   ' group by area_year, area_month '
                   
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '03012'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月大分類已收金額(銷售總額)(僅百貨, 3C, 化工, 工具, 電裝及其他)'
     set @Remark= '每月大分類銷售總額由 03002 加總而來'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       area_year,'+@CR+
                   '       area_month,'+@CR+
                   '       IsNull(area, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(area_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       IsNull(amt, 0) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and kind=''03002'''+@CR+
                   '   and (area =''AA'' Or area = ''AB'' Or area =''AC'' Or area = ''AD'' Or area = ''AE'')'+@CR+
                   ' union'+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       area_year,'+@CR+
                   '       area_month,'+@CR+
                   '       ''OT'' as area,'+@CR+
                   '       ''其他'' as area_name,'+@CR+
                   '       Sum(Isnull(amt, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and kind=''03002'''+@CR+
                   '   and (area <> ''AA'' and area <> ''AB'' and area = ''AC'' and area = ''AD'' and area = ''AE'')'+@CR+
                   ' group by area_year, area_month '
                   
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '03013'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月大分類已退金額(退貨總額)(僅百貨, 3C, 化工, 工具, 電裝及其他)'
     set @Remark= '每月大分類退貨總額由 03003 加總而來'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       area_year,'+@CR+
                   '       area_month,'+@CR+
                   '       IsNull(area, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(area_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       IsNull(amt, 0) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and kind=''03003'''+@CR+
                   '   and (area = ''AA'' Or area = ''AB'' Or area = ''AC'' Or area = ''AD'' Or area = ''AE'')'+@CR+
                   ' union'+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       area_year,'+@CR+
                   '       area_month,'+@CR+
                   '       ''OT'' as area,'+@CR+
                   '       ''其他'' as area_name,'+@CR+
                   '       Sum(IsNull(amt, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and kind=''03003'''+@CR+
                   '   and (area <> ''AA'' and area <> ''AB'' and area = ''AC'' and area = ''AD'' and area = ''AE'')'+@CR+
                   ' group by area_year, area_month '
                   
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

    --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '03014'
if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每年大分類實銷總額(已收退金額)(僅百貨, 3C, 化工, 工具, 電裝及其他)'
     set @Remark= '每年大分類實銷總額由 03001 加總而來'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       area_year,'+@CR+
                   '       ''00'' as area_month,'+@CR+
                   '       IsNull(area, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(area_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(amt, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and kind=''03001'''+@CR+
                   '   and (area =''AA'' Or area = ''AB'' Or area = ''AC'' Or area = ''AD'' Or area = ''AE'')'+@CR+
                   ' group by area_year, area, area_name'+@CR+
                   ' union'+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       area_year,'+@CR+
                   '       ''00'' as area_month,'+@CR+
                   '       ''OT'' as area,'+@CR+
                   '       ''其他'' as area_name,'+@CR+
                   '       Sum(IsNull(amt, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and kind=''03001'''+@CR+
                   '   and (area <> ''AA'' and area <> ''AB'' and area = ''AC'' and area = ''AD'' and area = ''AE'')'+@CR+
                   ' group by area_year '
                   
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '03015'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月大分類實銷佔比(已收退金額)(僅百貨, 3C, 化工, 工具, 電裝及其他)'
     set @Remark= '每月大分類實銷佔比 Formula: 03011 / Sum(03011)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       m.area_year,'+@CR+
                   '       m.area_month,'+@CR+
                   '       IsNull(m.area, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(m.area_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       case'+@CR+
                   '         when Isnull(d.tot_amt, 0) = 0 then 0'+@CR+
                   '         else Round(Isnull(m.amt, 0) / d.tot_amt, 4)'+@CR+
                   '       end as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       3 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '       inner join'+@CR+
                   '       (select area_year, area_month, Sum(IsNull(amt, 0)) as tot_amt'+@CR+
                   '          from '+@TB_RT_Data_tmp+@TB_Hint+@CR+
                   '         where 1=1'+@CR+
                   '           and kind =''03011'''+@CR+
                   '         group by area_year, area_month'+@CR+
                   '       ) d'+@CR+
                   '        on m.area_year=d.area_year'+@CR+
                   '       and m.area_month=d.area_month'+@CR+
                   ' where 1=1'+@CR+
                   '   and m.kind=''03011'' '
                   
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '03016'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每年大分類實銷佔比(已收退金額)(僅百貨, 3C, 化工, 工具, 電裝及其他)'
     set @Remark= '每年大分類實銷佔比 Formula: 03014 / Sum(03014)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       m.area_year,'+@CR+
                   '       m.area_month,'+@CR+
                   '       IsNull(m.area, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(m.area_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       case'+@CR+
                   '         when Isnull(d.tot_amt, 0) = 0 then 0'+@CR+
                   '         else Round(Isnull(m.amt, 0) / d.tot_amt, 4)'+@CR+
                   '       end as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       3 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '       inner join'+@CR+
                   '       (select area_year, Sum(IsNull(amt, 0)) as tot_amt'+@CR+
                   '          from '+@TB_RT_Data_tmp+@TB_Hint+@CR+
                   '         where 1=1'+@CR+
                   '           and kind =''03014'''+@CR+
                   '         group by area_year'+@CR+
                   '       ) d'+@CR+
                   '        on m.area_year=d.area_year'+@CR+
                   ' where 1=1'+@CR+
                   '   and m.kind=''03014'' '
                   
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '03017'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月大分類實銷達成率(已收退金額)(僅百貨, 3C, 化工, 工具, 電裝及其他)'
     set @Remark= '每月大分類實銷達成率 Formula: 03011 / 01004 加總而來'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       m.area_year,'+@CR+
                   '       m.area_month,'+@CR+
                   '       IsNull(m.area, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(m.area_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       case'+@CR+
                   '         when Isnull(d.amt, 0) = 0 then 0'+@CR+
                   '         else Round(Isnull(m.amt, 0) / d.amt, 4)'+@CR+
                   '       end as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       3 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '       inner join '+@TB_RT_Data_tmp+' d'+@TB_Hint+@CR+
                   '          on m.area_year=d.area_year'+@CR+
                   '         and m.area_month=d.area_month'+@CR+
                   '         and m.area=d.area'+@CR+
                   '         and d.kind =''01004'''+@CR+
                   ' where 1=1'+@CR+
                   '   and m.kind=''03011'' '
                   
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '03018'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每年大分類實銷達成率(已收退金額)(僅百貨, 3C, 化工, 工具, 電裝及其他)'
     set @Remark= '每年大分類實銷達成率 Formula: 03014 / 01005 加總而來'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       m.area_year,'+@CR+
                   '       m.area_month,'+@CR+
                   '       IsNull(m.area, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(m.area_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       case'+@CR+
                   '         when Isnull(d.amt, 0) = 0 then 0'+@CR+
                   '         else Round(Isnull(m.amt, 0) / d.amt, 4)'+@CR+
                   '       end as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       3 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '       inner join '+@TB_RT_Data_tmp+' d'+@TB_Hint+@CR+
                   '          on m.area_year=d.area_year'+@CR+
                   '         and m.area_month=d.area_month'+@CR+
                   '         and m.area=d.area'+@CR+
                   '         and d.kind =''01005'''+@CR+
                   ' where 1=1'+@CR+
                   '   and m.kind=''03014'' '
                   
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '03050'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月業務實銷金額(已收退金額)'
     set @Remark= '(請款日, 表身, 含稅, 含折讓, 含加減項)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_pdate_year,'+@CR+
                   '       chg_sp_pdate_month,'+@CR+
                   '       IsNull(sp_sales, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(chg_sales_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(Chg_SP_PAY) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslip'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and chg_sp_pdate_year >= ''2013'''+@CR+
                   '   and (sp_slip_fg =''2'' Or sp_slip_fg =''C'' Or sp_slip_fg = ''3'')'+@CR+
                   '   and chg_sp_pdate_year is not null'+@CR+
                   -- 2017/08/08 Rickliu 副總指示客編 IT、ZZ 開頭不認列業績，雖不認列業績但應收扔得收
                   '   and Chg_BU_NO Not In (''IT0000'', ''ZZ0000'') '+@CR+
                   ' group by chg_sp_pdate_year, chg_sp_pdate_month, sp_sales, chg_sales_name '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '03051'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月業務總計金額(實銷金額)'
     set @Remark= '(請款日, 表身, 含稅, 含折讓, 含加減項)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_pdate_year,'+@CR+
                   '       chg_sp_pdate_month,'+@CR+
                   '       IsNull(sp_sales, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(chg_sales_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       sum(sp_tot + sp_tax) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslip'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and chg_sp_pdate_year >= ''2013'''+@CR+
                   '   and (sp_slip_fg =''2'' Or sp_slip_fg =''C'' Or sp_slip_fg = ''3'')'+@CR+
                   '   and chg_sp_pdate_year is not null'+@CR+
                   -- 2017/08/08 Rickliu 副總指示客編 IT、ZZ 開頭不認列業績，雖不認列業績但應收扔得收
                   '   and Chg_BU_NO Not In (''IT0000'', ''ZZ0000'') '+@CR+
                   ' group by chg_sp_pdate_year, chg_sp_pdate_month, sp_sales, chg_sales_name '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '03052'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月業務實銷達成率(已收退金額)'
     set @Remark= '(請款日, 表身, 含稅, 含折讓, 含加減項)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_pdate_year,'+@CR+
                   '       chg_sp_pdate_month,'+@CR+
                   '       IsNull(sp_sales, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(chg_sales_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       case'+@CR+
                   '         when sum(sp_tot) <> 0 then Sum(Chg_SP_PAY) / sum(sp_tot + sp_tax)'+@CR+
                   '         else 0'+@CR+
                   '       end as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslip'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and chg_sp_pdate_year >= ''2013'''+@CR+
                   '   and (sp_slip_fg =''2'' Or sp_slip_fg =''C'' Or sp_slip_fg = ''3'')'+@CR+
                   '   and chg_sp_pdate_year is not null'+@CR+
                   -- 2017/08/08 Rickliu 副總指示客編 IT、ZZ 開頭不認列業績，雖不認列業績但應收扔得收
                   '   and Chg_BU_NO Not In (''IT0000'', ''ZZ0000'') '+@CR+
                   ' group by chg_sp_pdate_year, chg_sp_pdate_month, sp_sales, chg_sales_name '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '03060'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每年業務實銷金額(已收退金額)'
     set @Remark= '(請款日, 表身, 含稅, 含折讓, 含加減項)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_pdate_year,'+@CR+
                   '       '''+@NonSet+''' as chg_sp_pdate_month,'+@CR+
                   '       IsNull(sp_sales, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(chg_sales_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(Chg_SP_PAY) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslip'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and chg_sp_pdate_year >= ''2013'''+@CR+
                   '   and (sp_slip_fg =''2'' Or sp_slip_fg =''C'' Or sp_slip_fg = ''3'')'+@CR+
                   '   and chg_sp_pdate_year is not null'+@CR+
                   '   and e_dept like ''B2%'''+@CR+
                   -- 2017/08/08 Rickliu 副總指示客編 IT、ZZ 開頭不認列業績，雖不認列業績但應收扔得收
                   '   and Chg_BU_NO Not In (''IT0000'', ''ZZ0000'') '+@CR+
                   ' group by chg_sp_pdate_year, sp_sales, chg_sales_name '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '03061'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每年業務總計金額(實銷金額)'
     set @Remark= '(請款日, 表身, 含稅, 含折讓, 含加減項)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_pdate_year,'+@CR+
                   '       '''+@NonSet+''' as chg_sp_pdate_month,'+@CR+
                   '       IsNull(sp_sales, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(chg_sales_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       sum(sp_tot + sp_tax) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslip'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and chg_sp_pdate_year >= ''2013'''+@CR+
                   '   and (sp_slip_fg =''2'' Or sp_slip_fg =''C'' Or sp_slip_fg = ''3'')'+@CR+
                   '   and chg_sp_pdate_year is not null'+@CR+
                   -- 2017/08/08 Rickliu 副總指示客編 IT、ZZ 開頭不認列業績，雖不認列業績但應收扔得收
                   '   and Chg_BU_NO Not In (''IT0000'', ''ZZ0000'') '+@CR+
                   ' group by chg_sp_pdate_year, sp_sales, chg_sales_name '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '03062'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每年業務實銷達成率(已收退金額)'
     set @Remark= '(請款日, 表身, 含稅, 含折讓, 含加減項)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_pdate_year,'+@CR+
                   '       '''+@NonSet+''' as chg_sp_pdate_month,'+@CR+
                   '       IsNull(sp_sales, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(chg_sales_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       case'+@CR+
                   '         when sum(sp_tot) <> 0 then Sum(Chg_SP_PAY) / sum(sp_tot + sp_tax)'+@CR+
                   '         else 0'+@CR+
                   '       end as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslip'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and chg_sp_pdate_year >= ''2013'''+@CR+
                   '   and (sp_slip_fg =''2'' Or sp_slip_fg =''C'' Or sp_slip_fg = ''3'')'+@CR+
                   '   and chg_sp_pdate_year is not null'+@CR+
                   -- 2017/08/08 Rickliu 副總指示客編 IT、ZZ 開頭不認列業績，雖不認列業績但應收扔得收
                   '   and Chg_BU_NO Not In (''IT0000'', ''ZZ0000'') '+@CR+
                   ' group by chg_sp_pdate_year, sp_sales, chg_sales_name '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '03101'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月業務大分類實銷金額-百貨(個人業績)(已收退金額)'
     set @Remark= '(依請款日, 表身, 含稅, 含加減項, 含折讓, 含退貨, 含服務單, 大分類為 AA)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_pdate_year as area_year,'+@CR+
                   '       chg_sp_pdate_month as area_month,'+@CR+
                   '       IsNull(sp_sales, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(chg_sales_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(Chg_sd_pay_Sale, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslpdt'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and chg_sp_pdate_year >= ''2013'''+@CR+
                   '   and chg_skno_bkind = ''AA'''+@CR+
                   '   and (sd_slip_fg =''2'' Or sd_slip_fg =''C'' Or sd_slip_fg = ''3'')'+@CR+
                   '   and chg_sp_pdate_year is not null'+@CR+
                   -- 2017/08/08 Rickliu 副總指示客編 IT、ZZ 開頭不認列業績，雖不認列業績但應收扔得收
                   '   and Chg_BU_NO Not In (''IT0000'', ''ZZ0000'') '+@CR+
                   ' group by chg_sp_pdate_year, chg_sp_pdate_month, sp_sales, chg_sales_name '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '03102'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月業務大分類銷售金額(已收金額)-百貨(個人業績)'
     set @Remark='(依請款日, 表身, 含稅, 含加減項, 含折讓, 含服務單, 大分類為 AA)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_pdate_year as area_year,'+@CR+
                   '       chg_sp_pdate_month as area_month,'+@CR+
                   '       IsNull(sp_sales, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(chg_sales_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(Chg_sd_pay_Sale, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslpdt'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and chg_sp_pdate_year >= ''2013'''+@CR+
                   '   and chg_skno_bkind = ''AA'''+@CR+
                   '   and (sd_slip_fg =''2'' Or sd_slip_fg =''C'')'+@CR+
                   '   and chg_sp_pdate_year is not null'+@CR+
                   -- 2017/08/08 Rickliu 副總指示客編 IT、ZZ 開頭不認列業績，雖不認列業績但應收扔得收
                   '   and Chg_BU_NO Not In (''IT0000'', ''ZZ0000'') '+@CR+
                   ' group by chg_sp_pdate_year, chg_sp_pdate_month, sp_sales, chg_sales_name '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
    
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '03103'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月業務大分類退貨金額(已退金額)-百貨(個人業績)'
     set @Remark= '(依請款日, 表身, 含稅, 含加減項, 含折讓, 大分類為 AA)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_pdate_year as area_year,'+@CR+
                   '       chg_sp_pdate_month as area_month,'+@CR+
                   '       IsNull(sp_sales, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(chg_sales_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(Chg_sd_pay_Sale, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslpdt'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and chg_sp_pdate_year >= ''2013'''+@CR+
                   '   and chg_skno_bkind = ''AA'''+@CR+
                   '   and sd_slip_fg =''3'''+@CR+
                   '   and chg_sp_pdate_year is not null'+@CR+
                   -- 2017/08/08 Rickliu 副總指示客編 IT、ZZ 開頭不認列業績，雖不認列業績但應收扔得收
                   '   and Chg_BU_NO Not In (''IT0000'', ''ZZ0000'') '+@CR+
                   ' group by chg_sp_pdate_year, chg_sp_pdate_month, sp_sales, chg_sales_name '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '03104'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月業務大分類實銷金額達成率-百貨(個人業績)(已收退金額)'
     set @Remark= '(依請款日)(Formula: 03101 / 百貨目標, 目標資料: ori_xls#Target_Stock_BKind_Personal.Target_AA)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       m.area_year,'+@CR+
                   '       m.area_month,'+@CR+
                   '       IsNull(m.area, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(m.area_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       case'+@CR+
                   '         when IsNull(d.amt, 0) = 0 then 0'+@CR+
                   '         else Round(isnull(m.amt, 0) / isnull(d.amt, 0), 4)'+@CR+
                   '       end as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       3 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '       inner join '+@TB_RT_Data_tmp+' D'+@TB_Hint+@CR+
                   '          on m.area_year = d.area_year'+@CR+
                   '         and m.area_month = d.area_month'+@CR+
                   '         and m.area =d.area'+@CR+
                   ' where 1=1'+@CR+
                   '   and m.area_year >= ''2013'''+@CR+
                   '   and m.kind =''03101'''+@CR+
                   '   and d.kind =''01100'' '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '03105'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '年度業務大分類實銷金額達成率-百貨(個人業績)(已收退金額)'
     set @Remark= '(依請款日)(Formula: 03101 / 百貨目標, 目標資料: ori_xls#Target_Stock_BKind_Personal.Target_AA)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       m.area_year,'+@CR+
                   '       ''00'' as area_month,'+@CR+
                   '       IsNull(m.area, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(m.area_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       case'+@CR+
                   '         when IsNull(d.Target_Amt, 0) = 0 then 0'+@CR+
                   '         else Round(Sum(Isnull(amt, 0)) / d.Target_Amt, 4)'+@CR+
                   '       end as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       3 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '       inner join'+@CR+
                   '         (select area_year as year, area as emp_no, Sum(Isnull(amt, 0)) as Target_Amt'+@CR+
                   '            from '+@TB_RT_Data_tmp+' D'+@TB_Hint+@CR+
                   '           where kind =''01100'''+@CR+
                   '           group by area_year, area'+@CR+
                   '         ) D'+@CR+
                   '          on m.area_year=d.year'+@CR+
                   '         and m.area '+@Collate+'=d.emp_no '+@Collate+@CR+
                   ' where 1=1'+@CR+
                   '   and m.area_year >= ''2013'''+@CR+
                   '   and m.kind =''03101'''+@CR+
                   ' group by m.area_year, m.area, m.area_name, d.Target_Amt'

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end 

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '03201'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月業務大分類實銷金額-3C(個人業績)(已收退金額)'
     set @Remark= '(依請款日, 表身, 含稅, 含加減項, 含折讓, 含退貨, 含服務單, 大分類為 AB)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_pdate_year as area_year,'+@CR+
                   '       chg_sp_pdate_month as area_month,'+@CR+
                   '       IsNull(sp_sales, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(chg_sales_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(Chg_sd_pay_Sale, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslpdt'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and chg_sp_pdate_year >= ''2013'''+@CR+
                   '   and chg_skno_bkind = ''AB'''+@CR+
                   '   and (sd_slip_fg =''2'' Or sd_slip_fg =''C'' Or sd_slip_fg = ''3'')'+@CR+
                   '   and chg_sp_pdate_year is not null'+@CR+
                   -- 2017/08/08 Rickliu 副總指示客編 IT、ZZ 開頭不認列業績，雖不認列業績但應收扔得收
                   '   and Chg_BU_NO Not In (''IT0000'', ''ZZ0000'') '+@CR+
                   ' group by chg_sp_pdate_year, chg_sp_pdate_month, sp_sales, chg_sales_name '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '03202'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月業務大分類銷售金額(已收金額)-3C(個人業績)'
     set @Remark= '(依請款日, 表身, 含稅, 含加減項, 含折讓, 含服務單, 大分類為 AB)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_pdate_year as area_year,'+@CR+
                   '       chg_sp_pdate_month as area_month,'+@CR+
                   '       IsNull(sp_sales, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(chg_sales_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(Chg_sd_pay_Sale, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslpdt'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and chg_sp_pdate_year >= ''2013'''+@CR+
                   '   and chg_skno_bkind = ''AB'''+@CR+
                   '   and (sd_slip_fg =''2'' Or sd_slip_fg =''C'')'+@CR+
                   '   and chg_sp_pdate_year is not null'+@CR+
                   -- 2017/08/08 Rickliu 副總指示客編 IT、ZZ 開頭不認列業績，雖不認列業績但應收扔得收
                   '   and Chg_BU_NO Not In (''IT0000'', ''ZZ0000'') '+@CR+
                   ' group by chg_sp_pdate_year, chg_sp_pdate_month, sp_sales, chg_sales_name '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '03203'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月業務大分類退貨金額(已退金額)-3C(個人業績)'
     set @Remark= '(依請款日, 表身, 含稅, 含加減項, 含折讓, 大分類為 AB)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_pdate_year as area_year,'+@CR+
                   '       chg_sp_pdate_month as area_month,'+@CR+
                   '       IsNull(sp_sales, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(chg_sales_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(Chg_sd_pay_Sale, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslpdt'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and chg_sp_pdate_year >= ''2013'''+@CR+
                   '   and chg_skno_bkind = ''AB'''+@CR+
                   '   and sd_slip_fg =''3'''+@CR+
                   '   and chg_sp_pdate_year is not null'+@CR+
                   -- 2017/08/08 Rickliu 副總指示客編 IT、ZZ 開頭不認列業績，雖不認列業績但應收扔得收
                   '   and Chg_BU_NO Not In (''IT0000'', ''ZZ0000'') '+@CR+
                   ' group by chg_sp_pdate_year, chg_sp_pdate_month, sp_sales, chg_sales_name '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '03204'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月業務大分類實銷金額達成率-3C(個人業績)(已收退金額)'
     set @Remark= '(依請款日)(Formula: 03201 / 3C目標, 目標資料: ori_xls#Target_Stock_BKind_Personal.Target_AB)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       m.area_year,'+@CR+
                   '       m.area_month,'+@CR+
                   '       IsNull(m.area, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(m.area_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       case'+@CR+
                   '         when IsNull(d.amt, 0) = 0 then 0'+@CR+
                   '         else Round(isnull(m.amt, 0) / isnull(d.amt, 0), 4)'+@CR+
                   '       end as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       3 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '       inner join '+@TB_RT_Data_tmp+' D'+@TB_Hint+@CR+
                   '          on m.area_year = d.area_year'+@CR+
                   '         and m.area_month = d.area_month'+@CR+
                   '         and m.area =d.area'+@CR+
                   ' where 1=1'+@CR+
                   '   and m.area_year >= ''2013'''+@CR+
                   '   and m.kind =''03201'''+@CR+
                   '   and d.kind =''01200'' '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '03205'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '年度業務大分類實銷金額達成率-3C(個人業績)(已收退金額)'
     set @Remark= '(依請款日)(Formula: 03201 / 3C目標, 目標資料: ori_xls#Target_Stock_BKind_Personal.Target_AB)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       m.area_year,'+@CR+
                   '       ''00'' as area_month,'+@CR+
                   '       IsNull(m.area, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(m.area_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       case'+@CR+
                   '         when IsNull(d.Target_Amt, 0) = 0 then 0'+@CR+
                   '         else Round(Sum(Isnull(amt, 0)) / d.Target_Amt, 4)'+@CR+
                   '       end as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       3 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '       inner join'+@CR+
                   '         (select area_year as year, area as emp_no, Sum(Isnull(amt, 0)) as Target_Amt'+@CR+
                   '            from '+@TB_RT_Data_tmp+' D'+@TB_Hint+@CR+
                   '           where kind =''01200'''+@CR+
                   '           group by area_year, area'+@CR+
                   '         ) D'+@CR+
                   '          on m.area_year=d.year'+@CR+
                   '         and m.area '+@Collate+'=d.emp_no '+@Collate+@CR+
                   ' where 1=1'+@CR+
                   '   and m.area_year >= ''2013'''+@CR+
                   '   and m.kind =''03201'''+@CR+
                   ' group by m.area_year, m.area, m.area_name, d.Target_Amt'

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end 

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '03301'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月業務大分類實銷金額-化工(個人業績)(已收退金額)'
     set @Remark= '(依請款日, 表身, 含稅, 含加減項, 含折讓, 含退貨, 含服務單, 大分類為 AC)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_pdate_year as area_year,'+@CR+
                   '       chg_sp_pdate_month as area_month,'+@CR+
                   '       IsNull(sp_sales, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(chg_sales_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(Chg_sd_pay_Sale, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslpdt'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and chg_sp_pdate_year >= ''2013'''+@CR+
                   '   and chg_skno_bkind = ''AC'''+@CR+
                   '   and (sd_slip_fg =''2'' Or sd_slip_fg =''C'' Or sd_slip_fg = ''3'')'+@CR+
                   '   and chg_sp_pdate_year is not null'+@CR+
                   -- 2017/08/08 Rickliu 副總指示客編 IT、ZZ 開頭不認列業績，雖不認列業績但應收扔得收
                   '   and Chg_BU_NO Not In (''IT0000'', ''ZZ0000'') '+@CR+
                   ' group by chg_sp_pdate_year, chg_sp_pdate_month, sp_sales, chg_sales_name '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '03302'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月業務大分類銷售金額(已收金額)-化工(個人業績)'
     set @Remark= '(依請款日, 表身, 含稅, 含加減項, 含折讓, 含服務單, 大分類為 AC)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_pdate_year as area_year,'+@CR+
                   '       chg_sp_pdate_month as area_month,'+@CR+
                   '       IsNull(sp_sales, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(chg_sales_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(Chg_sd_pay_Sale, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslpdt'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and chg_sp_pdate_year >= ''2013'''+@CR+
                   '   and chg_skno_bkind = ''AC'''+@CR+
                   '   and (sd_slip_fg =''2'' Or sd_slip_fg =''C'')'+@CR+
                   '   and chg_sp_pdate_year is not null'+@CR+
                   -- 2017/08/08 Rickliu 副總指示客編 IT、ZZ 開頭不認列業績，雖不認列業績但應收扔得收
                   '   and Chg_BU_NO Not In (''IT0000'', ''ZZ0000'') '+@CR+
                   ' group by chg_sp_pdate_year, chg_sp_pdate_month, sp_sales, chg_sales_name '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
    
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '03303'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月業務大分類退貨金額(已退金額)-化工(個人業績)'
     set @Remark= '(依請款日, 表身, 含稅, 含加減項, 含折讓, 大分類為 AC)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_pdate_year as area_year,'+@CR+
                   '       chg_sp_pdate_month as area_month,'+@CR+
                   '       IsNull(sp_sales, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(chg_sales_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(Chg_sd_pay_Sale, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslpdt'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and chg_sp_pdate_year >= ''2013'''+@CR+
                   '   and chg_skno_bkind = ''AC'''+@CR+
                   '   and sd_slip_fg =''3'''+@CR+
                   '   and chg_sp_pdate_year is not null'+@CR+
                   -- 2017/08/08 Rickliu 副總指示客編 IT、ZZ 開頭不認列業績，雖不認列業績但應收扔得收
                   '   and Chg_BU_NO Not In (''IT0000'', ''ZZ0000'') '+@CR+
                   ' group by chg_sp_pdate_year, chg_sp_pdate_month, sp_sales, chg_sales_name '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '03304'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月業務大分類實銷金額達成率-化工(個人業績)(已收退金額)'
     set @Remark= '(依請款日)(Formula: 03301 / 化工目標, 目標資料: ori_xls#Target_Stock_BKind_Personal.Target_AC)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       m.area_year,'+@CR+
                   '       m.area_month,'+@CR+
                   '       IsNull(m.area, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(m.area_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       case'+@CR+
                   '         when IsNull(d.amt, 0) = 0 then 0'+@CR+
                   '         else Round(isnull(m.amt, 0) / isnull(d.amt, 0), 4)'+@CR+
                   '       end as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       3 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '       inner join '+@TB_RT_Data_tmp+' D'+@TB_Hint+@CR+
                   '          on m.area_year = d.area_year'+@CR+
                   '         and m.area_month = d.area_month'+@CR+
                   '         and m.area =d.area'+@CR+
                   ' where 1=1'+@CR+
                   '   and m.area_year >= ''2013'''+@CR+
                   '   and m.kind =''03301'''+@CR+
                   '   and d.kind =''01300'' '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '03305'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '年度業務大分類實銷金額達成率-化工(個人業績)(已收退金額)'
     set @Remark= '(依請款日)(Formula: 03301 / 化工目標, 目標資料: ori_xls#Target_Stock_BKind_Personal.AC)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       m.area_year,'+@CR+
                   '       ''00'' as area_month,'+@CR+
                   '       IsNull(m.area, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(m.area_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       case'+@CR+
                   '         when IsNull(d.Target_Amt, 0) = 0 then 0'+@CR+
                   '         else Round(Sum(Isnull(amt, 0)) / d.Target_Amt, 4)'+@CR+
                   '       end as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       3 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '       inner join'+@CR+
                   '         (select area_year as year, area as emp_no, Sum(Isnull(amt, 0)) as Target_Amt'+@CR+
                   '            from '+@TB_RT_Data_tmp+' D'+@TB_Hint+@CR+
                   '           where kind =''01300'''+@CR+
                   '           group by area_year, area'+@CR+
                   '         ) D'+@CR+
                   '          on m.area_year=d.year'+@CR+
                   '         and m.area '+@Collate+'=d.emp_no '+@Collate+@CR+
                   ' where 1=1'+@CR+
                   '   and m.area_year >= ''2013'''+@CR+
                   '   and m.kind =''03301'''+@CR+
                   ' group by m.area_year, m.area, m.area_name, d.Target_Amt'

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end 

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '03401'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月業務大分類實銷金額-工具(個人業績)(已收退金額)'
     set @Remark= '(依請款日, 表身, 含稅, 含加減項, 含折讓, 含退貨, 含服務單, 大分類為 AD)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_pdate_year as area_year,'+@CR+
                   '       chg_sp_pdate_month as area_month,'+@CR+
                   '       IsNull(sp_sales, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(chg_sales_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(Chg_sd_pay_Sale, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslpdt'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and chg_sp_pdate_year >= ''2013'''+@CR+
                   '   and chg_skno_bkind = ''AD'''+@CR+
                   '   and (sd_slip_fg =''2'' Or sd_slip_fg =''C'' Or sd_slip_fg = ''3'')'+@CR+
                   '   and chg_sp_pdate_year is not null'+@CR+
                   -- 2017/08/08 Rickliu 副總指示客編 IT、ZZ 開頭不認列業績，雖不認列業績但應收扔得收
                   '   and Chg_BU_NO Not In (''IT0000'', ''ZZ0000'') '+@CR+
                   ' group by chg_sp_pdate_year, chg_sp_pdate_month, sp_sales, chg_sales_name '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '03402'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月業務大分類銷售金額(已收金額)-工具(個人業績)'
     set @Remark= '(依請款日, 表身, 含稅, 含加減項, 含折讓, 含服務單, 大分類為 AD)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_pdate_year as area_year,'+@CR+
                   '       chg_sp_pdate_month as area_month,'+@CR+
                   '       IsNull(sp_sales, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(chg_sales_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(Chg_sd_pay_Sale, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslpdt'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and chg_sp_pdate_year >= ''2013'''+@CR+
                   '   and chg_skno_bkind = ''AD'''+@CR+
                   '   and (sd_slip_fg =''2'' Or sd_slip_fg =''C'')'+@CR+
                   '   and chg_sp_pdate_year is not null'+@CR+
                   -- 2017/08/08 Rickliu 副總指示客編 IT、ZZ 開頭不認列業績，雖不認列業績但應收扔得收
                   '   and Chg_BU_NO Not In (''IT0000'', ''ZZ0000'') '+@CR+
                   ' group by chg_sp_pdate_year, chg_sp_pdate_month, sp_sales, chg_sales_name '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
    
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '03403'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月業務大分類退貨金額(已退金額)-工具(個人業績)'
     set @Remark= '(依請款日, 表身, 含稅, 含加減項, 含折讓, 大分類為 AD)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_pdate_year as area_year,'+@CR+
                   '       chg_sp_pdate_month as area_month,'+@CR+
                   '       IsNull(sp_sales, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(chg_sales_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(chg_sd_stot_tax, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslpdt'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and chg_sp_pdate_year >= ''2013'''+@CR+
                   '   and chg_skno_bkind = ''AD'''+@CR+
                   '   and sd_slip_fg =''3'''+@CR+
                   '   and chg_sp_pdate_year is not null'+@CR+
                   -- 2017/08/08 Rickliu 副總指示客編 IT、ZZ 開頭不認列業績，雖不認列業績但應收扔得收
                   '   and Chg_BU_NO Not In (''IT0000'', ''ZZ0000'') '+@CR+
                   ' group by chg_sp_pdate_year, chg_sp_pdate_month, sp_sales, chg_sales_name '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '03404'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月業務大分類實銷金額達成率-工具(個人業績)(已收退金額)'
     set @Remark= '(依請款日)(Formula: 03401 / 工具目標, 目標資料: ori_xls#Target_Stock_BKind_Personal.Target_AD'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       m.area_year,'+@CR+
                   '       m.area_month,'+@CR+
                   '       IsNull(m.area, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(m.area_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       case'+@CR+
                   '         when IsNull(d.amt, 0) = 0 then 0'+@CR+
                   '         else Round(isnull(m.amt, 0) / isnull(d.amt, 0), 4)'+@CR+
                   '       end as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       3 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '       inner join '+@TB_RT_Data_tmp+' D'+@TB_Hint+@CR+
                   '          on m.area_year = d.area_year'+@CR+
                   '         and m.area_month = d.area_month'+@CR+
                   '         and m.area =d.area'+@CR+
                   ' where 1=1'+@CR+
                   '   and m.area_year >= ''2013'''+@CR+
                   '   and m.kind =''03401'''+@CR+
                   '   and d.kind =''01400'' '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '03405'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '年度業務大分類實銷金額達成率-工具(個人業績)(已收退金額)'
     set @Remark= '(依請款日)(Formula: 03401 / 工具目標, 目標資料: ori_xls#Target_Stock_BKind_Personal.Target_AD)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       m.area_year,'+@CR+
                   '       ''00'' as area_month,'+@CR+
                   '       IsNull(m.area, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(m.area_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       case'+@CR+
                   '         when IsNull(d.Target_Amt, 0) = 0 then 0'+@CR+
                   '         else Round(Sum(Isnull(amt, 0)) / d.Target_Amt, 4)'+@CR+
                   '       end as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       3 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '       inner join'+@CR+
                   '         (select area_year as year, area as emp_no, Sum(Isnull(amt, 0)) as Target_Amt'+@CR+
                   '            from '+@TB_RT_Data_tmp+' D'+@TB_Hint+@CR+
                   '           where kind =''01400'''+@CR+
                   '           group by area_year, area'+@CR+
                   '         ) D'+@CR+
                   '          on m.area_year=d.year'+@CR+
                   '         and m.area '+@Collate+'=d.emp_no '+@Collate+@CR+
                   ' where 1=1'+@CR+
                   '   and m.area_year >= ''2013'''+@CR+
                   '   and m.kind =''03401'''+@CR+
                   ' group by m.area_year, m.area, m.area_name, d.Target_Amt'

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end 
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '03501'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月業務大分類實銷金額-電裝(個人業績)(已收退金額)'
     set @Remark= '(依請款日, 表身, 含稅, 含加減項, 含折讓, 含退貨, 含服務單, 大分類為 AE)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_pdate_year as area_year,'+@CR+
                   '       chg_sp_pdate_month as area_month,'+@CR+
                   '       IsNull(sp_sales, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(chg_sales_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(Chg_sd_pay_Sale, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslpdt'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and chg_sp_pdate_year >= ''2013'''+@CR+
                   '   and chg_skno_bkind = ''AE'''+@CR+
                   '   and (sd_slip_fg =''2'' Or sd_slip_fg =''C'' Or sd_slip_fg = ''3'')'+@CR+
                   '   and chg_sp_pdate_year is not null'+@CR+
                   -- 2017/08/08 Rickliu 副總指示客編 IT、ZZ 開頭不認列業績，雖不認列業績但應收扔得收
                   '   and Chg_BU_NO Not In (''IT0000'', ''ZZ0000'') '+@CR+
                   ' group by chg_sp_pdate_year, chg_sp_pdate_month, sp_sales, chg_sales_name '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '03502'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月業務大分類銷售金額(已收金額)-電裝(個人業績)'
     set @Remark= '(依請款日, 表身, 含稅, 含加減項, 含折讓, 含服務單, 大分類為 AE)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_pdate_year as area_year,'+@CR+
                   '       chg_sp_pdate_month as area_month,'+@CR+
                   '       IsNull(sp_sales, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(chg_sales_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(Chg_sd_pay_Sale, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslpdt'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and chg_sp_pdate_year >= ''2013'''+@CR+
                   '   and chg_skno_bkind = ''AE'''+@CR+
                   '   and (sd_slip_fg =''2'' Or sd_slip_fg =''C'')'+@CR+
                   '   and chg_sp_pdate_year is not null'+@CR+
                   -- 2017/08/08 Rickliu 副總指示客編 IT、ZZ 開頭不認列業績，雖不認列業績但應收扔得收
                   '   and Chg_BU_NO Not In (''IT0000'', ''ZZ0000'') '+@CR+
                   ' group by chg_sp_pdate_year, chg_sp_pdate_month, sp_sales, chg_sales_name '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
    
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '03503'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月業務大分類退貨金額(已退金額)-電裝(個人業績)'
     set @Remark= '(依請款日, 表身, 含稅, 含加減項, 含折讓, 大分類為 AE)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_pdate_year as area_year,'+@CR+
                   '       chg_sp_pdate_month as area_month,'+@CR+
                   '       IsNull(sp_sales, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(chg_sales_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(Chg_sd_pay_Sale, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslpdt'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and chg_sp_pdate_year >= ''2013'''+@CR+
                   '   and chg_skno_bkind = ''AE'''+@CR+
                   '   and sd_slip_fg =''3'''+@CR+
                   '   and chg_sp_pdate_year is not null'+@CR+
                   -- 2017/08/08 Rickliu 副總指示客編 IT、ZZ 開頭不認列業績，雖不認列業績但應收扔得收
                   '   and Chg_BU_NO Not In (''IT0000'', ''ZZ0000'') '+@CR+
                   ' group by chg_sp_pdate_year, chg_sp_pdate_month, sp_sales, chg_sales_name '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '03504'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月業務大分類實銷金額達成率-電裝(個人業績)(已收退金額)'
     set @Remark= '(依請款日)(Formula: 03501 / 工具目標, 目標資料: ori_xls#Target_Stock_BKind_Personal.Target_AE)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       m.area_year,'+@CR+
                   '       m.area_month,'+@CR+
                   '       IsNull(m.area, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(m.area_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       case'+@CR+
                   '         when IsNull(d.amt, 0) = 0 then 0'+@CR+
                   '         else Round(isnull(m.amt, 0) / isnull(d.amt, 0), 4)'+@CR+
                   '       end as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       3 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '       inner join '+@TB_RT_Data_tmp+' D'+@TB_Hint+@CR+
                   '          on m.area_year = d.area_year'+@CR+
                   '         and m.area_month = d.area_month'+@CR+
                   '         and m.area =d.area'+@CR+
                   ' where 1=1'+@CR+
                   '   and m.area_year >= ''2013'''+@CR+
                   '   and m.kind =''03501'''+@CR+
                   '   and d.kind =''01500'' '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '03505'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '年度業務大分類實銷金額達成率-電裝(個人業績)(已收退金額)'
     set @Remark= '(依請款日)(Formula: 03501 / 工具目標, 目標資料: ori_xls#Target_Stock_BKind_Personal.Target_AE)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
                   
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       m.area_year,'+@CR+
                   '       ''00'' as area_month,'+@CR+
                   '       IsNull(m.area, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(m.area_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       case'+@CR+
                   '         when IsNull(d.Target_Amt, 0) = 0 then 0'+@CR+
                   '         else Round(Sum(Isnull(amt, 0)) / d.Target_Amt, 4)'+@CR+
                   '       end as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       3 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '       inner join'+@CR+
                   '         (select area_year as year, area as emp_no, Sum(Isnull(amt, 0)) as Target_Amt'+@CR+
                   '            from '+@TB_RT_Data_tmp+' D'+@TB_Hint+@CR+
                   '           where kind =''01500'''+@CR+
                   '           group by area_year, area'+@CR+
                   '         ) D'+@CR+
                   '          on m.area_year=d.year'+@CR+
                   '         and m.area '+@Collate+'=d.emp_no '+@Collate+@CR+
                   ' where 1=1'+@CR+
                   '   and m.area_year >= ''2013'''+@CR+
                   '   and m.kind =''03501'''+@CR+
                   ' group by m.area_year, m.area, m.area_name, d.Target_Amt'

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end 
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '03901'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月業務大分類實銷金額-其他(個人業績)(已收退金額)'
     set @Remark= '(依請款日, 表身, 含稅, 含加減項, 含折讓, 含退貨, 含服務單, 大分類扣除 百貨, 3C, 化工, 工具, 電裝及其他 以外的)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_pdate_year as area_year,'+@CR+
                   '       chg_sp_pdate_month as area_month,'+@CR+
                   '       IsNull(sp_sales, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(chg_sales_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(Chg_sd_pay_Sale, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslpdt'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and chg_sp_pdate_year >= ''2013'''+@CR+
                   '   and (chg_skno_bkind <> ''AA'' and chg_skno_bkind <> ''AB'' and chg_skno_bkind <> ''AC'' and'+@CR+
                   '        chg_skno_bkind <> ''AD'' and chg_skno_bkind <> ''AE'')'+@CR+
                   '   and (sd_slip_fg =''2'' Or sd_slip_fg =''C'' Or sd_slip_fg = ''3'')'+@CR+
                   '   and chg_sp_pdate_year is not null'+@CR+
                   -- 2017/08/08 Rickliu 副總指示客編 IT、ZZ 開頭不認列業績，雖不認列業績但應收扔得收
                   '   and Chg_BU_NO Not In (''IT0000'', ''ZZ0000'') '+@CR+
                   ' group by chg_sp_pdate_year, chg_sp_pdate_month, sp_sales, chg_sales_name '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '03902'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月業務大分類銷售金額(已收金額)-其他(個人業績)'
     set @Remark= '(依請款日, 表身, 含稅, 含加減項, 含折讓, 含服務單, 大分類扣除 百貨, 3C, 化工, 工具, 電裝及其他 以外的)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_pdate_year as area_year,'+@CR+
                   '       chg_sp_pdate_month as area_month,'+@CR+
                   '       IsNull(sp_sales, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(chg_sales_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(Chg_sd_pay_Sale, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslpdt'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and chg_sp_pdate_year >= ''2013'''+@CR+
                   '   and (chg_skno_bkind <> ''AA'' and chg_skno_bkind <> ''AB'' and chg_skno_bkind <> ''AC'' and'+@CR+
                   '        chg_skno_bkind <> ''AD'' and chg_skno_bkind <> ''AE'')'+@CR+
                   '   and (sd_slip_fg =''2'' Or sd_slip_fg =''C'')'+@CR+
                   '   and chg_sp_pdate_year is not null'+@CR+
                   -- 2017/08/08 Rickliu 副總指示客編 IT、ZZ 開頭不認列業績，雖不認列業績但應收扔得收
                   '   and Chg_BU_NO Not In (''IT0000'', ''ZZ0000'') '+@CR+
                   ' group by chg_sp_pdate_year, chg_sp_pdate_month, sp_sales, chg_sales_name '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Kind = '03903'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     set @Kind_Name = '每月業務大分類退貨金額(已退金額)-其他(個人業績)'
     set @Remark= '(依請款日, 表身, 含稅, 含加減項, 含折讓, 大分類扣除 百貨, 3C, 化工, 工具, 電裝及其他 以外的)'
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strWhere ='Where kind = '''+@Kind+''' '
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' '+@strWhere
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+''' '+@Collate+' as kind,'+@CR+
                   '       '''+@Kind+'.'+@Kind_Name+''' '+@Collate+' as kind_name,'+@CR+
                   '       chg_sp_pdate_year as area_year,'+@CR+
                   '       chg_sp_pdate_month as area_month,'+@CR+
                   '       IsNull(sp_sales, '''+@NonSet+''') as area,'+@CR+
                   '       IsNull(chg_sales_name, '''+@CNonSet+''') as area_name,'+@CR+
                   '       Sum(IsNull(Chg_sd_pay_Sale, 0)) as amt,'+@CR+
                   '       '''' as Not_Accumulate,'+@CR+
                   '       0 as Data_Type,'+@CR+
                   '       '''+@ReMark+''' as Remark'+@CR+
                   '  from fact_sslpdt'+@TB_Hint+@CR+
                   ' where 1=1'+@CR+
                   '   and chg_sp_pdate_year >= ''2013'''+@CR+
                   '   and (chg_skno_bkind <> ''AA'' and chg_skno_bkind <> ''AB'' and chg_skno_bkind <> ''AC'' and'+@CR+
                   '        chg_skno_bkind <> ''AD'' and chg_skno_bkind <> ''AE'')'+@CR+
                   '   and sd_slip_fg =''3'''+@CR+
                   '   and chg_sp_pdate_year is not null'+@CR+
                   -- 2017/08/08 Rickliu 副總指示客編 IT、ZZ 開頭不認列業績，雖不認列業績但應收扔得收
                   '   and Chg_BU_NO Not In (''IT0000'', ''ZZ0000'') '+@CR+
                   ' group by chg_sp_pdate_year, chg_sp_pdate_month, sp_sales, chg_sales_name '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--
-- 自動產生累計資料
--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--
  set @Kind = 'C'
  if (@in_Kind not like '[A-Z]%' Or @in_Kind ='')
  begin
     set @Kind_Name = '自動產生累計資料'
     
     Set @Msg = @Kind+'.刪除 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'Delete '+@TB_RT_Data_tmp+' where Kind like ''C'+Substring(@in_kind, 2, 4)+'%'' '
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
     
     Set @Msg = @Kind+'.新增 '+@Kind_Name+' 資料'
     set @strSQL = @SP_WTL_TB_RT_Data_tmp+@CR+
                   'insert into '+@TB_RT_Data_tmp+@CR+
                   'select '''+@Kind+'''+Substring(m.kind, 2, 4) as kind,'+@CR+
                   '       '''+@Kind+'''+Substring(m.kind_name, 2, Len(m.kind_name))+''(由 ''+m.kind+'' 累計)'' as kind_name,'+@CR+
                   '       m.area_year,'+@CR+
                   '       m.area_month,'+@CR+
                   '       m.area,'+@CR+ 
                   '       m.area_name,'+@CR+
                   '       Round(Sum(Isnull(d.amt, 0)), 4) as amt,'+@CR+
                   '       m.Not_Accumulate,'+@CR+
                   '       m.Data_Type,'+@CR+
                   '       m.Remark'+@CR+
                   '  from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '       inner join '+@TB_RT_Data_tmp+' d'+@TB_Hint+@CR+
                   '          on m.kind = d.kind and m.area = d.area'+@CR+
                   ' where 1=1'+@CR+
                   '   and m.area_year = d.area_year'+@CR+
                   '   and d.area_month <= m.area_month'+@CR+
                   '   and d.area_month <> 0'+@CR+
                   '   and m.kind like '''+@in_kind+'%'''+@CR+
                   -- 針對已經累計的資料就不進行累計加總
                   '   and m.Not_Accumulate = '''''+@CR+
                   -- 若資料只有一筆的就不進行累計
                   -- 2015/03/16 由於資料會累計成長，又加上使用 not exists 以及 distinct 方式 會直接使用到 tempdb，導致效能變差，
                   -- 因此直接取消使用以下語法。
                   /*
                   '   and not exists'+@CR+
                   '       (select distinct d1.kind, count(*) as cnt'+@CR+
                   '          from '+@TB_RT_Data_tmp+' d1'+@CR+
                   '         where m.kind = d1.kind'+@CR+
                   '           and d1.kind like '''+@in_kind+'%'''+@CR+
                   '         group by d1.kind'+@CR+
                   '        having count(*) = 1'+@CR+
                   '       )'+@CR+
                   */
                   ' group by m.kind, m.kind_name, m.area_year, m.area_month, m.area, m.area_name, m.amt, m.Not_Accumulate, m.Data_Type, m.Remark '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end 

--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--
--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--
  if (@in_Kind <> '')
  begin
     set @Msg = '刪除 '+@TB_RT_Data+' ['+@in_Kind+'] 資料'
     set @strSQL = 'Delete '+@TB_RT_Data+' Where Kind like '''+@in_kind+'%'' or Kind like ''[A-Z]'+Substring(@in_kind, 2, 4)+'%'' '
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
  else
  begin
     Set @Msg = '重整 ['+@TB_RT_Data_tmp+'] 將 Null 值補入初始值!!'
     set @strSQL = 'Update '+@TB_RT_Data_tmp+@CR+
                   '   set area_month ='+@CR+
                   '         case'+@CR+
                   '           when len(IsNull(area_month, '''')) = 0 then ''00'''+@CR+
                   '           else area_month'+@CR+
                   '         end,'+@CR+
                   '       area ='+@CR+
                   '         case'+@CR+
                   '           when len(IsNull(area_month, '''')) = 0 then '''+@NonSet+''''+@CR+
                   '           else Rtrim(LTrim(area))'+@CR+ 
                   '         end,'+@CR+
                   '       area_name ='+@CR+ 
                   '         case'+@CR+
                   '           when len(IsNull(area_name, '''')) = 0 then '''+@CNonSet+''''+@CR+
                   '           else Rtrim(LTrim(area_name))'+@CR+
                   '         end '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = '重建資料表 ['+@TB_RT_Data+']'
     set @strSQL = 'IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N''[dbo].['+@TB_RT_Data+']'') AND type in (N''U''))'+@CR+
                   '   Drop Table '+@TB_RT_Data+@CR+
                   'IF Not EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N''[dbo].['+@TB_RT_Data+']'') AND type in (N''U''))'+@CR+
                   '   CREATE TABLE [dbo].['+@TB_RT_Data+']('+@CR+
                   '     [kind] [varchar](10) NOT NULL,'+@CR+
                   '     [kind_name] [varchar](1000) NULL,'+@CR+
                   '     [Area_year] [int] NOT NULL,'+@CR+
                   '     [Area_month] [int] NOT NULL,'+@CR+
                   '     [area] [Varchar](20) NOT NULL,'+@CR+
                   '     [Area_name] [Varchar](50) NULL,'+@CR+
                   '     [amt] [float] NULL,'+@CR+
                   '     [Not_Accumulate] [Varchar] (1) NULL,'+@CR+
                   '     [Data_Type] [int] NULL,'+@CR+
                   '     [Remark] [Varchar] (Max) NULL'+@CR+
                   '     Primary Key (Kind, area_year, area_month, area)'+@CR+
                   '   )'

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
     
     set @strWhere = ''
  end

  if @in_Kind = ''
  begin
     Set @Msg = '比對並刪除異常資料 ['+@TB_RT_Data_tmp+'] 至 ['+@TB_RT_Data+'] '
     set @strSQL = 'delete from '+@TB_RT_Data+@CR+
                   ' where kind in'+@CR+
                   '      (Select distinct kind'+@CR+
                   '         from '+@TB_RT_Data_tmp+@TB_Hint+@CR+
                   '      ) '
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
  else 
  begin
     Set @Msg = '刪除 ['+@in_Kind+'] 類資料 ['+@TB_RT_Data_tmp+'] 至 ['+@TB_RT_Data+'] '
     set @strSQL = 'Delete from '+@TB_RT_Data+@CR+
                   'where kind in'+@CR+
                   '      (Select distinct kind'+@CR+
                   '         from '+@TB_RT_Data_tmp+@TB_Hint+@CR+
                   '        where kind like '''+@in_kind+'%'''+@CR+
                   '           or Kind like ''[A-Z]'+Substring(@in_kind, 2, 4)+'%'' ' +@CR+
                   '      ) '
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
     
  if @in_Kind Not like ('[A-Z]%')
  begin
     Set @Msg = '匯入 ['+@in_Kind+'] 類資料 ['+@TB_RT_Data_tmp+'] 至 ['+@TB_RT_Data+'] '
     set @strSQL = 'Insert Into '+@TB_RT_Data+@CR+
                   'Select * from '+@TB_RT_Data_tmp+@TB_Hint+@CR+
                   ' where kind like '''+@in_kind+'%'''+@CR+
                   '    or Kind like ''[A-Z]'+Substring(@in_kind, 2, 4)+'%'' '
                   
     set @strSQL = @strSQL +@CR+
                   ' Order by 1, 3, 4, 5 '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end
--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--
--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--
  set @Kind = 'A'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     -- 百貨
     Set @Msg = @Kind+'.刪除 年月業務大分類(百貨)實銷達成率排名暫存檔 資料 [Sale_Member_BKind_Rank_AA_By_YM_tmp]'
     set @strSQL = 'IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N''[dbo].[Sale_Member_BKind_Rank_AA_By_YM_tmp]'') AND type in (N''U''))'+@CR+
                   '   Drop Table Sale_Member_BKind_Rank_AA_By_YM_tmp '
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
     
     Set @Msg = @Kind+'.新增 年月業務大分類(百貨)實銷達成率排名暫存檔 資料 [Sale_Member_BKind_Rank_AA_By_YM_tmp]'

	 set @strSQL = ' select Row_Number() Over(Partition BY area_year, area_month order by area_year, area_month,'+@CR+
                   '                          case'+@CR+
                   '                            when IsNull(Target_Amt, 0) = 0 then 0'+@CR+
                   '                            else Round(Sum(IsNull(amt, 0)) / Target_Amt, 4)'+@CR+
                   '                          end desc) as Rank,'+@CR+
                   '        area_year as year,'+@CR+
                   '        area_month as month,'+@CR+
                   '        Rtrim(area) as emp_no,'+@CR+
                   '        Rtrim(area_name) as emp_name,'+@CR+
                   '        Round(Sum(IsNull(amt, 0)), 0) as sale_amt,'+@CR+
                   '        Target_Amt as target_amt,'+@CR+
                   '        case'+@CR+
                   '          when IsNull(Target_Amt, 0) = 0 then 0'+@CR+
                   '          else Round(Sum(Isnull(amt, 0)) / Target_Amt, 4)'+@CR+
                   '        end as rate'+@CR+
                   '       into Sale_Member_BKind_Rank_AA_By_YM_tmp'+@CR+
                   '   from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '        inner join'+@CR+
                   '          (select area_year as year, area_month as month, area as emp_no, Sum(IsNull(amt, 0)) as Target_Amt'+@CR+
                   '             from '+@TB_RT_Data_tmp+@TB_Hint+@CR+
                   '            where kind=''01100'''+@CR+
                   '            group by area_year, area_month, area'+@CR+
                   '          ) D'+@CR+
                   '           on m.area_year=d.year'+@CR+
                   '          and m.area_month=d.month'+@CR+
                   '          and m.area '+@Collate+'=d.emp_no '+@Collate+@CR+
                   '  where 1=1'+@CR+
                   '    and area_year >= ''2013'''+@CR+
                   '    and kind =''01101'''+@CR+
                   '  group by area_year, area_month, area, area_name, Target_Amt '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
	 -- 3C
     Set @Msg = @Kind+'.刪除 年月業務大分類(3C)實銷達成率排名暫存檔 資料 [Sale_Member_BKind_Rank_AB_By_YM_tmp]'
     set @strSQL = 'IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N''[dbo].[Sale_Member_BKind_Rank_AB_By_YM_tmp]'') AND type in (N''U''))'+@CR+
                   '   Drop Table Sale_Member_BKind_Rank_AB_By_YM_tmp '
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
     
     Set @Msg = @Kind+'.新增 年月業務大分類(3C)實銷達成率排名暫存檔 資料 [Sale_Member_BKind_Rank_AB_By_YM_tmp]'

	 set @strSQL = ' select Row_Number() Over(Partition BY area_year, area_month order by area_year, area_month,'+@CR+
                   '                          case'+@CR+
                   '                            when IsNull(Target_Amt, 0) = 0 then 0'+@CR+
                   '                            else Round(Sum(IsNull(amt, 0)) / Target_Amt, 4)'+@CR+
                   '                          end desc) as Rank,'+@CR+
                   '        area_year as year,'+@CR+
                   '        area_month as month,'+@CR+
                   '        Rtrim(area) as emp_no,'+@CR+
                   '        Rtrim(area_name) as emp_name,'+@CR+
                   '        Round(Sum(IsNull(amt, 0)), 0) as sale_amt,'+@CR+
                   '        Target_Amt as target_amt,'+@CR+
                   '        case'+@CR+
                   '          when IsNull(Target_Amt, 0) = 0 then 0'+@CR+
                   '          else Round(Sum(Isnull(amt, 0)) / Target_Amt, 4)'+@CR+
                   '        end as rate'+@CR+
                   '       into Sale_Member_BKind_Rank_AB_By_YM_tmp'+@CR+
                   '   from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '        inner join'+@CR+
                   '          (select area_year as year, area_month as month, area as emp_no, Sum(IsNull(amt, 0)) as Target_Amt'+@CR+
                   '             from '+@TB_RT_Data_tmp+@TB_Hint+@CR+
                   '            where kind=''01200'''+@CR+
                   '            group by area_year, area_month, area'+@CR+
                   '          ) D'+@CR+
                   '           on m.area_year=d.year'+@CR+
                   '          and m.area_month=d.month'+@CR+
                   '          and m.area '+@Collate+'=d.emp_no '+@Collate+@CR+
                   '  where 1=1'+@CR+
                   '    and area_year >= ''2013'''+@CR+
                   '    and kind =''01201'''+@CR+
                   '  group by area_year, area_month, area, area_name, Target_Amt '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
	 -- 化工
     Set @Msg = @Kind+'.刪除 年月業務大分類(化工)實銷達成率排名暫存檔 資料 [Sale_Member_BKind_Rank_AC_By_YM_tmp]'
     set @strSQL = 'IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N''[dbo].[Sale_Member_BKind_Rank_AC_By_YM_tmp]'') AND type in (N''U''))'+@CR+
                   '   Drop Table Sale_Member_BKind_Rank_AC_By_YM_tmp '
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
     
     Set @Msg = @Kind+'.新增 年月業務大分類(化工)實銷達成率排名暫存檔 資料 [Sale_Member_BKind_Rank_AC_By_YM_tmp]'

	 set @strSQL = ' select Row_Number() Over(Partition BY area_year, area_month order by area_year, area_month,'+@CR+
                   '                          case'+@CR+
                   '                            when IsNull(Target_Amt, 0) = 0 then 0'+@CR+
                   '                            else Round(Sum(IsNull(amt, 0)) / Target_Amt, 4)'+@CR+
                   '                          end desc) as Rank,'+@CR+
                   '        area_year as year,'+@CR+
                   '        area_month as month,'+@CR+
                   '        Rtrim(area) as emp_no,'+@CR+
                   '        Rtrim(area_name) as emp_name,'+@CR+
                   '        Round(Sum(IsNull(amt, 0)), 0) as sale_amt,'+@CR+
                   '        Target_Amt as target_amt,'+@CR+
                   '        case'+@CR+
                   '          when IsNull(Target_Amt, 0) = 0 then 0'+@CR+
                   '          else Round(Sum(Isnull(amt, 0)) / Target_Amt, 4)'+@CR+
                   '        end as rate'+@CR+
                   '       into Sale_Member_BKind_Rank_AC_By_YM_tmp'+@CR+
                   '   from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '        inner join'+@CR+
                   '          (select area_year as year, area_month as month, area as emp_no, Sum(IsNull(amt, 0)) as Target_Amt'+@CR+
                   '             from '+@TB_RT_Data_tmp+@TB_Hint+@CR+
                   '            where kind=''01300'''+@CR+
                   '            group by area_year, area_month, area'+@CR+
                   '          ) D'+@CR+
                   '           on m.area_year=d.year'+@CR+
                   '          and m.area_month=d.month'+@CR+
                   '          and m.area '+@Collate+'=d.emp_no '+@Collate+@CR+
                   '  where 1=1'+@CR+
                   '    and area_year >= ''2013'''+@CR+
                   '    and kind =''01301'''+@CR+
                   '  group by area_year, area_month, area, area_name, Target_Amt '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
	 -- 工具
     Set @Msg = @Kind+'.刪除 年月業務大分類(工具)實銷達成率排名暫存檔 資料 [Sale_Member_BKind_Rank_AD_By_YM_tmp]'
     set @strSQL = 'IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N''[dbo].[Sale_Member_BKind_Rank_AD_By_YM_tmp]'') AND type in (N''U''))'+@CR+
                   '   Drop Table Sale_Member_BKind_Rank_AD_By_YM_tmp '
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
     
     Set @Msg = @Kind+'.新增 年月業務大分類(工具)實銷達成率排名暫存檔 資料 [Sale_Member_BKind_Rank_AD_By_YM_tmp]'

	 set @strSQL = ' select Row_Number() Over(Partition BY area_year, area_month order by area_year, area_month,'+@CR+
                   '                          case'+@CR+
                   '                            when IsNull(Target_Amt, 0) = 0 then 0'+@CR+
                   '                            else Round(Sum(IsNull(amt, 0)) / Target_Amt, 4)'+@CR+
                   '                          end desc) as Rank,'+@CR+
                   '        area_year as year,'+@CR+
                   '        area_month as month,'+@CR+
                   '        Rtrim(area) as emp_no,'+@CR+
                   '        Rtrim(area_name) as emp_name,'+@CR+
                   '        Round(Sum(IsNull(amt, 0)), 0) as sale_amt,'+@CR+
                   '        Target_Amt as target_amt,'+@CR+
                   '        case'+@CR+
                   '          when IsNull(Target_Amt, 0) = 0 then 0'+@CR+
                   '          else Round(Sum(Isnull(amt, 0)) / Target_Amt, 4)'+@CR+
                   '        end as rate'+@CR+
                   '       into Sale_Member_BKind_Rank_AD_By_YM_tmp'+@CR+
                   '   from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '        inner join'+@CR+
                   '          (select area_year as year, area_month as month, area as emp_no, Sum(IsNull(amt, 0)) as Target_Amt'+@CR+
                   '             from '+@TB_RT_Data_tmp+@TB_Hint+@CR+
                   '            where kind=''01400'''+@CR+
                   '            group by area_year, area_month, area'+@CR+
                   '          ) D'+@CR+
                   '           on m.area_year=d.year'+@CR+
                   '          and m.area_month=d.month'+@CR+
                   '          and m.area '+@Collate+'=d.emp_no '+@Collate+@CR+
                   '  where 1=1'+@CR+
                   '    and area_year >= ''2013'''+@CR+
                   '    and kind =''01401'''+@CR+
                   '  group by area_year, area_month, area, area_name, Target_Amt '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
	 -- 電裝
     Set @Msg = @Kind+'.刪除 年月業務大分類(電裝)實銷達成率排名暫存檔 資料 [Sale_Member_BKind_Rank_AE_By_YM_tmp]'
     set @strSQL = 'IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N''[dbo].[Sale_Member_BKind_Rank_AE_By_YM_tmp]'') AND type in (N''U''))'+@CR+
                   '   Drop Table Sale_Member_BKind_Rank_AE_By_YM_tmp '
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
     
     Set @Msg = @Kind+'.新增 年月業務大分類(電裝)實銷達成率排名暫存檔 資料 [Sale_Member_BKind_Rank_AE_By_YM_tmp]'

	 set @strSQL = ' select Row_Number() Over(Partition BY area_year, area_month order by area_year, area_month,'+@CR+
                   '                          case'+@CR+
                   '                            when IsNull(Target_Amt, 0) = 0 then 0'+@CR+
                   '                            else Round(Sum(IsNull(amt, 0)) / Target_Amt, 4)'+@CR+
                   '                          end desc) as Rank,'+@CR+
                   '        area_year as year,'+@CR+
                   '        area_month as month,'+@CR+
                   '        Rtrim(area) as emp_no,'+@CR+
                   '        Rtrim(area_name) as emp_name,'+@CR+
                   '        Round(Sum(IsNull(amt, 0)), 0) as sale_amt,'+@CR+
                   '        Target_Amt as target_amt,'+@CR+
                   '        case'+@CR+
                   '          when IsNull(Target_Amt, 0) = 0 then 0'+@CR+
                   '          else Round(Sum(Isnull(amt, 0)) / Target_Amt, 4)'+@CR+
                   '        end as rate'+@CR+
                   '       into Sale_Member_BKind_Rank_AE_By_YM_tmp'+@CR+
                   '   from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '        inner join'+@CR+
                   '          (select area_year as year, area_month as month, area as emp_no, Sum(IsNull(amt, 0)) as Target_Amt'+@CR+
                   '             from '+@TB_RT_Data_tmp+@TB_Hint+@CR+
                   '            where kind=''01500'''+@CR+
                   '            group by area_year, area_month, area'+@CR+
                   '          ) D'+@CR+
                   '           on m.area_year=d.year'+@CR+
                   '          and m.area_month=d.month'+@CR+
                   '          and m.area '+@Collate+'=d.emp_no '+@Collate+@CR+
                   '  where 1=1'+@CR+
                   '    and area_year >= ''2013'''+@CR+
                   '    and kind =''01501'''+@CR+
                   '  group by area_year, area_month, area, area_name, Target_Amt '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
	 -- 其他
     Set @Msg = @Kind+'.刪除 年月業務大分類(其他)實銷達成率排名暫存檔 資料 [Sale_Member_BKind_Rank_OT_By_YM_tmp]'
     set @strSQL = 'IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N''[dbo].[Sale_Member_BKind_Rank_OT_By_YM_tmp]'') AND type in (N''U''))'+@CR+
                   '   Drop Table Sale_Member_BKind_Rank_OT_By_YM_tmp '
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
     
     Set @Msg = @Kind+'.新增 年月業務大分類(其他)實銷達成率排名暫存檔 資料 [Sale_Member_BKind_Rank_OT_By_YM_tmp]'

	 set @strSQL = ' select Row_Number() Over(Partition BY area_year, area_month order by area_year, area_month,'+@CR+
                   '                          case'+@CR+
                   '                            when IsNull(Target_Amt, 0) = 0 then 0'+@CR+
                   '                            else Round(Sum(IsNull(amt, 0)) / Target_Amt, 4)'+@CR+
                   '                          end desc) as Rank,'+@CR+
                   '        area_year as year,'+@CR+
                   '        area_month as month,'+@CR+
                   '        Rtrim(area) as emp_no,'+@CR+
                   '        Rtrim(area_name) as emp_name,'+@CR+
                   '        Round(Sum(IsNull(amt, 0)), 0) as sale_amt,'+@CR+
                   '        Target_Amt as target_amt,'+@CR+
                   '        case'+@CR+
                   '          when IsNull(Target_Amt, 0) = 0 then 0'+@CR+
                   '          else Round(Sum(Isnull(amt, 0)) / Target_Amt, 4)'+@CR+
                   '        end as rate'+@CR+
                   '       into Sale_Member_BKind_Rank_OT_By_YM_tmp'+@CR+
                   '   from '+@TB_RT_Data_tmp+' m'+@TB_Hint+@CR+
                   '        inner join'+@CR+
                   '          (select area_year as year, area_month as month, area as emp_no, Sum(IsNull(amt, 0)) as Target_Amt'+@CR+
                   '             from '+@TB_RT_Data_tmp+@TB_Hint+@CR+
                   '            where kind=''01900'''+@CR+
                   '            group by area_year, area_month, area'+@CR+
                   '          ) D'+@CR+
                   '           on m.area_year=d.year'+@CR+
                   '          and m.area_month=d.month'+@CR+
                   '          and m.area '+@Collate+'=d.emp_no '+@Collate+@CR+
                   '  where 1=1'+@CR+
                   '    and area_year >= ''2013'''+@CR+
                   '    and kind =''01901'''+@CR+
                   '  group by area_year, area_month, area, area_name, Target_Amt '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.刪除 年月業務大分類實銷達成率排名暫存檔 資料 [Sale_Member_BKind_Rank_By_YM_tmp]'
     set @strSQL = 'IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N''[dbo].[Sale_Member_BKind_Rank_By_YM_tmp]'') AND type in (N''U''))'+@CR+
                   '   Drop Table Sale_Member_BKind_Rank_By_YM_tmp '
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
     
     Set @Msg = @Kind+'.新增 年月業務大分類實銷達成率排名暫存檔 資料 [Sale_Member_BKind_Rank_By_YM_tmp]'
     set @strSQL = 'select Rank_AA as Rank, year_AA as year, month_AA as month,'+@CR+
                   '       ''百貨'' as kind_name_AA, emp_no_AA, emp_name_AA, sale_amt_AA, target_amt_AA, rate_AA,'+@CR+
                   '       ''3C''  as kind_name_AB, emp_no_AB, emp_name_AB, sale_amt_AB, target_amt_AB, rate_AB,'+@CR+
                   '       ''化工'' as kind_name_AC, emp_no_AC, emp_name_AC, sale_amt_AC, target_amt_AC, rate_AC,'+@CR+
                   '       ''工具'' as kind_name_AD, emp_no_AD, emp_name_AD, sale_amt_AD, target_amt_AD, rate_AD,'+@CR+
                   '       ''電裝'' as kind_name_AE, emp_no_AE, emp_name_AE, sale_amt_AE, target_amt_AE, rate_AE,'+@CR+
                   '       ''其他'' as kind_name_OT, emp_no_OT, emp_name_OT, sale_amt_OT, target_amt_OT, rate_OT'+@CR+
                   '       into Sale_Member_BKind_Rank_By_YM_tmp'+@CR+
                   '  from (select *'+@CR+
                   '          from (select ROW_NUMBER() Over (order by e_no) as Rank1'+@CR+
                   '                  from Fact_pemploy m'+@TB_Hint+@CR+
                   '                 where Chg_leave=''N'''+@CR+
                   '                   and (E_DEPT = ''B2100'' Or E_DEPT =''B2300'')'+@CR+
                   '               ) as newtable'+@CR+
                   '         where Rank1 <=10'+@CR+
                   '       ) a'+@CR+
                   -- 百貨
                   '       left join'+@CR+
                   '       (select Rank as Rank_AA,'+@CR+
                   '               year as year_AA,'+@CR+
                   '               month as month_AA,'+@CR+
                   '               Rtrim(emp_no) as emp_no_AA,'+@CR+
                   '               Rtrim(emp_name) as emp_name_AA,'+@CR+
                   '               Round(Sum(IsNull(sale_amt, 0)), 0) as sale_amt_AA,'+@CR+
                   '               target_amt as target_amt_AA,'+@CR+
                   '               case'+@CR+
                   '                 when IsNull(Target_Amt, 0) = 0 then 0'+@CR+
                   '                 else Round(Sum(Isnull(sale_amt, 0)) / Target_Amt, 4)'+@CR+
                   '               end as rate_AA'+@CR+
                   '          from Sale_Member_BKind_Rank_AA_By_YM_tmp'+@TB_Hint+@CR+
                   '         group by Rank, year, month, emp_no, emp_name, target_amt'+@CR+
                   '       )m'+@CR+
                   '       on a.Rank1=m.Rank_AA'+@CR+
                   -- 3C
                   '       left join'+@CR+
                   '       (select Rank as Rank_AB,'+@CR+
                   '               year as year_AB,'+@CR+
                   '               month as month_AB,'+@CR+
                   '               Rtrim(emp_no) as emp_no_AB,'+@CR+
                   '               Rtrim(emp_name) as emp_name_AB,'+@CR+
                   '               Round(Sum(IsNull(sale_amt, 0)), 0) as sale_amt_AB,'+@CR+
                   '               target_amt as target_amt_AB,'+@CR+
                   '               case'+@CR+
                   '                 when IsNull(Target_Amt, 0) = 0 then 0'+@CR+
                   '                 else Round(Sum(Isnull(sale_amt, 0)) / Target_Amt, 4)'+@CR+
                   '               end as rate_AB'+@CR+
                   '          from Sale_Member_BKind_Rank_AB_By_YM_tmp'+@TB_Hint+@CR+
                   '         group by Rank, year, month, emp_no, emp_name, target_amt'+@CR+
                   '       )d'+@CR+
                   '        on a.Rank1=d.Rank_AB'+@CR+
                   '       and m.year_AA=d.year_AB'+@CR+
                   '       and m.month_AA=d.month_AB'+@CR+
                   -- 化工
                   '       left join'+@CR+
                   '       (select Rank as Rank_AC,'+@CR+
                   '               year as year_AC,'+@CR+
                   '               month as month_AC,'+@CR+
                   '               Rtrim(emp_no) as emp_no_AC,'+@CR+
                   '               Rtrim(emp_name) as emp_name_AC,'+@CR+
                   '               Round(Sum(IsNull(sale_amt, 0)), 0) as sale_amt_AC,'+@CR+
                   '               target_amt as target_amt_AC,'+@CR+
                   '               case'+@CR+
                   '                 when IsNull(Target_Amt, 0) = 0 then 0'+@CR+
                   '                 else Round(Sum(Isnull(sale_amt, 0)) / Target_Amt, 4)'+@CR+
                   '               end as rate_AC'+@CR+
                   '          from Sale_Member_BKind_Rank_AC_By_YM_tmp'+@TB_Hint+@CR+
                   '         group by Rank, year, month, emp_no, emp_name, target_amt'+@CR+
                   '       )d1'+@CR+
                   '        on a.Rank1=d1.Rank_AC'+@CR+
                   '       and m.year_AA=d1.year_AC'+@CR+
                   '       and m.month_AA=d1.month_AC'+@CR+
                   -- 工具
                   '       left join'+@CR+
                   '       (select Rank as Rank_AD,'+@CR+
                   '               year as year_AD,'+@CR+
                   '               month as month_AD,'+@CR+
                   '               Rtrim(emp_no) as emp_no_AD,'+@CR+
                   '               Rtrim(emp_name) as emp_name_AD,'+@CR+
                   '               Round(Sum(IsNull(sale_amt, 0)), 0) as sale_amt_AD,'+@CR+
                   '               target_amt as target_amt_AD,'+@CR+
                   '               case'+@CR+
                   '                 when IsNull(Target_Amt, 0) = 0 then 0'+@CR+
                   '                 else Round(Sum(Isnull(sale_amt, 0)) / Target_Amt, 4)'+@CR+
                   '               end as rate_AD'+@CR+
                   '          from Sale_Member_BKind_Rank_AD_By_YM_tmp'+@TB_Hint+@CR+
                   '         group by Rank, year, month, emp_no, emp_name, target_amt'+@CR+
                   '       )d2'+@CR+
                   '        on a.Rank1=d2.Rank_AD'+@CR+
                   '       and m.year_AA=d2.year_AD'+@CR+
                   '       and m.month_AA=d2.month_AD'+@CR+
                   -- 電裝
                   '       left join'+@CR+
                   '       (select Rank as Rank_AE,'+@CR+
                   '               year as year_AE,'+@CR+
                   '               month as month_AE,'+@CR+
                   '               Rtrim(emp_no) as emp_no_AE,'+@CR+
                   '               Rtrim(emp_name) as emp_name_AE,'+@CR+
                   '               Round(Sum(IsNull(sale_amt, 0)), 0) as sale_amt_AE,'+@CR+
                   '               target_amt as target_amt_AE,'+@CR+
                   '               case'+@CR+
                   '                 when IsNull(Target_Amt, 0) = 0 then 0'+@CR+
                   '                 else Round(Sum(Isnull(sale_amt, 0)) / Target_Amt, 4)'+@CR+
                   '               end as rate_AE'+@CR+
                   '          from Sale_Member_BKind_Rank_AE_By_YM_tmp'+@TB_Hint+@CR+
                   '         group by Rank, year, month, emp_no, emp_name, target_amt'+@CR+
                   '       )d3'+@CR+
                   '        on a.Rank1=d3.Rank_AE'+@CR+
                   '       and m.year_AA=d3.year_AE'+@CR+
                   '       and m.month_AA=d3.month_AE'+@CR+
                   -- 其他
                   '       left join'+@CR+
                   '       (select Rank as Rank_OT,'+@CR+
                   '               year as year_OT,'+@CR+
                   '               month as month_OT,'+@CR+
                   '               Rtrim(emp_no) as emp_no_OT,'+@CR+
                   '               Rtrim(emp_name) as emp_name_OT,'+@CR+
                   '               Round(Sum(IsNull(sale_amt, 0)), 0) as sale_amt_OT,'+@CR+
                   '               target_amt as target_amt_OT,'+@CR+
                   '               case'+@CR+
                   '                 when IsNull(Target_Amt, 0) = 0 then 0'+@CR+
                   '                 else Round(Sum(Isnull(sale_amt, 0)) / Target_Amt, 4)'+@CR+
                   '               end as rate_OT'+@CR+
                   '          from Sale_Member_BKind_Rank_OT_By_YM_tmp'+@TB_Hint+@CR+
                   '         group by Rank, year, month, emp_no, emp_name, target_amt'+@CR+
                   '       )d4'+@CR+
                   '        on a.Rank1=d4.Rank_OT'+@CR+
                   '       and m.year_AA=d4.year_OT'+@CR+
                   '       and m.month_AA=d4.month_OT '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
     
     Set @Msg = @Kind+'.刪除 年月業務大分類實銷達成率排名 資料 [Sale_Member_BKind_Rank_By_YM]'
     set @strSQL = 'IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N''[dbo].[Sale_Member_BKind_Rank_By_YM]'') AND type in (N''U''))'+@CR+
                   '   Drop Table Sale_Member_BKind_Rank_By_YM '

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code


     Set @Msg = @Kind+'.新增 年月業務大分類實銷達成率排名 資料 [Sale_Member_BKind_Rank_By_YM]'
     set @strSQL = 'select Rank, m.year, m.month,'+@CR+                                                                                                                                                                                                                                                                                   
                   '       m.kind_name_AA, IsNull(emp_no_AA, '''') as emp_no_AA, IsNull(emp_name_AA, '''') as emp_name_AA, IsNull(sale_amt_AA, 0) as sale_amt_AA, IsNull(target_amt_AA, 0) as target_amt_AA, IsNull(rate_AA, 0) as rate_AA, IsNull(YM_Sale_amt_AA, 0) as YM_Sale_amt_AA, IsNull(YM_Target_AA, 0) as YM_Target_AA, IsNull(YM_rate_AA, 0) as YM_rate_AA,'+@CR+
                   '       m.kind_name_AB, IsNull(emp_no_AB, '''') as emp_no_AB, IsNull(emp_name_AB, '''') as emp_name_AB, IsNull(sale_amt_AB, 0) as sale_amt_AB, IsNull(target_amt_AB, 0) as target_amt_AB, IsNull(rate_AB, 0) as rate_AB, IsNull(YM_Sale_amt_AB, 0) as YM_Sale_amt_AB, IsNull(YM_Target_AB, 0) as YM_Target_AB, IsNull(YM_rate_AB, 0) as YM_rate_AB,'+@CR+
                   '       m.kind_name_AC, IsNull(emp_no_AC, '''') as emp_no_AC, IsNull(emp_name_AC, '''') as emp_name_AC, IsNull(sale_amt_AC, 0) as sale_amt_AC, IsNull(target_amt_AC, 0) as target_amt_AC, IsNull(rate_AC, 0) as rate_AC, IsNull(YM_Sale_amt_AC, 0) as YM_Sale_amt_AC, IsNull(YM_Target_AC, 0) as YM_Target_AC, IsNull(YM_rate_AC, 0) as YM_rate_AC,'+@CR+
                   '       m.kind_name_AD, IsNull(emp_no_AD, '''') as emp_no_AD, IsNull(emp_name_AD, '''') as emp_name_AD, IsNull(sale_amt_AD, 0) as sale_amt_AD, IsNull(target_amt_AD, 0) as target_amt_AD, IsNull(rate_AD, 0) as rate_AD, IsNull(YM_Sale_amt_AD, 0) as YM_Sale_amt_AD, IsNull(YM_Target_AD, 0) as YM_Target_AD, IsNull(YM_rate_AD, 0) as YM_rate_AD,'+@CR+
                   '       m.kind_name_AE, IsNull(emp_no_AE, '''') as emp_no_AE, IsNull(emp_name_AE, '''') as emp_name_AE, IsNull(sale_amt_AE, 0) as sale_amt_AE, IsNull(target_amt_AE, 0) as target_amt_AE, IsNull(rate_AE, 0) as rate_AE, IsNull(YM_Sale_amt_AE, 0) as YM_Sale_amt_AE, IsNull(YM_Target_AE, 0) as YM_Target_AE, IsNull(YM_rate_AE, 0) as YM_rate_AE,'+@CR+
                   '       m.kind_name_OT, IsNull(emp_no_OT, '''') as emp_no_OT, IsNull(emp_name_OT, '''') as emp_name_OT, IsNull(sale_amt_OT, 0) as sale_amt_OT, IsNull(target_amt_OT, 0) as target_amt_OT, IsNull(rate_OT, 0) as rate_OT, IsNull(YM_Sale_amt_OT, 0) as YM_Sale_amt_OT'+@CR+                                                                      
                   '       into Sale_Member_BKind_Rank_By_YM'+@CR+                                                                                                                                                                                                                                                                        
                   '  from Sale_Member_BKind_Rank_By_YM_tmp m'+@TB_Hint+@CR+                                                                                                                                                                                                                                                         
                   '       inner join'+@CR+                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                            
                   '       (select m.year, m.month,'+@CR+                                                                                                                                                                                                                                                                                 
                   -- 百貨                                                                                                                                                                                                                                                                                                                 
                   '               kind_name_AA, Sum(Isnull(sale_amt_AA, 0)) as YM_sale_amt_AA, Target_AA as YM_Target_AA,'+@CR+                                                                                                                                                                                                                  
                   '               case'+@CR+                                                                                                                                                                                                                                                                                             
                   '                 when Isnull(Target_AA, 0) = 0 then 0'+@CR+                                                                                                                                                                                                                                                           
                   '                 else Round(Sum(Isnull(sale_amt_AA, 0)) / Target_AA, 4)'+@CR+                                                                                                                                                                                                                                           
                   '               end as YM_rate_AA,'+@CR+                                                                                                                                                                                                                                                                                 
                   -- 3C                                                                                                                                                                                                                                                                                                                   
                   '               kind_name_AB, Sum(Isnull(sale_amt_AB, 0)) as YM_sale_amt_AB, Target_AB as YM_Target_AB,'+@CR+                                                                                                                                                                                                                  
                   '               case'+@CR+                                                                                                                                                                                                                                                                                             
                   '                 when Isnull(Target_AB, 0) = 0 then 0'+@CR+                                                                                                                                                                                                                                                           
                   '                 else Round(Sum(Isnull(sale_amt_AB, 0)) / Target_AB, 4)'+@CR+                                                                                                                                                                                                                                           
                   '               end as YM_rate_AB,'+@CR+                                                                                                                                                                                                                                                                                 
                   -- 化工                                                                                                                                                                                                                                                                                                                 
                   '               kind_name_AC, Sum(Isnull(sale_amt_AC, 0)) as YM_sale_amt_AC, Target_AC as YM_Target_AC,'+@CR+                                                                                                                                                                                                                  
                   '               case'+@CR+                                                                                                                                                                                                                                                                                             
                   '                 when Isnull(Target_AC, 0) = 0 then 0'+@CR+                                                                                                                                                                                                                                                           
                   '                 else Round(Sum(Isnull(sale_amt_AC, 0)) / Target_AC, 4)'+@CR+                                                                                                                                                                                                                                           
                   '               end as YM_rate_AC,'+@CR+                                                                                                                                                                                                                                                                                 
                   -- 工具                                                                                                                                                                                                                                                                                                                 
                   '               kind_name_AD, Sum(Isnull(sale_amt_AD, 0)) as YM_sale_amt_AD, Target_AD as YM_Target_AD,'+@CR+                                                                                                                                                                                                                  
                   '               case'+@CR+                                                                                                                                                                                                                                                                                             
                   '                 when Isnull(Target_AD, 0) = 0 then 0'+@CR+                                                                                                                                                                                                                                                           
                   '                 else Round(Sum(Isnull(sale_amt_AD, 0)) / Target_AD, 4)'+@CR+                                                                                                                                                                                                                                           
                   '               end as YM_rate_AD,'+@CR+                                                                                                                                                                                                                                                                                 
                   -- 電裝                                                                                                                                                                                                                                                                                                                 
                   '               kind_name_AE, Sum(Isnull(sale_amt_AE, 0)) as YM_sale_amt_AE, Target_AE as YM_Target_AE,'+@CR+                                                                                                                                                                                                                  
                   '               case'+@CR+                                                                                                                                                                                                                                                                                             
                   '                 when Isnull(Target_AE, 0) = 0 then 0'+@CR+                                                                                                                                                                                                                                                           
                   '                 else Round(Sum(Isnull(sale_amt_AE, 0)) / Target_AE, 4)'+@CR+                                                                                                                                                                                                                                           
                   '               end as YM_rate_AE,'+@CR+                                                                                                                                                                                                                                                                                 
                   -- 其他                                                                                                                                                                                                                                                                                                                 
                   '               kind_name_OT, Sum(Isnull(sale_amt_OT, 0)) as YM_sale_amt_OT'+@CR+                                                                                                                                                                                                                                            
                   '          from Sale_Member_BKind_Rank_By_YM_tmp M'+@TB_Hint+@CR+                                                                                                                                                                                                                                                 
                   '               left join Ori_Xls#Target_Stock_BKind_Business D'+@TB_Hint+@CR+                                                                                                                                                                                                                                   
                   '                  on m.Year = d.year'+@CR+                                                                                                                                                                                                                                                                            
                   '                 and m.month = d.month'+@CR+                                                                                                                                                                                                                                                                          
                   '         group by m.year, m.month, kind_name_AA, kind_name_AB, kind_name_AC, kind_name_AD, kind_name_AE, kind_name_OT, Target_AA, Target_AB, Target_AC, Target_AD, Target_AE'+@CR+                                                                                                                                                
                   '       ) D'+@CR+                                                                                                                                                                                                                                                                                                         
                   '        On m.year=d.year'+@CR+                                                                                                                                                                                                                                                                                        
                   '       and m.month=d.month '   

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end 

--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--
--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--
  set @Kind = 'M'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     
     Set @Msg = @Kind+'.刪除 將月份資料進行轉向 ['+@TB_RT_Data+'_MM_Pivot]'
     set @strSQL = 'IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N''[dbo].['+@TB_RT_Data+'_MM_Pivot]'') AND type in (N''U''))'+@CR+
                   '   Drop Table '+@TB_RT_Data+'_MM_Pivot '
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code

     Set @Msg = @Kind+'.新增 將月份資料進行轉向 ['+@TB_RT_Data+'_MM_Pivot]'
     set @strSQL = 'select kind, kind_name,'+@CR+
                   '       area_year,'+@CR+
                   '       area, area_name,'+@CR+
                   '       Data_Type,'+@CR+
                   '       IsNull([1], 0) as ''M01'', IsNull([2], 0) as ''M02'', IsNull([3], 0) as ''M03'','+@CR+
                   '       IsNull([4], 0) as ''M04'', IsNull([5], 0) as ''M05'', IsNull([6], 0) as ''M06'','+@CR+
                   '       IsNull([7], 0) as ''M07'', IsNull([8], 0) as ''M08'', IsNull([9], 0) as ''M09'','+@CR+
                   '       IsNull([10], 0) as ''M10'', IsNull([11], 0) as ''M11'', IsNull([12], 0) as ''M12'','+@CR+
                   '       (IsNull([1], 0) +  IsNull([2], 0) +  IsNull([3], 0) +'+@CR+
                   '        IsNull([4], 0) +  IsNull([5], 0) +  IsNull([6], 0) +'+@CR+
                   '        IsNull([7], 0) +  IsNull([8], 0) +  IsNull([9], 0) +'+@CR+
                   '        IsNull([10], 0) + IsNull([11], 0) + IsNull([12], 0)) as ''M_SUM'','+@CR+
                   '       Round([1]+[2]+[3]+[4]+[5]+[6]+[7]+[8]+[9]+[10]+[11]+[12] / 12, 4) as ''M_AVG'''+@CR+
                   '       into '+@TB_RT_Data+'_MM_Pivot'+@CR+
                   '  from (select kind, kind_name, Data_Type, area_year, area, area_name, area_month, Isnull(amt, 0) as amt'+@CR+
                   '          from '+@TB_RT_Data+@TB_Hint+@CR+
                   '         where 1=1'+@CR+
                   '           and area_month <> 0'+@CR+
                   '        ) m'+@CR+
                   '        pivot (Sum(amt) for area_month in ([0], [1], [2], [3], [4], [5], [6], [7], [8], [9], [10], [11], [12])) d'+@CR+
                   ' order by 1, 2, 3, 4, 5'

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--
--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--
  set @Kind = 'Y'
  if (@Kind like @in_Kind+'%' Or @in_Kind ='' Or @in_Kind like '%'+@Kind+'%')
  begin
     declare @Y_Str1 Varchar(1000) = ''
     declare @Y_Str2 Varchar(1000) = ''
     declare @area_year int
     
     declare cur_year_data cursor for
       select distinct area_year
         from realtime_saledata m
        where area_year <> 0 
          and area_month = 0 
        order by 1
          -- 2015/03/16 由於資料會累計成長，又加上使用 not exists 以及 distinct 方式 會直接使用到 tempdb，導致效能變差，
          -- 因此直接取消使用以下語法。
          /*
          and not exists
              (select distinct kind, count(*) as cnt
                 from realtime_saledata d1
                where m.kind = d1.kind
                group by kind
               having count(*) = 1
              )
           */
     Open cur_year_data
     Fetch Next from cur_year_data into @area_year
     
     While @@Fetch_Status = 0
     begin
       Set @Y_Str1 = @Y_Str1 + '[' +Convert(varchar(10), IsNull(@area_year, '00')) + '], '  
       Set @Y_Str2 = @Y_Str2 + 'IsNull([' +Convert(varchar(10), IsNull(@area_year, '00')) + '], 0) as ['+Convert(varchar(10), IsNull(@area_year, '0000'))+'], '
       Fetch Next from cur_year_data into @area_year
     end
     Close cur_year_data
     Deallocate cur_year_data

     if (Len(@Y_Str1) <> 0) Or (Len(@Y_Str2) <> 0)
     begin
        Set @Y_Str1 = Substring(@Y_Str1, 1, Len(@Y_Str1) - 1)
        Set @Y_Str2 = Substring(@Y_Str2, 1, Len(@Y_Str2) - 1)
     end

     Set @Msg = @Kind+'.刪除 將年度資料進行轉向 ['+@TB_RT_Data+'_YY_Pivot]'
     set @strSQL = 'IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N''[dbo].['+@TB_RT_Data+'_YY_Pivot]'') AND type in (N''U''))'+@CR+
                   '   Drop Table '+@TB_RT_Data+'_YY_Pivot '
     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code  

     Set @Msg = @Kind+'.新增 將年度資料進行轉向 ['+@TB_RT_Data+'_YY_Pivot]'
     set @strSQL = 'select kind, kind_name,'+@CR+
                   '       area, area_name,'+@CR+
                   '       Data_Type,'+@CR+
                   '       area_month,'+@CR+
                   '       '+@Y_Str2+@CR+
                   '       into '+@TB_RT_Data+'_YY_Pivot'+@CR+
                   '  from (select kind, kind_name, Data_Type, area_year, area_month, area, area_name, IsNull(amt, 0) as amt'+@CR+
                   '          from '+@TB_RT_Data+@CR+
                   '         where 1=1'+@CR+
                   '           and area_year <> 0'+@CR+
                   --'           and area_month = 0'+@CR+
                   '       ) m'+@CR+
                   '       pivot (Sum(amt) for area_year in ('+@Y_Str1+')) d'+@CR+
                   ' order by 1, 2, 3, 4, 5'

     Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
     if @Get_Result = @Err_Code Set @Result = @Err_Code
  end

  --Set @Msg = '清除臨時資料表 ['+@TB_RT_Data_tmp+']'
  --set @strSQL = 'IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N''[dbo].['+@TB_RT_Data_tmp+']'') AND type in (N''U''))'+@CR+
  --             '    Drop Table '+@TB_RT_Data_tmp+@CR
  --Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail

  -- 20151008 Rickliu Add Check Trans_Log Script
  Set @ETime = Getdate()
  Print 'Check Trans_Log End Scripts...'
  Print 'select * from trans_log where process =''uSP_Realtime_SaleData'' and trans_date Between '''+Convert(Varchar(100), @BTime, 120)+''' and '''+Convert(Varchar(100), @ETime, 120)+''' order by trans_date'
     
end
GO
