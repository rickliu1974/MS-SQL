USE [DW]
GO
/****** Object:  StoredProcedure [dbo].[uSP_Imp_Sales_Invoice_To_Receipt_Voucher]    Script Date: 08/18/2017 17:43:40 ******/
DROP PROCEDURE [dbo].[uSP_Imp_Sales_Invoice_To_Receipt_Voucher]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE Proc [dbo].[uSP_Imp_Sales_Invoice_To_Receipt_Voucher]
as
begin
  Declare @Proc Varchar(50) = 'uSP_Imp_Sales_Invoice_To_Receipt_Voucher'
  Declare @Cnt Int = 0
  Declare @RowCnt Int = 0
  Declare @Msg Varchar(4000) =''
  Declare @Errcode int = -1
  Declare @CR Varchar(4) = char(13)+char(10)
  Declare @db Varchar(100) = 'TYEIPDBS2.lytdbta16TX.dbo.'
  Declare @RM Varchar(100) = '銷項發票轉轉帳傳票'
  Declare @RowCount Table (cnt int)
  Declare @Sender Varchar(100) = 'fa@ta-yeh.com.tw;'
  Declare @strSQL Varchar(Max) = ''
  Declare @TB_Name Varchar(100) = 'Sales_Invoice_To_Receipt_Voucher'
  Declare @Xls_Name Varchar(100) = 'ori_xls#'+@TB_Name
  Declare @TB_tmp_Name Varchar(100) = @TB_Name+'_tmp'
  Declare @sys_date Varchar(20) = convert(varchar(10), getdate(), 120)

--=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
--=-[程式開始]=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
--=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
  set @Msg = '查詢傳票資料是否已存在。'
  set @strSQL = 'select count(1) as cnt from '+@Xls_Name
  delete @RowCount
  insert into @RowCount exec (@strSQL)
  select @Cnt=Rtrim(isnull(Cnt, '')) from @RowCount
  
  if @Cnt = 0
  begin
     set @Msg = @Xls_Name+' '+@RM+'檔 無資料，無須轉檔!!'
     Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt
  end
  else
  begin
     set @Msg = '查詢傳票資料是否已存在。'
     set @strSQL = 'select count(1) '+@CR+
                   '  from '+@db+'aslip '+@CR+
                   ' where sp_memo like ''%'+@RM+'%'+@sys_date+'%'' '
     delete @RowCount
     insert into @RowCount exec (@strSQL)
     select @Cnt=Rtrim(isnull(Cnt, '')) from @RowCount
     
     if @Cnt <> 0
     begin
        set @Msg ='已有 [ '+@Sys_Date+' '+@RM+' ] 資料，若需要重轉則先行清除所屬傳票後重轉即可。'
        Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt
     end
     else
     begin
        IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].['+@TB_tmp_Name+']') AND type in (N'U'))
        begin
           Set @Msg = '清除臨時轉入介面資料表 ['+@TB_tmp_Name+']。'
   
           Set @strSQL = 'DROP TABLE [dbo].['+@TB_tmp_Name+']'
           Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
        end
   
        set @Msg = '建立臨時轉入介面資料表 ['+@TB_tmp_Name+']。'
        set @strSQL = 'select '+@CR+
                      -- 傳票日期
                      '       m.sd_date,  '+@CR+
                      -- 傳票號碼
                      '       case '+@CR+
                      '         when isnull(convert(varchar(20), d2.sp_no), '''') = '''' '+@CR+
                      '         then Substring(Replace(m.sd_date, ''/'', ''''), 3, 6)+''0001'' '+@CR+
                      '         else Convert(Varchar(20), Convert(Varchar(10), d2.sp_no)+1) '+@CR+
                      '       end as sp_no, '+@CR+
                      -- 傳票明細摘要
                      -- 2017/04/11 財會主管鍾課要求更改摘要內容(客戶簡稱+發票號碼)
                      --'       ''發票號碼:''+m.sd_dcr as sd_dcr, '+@CR+
                      '       ''客戶(''+Rtrim(d.ct_sname collate Chinese_Taiwan_Stroke_CI_AS)+''),發票號碼(''+RTrim(m.sd_dcr)+'')'' as sd_dcr, '+@CR+
                      -- 科目代號
                      '       m.sd_atno, '+@CR+
                      -- 客戶編號, 客戶簡稱, 請款對象
                      '       isnull(m.sd_ctno, '''') as sd_ctno, isnull(d.ct_sname, '''') as ct_sname, '+@CR+
                      '       case '+@CR+
                      '         when sd_atno = ''1123'' '+@CR+
                      '         then isnull(ct_credit, '''') '+@CR+
                      '         else '''' '+@CR+
                      '       end as ct_credit, '+@CR+
                      -- 借貸別, 交易金額, 交易金額(原)
                      '       sd_doc, sd_amt, sd_amt as sd_oamt, '+@CR+
                      -- 經手人員編, 經手人姓名
                      '       m.sp_mkman, d1.e_name,  '+@CR+
                      -- 明細序號
                      '       Row_Number() over(Partition BY m.sd_date order by m.sd_date) as sd_seq, '+@CR+
                      -- 傳票備註
                      '       ''系統匯入 '+@RM+' 於 ''+convert(Varchar(20), getdate(), 120) as sp_memo '+@CR+
                      '       into '+@TB_tmp_Name+' '+@CR+
                      '  from (select sd_date, sd_dcr, sd_ctno,  '+@CR+
                      '               ''C'' as sd_doc, ''4101'' as sd_atno, C_4101_AMT as sd_amt, '+@CR+
                      '               sp_mkman '+@CR+
                      '          from '+@Xls_Name+@CR+
                      '         union '+@CR+
                      '        select sd_date, sd_dcr, sd_ctno, '+@CR+
                      '               ''C'' as sd_doc, ''2194'' as sd_atno, C_2194_AMT as sd_amt, '+@CR+
                      '               sp_mkman '+@CR+
                      '          from '+@Xls_Name+@CR+
                      '         union '+@CR+
                      '        select sd_date, sd_dcr, sd_ctno, '+@CR+
                      '               ''D'' as sd_doc, ''1123'' as sd_atno, D_1123_AMT as sd_amt, '+@CR+
                      '               sp_mkman '+@CR+
                      '          from '+@Xls_Name+@CR+
                      '       ) m '+@CR+
                      '      left join '+@db+'pcust d '+@CR+
                      '        on m.sd_ctno = d.ct_no collate Chinese_Taiwan_Stroke_CI_AS '+@CR+
                      '       and d.ct_class = ''1'' '+@CR+
                      '      left join '+@db+'pemploy d1 '+@CR+
                      '        on m.sp_mkman = d1.e_no collate Chinese_Taiwan_Stroke_CI_AS '+@CR+
                      '      left join '+@CR+
                      '      (select sp_date, Convert(decimal, max(sp_no)) as sp_no '+@CR+
                      '         from (select sp_date, sp_no '+@CR+
                      '                 from '+@db+'aslip '+@CR+
                      '                union '+@CR+
                      '               select sp_date, sp_no as sp_no '+@CR+
                      '                 from '+@db+'anpsp '+@CR+
                      '              ) m '+@CR+
                      '        group by sp_date '+@CR+
                      '      ) d2 '+@CR+
                      '      on m.sd_date = d2.sp_date '
   
        Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
     
        set @Msg = '新增傳票 主檔資料'
        set @strSQL = 'Insert Into '+@db+'anpsp '+@CR+
                      '(sp_class, sp_date, sp_no, sp_mkman, sp_memo) '+@CR+
                      'select distinct '+@CR+
                      -- 轉帳傳票
                      '       ''3'' as sp_class, '+@CR+
                      '       sd_date, sp_no, e_name, sp_memo '+@CR+
                      '  from '+@TB_tmp_Name+@CR+
                      ' where sd_dcr is not null'
        Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
        
        
        set @Msg = '新增傳票 明細檔資料'
        set @strSQL = 'Insert Into '+@db+'anpdt '+@CR+
                      '(sd_class, sd_date, sd_no, sd_atno, sd_dcr, sd_doc, '+@CR+
                      ' sd_ctno, sd_amt, sd_rate_nm, sd_oamt, sd_seq, sd_memo) '+@CR+
                      'select ''3'' as sd_class, sd_date, sp_no, sd_atno, sd_dcr, sd_doc, '+@CR+
                      '       ct_credit, sd_amt, ''NT'' as sd_rate_nm, sd_amt as sd_oamt, sd_seq, sd_dcr '+@CR+
                      '  from '+@TB_tmp_Name+@CR+
                      ' where sd_dcr is not null'
        Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
      
        set @Msg = '查詢筆數'
        set @strSQL = 'select count(1) '+@CR+
                      '  from '+@db+'anpsp '+@CR+
                      ' where sp_memo like ''%'+@RM+'%'+@sys_date+'%'' '
        delete @RowCount
        insert into @RowCount exec (@strSQL)
        select @Cnt=Rtrim(isnull(Cnt, '')) from @RowCount
          
        if isnull(@Cnt, 0) <> 0
        begin
           Set @RM = @RM+'...'+@Sys_Date+' 已產生 【'+Convert(Varchar(100), @cnt)+'】 張傳票!!'+
                     '請於 凌越系統 TA16TX 查詢傳票資料!!'
           Exec uSP_Sys_Send_Mail @Proc, @RM, @Sender, @RM, ''
           
           set @Msg = '刪除傳票 XLS 資料'
           set @strSQL = 'Delete '+@Xls_Name
           Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL        
        end
     end
  end
end
GO
