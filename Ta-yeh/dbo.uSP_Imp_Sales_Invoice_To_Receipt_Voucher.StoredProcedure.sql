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
  Declare @RM Varchar(100) = '�P���o������b�ǲ�'
  Declare @RowCount Table (cnt int)
  Declare @Sender Varchar(100) = 'fa@ta-yeh.com.tw;'
  Declare @strSQL Varchar(Max) = ''
  Declare @TB_Name Varchar(100) = 'Sales_Invoice_To_Receipt_Voucher'
  Declare @Xls_Name Varchar(100) = 'ori_xls#'+@TB_Name
  Declare @TB_tmp_Name Varchar(100) = @TB_Name+'_tmp'
  Declare @sys_date Varchar(20) = convert(varchar(10), getdate(), 120)

--=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
--=-[�{���}�l]=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
--=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
  set @Msg = '�d�߶ǲ���ƬO�_�w�s�b�C'
  set @strSQL = 'select count(1) as cnt from '+@Xls_Name
  delete @RowCount
  insert into @RowCount exec (@strSQL)
  select @Cnt=Rtrim(isnull(Cnt, '')) from @RowCount
  
  if @Cnt = 0
  begin
     set @Msg = @Xls_Name+' '+@RM+'�� �L��ơA�L������!!'
     Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt
  end
  else
  begin
     set @Msg = '�d�߶ǲ���ƬO�_�w�s�b�C'
     set @strSQL = 'select count(1) '+@CR+
                   '  from '+@db+'aslip '+@CR+
                   ' where sp_memo like ''%'+@RM+'%'+@sys_date+'%'' '
     delete @RowCount
     insert into @RowCount exec (@strSQL)
     select @Cnt=Rtrim(isnull(Cnt, '')) from @RowCount
     
     if @Cnt <> 0
     begin
        set @Msg ='�w�� [ '+@Sys_Date+' '+@RM+' ] ��ơA�Y�ݭn����h����M�����ݶǲ��᭫��Y�i�C'
        Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt
     end
     else
     begin
        IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].['+@TB_tmp_Name+']') AND type in (N'U'))
        begin
           Set @Msg = '�M���{����J������ƪ� ['+@TB_tmp_Name+']�C'
   
           Set @strSQL = 'DROP TABLE [dbo].['+@TB_tmp_Name+']'
           Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
        end
   
        set @Msg = '�إ��{����J������ƪ� ['+@TB_tmp_Name+']�C'
        set @strSQL = 'select '+@CR+
                      -- �ǲ����
                      '       m.sd_date,  '+@CR+
                      -- �ǲ����X
                      '       case '+@CR+
                      '         when isnull(convert(varchar(20), d2.sp_no), '''') = '''' '+@CR+
                      '         then Substring(Replace(m.sd_date, ''/'', ''''), 3, 6)+''0001'' '+@CR+
                      '         else Convert(Varchar(20), Convert(Varchar(10), d2.sp_no)+1) '+@CR+
                      '       end as sp_no, '+@CR+
                      -- �ǲ����ӺK�n
                      -- 2017/04/11 �]�|�D����ҭn�D���K�n���e(�Ȥ�²��+�o�����X)
                      --'       ''�o�����X:''+m.sd_dcr as sd_dcr, '+@CR+
                      '       ''�Ȥ�(''+Rtrim(d.ct_sname collate Chinese_Taiwan_Stroke_CI_AS)+''),�o�����X(''+RTrim(m.sd_dcr)+'')'' as sd_dcr, '+@CR+
                      -- ��إN��
                      '       m.sd_atno, '+@CR+
                      -- �Ȥ�s��, �Ȥ�²��, �дڹ�H
                      '       isnull(m.sd_ctno, '''') as sd_ctno, isnull(d.ct_sname, '''') as ct_sname, '+@CR+
                      '       case '+@CR+
                      '         when sd_atno = ''1123'' '+@CR+
                      '         then isnull(ct_credit, '''') '+@CR+
                      '         else '''' '+@CR+
                      '       end as ct_credit, '+@CR+
                      -- �ɶU�O, ������B, ������B(��)
                      '       sd_doc, sd_amt, sd_amt as sd_oamt, '+@CR+
                      -- �g��H���s, �g��H�m�W
                      '       m.sp_mkman, d1.e_name,  '+@CR+
                      -- ���ӧǸ�
                      '       Row_Number() over(Partition BY m.sd_date order by m.sd_date) as sd_seq, '+@CR+
                      -- �ǲ��Ƶ�
                      '       ''�t�ζפJ '+@RM+' �� ''+convert(Varchar(20), getdate(), 120) as sp_memo '+@CR+
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
     
        set @Msg = '�s�W�ǲ� �D�ɸ��'
        set @strSQL = 'Insert Into '+@db+'anpsp '+@CR+
                      '(sp_class, sp_date, sp_no, sp_mkman, sp_memo) '+@CR+
                      'select distinct '+@CR+
                      -- ��b�ǲ�
                      '       ''3'' as sp_class, '+@CR+
                      '       sd_date, sp_no, e_name, sp_memo '+@CR+
                      '  from '+@TB_tmp_Name+@CR+
                      ' where sd_dcr is not null'
        Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
        
        
        set @Msg = '�s�W�ǲ� �����ɸ��'
        set @strSQL = 'Insert Into '+@db+'anpdt '+@CR+
                      '(sd_class, sd_date, sd_no, sd_atno, sd_dcr, sd_doc, '+@CR+
                      ' sd_ctno, sd_amt, sd_rate_nm, sd_oamt, sd_seq, sd_memo) '+@CR+
                      'select ''3'' as sd_class, sd_date, sp_no, sd_atno, sd_dcr, sd_doc, '+@CR+
                      '       ct_credit, sd_amt, ''NT'' as sd_rate_nm, sd_amt as sd_oamt, sd_seq, sd_dcr '+@CR+
                      '  from '+@TB_tmp_Name+@CR+
                      ' where sd_dcr is not null'
        Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
      
        set @Msg = '�d�ߵ���'
        set @strSQL = 'select count(1) '+@CR+
                      '  from '+@db+'anpsp '+@CR+
                      ' where sp_memo like ''%'+@RM+'%'+@sys_date+'%'' '
        delete @RowCount
        insert into @RowCount exec (@strSQL)
        select @Cnt=Rtrim(isnull(Cnt, '')) from @RowCount
          
        if isnull(@Cnt, 0) <> 0
        begin
           Set @RM = @RM+'...'+@Sys_Date+' �w���� �i'+Convert(Varchar(100), @cnt)+'�j �i�ǲ�!!'+
                     '�Щ� ��V�t�� TA16TX �d�߶ǲ����!!'
           Exec uSP_Sys_Send_Mail @Proc, @RM, @Sender, @RM, ''
           
           set @Msg = '�R���ǲ� XLS ���'
           set @strSQL = 'Delete '+@Xls_Name
           Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL        
        end
     end
  end
end
GO
