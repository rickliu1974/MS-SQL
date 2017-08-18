USE [DW]
GO
/****** Object:  StoredProcedure [dbo].[xuSP_ETL_PCust_Test(201708����)]    Script Date: 08/18/2017 17:18:56 ******/
DROP PROCEDURE [dbo].[xuSP_ETL_PCust_Test(201708����)]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE Proc [dbo].[xuSP_ETL_PCust_Test(201708����)] (@Kind Varchar(1)='', @Value Timestamp = Null)
as
begin
  /*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
   Program Name: uSP_TA13#ETL_PCust
   Create Date: 2013/03/01
   Creator: Rickliu
   Updated Date: 2013/09/06 [�ק� Trans_Log �T�����e]
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
  
  -- Timestamp �i�H���ഫ�� Int ��A��^ Timestamp�A��r���i�� Timestamp�A���ƭȷ|�����T�C
  Declare @sTimestamp Varchar(1000) = Isnull(Convert(Varchar(1000), Convert(Int, @Value)), '')

  Begin try
 print '1'
    set @strSQL_Inx = 
        'if exists (select 1 from sysindexes where id = object_id(''@Tb_Name'') and name  = ''@Inx_Name'' and indid > 0 and indid < 255) drop index @Tb_Name.@Inx_Name '+@CR+
        '   create @Inx_clustered index @Inx_Name on @Tb_Name (@Inx_Columns) with fillfactor= 30 on "PRIMARY" '
  
 print '2'
    If @Kind Not In ('I', 'U', 'D', '') raiserror ('@Kind �Ѽƥ����� I, U, D, �ť�', 50005, 10, 1)

 print '3'
    IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].['+@Tb_name+']') AND type in (N'U'))
       Set @Chk_Tb_Exists = 'Y'

    IF (@Kind = '' And @Value = 0 )
    begin
       Set @Msg = '�M�Ÿ�ƪ� ['+@Tb_name+']'
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

    -- �p�G��ƪ�s�b�A�h�N���� Insert            
    If (@Chk_Tb_Exists = 'Y')
       set @strSQL = @strSQL +@CR+
                  'Insert into '+@Tb_tmp_name
   
    set @strSQL = @strSQL +@CR+
                  -- ���O, �s��, �W��, ²��
                  'select M.[ct_class], Rtrim(M.[ct_no]) as ct_no, Rtrim(M.[ct_name]) as ct_name, Rtrim(M.[ct_sname]) as ct_sname, '+@CR+
                  -- �Ȥ�K�X�s��
                  '       Case when m.ct_class = ''1'' and Len(Rtrim(M.[ct_no])) > 8 and Substring(M.ct_no, 1, 2) Not in (''IT'', ''IZ'') then Substring(Rtrim(M.[ct_no]), 1, 8) else '''' end as ct_no8, '+@CR+ 
                  -- �Ȥ�K�X�W��
                  '       LTrim(RTrim(Replace(Replace(Replace(Replace(Replace(Replace(m.ct_sname, ''#'', ''''), ''@'', ''''), ''-'', ''''), ''T'', ''''), ''P'', ''''), ''/'', ''''))) as ct_sname8, '+@CR+
                  -- ���q�a�}, �o���a�}, �e�f�a�}
                  '       [ct_addr1]=Rtrim(Convert(Varchar(255), M.ct_addr1)), [ct_addr2]=Rtrim(Convert(Varchar(255), M.ct_addr2)),  [ct_addr3]=Rtrim(Convert(Varchar(255), M.ct_addr3)), '+@CR+
                  -- �q��, �ǯu, �Τ@�s��, �t�d�H, �p���H
                  '       Rtrim(M.[ct_tel]) as ct_tel, Rtrim(M.[ct_fax]) as ct_fax, Rtrim(M.[ct_unino]) as ct_unino, Rtrim(M.[ct_presid]) as ct_presid, Rtrim(M.[ct_contact]) as ct_contact, '+@CR+
                  -- �������, �~�ȭ�, �b���B��, �H���B��, �Ȧ�b��, �Ȧ�W��
                  '       M.[ct_payfg], Rtrim(M.[ct_sales]) as ct_sales, M.[ct_p_limit], M.[ct_b_limit], Rtrim(M.[ct_bkno]) as ct_bkno, Rtrim(M.[ct_bknm]) as ct_bknm, '+@CR+
                  -- �Ƶ�, �����s��, ������, �o�����Y, �f�B���q
                  '       [ct_rem]=Rtrim(Convert(Varchar(255), M.ct_rem)), Rtrim(M.[ct_dept]) as ct_dept,  M.[ct_payrate], Rtrim(M.[ct_ivtitle]) as ct_ivtitle, Rtrim(M.[ct_porter]) as ct_porter, '+@CR+
                  -- �д�|�I�ڹ�H, �I�ڤ覡, �I�ڤѼ�, ���M��, �w����, �̪�����, �Ȥ�ݼt��, �ǿ�X��, ����, �a��,��O�N�X
                  '       Rtrim(M.[ct_credit]) as ct_credit, M.[ct_pmode], M.[ct_pdate], M.[ct_prenpay], M.[ct_prepay], M.[ct_last_dt], M.[ct_flg], M.[ct_t_fg], M.[ct_grade], M.[ct_area], '+@CR+
                  -- 2015/02/03 ������O
                  '       case when RTrim(isnull(M.ct_curt_id, '''')) = '''' then ''NT'' else RTrim(isnull(M.ct_curt_id, '''')) end as ct_curt_id, '+@CR+
                  -- �p���H¾��, �T���I�ڤ覡, �u�t�n�O��, ���u��, �ꥻ�B, �T���w����, �T�����M��, �T�����I�b���, �T�����I�����
                  '       Rtrim(M.[ct_cont_sp]) as ct_cont_sp, M.[ct_pay], Rtrim(M.[ct_regist]) as ct_regist, M.[ct_worker],  M.[ct_capital], M.[ct_skpay], M.[ct_sknpay], M.[ct_accno2], M.[ct_chkno2], '+@CR+
                  -- ���ɤ��, �i�P�s�дڹ�H, �i�P�s�w����, �i�P�s���M��, �ꤺ.��~
                  '       M.[ct_cdate], M.[ct_payer], M.[ct_advance], M.[ct_debt], M.[ct_abroad], '+@CR+
                  -- �Ȥ�|�t�Ӫ����@(�ȡG�}�o����B�t�G�t�X�_�l��), �Ȥ�|�t�Ӫ����G(�ȡG��������B�t�G�����t�X��), �Ȥ�|�t�Ӫ����T(�ȡG�`���q�O�B�t�G���ϥ�), �Ȥ�|�t�Ӫ����|(�ȡG�����O�B�t�G�t�ӵ���), �Ȥ�|�t�Ӫ�����(�ȡG�ȪA�M��1�B�t�Gñ�������), �Ȥ�|�t�Ӫ�����(�ȡG�ȪA�M��2�B�t�G���ϥ�)
                  '       Rtrim(M.[ct_fld1]) as ct_fld1, Rtrim(M.[ct_fld2]) as ct_fld2, Rtrim(M.[ct_fld3]) as ct_fld3, Rtrim(M.[ct_fld4]) as ct_fld4, Rtrim(M.[ct_fld5]) as ct_fld5, Rtrim(M.[ct_fld6]) as ct_fld6, '+@CR+
                  -- �Ȥ�|�t�Ӫ����C(�ȡG�ȪA�M��3�B�t�G���ϥ�), �Ȥ�|�t�Ӫ����K(�ȡG�ȪA�M��4�B�t�G���ϥ�), �Ȥ�|�t�Ӫ����E(�ȡG���ڤ覡�B�t�G���ϥ�), �Ȥ�|�t�Ӫ����Q(�ȡG�o���W�١B�t�G���ϥ�), �Ȥ�|�t�Ӫ����Q�@(�ȡG�I�ڱ���B�t�G���ϥ�), �Ȥ�|�t�Ӫ����Q�G(�ȡG�½s�� 2012�}�b�ϥ�)
                  '       Rtrim(M.[ct_fld7]) as ct_fld7, Rtrim(M.[ct_fld8]) as ct_fld8, Rtrim(M.[ct_fld9]) as ct_fld9, Rtrim(M.[ct_fld10]) as ct_fld10, Rtrim(M.[ct_fld11]) as ct_fld11, Rtrim(M.[ct_fld12]) as ct_fld12, '+@CR+
                  --�T���B�e�覡, �}�ߵo������, ���������p�Ʀ��, ������B���p�Ʀ��, ��~�O, ��׻Ȧ�, �ϰ�s��, �Ȥ�ӷ�, �Ȥ����O, ���X�g��
                  '       M.[ct_sea], M.[ct_invofg], M.[ct_udec], M.[ct_tdec], M.[ct_busine], Rtrim(M.[ct_banpay]) as ct_banpay, M.[ct_loc], M.[ct_sour], M.[ct_kind], M.[ct_tkday], '+@CR+
                  '       Chg_ctclass = Case when M.ct_class = ''1'' and Len(m.ct_no) = 9 then ''�Ȥ�'' when M.ct_class = ''2'' and Len(m.ct_no) = 5 then ''�t��'' end, '+@CR+
                  -- �Ȥ�
                  -- 2016/09/09 Rickliu �]��´�X�s�A�W�[�����q���A�]���N�쥻 2~5 �אּ 2~6
                  -- Chg_BU_NO = Substring(Rtrim(M.[ct_no]), 1, 5), -- �`���q�s��
                  -- �`���q�s��
                  '       Chg_BU_NO = Substring(Rtrim(M.[ct_no]), 1, 6), Chg_ctno_Port_Office = Rtrim(V1.Code_Name), Chg_ctno_CustKind_CustCity = Rtrim(V2.Code_Name), Chg_ctno_CustChain = Rtrim(V3.Code_Name), '+@CR+
                  '       Chg_ct_dept_Name = Rtrim(D1.DP_NAME), Chg_ct_sales_Name = Rtrim(D2.E_NAME), Chg_credit_Name = Rtrim(D3.ct_sname), Chg_busine_Name = Rtrim(P4.Tn_Contact), Chg_loc_Name = Rtrim(P5.Tn_Contact), '+@CR+
                  '       Chg_sour_Name = Rtrim(P6.Tn_Contact), Chg_Customer_kind_Name = Rtrim(P7.Tn_Contact), Chg_payfg_Name = Rtrim(V4.code_name), Chg_porter_Name = Rtrim(P8.tr_name), Chg_pmode_Name = Rtrim(V5.code_name),  '+@CR+
                  '       Chg_fld1_Year = substring(M.[ct_fld1], 1, 4), Chg_fld1_Month = substring(M.[ct_fld1], 6, 2), Chg_fld2_Year = substring(M.[ct_fld2], 1, 4), Chg_fld2_Month = substring(M.[ct_fld2], 6, 2),  '+@CR+
                  -- �~�ȫȤ�ʤj�W��
                  '       Chg_Hunderd_Customer = Case when isnull(P9.Customer, '''') <> '''' then ''Y'' else ''N'' end, Chg_Hunderd_Customer_Name = Case when isnull(P9.Customer, '''') <> '''' then Rtrim(Customer) else '''' end, '+@CR+
                  -- 2014/1/27 Rick �W�[�P�_�O�_�����q�Ȥ�
                  '       Case when ((Upper(substring(m.ct_no, 1, 2)) = ''I9'' Or Upper(substring(m.ct_no, 1, 5)) = ''I1826'' Or  Upper(substring(m.ct_no, 1, 5)) = ''IZ000'') and (m.ct_class = ''1'')) then ''Y'' else ''N'' end as Chg_IS_Lan_Custom, '+@CR+
                  -- 2014/10/09 Rickliu �W�[�P�O���īȤ�μt��
                  '       Case when (LTrim(RTrim(m.ct_name)) = '''' or m.ct_name like ''%����%'' or m.ct_name like ''����%'' or m.ct_name like ''���Ф���%'' or m.ct_name like ''����%'' or m.ct_name like ''����%'' or '+@CR+
                  '            (m.ct_class = ''1'' and Rtrim(m.ct_fld2) <> '''' ) Or (m.ct_class = ''1'' and substring(m.ct_no, 1, 1) Not in (''I'', ''E'', ''B'')) Or '+@CR+
                  '            (m.ct_class = ''1'' and substring(m.ct_no, 2, 1) in (''P'', ''A'', ''T'', ''Z''))) '+@CR+
                  '            then ''Y'' else '''' end as Chg_ct_close, '+@CR+
                  -- 2015/02/03 Rickliu �s�W�ײv
                  '       Chg_rate_date = P10.r_date, Chg_rate = P10.r_rate, '+@CR+
                  -- 2015/02/26 Rickliu �W�[ �Ȥ� ��������, ���s���ĤE�X 1: �ʳf, 2: 3C, 3: �H��, 4:�N�u(OEM)
                  '       Case when m.ct_class = ''1'' and Len(m.ct_no) = 9 then Substring(RTrim(m.ct_no), 9, 1) when m.ct_class = ''1'' and Len(m.ct_no) = 5 then ''0'' else '''' end as Chg_Cust_Sale_Class, '+@CR+
                  '       Case when m.ct_class = ''1'' and Len(m.ct_no) = 5 then ''�t��'' when m.ct_class = ''1'' and Substring(RTrim(m.ct_no), 9, 1) = ''1'' then ''�ʳf'' '+@CR+
                  '            when m.ct_class = ''1'' and Substring(RTrim(m.ct_no), 9, 1) = ''2'' then ''3C'' when m.ct_class = ''1'' and Substring(RTrim(m.ct_no), 9, 1) = ''3'' then ''�H��'' '+@CR+
                  '            when m.ct_class = ''1'' and Substring(RTrim(m.ct_no), 9, 1) = ''4'' then ''�N�u(OEM)'' else '''' end as Chg_Cust_Sale_Class_Name, '+@CR+
                  '       Case when m.ct_class = ''1'' and Len(m.ct_no) = 5 then RTrim(m.ct_sname)+''-�t��'' '+@CR+
                  '            when m.ct_class = ''1'' and Substring(RTrim(m.ct_no), 9, 1) = ''1'' then Replace(Replace(RTrim(m.ct_fld3), ''@'', ''), ''#'', '''')+''-�ʳf'' '+@CR+
                  '            when m.ct_class = ''1'' and Substring(RTrim(m.ct_no), 9, 1) = ''2'' then Replace(Replace(RTrim(m.ct_fld3), ''@'', ''), ''#'', '''')+''-3C'' '+@CR+
                  '            when m.ct_class = ''1'' and Substring(RTrim(m.ct_no), 9, 1) = ''3'' then Replace(Replace(RTrim(m.ct_fld3), ''@'', ''), ''#'', '''')+''-�H��'' '+@CR+
                  '            when m.ct_class = ''1'' and Substring(RTrim(m.ct_no), 9, 1) = ''4'' then Replace(Replace(RTrim(m.ct_fld3), ''@'', ''), ''#'', '''')+''-�N�u(OEM)'' '+@CR+
                  '            else '''' end as Chg_Cust_Sale_Class_sName, '+@CR+
                  -- 2015/03/05 Rickliu �W�[ �Ȥ᦬�����������f�~�������O
                  '       Case when m.ct_class = ''1'' and Substring(RTrim(m.ct_no), 9, 1) = ''1'' then ''A'' when m.ct_class = ''1'' and Substring(RTrim(m.ct_no), 9, 1) = ''2'' then ''B'' '+@CR+
                  '            when m.ct_class = ''1'' and Substring(RTrim(m.ct_no), 9, 1) = ''3'' then ''C'' when m.ct_class = ''1'' and Substring(RTrim(m.ct_no), 9, 1) = ''4'' then ''F'' '+@CR+
                  '            else '''' end as Chg_Cust_Dis_Mapping, '+@CR+
                  '       pcust_update_datetime = getdate(), pcust_timestamp = m.timestamp_column '
    If (@Chk_Tb_Exists = 'N')
       set @strSQL = @strSQL +@CR+
                  '       Into '+@Tb_name

    set @strSQL = @strSQL +@CR+
                  '  From CTE_Q3 m '+@CR+
                  -- �������
                  '       left join SYNC_TA13.dbo.pdept D1 On M.ct_dept = D1.DP_NO '+@CR+
                  -- ���u���
                  '       left join SYNC_TA13.dbo.Pemploy D2 on M.ct_sales = D2.E_NO '+@CR+
                  -- �Ȥ����O
                  '       left join SYNC_TA13.dbo.PCust D3 On M.ct_class = D3.ct_class and M.ct_credit = D3.ct_no '+@CR+
                  -- �ϰ�O
                  '       left join SYNC_TA13.dbo.pattn P4 On P4.tn_class=''4'' and M.ct_busine = P4.TN_NO '+@CR+
                  -- �Ȥ�ӷ��O
                  '       left join SYNC_TA13.dbo.pattn P5 On P5.tn_class=''5'' and M.ct_loc = P5.TN_NO '+@CR+
                  -- �Ȥ����O 
                  '       left join SYNC_TA13.dbo.pattn P6 On P6.tn_class=''6'' and M.ct_sour = P6.TN_NO '+@CR+
                  -- �f�B���q
                  '       left join SYNC_TA13.dbo.pattn P7 On P7.tn_class=''7'' and M.ct_kind = P7.TN_NO '+@CR+
                  '       left join SYNC_TA13.dbo.struc P8 On M.ct_porter = P8.tr_no '+@CR+
                  -- 2015/03/06 Rickliu �ܧ�ĥ� Ori_Xls#Sys_Code �t�ӲĤ@�X���q�O �� �Ȥ�Ĥ@�X���~�P
                  '       left join Ori_Xls#Sys_Code V1 '+@CR+
                  '         On (V1.code_class =''6'' and M.ct_class = ''1'' and V1.Code_Class = ''1'' And Substring(M.[ct_no], 1, 1)= V1.Code_End Collate Chinese_Taiwan_Stroke_CI_AS) '+@CR+
                  '         Or (V1.code_class =''6'' and M.ct_class = ''2'' and V1.Code_Class = ''2'' And Substring(M.[ct_no], 1, 1)= V1.Code_End Collate Chinese_Taiwan_Stroke_CI_AS) '+@CR+
                  '       left join Ori_Xls#Sys_Code V2 '+@CR+
                  -- 2016/09/09 Rickliu �]��´�X�s�A�W�[�����q���A�]���N�쥻 2~5 �אּ 2~6
                  '        On (V2.Code_class = ''1'' and M.ct_class = ''1'' AND Substring(M.[ct_no], 2, 6) COLLATE Chinese_Taiwan_Stroke_CI_AS '+@CR+ 
                  '           between V2.Code_Begin COLLATE Chinese_Taiwan_Stroke_CI_AS and V2.Code_End COLLATE Chinese_Taiwan_Stroke_CI_AS) Or '+@CR+
                  '           (V2.Code_class = ''1'' and M.ct_class = ''2'' And Substring(M.[ct_no], 2, 1) COLLATE Chinese_Taiwan_Stroke_CI_AS '+@CR+
                  '           between V2.Code_Begin COLLATE Chinese_Taiwan_Stroke_CI_AS and V2.Code_End COLLATE Chinese_Taiwan_Stroke_CI_AS) '+@CR+
                  -- 2015/03/06 Rickliu �ܧ�ĥ� Ori_Xls#Sys_Code �Ȥ�ĤE�X�������O�A�ӫ~���������j�����O�X
                  '       left join Ori_Xls#Sys_Code V3 On V3.code_class =''3'' and Substring(M.[ct_no], 9, 1) = V3.Code_Begin Collate Chinese_Taiwan_Stroke_CI_AS'+@CR+
                  -- 2015/03/06 Rickliu �ܧ�ĥ� Ori_Xls#Sys_Code ����ζi������
                  '       left join Ori_Xls#Sys_Code V4 On V3.code_class =''4'' and M.ct_payfg = V4.Code_Begin Collate Chinese_Taiwan_Stroke_CI_AS'+@CR+
                  -- 2015/03/06 Rickliu �ܧ�ĥ� Ori_Xls#Sys_Code �ӫ~ ��/�I�ڤѼƤ覡
                  '       left join Ori_Xls#Sys_Code V5 On V3.code_class =''5'' and M.ct_pmode = V5.Code_Begin Collate Chinese_Taiwan_Stroke_CI_AS'+@CR+
                  -- 2013/2/2 �P���ɳX�ͫ�T�{�O�ϥ�²�٥h����
                  '       left join ori_xls#Hunderd_Customer P9 on m.ct_class=''1'' and m.ct_name collate Chinese_PRC_BIN like ''%''+P9.customer+''%'' collate Chinese_PRC_BIN '+@CR+
                  -- 2015/07/01 NanLiao��g�޿� ���o�ߤ@��
                  '       left join CTE_Q2 P10 on Case when RTrim(Isnull(M.ct_curt_id, '')) = '''' then ''NT'' else M.ct_curt_id end = P10.r_name '
     Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL

    -- 2017/04/06 Rickliu �����ɤ��b�h����
    -- Index: no
    set @Inx_Name = 'no'
    set @Inx_Columns = 'ct_no, ct_class'
    set @Inx_clustered = 'clustered'
    Set @Msg = '�ˮ֯����� ['+@Tb_Name+'.'+@Inx_Name+'('+@Inx_Columns+')]'
    set @strSQL= Replace(Replace(Replace(Replace(@strSQL_Inx, '@Tb_Name', @Tb_Name), '@Inx_Name', @Inx_Name), '@Inx_Columns', @Inx_Columns), '@Inx_clustered', @Inx_clustered)
    Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
    
    -- Index: class
    set @Inx_Name = 'class'
    set @Inx_Columns = 'ct_class'
    set @Inx_clustered = ''
    Set @Msg = '�ˮ֯����� ['+@Tb_Name+'.'+@Inx_Name+'('+@Inx_Columns+')]'
    set @strSQL= Replace(Replace(Replace(Replace(@strSQL_Inx, '@Tb_Name', @Tb_Name), '@Inx_Name', @Inx_Name), '@Inx_Columns', @Inx_Columns), '@Inx_clustered', @Inx_clustered)
    Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
  
    -- Index: ct_payrate
    set @Inx_Name = 'ct_payrate'
    set @Inx_Columns = 'ct_payer, ct_class'
    set @Inx_clustered = ''
    Set @Msg = '�ˮ֯����� ['+@Tb_Name+'.'+@Inx_Name+'('+@Inx_Columns+')]'
    set @strSQL= Replace(Replace(Replace(Replace(@strSQL_Inx, '@Tb_Name', @Tb_Name), '@Inx_Name', @Inx_Name), '@Inx_Columns', @Inx_Columns), '@Inx_clustered', @Inx_clustered)
    Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
  
    -- Index: ctbus
    set @Inx_Name = 'ctbus'
    set @Inx_Columns = 'ct_busine, ct_class'
    set @Inx_clustered = ''
    Set @Msg = '�ˮ֯����� ['+@Tb_Name+'.'+@Inx_Name+'('+@Inx_Columns+')]'
    set @strSQL= Replace(Replace(Replace(Replace(@strSQL_Inx, '@Tb_Name', @Tb_Name), '@Inx_Name', @Inx_Name), '@Inx_Columns', @Inx_Columns), '@Inx_clustered', @Inx_clustered)
    Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
  
    -- Index: ctkind
    set @Inx_Name = 'ctkind'
    set @Inx_Columns = 'ct_kind, ct_class'
    set @Inx_clustered = ''
    Set @Msg = '�ˮ֯����� ['+@Tb_Name+'.'+@Inx_Name+'('+@Inx_Columns+')]'
    set @strSQL= Replace(Replace(Replace(Replace(@strSQL_Inx, '@Tb_Name', @Tb_Name), '@Inx_Name', @Inx_Name), '@Inx_Columns', @Inx_Columns), '@Inx_clustered', @Inx_clustered)
    Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
  
    -- Index: ctloc
    set @Inx_Name = 'ctloc'
    set @Inx_Columns = 'ct_loc, ct_class'
    set @Inx_clustered = ''
    Set @Msg = '�ˮ֯����� ['+@Tb_Name+'.'+@Inx_Name+'('+@Inx_Columns+')]'
    set @strSQL= Replace(Replace(Replace(Replace(@strSQL_Inx, '@Tb_Name', @Tb_Name), '@Inx_Name', @Inx_Name), '@Inx_Columns', @Inx_Columns), '@clustered', @Inx_clustered)
    Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
  
    -- Index: ctname
    set @Inx_Name = 'ctname'
    set @Inx_Columns = 'ct_name'
    set @Inx_clustered = ''
    Set @Msg = '�ˮ֯����� ['+@Tb_Name+'.'+@Inx_Name+'('+@Inx_Columns+')]'
    set @strSQL= Replace(Replace(Replace(Replace(@strSQL_Inx, '@Tb_Name', @Tb_Name), '@Inx_Name', @Inx_Name), '@Inx_Columns', @Inx_Columns), '@Inx_clustered', @Inx_clustered)
    Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
  
    -- Index: sname
    set @Inx_Name = 'sname'
    set @Inx_Columns = 'ct_sname, ct_class'
    set @Inx_clustered = ''
    Set @Msg = '�ˮ֯����� ['+@Tb_Name+'.'+@Inx_Name+'('+@Inx_Columns+')]'
    set @strSQL= Replace(Replace(Replace(Replace(@strSQL_Inx, '@Tb_Name', @Tb_Name), '@Inx_Name', @Inx_Name), '@Inx_Columns', @Inx_Columns), '@Inx_clustered', @Inx_clustered)
    Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
  
  end try
  begin catch
    set @Result = -1
    set @Msg = @Msg+'(���~�T��:'+ERROR_MESSAGE()+')'

    Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Result
  end catch

  Return(@Result)
end
GO
