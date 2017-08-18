USE [DW]
GO
/****** Object:  StoredProcedure [dbo].[uSP_Imp_SCM_Comp_EC_Consign_Order_MyFone]    Script Date: 07/24/2017 14:43:59 ******/
DROP PROCEDURE [dbo].[uSP_Imp_SCM_Comp_EC_Consign_Order_MyFone]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE Proc [dbo].[uSP_Imp_SCM_Comp_EC_Consign_Order_MyFone]
as
begin
  /***********************************************************************************************************
     Tw_Mobile ��x�Գ��b��öi���V���
  ***********************************************************************************************************/
  Declare @Proc Varchar(50) = 'uSP_Imp_SCM_Comp_EC_Consign_Order_MyFone'
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
       
  --���ͳf�����A2014/6/23 �p�_��ܭY XLS �����Ӵ����� 4/1~4/30 �h�O 3������b�A�ӭn���V 3/21~4/20 ��b������P�h�f��
  set @Last_Date = '20'
  set @sp_date = Substring(Convert(Varchar(10), getdate(), 111), 1, 8)+@Last_Date
  set @sd_date = @sp_date

  Declare @CompanyName Varchar(100), @CompanyLikeName Varchar(100), @Rm Varchar(100), @sp_maker Varchar(20), @str Varchar(200)
  Declare @Pos Int
  
  Set @CompanyName = '�x�W�j���j' --> �ФŶ��ܰ�
  Set @CompanyLikeName = '%'+@CompanyName+'%'
  Set @Rm = '�t�ζפJ'+@CompanyName
  Set @sp_maker = 'Admin'
  Set @Collate = 'Collate Chinese_Taiwan_Stroke_CI_AS'
  set @ct_no ='I90240011'

  Set @xls_head = 'Ori_Xls#'
  Set @TB_head = 'Comp_EC_Consign_Order_MyFone'
  Set @TB_head_tmp = @TB_head+'_tmp'
  Set @TB_xls_Name = @xls_head+@TB_Head -- Ex: Ori_Xls#Comp_EC_Consign_Order_MyFone -- Excel ���ɸ��

  -- Check �פJ�ɮ׬O�_�s�b
  IF Not EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].['+@TB_xls_Name+']') AND type in (N'U'))
  begin
print '1'
     --Send Message
     set @Cnt = @Errcode
     set @strSQL = ''
     Set @Msg = '�~�� Excel �פJ��ƪ� ['+@TB_xls_Name+']���s�b�A�פ�i�����ɧ@�~�C'
     Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt
     -- 2013/11/28 �W�[���Ѧ^�ǭ�
     close Cur_Tw_Mobile_DataKind
     deallocate Cur_Tw_Mobile_DataKind
     Return(@Errcode)
  end
    
  IF Exists(SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].['+@TB_head_tmp+']') AND type in (N'U'))
  begin
print '2'
    Set @Msg = '�M�� '+@CompanyName+' ��b��Ȧs��ƪ� ['+@TB_head_tmp+']�C'

    Set @strSQL = 'DROP TABLE [dbo].['+@TB_head_tmp+']'
    Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
  end
  
  IF Exists(SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].['+@TB_head+']') AND type in (N'U'))
  begin
print '3'
    Set @Msg = '�M�� '+@CompanyName+' ��b�椶����ƪ� ['+@TB_head+']�C'

    Set @strSQL = 'DROP TABLE [dbo].['+@TB_head+']'
    Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
  end

print '4'
  Set @Msg = '���X�X�h�f��ơC'
  Set @strSQL = 'Declare @YM Varchar(7) '+@CR+
                'Declare @B_Pdate Varchar(10) '+@CR+
                'Declare @E_Pdate Varchar(10) '+@CR+
                ' '+@CR+
                'select Top 1 @YM = Substring(F4, 1, 7)  '+@CR+
                '  from '+@TB_xls_name+@CR+
                ' where F3 like ''%���Ӵ���%'' '+@CR+
                ' '+@CR+
                'Exec uSP_Get_Cust_Pdate_Range '''+@ct_no+''', @YM, @b_pdate output, @e_pdate output '+@CR+
                ' '+@CR+
                'select F1, '+@CR+ -- ����
                '       F2, F3, F4, F5, '+@CR+ -- ������	�X�f���	�e�F���	�h�f���
                '       rtrim(F6) as F6, F7, F8, '+@CR+ -- �q��s��	�X�f�渹	��P����
                '       F9, F10, F11, F12, F13, '+@CR+ -- �����Ӱӫ~�f��	�ӫ~��l�f��	�ӫ~�f��	�ӫ~�W��	�˦�
                '       Convert(float, Replace(F14, '','', '''')) as F14, '+@CR+ -- �ӫ����
                '       Convert(float, Replace(F15, '','', '''')) as F15, '+@CR+ -- �ƶq
                '       Convert(float, Replace(F16, '','', '''')) as F16, '+@CR+ -- �ѳf��(���|)
                '       Convert(float, Replace(F17, '','', '''')) as F17, '+@CR+ -- �ѳf��(�t�|)
                '       Convert(float, Replace(F18, '','', '''')) as F18, '+@CR+  -- �p�p(���|)
                '       @B_Pdate as B_sp_pdate, @E_PDate as E_sp_pdate, '+@CR+  -- ��ڰ_��
                '       '''' as F19, '''' as F20, '+@CR+
                '       Getdate() as Create_datetime '+@CR+
                '       into '+@TB_head_tmp+@CR+
                '  from '+@TB_xls_name+@CR+
                ' where Isnumeric(F18) = 1'
  Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
                
print '5'
  -- �ѩ� F20 ���S���ϥΩҥH���ӷ�@�X�h�f�аO
  
  Set @Msg = '�N�X�f�ΰh�f��Ƥ��H�аO�C'
  Set @strSQL = 'update '+@TB_head_tmp+@CR+
                '   set F20 = ''2'' '+@CR+
                '  from '+@TB_head_tmp+@CR+
                ' where F1 like ''�X%'' '+@CR+
                ' '+@CR+
                'update '+@TB_head_tmp+@CR+
                '   set F20 = ''3'' '+@CR+
                --'       F18 = F18'+@CR+
                '  from '+@TB_head_tmp+@CR+
                ' where F1 like ''�h%'' '+@CR+
                ' '
               
  Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL

  --�b���i����ӫ~�s���A�|�����U�s���H�ΰӫ~�򥻸���ɶi����@�~
print '6'
  Set @Msg = '�i�� XLS ��� ��V �ӫ~�s���üg�J ['+@TB_OD_Name+']��ƪ�C'
  set @strSQL = 'select ''XLS'' as master, '+@CR+
                '       XLS_YM,'+@CR+
                '       Date_Range, '+@CR+
                '       rtrim(slip_fg) as slip_fg, '+@CR+
                '       case when rtrim(slip_fg) = ''2'' then ''�X'' else ''�h'' end as xls_slip_fg_name, '+@CR+
                '       Convert(VarChar(20), rtrim(order_no)) as order_no, '+@CR+
                '       Convert(Varchar(255), isnull(Chg_sp_slip_name, '''')) as sd_spec,'+@CR+
                '       Round(isnull(xls_tot, 0), 0) as xls_tot, isnull(chg_sd_stot, 0) as chg_sd_stot, '+@CR+
                '       0 as order_total, '+@CR+
                '       isnull(cnt, 0) as order_cnt, '+@CR+
                '       case '+@CR+
                '         when xls_tot = chg_sd_stot then ''N'' '+@CR+
                '         else ''Y'' '+@CR+
                '       end as Equal '+@CR+
                '       into '+@TB_head+@CR+
                '  from (select Rtrim(F6) as Order_no, Convert(Varchar(7), E_sp_pdate, 111) as XLS_YM, B_sp_pdate, E_sp_pdate, '+@CR+
                '               F20 as slip_fg, sum(F18*1.05) as xls_tot, count(1) as cnt  '+@CR+
                '          from '+@TB_head_tmp+' '+@CR+
                '         group by F20, F6, Convert(Varchar(7), E_sp_pdate, 111), B_sp_pdate, E_sp_pdate '+@CR+
                '       ) m '+@CR+
                '       left join '+@CR+
                '       (select sd_slip_fg, chg_sd_slip_fg, '+@CR+
                '               sd_spec, Chg_sp_slip_name, '+@CR+
                '               Round(sum(chg_sd_stot * 1.05), 0) as chg_sd_stot '+@CR+
                '          from Fact_sslpdt m '+@CR+
                '         where sd_class = ''1'' '+@CR+
                '           and sd_ctno ='''+@ct_no+''' '+@CR+
                '           and sd_spec like ''5%'' '+@CR+

                '           and sp_pdate >= '+@CR+
                '               (select top 1 B_sp_pdate from '+@TB_head_tmp+') '+@CR+
                '           and sp_pdate <= '+@CR+
                '               (select top 1 E_sp_pdate from '+@TB_head_tmp+') '+@CR+

                '         group by sd_slip_fg, chg_sd_slip_fg, sd_spec, Chg_sp_slip_name '+@CR+
                '       ) d '+@CR+
                '        on Order_no '+@Collate+' = Rtrim(sd_spec) '+@Collate+' '+@CR+
                '       and slip_fg '+@Collate+' = sd_slip_fg '+@Collate+''+@CR+
                '       Cross join '+@CR+
                '       (select Top 1 B_SP_PDate+'' ~ ''+ E_SP_PDate as Date_Range '+@CR+
                '          from '+@TB_head_tmp+@CR+
                '       ) D1 '+@CR+
                ' where 1=1 '+@CR+
                ' order by Order_no '
  
  Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL

print '7'
  Set @Msg = '�i�� ��V ��� XLS �ӫ~�s���üg�J ['+@TB_OD_Name+']��ƪ�C'
  set @strSQL = 'insert into '+@TB_head+@CR+
                'select ''LYS'' as master, '+@CR+
                '       isnull(Chg_sp_pdate_ym, xls_ym) as XLS_YM,  '+@CR+
                '       Date_Range, '+@CR+
                '       isnull(sd_slip_fg, slip_fg) as sd_slip_fg, '+@CR+
                '       case when rtrim(isnull(sd_slip_fg, slip_fg)) = ''2'' then ''�X'' else ''�h'' end as xls_slip_fg_name, '+@CR+
                '       Case '+@CR+
                '         when order_no '+@Collate+' is null then rtrim(substring(sd_spec, 1, 20)) '+@Collate+' '+@CR+
                '         else rtrim(substring(order_no, 1, 20)) '+@Collate+' '+@CR+
                '       end as order_no, '+@CR+
                '       Convert(Varchar(255), isnull(Chg_sp_slip_name, '''')) as sd_spec, '+@CR+
                '       Round(isnull(xls_tot, 0), 0) as xls_tot, isnull(chg_sd_stot, 0) as chg_sd_stot, '+@CR+
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

                '           and sd_spec like ''5%'' '+@CR+
                '         group by chg_sp_pdate_YM, sd_slip_fg, chg_sd_slip_fg, sd_spec, Chg_sp_slip_name '+@CR+
                '       ) m '+@CR+
                '       Left join '+@CR+
                '       (select Rtrim(F6) as Order_no, Convert(Varchar(7), E_sp_pdate, 111) as xls_YM, B_sp_pdate, E_sp_pdate,'+@CR+
                '               F20 as slip_fg, sum(F18*1.05) as xls_tot, count(1) as cnt  '+@CR+
                '          from '+@TB_head_tmp+' '+@CR+
                '         group by F20, F6, Convert(Varchar(7), E_sp_pdate, 111), B_sp_pdate, E_sp_pdate '+@CR+
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
  Set @Msg = '�̺��q�渹��s�̫ᵲ�l���B�C'
  set @strSQL =  -- ���������������q�s���A�ñN��ڽs�����H��J
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
                '           and isnull(m.sd_spec, '''')= '''' '+@CR+
                '         group by m.order_no, slip_fg, d.Chg_sp_slip_name '+@CR+
                '       ) d '+@CR+
                '        on m.slip_fg '+@Collate+' = d.slip_fg '+@Collate+' '+@CR+
                '       and m.order_no '+@Collate+'= d.order_no '+@Collate+' '+@CR+
                '       and isnull(m.sd_spec, '''')= '''' '+@CR+
                '       and isnull(d.sd_spec, '''')<> '''' '+@CR+
                ' '+@CR+
                
                -- �̺��q�渹��s�̫ᵲ�l���B
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
               
                -- 2015/09/02 Rickliu �ܧ���~���O(PS, * �h�N��������`���)
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
