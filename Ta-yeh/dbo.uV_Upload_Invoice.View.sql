USE [DW]
GO
/****** Object:  View [dbo].[uV_Upload_Invoice]    Script Date: 08/18/2017 17:43:39 ******/
DROP VIEW [dbo].[uV_Upload_Invoice]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
Create View [dbo].[uV_Upload_Invoice]
as
  select '###_Invoice_Master_###' as '###_Invoice_Master_###',
         Convert(Varchar(10), isnull(in_no, '')) as M_in_no, -- �o�����X
         Convert(Varchar(10), in_date, 111) as M_in_date, -- �o�����
         -- 2015/01/29 Rickliu �ѩ��V�t�ΨõL�q�l�o���ﶵ�A�S�[�W RPT_13010 ����@�v���O�ιq�l�o���Ҧ��W�ǡA�ҥH�T�w�g 07
         /*
         case
           when in_frm = '31' then '01' -- �T�p���o��
           when in_frm = '32' then '02' -- �G�p���o��
           when in_frm = '35' then '03' -- �T�p�����Ⱦ��o��
           else ''
         end as M_in_frm, -- �o�����O
         */
         '07' as M_in_frm, -- �o�����O
         Convert(Varchar(10), isnull(in_bno, '')) as M_in_bno, -- �R��Τ@�s��
         convert(char(1), in_tcd) as M_in_tcd, -- �ҵ|�O 1:���|, 2:�s�|�v, 3:�K�|
         Convert(Varchar(30),
           case
             when isnull(in_amt, 0) = 0 then 0
             else Convert(Float, Round(in_tax / in_amt, 2) * 100)
           end) as M_in_tax, -- �|�v
         '' as M_in_bound_fg, -- �q���覡���O 1: �D�g�����X�f, 2:�g�����X�f
         '' as M_bay_role_fg, -- ��~�H������O
         Convert(Varchar(30), Convert(Money, isnull(in_amt, 0))) as M_in_amt, -- ������B
         '1' as M_Exchange_rate, -- �ײv
         'TWD' as M_Currency, -- ���O
         -- �J�}���O�G�Y���J�}���O�A�h��*�A2014/11/04 ���p���P�q�h��|���A��ܤ@����~�H�O�Τ��쪺�A�ҥH�i�H�ť�
         -- Convert(Varchar(1), Replace(Rtrim(Isnull(in_scd, '')), 'A', '*')) as M_in_scd, -- �J�}���O *:�J�}
         '' as M_in_scd, -- �J�}���O
         '0' as M_sale_kind, -- �P�����O 0:�@��P�� 1:�v�Ұs�� 2:�T�w�겣 3:�g�a
         '' as M_remark, -- ���O
         Convert(Varchar(20), m.in_ctno) as M_in_ctno, -- �Ȥ�s��
/*
         '###_Invoice_Detail_###' as '###_Invoice_Detail_###',
         Convert(Varchar(10), isnull(in_no, '')) as D_in_no, -- �o�����X
         Convert(Varchar(20), Rtrim(Isnull(id_sknm, ''))) as D_id_sknm, -- �f�~�s��
         Convert(Varchar(20), Replace(Replace(Replace(Replace(Replace(Convert(Varchar(100), Rtrim(Isnull(id_skno, ''))), '<', ''), '>', ''), '%', ''), '&', ''), '!', '')) as D_id_skno, -- �f�~�W��
         '' as D_Related_number, -- �������X
  
         Convert(Varchar(30), Convert(Money, isnull(id_price, 0))) as D_id_price1, -- ���1
         Convert(Varchar(10), Rtrim(isnull(id_unit, ''))) as D_id_unit1, -- ���1
         Convert(Varchar(30), Convert(Money, isnull(id_qty, 0))) as D_id_qty1, -- �ƶq1
  
         '' as D_id_price2, -- ���2
         '' as D_id_unit2, -- ���2
         '' as D_id_qty2, -- �ƶq2
         
         '' as D_remark, -- ��@�����O
*/
         /**************************************************************************************************************************************
          2014/12/30 Rickliu �L�Ҫ�ܼW�[��~�H�P���W�ǰ�|����Ʈ榡��
          Upload Bessiness Sale Invoice Data ==> BSI          
          
          Field NO		Field Name				Size	Begin	End
          Field(01)		�榡�N��				X(002)		1	  2
          Field(02)		�ӳ���~�H�|�y�s��		X(009)	    3	 11
          Field(03)		�y����					X(007)	   12	 18
          Field(04)		��Ʃ��ݦ~��			9(005)	   19    23
          Field(05)		�R���H�Τ@�s��			X(008)	   24	 31
          
         **************************************************************************************************************************************/
         '### ��~�|�P����Ʈ榡�� ###' as '### ��~�|�P����Ʈ榡�� ###',
         -- Field(01) �榡�N��
         in_frm as UBI_Frm, 
         -- Field(02) �O�d�� �ӳ���~�H�|�y�s�� Tax Registration Number
         '#AAAAAAA#' as UBI_Trn, 
         -- Field(03) �y����
         '#BBBBB#' as UBI_NO, -- �y����
         -- Field(04-01) ��Ʃ��ݦ~
--         Convert(Varchar(4), year(in_date)-1911) Collate Chinese_PRC_BIN as UBI_cYear, -- �o�����ݦ~(����~), �Фŧ��w��
         -- Field(04-02) ��Ʃ��ݤ�
--         Convert(Varchar(2), month(in_date)) Collate Chinese_PRC_BIN as UBI_Month, -- �o�����ݤ�
         Convert(Varchar(4), Convert(Int, Substring(in_yrmn, 1, 4)) -1911) Collate Chinese_PRC_BIN as UBI_cYear, -- �o�����ݦ~(����~), �Фŧ��w��
         -- Field(04-02) ��Ʃ��ݤ�
         Substring(in_yrmn, 5, 2) Collate Chinese_PRC_BIN as UBI_Month, -- �o�����ݤ�
         
         -- Field(05) �R���H�νs
         Convert(Char(8), Convert(Varchar(10), isnull(in_bno, ''))) as UBI_Buy_GUI,
         -- Field(06) �O�d�� �P��H�νs Government Uniform Invoice(GUI) Number ==> GUI Number
         '#CCCCCC#' as UBI_Sale_GUI,
         -- Field(07) �Τ@�o�� GUI
         Reverse(Convert(Char(10), Reverse(RTrim(Convert(Varchar(10), isnull(in_no, '')))))) as UBI_GUI,
         -- Field(13) �P����B 
         Substring(Convert(Varchar(30), Convert(Money, isnull(in_amt, 0) + 1000000000000)), 2, 12) as UBI_Sale_AMT,
         -- Field(14) �ҵ|�O
         convert(char(1), in_tcd) as UBI_Tcd,
         -- Field(15) ��~�|�|�B
         Substring(Convert(Varchar(30), Convert(Money, isnull(in_tax, 0) + 10000000000)), 2, 10) as UBI_Tax,
         -- Field(16) ����N��
         convert(char(1), in_mcd) as UBI_Mcd,
         -- Field(17) �ť�
         Space(5) as UBI_Keep1,
         -- Field(18) �S�ص|�B�|�v
         Space(1) as UBI_Keep2,
         -- Field(19) �J�[���O
         convert(char(1), in_scd) as UBI_Scd,
         -- Field(20) �q���覡���O
         Space(1) as UBI_Keep3,
         case
           when m.in_frm in ('31', '32', '33', '34', '35')
           then Convert(Char(2), in_frm) + -- �榡�N��
                '#AAAAAAA#'+ -- �O�d�� �ӳ���~�H�|�y�s��
                '#BBBBB#'+ -- �y����
                Convert(Char(3), Convert(Varchar(4), substring(in_yrmn, 1, 4)-1911) Collate Chinese_PRC_BIN) + -- ��Ʃ��ݦ~ ����~
                Substring(Convert(Char(3), Convert(Varchar(2), substring(in_yrmn, 5, 2)) Collate Chinese_PRC_BIN +100), 2, 2)+ -- ��Ʃ��ݤ��
                Convert(Char(8), Convert(Varchar(10), isnull(in_bno, '')))+ -- �R���H�νs         
                '#CCCCCC#'+ -- �O�d�� �P��H�νs
                Reverse(Convert(Char(10), Reverse(RTrim(Convert(Varchar(10), isnull(in_no, ''))))))+ -- �o�����X
                Substring(Convert(Varchar(30), Convert(Money, isnull(in_amt, 0) + 1000000000000)), 2, 12)+ -- �P����B
                convert(char(1), in_tcd)+ -- �ҵ|�O
                Substring(Convert(Varchar(30), Convert(Money, isnull(in_tax, 0) + 10000000000)), 2, 10)+ -- ��~�|�|�B
                convert(char(1), in_mcd)+ -- 
                Space(5)+
                Space(1)+
                convert(char(1), in_scd)+
                Space(1)
                else ''
         end as Business_Sale_Invoice_OutData
    from tyeipdbs2.lytdbta13.dbo.pinvo m
--          left join tyeipdbs1.lytdbta13.dbo.pinvodt d
--            on m.in_prono = d.id_prono
   where 1=1
     --and m.in_tcd <> 'F'
     and m.in_no <> ''
     and m.in_frm in ('31', '32', '33', '34', '35')


/***[�o���D��]*****************************************************************************************************************************************

�o�����X�G�t�o���r�y�@10�X

�o������G�п�J�褸�~

�o�����O�G
01�G�T�p���o��
02�G�G�p���o��
03�G�G�p�����Ⱦ��o��
04�G�S�ص|�B�o��
05�G�q�l�p����o��
06�G�T�p�����Ⱦ��o��

�ҵ|���O�N�X�G
1�G���|
2�G�s�|�v
3�G�K�|

�ײv�G�K��%�r�ˡ]�ȿ�J�Ʀr�^

�q���覡���O�G
1�G�D�g�����X�f
2�G�g�����X�f

������B�G(���12�Ϥp��4)

�ײv�G(���8�Ϥp��4)

���O�G
TWD�G�s�x��
USD�G����
GBP�G�^��
DEM�G�w�갨�J
AUD�G�D�j�Q�ȹ�
HKD�G���
SGD�G�s�[�Y��
CAD�G�[���j��
CHF�G��h�k��
MYR�G���Ӧ�ȹ�
FRF�G�k��k��
BEF�G��Q�ɪk��
SEK�G����
JPY�G���
ITL�G�q�j�Q����
THB�G����
NTD�GCURRENCY_NTD
EUR�G�ڬw�@�P�f��
NZD�G�æ�����

�J�}���O�G�Y���J�}���O�A�h��*�A2014/11/04 ���p���P�q�h��|���A��ܤ@����~�H�O�Τ��쪺�A�ҥH�i�H�ť�

�P�����O�N�X�G
0�G�@��P��
1�G�v�ϰs��
2�G�T�w�겣
3�G�g�a

-- [�o��������]**********************************************************************************************************************************
�o�����X�G���P�o���D�ɸ��X�ۦP ���� 10
�~�W�s���G�u��^�� ���� 20
�o���~�W�G���� 256
�������X�G���� 20
����G���� 17
���G���� 6
�ƶq�G���� 17
���2�G���� 17
���2�G���� 6
�ƶq2�G���� 17
��@���Ƶ��G���� 40

--**********************************************************************************************************************************************

�Ƶ��G
1.�o���D�ɤΩ����ɬ��קK��ǰO�����D���͡A�бĥΤ�r�榡
2.�פJ�o�����X�̦h���i�W�L12�i�o���A�C�i�o�������ɭ���999��
3.�D�ɵo�����X�P�����ɵo�����X�ۦP�����P�@�i�o��
4. �ҵ|�O���s�|�v�ɡA�~�ݶ�g�q���覡���O�B������B�B�ײv�ι��O
5.�o���D�ɶ��H�o�����X�ƧǡA�B�W�@�i�o�����}�ߤ�����o�j��U�@�i���o������A�H�קK���������D
6.�ФŨϥίS��r���]�Ҧp�G<, >, %,&,!�^
7.���2�μƶq2�Ȩѿ��K�o���ϥ�
8.���O�ζײv���ȨѳƵ��ϥΡA����B��\��
9.�o�������ɤ�����γ��2���ХH�s�x����J
*************************************************************************************************************************************************/
GO
