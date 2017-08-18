USE [DW]
GO
/****** Object:  UserDefinedFunction [dbo].[uFn_GetDate]    Script Date: 08/18/2017 17:18:57 ******/
DROP FUNCTION [dbo].[uFn_GetDate]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
Create Function [dbo].[uFn_GetDate](@Kind Int = 0, @vDate Varchar(10) = '', @AddDate Int=1)
Returns Varchar(1000)
as
begin
  Declare @aVDate DateTime = Getdate()
  Declare @RDate Varchar(2000)= ''
  Declare @CR Varchar(4) = char(13)+char(10)
  Declare @Last_Day Int = 0
  Declare @Month_Last_Day Int = 0
  Declare @Month_Next_Last_Day Int = 0
  
  if @vDate = ''
     Set @vDate = Convert(Varchar(10), getdate())

  Set @RDate = ''
  if IsDate(@vDate) = 1
  begin
     Set @aVDate = Convert(DateTime, @vDate)
     Set @vDate = Convert(Varchar(10), @aVDate, 111)
     
     set @Month_Last_Day = Day(dateadd(dd, -1, DATEADD(mm, DATEDIFF(m, 0, @aVDate) +1, 0)))
     set @Month_Next_Last_Day = Day(dateadd(dd, -1, DATEADD(mm, DATEDIFF(m, 0, @aVDate) +2, 0)))

     -- �ХI�ڤ覡�G�T�w�Ѽ�
     if @Kind = 2
     begin
        /*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-
         �ХI�ڤ覡�G1=�p��ѼơA ����� + n  �ѽХI�� 
               �Ҧp�G��ڤ���� 2/3�A�p��ѼƬ� 19�A�h�ХI�ڤѼƫh�� 3 + 19 = 22 ==> 2/22

         �ХI�ڤ覡�G2=�T�w�ѼơA�Y�ѼƤp��t�Τ�A�h����몺�T�w��ơA�Ϥ��h������T�w���
                �Ҧp�R��ڤ���� 2/3 �A�T�w�ѼƬ� 19�A�h�� 2/19 �鬰�ХI�ڤ�A�Y��ڤ���� 2/21 �h�ХI�ڤ鬰 3/19�C

         2015/02/11 Rickliu �g�߰� ���p�B�p�_��T�{ ��Ѫ���ڻ{�C��몺�b�A�ҥH�Y�K�t�Τ�����멳�B�T�w�дڤ鬰�멳�A�h�{�C��몺�b�A
         �Y���n�אּ �t�Τ�����멳�B�T�w�дڤ鬰�멳 �{�C���몺�b�ɡA�u�n�N���U���P�_�אּ if Day(@aVDate) < @AddDate �Y�i�C
         -*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*/
        if Day(@aVDate) <= @AddDate
        begin
           if isDate(Convert(Varchar(8), @aVDate, 111) + Convert(Varchar(2), @AddDate)) = 1
              set @RDate = Convert(Varchar(8), @aVDate, 111) +Convert(Varchar(2), @AddDate)

           if @RDate = '' And (@AddDate >= @Month_Last_Day) 
              if isDate(Convert(Varchar(8), DateAdd(mm, +1, @aVDate), 111) + Convert(Varchar(2), @AddDate)) = 1
                 set @RDate = Convert(Varchar(8), DateAdd(mm, +1, @aVDate), 111) + Convert(Varchar(2), @AddDate)
              else
                 set @RDate = Convert(Varchar(8), DateAdd(mm, +1, @aVDate), 111) + Convert(Varchar(2), @Month_Next_Last_Day)
        end
        else
        begin
           if (@AddDate < @Month_Next_Last_Day)
              set @Last_Day = Convert(Varchar(2), @AddDate)
           else
              set @Last_Day = @Month_Next_Last_Day
           
           if (@AddDate = 0) -- 2015/07/15 Rickliu �w���дڤ���Y��g 0 �ɡA�h�^�ǿ��~����C
              set @RDate = @vDate
           else 
              set @RDate = substring(Convert(Varchar(10), DateAdd(mm, +1, @aVDate), 111), 1, 8) + Convert(Varchar(2), @Last_Day)
        end
     end
     else
     begin 
       Select @RDate = 
                Case
                  -- �ХI�ڤ覡�G�p��Ѽ�
                  When @Kind = 1 then 
                       Convert(Varchar(10), DateAdd(dd, @AddDate, @aVDate) , 111)

                  /*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-
                   �s�X�W�h�G �� 1 �X�� 1, 2, 3 �N�� �W��, ����, ����
                              �� 2 �X�� �g�B��B�u�B�~
                              �� 3 �X�� �Ĥ@�� �� �̫�@�� �� �Ѽ�
                              
                   �����@�h�Ш̤W�z�s�X�W�h�i��s�C
                  -*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*/

                  --*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-
                  --���o�W�g���}�l���(�w�]�P���鬰���g�}�l��)
                  When @Kind = 111 then Convert(Varchar(20), DATEADD(wk, DATEDIFF(wk, 0, @aVDate),-1), 111)
                  --���o�W�몺�Ĥ@��
                  when @Kind = 121 then Convert(Varchar(20), Dateadd(mm, -1, Dateadd(mm, DATEDIFF(mm, 0, @aVDate), 0)), 111)
                  --���o�W�몺�̫�@��
                  when @Kind = 122 then Convert(Varchar(20), Dateadd(dd, -1, Dateadd(mm,DATEDIFF(mm, 0, @aVDate),0)), 111)
                  --���o�W�u���Ĥ@��
                  when @Kind = 131 then Convert(Varchar(20), DATEADD(qq, -1, DATEADD(qq, DATEDIFF(qq, 0, DATEADD(qq, 0, @aVDate)), 0)), 111)
                  --���o�W�u���̫�@��
                  when @Kind = 132 then Convert(Varchar(20), DateAdd(mm, +3, DATEADD(qq, -1, DATEADD(qq, DATEDIFF(qq, 0, DATEADD(qq, 0, @aVDate)), 0))) -1, 111)
                  --���o�h�~���Ĥ@��
                  When @Kind = 141 then Convert(Varchar(20), DATEADD(yy, -1, DATEADD(yy, DATEDIFF(yy, 0, @aVDate),0)), 111)
                  --���o�h�~���̫�@��
                  When @Kind = 142 then Convert(Varchar(20), DATEADD(yy, DATEDIFF(yy, 0, @aVDate), -1), 111)
                  --���o�W��Ѽ�
                  when @Kind = 151 then Convert(Varchar(20), Day(dateadd(dd, -1, DATEADD(mm, DATEDIFF(m, 0, @aVDate)+0, 0))))
                  --���o�W�u�Ѽ�
                  when @Kind = 152 then Convert(Varchar(20), DateDiff(dd, DATEADD(qq, -1, DATEADD(qq, DATEDIFF(qq, 0, DATEADD(qq, 0, @aVDate)), 0)), DateAdd(mm, +3, DATEADD(qq, -1, DATEADD(qq, DATEDIFF(qq, 0, DATEADD(qq, 0, @aVDate)), 0))) -1) +1)



                  --*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-
                  --���o���g���}�l���(�w�]�P���鬰���g�}�l��)
                  When @Kind = 211 then Convert(Varchar(20), DATEADD(wk, DATEDIFF(wk, 0, @aVDate),-1), 111)
                  --���o���몺�Ĥ@��
                  when @Kind = 221 then Convert(Varchar(20), Dateadd(mm, DATEDIFF(mm, 0, @aVDate), 0), 111)
                  --���o���몺�̫�@��
                  when @Kind = 222 then Convert(Varchar(20), dateadd(dd, -1, DATEADD(mm, DATEDIFF(m, 0, @aVDate) +1, 0)), 111)
                  --���o���u���Ĥ@��
                  when @Kind = 231 then Convert(Varchar(20), DATEADD(qq, +0, DATEADD(qq, DATEDIFF(qq, 0, DATEADD(qq, 0, @aVDate)), 0)), 111)
                  --���o���u���̫�@��
                  when @Kind = 232 then Convert(Varchar(20), DateAdd(mm, +3, DATEADD(qq, +0, DATEADD(qq, DATEDIFF(qq, 0, DATEADD(qq, 0, @aVDate)), 0))) -1, 111)
                  --���o���~���Ĥ@��
                  when @Kind = 241 then Convert(Varchar(20), DATEADD(yy, DATEDIFF(yy, 0, @aVDate), 0), 111)
                  --���o���~���̫�@��
                  When @Kind = 242 then Convert(Varchar(20), DATEADD(yy, DATEDIFF(yy, 0, @aVDate)+1, -1), 111)
                  --���o����Ѽ�
                  when @Kind = 251 then Convert(Varchar(20), Day(dateadd(dd, -1, DATEADD(mm, DATEDIFF(m, 0, @aVDate)+1, 0))))
                  --���o���u�Ѽ�
                  when @Kind = 252 then Convert(Varchar(20), DateDiff(dd, DATEADD(qq, +0, DATEADD(qq, DATEDIFF(qq, 0, DATEADD(qq, 0, @aVDate)), 0)), DateAdd(mm, +3, DATEADD(qq, +0, DATEADD(qq, DATEDIFF(qq, 0, DATEADD(qq, 0, @aVDate)), 0))) -1) +1)



                  --*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-
                  --���o�U�g���}�l���(�w�]�P���鬰���g�}�l��)
                  When @Kind = 311 then Convert(Varchar(20), DATEADD(wk, DATEDIFF(wk,0, @aVDate), -1+7), 111)
                  --���o�U�몺�Ĥ@��
                  when @Kind = 321 then Convert(Varchar(20), Dateadd(mm, 1, Dateadd(mm, DATEDIFF(mm, 0, @aVDate), 0)), 111)
                  --���o�U�몺�̫�@��
                  when @Kind = 322 then Convert(Varchar(20), dateadd(dd, -1, DATEADD(mm, DATEDIFF(m, 0, @aVDate) +2, 0)), 111)
                  --���o�U�u���Ĥ@��
                  when @Kind = 331 then Convert(Varchar(20), DATEADD(qq, +1, DATEADD(qq, DATEDIFF(qq, 0, DATEADD(qq, 0, @aVDate)), 0)), 111)
                  --���o�U�u���̫�@��
                  when @Kind = 332 then Convert(Varchar(20), DateAdd(mm, +3, DATEADD(qq, +1, DATEADD(qq, DATEDIFF(qq, 0, DATEADD(qq, 0, @aVDate)), 0))) -1, 111)
                  --���o���~���Ĥ@��
                  When @Kind = 341 then Convert(Varchar(20), DATEADD(yy, 1, DATEADD(yy, DATEDIFF(yy, 0, @aVDate),0)), 111)
                  --���o���~���̫�@��
                  When @Kind = 342 then Convert(Varchar(20), DATEADD(yy, DATEDIFF(yy, 0, @aVDate)+2, -1), 111)
                  --���o�U��Ѽ�
                  when @Kind = 351 then Convert(Varchar(20), Day(dateadd(dd, -1, DATEADD(mm, DATEDIFF(m, 0, @aVDate)+2, 0))))
                  --���o�U�u�Ѽ�
                  when @Kind = 352 then Convert(Varchar(20), DateDiff(dd, DATEADD(qq, +1, DATEADD(qq, DATEDIFF(qq, 0, DATEADD(qq, 0, @aVDate)), 0)), DateAdd(mm, +3, DATEADD(qq, +1, DATEADD(qq, DATEDIFF(qq, 0, DATEADD(qq, 0, @aVDate)), 0))) -1) +1)
                  else 
                    ''
                end
     end
  end
  
  
  Return(@RDate)
end
GO
