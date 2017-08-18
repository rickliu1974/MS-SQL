USE [DW]
GO
/****** Object:  StoredProcedure [dbo].[uSP_Exec_IMP_Orders]    Script Date: 08/18/2017 17:43:40 ******/
DROP PROCEDURE [dbo].[uSP_Exec_IMP_Orders]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE Procedure [dbo].[uSP_Exec_IMP_Orders]
as
begin
  Declare @Proc Varchar(50) = 'uSP_Exec_IMP_Orders'
  Declare @CR Varchar(4) =Char(13)+Char(10)
  Declare @Body_Msg NVarchar(Max) = '', @Run_Proc_Msg NVarchar(Max) = '', @Run_Msg NVarchar(1000) = ''
  Declare @Cnt_xls int = 0, @Cnt_txt int = 0, @Run_Err int = 0, @Errcode int = -1
  Declare @Cnt_Err int =0,  @Cnt_Suss int = 0
  Declare @Result int = 0
  Declare @Proce_SubName NVarchar(Max) = '�~���ɮ׶פJ�� DW ��Ʈw�{��'
  Declare @BTime DateTime, @ETime DateTime
  Declare @Hour Varchar(2) = ''
  
  set @BTime = GetDate()
  
  set @Hour =(select left(convert(varchar(20), @BTime, 114), 2))
  
  if (@Hour in ('12', '00')) --< �]�w�T��B��ɶ� (24 �p�ɨ�)
  Begin
     Set @body_Msg = @body_Msg + '['+@Proc+'] �w�Q�]�w '+@Hour+' �ɸT��B��!!'+@CR
     Goto End_Exit
  end
  
  Begin Try
    Exec @Cnt_xls = DW.dbo.uSP_Sys_Imp_xls_to_db
    Exec @Cnt_txt = DW.dbo.uSP_Sys_Imp_txt_to_db

    set @Cnt_Err = @Cnt_xls + @Cnt_txt
    if (@Cnt_Err in (0, -2))
    begin
       Set @body_Msg = @body_Msg + '�L�����Ƥ��i������!!'+@CR
       Goto End_Exit
    end
    
    if @Cnt_xls <> @Errcode
    begin
       set @Body_Msg = @Body_Msg + REPLICATE('#', 100)+@CR
       set @Body_Msg = @Body_Msg + '���� uSP_Imp_xls_to_db�i�@�פJ Excel �ɮ׼ơj...('+CONVERT(Varchar(100), @Cnt_xls)+'��)�C'+@CR
       set @Body_Msg = @Body_Msg + REPLICATE('#', 100)+@CR
 
       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       Begin Try
          set @Run_Proc_Msg = 'EC01.���� uSP_Imp_SCM_EC_Books_Order�i�իȨ� EC��J�j... '
          Exec @Run_Err = DW.dbo.uSP_Imp_SCM_EC_Books_Order
          if @Run_Err = @Errcode
          begin
             set @Run_Msg = '����'
             set @Cnt_Err = @Cnt_Err +1
          end
          else
          begin
            set @Run_Msg = '���\'
            set @Cnt_Suss = @Cnt_Suss +1
          end
       End Try
       Begin Catch 
          set @Run_Msg = '����'
          set @Cnt_Err = @Cnt_Err +1
       End Catch

       set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
       set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
       Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1


       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       Begin Try
          set @Run_Proc_Msg = 'EC02.���� uSP_Imp_SCM_EC_CrazyMike_Order�i�ƨg��� EC��J�j... '
          Exec @Run_Err = DW.dbo.uSP_Imp_SCM_EC_CrazyMike_Order
          if @Run_Err = @Errcode
          begin
             set @Run_Msg = '����'
             set @Cnt_Err = @Cnt_Err +1
          end
          else
          begin
            set @Run_Msg = '���\'
            set @Cnt_Suss = @Cnt_Suss +1
          end
       End Try
       Begin Catch 
          set @Run_Msg = '����'
          set @Cnt_Err = @Cnt_Err +1
       End Catch

       set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
       set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
       Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1     

       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       Begin Try
          set @Run_Proc_Msg = 'EC03.���� uSP_Imp_SCM_EC_ETmall_Order�i�F���ʪ� EC��J�j... '
          Exec @Run_Err = DW.dbo.uSP_Imp_SCM_EC_ETmall_Order
          if @Run_Err = @Errcode
          begin
             set @Run_Msg = '����'
             set @Cnt_Err = @Cnt_Err +1
          end
          else
          begin
            set @Run_Msg = '���\'
            set @Cnt_Suss = @Cnt_Suss +1
          end
       End Try
       Begin Catch 
          set @Run_Msg = '����'
          set @Cnt_Err = @Cnt_Err +1
       End Catch

       set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
       set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
       Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1 

       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       Begin Try
          set @Run_Proc_Msg = 'EC04.���� uSP_Imp_SCM_EC_ForMall_Order�i��MALL EC��J�j... '
          Exec @Run_Err = DW.dbo.uSP_Imp_SCM_EC_ForMall_Order
          if @Run_Err = @Errcode
          begin
             set @Run_Msg = '����'
             set @Cnt_Err = @Cnt_Err +1
          end
          else
          begin
            set @Run_Msg = '���\'
            set @Cnt_Suss = @Cnt_Suss +1
          end
       End Try
       Begin Catch 
          set @Run_Msg = '����'
          set @Cnt_Err = @Cnt_Err +1
       End Catch

       set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
       set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
       Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1 

       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       Begin Try
          set @Run_Proc_Msg = 'EC05.���� uSP_Imp_SCM_EC_Friday_Order�i�ɶ��b EC��J�j... '
          Exec @Run_Err = DW.dbo.uSP_Imp_SCM_EC_Friday_Order
          if @Run_Err = @Errcode
          begin
             set @Run_Msg = '����'
             set @Cnt_Err = @Cnt_Err +1
          end
          else
          begin
            set @Run_Msg = '���\'
            set @Cnt_Suss = @Cnt_Suss +1
          end
       End Try
       Begin Catch 
          set @Run_Msg = '����'
          set @Cnt_Err = @Cnt_Err +1
       End Catch

       set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
       set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
       Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1 

       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       Begin Try
          set @Run_Proc_Msg = 'EC06.���� uSP_Imp_SCM_EC_GoHappy_Order�i�ּ��� EC��J�j... '
          Exec @Run_Err = DW.dbo.uSP_Imp_SCM_EC_GoHappy_Order
          if @Run_Err = @Errcode
          begin
             set @Run_Msg = '����'
             set @Cnt_Err = @Cnt_Err +1
          end
          else
          begin
            set @Run_Msg = '���\'
            set @Cnt_Suss = @Cnt_Suss +1
          end
       End Try
       Begin Catch 
          set @Run_Msg = '����'
          set @Cnt_Err = @Cnt_Err +1
       End Catch

       set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
       set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
       Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1 

       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       Begin Try
          set @Run_Proc_Msg = 'EC07.���� uSP_Imp_SCM_EC_GoMy_Order�iGOMY EC��J�j... '
          Exec @Run_Err = DW.dbo.uSP_Imp_SCM_EC_GoMy_Order
          if @Run_Err = @Errcode
          begin
             set @Run_Msg = '����'
             set @Cnt_Err = @Cnt_Err +1
          end
          else
          begin
            set @Run_Msg = '���\'
            set @Cnt_Suss = @Cnt_Suss +1
          end
       End Try
       Begin Catch 
          set @Run_Msg = '����'
          set @Cnt_Err = @Cnt_Err +1
       End Catch

       set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
       set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
       Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1 

       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       Begin Try
          set @Run_Proc_Msg = 'EC08.���� uSP_Imp_SCM_EC_MOMO_Order�i�I��MOMO EC��J�j... '
          Exec @Run_Err = DW.dbo.uSP_Imp_SCM_EC_MOMO_Order
          if @Run_Err = @Errcode
          begin
             set @Run_Msg = '����'
             set @Cnt_Err = @Cnt_Err +1
          end
          else
          begin
            set @Run_Msg = '���\'
            set @Cnt_Suss = @Cnt_Suss +1
          end
       End Try
       Begin Catch 
          set @Run_Msg = '����'
          set @Cnt_Err = @Cnt_Err +1
       End Catch

       set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
       set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
       Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1 

       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       Begin Try
          set @Run_Proc_Msg = 'EC09.���� uSP_Imp_SCM_EC_myfone_Order�i�x�W�j���j EC��J�j... '
          Exec @Run_Err = DW.dbo.uSP_Imp_SCM_EC_myfone_Order
          if @Run_Err = @Errcode
          begin
             set @Run_Msg = '����'
             set @Cnt_Err = @Cnt_Err +1
          end
          else
          begin
            set @Run_Msg = '���\'
            set @Cnt_Suss = @Cnt_Suss +1
          end
       End Try
       Begin Catch 
          set @Run_Msg = '����'
          set @Cnt_Err = @Cnt_Err +1
       End Catch

       set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
       set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
       Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1 

       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       Begin Try
          set @Run_Proc_Msg = 'EC10.���� uSP_Imp_SCM_EC_PayEasy_Order�iPayEasy EC��J�j... '
          Exec @Run_Err = DW.dbo.uSP_Imp_SCM_EC_PayEasy_Order
          if @Run_Err = @Errcode
          begin
             set @Run_Msg = '����'
             set @Cnt_Err = @Cnt_Err +1
          end
          else
          begin
            set @Run_Msg = '���\'
            set @Cnt_Suss = @Cnt_Suss +1
          end
       End Try
       Begin Catch 
          set @Run_Msg = '����'
          set @Cnt_Err = @Cnt_Err +1
       End Catch

       set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
       set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
       Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1 

       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       Begin Try
          set @Run_Proc_Msg = 'EC11.���� uSP_Imp_SCM_EC_PCShop_Order�iPCHome�ʪ� EC��J�j... '
          Exec @Run_Err = DW.dbo.uSP_Imp_SCM_EC_PCShop_Order
          if @Run_Err = @Errcode
          begin
             set @Run_Msg = '����'
             set @Cnt_Err = @Cnt_Err +1
          end
          else
          begin
            set @Run_Msg = '���\'
            set @Cnt_Suss = @Cnt_Suss +1
          end
       End Try
       Begin Catch 
          set @Run_Msg = '����'
          set @Cnt_Err = @Cnt_Err +1
       End Catch

       set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
       set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
       Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1 

       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       Begin Try
          set @Run_Proc_Msg = 'EC12.���� uSP_Imp_SCM_EC_PCStore_Order�iPCHome�ө��� EC��J�j... '
          Exec @Run_Err = DW.dbo.uSP_Imp_SCM_EC_PCStore_Order
          if @Run_Err = @Errcode
          begin
             set @Run_Msg = '����'
             set @Cnt_Err = @Cnt_Err +1
          end
          else
          begin
            set @Run_Msg = '���\'
            set @Cnt_Suss = @Cnt_Suss +1
          end
       End Try
       Begin Catch 
          set @Run_Msg = '����'
          set @Cnt_Err = @Cnt_Err +1
       End Catch

       set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
       set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
       Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1 

       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       Begin Try
          set @Run_Proc_Msg = 'EC13.���� uSP_Imp_SCM_EC_RELO_Order�i�Q�� EC��J�j... '
          Exec @Run_Err = DW.dbo.uSP_Imp_SCM_EC_RELO_Order
          if @Run_Err = @Errcode
          begin
             set @Run_Msg = '����'
             set @Cnt_Err = @Cnt_Err +1
          end
          else
          begin
            set @Run_Msg = '���\'
            set @Cnt_Suss = @Cnt_Suss +1
          end
       End Try
       Begin Catch 
          set @Run_Msg = '����'
          set @Cnt_Err = @Cnt_Err +1
       End Catch

       set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
       set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
       Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1 

       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       Begin Try
          set @Run_Proc_Msg = 'EC14.���� uSP_Imp_SCM_EC_UDN_Order�iUDN EC��J�j... '
          Exec @Run_Err = DW.dbo.uSP_Imp_SCM_EC_UDN_Order
          if @Run_Err = @Errcode
          begin
             set @Run_Msg = '����'
             set @Cnt_Err = @Cnt_Err +1
          end
          else
          begin
            set @Run_Msg = '���\'
            set @Cnt_Suss = @Cnt_Suss +1
          end
       End Try
       Begin Catch 
          set @Run_Msg = '����'
          set @Cnt_Err = @Cnt_Err +1
       End Catch

       set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
       set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
       Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1 

       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       Begin Try
          set @Run_Proc_Msg = 'EC15.���� uSP_Imp_SCM_EC_Umall_Order�i�˴��ʪ� EC��J�j... '
          Exec @Run_Err = DW.dbo.uSP_Imp_SCM_EC_Umall_Order
          if @Run_Err = @Errcode
          begin
             set @Run_Msg = '����'
             set @Cnt_Err = @Cnt_Err +1
          end
          else
          begin
            set @Run_Msg = '���\'
            set @Cnt_Suss = @Cnt_Suss +1
          end
       End Try
       Begin Catch 
          set @Run_Msg = '����'
          set @Cnt_Err = @Cnt_Err +1
       End Catch

       set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
       set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
       Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1 

       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       Begin Try
          set @Run_Proc_Msg = 'EC16.���� uSP_Imp_SCM_EC_YahooBuy_Order�iYahoo�ʪ� EC��J�j... '
          Exec @Run_Err = DW.dbo.uSP_Imp_SCM_EC_YahooBuy_Order
          if @Run_Err = @Errcode
          begin
             set @Run_Msg = '����'
             set @Cnt_Err = @Cnt_Err +1
          end
          else
          begin
            set @Run_Msg = '���\'
            set @Cnt_Suss = @Cnt_Suss +1
          end
       End Try
       Begin Catch 
          set @Run_Msg = '����'
          set @Cnt_Err = @Cnt_Err +1
       End Catch

       set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
       set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
       Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1   

       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       Begin Try
          set @Run_Proc_Msg = 'EC17.���� uSP_Imp_SCM_EC_YahooMall_Order�iYahoo�W�Űӫ� EC��J�j... '
          Exec @Run_Err = DW.dbo.uSP_Imp_SCM_EC_YahooMall_Order
          if @Run_Err = @Errcode
          begin
             set @Run_Msg = '����'
             set @Cnt_Err = @Cnt_Err +1
          end
          else
          begin
            set @Run_Msg = '���\'
            set @Cnt_Suss = @Cnt_Suss +1
          end
       End Try
       Begin Catch 
          set @Run_Msg = '����'
          set @Cnt_Err = @Cnt_Err +1
       End Catch

       set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
       set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
       Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1    
       
       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       Begin Try
          set @Run_Proc_Msg = 'EC18.���� uSP_Imp_SCM_EC_SharedFormat_Order�i�@�ή榡��J�j... '
          Exec @Run_Err = DW.dbo.uSP_Imp_SCM_EC_SharedFormat_Order
          if @Run_Err = @Errcode
          begin
             set @Run_Msg = '����'
             set @Cnt_Err = @Cnt_Err +1
          end
          else
          begin
            set @Run_Msg = '���\'
            set @Cnt_Suss = @Cnt_Suss +1
          end
       End Try
       Begin Catch 
          set @Run_Msg = '����'
          set @Cnt_Err = @Cnt_Err +1
       End Catch

       set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
       set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
       Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1    

       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
    --   Begin Try
    --      set @Run_Proc_Msg = '01.���� SP_Imp_SCM_Consignment_Car1�i���R��-�H�ܰU��^�f���ɡj... '
    --      Exec @Run_Err = DW.dbo.SP_Imp_SCM_Consignment_Car1
    --      if @Run_Err = @Errcode
    --      begin
    --         set @Run_Msg = '����'
    --         set @Cnt_Err = @Cnt_Err +1
    --      end
    --      else
    --      begin
    --        set @Run_Msg = '���\'
    --        set @Cnt_Suss = @Cnt_Suss +1
    --      end
    --   End Try
    --   Begin Catch 
    --      set @Run_Msg = '����'
    --      set @Cnt_Err = @Cnt_Err +1
    --   End Catch

    --   set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
    --   set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
    --   Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1
       
        
    --   ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
    --   Begin Try
    --     set @Run_Proc_Msg = '02.���� SP_Imp_SCM_Consignment_PCHome�iPCHome-�H�ܰU��^�f���ɡj... '
    --     Exec @Run_Err = DW.dbo.SP_Imp_SCM_Consignment_PCHome
    --     if @Run_Err = @Errcode
    --     begin
    --        set @Run_Msg = '����'
    --        set @Cnt_Err = @Cnt_Err +1
    --     end
    --     else
    --     begin
    --       set @Run_Msg = '���\'
    --       set @Cnt_Suss = @Cnt_Suss +1
    --     end
    --   End Try
    --   Begin Catch
    --        set @Run_Msg = '����'
    --        set @Cnt_Err = @Cnt_Err +1
    --   End Catch
       
    --   set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
    --   set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
    --   Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1
  
    --   ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
    --   Begin Try
    --     set @Run_Proc_Msg = '03.���� SP_Imp_SCM_Consignment_Yahoo�iYahoo-�U��^�f���ɡj... '
    --     Exec @Run_Err = DW.dbo.SP_Imp_SCM_Consignment_Yahoo
    --     if @Run_Err = @Errcode
    --     begin
    --        set @Run_Msg = '����'
    --        set @Cnt_Err = @Cnt_Err +1
    --     end
    --     else
    --     begin
    --       set @Run_Msg = '���\'
    --       set @Cnt_Suss = @Cnt_Suss +1
    --     end
    --   End Try
    --   Begin Catch
    --        set @Run_Msg = '����'
    --        set @Cnt_Err = @Cnt_Err +1
    --   End Catch
         
    --   set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
    --   set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
    --   Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1

    --   ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
    --   Begin Try
    --      set @Run_Proc_Msg = '04.���� SP_Imp_SCM_Comp_Stock_Car1�i���R��-�H�ܮw�s��b���ɡj... '
    --      Exec @Run_Err = DW.dbo.SP_Imp_SCM_Comp_Stock_Car1
    --      if @Run_Err = @Errcode
    --      begin
    --         set @Run_Msg = '����'
    --         set @Cnt_Err = @Cnt_Err +1
    --      end
    --      else
    --      begin
    --        set @Run_Msg = '���\'
    --        set @Cnt_Suss = @Cnt_Suss +1
    --      end
    --   End Try
    --   Begin Catch
    --        set @Run_Msg = '����'
    --        set @Cnt_Err = @Cnt_Err +1
    --   End Catch

    --   set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
    --   set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
    --   Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1
  
    --   ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
    --   Begin Try
    --      set @Run_Proc_Msg = '05.���� SP_Imp_SCM_Comp_Sale_Order_Yahoo�iYahoo-�H�ܮw�s��b���ɡj... '
    --      Exec @Run_Err = DW.dbo.SP_Imp_SCM_Comp_Sale_Order_Yahoo
    --      if @Run_Err = @Errcode
    --      begin
    --         set @Run_Msg = '����'
    --         set @Cnt_Err = @Cnt_Err +1
    --      end
    --      else
    --      begin
    --        set @Run_Msg = '���\'
    --        set @Cnt_Suss = @Cnt_Suss +1
    --      end
    --   End Try
    --   Begin Catch
    --        set @Run_Msg = '����'
    --        set @Cnt_Err = @Cnt_Err +1
    --   End Catch

    --   set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
    --   set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
    --   Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1

    --   ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
    --   Begin Try
    --      set @Run_Proc_Msg = '06.���� SP_Imp_SCM_Comp_Sale_Order_TwMobile�i�x�W�j���j��b���ɡj... '
    --      Exec @Run_Err = DW.dbo.SP_Imp_SCM_Comp_Sale_Order_TwMobile
    --      if @Run_Err = @Errcode
    --      begin
    --         set @Run_Msg = '����'
    --         set @Cnt_Err = @Cnt_Err +1
    --      end
    --      else
    --      begin
    --        set @Run_Msg = '���\'
    --        set @Cnt_Suss = @Cnt_Suss +1
    --      end
    --   End Try
    --   Begin Catch
    --        set @Run_Msg = '����'
    --        set @Cnt_Err = @Cnt_Err +1
    --   End Catch

    --   set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
    --   set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
    --   Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1
       
    --   ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
    --   Begin Try
    --      set @Run_Proc_Msg = '07.���� SP_Imp_SCM_Comp_Sale_Order_SuperMarket�i�W�Űӫ� ��b���ɡj... '
    --      Exec @Run_Err = DW.dbo.SP_Imp_SCM_Comp_Sale_Order_SuperMarket
    --      if @Run_Err = @Errcode
    --      begin
    --         set @Run_Msg = '����'
    --         set @Cnt_Err = @Cnt_Err +1
    --      end
    --      else
    --      begin
    --        set @Run_Msg = '���\'
    --        set @Cnt_Suss = @Cnt_Suss +1
    --      end
    --   End Try
    --   Begin Catch
    --        set @Run_Msg = '����'
    --        set @Cnt_Err = @Cnt_Err +1
    --   End Catch

    --   set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
    --   set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
    --   Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1
           
    --   ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
    --   Begin Try
    --      set @Run_Proc_Msg = '08.���� SP_Imp_SCM_Comp_Sale_Order_Momo�i�I�� MOMO ��b���ɡj... '
    --      Exec @Run_Err = DW.dbo.SP_Imp_SCM_Comp_Sale_Order_Momo
    --      if @Run_Err = @Errcode
    --      begin
    --         set @Run_Msg = '����'
    --         set @Cnt_Err = @Cnt_Err +1
    --      end
    --      else
    --      begin
    --        set @Run_Msg = '���\'
    --        set @Cnt_Suss = @Cnt_Suss +1
    --      end
    --   End Try
    --   Begin Catch
    --        set @Run_Msg = '����'
    --        set @Cnt_Err = @Cnt_Err +1
    --   End Catch

    --   set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
    --   set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
    --   Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1

    --   ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
    --   Begin Try
    --      set @Run_Proc_Msg = '09.���� SP_Imp_SCM_Consignment_Momo�i�I�� MOMO �H�����ɡj... '
    --      Exec @Run_Err = DW.dbo.SP_Imp_SCM_Consignment_Momo
    --      if @Run_Err = @Errcode
    --      begin
    --         set @Run_Msg = '����'
    --         set @Cnt_Err = @Cnt_Err +1
    --      end
    --      else
    --      begin
    --        set @Run_Msg = '���\'
    --        set @Cnt_Suss = @Cnt_Suss +1
    --      end
    --   End Try
    --   Begin Catch
    --        set @Run_Msg = '����'
    --        set @Cnt_Err = @Cnt_Err +1
    --   End Catch

    --   set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
    --   set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
    --   Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1
       
    --   ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
    --   Begin Try
    --      set @Run_Proc_Msg = '13.���� SP_Imp_Sales_Invoice_To_Receipt_Voucher�i�P���o������b�ǲ��j... '
    --      Exec @Run_Err = DW.dbo.SP_Imp_Sales_Invoice_To_Receipt_Voucher
    --      if @Run_Err = @Errcode
    --      begin
    --         set @Run_Msg = '����'
    --         set @Cnt_Err = @Cnt_Err +1
    --      end
    --      else
    --      begin
    --        set @Run_Msg = '���\'
    --        set @Cnt_Suss = @Cnt_Suss +1
    --      end
    --   End Try
    --   Begin Catch
    --        set @Run_Msg = '����'
    --        set @Cnt_Err = @Cnt_Err +1
    --   End Catch

    --   set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
    --   set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
    --   Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1
    --   ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
	   --Begin Try
    --      set @Run_Proc_Msg = '14.���� SP_Imp_SCM_Consignment_Car1_Order�i���R��-�z�f��q��j... '
    --      Exec @Run_Err = DW.dbo.SP_Imp_SCM_Consignment_Car1_Order
    --      if @Run_Err = @Errcode
    --      begin
    --         set @Run_Msg = '����'
    --         set @Cnt_Err = @Cnt_Err +1
    --      end
    --      else
    --      begin
    --        set @Run_Msg = '���\'
    --        set @Cnt_Suss = @Cnt_Suss +1
    --      end
    --   End Try
    --   Begin Catch
    --        set @Run_Msg = '����'
    --        set @Cnt_Err = @Cnt_Err +1
    --   End Catch


    --   set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
    --   set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
    --   Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1
    --   ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
	   --Begin Try
    --      set @Run_Proc_Msg = '15.���� SP_Imp_SCM_Consignment_Secar_Order�i���q-�z�f��q��j... '
    --      Exec @Run_Err = DW.dbo.SP_Imp_SCM_Consignment_Secar_Order
    --      if @Run_Err = @Errcode
    --      begin
    --         set @Run_Msg = '����'
    --         set @Cnt_Err = @Cnt_Err +1
    --      end
    --      else
    --      begin
    --        set @Run_Msg = '���\'
    --        set @Cnt_Suss = @Cnt_Suss +1
    --      end
    --   End Try
    --   Begin Catch
    --        set @Run_Msg = '����'
    --        set @Cnt_Err = @Cnt_Err +1
    --   End Catch

	   
    --   set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
    --   set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
    --   Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1
    end
    else
       set @Body_Msg = '�L Excel ��ƱN���i�����ɡC'
       
---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---

    if @Cnt_txt <> @Errcode
    begin
       set @Body_Msg = @Body_Msg + REPLICATE('#', 100)+@CR
       set @Body_Msg = @Body_Msg + '���� uSP_Sys_Imp_txt_to_db�i�@�פJ Text �ɮ׼ơj...('+CONVERT(Varchar(100), @Cnt_txt)+'��)�C'+@CR
       set @Body_Msg = @Body_Msg + REPLICATE('#', 100)+@CR
  
       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       Begin Try
          set @Run_Proc_Msg = '01.���� uSP_Imp_SCM_Consignment_800yaoya�i�K�ʫ�-�z�f��q��j... '
          Exec @Run_Err = DW.dbo.uSP_Imp_SCM_Consignment_800yaoya
          if @Run_Err = @Errcode
          begin
             set @Run_Msg = '����'
             set @Cnt_Err = @Cnt_Err +1
          end
          else
          begin
            set @Run_Msg = '���\'
            set @Cnt_Suss = @Cnt_Suss +1
          end
       End Try
       Begin Catch
            set @Run_Msg = '����'
            set @Cnt_Err = @Cnt_Err +1
       End Catch

       set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
       set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
       Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1
        
       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       Begin Try
          set @Run_Proc_Msg = '02.���� uSP_Imp_SCM_Consignment_Carno1�i������-�z�f��q��j... '
          Exec @Run_Err = DW.dbo.uSP_Imp_SCM_Consignment_Carno1
          if @Run_Err = @Errcode
          begin
             set @Run_Msg = '����'
             set @Cnt_Err = @Cnt_Err +1
          end
          else
          begin
            set @Run_Msg = '���\'
            set @Cnt_Suss = @Cnt_Suss +1
          end
       End Try
       Begin Catch
            set @Run_Msg = '����'
            set @Cnt_Err = @Cnt_Err +1
       End Catch

       set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
       set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
       Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1
    end
    else
       set @Body_Msg = '�L Text ��ƱN���i�����ɡC'

    Set @Run_Msg = '���� '+@Proce_SubName+'...���\ ('+CONVERT(Varchar(2), @Cnt_Suss)+'), ���� ('+CONVERT(Varchar(2), @Cnt_Err)+')'
  end try
  begin catch
    Set @Run_Msg = '���� '+@Run_Proc_Msg+'...�o���Y�����~!!'
    Set @Body_Msg = @Run_Proc_Msg+'���~�T���G'+@CR+ERROR_MESSAGE()
    Exec uSP_Sys_Write_Log @Proc, @Run_Msg, @Body_Msg, @ErrCode, 1
  end catch

  End_Exit:
  Set @ETime = GetDate()
  Set @Run_Msg =@Proc+' '+@Proce_SubName+'...����ɶ��G[ '+CONVERT(Varchar(100), @ETime - @BTime, 114)+' ].'+@CR
  Set @Body_Msg = @Body_Msg + @Run_Msg
  Exec uSP_Sys_Write_Log @Proc, @Run_Msg, @Body_Msg, 0, 1

End
GO
