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
  Declare @Proce_SubName NVarchar(Max) = '外部檔案匯入至 DW 資料庫程序'
  Declare @BTime DateTime, @ETime DateTime
  Declare @Hour Varchar(2) = ''
  
  set @BTime = GetDate()
  
  set @Hour =(select left(convert(varchar(20), @BTime, 114), 2))
  
  if (@Hour in ('12', '00')) --< 設定禁止運行時間 (24 小時制)
  Begin
     Set @body_Msg = @body_Msg + '['+@Proc+'] 已被設定 '+@Hour+' 時禁止運行!!'+@CR
     Goto End_Exit
  end
  
  Begin Try
    Exec @Cnt_xls = DW.dbo.uSP_Sys_Imp_xls_to_db
    Exec @Cnt_txt = DW.dbo.uSP_Sys_Imp_txt_to_db

    set @Cnt_Err = @Cnt_xls + @Cnt_txt
    if (@Cnt_Err in (0, -2))
    begin
       Set @body_Msg = @body_Msg + '無任何資料不進行轉檔!!'+@CR
       Goto End_Exit
    end
    
    if @Cnt_xls <> @Errcode
    begin
       set @Body_Msg = @Body_Msg + REPLICATE('#', 100)+@CR
       set @Body_Msg = @Body_Msg + '執行 uSP_Imp_xls_to_db【共匯入 Excel 檔案數】...('+CONVERT(Varchar(100), @Cnt_xls)+'個)。'+@CR
       set @Body_Msg = @Body_Msg + REPLICATE('#', 100)+@CR
 
       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       Begin Try
          set @Run_Proc_Msg = 'EC01.執行 uSP_Imp_SCM_EC_Books_Order【博客來 EC轉入】... '
          Exec @Run_Err = DW.dbo.uSP_Imp_SCM_EC_Books_Order
          if @Run_Err = @Errcode
          begin
             set @Run_Msg = '失敗'
             set @Cnt_Err = @Cnt_Err +1
          end
          else
          begin
            set @Run_Msg = '成功'
            set @Cnt_Suss = @Cnt_Suss +1
          end
       End Try
       Begin Catch 
          set @Run_Msg = '失敗'
          set @Cnt_Err = @Cnt_Err +1
       End Catch

       set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
       set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
       Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1


       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       Begin Try
          set @Run_Proc_Msg = 'EC02.執行 uSP_Imp_SCM_EC_CrazyMike_Order【瘋狂賣客 EC轉入】... '
          Exec @Run_Err = DW.dbo.uSP_Imp_SCM_EC_CrazyMike_Order
          if @Run_Err = @Errcode
          begin
             set @Run_Msg = '失敗'
             set @Cnt_Err = @Cnt_Err +1
          end
          else
          begin
            set @Run_Msg = '成功'
            set @Cnt_Suss = @Cnt_Suss +1
          end
       End Try
       Begin Catch 
          set @Run_Msg = '失敗'
          set @Cnt_Err = @Cnt_Err +1
       End Catch

       set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
       set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
       Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1     

       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       Begin Try
          set @Run_Proc_Msg = 'EC03.執行 uSP_Imp_SCM_EC_ETmall_Order【東森購物 EC轉入】... '
          Exec @Run_Err = DW.dbo.uSP_Imp_SCM_EC_ETmall_Order
          if @Run_Err = @Errcode
          begin
             set @Run_Msg = '失敗'
             set @Cnt_Err = @Cnt_Err +1
          end
          else
          begin
            set @Run_Msg = '成功'
            set @Cnt_Suss = @Cnt_Suss +1
          end
       End Try
       Begin Catch 
          set @Run_Msg = '失敗'
          set @Cnt_Err = @Cnt_Err +1
       End Catch

       set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
       set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
       Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1 

       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       Begin Try
          set @Run_Proc_Msg = 'EC04.執行 uSP_Imp_SCM_EC_ForMall_Order【瘋MALL EC轉入】... '
          Exec @Run_Err = DW.dbo.uSP_Imp_SCM_EC_ForMall_Order
          if @Run_Err = @Errcode
          begin
             set @Run_Msg = '失敗'
             set @Cnt_Err = @Cnt_Err +1
          end
          else
          begin
            set @Run_Msg = '成功'
            set @Cnt_Suss = @Cnt_Suss +1
          end
       End Try
       Begin Catch 
          set @Run_Msg = '失敗'
          set @Cnt_Err = @Cnt_Err +1
       End Catch

       set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
       set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
       Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1 

       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       Begin Try
          set @Run_Proc_Msg = 'EC05.執行 uSP_Imp_SCM_EC_Friday_Order【時間軸 EC轉入】... '
          Exec @Run_Err = DW.dbo.uSP_Imp_SCM_EC_Friday_Order
          if @Run_Err = @Errcode
          begin
             set @Run_Msg = '失敗'
             set @Cnt_Err = @Cnt_Err +1
          end
          else
          begin
            set @Run_Msg = '成功'
            set @Cnt_Suss = @Cnt_Suss +1
          end
       End Try
       Begin Catch 
          set @Run_Msg = '失敗'
          set @Cnt_Err = @Cnt_Err +1
       End Catch

       set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
       set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
       Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1 

       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       Begin Try
          set @Run_Proc_Msg = 'EC06.執行 uSP_Imp_SCM_EC_GoHappy_Order【快樂購 EC轉入】... '
          Exec @Run_Err = DW.dbo.uSP_Imp_SCM_EC_GoHappy_Order
          if @Run_Err = @Errcode
          begin
             set @Run_Msg = '失敗'
             set @Cnt_Err = @Cnt_Err +1
          end
          else
          begin
            set @Run_Msg = '成功'
            set @Cnt_Suss = @Cnt_Suss +1
          end
       End Try
       Begin Catch 
          set @Run_Msg = '失敗'
          set @Cnt_Err = @Cnt_Err +1
       End Catch

       set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
       set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
       Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1 

       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       Begin Try
          set @Run_Proc_Msg = 'EC07.執行 uSP_Imp_SCM_EC_GoMy_Order【GOMY EC轉入】... '
          Exec @Run_Err = DW.dbo.uSP_Imp_SCM_EC_GoMy_Order
          if @Run_Err = @Errcode
          begin
             set @Run_Msg = '失敗'
             set @Cnt_Err = @Cnt_Err +1
          end
          else
          begin
            set @Run_Msg = '成功'
            set @Cnt_Suss = @Cnt_Suss +1
          end
       End Try
       Begin Catch 
          set @Run_Msg = '失敗'
          set @Cnt_Err = @Cnt_Err +1
       End Catch

       set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
       set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
       Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1 

       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       Begin Try
          set @Run_Proc_Msg = 'EC08.執行 uSP_Imp_SCM_EC_MOMO_Order【富邦MOMO EC轉入】... '
          Exec @Run_Err = DW.dbo.uSP_Imp_SCM_EC_MOMO_Order
          if @Run_Err = @Errcode
          begin
             set @Run_Msg = '失敗'
             set @Cnt_Err = @Cnt_Err +1
          end
          else
          begin
            set @Run_Msg = '成功'
            set @Cnt_Suss = @Cnt_Suss +1
          end
       End Try
       Begin Catch 
          set @Run_Msg = '失敗'
          set @Cnt_Err = @Cnt_Err +1
       End Catch

       set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
       set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
       Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1 

       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       Begin Try
          set @Run_Proc_Msg = 'EC09.執行 uSP_Imp_SCM_EC_myfone_Order【台灣大哥大 EC轉入】... '
          Exec @Run_Err = DW.dbo.uSP_Imp_SCM_EC_myfone_Order
          if @Run_Err = @Errcode
          begin
             set @Run_Msg = '失敗'
             set @Cnt_Err = @Cnt_Err +1
          end
          else
          begin
            set @Run_Msg = '成功'
            set @Cnt_Suss = @Cnt_Suss +1
          end
       End Try
       Begin Catch 
          set @Run_Msg = '失敗'
          set @Cnt_Err = @Cnt_Err +1
       End Catch

       set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
       set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
       Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1 

       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       Begin Try
          set @Run_Proc_Msg = 'EC10.執行 uSP_Imp_SCM_EC_PayEasy_Order【PayEasy EC轉入】... '
          Exec @Run_Err = DW.dbo.uSP_Imp_SCM_EC_PayEasy_Order
          if @Run_Err = @Errcode
          begin
             set @Run_Msg = '失敗'
             set @Cnt_Err = @Cnt_Err +1
          end
          else
          begin
            set @Run_Msg = '成功'
            set @Cnt_Suss = @Cnt_Suss +1
          end
       End Try
       Begin Catch 
          set @Run_Msg = '失敗'
          set @Cnt_Err = @Cnt_Err +1
       End Catch

       set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
       set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
       Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1 

       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       Begin Try
          set @Run_Proc_Msg = 'EC11.執行 uSP_Imp_SCM_EC_PCShop_Order【PCHome購物 EC轉入】... '
          Exec @Run_Err = DW.dbo.uSP_Imp_SCM_EC_PCShop_Order
          if @Run_Err = @Errcode
          begin
             set @Run_Msg = '失敗'
             set @Cnt_Err = @Cnt_Err +1
          end
          else
          begin
            set @Run_Msg = '成功'
            set @Cnt_Suss = @Cnt_Suss +1
          end
       End Try
       Begin Catch 
          set @Run_Msg = '失敗'
          set @Cnt_Err = @Cnt_Err +1
       End Catch

       set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
       set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
       Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1 

       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       Begin Try
          set @Run_Proc_Msg = 'EC12.執行 uSP_Imp_SCM_EC_PCStore_Order【PCHome商店街 EC轉入】... '
          Exec @Run_Err = DW.dbo.uSP_Imp_SCM_EC_PCStore_Order
          if @Run_Err = @Errcode
          begin
             set @Run_Msg = '失敗'
             set @Cnt_Err = @Cnt_Err +1
          end
          else
          begin
            set @Run_Msg = '成功'
            set @Cnt_Suss = @Cnt_Suss +1
          end
       End Try
       Begin Catch 
          set @Run_Msg = '失敗'
          set @Cnt_Err = @Cnt_Err +1
       End Catch

       set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
       set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
       Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1 

       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       Begin Try
          set @Run_Proc_Msg = 'EC13.執行 uSP_Imp_SCM_EC_RELO_Order【利樂 EC轉入】... '
          Exec @Run_Err = DW.dbo.uSP_Imp_SCM_EC_RELO_Order
          if @Run_Err = @Errcode
          begin
             set @Run_Msg = '失敗'
             set @Cnt_Err = @Cnt_Err +1
          end
          else
          begin
            set @Run_Msg = '成功'
            set @Cnt_Suss = @Cnt_Suss +1
          end
       End Try
       Begin Catch 
          set @Run_Msg = '失敗'
          set @Cnt_Err = @Cnt_Err +1
       End Catch

       set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
       set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
       Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1 

       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       Begin Try
          set @Run_Proc_Msg = 'EC14.執行 uSP_Imp_SCM_EC_UDN_Order【UDN EC轉入】... '
          Exec @Run_Err = DW.dbo.uSP_Imp_SCM_EC_UDN_Order
          if @Run_Err = @Errcode
          begin
             set @Run_Msg = '失敗'
             set @Cnt_Err = @Cnt_Err +1
          end
          else
          begin
            set @Run_Msg = '成功'
            set @Cnt_Suss = @Cnt_Suss +1
          end
       End Try
       Begin Catch 
          set @Run_Msg = '失敗'
          set @Cnt_Err = @Cnt_Err +1
       End Catch

       set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
       set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
       Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1 

       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       Begin Try
          set @Run_Proc_Msg = 'EC15.執行 uSP_Imp_SCM_EC_Umall_Order【森森購物 EC轉入】... '
          Exec @Run_Err = DW.dbo.uSP_Imp_SCM_EC_Umall_Order
          if @Run_Err = @Errcode
          begin
             set @Run_Msg = '失敗'
             set @Cnt_Err = @Cnt_Err +1
          end
          else
          begin
            set @Run_Msg = '成功'
            set @Cnt_Suss = @Cnt_Suss +1
          end
       End Try
       Begin Catch 
          set @Run_Msg = '失敗'
          set @Cnt_Err = @Cnt_Err +1
       End Catch

       set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
       set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
       Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1 

       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       Begin Try
          set @Run_Proc_Msg = 'EC16.執行 uSP_Imp_SCM_EC_YahooBuy_Order【Yahoo購物 EC轉入】... '
          Exec @Run_Err = DW.dbo.uSP_Imp_SCM_EC_YahooBuy_Order
          if @Run_Err = @Errcode
          begin
             set @Run_Msg = '失敗'
             set @Cnt_Err = @Cnt_Err +1
          end
          else
          begin
            set @Run_Msg = '成功'
            set @Cnt_Suss = @Cnt_Suss +1
          end
       End Try
       Begin Catch 
          set @Run_Msg = '失敗'
          set @Cnt_Err = @Cnt_Err +1
       End Catch

       set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
       set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
       Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1   

       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       Begin Try
          set @Run_Proc_Msg = 'EC17.執行 uSP_Imp_SCM_EC_YahooMall_Order【Yahoo超級商城 EC轉入】... '
          Exec @Run_Err = DW.dbo.uSP_Imp_SCM_EC_YahooMall_Order
          if @Run_Err = @Errcode
          begin
             set @Run_Msg = '失敗'
             set @Cnt_Err = @Cnt_Err +1
          end
          else
          begin
            set @Run_Msg = '成功'
            set @Cnt_Suss = @Cnt_Suss +1
          end
       End Try
       Begin Catch 
          set @Run_Msg = '失敗'
          set @Cnt_Err = @Cnt_Err +1
       End Catch

       set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
       set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
       Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1    
       
       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       Begin Try
          set @Run_Proc_Msg = 'EC18.執行 uSP_Imp_SCM_EC_SharedFormat_Order【共用格式轉入】... '
          Exec @Run_Err = DW.dbo.uSP_Imp_SCM_EC_SharedFormat_Order
          if @Run_Err = @Errcode
          begin
             set @Run_Msg = '失敗'
             set @Cnt_Err = @Cnt_Err +1
          end
          else
          begin
            set @Run_Msg = '成功'
            set @Cnt_Suss = @Cnt_Suss +1
          end
       End Try
       Begin Catch 
          set @Run_Msg = '失敗'
          set @Cnt_Err = @Cnt_Err +1
       End Catch

       set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
       set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
       Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1    

       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
    --   Begin Try
    --      set @Run_Proc_Msg = '01.執行 SP_Imp_SCM_Consignment_Car1【車麗屋-寄倉託售回貨轉檔】... '
    --      Exec @Run_Err = DW.dbo.SP_Imp_SCM_Consignment_Car1
    --      if @Run_Err = @Errcode
    --      begin
    --         set @Run_Msg = '失敗'
    --         set @Cnt_Err = @Cnt_Err +1
    --      end
    --      else
    --      begin
    --        set @Run_Msg = '成功'
    --        set @Cnt_Suss = @Cnt_Suss +1
    --      end
    --   End Try
    --   Begin Catch 
    --      set @Run_Msg = '失敗'
    --      set @Cnt_Err = @Cnt_Err +1
    --   End Catch

    --   set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
    --   set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
    --   Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1
       
        
    --   ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
    --   Begin Try
    --     set @Run_Proc_Msg = '02.執行 SP_Imp_SCM_Consignment_PCHome【PCHome-寄倉託售回貨轉檔】... '
    --     Exec @Run_Err = DW.dbo.SP_Imp_SCM_Consignment_PCHome
    --     if @Run_Err = @Errcode
    --     begin
    --        set @Run_Msg = '失敗'
    --        set @Cnt_Err = @Cnt_Err +1
    --     end
    --     else
    --     begin
    --       set @Run_Msg = '成功'
    --       set @Cnt_Suss = @Cnt_Suss +1
    --     end
    --   End Try
    --   Begin Catch
    --        set @Run_Msg = '失敗'
    --        set @Cnt_Err = @Cnt_Err +1
    --   End Catch
       
    --   set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
    --   set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
    --   Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1
  
    --   ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
    --   Begin Try
    --     set @Run_Proc_Msg = '03.執行 SP_Imp_SCM_Consignment_Yahoo【Yahoo-託售回貨轉檔】... '
    --     Exec @Run_Err = DW.dbo.SP_Imp_SCM_Consignment_Yahoo
    --     if @Run_Err = @Errcode
    --     begin
    --        set @Run_Msg = '失敗'
    --        set @Cnt_Err = @Cnt_Err +1
    --     end
    --     else
    --     begin
    --       set @Run_Msg = '成功'
    --       set @Cnt_Suss = @Cnt_Suss +1
    --     end
    --   End Try
    --   Begin Catch
    --        set @Run_Msg = '失敗'
    --        set @Cnt_Err = @Cnt_Err +1
    --   End Catch
         
    --   set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
    --   set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
    --   Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1

    --   ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
    --   Begin Try
    --      set @Run_Proc_Msg = '04.執行 SP_Imp_SCM_Comp_Stock_Car1【車麗屋-寄倉庫存對帳轉檔】... '
    --      Exec @Run_Err = DW.dbo.SP_Imp_SCM_Comp_Stock_Car1
    --      if @Run_Err = @Errcode
    --      begin
    --         set @Run_Msg = '失敗'
    --         set @Cnt_Err = @Cnt_Err +1
    --      end
    --      else
    --      begin
    --        set @Run_Msg = '成功'
    --        set @Cnt_Suss = @Cnt_Suss +1
    --      end
    --   End Try
    --   Begin Catch
    --        set @Run_Msg = '失敗'
    --        set @Cnt_Err = @Cnt_Err +1
    --   End Catch

    --   set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
    --   set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
    --   Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1
  
    --   ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
    --   Begin Try
    --      set @Run_Proc_Msg = '05.執行 SP_Imp_SCM_Comp_Sale_Order_Yahoo【Yahoo-寄倉庫存對帳轉檔】... '
    --      Exec @Run_Err = DW.dbo.SP_Imp_SCM_Comp_Sale_Order_Yahoo
    --      if @Run_Err = @Errcode
    --      begin
    --         set @Run_Msg = '失敗'
    --         set @Cnt_Err = @Cnt_Err +1
    --      end
    --      else
    --      begin
    --        set @Run_Msg = '成功'
    --        set @Cnt_Suss = @Cnt_Suss +1
    --      end
    --   End Try
    --   Begin Catch
    --        set @Run_Msg = '失敗'
    --        set @Cnt_Err = @Cnt_Err +1
    --   End Catch

    --   set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
    --   set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
    --   Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1

    --   ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
    --   Begin Try
    --      set @Run_Proc_Msg = '06.執行 SP_Imp_SCM_Comp_Sale_Order_TwMobile【台灣大哥大對帳轉檔】... '
    --      Exec @Run_Err = DW.dbo.SP_Imp_SCM_Comp_Sale_Order_TwMobile
    --      if @Run_Err = @Errcode
    --      begin
    --         set @Run_Msg = '失敗'
    --         set @Cnt_Err = @Cnt_Err +1
    --      end
    --      else
    --      begin
    --        set @Run_Msg = '成功'
    --        set @Cnt_Suss = @Cnt_Suss +1
    --      end
    --   End Try
    --   Begin Catch
    --        set @Run_Msg = '失敗'
    --        set @Cnt_Err = @Cnt_Err +1
    --   End Catch

    --   set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
    --   set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
    --   Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1
       
    --   ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
    --   Begin Try
    --      set @Run_Proc_Msg = '07.執行 SP_Imp_SCM_Comp_Sale_Order_SuperMarket【超級商城 對帳轉檔】... '
    --      Exec @Run_Err = DW.dbo.SP_Imp_SCM_Comp_Sale_Order_SuperMarket
    --      if @Run_Err = @Errcode
    --      begin
    --         set @Run_Msg = '失敗'
    --         set @Cnt_Err = @Cnt_Err +1
    --      end
    --      else
    --      begin
    --        set @Run_Msg = '成功'
    --        set @Cnt_Suss = @Cnt_Suss +1
    --      end
    --   End Try
    --   Begin Catch
    --        set @Run_Msg = '失敗'
    --        set @Cnt_Err = @Cnt_Err +1
    --   End Catch

    --   set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
    --   set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
    --   Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1
           
    --   ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
    --   Begin Try
    --      set @Run_Proc_Msg = '08.執行 SP_Imp_SCM_Comp_Sale_Order_Momo【富邦 MOMO 對帳轉檔】... '
    --      Exec @Run_Err = DW.dbo.SP_Imp_SCM_Comp_Sale_Order_Momo
    --      if @Run_Err = @Errcode
    --      begin
    --         set @Run_Msg = '失敗'
    --         set @Cnt_Err = @Cnt_Err +1
    --      end
    --      else
    --      begin
    --        set @Run_Msg = '成功'
    --        set @Cnt_Suss = @Cnt_Suss +1
    --      end
    --   End Try
    --   Begin Catch
    --        set @Run_Msg = '失敗'
    --        set @Cnt_Err = @Cnt_Err +1
    --   End Catch

    --   set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
    --   set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
    --   Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1

    --   ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
    --   Begin Try
    --      set @Run_Proc_Msg = '09.執行 SP_Imp_SCM_Consignment_Momo【富邦 MOMO 寄倉轉檔】... '
    --      Exec @Run_Err = DW.dbo.SP_Imp_SCM_Consignment_Momo
    --      if @Run_Err = @Errcode
    --      begin
    --         set @Run_Msg = '失敗'
    --         set @Cnt_Err = @Cnt_Err +1
    --      end
    --      else
    --      begin
    --        set @Run_Msg = '成功'
    --        set @Cnt_Suss = @Cnt_Suss +1
    --      end
    --   End Try
    --   Begin Catch
    --        set @Run_Msg = '失敗'
    --        set @Cnt_Err = @Cnt_Err +1
    --   End Catch

    --   set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
    --   set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
    --   Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1
       
    --   ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
    --   Begin Try
    --      set @Run_Proc_Msg = '13.執行 SP_Imp_Sales_Invoice_To_Receipt_Voucher【銷項發票轉轉帳傳票】... '
    --      Exec @Run_Err = DW.dbo.SP_Imp_Sales_Invoice_To_Receipt_Voucher
    --      if @Run_Err = @Errcode
    --      begin
    --         set @Run_Msg = '失敗'
    --         set @Cnt_Err = @Cnt_Err +1
    --      end
    --      else
    --      begin
    --        set @Run_Msg = '成功'
    --        set @Cnt_Suss = @Cnt_Suss +1
    --      end
    --   End Try
    --   Begin Catch
    --        set @Run_Msg = '失敗'
    --        set @Cnt_Err = @Cnt_Err +1
    --   End Catch

    --   set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
    --   set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
    --   Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1
    --   ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
	   --Begin Try
    --      set @Run_Proc_Msg = '14.執行 SP_Imp_SCM_Consignment_Car1_Order【車麗屋-揀貨轉訂單】... '
    --      Exec @Run_Err = DW.dbo.SP_Imp_SCM_Consignment_Car1_Order
    --      if @Run_Err = @Errcode
    --      begin
    --         set @Run_Msg = '失敗'
    --         set @Cnt_Err = @Cnt_Err +1
    --      end
    --      else
    --      begin
    --        set @Run_Msg = '成功'
    --        set @Cnt_Suss = @Cnt_Suss +1
    --      end
    --   End Try
    --   Begin Catch
    --        set @Run_Msg = '失敗'
    --        set @Cnt_Err = @Cnt_Err +1
    --   End Catch


    --   set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
    --   set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
    --   Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1
    --   ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
	   --Begin Try
    --      set @Run_Proc_Msg = '15.執行 SP_Imp_SCM_Consignment_Secar_Order【旭益-揀貨轉訂單】... '
    --      Exec @Run_Err = DW.dbo.SP_Imp_SCM_Consignment_Secar_Order
    --      if @Run_Err = @Errcode
    --      begin
    --         set @Run_Msg = '失敗'
    --         set @Cnt_Err = @Cnt_Err +1
    --      end
    --      else
    --      begin
    --        set @Run_Msg = '成功'
    --        set @Cnt_Suss = @Cnt_Suss +1
    --      end
    --   End Try
    --   Begin Catch
    --        set @Run_Msg = '失敗'
    --        set @Cnt_Err = @Cnt_Err +1
    --   End Catch

	   
    --   set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
    --   set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
    --   Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1
    end
    else
       set @Body_Msg = '無 Excel 資料將不進行轉檔。'
       
---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---

    if @Cnt_txt <> @Errcode
    begin
       set @Body_Msg = @Body_Msg + REPLICATE('#', 100)+@CR
       set @Body_Msg = @Body_Msg + '執行 uSP_Sys_Imp_txt_to_db【共匯入 Text 檔案數】...('+CONVERT(Varchar(100), @Cnt_txt)+'個)。'+@CR
       set @Body_Msg = @Body_Msg + REPLICATE('#', 100)+@CR
  
       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       Begin Try
          set @Run_Proc_Msg = '01.執行 uSP_Imp_SCM_Consignment_800yaoya【八百屋-揀貨轉訂單】... '
          Exec @Run_Err = DW.dbo.uSP_Imp_SCM_Consignment_800yaoya
          if @Run_Err = @Errcode
          begin
             set @Run_Msg = '失敗'
             set @Cnt_Err = @Cnt_Err +1
          end
          else
          begin
            set @Run_Msg = '成功'
            set @Cnt_Suss = @Cnt_Suss +1
          end
       End Try
       Begin Catch
            set @Run_Msg = '失敗'
            set @Cnt_Err = @Cnt_Err +1
       End Catch

       set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
       set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
       Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1
        
       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       Begin Try
          set @Run_Proc_Msg = '02.執行 uSP_Imp_SCM_Consignment_Carno1【車之輪-揀貨轉訂單】... '
          Exec @Run_Err = DW.dbo.uSP_Imp_SCM_Consignment_Carno1
          if @Run_Err = @Errcode
          begin
             set @Run_Msg = '失敗'
             set @Cnt_Err = @Cnt_Err +1
          end
          else
          begin
            set @Run_Msg = '成功'
            set @Cnt_Suss = @Cnt_Suss +1
          end
       End Try
       Begin Catch
            set @Run_Msg = '失敗'
            set @Cnt_Err = @Cnt_Err +1
       End Catch

       set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
       set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
       Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1
    end
    else
       set @Body_Msg = '無 Text 資料將不進行轉檔。'

    Set @Run_Msg = '執行 '+@Proce_SubName+'...成功 ('+CONVERT(Varchar(2), @Cnt_Suss)+'), 失敗 ('+CONVERT(Varchar(2), @Cnt_Err)+')'
  end try
  begin catch
    Set @Run_Msg = '執行 '+@Run_Proc_Msg+'...發生嚴重錯誤!!'
    Set @Body_Msg = @Run_Proc_Msg+'錯誤訊息：'+@CR+ERROR_MESSAGE()
    Exec uSP_Sys_Write_Log @Proc, @Run_Msg, @Body_Msg, @ErrCode, 1
  end catch

  End_Exit:
  Set @ETime = GetDate()
  Set @Run_Msg =@Proc+' '+@Proce_SubName+'...執行時間：[ '+CONVERT(Varchar(100), @ETime - @BTime, 114)+' ].'+@CR
  Set @Body_Msg = @Body_Msg + @Run_Msg
  Exec uSP_Sys_Write_Log @Proc, @Run_Msg, @Body_Msg, 0, 1

End
GO
