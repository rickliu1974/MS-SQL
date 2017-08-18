USE [DW]
GO
/****** Object:  StoredProcedure [dbo].[uSP_ETL_New_Stock_Lists]    Script Date: 07/24/2017 14:43:59 ******/
DROP PROCEDURE [dbo].[uSP_ETL_New_Stock_Lists]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE Proc [dbo].[uSP_ETL_New_Stock_Lists](@Run_ETL Int = 1)
as
begin
  /*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
   Program Name: uSP_ETL_New_Stock_Lists
   Create Date: 2014/06/05
   Creator: Rickliu
  *-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*/  

  Declare @Proc Varchar(50) = 'uSP_ETL_New_Stock_Lists'
  Declare @Cnt Int =0
  Declare @Msg Varchar(Max) =''
  Declare @strSQL Varchar(Max)=''
  Declare @CR Varchar(4) = char(13)+char(10)
  Declare @Result Int = 0
  
  set @Msg = '更新商品的新品欄位'
  begin try
     Exec uSP_Sys_Waiting_Table_Lock 'Ori_XLS#New_Stock_Lists'

  -- 2014/07/02 -- Rickliu 採購美如要求將他所提供的新品列表之到貨年月，更新至凌越的商品基本資料內。
     Set @Msg = '以 [Ori_XLS#New_Stock_Lists] 更新 TYEIPDBS2.lytdbta13.dbo.sstock 商品資料內的新品欄位'
     Update TYEIPDBS2.lytdbta13.dbo.sstock 
        set sk_size = 
              Case 
                when isnull(m.arrival_date, '') ='' then '無'
                else Convert(Varchar(6), Convert(DateTime, m.arrival_date), 112) 
              End 
       from (select sk_no as n_skno, arrival_date 
               from Ori_XLS#New_Stock_Lists With(NoLock) ) m 
      where sk_no Collate Chinese_Taiwan_Stroke_CI_AS = m.n_skno Collate Chinese_Taiwan_Stroke_CI_AS
        and m.n_skno > ''
      
     if @Run_ETL = 1
     begin
        Exec @Result = DW.dbo.uSp_ETL_sstock
        if @Result <> -1 set @Result = 0
        Exec DW.dbo.uSP_ETL_sslpdt
        if @Result <> -1 set @Result = 0
     end

     Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt
  end try
  begin catch
    set @Result = -1
    set @Msg = @Msg+'(錯誤訊息:'+ERROR_MESSAGE()+')'

    Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Result
  end catch

  Return(@Result)
end
GO
