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

     -- 請付款方式：固定天數
     if @Kind = 2
     begin
        /*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-
         請付款方式：1=計算天數， 交易後 + n  天請付款 
               例如：單據日期為 2/3，計算天數為 19，則請付款天數則為 3 + 19 = 22 ==> 2/22

         請付款方式：2=固定天數，若天數小於系統日，則為當月的固定日數，反之則為次月固定日數
                例如﹕單據日期為 2/3 ，固定天數為 19，則為 2/19 日為請付款日，若單據日期為 2/21 則請付款日為 3/19。

         2015/02/11 Rickliu 經詢問 美如、小嫻後確認 當天的單據認列當月的帳，所以即便系統日期為月底且固定請款日為月底，則認列當月的帳，
         若日後要改為 系統日期為月底且固定請款日為月底 認列次月的帳時，只要將底下的判斷改為 if Day(@aVDate) < @AddDate 即可。
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
           
           if (@AddDate = 0) -- 2015/07/15 Rickliu 預防請款日期若填寫 0 時，則回傳錯誤日期。
              set @RDate = @vDate
           else 
              set @RDate = substring(Convert(Varchar(10), DateAdd(mm, +1, @aVDate), 111), 1, 8) + Convert(Varchar(2), @Last_Day)
        end
     end
     else
     begin 
       Select @RDate = 
                Case
                  -- 請付款方式：計算天數
                  When @Kind = 1 then 
                       Convert(Varchar(10), DateAdd(dd, @AddDate, @aVDate) , 111)

                  /*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-
                   編碼規則： 第 1 碼為 1, 2, 3 代表 上月, 本月, 次月
                              第 2 碼為 週、月、季、年
                              第 3 碼為 第一天 或 最後一天 或 天數
                              
                   日後維護則請依上述編碼規則進行編列
                  -*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*/

                  --*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-
                  --取得上週的開始日期(預設星期日為本週開始日)
                  When @Kind = 111 then Convert(Varchar(20), DATEADD(wk, DATEDIFF(wk, 0, @aVDate),-1), 111)
                  --取得上月的第一天
                  when @Kind = 121 then Convert(Varchar(20), Dateadd(mm, -1, Dateadd(mm, DATEDIFF(mm, 0, @aVDate), 0)), 111)
                  --取得上月的最後一天
                  when @Kind = 122 then Convert(Varchar(20), Dateadd(dd, -1, Dateadd(mm,DATEDIFF(mm, 0, @aVDate),0)), 111)
                  --取得上季的第一天
                  when @Kind = 131 then Convert(Varchar(20), DATEADD(qq, -1, DATEADD(qq, DATEDIFF(qq, 0, DATEADD(qq, 0, @aVDate)), 0)), 111)
                  --取得上季的最後一天
                  when @Kind = 132 then Convert(Varchar(20), DateAdd(mm, +3, DATEADD(qq, -1, DATEADD(qq, DATEDIFF(qq, 0, DATEADD(qq, 0, @aVDate)), 0))) -1, 111)
                  --取得去年的第一天
                  When @Kind = 141 then Convert(Varchar(20), DATEADD(yy, -1, DATEADD(yy, DATEDIFF(yy, 0, @aVDate),0)), 111)
                  --取得去年的最後一天
                  When @Kind = 142 then Convert(Varchar(20), DATEADD(yy, DATEDIFF(yy, 0, @aVDate), -1), 111)
                  --取得上月天數
                  when @Kind = 151 then Convert(Varchar(20), Day(dateadd(dd, -1, DATEADD(mm, DATEDIFF(m, 0, @aVDate)+0, 0))))
                  --取得上季天數
                  when @Kind = 152 then Convert(Varchar(20), DateDiff(dd, DATEADD(qq, -1, DATEADD(qq, DATEDIFF(qq, 0, DATEADD(qq, 0, @aVDate)), 0)), DateAdd(mm, +3, DATEADD(qq, -1, DATEADD(qq, DATEDIFF(qq, 0, DATEADD(qq, 0, @aVDate)), 0))) -1) +1)



                  --*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-
                  --取得本週的開始日期(預設星期日為本週開始日)
                  When @Kind = 211 then Convert(Varchar(20), DATEADD(wk, DATEDIFF(wk, 0, @aVDate),-1), 111)
                  --取得本月的第一天
                  when @Kind = 221 then Convert(Varchar(20), Dateadd(mm, DATEDIFF(mm, 0, @aVDate), 0), 111)
                  --取得本月的最後一天
                  when @Kind = 222 then Convert(Varchar(20), dateadd(dd, -1, DATEADD(mm, DATEDIFF(m, 0, @aVDate) +1, 0)), 111)
                  --取得本季的第一天
                  when @Kind = 231 then Convert(Varchar(20), DATEADD(qq, +0, DATEADD(qq, DATEDIFF(qq, 0, DATEADD(qq, 0, @aVDate)), 0)), 111)
                  --取得本季的最後一天
                  when @Kind = 232 then Convert(Varchar(20), DateAdd(mm, +3, DATEADD(qq, +0, DATEADD(qq, DATEDIFF(qq, 0, DATEADD(qq, 0, @aVDate)), 0))) -1, 111)
                  --取得本年的第一天
                  when @Kind = 241 then Convert(Varchar(20), DATEADD(yy, DATEDIFF(yy, 0, @aVDate), 0), 111)
                  --取得本年的最後一天
                  When @Kind = 242 then Convert(Varchar(20), DATEADD(yy, DATEDIFF(yy, 0, @aVDate)+1, -1), 111)
                  --取得本月天數
                  when @Kind = 251 then Convert(Varchar(20), Day(dateadd(dd, -1, DATEADD(mm, DATEDIFF(m, 0, @aVDate)+1, 0))))
                  --取得本季天數
                  when @Kind = 252 then Convert(Varchar(20), DateDiff(dd, DATEADD(qq, +0, DATEADD(qq, DATEDIFF(qq, 0, DATEADD(qq, 0, @aVDate)), 0)), DateAdd(mm, +3, DATEADD(qq, +0, DATEADD(qq, DATEDIFF(qq, 0, DATEADD(qq, 0, @aVDate)), 0))) -1) +1)



                  --*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-
                  --取得下週的開始日期(預設星期日為本週開始日)
                  When @Kind = 311 then Convert(Varchar(20), DATEADD(wk, DATEDIFF(wk,0, @aVDate), -1+7), 111)
                  --取得下月的第一天
                  when @Kind = 321 then Convert(Varchar(20), Dateadd(mm, 1, Dateadd(mm, DATEDIFF(mm, 0, @aVDate), 0)), 111)
                  --取得下月的最後一天
                  when @Kind = 322 then Convert(Varchar(20), dateadd(dd, -1, DATEADD(mm, DATEDIFF(m, 0, @aVDate) +2, 0)), 111)
                  --取得下季的第一天
                  when @Kind = 331 then Convert(Varchar(20), DATEADD(qq, +1, DATEADD(qq, DATEDIFF(qq, 0, DATEADD(qq, 0, @aVDate)), 0)), 111)
                  --取得下季的最後一天
                  when @Kind = 332 then Convert(Varchar(20), DateAdd(mm, +3, DATEADD(qq, +1, DATEADD(qq, DATEDIFF(qq, 0, DATEADD(qq, 0, @aVDate)), 0))) -1, 111)
                  --取得明年的第一天
                  When @Kind = 341 then Convert(Varchar(20), DATEADD(yy, 1, DATEADD(yy, DATEDIFF(yy, 0, @aVDate),0)), 111)
                  --取得明年的最後一天
                  When @Kind = 342 then Convert(Varchar(20), DATEADD(yy, DATEDIFF(yy, 0, @aVDate)+2, -1), 111)
                  --取得下月天數
                  when @Kind = 351 then Convert(Varchar(20), Day(dateadd(dd, -1, DATEADD(mm, DATEDIFF(m, 0, @aVDate)+2, 0))))
                  --取得下季天數
                  when @Kind = 352 then Convert(Varchar(20), DateDiff(dd, DATEADD(qq, +1, DATEADD(qq, DATEDIFF(qq, 0, DATEADD(qq, 0, @aVDate)), 0)), DateAdd(mm, +3, DATEADD(qq, +1, DATEADD(qq, DATEDIFF(qq, 0, DATEADD(qq, 0, @aVDate)), 0))) -1) +1)
                  else 
                    ''
                end
     end
  end
  
  
  Return(@RDate)
end
GO
