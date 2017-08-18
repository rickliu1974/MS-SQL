USE [DW]
GO
/****** Object:  StoredProcedure [dbo].[uSP_Range_Swaredt]    Script Date: 08/18/2017 17:18:56 ******/
DROP PROCEDURE [dbo].[uSP_Range_Swaredt]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
Create Proc [dbo].[uSP_Range_Swaredt](@sYM Varchar(7), @sWD_skno Varchar(20)= '')
as
begin
  Declare @Proc Varchar(50) = 'uSP_Range_Swaredt'
  Declare @CR Varchar(4) = char(13)+char(10)
  Declare @Msg Varchar(Max)

  Declare @wd_skno Varchar(10), @sk_name Varchar(30), @sk_save float
   
  Declare @Range_Qty1 Int, @Range_Qty2 Int, @Range_Qty3 Int,
          @Range_Qty4 Int, @Range_Qty5 Int, @Range_Qty6 Int,
          @wd_Qty_atot Int
           
  Declare @Range_Save1 Float, @Range_Save2 Float, @Range_Save3 Float,
          @Range_Save4 Float, @Range_Save5 Float, @Range_Save6 Float,
          @wd_Save_atot Float

  if Len(RTrim(Isnull(@sYM, ''))) <> 7 
     Set @sYM = Convert(Varchar(7), Getdate(), 111)

  IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[TB_Range_Swaredt]') AND type in (N'U'))
     Drop Table TB_Range_Swaredt

  Set @Msg = '產出庫存臨時帳齡表'
  Exec uSP_Sys_Write_Log @Proc, @Msg, '', 0
  select RTrim(wd_skno) as wd_skno, RTrim(sk_name) as sk_name, sk_save,
         Sum(wd_Qty_atot) as wd_Qty_atot, Sum(wd_Save_atot) as wd_Save_atot,
         Sum(Range_Qty1) as Range_Qty1, Sum(Range_Save1) as Range_Save1,
         Sum(Range_Qty2) as Range_Qty2, Sum(Range_Save2) as Range_Save2,
         Sum(Range_Qty3) as Range_Qty3, Sum(Range_Save3) as Range_Save3,
         Sum(Range_Qty4) as Range_Qty4, Sum(Range_Save4) as Range_Save4,
         Sum(Range_Qty5) as Range_Qty5, Sum(Range_Save5) as Range_Save5,
         Sum(Range_Qty6) as Range_Qty6, Sum(Range_Save6) as Range_Save6
         into TB_Range_Swaredt
    from (select wd_ym, wd_skno, sk_name, sk_save,
                 case
                   when amonth = 1 then sum(wd_amt)
                   else 0
                 end Range_Qty1,
                 case
                   when amonth = 2 then sum(wd_amt)
                   else 0
                 end Range_Qty2,
                 case
                   when amonth = 3 then sum(wd_amt)
                   else 0
                 end Range_Qty3,
                 case
                   when amonth > 3 and amonth <=6 then sum(wd_amt)
                   else 0
                 end Range_Qty4,
                 case
                   when amonth > 6 and amonth <=12 then sum(wd_amt)
                   else 0
                 end Range_Qty5,
                 case
                   when amonth > 12 then sum(wd_amt)
                   else 0
                 end Range_Qty6,
                 case
                   when amonth = 1 then sum(wd_save_tot)
                   else 0
                 end Range_Save1,
                 case
                   when amonth = 2 then sum(wd_save_tot)
                   else 0
                 end Range_Save2,
                 case
                   when amonth = 3 then sum(wd_save_tot)
                   else 0
                 end Range_Save3,
                 case
                   when amonth > 3 and amonth <=6 then sum(wd_save_tot)
                   else 0
                 end Range_Save4,
                 case
                   when amonth > 6 and amonth <=12 then sum(wd_save_tot)
                   else 0
                 end Range_Save5,
                 case
                   when amonth > 12 then sum(wd_save_tot)
                   else 0
                 end Range_Save6,
                 sum(wd_amt) as wd_Qty_atot, sum(wd_save_tot) as wd_save_atot
            from (select wd_ym, 
                         datediff(mm, dateadd(mm, 1, convert(datetime, wd_ym+'/01')-1), getdate()) as amonth,
                         wd_skno, d.sk_name, d.sk_save,
                         wd_amt, wd_save_tot
                    from fact_swaredt m
                         left join fact_sstock d
                           on m.wd_skno = d.sk_no
                   where wd_ym <= @sYM
                     and wd_skno like '%'+@sWD_skno+'%'
                     and wd_no in ('AA', 'AB', 'AC')
                 ) m
           group by wd_ym, wd_skno, sk_name, amonth, sk_save
         ) m
   group by wd_skno, sk_name, sk_save
   order by wd_skno, sk_name, sk_save

  Declare Cur_Range_Swared Cursor for
    Select wd_skno, sk_name, sk_save,
           Range_Qty1, Range_Qty2, Range_Qty3,
           Range_Qty4, Range_Qty5, Range_Qty6
      from TB_Range_Swaredt

  open Cur_Range_Swared
  fetch next from Cur_Range_Swared into @wd_skno, @sk_name, @sk_save,
                                        @Range_Qty1, @Range_Qty2, @Range_Qty3,
                                        @Range_Qty4, @Range_Qty5, @Range_Qty6

  while @@fetch_status =0
  begin
     Set @Msg = 'Before Data..Process sk_no:['+@wd_skno+'], sk_name:['+@sk_name+']'+@CR+
                '@Range_Qty6:['+Convert(Varchar(100), @Range_Qty6)+'], @Range_Qty5:['+Convert(Varchar(100), @Range_Qty5)+'], '+
                '@Range_Qty4:['+Convert(Varchar(100), @Range_Qty4)+'], @Range_Qty3:['+Convert(Varchar(100), @Range_Qty3)+'], '+
                '@Range_Qty2:['+Convert(Varchar(100), @Range_Qty2)+'], @Range_Qty1:['+Convert(Varchar(100), @Range_Qty1)+']'
     Exec uSP_Sys_Write_Log @Proc, @Msg, '', 0

     While @Range_Qty6 > 0 
     begin
        print 'Run Range_Qty6'
        if @Range_Qty5 < 0
        begin
          If @Range_Qty6 + @Range_Qty5 <= 0
          begin
             Set @Range_Qty5 = @Range_Qty6 + @Range_Qty5
             Set @Range_Qty6 = 0
          end
          else
          begin
             Set @Range_Qty6 = @Range_Qty6 + @Range_Qty5
             Set @Range_Qty5 = 0
          end
          Set @Msg = 'Run Range_Qty6 >0 and Range_Qty5 <0, @Range_Qty6:['+Convert(Varchar(100), @Range_Qty6)+'], @Range_Qty5:['+Convert(Varchar(100), @Range_Qty5)+']'
          Exec uSP_Sys_Write_Log @Proc, @Msg, '', 0
          Continue
        end
        
        if @Range_Qty4 < 0
        begin
          If @Range_Qty6 + @Range_Qty4 <= 0
          begin
             Set @Range_Qty4 = @Range_Qty6 + @Range_Qty4
             Set @Range_Qty6 = 0
          end
          else
          begin
             Set @Range_Qty6 = @Range_Qty6 + @Range_Qty4
             Set @Range_Qty4 = 0
          end
          Set @Msg = 'Run Range_Qty6 >0 and Range_Qty4 <0, @Range_Qty6:['+Convert(Varchar(100), @Range_Qty6)+'], @Range_Qty4:['+Convert(Varchar(100), @Range_Qty4)+']'
          Exec uSP_Sys_Write_Log @Proc, @Msg, '', 0
          Continue
        end
        
        if @Range_Qty3 < 0
        begin
          If @Range_Qty6 + @Range_Qty3 <= 0
          begin
             Set @Range_Qty3 = @Range_Qty6 + @Range_Qty3
             Set @Range_Qty6 = 0
          end
          else
          begin
             Set @Range_Qty6 = @Range_Qty6 + @Range_Qty3
             Set @Range_Qty3 = 0
          end
          Set @Msg = 'Run Range_Qty6 >0 and Range_Qty3 <0, @Range_Qty6:['+Convert(Varchar(100), @Range_Qty6)+'], @Range_Qty3:['+Convert(Varchar(100), @Range_Qty3)+']'
          Exec uSP_Sys_Write_Log @Proc, @Msg, '', 0
          Continue
        end

        if @Range_Qty2 < 0
        begin
          If @Range_Qty6 + @Range_Qty2 <= 0
          begin
             Set @Range_Qty2 = @Range_Qty6 + @Range_Qty2
             Set @Range_Qty6 = 0
          end
          else
          begin
             Set @Range_Qty6 = @Range_Qty6 + @Range_Qty2
             Set @Range_Qty2 = 0
          end
          Set @Msg = 'Run Range_Qty6 >0 and Range_Qty2 <0, @Range_Qty6:['+Convert(Varchar(100), @Range_Qty6)+'], @Range_Qty2:['+Convert(Varchar(100), @Range_Qty2)+']'
          Exec uSP_Sys_Write_Log @Proc, @Msg, '', 0
          Continue
        end

        if @Range_Qty1 < 0
        begin
          If @Range_Qty6 + @Range_Qty1 <= 0
          begin
              Set @Range_Qty1 = @Range_Qty6 + @Range_Qty1
              Set @Range_Qty6 = 0
         end
          else
          begin
             Set @Range_Qty6 = @Range_Qty6 + @Range_Qty1
             Set @Range_Qty1 = 0
          end
          Set @Msg = 'Run Range_Qty6 >0 and Range_Qty1 <0, @Range_Qty6:['+Convert(Varchar(100), @Range_Qty6)+'], @Range_Qty1:['+Convert(Varchar(100), @Range_Qty1)+']'
          Exec uSP_Sys_Write_Log @Proc, @Msg, '', 0
          Continue
        end
        Break
     end

     While @Range_Qty5 > 0 
     begin
        print 'Run Range_Qty5'
        if @Range_Qty4 < 0
        begin
          If @Range_Qty5 + @Range_Qty4 <= 0
          begin
             Set @Range_Qty4 = @Range_Qty5 + @Range_Qty4
             Set @Range_Qty5 = 0
          end
          else
          begin
             Set @Range_Qty5 = @Range_Qty5 + @Range_Qty4
             Set @Range_Qty4 = 0
          end
          Set @Msg = 'Run Range_Qty5 >0 and Range_Qty4 <0, @Range_Qty5:['+Convert(Varchar(100), @Range_Qty5)+'], @Range_Qty4:['+Convert(Varchar(100), @Range_Qty4)+']'
          Exec uSP_Sys_Write_Log @Proc, @Msg, '', 0
          Continue
        end
        
        if @Range_Qty3 < 0
        begin
          If @Range_Qty5 + @Range_Qty3 <= 0
          begin
             Set @Range_Qty3 = @Range_Qty5 + @Range_Qty3
             Set @Range_Qty5 = 0
          end
          else
          begin
             Set @Range_Qty5 = @Range_Qty5 + @Range_Qty3
             Set @Range_Qty3 = 0
          end
          Set @Msg = 'Run Range_Qty5 >0 and Range_Qty3 <0, @Range_Qty5:['+Convert(Varchar(100), @Range_Qty5)+'], @Range_Qty3:['+Convert(Varchar(100), @Range_Qty3)+']'
          Exec uSP_Sys_Write_Log @Proc, @Msg, '', 0
          Continue
        end
        
        if @Range_Qty2 < 0
        begin
          If @Range_Qty5 + @Range_Qty2 <= 0
          begin
             Set @Range_Qty2 = @Range_Qty5 + @Range_Qty2
             Set @Range_Qty5 = 0
          end
          else
          begin
             Set @Range_Qty5 = @Range_Qty5 + @Range_Qty2
             Set @Range_Qty2 = 0
          end
          Set @Msg = 'Run Range_Qty5 >0 and Range_Qty2 <0, @Range_Qty5:['+Convert(Varchar(100), @Range_Qty5)+'], @Range_Qty2:['+Convert(Varchar(100), @Range_Qty2)+']'
          Exec uSP_Sys_Write_Log @Proc, @Msg, '', 0
          Continue
        end

        if @Range_Qty1 < 0
        begin
          If @Range_Qty5 + @Range_Qty1 <= 0
          begin
             Set @Range_Qty1 = @Range_Qty5 + @Range_Qty1
             Set @Range_Qty5 = 0
          end
          else
          begin
             Set @Range_Qty5 = @Range_Qty5 + @Range_Qty1
             Set @Range_Qty1 = 0
          end
          Set @Msg = 'Run Range_Qty5 >0 and Range_Qty1 <0, @Range_Qty5:['+Convert(Varchar(100), @Range_Qty5)+'], @Range_Qty1:['+Convert(Varchar(100), @Range_Qty1)+']'
          Exec uSP_Sys_Write_Log @Proc, @Msg, '', 0
          Continue
        end
        Break
     end

     While @Range_Qty4 > 0 
     begin
        print 'Run Range_Qty4'
        if @Range_Qty3 < 0
        begin
          If @Range_Qty4 + @Range_Qty3 <= 0
          begin
             Set @Range_Qty3 = @Range_Qty4 + @Range_Qty3
             Set @Range_Qty4 = 0
          end
          else
          begin
             Set @Range_Qty4 = @Range_Qty4 + @Range_Qty3
             Set @Range_Qty3 = 0
          end
          Set @Msg = 'Run Range_Qty4 >0 and Range_Qty3 <0, @Range_Qty4:['+Convert(Varchar(100), @Range_Qty4)+'], @Range_Qty3:['+Convert(Varchar(100), @Range_Qty3)+']'
          Exec uSP_Sys_Write_Log @Proc, @Msg, '', 0
          Continue
        end
        
        if @Range_Qty2 < 0
        begin
          If @Range_Qty4 + @Range_Qty2 <= 0
          begin
             Set @Range_Qty2 = @Range_Qty4 + @Range_Qty2
             Set @Range_Qty4 = 0
          end
          else
          begin
             Set @Range_Qty4 = @Range_Qty4 + @Range_Qty2
             Set @Range_Qty2 = 0
          end
          Set @Msg = 'Run Range_Qty4 >0 and Range_Qty2 <0, @Range_Qty4:['+Convert(Varchar(100), @Range_Qty4)+'], @Range_Qty2:['+Convert(Varchar(100), @Range_Qty2)+']'
          Exec uSP_Sys_Write_Log @Proc, @Msg, '', 0
          Continue
        end

        if @Range_Qty1 < 0
        begin
          If @Range_Qty4 + @Range_Qty1 <= 0
          begin
             Set @Range_Qty1 = @Range_Qty4 + @Range_Qty1
             Set @Range_Qty4 = 0
          end
          else
          begin
             Set @Range_Qty4 = @Range_Qty4 + @Range_Qty1
             Set @Range_Qty1 = 0
          end
          Set @Msg = 'Run Range_Qty4 >0 and Range_Qty1 <0, @Range_Qty4:['+Convert(Varchar(100), @Range_Qty4)+'], @Range_Qty1:['+Convert(Varchar(100), @Range_Qty1)+']'
          Exec uSP_Sys_Write_Log @Proc, @Msg, '', 0
          Continue
        end
        Break
     end

     While @Range_Qty3 > 0 
     begin
        print 'Run Range_Qty3'
        if @Range_Qty2 < 0
        begin
          If @Range_Qty3 + @Range_Qty2 <= 0
          begin
             Set @Range_Qty2 = @Range_Qty3 + @Range_Qty2
             Set @Range_Qty3 = 0
          end
          else
          begin
             Set @Range_Qty3 = @Range_Qty3 + @Range_Qty2
             Set @Range_Qty2 = 0
          end
          Set @Msg = 'Run Range_Qty3 >0 and Range_Qty1 <0, @Range_Qty3:['+Convert(Varchar(100), @Range_Qty3)+'], @Range_Qty2:['+Convert(Varchar(100), @Range_Qty2)+']'
          Exec uSP_Sys_Write_Log @Proc, @Msg, '', 0
          Continue
        end
        
        if @Range_Qty1 < 0
        begin
          If @Range_Qty3 + @Range_Qty1 <= 0
          begin
             Set @Range_Qty1 = @Range_Qty3 + @Range_Qty1
             Set @Range_Qty3 = 0
          end
          else
          begin
             Set @Range_Qty3 = @Range_Qty3 + @Range_Qty1
             Set @Range_Qty1 = 0
          end
          Set @Msg = 'Run Range_Qty3 >0 and Range_Qty1 <0, @Range_Qty3:['+Convert(Varchar(100), @Range_Qty3)+'], @Range_Qty1:['+Convert(Varchar(100), @Range_Qty1)+']'
          Exec uSP_Sys_Write_Log @Proc, @Msg, '', 0
          Continue
        end

        Set @Msg = '@Range_Qty6:['+Convert(Varchar(100), @Range_Qty6)+'], @Range_Qty5:['+Convert(Varchar(100), @Range_Qty5)+'], '+
                   '@Range_Qty4:['+Convert(Varchar(100), @Range_Qty4)+'], @Range_Qty3:['+Convert(Varchar(100), @Range_Qty3)+'], '+
                   '@Range_Qty2:['+Convert(Varchar(100), @Range_Qty2)+'], @Range_Qty1:['+Convert(Varchar(100), @Range_Qty1)+']'        
        Exec uSP_Sys_Write_Log @Proc, @Msg, '', 0
        Break
     end

     While @Range_Qty2 > 0 
     begin
        print 'Run Range_Qty2'
        if @Range_Qty1 < 0
        begin
          If @Range_Qty2 + @Range_Qty1 <= 0
          begin
             Set @Range_Qty1 = @Range_Qty2 + @Range_Qty1
             Set @Range_Qty2 = 0
          end
          else
          begin
             Set @Range_Qty2 = @Range_Qty2 + @Range_Qty1
             Set @Range_Qty1 = 0
          end
          Set @Msg = 'Run Range_Qty2 >0 and Range_Qty1 <0, @Range_Qty2:['+Convert(Varchar(100), @Range_Qty2)+'], @Range_Qty1:['+Convert(Varchar(100), @Range_Qty1)+']'
          Exec uSP_Sys_Write_Log @Proc, @Msg, '', 0
          Continue
        end
        Break
     end
     Set @Msg = 'After Data..Process sk_no:['+@wd_skno+'], sk_name:['+@sk_name+']'+@CR+
                '@Range_Qty6:['+Convert(Varchar(100), @Range_Qty6)+'], @Range_Qty5:['+Convert(Varchar(100), @Range_Qty5)+'], '+
                '@Range_Qty4:['+Convert(Varchar(100), @Range_Qty4)+'], @Range_Qty3:['+Convert(Varchar(100), @Range_Qty3)+'], '+
                '@Range_Qty2:['+Convert(Varchar(100), @Range_Qty2)+'], @Range_Qty1:['+Convert(Varchar(100), @Range_Qty1)+']'           
     Exec uSP_Sys_Write_Log @Proc, @Msg, '', 0

     update TB_Range_Swaredt
        set wd_Qty_atot = (@Range_Qty1 + @Range_Qty2 + @Range_Qty3 + @Range_Qty4 + @Range_Qty5 + @Range_Qty6),
            wd_Save_atot = (@Range_Qty1 + @Range_Qty2 + @Range_Qty3 + @Range_Qty4 + @Range_Qty5 + @Range_Qty6) * @sk_Save,
            Range_Qty6 = @Range_Qty6,
            Range_Qty5 = @Range_Qty5,
            Range_Qty4 = @Range_Qty4,
            Range_Qty3 = @Range_Qty3,
            Range_Qty2 = @Range_Qty2,
            Range_Qty1 = @Range_Qty1,
            
            Range_Save6 = @Range_Qty6 * @sk_Save,
            Range_Save5 = @Range_Qty5 * @sk_Save,
            Range_Save4 = @Range_Qty4 * @sk_Save,
            Range_Save3 = @Range_Qty3 * @sk_Save,
            Range_Save2 = @Range_Qty2 * @sk_Save,
            Range_Save1 = @Range_Qty1 * @sk_Save
      where wd_skno = @wd_skno
            
      fetch next from Cur_Range_Swared into @wd_skno, @sk_name, @sk_save,
                                            @Range_Qty1, @Range_Qty2, @Range_Qty3,
                                            @Range_Qty4, @Range_Qty5, @Range_Qty6
  end
  close Cur_Range_Swared
  deallocate Cur_Range_Swared
  
  Select wd_skno, sk_name, sk_save,
         wd_Qty_atot, wd_Save_atot,
         Range_Qty1, Range_Save1,
         Range_Qty2, Range_Save2,
         Range_Qty3, Range_Save3,
         Range_Qty4, Range_Save4,
         Range_Qty5, Range_Save5,
         Range_Qty6, Range_Save6
    from TB_Range_Swaredt
   order by wd_skno
end
GO
