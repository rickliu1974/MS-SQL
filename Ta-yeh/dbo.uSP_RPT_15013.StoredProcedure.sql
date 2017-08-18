USE [DW]
GO
/****** Object:  StoredProcedure [dbo].[uSP_RPT_15013]    Script Date: 07/24/2017 14:44:00 ******/
DROP PROCEDURE [dbo].[uSP_RPT_15013]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
Create Procedure [dbo].[uSP_RPT_15013](@sBDate Varchar(10) ='', @nMonth Int = 3, @nNo Int = 1)
as
Begin
  Declare @BDate DateTime
  
  If @sBDate = ''
     Set @BDate = Getdate()
  else
     Set @BDate = Convert(DateTime, @sBDate)
  
  Set @nMonth = @nMonth * -1

-- 2015/01/27 ¨ÑÀ³°ÓµûÅ²
select m.ct_no, m.ct_name, m.ct_sname,
       Convert(Varchar(4), Year(getdate())) as YY,
       Convert(Varchar(10), DateAdd(mm, @nMonth, @BDate), 111) as BDate,
       Convert(Varchar(10), getdate(), 111) as EDate,
       max(sp_date) as sp_date,
       count(sp_no) as sp_cnt,
       Sum(sp_tot) as sp_tot,
       Replace(RTrim((select distinct chg_skno_bkind_name +'  ' 
                       from fact_sslpdt d
                      where sd_slip_fg = '0'
                        and sd_date >= Convert(Varchar(10), DateAdd(mm, @nMonth, @BDate), 111)
                        and m.ct_no = d.sd_ctno
                        for xml path(''))
                    ), '  ', '¡B'
              ) as cc,
       'TA'+Substring(Convert(Varchar(4), Year(@BDate)), 3, 2)+
       Substring(Convert(Varchar(1000), Row_number() Over(Order by m.ct_no) + 1000+@nNo-1), 2, 4)
       as sno
  from fact_pcust m
       inner join
       (select *
          from fact_sslip d
         where sp_slip_fg = '0'
           and sp_date >= Convert(Varchar(10), DateAdd(mm, @nMonth, @BDate), 111)
       ) d 
        on m.ct_no = d.sp_ctno
       and m.ct_class = '2'
 where 1=1
 group by m.ct_no, m.ct_name, m.ct_sname
 order by 1, 2
end
GO
