USE [DW]
GO
/****** Object:  View [dbo].[xV_ON_ECStatement_TA13(201708����)]    Script Date: 08/18/2017 17:43:39 ******/
DROP VIEW [dbo].[xV_ON_ECStatement_TA13(201708����)]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE View [dbo].[xV_ON_ECStatement_TA13(201708����)]
AS

select a.sp_slip_fg, a.sp_date, a.sp_pdate, a.sp_no,a.sp_ordno,a.sp_ctno, a.sp_ctname, a.sp_sales, a.sp_maker, a.sp_rem, 
       b.sd_no, b.sd_skno, b.sd_name, b.sd_whno, b.sd_spec, b.sd_unit, 
       sd_qty= case sp_slip_fg when '2' then sd_qty when '3' then -sd_qty end, 
       sd_price=case sp_slip_fg when '2' then sd_price*1.05 when '3' then -sd_price*1.05 end, 
       c.cusname, c.telday, c.cellphone, c.address, c.product, c.priceinfo, c.sender, c.sendaddr 
  from Ori_TA#SSLIP a 
  join Ori_TA#sslpdt b on a.sp_class=b.sd_class and a.sp_slip_fg=b.sd_slip_fg and a.sp_no=b.sd_no 
  join Ori_TA#extratayeh c on b.sd_spec COLLATE Chinese_Taiwan_Stroke_CI_AS=c.so COLLATE Chinese_Taiwan_Stroke_CI_AS and c.priceinfo<>'' 
 where a.sp_class='1' and a.sp_slip_fg IN ('2', '3')
GO
