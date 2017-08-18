USE [DW]
GO
/****** Object:  View [dbo].[uV_Year_VendorTrans_Detail]    Script Date: 08/18/2017 17:18:53 ******/
DROP VIEW [dbo].[uV_Year_VendorTrans_Detail]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE VIEW [dbo].[uV_Year_VendorTrans_Detail]
AS
--

select distinct 
       m.sp_class,
       m.sp_slip_fg,
       m.Chg_sp_slip_fg,
       Convert(Varchar(7), m.sp_date, 111) as sp_ym,
       m.sp_ctno,
       m.sp_ctname,
       m.ct_sname,
       m.sp_no,
       m.sp_date,
       d.sd_skno,
       d.sd_name,
       d.sk_bcode,
       d.sd_seqfld,
       [sd_qty]= -- 數量
         Case 
           WHEN d.SD_SLIP_FG in ('0', '2', '5', '6', '8', 'C', 'R') THEN IsNull(d.SD_QTY, 0)
           WHEN d.SD_SLIP_FG in ('1', '3', '4', '7', '9') THEN Isnull(-d.SD_QTY, 0)
           else 0
         END, 
       d.sd_price, -- 單價
       [sd_stot]= -- 小計
         Case 
           WHEN d.SD_SLIP_FG in ('0', '2', '4', '6', '8', 'C', 'R') THEN Isnull(d.SD_STOT, 0)
           WHEN d.SD_SLIP_FG in ('1', '3', '4', '7', '9') THEN Isnull(-d.SD_STOT, 0)
           else 0
         END,
      [sp_tax]= -- 單據稅金
       Case
         when d.sd_seqfld = 1 
         then 
           Case 
             WHEN d.SD_SLIP_FG in ('0', '2', '4', '6', '8', 'C', 'R') THEN Isnull(m.sp_tax, 0)
             WHEN d.SD_SLIP_FG in ('1', '3', '4', '7', '9') THEN Isnull(-m.sp_tax, 0)
             else 0
           END
         else 0
       end,
       Case
         when d.sd_seqfld = 1 then m.sp_tot
         else 0
       end sp_tot, -- 單據總金額
       m.sp_invoice, -- 發票號碼
       d1.last_date -- 最後一次交易日期
  from Fact_sslip m
       inner join Fact_sslpdt d
          on m.sp_class = d.sd_class
         and m.sp_slip_fg = d.sd_slip_fg
         and m.sp_no = d.sd_no
         and m.sp_date = d.sd_date 
        left join 
          (select sd_class, sd_slip_fg, sd_ctno, MAX(sd_date) AS last_date 
             from Fact_sslpdt
            group by sd_class, sd_slip_fg, sd_ctno
          ) d1
          on m.sp_class = d1.sd_class
         and m.sp_slip_fg = d1.sd_slip_fg
         and m.sp_ctno = d1.sd_ctno
 where 1=1
   and m.sp_class ='0' -- 進退貨
 --order by m.sp_slip_fg, m.sp_date, m.sp_ctno, m.sp_no, d.sd_seqfld
GO
