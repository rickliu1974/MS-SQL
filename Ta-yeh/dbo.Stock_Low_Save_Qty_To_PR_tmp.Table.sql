USE [DW]
GO
/****** Object:  Table [dbo].[Stock_Low_Save_Qty_To_PR_tmp]    Script Date: 07/24/2017 14:43:53 ******/
DROP TABLE [dbo].[Stock_Low_Save_Qty_To_PR_tmp]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[Stock_Low_Save_Qty_To_PR_tmp](
	[### 主檔資料 ###] [varchar](16) NOT NULL,
	[br_no] [nvarchar](11) NULL,
	[br_date1] [varchar](10) NULL,
	[br_sales] [varchar](5) NOT NULL,
	[br_dpno] [varchar](5) NOT NULL,
	[br_maker] [varchar](5) NOT NULL,
	[br_tot] [decimal](38, 7) NULL,
	[br_tal_rec] [int] NULL,
	[br_rem] [varchar](40) NULL,
	[br_ispack] [varchar](1) NOT NULL,
	[br_npack] [int] NULL,
	[br_surefg] [int] NOT NULL,
	[### 明細資料 ###] [varchar](16) NOT NULL,
	[bd_date2] [varchar](10) NULL,
	[bd_ctno] [nchar](10) NULL,
	[bd_ctname] [nchar](60) NULL,
	[bd_skno] [nvarchar](30) NULL,
	[bd_name] [nvarchar](60) NULL,
	[bd_unit] [nvarchar](8) NULL,
	[bd_qty] [decimal](22, 6) NULL,
	[bd_price] [decimal](20, 6) NOT NULL,
	[bd_stot] [decimal](38, 7) NULL,
	[bd_unit_fg] [int] NOT NULL,
	[bd_rem] [varchar](40) NULL,
	[bd_rate_nm] [nvarchar](8) NULL,
	[bd_rate] [decimal](20, 6) NULL,
	[bd_is_pack] [int] NOT NULL,
	[bd_surefg] [int] NOT NULL,
	[bd_seqfld] [bigint] NULL,
	[### 數量檢查 ###] [varchar](16) NOT NULL,
	[chk_sk_bqty] [decimal](20, 6) NULL,
	[chk_wd_save_qty] [decimal](22, 6) NULL,
	[chk_wd_last_diff_qty] [decimal](34, 6) NULL,
	[chk_od_qty] [numeric](38, 6) NOT NULL,
	[chk_bd_qty] [numeric](38, 6) NOT NULL
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
