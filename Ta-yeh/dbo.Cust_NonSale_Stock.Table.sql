USE [DW]
GO
/****** Object:  Table [dbo].[Cust_NonSale_Stock]    Script Date: 07/24/2017 14:43:39 ******/
DROP TABLE [dbo].[Cust_NonSale_Stock]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[Cust_NonSale_Stock](
	[ct_no8] [nvarchar](8) NULL,
	[ct_sname8] [nvarchar](121) NULL,
	[ct_sales] [nvarchar](10) NULL,
	[chg_ct_sales_name] [nvarchar](10) NULL,
	[ct_fld3] [nvarchar](60) NULL,
	[chg_hunderd_customer] [varchar](1) NOT NULL,
	[Chg_Hunderd_Customer_Name] [nvarchar](255) NULL,
	[chg_is_lan_custom] [varchar](1) NOT NULL,
	[sk_no] [nvarchar](30) NULL,
	[sk_name] [nvarchar](60) NULL,
	[Chg_Wd_AA_first_Qty] [decimal](20, 6) NULL,
	[Chg_Wd_AB_first_Qty] [decimal](20, 6) NOT NULL,
	[Chg_Wd_AA_last_Qty] [decimal](32, 6) NOT NULL,
	[Chg_Wd_AB_last_Qty] [decimal](32, 6) NOT NULL,
	[Chg_Wd_AA_first_amt] [decimal](38, 9) NULL,
	[Chg_Wd_AB_first_amt] [decimal](38, 9) NULL,
	[Chg_Wd_AA_last_amt] [decimal](38, 6) NULL,
	[Chg_Wd_AB_last_amt] [decimal](38, 6) NULL,
	[Chg_skno_accno] [nvarchar](1) NOT NULL,
	[Chg_skno_accno_Name] [nvarchar](255) NOT NULL,
	[Chg_skno_BKind] [nvarchar](2) NOT NULL,
	[Chg_skno_Bkind_Name] [nvarchar](255) NOT NULL,
	[Chg_skno_SKind] [nvarchar](4) NULL,
	[Chg_skno_SKind_Name] [nvarchar](255) NULL,
	[Chg_kind_Name] [nchar](30) NULL,
	[Chg_is_dead_stock] [varchar](1) NOT NULL,
	[chg_dead_stock_ym] [varchar](7) NULL,
	[chg_new_arrival_ym] [varchar](7) NULL,
	[sd_qty] [decimal](38, 6) NOT NULL,
	[sd_stot] [decimal](38, 6) NOT NULL,
	[sd2_qty] [decimal](38, 6) NOT NULL,
	[sd2_stot] [decimal](38, 6) NOT NULL,
	[sd6_qty] [decimal](38, 6) NOT NULL,
	[sd6_stot] [decimal](38, 6) NOT NULL,
	[sd7_qty] [decimal](38, 6) NOT NULL,
	[sd7_stot] [decimal](38, 6) NOT NULL,
	[sale_flag] [varchar](1) NOT NULL,
	[dcount_ctno] [int] NOT NULL,
	[dcount_ctno8] [int] NOT NULL,
	[dcount_skno] [int] NOT NULL,
	[update_datetime] [datetime] NOT NULL
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
