USE [DW]
GO
/****** Object:  Table [dbo].[Cust_MonthSale_Stock]    Script Date: 07/24/2017 14:43:39 ******/
DROP TABLE [dbo].[Cust_MonthSale_Stock]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[Cust_MonthSale_Stock](
	[kind] [varchar](1) NOT NULL,
	[kind_Name] [varchar](4) NOT NULL,
	[ct_no8] [nvarchar](8) NULL,
	[ct_sname8] [nvarchar](121) NULL,
	[chg_ct_fld3] [nvarchar](60) NULL,
	[Chg_Hunderd_Customer_Name] [nvarchar](255) NULL,
	[Chg_IS_Lan_Custom] [varchar](1) NULL,
	[sd_skno] [nvarchar](30) NULL,
	[sd_name] [nvarchar](60) NULL,
	[Chg_Wd_AA_first_Qty] [decimal](20, 6) NULL,
	[Chg_Wd_AB_first_Qty] [decimal](20, 6) NULL,
	[Chg_Wd_AA_last_Qty] [decimal](32, 6) NULL,
	[Chg_Wd_AB_last_Qty] [decimal](32, 6) NULL,
	[Chg_skno_accno] [nvarchar](1) NULL,
	[Chg_skno_accno_Name] [nvarchar](255) NULL,
	[Chg_skno_BKind] [nvarchar](2) NULL,
	[Chg_skno_Bkind_Name] [nvarchar](255) NULL,
	[Chg_skno_SKind] [nvarchar](4) NULL,
	[Chg_skno_SKind_Name] [nvarchar](255) NULL,
	[Chg_kind_Name] [nchar](30) NULL,
	[Chg_sp_date_YM] [varchar](10) NULL,
	[Chg_sp_sales] [nvarchar](10) NULL,
	[Chg_sales_Name] [nvarchar](10) NULL,
	[Chg_sale_Month_Master] [varchar](1) NOT NULL,
	[Chg_sale_Month_MasterName] [nvarchar](255) NULL,
	[Chg_Hunderd_Customer] [varchar](1) NULL,
	[Chg_Stock_NonSales] [varchar](1) NULL,
	[Chg_is_dead_stock] [varchar](1) NULL,
	[chg_dead_stock_ym] [varchar](7) NULL,
	[chg_new_arrival_ym] [varchar](7) NULL,
	[sd2_qty] [decimal](38, 6) NULL,
	[sd6_qty] [decimal](38, 6) NULL,
	[sd7_qty] [decimal](38, 6) NULL,
	[sd2_stot] [decimal](38, 6) NULL,
	[sd6_stot] [decimal](38, 6) NULL,
	[sd7_stot] [decimal](38, 6) NULL,
	[all_count_ctno] [int] NOT NULL,
	[all_count_skno] [int] NOT NULL,
	[master_count_ctno] [int] NOT NULL,
	[master_count_skno] [int] NOT NULL,
	[nonsale_count_ctno] [int] NOT NULL,
	[nonsale_count_skno] [int] NOT NULL,
	[newstock_count_ctno] [int] NOT NULL,
	[newstock_count_skno] [int] NOT NULL,
	[deadstock_count_ctno] [int] NOT NULL,
	[deadstock_count_skno] [int] NOT NULL,
	[update_datetime] [datetime] NOT NULL
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
