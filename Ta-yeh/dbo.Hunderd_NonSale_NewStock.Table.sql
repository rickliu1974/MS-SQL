USE [DW]
GO
/****** Object:  Table [dbo].[Hunderd_NonSale_NewStock]    Script Date: 07/24/2017 14:43:43 ******/
DROP TABLE [dbo].[Hunderd_NonSale_NewStock]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[Hunderd_NonSale_NewStock](
	[kind] [varchar](1) NOT NULL,
	[kind_Name] [varchar](4) NOT NULL,
	[ct_no8] [nvarchar](8) NULL,
	[ct_sname8] [nvarchar](127) NULL,
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
	[Chg_New_Arrival_YM] [varchar](7) NULL,
	[Chg_New_First_Qty] [float] NULL,
	[qty] [decimal](38, 6) NOT NULL,
	[amt] [decimal](38, 6) NOT NULL,
	[update_datetime] [datetime] NOT NULL
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
