USE [DW]
GO
/****** Object:  Table [dbo].[Comp_EC_Consign_Order_Yahoo]    Script Date: 07/24/2017 14:43:37 ******/
DROP TABLE [dbo].[Comp_EC_Consign_Order_Yahoo]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[Comp_EC_Consign_Order_Yahoo](
	[master] [varchar](3) NOT NULL,
	[XLS_YM] [varchar](10) NULL,
	[Date_Range] [varchar](23) NULL,
	[slip_fg] [char](1) NULL,
	[xls_slip_fg_name] [varchar](2) NOT NULL,
	[order_no] [varchar](20) NULL,
	[cust_name] [varchar](20) NULL,
	[sd_spec] [varchar](255) NULL,
	[xls_tot] [int] NOT NULL,
	[chg_sd_stot] [numeric](38, 7) NOT NULL,
	[order_total] [int] NOT NULL,
	[order_cnt] [int] NOT NULL,
	[Equal] [varchar](1) NOT NULL
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO