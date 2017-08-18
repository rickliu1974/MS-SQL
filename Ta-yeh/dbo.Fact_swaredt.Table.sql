USE [DW]
GO
/****** Object:  Table [dbo].[Fact_swaredt]    Script Date: 07/24/2017 14:43:43 ******/
DROP TABLE [dbo].[Fact_swaredt]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[Fact_swaredt](
	[wd_no] [nvarchar](10) NULL,
	[wh_name] [nvarchar](30) NULL,
	[wd_yr] [char](4) NULL,
	[wd_ym] [varchar](7) NULL,
	[wd_skno] [nvarchar](30) NULL,
	[sk_name] [nvarchar](60) NULL,
	[wd_amt] [decimal](20, 6) NOT NULL,
	[wd_ave] [decimal](20, 6) NOT NULL,
	[sk_save] [decimal](20, 6) NOT NULL,
	[wd_save_tot] [decimal](38, 9) NULL,
	[cal_year] [int] NULL,
	[cal_ym] [varchar](7) NULL,
	[chg_skno_bkind3] [varchar](2) NULL,
	[chg_skno_bkind_name3] [varchar](4) NULL,
	[swaredt_update_datetime] [datetime] NOT NULL
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
