USE [DW]
GO
/****** Object:  Table [dbo].[Realtime_SaleData_YY_Pivot]    Script Date: 07/24/2017 14:43:51 ******/
DROP TABLE [dbo].[Realtime_SaleData_YY_Pivot]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[Realtime_SaleData_YY_Pivot](
	[kind] [varchar](10) NOT NULL,
	[kind_name] [varchar](1000) NULL,
	[area] [varchar](20) NOT NULL,
	[area_name] [varchar](50) NULL,
	[Data_Type] [int] NULL,
	[area_month] [int] NOT NULL,
	[2008] [float] NOT NULL,
	[2009] [float] NOT NULL,
	[2010] [float] NOT NULL,
	[2011] [float] NOT NULL,
	[2012] [float] NOT NULL,
	[2013] [float] NOT NULL,
	[2014] [float] NOT NULL,
	[2015] [float] NOT NULL,
	[2016] [float] NOT NULL,
	[2017] [float] NOT NULL
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
