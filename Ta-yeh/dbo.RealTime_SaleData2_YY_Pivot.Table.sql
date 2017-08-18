USE [DW]
GO
/****** Object:  Table [dbo].[RealTime_SaleData2_YY_Pivot]    Script Date: 07/24/2017 14:43:51 ******/
DROP TABLE [dbo].[RealTime_SaleData2_YY_Pivot]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[RealTime_SaleData2_YY_Pivot](
	[kind] [varchar](10) NOT NULL,
	[kind_name] [varchar](1000) NULL,
	[area] [varchar](20) NOT NULL,
	[area_name] [varchar](50) NULL,
	[Data_Type] [int] NULL,
	[Y2008] [float] NOT NULL,
	[Y2009] [float] NOT NULL,
	[Y2010] [float] NOT NULL,
	[Y2011] [float] NOT NULL,
	[Y2012] [float] NOT NULL,
	[Y2013] [float] NOT NULL,
	[Y2014] [float] NOT NULL,
	[Y2015] [float] NOT NULL
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
