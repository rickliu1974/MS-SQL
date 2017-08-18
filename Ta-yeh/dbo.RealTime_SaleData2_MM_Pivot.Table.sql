USE [DW]
GO
/****** Object:  Table [dbo].[RealTime_SaleData2_MM_Pivot]    Script Date: 07/24/2017 14:43:51 ******/
DROP TABLE [dbo].[RealTime_SaleData2_MM_Pivot]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[RealTime_SaleData2_MM_Pivot](
	[kind] [varchar](10) NOT NULL,
	[kind_name] [varchar](1000) NULL,
	[area_year] [int] NOT NULL,
	[area] [varchar](20) NOT NULL,
	[area_name] [varchar](50) NULL,
	[Data_Type] [int] NULL,
	[M01] [float] NOT NULL,
	[M02] [float] NOT NULL,
	[M03] [float] NOT NULL,
	[M04] [float] NOT NULL,
	[M05] [float] NOT NULL,
	[M06] [float] NOT NULL,
	[M07] [float] NOT NULL,
	[M08] [float] NOT NULL,
	[M09] [float] NOT NULL,
	[M10] [float] NOT NULL,
	[M11] [float] NOT NULL,
	[M12] [float] NOT NULL,
	[M_SUM] [float] NOT NULL,
	[M_AVG] [float] NULL
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
