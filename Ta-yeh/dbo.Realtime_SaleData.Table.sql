USE [DW]
GO
/****** Object:  Table [dbo].[Realtime_SaleData]    Script Date: 07/24/2017 14:43:51 ******/
DROP TABLE [dbo].[Realtime_SaleData]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[Realtime_SaleData](
	[kind] [varchar](10) NOT NULL,
	[kind_name] [varchar](1000) NULL,
	[Area_year] [int] NOT NULL,
	[Area_month] [int] NOT NULL,
	[area] [varchar](20) NOT NULL,
	[Area_name] [varchar](50) NULL,
	[amt] [float] NULL,
	[Not_Accumulate] [varchar](1) NULL,
	[Data_Type] [int] NULL,
	[Remark] [varchar](max) NULL,
PRIMARY KEY CLUSTERED 
(
	[kind] ASC,
	[Area_year] ASC,
	[Area_month] ASC,
	[area] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
