USE [DW]
GO
/****** Object:  Table [dbo].[RealTime_SaleDataV2_tmp]    Script Date: 07/24/2017 14:43:52 ******/
DROP TABLE [dbo].[RealTime_SaleDataV2_tmp]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[RealTime_SaleDataV2_tmp](
	[kind] [varchar](10) NOT NULL,
	[kind_name] [varchar](255) NULL,
	[Area_year] [int] NOT NULL,
	[Area_month] [int] NOT NULL,
	[Area] [varchar](20) NOT NULL,
	[Area_name] [varchar](50) NULL,
	[Amt] [float] NULL,
	[Not_Accumulate] [varchar](1) NULL,
	[Data_Type] [int] NULL,
	[Remark] [varchar](255) NULL,
PRIMARY KEY CLUSTERED 
(
	[kind] ASC,
	[Area_year] ASC,
	[Area_month] ASC,
	[Area] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
