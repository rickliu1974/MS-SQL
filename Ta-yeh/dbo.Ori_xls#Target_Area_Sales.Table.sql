USE [DW]
GO
/****** Object:  Table [dbo].[Ori_xls#Target_Area_Sales]    Script Date: 07/24/2017 14:43:50 ******/
DROP TABLE [dbo].[Ori_xls#Target_Area_Sales]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[Ori_xls#Target_Area_Sales](
	[Year] [float] NULL,
	[Month] [float] NULL,
	[Area] [nvarchar](255) NULL,
	[Area_name ] [nvarchar](255) NULL,
	[Amt] [float] NULL,
	[xlsFileName] [varchar](255) NULL,
	[SplitFileName] [varchar](255) NULL,
	[imp_date] [datetime] NOT NULL,
	[rowid] [int] IDENTITY(1,1) NOT NULL
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
