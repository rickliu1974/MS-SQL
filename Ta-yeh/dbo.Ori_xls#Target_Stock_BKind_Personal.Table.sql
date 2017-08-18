USE [DW]
GO
/****** Object:  Table [dbo].[Ori_xls#Target_Stock_BKind_Personal]    Script Date: 07/24/2017 14:43:51 ******/
DROP TABLE [dbo].[Ori_xls#Target_Stock_BKind_Personal]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[Ori_xls#Target_Stock_BKind_Personal](
	[year] [float] NULL,
	[month] [float] NULL,
	[emp_no] [nvarchar](255) NULL,
	[target_AA] [float] NULL,
	[target_AB] [float] NULL,
	[target_AC] [float] NULL,
	[target_AD] [float] NULL,
	[target_AE] [float] NULL,
	[target_OT] [float] NULL,
	[xlsFileName] [varchar](255) NULL,
	[SplitFileName] [varchar](255) NULL,
	[imp_date] [datetime] NOT NULL,
	[rowid] [int] IDENTITY(1,1) NOT NULL
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
