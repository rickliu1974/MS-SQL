USE [DW]
GO
/****** Object:  Table [dbo].[Ori_xls#Target_Stock_BKind_Business]    Script Date: 07/24/2017 14:43:51 ******/
DROP TABLE [dbo].[Ori_xls#Target_Stock_BKind_Business]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[Ori_xls#Target_Stock_BKind_Business](
	[Year] [float] NULL,
	[Month] [float] NULL,
	[Target_AA] [float] NULL,
	[Target_AB] [float] NULL,
	[Target_AC] [float] NULL,
	[Target_AD] [float] NULL,
	[Target_AE] [float] NULL,
	[Target_OT] [float] NULL,
	[xlsFileName] [varchar](255) NULL,
	[SplitFileName] [varchar](255) NULL,
	[imp_date] [datetime] NOT NULL,
	[rowid] [int] IDENTITY(1,1) NOT NULL
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
