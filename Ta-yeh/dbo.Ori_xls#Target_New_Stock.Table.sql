USE [DW]
GO
/****** Object:  Table [dbo].[Ori_xls#Target_New_Stock]    Script Date: 07/24/2017 14:43:51 ******/
DROP TABLE [dbo].[Ori_xls#Target_New_Stock]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Ori_xls#Target_New_Stock](
	[Year] [float] NULL,
	[Month] [float] NULL,
	[Target] [float] NULL,
	[rowid] [int] IDENTITY(1,1) NOT NULL
) ON [PRIMARY]
GO
