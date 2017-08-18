USE [DW]
GO
/****** Object:  Table [dbo].[Ori_xls#Target_Before_Sale_By_2012]    Script Date: 07/24/2017 14:43:50 ******/
DROP TABLE [dbo].[Ori_xls#Target_Before_Sale_By_2012]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Ori_xls#Target_Before_Sale_By_2012](
	[Year] [float] NULL,
	[month] [float] NULL,
	[Sale_Amt] [float] NULL,
	[Target_Amt] [float] NULL,
	[rowid] [int] IDENTITY(1,1) NOT NULL
) ON [PRIMARY]
GO
