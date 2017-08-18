USE [DW]
GO
/****** Object:  Table [dbo].[Ori_xls#Product_Characteristic]    Script Date: 07/24/2017 14:43:49 ******/
DROP TABLE [dbo].[Ori_xls#Product_Characteristic]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Ori_xls#Product_Characteristic](
	[貨品編號] [nvarchar](255) NULL,
	[系列] [nvarchar](255) NULL,
	[貨品名稱] [nvarchar](255) NULL,
	[rowid] [int] IDENTITY(1,1) NOT NULL
) ON [PRIMARY]
GO
