USE [DW]
GO
/****** Object:  Table [dbo].[TB_Range_Swaredt]    Script Date: 07/24/2017 14:43:54 ******/
DROP TABLE [dbo].[TB_Range_Swaredt]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[TB_Range_Swaredt](
	[wd_skno] [nvarchar](30) NULL,
	[sk_name] [nvarchar](60) NULL,
	[sk_save] [decimal](20, 6) NULL,
	[wd_Qty_atot] [decimal](38, 6) NULL,
	[wd_Save_atot] [decimal](38, 9) NULL,
	[Range_Qty1] [decimal](38, 6) NULL,
	[Range_Save1] [decimal](38, 9) NULL,
	[Range_Qty2] [decimal](38, 6) NULL,
	[Range_Save2] [decimal](38, 9) NULL,
	[Range_Qty3] [decimal](38, 6) NULL,
	[Range_Save3] [decimal](38, 9) NULL,
	[Range_Qty4] [decimal](38, 6) NULL,
	[Range_Save4] [decimal](38, 9) NULL,
	[Range_Qty5] [decimal](38, 6) NULL,
	[Range_Save5] [decimal](38, 9) NULL,
	[Range_Qty6] [decimal](38, 6) NULL,
	[Range_Save6] [decimal](38, 9) NULL
) ON [PRIMARY]
GO
