USE [DW]
GO
/****** Object:  Table [dbo].[Ori_xls#Sales_Invoice_To_Receipt_Voucher]    Script Date: 07/24/2017 14:43:49 ******/
DROP TABLE [dbo].[Ori_xls#Sales_Invoice_To_Receipt_Voucher]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[Ori_xls#Sales_Invoice_To_Receipt_Voucher](
	[SD_DATE] [nvarchar](255) NULL,
	[SD_DCR] [nvarchar](255) NULL,
	[SD_CTNO] [nvarchar](255) NULL,
	[SD_CTNAME] [nvarchar](255) NULL,
	[C_4101_AMT] [float] NULL,
	[C_2194_AMT] [float] NULL,
	[D_1123_AMT] [float] NULL,
	[SP_MKMAN] [nvarchar](255) NULL,
	[XlsFileName] [varchar](255) NULL,
	[SplitFileName] [varchar](255) NULL,
	[Imp_date] [datetime] NOT NULL,
	[rowid] [int] IDENTITY(1,1) NOT NULL
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
