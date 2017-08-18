USE [DW]
GO
/****** Object:  Table [dbo].[Comp_EC_Order_SuperMarket_tmp]    Script Date: 07/24/2017 14:43:37 ******/
DROP TABLE [dbo].[Comp_EC_Order_SuperMarket_tmp]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[Comp_EC_Order_SuperMarket_tmp](
	[slip_fg] [varchar](1) NOT NULL,
	[Order_No] [nvarchar](255) NULL,
	[xls_tot] [int] NULL,
	[xlsfileName] [varchar](255) NULL,
	[rowid] [int] IDENTITY(1,1) NOT NULL,
	[XLS_YM] [varchar](7) NULL,
	[b_sp_pdate] [varchar](10) NULL,
	[e_sp_pdate] [varchar](10) NULL,
	[Create_datetime] [datetime] NOT NULL
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
