USE [DW]
GO
/****** Object:  Table [dbo].[EC_Consign_Order_Yahoo_tmp]    Script Date: 07/24/2017 14:43:40 ******/
DROP TABLE [dbo].[EC_Consign_Order_Yahoo_tmp]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[EC_Consign_Order_Yahoo_tmp](
	[F1] [nvarchar](255) NULL,
	[F2] [nvarchar](255) NULL,
	[F3] [nvarchar](255) NULL,
	[F4] [nvarchar](255) NULL,
	[F5] [nvarchar](255) NULL,
	[F6] [nvarchar](255) NULL,
	[F7] [nvarchar](255) NULL,
	[F8] [nvarchar](255) NULL,
	[F9] [nvarchar](255) NULL,
	[F10] [nvarchar](255) NULL,
	[F11] [nvarchar](255) NULL,
	[F12] [nvarchar](255) NULL,
	[F13] [nvarchar](255) NULL,
	[F14] [nvarchar](255) NULL,
	[XlsFileName] [varchar](255) NULL,
	[SplitFileName] [varchar](255) NULL,
	[Imp_date] [datetime] NOT NULL,
	[rowid] [int] IDENTITY(1,1) NOT NULL,
	[slip_fg] [varchar](1) NOT NULL,
	[Xls_YM] [varchar](7) NULL,
	[B_sp_pdate] [varchar](10) NULL,
	[E_sp_pdate] [varchar](10) NULL,
	[Order_no] [nvarchar](15) NULL,
	[XLS_skno] [nvarchar](4000) NULL,
	[cust_name] [nvarchar](255) NULL,
	[XLS_Price_TAX] [int] NULL,
	[XLS_QTY] [int] NULL,
	[XLS_Price] [numeric](29, 17) NULL,
	[XLS_TOT] [numeric](18, 6) NULL,
	[print_date] [date] NULL,
	[SP_Exec_date] [datetime] NOT NULL
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
