USE [DW]
GO
/****** Object:  Table [dbo].[Mall_Consign_Order_Momo_tmp]    Script Date: 07/24/2017 14:43:44 ******/
DROP TABLE [dbo].[Mall_Consign_Order_Momo_tmp]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[Mall_Consign_Order_Momo_tmp](
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
	[F15] [nvarchar](255) NULL,
	[F16] [nvarchar](255) NULL,
	[F17] [nvarchar](255) NULL,
	[F18] [nvarchar](255) NULL,
	[F19] [nvarchar](255) NULL,
	[F20] [nvarchar](255) NULL,
	[F21] [nvarchar](255) NULL,
	[XlsFileName] [varchar](255) NULL,
	[SplitFileName] [varchar](255) NULL,
	[Imp_date] [datetime] NOT NULL,
	[rowid] [int] IDENTITY(1,1) NOT NULL,
	[slip_fg] [varchar](1) NOT NULL,
	[Order_no] [nvarchar](15) NULL,
	[sp_pdate] [varchar](7) NULL,
	[XLS_CT_Name] [nvarchar](255) NULL,
	[XLS_SKNO] [nvarchar](510) NULL,
	[XLS_Qty] [int] NULL,
	[XLS_Amt] [int] NULL,
	[XLS_Tot] [int] NULL,
	[YM] [varchar](7) NOT NULL,
	[B_Pdate] [varchar](10) NOT NULL,
	[E_Pdate] [varchar](10) NOT NULL,
	[SP_Exec_date] [datetime] NOT NULL
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
