USE [DW]
GO
/****** Object:  Table [dbo].[Ori_xls#Target_Cust_For_LAN]    Script Date: 07/24/2017 14:43:51 ******/
DROP TABLE [dbo].[Ori_xls#Target_Cust_For_LAN]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[Ori_xls#Target_Cust_For_LAN](
	[Year] [float] NULL,
	[Month] [float] NULL,
	[ct_no] [nvarchar](255) NULL,
	[ct_name] [nvarchar](255) NULL,
	[amt] [float] NULL,
	[xlsFileName] [varchar](255) NULL,
	[SplitFileName] [varchar](255) NULL,
	[imp_date] [datetime] NOT NULL,
	[rowid] [int] IDENTITY(1,1) NOT NULL
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
