USE [DW]
GO
/****** Object:  Table [dbo].[Ori_xls#Dog_RecordLists]    Script Date: 07/24/2017 14:43:47 ******/
DROP TABLE [dbo].[Ori_xls#Dog_RecordLists]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Ori_xls#Dog_RecordLists](
	[日期] [datetime] NULL,
	[業務] [nvarchar](255) NULL,
	[周行程] [nvarchar](255) NULL,
	[洽談客戶名稱] [nvarchar](255) NULL,
	[工作日報] [nvarchar](255) NULL,
	[訪談目的] [nvarchar](255) NULL,
	[開始時間] [nvarchar](255) NULL,
	[結束時間] [nvarchar](255) NULL,
	[拜訪時間] [datetime] NULL,
	[衛星犬計畫] [nvarchar](255) NULL,
	[衛星犬] [nvarchar](255) NULL,
	[停留時間] [datetime] NULL,
	[rowid] [int] IDENTITY(1,1) NOT NULL
) ON [PRIMARY]
GO
