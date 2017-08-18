USE [DW]
GO
/****** Object:  Table [dbo].[Sale_Member_BKind_Rank_OT_By_YM_tmp]    Script Date: 07/24/2017 14:43:52 ******/
DROP TABLE [dbo].[Sale_Member_BKind_Rank_OT_By_YM_tmp]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[Sale_Member_BKind_Rank_OT_By_YM_tmp](
	[Rank] [bigint] NULL,
	[year] [int] NOT NULL,
	[month] [int] NOT NULL,
	[emp_no] [varchar](20) NULL,
	[emp_name] [varchar](50) NULL,
	[sale_amt] [float] NULL,
	[target_amt] [float] NULL,
	[rate] [float] NULL
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
