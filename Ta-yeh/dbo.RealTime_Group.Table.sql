USE [DW]
GO
/****** Object:  Table [dbo].[RealTime_Group]    Script Date: 07/24/2017 14:43:51 ******/
DROP TABLE [dbo].[RealTime_Group]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[RealTime_Group](
	[Group_ID] [varchar](20) NULL,
	[Group_Type] [varchar](20) NULL,
	[SubSys] [varchar](20) NULL,
	[Query] [varchar](50) NULL,
	[Query_Name] [varchar](200) NULL,
	[Timer] [int] NULL,
	[Query_Seq] [int] NULL
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
