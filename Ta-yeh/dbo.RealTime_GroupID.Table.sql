USE [DW]
GO
/****** Object:  Table [dbo].[RealTime_GroupID]    Script Date: 07/24/2017 14:43:51 ******/
DROP TABLE [dbo].[RealTime_GroupID]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[RealTime_GroupID](
	[Group_ID] [varchar](20) NULL,
	[Use_Type] [varchar](10) NULL,
	[Group_Type] [varchar](20) NULL,
	[Group_Seq] [int] NULL
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
