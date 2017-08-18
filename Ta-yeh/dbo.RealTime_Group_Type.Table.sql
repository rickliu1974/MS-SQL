USE [DW]
GO
/****** Object:  Table [dbo].[RealTime_Group_Type]    Script Date: 07/24/2017 14:43:51 ******/
DROP TABLE [dbo].[RealTime_Group_Type]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[RealTime_Group_Type](
	[Group_Type] [varchar](20) NULL,
	[Type_Name] [varchar](50) NULL,
	[Type_Seq] [int] NULL
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
