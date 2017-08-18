USE [DW]
GO
/****** Object:  Table [dbo].[RealTime_VideoPath]    Script Date: 07/24/2017 14:43:52 ******/
DROP TABLE [dbo].[RealTime_VideoPath]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[RealTime_VideoPath](
	[Video_Name] [varchar](30) NOT NULL,
	[Video_Time] [int] NULL,
	[FilePath] [varchar](200) NULL
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
