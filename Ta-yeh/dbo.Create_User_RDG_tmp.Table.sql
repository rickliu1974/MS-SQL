USE [DW]
GO
/****** Object:  Table [dbo].[Create_User_RDG_tmp]    Script Date: 07/24/2017 14:43:39 ******/
DROP TABLE [dbo].[Create_User_RDG_tmp]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[Create_User_RDG_tmp](
	[RowID] [int] NULL,
	[Kind] [varchar](1) NULL,
	[Data] [varchar](4000) NULL
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
