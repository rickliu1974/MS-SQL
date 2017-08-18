USE [DW]
GO
/****** Object:  Table [dbo].[Zip_Address]    Script Date: 07/24/2017 14:43:54 ******/
DROP TABLE [dbo].[Zip_Address]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[Zip_Address](
	[Zip_Code] [varchar](5) NULL,
	[City] [varchar](50) NULL,
	[Area] [varchar](50) NULL,
	[Road] [varchar](50) NULL,
	[Scope] [varchar](50) NULL,
	[Road_Addr] [varchar](100) NULL,
	[Old_City] [varchar](100) NULL,
	[Old_Area] [varchar](100) NULL
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
