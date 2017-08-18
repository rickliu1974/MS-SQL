USE [DW]
GO
/****** Object:  Table [dbo].[Stock_taking_TA13]    Script Date: 07/24/2017 14:43:53 ******/
DROP TABLE [dbo].[Stock_taking_TA13]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[Stock_taking_TA13](
	[F1] [nvarchar](255) NULL,
	[F2] [nvarchar](255) NULL,
	[F3] [nvarchar](255) NULL,
	[F4] [nvarchar](255) NULL,
	[SP_date] [varchar](10) NOT NULL,
	[SP_sales] [varchar](6) NOT NULL,
	[SP_dpno] [varchar](5) NOT NULL,
	[import_date] [date] NULL
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
