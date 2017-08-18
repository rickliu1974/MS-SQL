USE [DW]
GO
/****** Object:  Table [dbo].[Stock_Error_Reason_tmp]    Script Date: 07/24/2017 14:43:53 ******/
DROP TABLE [dbo].[Stock_Error_Reason_tmp]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Stock_Error_Reason_tmp](
	[year] [float] NULL,
	[month] [float] NULL,
	[sk_no] [nvarchar](255) NULL,
	[sk_name] [nvarchar](255) NULL,
	[qty] [float] NULL,
	[error_reason] [nvarchar](255) NULL
) ON [PRIMARY]
GO
