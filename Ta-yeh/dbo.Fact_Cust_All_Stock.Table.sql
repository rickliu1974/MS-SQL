USE [DW]
GO
/****** Object:  Table [dbo].[Fact_Cust_All_Stock]    Script Date: 07/24/2017 14:43:43 ******/
DROP TABLE [dbo].[Fact_Cust_All_Stock]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Fact_Cust_All_Stock](
	[ct_no] [nvarchar](10) NULL,
	[ct_sname] [nvarchar](12) NULL,
	[ct_name] [nvarchar](60) NULL,
	[ct_fld3] [nvarchar](60) NULL,
	[sk_no] [nvarchar](30) NULL,
	[sk_name] [nvarchar](60) NULL
) ON [PRIMARY]
GO
