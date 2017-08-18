USE [DW]
GO
/****** Object:  Table [dbo].[Target_Sales_Cust_Bkind_Amt_tmp]    Script Date: 07/24/2017 14:43:53 ******/
DROP TABLE [dbo].[Target_Sales_Cust_Bkind_Amt_tmp]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[Target_Sales_Cust_Bkind_Amt_tmp](
	[year] [varchar](4) NOT NULL,
	[month] [varchar](2) NOT NULL,
	[ct_sales] [varchar](20) NOT NULL,
	[sales_name] [varchar](255) NULL,
	[ct_no] [varchar](20) NOT NULL,
	[ct_name] [varchar](255) NULL,
	[Bkind] [varchar](2) NOT NULL,
	[Amt] [int] NULL,
	[imp_date] [datetime] NULL
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
