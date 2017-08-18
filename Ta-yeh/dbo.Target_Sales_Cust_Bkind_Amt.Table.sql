USE [DW]
GO
/****** Object:  Table [dbo].[Target_Sales_Cust_Bkind_Amt]    Script Date: 07/24/2017 14:43:53 ******/
DROP TABLE [dbo].[Target_Sales_Cust_Bkind_Amt]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[Target_Sales_Cust_Bkind_Amt](
	[year] [varchar](4) NOT NULL,
	[month] [varchar](2) NOT NULL,
	[ct_sales] [varchar](20) NOT NULL,
	[sales_name] [varchar](255) NULL,
	[ct_no] [varchar](20) NOT NULL,
	[ct_name] [varchar](255) NULL,
	[Bkind] [varchar](2) NOT NULL,
	[Amt] [int] NULL,
	[imp_date] [datetime] NULL,
 CONSTRAINT [PK_Target_Sales_Cust_Bkind_Amt] PRIMARY KEY CLUSTERED 
(
	[year] ASC,
	[month] ASC,
	[ct_sales] ASC,
	[ct_no] ASC,
	[Bkind] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
