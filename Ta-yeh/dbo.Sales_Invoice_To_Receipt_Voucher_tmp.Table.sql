USE [DW]
GO
/****** Object:  Table [dbo].[Sales_Invoice_To_Receipt_Voucher_tmp]    Script Date: 07/24/2017 14:43:52 ******/
DROP TABLE [dbo].[Sales_Invoice_To_Receipt_Voucher_tmp]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[Sales_Invoice_To_Receipt_Voucher_tmp](
	[sd_date] [nvarchar](255) NULL,
	[sp_no] [nvarchar](20) NULL,
	[sd_dcr] [nvarchar](284) NULL,
	[sd_atno] [varchar](4) NOT NULL,
	[sd_ctno] [nvarchar](255) NOT NULL,
	[ct_sname] [nchar](12) NOT NULL,
	[ct_credit] [nvarchar](10) NOT NULL,
	[sd_doc] [varchar](1) NOT NULL,
	[sd_amt] [float] NULL,
	[sd_oamt] [float] NULL,
	[sp_mkman] [nvarchar](255) NULL,
	[e_name] [nchar](10) NULL,
	[sd_seq] [bigint] NULL,
	[sp_memo] [varchar](51) NULL
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
