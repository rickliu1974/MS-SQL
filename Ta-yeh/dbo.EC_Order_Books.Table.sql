USE [DW]
GO
/****** Object:  Table [dbo].[EC_Order_Books]    Script Date: 07/24/2017 14:43:41 ******/
DROP TABLE [dbo].[EC_Order_Books]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[EC_Order_Books](
	[rowid] [int] NOT NULL,
	[ct_no] [nvarchar](10) NULL,
	[ct_sname] [nvarchar](12) NULL,
	[ct_ssname] [nvarchar](73) NULL,
	[sk_no] [nchar](30) NULL,
	[sk_name] [nchar](60) NULL,
	[f5] [nvarchar](255) NULL,
	[sk_bcode] [nchar](30) NULL,
	[f4] [varchar](max) NULL,
	[Ori_sale_qty] [numeric](20, 6) NULL,
	[sale_qty] [numeric](20, 6) NULL,
	[Ori_sale_amt] [numeric](20, 6) NULL,
	[sale_amt] [numeric](20, 6) NULL,
	[Ori_sale_tot] [numeric](38, 9) NULL,
	[sale_tot] [numeric](38, 9) NULL,
	[F11] [varchar](100) NULL,
	[F12] [varchar](10) NOT NULL,
	[F13] [varchar](10) NOT NULL,
	[source_file] [varchar](255) NULL,
	[source_order] [varchar](max) NULL,
	[buyer] [nvarchar](255) NULL,
	[cust] [nvarchar](255) NULL,
	[cust_phone] [nvarchar](255) NULL,
	[cust_mobile] [nvarchar](255) NULL,
	[cust_addr1] [nvarchar](255) NULL,
	[cust_addr2] [nvarchar](255) NULL,
	[MEMO] [nvarchar](255) NULL,
	[isfound] [varchar](1) NOT NULL,
	[sct_ss_csno] [nchar](20) NOT NULL,
	[sct_ss_csname] [nchar](80) NOT NULL,
	[sct_ss_rec] [decimal](20, 6) NOT NULL,
	[sct_ss_ctkind] [nchar](1) NOT NULL,
	[sct_ss_ctno] [ntext] NOT NULL,
	[sct_ss_nokind] [nchar](1) NOT NULL,
	[sct_ss_no] [nchar](30) NOT NULL,
	[sct_ss_sendno] [nchar](30) NOT NULL,
	[sct_ss_noqty] [decimal](20, 6) NOT NULL,
	[sct_ss_sendqty] [decimal](20, 6) NOT NULL,
	[sct_ss_oneqty] [decimal](20, 6) NOT NULL,
	[sct_ss_itmqty] [decimal](20, 6) NOT NULL,
	[sct_sendtot] [decimal](20, 6) NOT NULL,
	[sct_ss_sdate] [datetime] NOT NULL,
	[sct_ss_edate] [datetime] NOT NULL,
	[sct_ss_sqty] [decimal](20, 6) NOT NULL,
	[sct_ss_eqty] [decimal](20, 6) NOT NULL,
	[sct_ss_form] [ntext] NOT NULL,
	[isSend] [nvarchar](255) NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
