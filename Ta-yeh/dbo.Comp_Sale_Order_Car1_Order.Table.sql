USE [DW]
GO
/****** Object:  Table [dbo].[Comp_Sale_Order_Car1_Order]    Script Date: 07/24/2017 14:43:38 ******/
DROP TABLE [dbo].[Comp_Sale_Order_Car1_Order]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[Comp_Sale_Order_Car1_Order](
	[rowid] [int] NOT NULL,
	[ct_no] [nchar](10) NULL,
	[ct_sname] [nvarchar](12) NULL,
	[ct_ssname] [nvarchar](73) NULL,
	[sk_no] [nchar](30) NULL,
	[sk_name] [nchar](60) NULL,
	[f5] [nvarchar](255) NULL,
	[sk_bcode] [nchar](30) NULL,
	[f4] [nvarchar](255) NULL,
	[Ori_sale_qty] [numeric](20, 6) NULL,
	[sale_qty] [numeric](20, 6) NULL,
	[Ori_sale_amt] [numeric](38, 18) NULL,
	[sale_amt] [numeric](38, 18) NULL,
	[Ori_sale_tot] [numeric](20, 6) NULL,
	[sale_tot] [numeric](20, 6) NULL,
	[F11] [varchar](100) NULL,
	[F12] [varchar](10) NOT NULL,
	[F13] [nvarchar](255) NULL,
	[isfound] [varchar](1) NOT NULL,
	[sct_ss_csno] [nchar](20) NOT NULL,
	[sct_ss_csname] [nchar](80) NOT NULL,
	[sct_ss_rec] [numeric](20, 6) NOT NULL,
	[sct_ss_ctkind] [nchar](1) NOT NULL,
	[sct_ss_ctno] [ntext] NOT NULL,
	[sct_ss_nokind] [nchar](1) NOT NULL,
	[sct_ss_no] [nchar](30) NOT NULL,
	[sct_ss_sendno] [nchar](30) NOT NULL,
	[sct_ss_noqty] [numeric](20, 6) NOT NULL,
	[sct_ss_sendqty] [numeric](20, 6) NOT NULL,
	[sct_ss_oneqty] [numeric](20, 6) NOT NULL,
	[sct_ss_itmqty] [numeric](20, 6) NOT NULL,
	[sct_sendtot] [numeric](20, 6) NOT NULL,
	[sct_ss_sdate] [datetime] NOT NULL,
	[sct_ss_edate] [datetime] NOT NULL,
	[sct_ss_sqty] [numeric](20, 6) NOT NULL,
	[sct_ss_eqty] [numeric](20, 6) NOT NULL,
	[sct_ss_form] [ntext] NOT NULL,
	[isSend] [int] NOT NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
