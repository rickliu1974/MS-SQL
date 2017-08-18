USE [DW]
GO
/****** Object:  Table [dbo].[Mall_Normal_Order_Carno1_Order]    Script Date: 07/24/2017 14:43:45 ******/
DROP TABLE [dbo].[Mall_Normal_Order_Carno1_Order]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[Mall_Normal_Order_Carno1_Order](
	[rowid] [int] NOT NULL,
	[ct_no] [nchar](10) NULL,
	[ct_sname] [nvarchar](12) NULL,
	[ct_ssname] [nvarchar](73) NULL,
	[sk_no] [nchar](30) NULL,
	[sk_name] [nchar](60) NULL,
	[f5] [varchar](100) NULL,
	[sk_bcode] [nchar](30) NULL,
	[f4] [varchar](100) NULL,
	[Ori_sale_qty] [numeric](20, 6) NULL,
	[sale_qty] [numeric](20, 6) NULL,
	[Ori_sale_amt] [numeric](38, 18) NULL,
	[sale_amt] [numeric](38, 18) NULL,
	[Ori_sale_tot] [numeric](20, 6) NULL,
	[sale_tot] [numeric](20, 6) NULL,
	[F11] [varchar](100) NULL,
	[F12] [varchar](100) NULL,
	[F13] [varchar](100) NULL,
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
	[isSend] [int] NOT NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
