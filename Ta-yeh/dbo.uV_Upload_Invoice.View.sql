USE [DW]
GO
/****** Object:  View [dbo].[uV_Upload_Invoice]    Script Date: 08/18/2017 17:43:39 ******/
DROP VIEW [dbo].[uV_Upload_Invoice]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
Create View [dbo].[uV_Upload_Invoice]
as
  select '###_Invoice_Master_###' as '###_Invoice_Master_###',
         Convert(Varchar(10), isnull(in_no, '')) as M_in_no, -- 發票號碼
         Convert(Varchar(10), in_date, 111) as M_in_date, -- 發票日期
         -- 2015/01/29 Rickliu 由於凌越系統並無電子發票選項，又加上 RPT_13010 報表一率都是用電子發票模式上傳，所以固定寫 07
         /*
         case
           when in_frm = '31' then '01' -- 三聯式發票
           when in_frm = '32' then '02' -- 二聯式發票
           when in_frm = '35' then '03' -- 三聯式收銀機發票
           else ''
         end as M_in_frm, -- 發票類別
         */
         '07' as M_in_frm, -- 發票類別
         Convert(Varchar(10), isnull(in_bno, '')) as M_in_bno, -- 買方統一編號
         convert(char(1), in_tcd) as M_in_tcd, -- 課稅別 1:應稅, 2:零稅率, 3:免稅
         Convert(Varchar(30),
           case
             when isnull(in_amt, 0) = 0 then 0
             else Convert(Float, Round(in_tax / in_amt, 2) * 100)
           end) as M_in_tax, -- 稅率
         '' as M_in_bound_fg, -- 通關方式註記 1: 非經海關出口, 2:經海關出口
         '' as M_bay_role_fg, -- 營業人角色註記
         Convert(Varchar(30), Convert(Money, isnull(in_amt, 0))) as M_in_amt, -- 原幣金額
         '1' as M_Exchange_rate, -- 匯率
         'TWD' as M_Currency, -- 幣別
         -- 彙開註記：若有彙開註記，則打*，2014/11/04 素如有致電去國稅局，表示一般營業人是用不到的，所以可以空白
         -- Convert(Varchar(1), Replace(Rtrim(Isnull(in_scd, '')), 'A', '*')) as M_in_scd, -- 彙開註記 *:彙開
         '' as M_in_scd, -- 彙開註記
         '0' as M_sale_kind, -- 銷售類別 0:一般銷售 1:洋菸酒類 2:固定資產 3:土地
         '' as M_remark, -- 註記
         Convert(Varchar(20), m.in_ctno) as M_in_ctno, -- 客戶編號
/*
         '###_Invoice_Detail_###' as '###_Invoice_Detail_###',
         Convert(Varchar(10), isnull(in_no, '')) as D_in_no, -- 發票號碼
         Convert(Varchar(20), Rtrim(Isnull(id_sknm, ''))) as D_id_sknm, -- 貨品編號
         Convert(Varchar(20), Replace(Replace(Replace(Replace(Replace(Convert(Varchar(100), Rtrim(Isnull(id_skno, ''))), '<', ''), '>', ''), '%', ''), '&', ''), '!', '')) as D_id_skno, -- 貨品名稱
         '' as D_Related_number, -- 相關號碼
  
         Convert(Varchar(30), Convert(Money, isnull(id_price, 0))) as D_id_price1, -- 單價1
         Convert(Varchar(10), Rtrim(isnull(id_unit, ''))) as D_id_unit1, -- 單位1
         Convert(Varchar(30), Convert(Money, isnull(id_qty, 0))) as D_id_qty1, -- 數量1
  
         '' as D_id_price2, -- 單價2
         '' as D_id_unit2, -- 單位2
         '' as D_id_qty2, -- 數量2
         
         '' as D_remark, -- 單一欄位註記
*/
         /**************************************************************************************************************************************
          2014/12/30 Rickliu 林課表示增加營業人銷項上傳國稅局資料格式檔
          Upload Bessiness Sale Invoice Data ==> BSI          
          
          Field NO		Field Name				Size	Begin	End
          Field(01)		格式代號				X(002)		1	  2
          Field(02)		申報營業人稅籍編號		X(009)	    3	 11
          Field(03)		流水號					X(007)	   12	 18
          Field(04)		資料所屬年月			9(005)	   19    23
          Field(05)		買受人統一編號			X(008)	   24	 31
          
         **************************************************************************************************************************************/
         '### 營業稅銷項資料格式檔 ###' as '### 營業稅銷項資料格式檔 ###',
         -- Field(01) 格式代號
         in_frm as UBI_Frm, 
         -- Field(02) 保留給 申報營業人稅籍編號 Tax Registration Number
         '#AAAAAAA#' as UBI_Trn, 
         -- Field(03) 流水號
         '#BBBBB#' as UBI_NO, -- 流水號
         -- Field(04-01) 資料所屬年
--         Convert(Varchar(4), year(in_date)-1911) Collate Chinese_PRC_BIN as UBI_cYear, -- 發票所屬年(民國年), 請勿更改定序
         -- Field(04-02) 資料所屬月
--         Convert(Varchar(2), month(in_date)) Collate Chinese_PRC_BIN as UBI_Month, -- 發票所屬月
         Convert(Varchar(4), Convert(Int, Substring(in_yrmn, 1, 4)) -1911) Collate Chinese_PRC_BIN as UBI_cYear, -- 發票所屬年(民國年), 請勿更改定序
         -- Field(04-02) 資料所屬月
         Substring(in_yrmn, 5, 2) Collate Chinese_PRC_BIN as UBI_Month, -- 發票所屬月
         
         -- Field(05) 買受人統編
         Convert(Char(8), Convert(Varchar(10), isnull(in_bno, ''))) as UBI_Buy_GUI,
         -- Field(06) 保留給 銷售人統編 Government Uniform Invoice(GUI) Number ==> GUI Number
         '#CCCCCC#' as UBI_Sale_GUI,
         -- Field(07) 統一發票 GUI
         Reverse(Convert(Char(10), Reverse(RTrim(Convert(Varchar(10), isnull(in_no, '')))))) as UBI_GUI,
         -- Field(13) 銷售金額 
         Substring(Convert(Varchar(30), Convert(Money, isnull(in_amt, 0) + 1000000000000)), 2, 12) as UBI_Sale_AMT,
         -- Field(14) 課稅別
         convert(char(1), in_tcd) as UBI_Tcd,
         -- Field(15) 營業稅稅額
         Substring(Convert(Varchar(30), Convert(Money, isnull(in_tax, 0) + 10000000000)), 2, 10) as UBI_Tax,
         -- Field(16) 扣抵代號
         convert(char(1), in_mcd) as UBI_Mcd,
         -- Field(17) 空白
         Space(5) as UBI_Keep1,
         -- Field(18) 特種稅額稅率
         Space(1) as UBI_Keep2,
         -- Field(19) 彙加註記
         convert(char(1), in_scd) as UBI_Scd,
         -- Field(20) 通關方式註記
         Space(1) as UBI_Keep3,
         case
           when m.in_frm in ('31', '32', '33', '34', '35')
           then Convert(Char(2), in_frm) + -- 格式代號
                '#AAAAAAA#'+ -- 保留給 申報營業人稅籍編號
                '#BBBBB#'+ -- 流水號
                Convert(Char(3), Convert(Varchar(4), substring(in_yrmn, 1, 4)-1911) Collate Chinese_PRC_BIN) + -- 資料所屬年 民國年
                Substring(Convert(Char(3), Convert(Varchar(2), substring(in_yrmn, 5, 2)) Collate Chinese_PRC_BIN +100), 2, 2)+ -- 資料所屬月份
                Convert(Char(8), Convert(Varchar(10), isnull(in_bno, '')))+ -- 買受人統編         
                '#CCCCCC#'+ -- 保留給 銷售人統編
                Reverse(Convert(Char(10), Reverse(RTrim(Convert(Varchar(10), isnull(in_no, ''))))))+ -- 發票號碼
                Substring(Convert(Varchar(30), Convert(Money, isnull(in_amt, 0) + 1000000000000)), 2, 12)+ -- 銷售金額
                convert(char(1), in_tcd)+ -- 課稅別
                Substring(Convert(Varchar(30), Convert(Money, isnull(in_tax, 0) + 10000000000)), 2, 10)+ -- 營業稅稅額
                convert(char(1), in_mcd)+ -- 
                Space(5)+
                Space(1)+
                convert(char(1), in_scd)+
                Space(1)
                else ''
         end as Business_Sale_Invoice_OutData
    from tyeipdbs2.lytdbta13.dbo.pinvo m
--          left join tyeipdbs1.lytdbta13.dbo.pinvodt d
--            on m.in_prono = d.id_prono
   where 1=1
     --and m.in_tcd <> 'F'
     and m.in_no <> ''
     and m.in_frm in ('31', '32', '33', '34', '35')


/***[發票主檔]*****************************************************************************************************************************************

發票號碼：含發票字軌共10碼

發票日期：請輸入西元年

發票類別：
01：三聯式發票
02：二聯式發票
03：二聯式收銀機發票
04：特種稅額發票
05：電子計算機發票
06：三聯式收銀機發票

課稅類別代碼：
1：應稅
2：零稅率
3：免稅

匯率：免填%字樣（僅輸入數字）

通關方式註記：
1：非經海關出口
2：經海關出口

原幣金額：(整數12＋小數4)

匯率：(整數8＋小數4)

幣別：
TWD：新台幣
USD：美金
GBP：英鎊
DEM：德國馬克
AUD：澳大利亞幣
HKD：港幣
SGD：新加坡幣
CAD：加拿大幣
CHF：瑞士法郎
MYR：馬來西亞幣
FRF：法國法郎
BEF：比利時法郎
SEK：瑞典幣
JPY：日圓
ITL：義大利里拉
THB：泰銖
NTD：CURRENCY_NTD
EUR：歐洲共同貨幣
NZD：紐西蘭幣

彙開註記：若有彙開註記，則打*，2014/11/04 素如有致電去國稅局，表示一般營業人是用不到的，所以可以空白

銷售類別代碼：
0：一般銷售
1：洋煙酒類
2：固定資產
3：土地

-- [發票明細檔]**********************************************************************************************************************************
發票號碼：須與發票主檔號碼相同 長度 10
品名編號：只能英數 長度 20
發票品名：長度 256
相關號碼：長度 20
單價：長度 17
單位：長度 6
數量：長度 17
單價2：長度 17
單位2：長度 6
數量2：長度 17
單一欄位備註：長度 40

--**********************************************************************************************************************************************

備註：
1.發票主檔及明細檔為避免科學記號問題產生，請採用文字格式
2.匯入發票號碼最多不可超過12張發票，每張發票明細檔限制999筆
3.主檔發票號碼與明細檔發票號碼相同視為同一張發票
4. 課稅別為零稅率時，才需填寫通關方式註記、原幣金額、匯率及幣別
5.發票主檔須以發票號碼排序，且上一張發票之開立日期不得大於下一張的發票日期，以避免跳號之問題
6.請勿使用特殊字元（例如：<, >, %,&,!）
7.單價2及數量2僅供鋼鐵發票使用
8.幣別及匯率欄位僅供備註使用，不具運算功能
9.發票明細檔中單價及單價2欄位請以新台幣輸入
*************************************************************************************************************************************************/
GO
