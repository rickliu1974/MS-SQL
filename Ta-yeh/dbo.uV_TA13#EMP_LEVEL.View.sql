USE [DW]
GO
/****** Object:  View [dbo].[uV_TA13#EMP_LEVEL]    Script Date: 08/18/2017 17:43:39 ******/
DROP VIEW [dbo].[uV_TA13#EMP_LEVEL]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
Create View [dbo].[uV_TA13#EMP_LEVEL]
as
  -- 2017/03/17 Rickliu 原本是以 TYEIPDBS2.lytdbta13.SP_EMP_LEVEL Store_Proce 產出 TB_EMP_LEVEL 資料表，因改用訂閱資料庫方式，所以改為 VIEW 取得資料。
  -- 後續思考看能否用 CTE 方式再加速資料處理。
  select *,
         case 
           when dp_no <> Chg_Dept_Level then  ''
           else rtrim(e_no)
         end as Result_e_no, -- 此欄位提供給 SmartQuery 派送參數所用
         case 
           when Chg_Duty_Level >=3 then  ''
           else rtrim(Convert(Varchar(5), Chg_dept_level))
         end as Result_dept_level, -- 此欄位提供給 SmartQuery 派送參數所用
         case 
           when Chg_Duty_Level >=3 then  ''
           else rtrim(Convert(Varchar(5), Chg_duty_level))
         end as Result_duty_level, -- 此欄位提供給 SmartQuery 派送參數所用
         case 
           when Chg_Duty_Level >=3 then  ''
           else rtrim(e_mstno)
         end as Result_e_mstno, -- 此欄位提供給 SmartQuery 派送參數所用
         getdate() as cr_date
    from (Select distinct
                 Rtrim(M.e_no) as e_no, 
                 Rtrim(M.e_name) as e_name, 
                 replace(replace(dp_no, 'Z', '1'), 'B2', 'B1') as Org_Order, --
                 Case
                   When dp_no in ('Z1000', 'Z3000', '11000', '13000') then 1
                   When substring(dp_no, 1, 2) in ('Z2', '12') then 2
                   When substring(dp_no, 2, 1)  = '0'  then 3
                   When substring(dp_no, 1, 3) in ('Z11', 'Z12', 'Z13', '111', '112', '113') or substring(dp_no, 3, 1) ='0' then 4
                   
                   When substring(dp_no, 4, 1) ='0' then 5
                   When substring(dp_no, 5, 1) ='0' then 6
                   else 7
                 end as dp_lv, -- 組織階層編號
                 Rtrim(D1.dp_no) as dp_no, 
                 Substring(Rtrim(D1.dp_no), 1, Len(Rtrim(D1.dp_no))-1) as upper_dp, -- 上層組織
                 Rtrim(D1.dp_name) as dp_name, 
                 Case 
                   When e_no like '%ZZ%' then '虛擬帳號'
                   else Rtrim(M.e_duty)
                 end as e_duty,
                 M.e_mstno, -- 組別編號
                 rtrim(Replace(Convert(Varchar(100),dp_rem),'舊部門編號:','')) as 'old_dpno',
                 M.e_rdate, -- 到職日期
                 M.e_ldate, -- 離職日期
                 case 
                   -- 2015/11/13 Rickliu 此條件請勿更動，因出勤會抓取此資料作為離職判斷依據。
                   when (convert(varchar(10), e_ldate, 111)='1900/01/01') and (D2.status ='1')
                   then 'Y'
                   else 'N'
                 end as status, -- EIP 帳號啟用
                 rtrim(substring(dp_no, 1, 2)) as Chg_Master_Dept,
/**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**
 2015/11/11 Rickliu 職稱等級，L3(含)以上等級可以看到所有資料
    7: 董事長級
    6: 總經理、副總級
    5: 管理部協理
    4: 經理級、客服組長、業務助理、採購組長、業一課長
    3: 總管理處、管理部、財會部
       課長級、組別設定 LS 者
       只有美編企劃組不能看，其他行銷企劃課成員皆可以看
    -----------------------------------------------------------^ 可以看得到所有資料
    2: 組長、主任，非業一課業務課長
    1: 專員或助理
    0: 離職員工
 --**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**/
                 Case --員工職階(作為可查詢資料所需職階，請勿將以下順序調動，判斷式順序是由上往下判斷
                   -- L0
                   When e_ldate <> '1900/01/01' then 0 --離職員工

                   -- L7
                   When e_Duty like '%董事長%' or substring(dp_no, 1, 2) in ('A3') then 7

                   -- L6
                   When e_Duty like '%總經理%' or e_Duty like '%副總%' then 6

                   -- L5
                   When e_Duty like '%管理%協理%' then 5

                   -- L4
                   When e_Duty like '%經理%' Or e_Duty like '%客服組長%' or e_Duty like '%業%助%' or e_Duty like '%採購組長%' then 4 
                   When substring(dp_no, 1, 2) = 'B2' and e_Duty = '%業務課長%' then 4

                   -- L2
                   When e_Duty like '%組長%' or e_Duty like '%主任%' then 2
                   When substring(dp_no, 1, 2) <> 'B2' and e_Duty = '%業務課長%' then 2

                   -- L3
                   When (e_Duty like '%課長%' or e_mstno = 'LS') then 3
                   When (e_Duty like '%警衛%' or e_mstno = 'LS') then 1 
                   When (substring(dp_no, 1, 2) = 'B2' and e_Duty = '%業務課長%') then 3
                   When (substring(dp_no, 1, 2) in ('A0', 'A1', 'A2')) then 3 --總管理處、管理部、財會部
                        --2015/11/11 Rickliu 副總指示 只有美編企劃組不能看，其他行銷企劃課成員皆可以看
                   When (substring(dp_no, 1, 2) = 'A4' And dp_no <> 'A4110') then 3
				        --商開
                   When (substring(dp_no, 1, 2) = 'A5') then 3

                   -- L1
                   else 1 --專員或助理
                 end as Chg_Duty_Level, 

/**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**
 2015/11/11 Rickliu 部門等級
   全資料視野(清空部門編號   )：董事長、總經理、副總級、總管理處、管理部、財會部
   事業群視野(部門編號前 1 碼)：暫訂管理職之協理職稱者
   部級視野  (部門編號前 2 碼)：經理、客服組長、業助、業一課長
   課級視野  (部門編號前 3 碼)：課長級、組別設定 LS 者
   組別視野  (部門編號前 4 碼)：凡離職人員以及上述未列條件者
 --**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**/
                 Case --部門階級(作為資料對應部門欄位)，請勿將以下順序調動，判斷式順序是由上往下判斷
                   -- 最小單位視野
                   When e_ldate <> '1900/01/01' then rtrim(e_dept) --離職員工

                   -- 全資料視野
                   When e_Duty like '%董事長%' then '' 
                   When e_Duty like '%總經理%' or e_Duty like '%副總%' then ''
                   When substring(rtrim(e_dept), 1, 2) in ('A0', 'A1', 'A2', 'A3', 'A5') then ''

                   -- 事業群視野
                   When e_Duty like '%管理%協理%' then substring(rtrim(e_dept), 1, 1)  --協理(可查詢各部門以下資料)

                   -- 部級視野
                   When e_Duty like '%經理%' then substring(rtrim(e_dept), 1, 2) 
                   When e_Duty like '%客服組長%' or e_Duty like '%業%助%' then substring(rtrim(e_dept), 1, 2)
                   When (substring(dp_no, 1, 2) = 'B2' and e_Duty = '%業務課長%') then substring(rtrim(e_dept), 1, 2) 

                   -- 課級視野
                   When (e_Duty like '%課長%' or e_mstno = 'LS') then substring(e_dept, 1, 3) --課長(可查詢自己部門以下的資料)

                   -- 組別視野
                   When (substring(dp_no, 1, 2) <> 'B2' and e_Duty = '%業務課長%') then substring(rtrim(e_dept), 1, 4)
                   When e_Duty like '%組長%' or e_Duty like '%主任%' then substring(rtrim(e_dept), 1, 4) --組長 or 主任(僅可查詢部屬資料)

                   -- 最小單位視野
                   else rtrim(e_dept)
                 end as Chg_Dept_Level, 

/**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**
 2015/11/11 Rickliu 成本等級
                  |       |
    成本 < 大盤 < | 中 盤 | < 小盤 < 一般 < 市價
    (     2      )|(  1  )|(         0         ) ==> 成本等級
    L2：董事長、總經理、副總級、財會部、資訊部、行企部(不含美編企劃)、採購部
    L1：總管理處、管理部、經理級、課長級、組長級、業助級、主任級、組別設定 LS 者
    L0：凡離職人員以及上述未列條件者
 --**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**/
                 
                 Case --倉庫成本顯示權限，請勿將以下順序調動，判斷式順序是由上往下判斷
                   -- L0
                   When e_ldate <> '1900/01/01' then 0  --離職員工

                   -- L2
                   When e_Duty like '%董事長%' then 2
                   When e_Duty like '%總經理%' or e_Duty like '%副總%' then 2
                   When dp_no like 'A[2345]000%' then 2
                   When dp_no like 'A[235]%' then 2
                   -- 2015/11/11 Rickliu 副總指示 銷除了美編企劃組不能看其他行銷人員允許查看
                   When (dp_no like 'A4%') And (dp_no not like 'A411%') then 2 
				   -- 商開
                   When (dp_no like 'A5%')  then 2 
                   
                   -- L1
                   When e_Duty like '%管理%協理%' then 1 
                   When e_Duty like '%經理%' or e_Duty like '%課長%' or e_Duty like '%組長%' then 1
                   When e_Duty like '%業%助%' or e_Duty like '%主任%' or e_mstno = 'LS' then 1

                   -- L0
                   else 0 --無成本欄位權限
                 end as stock_amt_level, 
          
                 Rtrim(D3.PassWord) as EIP_Password, -- EIP 密碼
                 Rtrim(D2.EMail) as EMail, -- 大業個人 MAIL
                 Rtrim(D2.MSNAccount) as MSNAccount, 
                 Rtrim(D2.SKYPEAccount) as SKYPEAccount, 
                 case
                   when (convert(varchar(10), e_ldate, 111) >= '2013/10/01') Or
                        (convert(varchar(10), e_ldate, 111) = '1900/01/01')
                   then Substring(D2.AccountID, 1, 2) 
                   else ''
                 end as sDept, --
                 Rtrim(D2.AccountID) as AccountID,
                 Case 
                   When rtrim(D2.OfficeExt) = '' then '000'
                   else rtrim(D2.OfficeExt)
                 end as OfficeExt, -- 分機號
                 Case 
                   When rtrim(D2.MsnAccount) = '' then '000'
                   else rtrim(D2.MsnAccount)
                 end as PC_Name, -- 預設電腦名稱
                 case 
                   when convert(varchar(10), e_ldate, 111)='1900/01/01' 
                   then 'N' -- 在職
                   else 'Y' -- 離職
                 end as 'leave', -- 離職否
                 Rtrim(D2.Remark) as Remark
            From SYNC_TA13.dbo.pemploy as M 
                 Inner join SYNC_TA13.dbo.pdept as D1 
                    on M.e_dept = D1.dp_no 
                  -- 2014/03/20 請勿使用 Inner join  因為 EIP 可能無此帳號，但 SQ 仍須有此員工資料才能調閱出所屬員報表。
                  left join WebEIP2.dbo.Account_Data_View as D2 
                    on m.e_no = D2.pager Collate Chinese_Taiwan_Stroke_CI_AS
                   and M.e_name Collate Chinese_Taiwan_Stroke_CI_AS = D2.FullName
                  Left outer join WebEIP2.dbo.AFS_AccountView as D3
                    on D2.AccountID = D3.AccountID
           Where (M.e_ldate >= '1900/1/1') 
             And (M.e_no <> 'LY') 
             And (substring(M.e_no, 1, 3) <> 'TAZ')
         ) m
GO
