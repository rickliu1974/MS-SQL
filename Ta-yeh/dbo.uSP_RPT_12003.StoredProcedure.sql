USE [DW]
GO
/****** Object:  StoredProcedure [dbo].[uSP_RPT_12003]    Script Date: 08/18/2017 17:43:41 ******/
DROP PROCEDURE [dbo].[uSP_RPT_12003]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE   PROCEDURE [dbo].[uSP_RPT_12003]
AS
BEGIN
/*
20140109 brian 
1.不包含庫吉人員 2位
2.不包含CZZ002彭紹強帳號 1位


==============================================================================================
-- 自訂一個 DATETIME 型別的變數  
DECLARE @myDate AS DATETIME  
  
-- 以 2008 年 3 月 19 日為基準點，找出 2008 年第一天的日期  
-- 注意：改以「年」為單位  
SET @myDate = DATEADD(yy, DATEDIFF(yy, '', '2008/3/19'), '')  
  
-- 以「日」為單位，扣掉 1 天  
SELECT DATEADD(day, -1, @myDate) [2008年3月19日的前一年最後一天(精確到毫秒)]  
================================================================================================
*/

  --目前各部門人數
  select M.deptname,count(M.deptname) as 'dept_people'  
         into #tmp_people
    from (select Org_Order,e_rdate,
                 case 
                     When Org_Order like '110%'  then '董事長室' 
                     When Org_Order like '120%'  then '總經理室'
                     When Org_Order like '130%'  then '執行副總'
                     When Org_Order like 'A00%'  then '總管理處'
                     When Org_Order like 'A11%'  then '管理部'
                     When Org_Order like 'A21%'  then '財會部'
                     When Org_Order like 'A31%'  then '資訊部'
                     When Org_Order like 'A41%'  then '行企部'
                     When Org_Order like 'A51%'  then '採購部'
                     When Org_Order like 'A61%'  then '資材部'
                     When Org_Order like 'B11%' and e_no like 'T%' then '汽車精品課'
                     When Org_Order like 'B12%'  then '家庭五金課'
                     When Org_Order like 'B13%'  then '電子商務課'
                     When Org_Order like 'B16%'  then '國際貿易課'
                     When Org_Order like 'B19%'  then '客服部'
                     When Org_Order like 'D11%'  then '生產部'
                     When Org_Order like 'D21%'  then '品研部'
                     When Org_Order like 'E10%'  then '廣州-庫吉'
                     else  '其他'
                 end as deptname
            from uV_TA13#EMP_LEVEL
           Where leave='N' and e_no like 'T%'  
         )  M 
   group by M.deptname
    
     
    
   --上月底人數
   select deptname,count(1) as pre_m_people 
          into #tmp_lastm_people
     from (select Org_Order,e_rdate,
                  case 
                    When Org_Order like '110%'  then '董事長室' 
                    When Org_Order like '120%'  then '總經理室'
                    When Org_Order like '130%'  then '執行副總'
                    When Org_Order like 'A00%'  then '總管理處'
                    When Org_Order like 'A11%'  then '管理部'
                    When Org_Order like 'A21%'  then '財會部'
                    When Org_Order like 'A31%'  then '資訊部'
                    When Org_Order like 'A41%'  then '行企部'
                    When Org_Order like 'A51%'  then '採購部'
                    When Org_Order like 'A61%'  then '資材部'
                    When Org_Order like 'B11%' and e_no like 'T%' then '汽車精品課'
                    When Org_Order like 'B12%'  then '家庭五金課'
                    When Org_Order like 'B13%'  then '電子商務課'
                    When Org_Order like 'B16%'  then '國際貿易課'
                    When Org_Order like 'B19%'  then '客服部'
                    When Org_Order like 'D11%'  then '生產部'
                    When Org_Order like 'D21%'  then '品研部'
                    When Org_Order like 'E10%'  then '廣州-庫吉'
                    else  '其他'
                  end as deptname
             from uV_TA13#EMP_LEVEL
             --上月底人數  離職=Y 及 離職日>=本月月初  +  未離職 
            Where (leave='Y' and e_ldate >= dateadd(mm,datediff(mm,'',getdate()),'') ) 
               or leave='N' and e_no like 'T%'
          ) Y
    group by deptname 
    
    --新進人員
    select deptname,count(1) as new_people 
           into #tmp_new_people
      from (select *,
                   case 
                     When Org_Order like '110%'  then '董事長室' 
                     When Org_Order like '120%'  then '總經理室'
                     When Org_Order like '130%'  then '執行副總'
                     When Org_Order like 'A00%'  then '總管理處'
                     When Org_Order like 'A11%'  then '管理部'
                     When Org_Order like 'A21%'  then '財會部'
                     When Org_Order like 'A31%'  then '資訊部'
                     When Org_Order like 'A41%'  then '行企部'
                     When Org_Order like 'A51%'  then '採購部'
                     When Org_Order like 'A61%'  then '資材部'
                     When Org_Order like 'B11%' and e_no like 'T%' then '汽車精品課'
                     When Org_Order like 'B12%'  then '家庭五金課'
                     When Org_Order like 'B13%'  then '電子商務課'
                     When Org_Order like 'B16%'  then '國際貿易課'
                     When Org_Order like 'B19%'  then '客服部'
                     When Org_Order like 'D11%'  then '生產部'
                     When Org_Order like 'D21%'  then '品研部'
                     When Org_Order like 'E10%'  then '廣州-庫吉'
                     else  '其他'
                   end as deptname
              from uV_TA13#EMP_LEVEL
              --到職日期>=本月月初
             Where e_rdate >= dateadd(mm,datediff(mm,'',getdate()),'') 
               and e_no like 'T%'
           ) Y
     group by deptname 
    
    --本月離職
    select deptname,count(1) as leave_people
           into #tmp_leave_people
      from (select *,
                   case 
                     When Org_Order like '110%'  then '董事長室' 
                     When Org_Order like '120%'  then '總經理室'
                     When Org_Order like '130%'  then '執行副總'
                     When Org_Order like 'A00%'  then '總管理處'
                     When Org_Order like 'A11%'  then '管理部'
                     When Org_Order like 'A21%'  then '財會部'
                     When Org_Order like 'A31%'  then '資訊部'
                     When Org_Order like 'A41%'  then '行企部'
                     When Org_Order like 'A51%'  then '採購部'
                     When Org_Order like 'A61%'  then '資材部'
                     When Org_Order like 'B11%' and e_no like 'T%' then '汽車精品課'
                     When Org_Order like 'B12%'  then '家庭五金課'
                     When Org_Order like 'B13%'  then '電子商務課'
                     When Org_Order like 'B16%'  then '國際貿易課'
                     When Org_Order like 'B19%'  then '客服部'
                     When Org_Order like 'D11%'  then '生產部'
                     When Org_Order like 'D21%'  then '品研部'
                     When Org_Order like 'E10%'  then '廣州-庫吉'
                     else  '其他'
                   end as deptname
              from uV_TA13#EMP_LEVEL
              --離職=Y and 離職日期 >= 本月月初
             Where (leave='Y') and (e_ldate >= dateadd(mm,datediff(mm,'',getdate()),'') ) 
               and e_no like 'T%'
           ) Y
     group by deptname 
    
    
    --本月轉入
    --WHERE 條件再依據設定值調整
    select deptname,count(1) as deptin
           into #tmp_deptin
      from (select Org_Order,m.e_rdate,
                   case 
                     When Org_Order like '110%'  then '董事長室' 
                     When Org_Order like '120%'  then '總經理室'
                     When Org_Order like '130%'  then '執行副總'
                     When Org_Order like 'A00%'  then '總管理處'
                     When Org_Order like 'A11%'  then '管理部'
                     When Org_Order like 'A21%'  then '財會部'
                     When Org_Order like 'A31%'  then '資訊部'
                     When Org_Order like 'A41%'  then '行企部'
                     When Org_Order like 'A51%'  then '採購部'
                     When Org_Order like 'A61%'  then '資材部'
                     When Org_Order like 'B11%' and m.e_no like 'T%' then '汽車精品課'
                     When Org_Order like 'B12%'  then '家庭五金課'
                     When Org_Order like 'B13%'  then '電子商務課'
                     When Org_Order like 'B16%'  then '國際貿易課'
                     When Org_Order like 'B19%'  then '客服部'
                     When Org_Order like 'D11%'  then '生產部'
                     When Org_Order like 'D21%'  then '品研部'
                     When Org_Order like 'E10%'  then '廣州-庫吉'
                     else  '其他'
                   end as deptname
              from uV_TA13#EMP_LEVEL m
                   join SYNC_TA13.dbo.pemploy d on m.e_no=d.e_no 
                   --離職=Y and 離職日期 >= 本月月初
             Where m.e_rdate = dateadd(mm,datediff(mm,'',getdate()),'') 
               and m.e_no like 'T%'
           ) Y
     group by deptname
    
    
    
    --本月轉出
    --WHERE 條件再依據設定值調整
    select deptname,count(1) as deptout
           into #tmp_deptout
      from (select Org_Order,m.e_rdate,
                   case 
                     When Org_Order like '110%'  then '董事長室' 
                     When Org_Order like '120%'  then '總經理室'
                     When Org_Order like '130%'  then '執行副總'
                     When Org_Order like 'A00%'  then '總管理處'
                     When Org_Order like 'A11%'  then '管理部'
                     When Org_Order like 'A21%'  then '財會部'
                     When Org_Order like 'A31%'  then '資訊部'
                     When Org_Order like 'A41%'  then '行企部'
                     When Org_Order like 'A51%'  then '採購部'
                     When Org_Order like 'A61%'  then '資材部'
                     When Org_Order like 'B11%' and m.e_no like 'T%' then '汽車精品課'
                     When Org_Order like 'B12%'  then '家庭五金課'
                     When Org_Order like 'B13%'  then '電子商務課'
                     When Org_Order like 'B16%'  then '國際貿易課'
                     When Org_Order like 'B19%'  then '客服部'
                     When Org_Order like 'D11%'  then '生產部'
                     When Org_Order like 'D21%'  then '品研部'
                     When Org_Order like 'E10%'  then '廣州-庫吉'
                     else  '其他'
                   end as deptname
              from uV_TA13#EMP_LEVEL m
                   join SYNC_TA13.dbo.pemploy d on m.e_no=d.e_no 
                   --離職日期 >= 本月月初
             Where m.e_rdate = dateadd(mm,datediff(mm,'',getdate()),'') 
               and m.e_no like 'T%'
           ) Y
     group by deptname
    
    --SQ Query
    
    select a.deptname,a.dept_people,b.pre_m_people,c.new_people,d.leave_people,e.deptin,f.deptout
        from #tmp_people a
             left join #tmp_lastm_people b 
               on a.deptname=b.deptname
             left join #tmp_new_people c 
               on a.deptname=c.deptname
             left join #tmp_leave_people d 
               on a.deptname=d.deptname
             left join #tmp_deptin e 
               on a.deptname=e.deptname
             left join #tmp_deptout f 
               on a.deptname=e.deptname
  
End
GO
