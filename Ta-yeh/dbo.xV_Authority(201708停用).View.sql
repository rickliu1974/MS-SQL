USE [DW]
GO
/****** Object:  View [dbo].[xV_Authority(201708停用)]    Script Date: 08/18/2017 17:43:39 ******/
DROP VIEW [dbo].[xV_Authority(201708停用)]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
Create View [dbo].[xV_Authority(201708停用)]
as
select pw_comp=RTrim(isnull(pw_comp, '')), kind,
       m_no, m_name=Replicate('－', Len(RTrim(isnull(m_no, '   ')))-3)+RTrim(isnull(m_name, '')), 
       e_dept=Rtrim(isnull(d.e_dept, '')), dp_name=RTrim(isnull(dp_name, '')), 
       pw_epno=Rtrim(isnull(pw_epno, '')), pw_user=Rtrim(isnull(pw_user, '')), 
       e_duty=RTrim(isnull(d.e_duty, '')),
       '執行'=case When substring(pw_fg, 2, 1)='E' then 'V' else '' end,
       '新增'=case When substring(pw_fg, 4, 1)='E' then 'V' else '' end,
       '查詢'=case When substring(pw_fg, 5, 1)='E' then 'V' else '' end,
       '修改'=case When substring(pw_fg, 7, 1)='E' then 'V' else '' end,
       '刪除'=case When substring(pw_fg, 8, 1)='E' then 'V' else '' end,
       '列印'=case When substring(pw_fg, 10, 1)='E' then 'V' else '' end      
  from (select Kind='生產製造', * from Ori_ALL#menubar
        union all
        select Kind='財會系統', * from Ori_ALL#amenubar
        union all
        select Kind='國貿系統', * from Ori_ALL#tmenubar
        union all
        select Kind='業管系統', * from Ori_ALL#masmenu
        union all
        select Kind='維修系統', * from Ori_ALL#mafmenu
       ) m
       left join Ori_TA13#pemploy d on m.pw_user=d.e_name
       left join Ori_TA13#pdept d1 on d.e_dept=d1.dp_no
GO
