USE [DW]
GO
/****** Object:  View [dbo].[xV_Authority(201708����)]    Script Date: 08/18/2017 17:43:39 ******/
DROP VIEW [dbo].[xV_Authority(201708����)]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
Create View [dbo].[xV_Authority(201708����)]
as
select pw_comp=RTrim(isnull(pw_comp, '')), kind,
       m_no, m_name=Replicate('��', Len(RTrim(isnull(m_no, '   ')))-3)+RTrim(isnull(m_name, '')), 
       e_dept=Rtrim(isnull(d.e_dept, '')), dp_name=RTrim(isnull(dp_name, '')), 
       pw_epno=Rtrim(isnull(pw_epno, '')), pw_user=Rtrim(isnull(pw_user, '')), 
       e_duty=RTrim(isnull(d.e_duty, '')),
       '����'=case When substring(pw_fg, 2, 1)='E' then 'V' else '' end,
       '�s�W'=case When substring(pw_fg, 4, 1)='E' then 'V' else '' end,
       '�d��'=case When substring(pw_fg, 5, 1)='E' then 'V' else '' end,
       '�ק�'=case When substring(pw_fg, 7, 1)='E' then 'V' else '' end,
       '�R��'=case When substring(pw_fg, 8, 1)='E' then 'V' else '' end,
       '�C�L'=case When substring(pw_fg, 10, 1)='E' then 'V' else '' end      
  from (select Kind='�Ͳ��s�y', * from Ori_ALL#menubar
        union all
        select Kind='�]�|�t��', * from Ori_ALL#amenubar
        union all
        select Kind='��T�t��', * from Ori_ALL#tmenubar
        union all
        select Kind='�~�ިt��', * from Ori_ALL#masmenu
        union all
        select Kind='���רt��', * from Ori_ALL#mafmenu
       ) m
       left join Ori_TA13#pemploy d on m.pw_user=d.e_name
       left join Ori_TA13#pdept d1 on d.e_dept=d1.dp_no
GO
