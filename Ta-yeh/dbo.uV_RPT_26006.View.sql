USE [DW]
GO
/****** Object:  View [dbo].[uV_RPT_26006]    Script Date: 07/24/2017 14:43:55 ******/
DROP VIEW [dbo].[uV_RPT_26006]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
Create View [dbo].[uV_RPT_26006]
as
select m.*,f.sk_fld9+f.sk_fld10 as feature
       ,substring(m.brcodeno1,1,12) as brcode1
       ,substring(m.brcodeno2,1,12) as brcode2
       ,substring(m.brcodeno3,1,12) as brcode3
       ,substring(m.brcodeno4,1,12) as brcode4
       ,substring(m.brcodeno5,1,12) as brcode5 

from(
select distinct
       kind,
       sk_no as group_no,
       sk_name as group_name

	   ,pic_6,pic_4,pic_2
	   ,main_pic,pic1,pic2,pic3
	   

	   ,REPLACE( REPLACE(main_pic,'W:\5.產品圖','http://192.168.1.51/Smart-Query/pic'),'\','/') as main_pic_1
	   ,REPLACE( REPLACE(pic1,'W:\5.產品圖','http://192.168.1.51/Smart-Query/pic'),'\','/') as pic1_1
	   ,REPLACE( REPLACE(pic2,'W:\5.產品圖','http://192.168.1.51/Smart-Query/pic'),'\','/') as pic2_1
	   ,REPLACE( REPLACE(pic3,'W:\5.產品圖','http://192.168.1.51/Smart-Query/pic'),'\','/') as pic3_1
	   


	   ,avg_price,gross_property,package,price,product_property,color,size
	   ,replace(replace((
                         select [1] as brcodeno
                           from
                               (
                                select distinct kind
	                                  ,group_no
	                                  ,barcode_name
	                                  ,sk_bcode
									  ,b.sk_no
	                                  ,RANK() OVER (ORDER BY sk_bcode) AS brcodeno
                                 from Ori_xls#Stock_Property as a left join Fact_sstock as b on a.sk_no=b.sk_no collate Chinese_Taiwan_Stroke_CI_AS
                                where a.group_no = m.sk_no and b.sk_bcode > '') a
                                pivot (MAX(sk_bcode) for brcodeno in ([1], [2], [3], [4], [5])) b
                          group by kind,group_no,[1], [2], [3], [4], [5]
                            for xml path('')
                         ),'<brcodeno>',''),'</brcodeno>','') as brcodeno1
	   ,replace(replace((
                         select [2] as brcodeno
                           from
                               (
                                select distinct kind
	                                  ,group_no
	                                  ,barcode_name
	                                  ,sk_bcode
									  ,b.sk_no
	                                  ,RANK() OVER (ORDER BY sk_bcode) AS brcodeno
                                 from Ori_xls#Stock_Property as a left join Fact_sstock as b on a.sk_no=b.sk_no collate Chinese_Taiwan_Stroke_CI_AS
                                where a.group_no = m.sk_no and b.sk_bcode > '') a
                                pivot (MAX(sk_bcode) for brcodeno in ([1], [2], [3], [4], [5])) b
                          group by kind,group_no,[1], [2], [3], [4], [5]
                            for xml path('')
                         ),'<brcodeno>',''),'</brcodeno>','') as brcodeno2
	   ,replace(replace((
                         select [3] as brcodeno
                           from
                               (
                                select distinct kind
	                                  ,group_no
	                                  ,barcode_name
	                                  ,sk_bcode
									  ,b.sk_no
	                                  ,RANK() OVER (ORDER BY sk_bcode) AS brcodeno
                                 from Ori_xls#Stock_Property as a left join Fact_sstock as b on a.sk_no=b.sk_no collate Chinese_Taiwan_Stroke_CI_AS
                                where a.group_no = m.sk_no and b.sk_bcode > '') a
                                pivot (MAX(sk_bcode) for brcodeno in ([1], [2], [3], [4], [5])) b
                          group by kind,group_no,[1], [2], [3], [4], [5]
                            for xml path('')
                         ),'<brcodeno>',''),'</brcodeno>','') as brcodeno3
	   ,replace(replace((
                         select [4] as brcodeno
                           from
                               (
                                select distinct kind
	                                  ,group_no
	                                  ,barcode_name
	                                  ,sk_bcode
									  ,b.sk_no
	                                  ,RANK() OVER (ORDER BY sk_bcode) AS brcodeno
                                 from Ori_xls#Stock_Property as a left join Fact_sstock as b on a.sk_no=b.sk_no collate Chinese_Taiwan_Stroke_CI_AS
                                where a.group_no = m.sk_no and b.sk_bcode > '') a
                                pivot (MAX(sk_bcode) for brcodeno in ([1], [2], [3], [4], [5])) b
                          group by kind,group_no,[1], [2], [3], [4], [5]
                            for xml path('')
                         ),'<brcodeno>',''),'</brcodeno>','') as brcodeno4
	   ,replace(replace((
                         select [5] as brcodeno
                           from
                               (
                                select distinct kind
	                                  ,group_no
	                                  ,barcode_name
	                                  ,sk_bcode
									  ,b.sk_no
	                                  ,RANK() OVER (ORDER BY sk_bcode) AS brcodeno
                                 from Ori_xls#Stock_Property as a left join Fact_sstock as b on a.sk_no=b.sk_no collate Chinese_Taiwan_Stroke_CI_AS
                                where a.group_no = m.sk_no and b.sk_bcode > '') a
                                pivot (MAX(sk_bcode) for brcodeno in ([1], [2], [3], [4], [5])) b
                          group by kind,group_no,[1], [2], [3], [4], [5]
                            for xml path('')
                         ),'<brcodeno>',''),'</brcodeno>','') as brcodeno5
	   ,replace(replace((
                         select [1] as brcodename
                           from
                               (
                                select distinct kind
	                                  ,group_no
	                                  ,barcode_name
	                                  ,sk_bcode
									  ,b.sk_no
	                                  ,RANK() OVER (ORDER BY sk_bcode) AS brcodeno
                                 from Ori_xls#Stock_Property as a left join Fact_sstock as b on a.sk_no=b.sk_no collate Chinese_Taiwan_Stroke_CI_AS
                                where a.group_no = m.sk_no and b.sk_bcode > '') a
                                pivot (MAX(barcode_name) for brcodeno in ([1], [2], [3], [4], [5])) b
                          group by kind,group_no,[1], [2], [3], [4], [5]
                            for xml path('')
                         ),'<brcodename>',''),'</brcodename>','') as brcodename1
	   ,replace(replace((
                         select [2] as brcodename
                           from
                               (
                                select distinct kind
	                                  ,group_no
	                                  ,barcode_name
	                                  ,sk_bcode
									  ,b.sk_no
	                                  ,RANK() OVER (ORDER BY sk_bcode) AS brcodeno
                                 from Ori_xls#Stock_Property as a left join Fact_sstock as b on a.sk_no=b.sk_no collate Chinese_Taiwan_Stroke_CI_AS
                                where a.group_no = m.sk_no and b.sk_bcode > '') a
                                pivot (MAX(barcode_name) for brcodeno in ([1], [2], [3], [4], [5])) b
                          group by kind,group_no,[1], [2], [3], [4], [5]
                            for xml path('')
                         ),'<brcodename>',''),'</brcodename>','') as brcodename2
	   ,replace(replace((
                         select [3] as brcodename
                           from
                               (
                                select distinct kind
	                                  ,group_no
	                                  ,barcode_name
	                                  ,sk_bcode
									  ,b.sk_no
	                                  ,RANK() OVER (ORDER BY sk_bcode) AS brcodeno
                                 from Ori_xls#Stock_Property as a left join Fact_sstock as b on a.sk_no=b.sk_no collate Chinese_Taiwan_Stroke_CI_AS
                                where a.group_no = m.sk_no and b.sk_bcode > '') a
                                pivot (MAX(barcode_name) for brcodeno in ([1], [2], [3], [4], [5])) b
                          group by kind,group_no,[1], [2], [3], [4], [5]
                            for xml path('')
                         ),'<brcodename>',''),'</brcodename>','') as brcodename3
	   ,replace(replace((
                         select [4] as brcodename
                           from
                               (
                                select distinct kind
	                                  ,group_no
	                                  ,barcode_name
	                                  ,sk_bcode
									  ,b.sk_no
	                                  ,RANK() OVER (ORDER BY sk_bcode) AS brcodeno
                                 from Ori_xls#Stock_Property as a left join Fact_sstock as b on a.sk_no=b.sk_no collate Chinese_Taiwan_Stroke_CI_AS
                                where a.group_no = m.sk_no and b.sk_bcode > '') a
                                pivot (MAX(barcode_name) for brcodeno in ([1], [2], [3], [4], [5])) b
                          group by kind,group_no,[1], [2], [3], [4], [5]
                            for xml path('')
                         ),'<brcodename>',''),'</brcodename>','') as brcodename4
	   ,replace(replace((
                         select [5] as brcodename
                           from
                               (
                                select distinct kind
	                                  ,group_no
	                                  ,barcode_name
	                                  ,sk_bcode
									  ,b.sk_no
	                                  ,RANK() OVER (ORDER BY sk_bcode) AS brcodeno
                                 from Ori_xls#Stock_Property as a left join Fact_sstock as b on a.sk_no=b.sk_no collate Chinese_Taiwan_Stroke_CI_AS
                                where a.group_no = m.sk_no and b.sk_bcode > '') a
                                pivot (MAX(barcode_name) for brcodeno in ([1], [2], [3], [4], [5])) b
                          group by kind,group_no,[1], [2], [3], [4], [5]
                            for xml path('')
                         ),'<brcodename>',''),'</brcodename>','') as brcodename5
	   ,replace(replace((
                         select [1] as brcodeskno
                           from
                               (
                                select distinct kind
	                                  ,group_no
	                                  ,barcode_name
	                                  ,sk_bcode
									  ,b.sk_no
	                                  ,RANK() OVER (ORDER BY sk_bcode) AS brcodeno
                                 from Ori_xls#Stock_Property as a left join Fact_sstock as b on a.sk_no=b.sk_no collate Chinese_Taiwan_Stroke_CI_AS
                                where a.group_no = m.sk_no and b.sk_bcode > '') a
                                pivot (MAX(a.sk_no) for brcodeno in ([1], [2], [3], [4], [5])) b
                          group by kind,group_no,[1], [2], [3], [4], [5]
                            for xml path('')
                         ),'<brcodeskno>',''),'</brcodeskno>','') as brcodeskno1
	   ,replace(replace((
                         select [2] as brcodeskno
                           from
                               (
                                select distinct kind
	                                  ,group_no
	                                  ,barcode_name
	                                  ,sk_bcode
									  ,b.sk_no
	                                  ,RANK() OVER (ORDER BY sk_bcode) AS brcodeno
                                 from Ori_xls#Stock_Property as a left join Fact_sstock as b on a.sk_no=b.sk_no collate Chinese_Taiwan_Stroke_CI_AS
                                where a.group_no = m.sk_no and b.sk_bcode > '') a
                                pivot (MAX(a.sk_no) for brcodeno in ([1], [2], [3], [4], [5])) b
                          group by kind,group_no,[1], [2], [3], [4], [5]
                            for xml path('')
                         ),'<brcodeskno>',''),'</brcodeskno>','') as brcodeskno2
	   ,replace(replace((
                         select [3] as brcodeskno
                           from
                               (
                                select distinct kind
	                                  ,group_no
	                                  ,barcode_name
	                                  ,sk_bcode
									  ,b.sk_no
	                                  ,RANK() OVER (ORDER BY sk_bcode) AS brcodeno
                                 from Ori_xls#Stock_Property as a left join Fact_sstock as b on a.sk_no=b.sk_no collate Chinese_Taiwan_Stroke_CI_AS
                                where a.group_no = m.sk_no and b.sk_bcode > '') a
                                pivot (MAX(a.sk_no) for brcodeno in ([1], [2], [3], [4], [5])) b
                          group by kind,group_no,[1], [2], [3], [4], [5]
                            for xml path('')
                         ),'<brcodeskno>',''),'</brcodeskno>','') as brcodeskno3
	   ,replace(replace((
                         select [4] as brcodeskno
                           from
                               (
                                select distinct kind
	                                  ,group_no
	                                  ,barcode_name
	                                  ,sk_bcode
									  ,b.sk_no
	                                  ,RANK() OVER (ORDER BY sk_bcode) AS brcodeno
                                 from Ori_xls#Stock_Property as a left join Fact_sstock as b on a.sk_no=b.sk_no collate Chinese_Taiwan_Stroke_CI_AS
                                where a.group_no = m.sk_no and b.sk_bcode > '') a
                                pivot (MAX(a.sk_no) for brcodeno in ([1], [2], [3], [4], [5])) b
                          group by kind,group_no,[1], [2], [3], [4], [5]
                            for xml path('')
                         ),'<brcodeskno>',''),'</brcodeskno>','') as brcodeskno4
	   ,replace(replace((
                         select [5] as brcodeskno
                           from
                               (
                                select distinct kind
	                                  ,group_no
	                                  ,barcode_name
	                                  ,sk_bcode
									  ,b.sk_no
	                                  ,RANK() OVER (ORDER BY sk_bcode) AS brcodeno
                                 from Ori_xls#Stock_Property as a left join Fact_sstock as b on a.sk_no=b.sk_no collate Chinese_Taiwan_Stroke_CI_AS
                                where a.group_no = m.sk_no and b.sk_bcode > '') a
                                pivot (MAX(a.sk_no) for brcodeno in ([1], [2], [3], [4], [5])) b
                          group by kind,group_no,[1], [2], [3], [4], [5]
                            for xml path('')
                         ),'<brcodeskno>',''),'</brcodeskno>','') as brcodeskno5
  from Ori_xls#Stock_Property m) as m left join Fact_sstock as f on m.brcodeno1=f.sk_bcode
 where m.kind = 'G'
GO
