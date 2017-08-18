USE [DW]
GO
/****** Object:  StoredProcedure [dbo].[uSP_Imp_Warehouse_LA_Adjust]    Script Date: 08/18/2017 17:18:56 ******/
DROP PROCEDURE [dbo].[uSP_Imp_Warehouse_LA_Adjust]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE Proc [dbo].[uSP_Imp_Warehouse_LA_Adjust]
as
begin
  /***********************************************************************************************************
    2014/10/14 Rickliu 
    每月固定 10 日產生調整單據至 TA13，該調整單據僅針對 LA 倉，並將所有品項調整為零，調整單據僅產生未確認單     
  ***********************************************************************************************************/
  Declare @Proc Varchar(50) = 'uSP_Imp_Warehouse_LA_Adjust'
  Declare @Cnt Int = 0
  Declare @RowCnt Int = 0
  Declare @Msg Varchar(4000) =''
  Declare @Errcode int = -1
  Declare @subject Varchar(100)

  Declare @Kind Varchar(1), @KindName Varchar(10), @sd_class Varchar(1), @sd_slip_fg Varchar(1), @OD_Name Varchar(20), @OD_CName Varchar(20), @oper Varchar(1)
  Declare @xls_head Varchar(100), @TB_head Varchar(200), @TB_xls_Name Varchar(200), @TB_OD_Name Varchar(200), @TB_tmp_name Varchar(200)

  Declare @strSQL Varchar(Max)
  Declare @CR Varchar(4) = char(13)+char(10)
  Declare @RowCount Table (cnt int)
  Declare @RowData Table (aData Varchar(255))
  
  Declare @new_sdno varchar(10)
  Declare @sd_date varchar(10)
  Declare @sp_date varchar(10)


  Declare @Print_Date varchar(10)
  Declare @Last_Date Varchar(2)
  Declare @sp_cnt int
       
  --產生調整日期，2014/10/14 林課表示每月固定以 10 號產生調整單據
  set @Last_Date = '10'
  set @sp_date = Substring(Convert(Varchar(10), getdate(), 111), 1, 8)+@Last_Date
  set @sd_date = @sp_date
  Set @sd_class = '5' -- 調整.調撥單
  Set @sd_slip_fg = 'A' -- 調整單
  
  Declare @Rm Varchar(100), @sp_maker Varchar(20), @str Varchar(200)
  Declare @Pos Int

  Set @Rm = '系統匯入調整單('+Convert(Varchar(10), getdate(), 111)+')'
  Set @sp_maker = 'Admin'

  -- 判別單據是否存在
  select @sp_cnt = sum(cnt)
    from (select count(1) as cnt
            from SYNC_TA13.dbo.sslpdt m
           where sd_class = @sd_class
             and sd_slip_fg = @sd_slip_fg
             and sd_date = @sp_date
             and sd_lotno like '%'+@Rm+'%'
           union
          select count(1) as cnt
            from SYNC_TA13.dbo.sslpdttmp m
           where sd_class = @sd_class
             and sd_slip_fg = @sd_slip_fg
             and sd_date = @sp_date
             and sd_lotno like '%'+@Rm+'%') m

  if @sp_cnt > 0
  begin
     set @Cnt = @Errcode
     set @strSQL = ''
     Set @subject = 'Exec '+@Proc+' LA倉產生調整單據失敗!!'
     set @Msg ='已有 ['+@sp_Date+'] 調整單據(已/未確認)資料，若需要重轉則先行清除確認及未確認單據後重轉即可。sslpdt, @sd_class=['+@sd_class+'], @sd_slip_fg=['+@sd_slip_fg+'], @sd_date=['+@sp_Date+']。'
     Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt
     Execute uSP_Sys_Send_Mail @Proc, @Subject, 'fa@ta-yeh.com.tw', @Msg, ''

     Return(@Errcode)
  end
  else
  begin
     --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
     -- 取得最新單號
     set @new_sdno = (select convert(varchar(10), isnull(max(sd_no), replace(substring(@sd_date, 3, 8), '/', '') +'0000')) 
                        from (select distinct sd_no
                                from SYNC_TA13.dbo.sslpdt m
                               where sd_class = @sd_class
                                 and sd_slip_fg = @sd_slip_fg
                                 and sd_date = @sd_date
                               union
                              select distinct sd_no
                                from SYNC_TA13.dbo.sslpdttmp m
                               where sd_class = @sd_class
                                 and sd_slip_fg = @sd_slip_fg
                                 and sd_date = @sd_date
                             ) m
                     )
     set @New_sdno = Convert(Varchar(10), Convert(Int, @New_sdno) +1)
     set @Msg = '取得最新調整單號(必須從已核准以及未核准的表單取號),new_sdno=['+@new_sdno+'], sd_class=['+@sd_class+'], sd_slip_fg=['+@sd_slip_fg+'], sd_date=['+@sd_date+'].'
     Exec uSP_Sys_Write_Log @Proc, @Msg, '', @Cnt
     --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
     set @Msg = '新增['+@sd_date+'] 日調整明細。'
     set @strSQL ='insert into TYEIPDBS2.lytdbta13.dbo.sslpdttmp '+@CR+
                  '(sd_class, sd_slip_fg, sd_date, sd_no, '+@CR+
                  ' sd_skno, sd_name, sd_whno, sd_whno2, sd_qty, '+@CR+
                  ' sd_price, sd_dis, sd_stot, sd_lotno, sd_unit, '+@CR+
                  ' sd_unit_fg, sd_ave_p, sd_rate, sd_seqfld, sd_ordno '+@CR+
                  ') '+@CR+
                  'select '''+@sd_class+''' as sd_class, '+@CR+ -- 類別
                  '       '''+@sd_slip_fg+''' sd_slip_fg, '+@CR+ -- 單據種類
                  '       convert(datetime, '''+@sd_date+''') as sd_date, '+@CR+ -- 貨單日期
                  '       '''+@New_sdno+''' as sd_no, '+@CR+ -- 貨單編號
                  ''+@CR+
                  '       m.wd_skno as sd_skno, '+@CR+ -- 貨品編號
                  '       m.sk_name as sd_name, '+@CR+ -- 品名規格
                  '       ''LA'' as sd_whno, '+@CR+ -- 倉庫(入)
                  '       '''' as sd_whno2, '+@CR+ -- 倉庫(出)
                  '       m.wd_last_qty * -1 as sd_qty, '+@CR+ -- 交易數量
                  ''+@CR+
                  '       m.sk_save as sd_price, '+@CR+ -- 價格
                  '       0 as sd_dis, '+@CR+ -- 折數
                  '       m.wd_last_qty * -1 * m.sk_save as sd_stot, '+@CR+ -- 小計
                  '       '''+@Rm+''' as sd_lotno, '+@CR+ -- 備用欄位
                  '       m.sk_unit as sd_unit, '+@CR+ -- 單位
                  ' '+@CR+
                  '       0 as sd_unit_fg, '+@CR+ -- 單位旗標
                  '       m.sk_save as sd_ave_p, '+@CR+ -- 單位成本
                  '       1 as sd_rate, '+@CR+ -- 匯率
                  '       rowid as sd_seqfld, '+@CR+ -- 明細序號, 此序號會因為凌越修改而加以變更，所以另存一份到 sd_ordno
                  '       rowid as sd_ordno'+@CR+ -- XLS 明細序號
                  '  from (select m.wd_skno, d.sk_name, sk_unit, sk_save, '+@CR+
                  '               sum(isnull(wd_amt0+wd_amt1+wd_amt2+wd_amt3+wd_amt4+wd_amt5+wd_amt6+ '+@CR+
                  '                          wd_amt7+wd_amt8+wd_amt9+wd_amt10+wd_amt11+wd_amt12, 0)) as wd_last_qty, '+@CR+
                  '               row_number() over(order by m.wd_skno) as rowid  '+@CR+
                  '          from SYNC_TA13.dbo.swaredt m '+@CR+
                  '               left join SYNC_TA13.dbo.sstock d '+@CR+
                  '                 on m.wd_skno = d.sk_no '+@CR+
                  -- 2014/10/15 Rickliu 原本使用取當年度的庫存資料，但發現凌越的庫存資料是有異動才會進行過帳，
                  -- 因此若去年度的商品若沒異動則庫存數仍停留在去年，所以改採抓取每個商品最後一年的庫存資料。
                  '               inner join '+@CR+
                  '               (select wd_no, max(wd_yr) as wd_yr, wd_skno '+@CR+
                  '                  from SYNC_TA13.dbo.swaredt '+@CR+
                  '                 group by wd_no, wd_skno '+@CR+
                  '               ) d1 '+@CR+
                  '                 on m.wd_no = d1.wd_no '+@CR+
                  '                and m.wd_yr = d1.wd_yr '+@CR+
                  '                and m.wd_skno = d1.wd_skno '+@CR+
                  '         where 1=1 '+@CR+
                  '           and m.wd_no = ''LA'' '+@CR+
                  '           and m.wd_class=''0'' '+@CR+
                  '           and m.wd_skno like ''A%'' '+@CR+
                  '         group by m.wd_skno, d.sk_name, sk_unit, sk_save '+@CR+
                  '         having sum(isnull(wd_amt0+wd_amt1+wd_amt2+wd_amt3+wd_amt4+wd_amt5+wd_amt6+ '+@CR+
                  '                           wd_amt7+wd_amt8+wd_amt9+wd_amt10+wd_amt11+wd_amt12, 0)) <> 0 '+@CR+
                  '       ) m '+@CR+
                  ' Order by 1, 2, 3'
     Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
     
     --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
     -- 新增調整主檔 PKey: sp_no (ASC), sp_slip_fg (ASC)
     set @Msg = '新增調整主檔。'
     set @strSQL ='insert into TYEIPDBS2.lytdbta13.dbo.ssliptmp '+@CR+
                  '(sp_class, sp_slip_fg, sp_date, sp_pdate, sp_no, '+@CR+
                  ' sp_maker, sp_tot, sp_tax, sp_dis, '+@CR+
                  ' sp_pay_kd, sp_rate_nm, sp_rate, sp_itot, sp_inv_kd, '+@CR+
                  ' sp_tax_kd, sp_invtype, sp_rem, sp_tal_rec '+@CR+
                  ') '+@CR+
                  'select distinct sd_class as sp_class, '+@CR+
                  '       sd_slip_fg as sp_slip_fg, '+@CR+
                  '       sd_date as sp_date, '+@CR+
                  '       sd_date as sp_pdate, '+@CR+
                  '       sd_no as sp_no, '+@CR+
                  ' '+@CR+
                  '       '''+@sp_maker+''' as sp_maker, '+@CR+ -- 製單人員
                  '       sum(sd_stot) as sp_tot, '+@CR+ -- 小計
                  
                  -- 2013/6/10 托售回貨不需要稅金
                  '       0 as sp_tax, '+@CR+ -- 營業稅(原)
                  
                  '       0 as sp_dis, '+@CR+ -- 折讓金額(原)
                  ' '+@CR+
                  '       '''' as sp_pay_kd, '+@CR+ -- 售價種類
                  '       ''NT'' as sp_rate_nm, '+@CR+ -- 匯率名稱
                  '       1 as sp_rate, '+@CR+ -- 匯率
                  '       0 as sp_itot, '+@CR+ -- 發票金額
                  '       0 as sp_inv_kd, '+@CR+ -- 發票類別(=1  三聯式,=2  二聯式,=3  收銀機)
                  ' '+@CR+
                  '       0 as sp_tax_kd, '+@CR+ -- 稅別(=1應稅,=2零稅 )
                  '       0 as sp_invtype, '+@CR+ -- 開立方式(=1未開, =2隨單開立, =3批次開立)
                  '       '''+@Rm+''' as sp_rem, '+@CR+ -- 備註
                  '       Count(1) as sp_tal_rec '+@CR+ -- 明細總筆數
                  '  from TYEIPDBS2.lytdbta13.dbo.sslpdttmp m '+@CR+
                  ' where sd_class = '''+@sd_class+''' '+@CR+
                  '   and sd_slip_fg = '''+@sd_slip_fg+''' '+@CR+
                  '   and sd_date = '''+@sd_date+''' '+@CR+
                  '   and sd_lotno like '''+@RM+'%'' '+@CR+
                  ' group by sd_class, '+@CR+
                  '       sd_slip_fg, '+@CR+
                  '       sd_date, '+@CR+
                  '       sd_no '

     Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL

     Set @subject = 'Exec '+@Proc+' 成功!!'
     set @Msg ='成功匯入 ['+@sp_Date+'] LA倉調整單據(未確認)資料。'
     Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt
     Execute uSP_Sys_Send_Mail @Proc, @Subject, 'fa@ta-yeh.com.tw', @Msg, ''
  end

  Return(0)
end
GO
