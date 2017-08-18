USE [DW]
GO
/****** Object:  StoredProcedure [dbo].[uSP_Sys_Get_Columns]    Script Date: 08/18/2017 17:18:56 ******/
DROP PROCEDURE [dbo].[uSP_Sys_Get_Columns]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
Create Procedure [dbo].[uSP_Sys_Get_Columns](@Src_DBName Varchar(100), @TableName Varchar(100), 
                                           @String Varchar(Max), @Token Varchar(100), @FieldNames Varchar(Max) output)
as
Begin
  /*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
   Program Name: uSP_Sys_Get_Columns
   Create Date: 2013/03/01
   Creator: Rickliu
   Updated Date: 2013/09/06 [變更 Trans_Log 訊息顯示內容]
   
   
   XType Datatype 
   ----------------
   34 image
   35 text
   36 uniqueidentifier
   48 tinyint
   52 smallint
   56 int
   58 smalldatetime
   59 real
   60 money
   61 datetime
   62 float
   98 sql_variant
   99 ntext
   104 bit
   106 decimal
   108 numeric
   122 smallmoney
   127 bigint
   165 varbinary
   167 varchar
   173 binary
   175 char
   189 timestamp
   231 nvarchar
   231 sysname
   239 nchar
   241 xml

  *-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*/  

  Declare @Proc Varchar(50) = 'uSP_Sys_Get_Columns'
  Declare @Msg Varchar(Max) =''
  declare @strSQL Varchar(Max)=''
  declare @Cr Varchar(10) = Char(13)+Char(10)
  declare @ErrCode Int = -1
  
  declare @RString Varchar(Max)=''
  
  declare @Column_Name Varchar(100)=''
  declare @Data_Type Varchar(100)=''
  declare @Field_Length Integer
  declare @Field_scale Integer
  declare @Can_Null Varchar(100)
  declare @Collation_Name Varchar(255)
  
  declare @Column_String Varchar(100)=''
  
  
  declare @sField Varchar(Max)=''
  declare @Cnt Int=0

  if object_id('tempdb..##tmp_Columns') is not null 
  begin
     Set @Msg = '刪除資料表 [tempdb..##tmp_Columns]'
     set @strSQL= 'DROP TABLE ##tmp_Columns'

     Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
  end

  set @Msg = '寫入暫存資料表中。'
  if @Src_DBName <> ''
     Set @strSQL =
            'select Column_Name=D.Name, '+@Cr+
            '       Data_Type=D1.Name, '+@Cr+
            '       Field_Length=D.Prec, '+@Cr+
            '       Field_scale=D.scale, '+@Cr+
            '       Can_Null=Case isnullable when 1 then ''Null'' else ''Not Null'' end, '+@Cr+
            '       Collation_Name=D.Collation '+@Cr+
            '       into ##tmp_Columns'+@Cr+
            '  from '+@Src_DbName+'.[dbo].[SysObjects] m '+@Cr+
            '        left join '+@Src_DbName+'.dbo.syscolumns D on M.id=D.id '+@Cr+
            '        left join '+@Src_DbName+'.dbo.systypes D1 on D.xtype = D1.xtype and D.xusertype=D1.xusertype '+@Cr+
            ' where m.name = '''+@TableName+''' '+@Cr+
            '   and (d.xtype not in (36, 189) or d.xusertype not in (36, 189)) '+@Cr+
            ' order by D.colorder '
  else
     Set @strSQL =
            'select Column_Name=D.Name, '+@Cr+
            '       Data_Type=D1.Name, '+@Cr+
            '       Field_Length=D.Prec, '+@Cr+
            '       Field_scale=D.scale, '+@Cr+
            '       Can_Null=Case isnullable when 1 then ''Null'' else ''Not Null'' end, '+@Cr+
            '       Collation_Name=D.Collation '+@Cr+
            '       into ##tmp_Columns'+@Cr+
            '  from [SysObjects] m '+@Cr+
            '        left join syscolumns D on M.id=D.id '+@Cr+
            '        left join systypes D1 on D.xtype = D1.xtype and D.xusertype=D1.xusertype '+@Cr+
            ' where m.name = '''+@TableName+''' '+@Cr+
            '   and (d.xtype not in (36, 189) or d.xusertype not in (36, 189)) '+@Cr+
            ' order by D.colorder '

  Exec @Cnt = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
  
  if @Cnt = @ErrCode Goto End_Proc

  Declare Cur_Columns Cursor for
    select * from ##tmp_Columns
  
  set @FieldNames =''
  set @Cnt =0
  Open Cur_Columns
  Fetch Next From Cur_Columns into @Column_Name, @Data_Type, @Field_Length, @Field_scale, @Can_Null, @Collation_Name

  While @@FETCH_STATUS = 0
  Begin
    if @String = ''
       Set @String = '@COL_NAME' 
    if @Token = ''
       Set @Token = ', '
    Set @String = Upper(@String)

    Set @RString = 
          Replace(
          Replace(
          Replace(
          Replace(
          Replace(@String, 
                  '@COL_DEFAULT', 
                       Case 
                         when Upper(@Data_Type) in ('INT', 'FLOAT', 'DECIMAL', 'BIT') then '['+@Column_Name+']=0 '
                         else '['+@Column_Name+']='''' '
                       end),
                  '@COL_COLLATION', IsNull(@Collation_Name, '')), 
                  '@COL_LEN', IsNull('('+Convert(Varchar(100), @Field_Length)+IsNull(Convert(Varchar(100), @Field_scale), '')+')', '')),
                  '@COL_TYPE', '['+@Data_Type+']'),
                  '@COL_NAME', '['+@Column_Name+']')

    Set @RString = Upper(@RString)
    if @FieldNames = ''
       Set @FieldNames = @RString
    else
       Set @FieldNames = @FieldNames+@Token+@RString
    set @Cnt = @Cnt +1
    
    if @Cnt >=5
    begin
       Set @FieldNames = @FieldNames + @CR
       Set @Cnt = 0
    end
    Fetch Next From Cur_Columns into @Column_Name, @Data_Type, @Field_Length, @Field_scale, @Can_Null, @Collation_Name
  end
  set @FieldNames = '  '+@FieldNames
  print '參數：@COL_NAME 欄位名稱, @COL_COLLATION 定序, @COL_DEFAULT 預設值, @COL_LEN 長度, @COL_TYPE 型態'
  print 'Result Print:'
  Print @FieldNames
  print 'Print End.'

  Close Cur_Columns
  DEALLOCATE Cur_Columns

End_Proc:
  Return(@Cnt)
  
end;
GO
