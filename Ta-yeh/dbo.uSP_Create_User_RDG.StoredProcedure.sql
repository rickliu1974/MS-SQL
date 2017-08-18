USE [DW]
GO
/****** Object:  StoredProcedure [dbo].[uSP_Create_User_RDG]    Script Date: 08/18/2017 17:18:56 ******/
DROP PROCEDURE [dbo].[uSP_Create_User_RDG]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
create procedure [dbo].[uSP_Create_User_RDG](@sDomain Varchar(100), @sAdmin_PWD varchar(100), @sLoca_Admin_PWD varchar(100))
as
begin
  Declare @Proc Varchar(50) = 'uSP_Create_User_RDG'
  Declare @sServer_Name Varchar(500), @sDisplay Varchar(500)= '', @sAccountID Varchar(500)= '', @sPassWord Varchar(500)= ''

  Declare @sText Varchar(Max) = ''
  Declare @RowID Int = 0
  Declare @CR Varchar(4) =Char(13)+Char(10)
  Declare @sSQL Varchar(2000)
  
  IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[Create_User_RDG_tmp]') AND type in (N'U'))
     Drop Table Create_User_RDG_tmp

  Create Table Create_User_RDG_tmp (
    RowID Int,
    Kind Varchar(1), -- C: Common, U:User, A:Admin
    Data Varchar(4000)
  )
  
  set @RowID = 1
  Set @sText = 
        -- 2015/11/09 Rickliu 原本 encoding = 'UTF-8', 但 MSSQL 2008 R2 已經不支援，因此只好將 xml 內的 encoding 改為 unicode，而 Bcp 參數則改為 -w 而非 -c
        '<?xml version="1.0" encoding="unicode"?>'+@CR+
        '<RDCMan schemaVersion="1">'+@CR+
        '    <version>2.2</version>'+@CR+
        '    <file>'+@CR+
        '        <properties>'+@CR+
        '            <name>大業用戶端</name>'+@CR+
        '            <expanded>True</expanded>'+@CR+
        '            <comment />'+@CR+
        '            <logonCredentials inherit="FromParent" />'+@CR+
        '            <connectionSettings inherit="FromParent" />'+@CR+
        '            <gatewaySettings inherit="FromParent" />'+@CR+
        '            <remoteDesktop inherit="FromParent" />'+@CR+
        '            <localResources inherit="FromParent" />'+@CR+
        '            <securitySettings inherit="FromParent" />'+@CR+
        '            <displaySettings inherit="FromParent" />'+@CR+
        '        </properties>'
  Insert Into Create_User_RDG_tmp Values(@RowID, 'C', RTrim(@sText))

  set @RowID = 1000
  Set @sText = 
        '        <group>'+@CR+
        '            <properties>'+@CR+
        '                <name>用戶PC端</name>'+@CR+
        '                <expanded>False</expanded>'+@CR+
        '                <comment />'+@CR+
        '                <logonCredentials inherit="None">'+@CR+
        '                    <userName>Administrator</userName>'+@CR+
        '                    <domain>'+@sDomain+'</domain>'+@CR+
        '                    <password storeAsClearText="True">'+@sAdmin_PWD+'</password>'+@CR+
        '                </logonCredentials>'+@CR+
        '                <connectionSettings inherit="FromParent" />'+@CR+
        '                <gatewaySettings inherit="FromParent" />'+@CR+
        '                <remoteDesktop inherit="FromParent" />'+@CR+
        '                <localResources inherit="FromParent" />'+@CR+
        '                <securitySettings inherit="FromParent" />'+@CR+
        '                <displaySettings inherit="FromParent" />'+@CR+
        '            </properties> '
        
  Insert Into Create_User_RDG_tmp Values(@RowID, 'U', RTrim(@sText))
  Insert Into Create_User_RDG_tmp Values(@RowID+2000, 'A', 
    Replace(Replace(Replace(@sText, 
        '用戶PC端', '用戶PC管理者'),
        '<domain>'+@sDomain+'</domain>', '<domain />'),
        '<password storeAsClearText="True">'+@sAdmin_PWD+'</password>', '<password storeAsClearText="True">'+@sLoca_Admin_PWD+'</password>')
    )
    
  Declare cur_tb_emp cursor for
    select DP_Name+'-'+E_No+'('+E_Name+' '+E_Duty+'), '+OfficeExt+
           Case
             when Leave = 'Y' then ', 離職'
             else ''
           end COLLATE Chinese_Taiwan_Stroke_CI_AS as Display,
           'ta'+OfficeExt as Server_Name, AccountID, EIP_PassWord
      from ori_ta13#tb_emp_level
     where EIP_PassWord is not null
     order by dp_no, e_no

  Open cur_tb_emp
  Fetch Next From cur_tb_emp Into @sServer_Name, @sDisplay, @sAccountID, @sPassWord

  While @@Fetch_Status = 0
  begin
     Set @RowID = @RowID +1
     Set @sText = 
           '            <server>'+@CR+
           '                <name>'+@sDisplay+'</name>'+@CR+
           '                <displayName>'+@sServer_Name+'</displayName>'+@CR+
           '                <comment />'+@CR+
           '                <logonCredentials inherit="None">'+@CR+
           '                    <userName>'+@sAccountID+'</userName>'+@CR+
           '                    <domain>'+@sDomain+'</domain>'+@CR+
           '                    <password storeAsClearText="True">'+@sPassWord+'</password>'+@CR+
           '                </logonCredentials>'+@CR+
           '                <connectionSettings inherit="FromParent" />'+@CR+
           '                <gatewaySettings inherit="FromParent" />'+@CR+
           '                <remoteDesktop inherit="FromParent" />'+@CR+
           '                <localResources inherit="FromParent" />'+@CR+
           '                <securitySettings inherit="FromParent" />'+@CR+
           '                <displaySettings inherit="FromParent" />'+@CR+
           '            </server>'
     Insert Into Create_User_RDG_tmp Values(@RowID, 'U', RTrim(@sText))
     Insert Into Create_User_RDG_tmp Values(@RowID+2000, 'A', 
       Replace(Replace(Replace(Replace(Replace(@sText, 
           '<logonCredentials inherit="None">', '<logonCredentials inherit="FromParent" />'),
           '                    <userName>'+@sAccountID+'</userName>'+@CR, ''),
           '                    <domain>'+@sDomain+'</domain>'+@CR, ''),
           '                    <password storeAsClearText="True">'+@sPassWord+'</password>'+@CR, ''),
           '                </logonCredentials>'+@CR,  '')
       )

     Fetch Next From cur_tb_emp Into @sServer_Name, @sDisplay, @sAccountID, @sPassWord
  end
  Close cur_tb_emp
  Deallocate cur_tb_emp

  Set @sText = 
        '        </group>'
  Insert Into Create_User_RDG_tmp Values(@RowID+1000, 'U', RTrim(@sText))
  Insert Into Create_User_RDG_tmp Values(@RowID+2000, 'A', RTrim(@sText))


  Set @RowID = 999999
  Set @sText = 
        '    </file>'+@CR+
        '</RDCMan>'
  Insert Into Create_User_RDG_tmp Values(@RowID, 'C', RTrim(@sText))
  
  Delete Create_User_RDG_tmp where RTrim(Isnull(Data, '')) = ''

  Exec uSP_Sys_Advanced_Options 1
  set @sSQL = 'bcp "select Replace(Data, ''&'', ''&amp;'') as Data from dw.dbo.Create_User_RDG_tmp order by rowid" queryout "D:\USER.rdg" -w -T -SVMS\VMSOP'
  print @sSQL
  exec master..xp_cmdshell @sSQL
  Exec uSP_Sys_Advanced_Options 0
end
GO
