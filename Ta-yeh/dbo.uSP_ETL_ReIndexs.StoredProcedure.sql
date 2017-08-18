USE [DW]
GO
/****** Object:  StoredProcedure [dbo].[uSP_ETL_ReIndexs]    Script Date: 08/18/2017 17:18:56 ******/
DROP PROCEDURE [dbo].[uSP_ETL_ReIndexs]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE Procedure [dbo].[uSP_ETL_ReIndexs]
as
begin
  Declare @Proc Varchar(50) = 'uSP_ETL_ReIndexs'
  Declare @Cnt Int =0
  Declare @Msg Varchar(Max) =''
  Declare @strSQL Varchar(Max)=''

  begin try
    if exists (select 1
                 from sysindexes
                where id = object_id('dbo.fact_pcust')
                  and name  = 'sname'
                  and indid > 0
                  and indid < 255)
       drop index dbo.fact_pcust.sname
    
    
    if exists (select 1
                 from sysindexes
                where id = object_id('dbo.fact_pcust')
                and name = 'ctname'
                and indid > 0
                and indid < 255)
       drop index dbo.fact_pcust.ctname
    
    
    if exists (select 1
                 from sysindexes
                where id = object_id('dbo.fact_pcust')
                  and name  = 'ctloc'
                  and indid > 0
                  and indid < 255)
       drop index dbo.fact_pcust.ctloc
    
    
    if exists (select 1
                 from sysindexes
                where id = object_id('dbo.fact_pcust')
                  and name  = 'ctkind'
                  and indid > 0
                  and indid < 255)
       drop index dbo.fact_pcust.ctkind
    
    
    if exists (select 1
                 from sysindexes
                where id = object_id('dbo.fact_pcust')
                  and name  = 'ctbus'
                  and indid > 0
                  and indid < 255)
       drop index dbo.fact_pcust.ctbus
    
    
    if exists (select 1
                 from sysindexes
                where id = object_id('dbo.fact_pcust')
                  and name  = 'ct_payrate'
                  and indid > 0
                  and indid < 255)
       drop index dbo.fact_pcust.ct_payrate
    
    
    if exists (select 1
                 from sysindexes
                where id = object_id('dbo.fact_pcust')
                  and name  = 'class'
                  and indid > 0
                  and indid < 255)
       drop index dbo.fact_pcust.class
    
    
    if exists (select 1
                 from sysindexes
                where id = object_id('dbo.fact_pcust')
                  and name  = 'no'
                  and indid > 0
                  and indid < 255)
       drop index dbo.fact_pcust.no
    
    
    /*==============================================================*/
    /* Index: no                                                    */
    /*==============================================================*/
    create clustered index no on dbo.fact_pcust (
    ct_no,
    ct_class
    )
    with fillfactor= 30
    on "PRIMARY"
    
    
    /*==============================================================*/
    /* Index: class                                                 */
    /*==============================================================*/
    create index class on dbo.fact_pcust (
    ct_class
    )
    with fillfactor= 30
    on "PRIMARY"
    
    
    /*==============================================================*/
    /* Index: ct_payrate                                            */
    /*==============================================================*/
    create index ct_payrate on dbo.fact_pcust (
    ct_payer,
    ct_class
    )
    with fillfactor= 30
    on "PRIMARY"
    
    
    /*==============================================================*/
    /* Index: ctbus                                                 */
    /*==============================================================*/
    create index ctbus on dbo.fact_pcust (
    ct_busine,
    ct_class
    )
    with fillfactor= 30
    on "PRIMARY"
    
    
    /*==============================================================*/
    /* Index: ctkind                                                */
    /*==============================================================*/
    create index ctkind on dbo.fact_pcust (
    ct_kind,
    ct_class
    )
    with fillfactor= 30
    on "PRIMARY"
    
    
    /*==============================================================*/
    /* Index: ctloc                                                 */
    /*==============================================================*/
    create index ctloc on dbo.fact_pcust (
    ct_loc,
    ct_class
    )
    with fillfactor= 30
    on "PRIMARY"
    
    
    /*==============================================================*/
    /* Index: ctname                                                */
    /*==============================================================*/
    create index ctname on dbo.fact_pcust (
    ct_name
    )
    with fillfactor= 30
    on "PRIMARY"
    
    
    /*==============================================================*/
    /* Index: sname                                                 */
    /*==============================================================*/
    create index sname on dbo.fact_pcust (
    ct_sname,
    ct_class
    )
    with fillfactor= 30
    on "PRIMARY"
    
    /****************************************************************************************************************************************/
    /****************************************************************************************************************************************/
    if exists (select 1
                 from sysindexes
                where id = object_id('dbo.fact_pemploy')
                  and name  = 'mstno'
                  and indid > 0
                  and indid < 255)
       drop index dbo.fact_pemploy.mstno
    
    
    if exists (select 1
                 from sysindexes
                where id = object_id('dbo.fact_pemploy')
                  and name  = 'ldate'
                  and indid > 0
                  and indid < 255)
       drop index dbo.fact_pemploy.ldate
    
    
    if exists (select 1
                 from sysindexes
                where id = object_id('dbo.fact_pemploy')
                  and name  = 'epnm'
                  and indid > 0
                  and indid < 255)
       drop index dbo.fact_pemploy.epnm
    
    
    if exists (select 1
                 from sysindexes
                where id = object_id('dbo.fact_pemploy')
                  and name  = 'dept'
                  and indid > 0
                  and indid < 255)
       drop index dbo.fact_pemploy.dept
    
    
    if exists (select 1
                 from sysindexes
                where id = object_id('dbo.fact_pemploy')
                  and name  = 'bkno'
                  and indid > 0
                  and indid < 255)
       drop index dbo.fact_pemploy.bkno
    
    
    if exists (select 1
                 from sysindexes
                where id = object_id('dbo.fact_pemploy')
                  and name  = 'epno'
                  and indid > 0
                  and indid < 255)
       drop index dbo.fact_pemploy.epno
    
    
    /*==============================================================*/
    /* Index: epno                                                  */
    /*==============================================================*/
    create clustered index epno on dbo.fact_pemploy (
    e_no
    )
    with fillfactor= 30
    on "PRIMARY"
    
    
    /*==============================================================*/
    /* Index: bkno                                                  */
    /*==============================================================*/
    create index bkno on dbo.fact_pemploy (
    e_bkno
    )
    with fillfactor= 30
    on "PRIMARY"
    
    
    /*==============================================================*/
    /* Index: dept                                                  */
    /*==============================================================*/
    create index dept on dbo.fact_pemploy (
    e_dept
    )
    with fillfactor= 30
    on "PRIMARY"
    
    
    /*==============================================================*/
    /* Index: epnm                                                  */
    /*==============================================================*/
    create index epnm on dbo.fact_pemploy (
    e_name
    )
    with fillfactor= 30
    on "PRIMARY"
    
    
    /*==============================================================*/
    /* Index: ldate                                                 */
    /*==============================================================*/
    create index ldate on dbo.fact_pemploy (
    e_ldate
    )
    with fillfactor= 30
    on "PRIMARY"
    
    
    /*==============================================================*/
    /* Index: mstno                                                 */
    /*==============================================================*/
    create index mstno on dbo.fact_pemploy (
    e_mstno
    )
    with fillfactor= 30
    on "PRIMARY"
    
    /****************************************************************************************************************************************/
    /****************************************************************************************************************************************/
    if exists (select 1
                 from sysindexes
                where id = object_id('dbo.fact_sstock')
                  and name  = 'SUPP'
                  and indid > 0
                  and indid < 255)
       drop index dbo.fact_sstock.SUPP
    
    
    if exists (select 1
                 from sysindexes
                where id = object_id('dbo.fact_sstock')
                  and name  = 'SKNM'
                  and indid > 0
                  and indid < 255)
       drop index dbo.fact_sstock.SKNM
    
    
    if exists (select 1
                 from sysindexes
                where id = object_id('dbo.fact_sstock')
                  and name  = 'SKKD'
                  and indid > 0
                  and indid < 255)
       drop index dbo.fact_sstock.SKKD
    
    
    if exists (select 1
                 from sysindexes
                where id = object_id('dbo.fact_sstock')
                  and name  = 'SKBNO'
                  and indid > 0
                  and indid < 255)
       drop index dbo.fact_sstock.SKBNO
    
    
    if exists (select 1
                 from sysindexes
                where id = object_id('dbo.fact_sstock')
                  and name  = 'SKABNO'
                  and indid > 0
                  and indid < 255)
       drop index dbo.fact_sstock.SKABNO
    
    
    if exists (select 1
                 from sysindexes
                where id = object_id('dbo.fact_sstock')
                  and name  = 'SKNO'
                  and indid > 0
                  and indid < 255)
       drop index dbo.fact_sstock.SKNO
    
    
    /*==============================================================*/
    /* Index: SKNO                                                  */
    /*==============================================================*/
    create clustered index SKNO on dbo.fact_sstock (
    sk_no
    )
    with fillfactor= 30
    on "PRIMARY"
    
    
    /*==============================================================*/
    /* Index: SKABNO                                                */
    /*==============================================================*/
    create index SKABNO on dbo.fact_sstock (
    sk_abcode,
    sk_no
    )
    with fillfactor= 30
    on "PRIMARY"
    
    
    /*==============================================================*/
    /* Index: SKBNO                                                 */
    /*==============================================================*/
    create index SKBNO on dbo.fact_sstock (
    sk_bcode,
    sk_no
    )
    with fillfactor= 30
    on "PRIMARY"
    
    
    /*==============================================================*/
    /* Index: SKKD                                                  */
    /*==============================================================*/
    create index SKKD on dbo.fact_sstock (
    sk_kind,
    sk_no
    )
    with fillfactor= 30
    on "PRIMARY"
    
    
    /*==============================================================*/
    /* Index: SKNM                                                  */
    /*==============================================================*/
    create index SKNM on dbo.fact_sstock (
    sk_name
    )
    with fillfactor= 30
    on "PRIMARY"
    
    
    /*==============================================================*/
    /* Index: SUPP                                                  */
    /*==============================================================*/
    create index SUPP on dbo.fact_sstock (
    s_supp
    )
    with fillfactor= 30
    on "PRIMARY"
    
    /****************************************************************************************************************************************/
    /****************************************************************************************************************************************/
    if exists (select 1
                 from sysindexes
                where id = object_id('dbo.fact_sslip')
                  and name  = 'spda'
                  and indid > 0
                  and indid < 255)
       drop index dbo.fact_sslip.spda
    
    
    if exists (select 1
                 from sysindexes
                where id = object_id('dbo.fact_sslip')
                  and name  = 'spctp'
                  and indid > 0
                  and indid < 255)
       drop index dbo.fact_sslip.spctp
    
    
    if exists (select 1
                 from sysindexes
                where id = object_id('dbo.fact_sslip')
                  and name  = 'slip_fg'
                  and indid > 0
                  and indid < 255)
       drop index dbo.fact_sslip.slip_fg
    
    
    if exists (select 1
                 from sysindexes
                where id = object_id('dbo.fact_sslip')
                  and name  = 'ordno'
                  and indid > 0
                  and indid < 255)
       drop index dbo.fact_sslip.ordno
    
    
    if exists (select 1
                 from sysindexes
                where id = object_id('dbo.fact_sslip')
                  and name  = 'mafno'
                  and indid > 0
                  and indid < 255)
       drop index dbo.fact_sslip.mafno
    
    
    if exists (select 1
                 from sysindexes
                where id = object_id('dbo.fact_sslip')
                  and name  = 'ctno'
                  and indid > 0
                  and indid < 255)
       drop index dbo.fact_sslip.ctno
    
    
    if exists (select 1
                 from sysindexes
                where id = object_id('dbo.fact_sslip')
                  and name  = 'class'
                  and indid > 0
                  and indid < 255)
       drop index dbo.fact_sslip.class
    
    
    if exists (select 1
                 from sysindexes
                where id = object_id('dbo.fact_sslip')
                  and name  = 'caseno'
                  and indid > 0
                  and indid < 255)
       drop index dbo.fact_sslip.caseno
    
    
    /*==============================================================*/
    /* Index: caseno                                                */
    /*==============================================================*/
    create index caseno on dbo.fact_sslip (
    sp_caseno
    )
    with fillfactor= 30
    on "PRIMARY"
    
    
    /*==============================================================*/
    /* Index: class                                                 */
    /*==============================================================*/
    create index class on dbo.fact_sslip (
    sp_class
    )
    with fillfactor= 30
    on "PRIMARY"
    
    
    /*==============================================================*/
    /* Index: ctno                                                  */
    /*==============================================================*/
    create index ctno on dbo.fact_sslip (
    sp_ctno,
    sp_class
    )
    with fillfactor= 30
    on "PRIMARY"
    
    
    /*==============================================================*/
    /* Index: mafno                                                 */
    /*==============================================================*/
    create index mafno on dbo.fact_sslip (
    sp_mafno,
    sp_mafkd
    )
    with fillfactor= 30
    on "PRIMARY"
    
    
    /*==============================================================*/
    /* Index: ordno                                                 */
    /*==============================================================*/
    create index ordno on dbo.fact_sslip (
    sp_ordno,
    sp_class
    )
    with fillfactor= 30
    on "PRIMARY"
    
    
    /*==============================================================*/
    /* Index: slip_fg                                               */
    /*==============================================================*/
    create index slip_fg on dbo.fact_sslip (
    sp_slip_fg
    )
    with fillfactor= 30
    on "PRIMARY"
    
    
    /*==============================================================*/
    /* Index: spctp                                                 */
    /*==============================================================*/
    create index spctp on dbo.fact_sslip (
    sp_pdate,
    sp_ctno,
    sp_class
    )
    with fillfactor= 30
    on "PRIMARY"
    
    
    /*==============================================================*/
    /* Index: spda                                                  */
    /*==============================================================*/
    create index spda on dbo.fact_sslip (
    sp_date,
    sp_slip_fg
    )
    with fillfactor= 30
    on "PRIMARY"
    
    
    /****************************************************************************************************************************************/
    /****************************************************************************************************************************************/
    if exists (select 1
                 from sysindexes
                where id = object_id('dbo.fact_sslpdt')
                  and name  = 'slip_fg'
                  and indid > 0
                  and indid < 255)
       drop index dbo.fact_sslpdt.slip_fg
    
    
    if exists (select 1
                 from sysindexes
                where id = object_id('dbo.fact_sslpdt')
                  and name  = 'skno'
                  and indid > 0
                  and indid < 255)
       drop index dbo.fact_sslpdt.skno
    
    
    if exists (select 1
                 from sysindexes
                where id = object_id('dbo.fact_sslpdt')
                  and name  = 'skda'
                  and indid > 0
                  and indid < 255)
       drop index dbo.fact_sslpdt.skda
    
    
    if exists (select 1
                 from sysindexes
                where id = object_id('dbo.fact_sslpdt')
                  and name  = 'sdda'
                  and indid > 0
                  and indid < 255)
       drop index dbo.fact_sslpdt.sdda
    
    
    if exists (select 1
                 from sysindexes
                where id = object_id('dbo.fact_sslpdt')
                  and name  = 'ordno'
                  and indid > 0
                  and indid < 255)
       drop index dbo.fact_sslpdt.ordno
    
    
    if exists (select 1
                 from sysindexes
                where id = object_id('dbo.fact_sslpdt')
                  and name  = 'ctno'
                  and indid > 0
                  and indid < 255)
       drop index dbo.fact_sslpdt.ctno
    
    
    if exists (select 1
                 from sysindexes
                where id = object_id('dbo.fact_sslpdt')
                  and name  = 'CSNO'
                  and indid > 0
                  and indid < 255)
       drop index dbo.fact_sslpdt.CSNO
    
    
    if exists (select 1
                 from sysindexes
                where id = object_id('dbo.fact_sslpdt')
                  and name  = 'class'
                  and indid > 0
                  and indid < 255)
       drop index dbo.fact_sslpdt.class
    
    
    if exists (select 1
                 from sysindexes
                where id = object_id('dbo.fact_sslpdt')
                  and name  = 'sdno'
                  and indid > 0
                  and indid < 255)
       drop index dbo.fact_sslpdt.sdno
    
    
    /*==============================================================*/
    /* Index: sdno                                                  */
    /*==============================================================*/
    create clustered index sdno on dbo.fact_sslpdt (
    sd_no,
    sd_slip_fg
    )
    with fillfactor= 30
    on "PRIMARY"
    
    
    /*==============================================================*/
    /* Index: class                                                 */
    /*==============================================================*/
    create index class on dbo.fact_sslpdt (
    sd_class
    )
    with fillfactor= 30
    on "PRIMARY"
    
    
    /*==============================================================*/
    /* Index: CSNO                                                  */
    /*==============================================================*/
    create index CSNO on dbo.fact_sslpdt (
    sd_csno
    )
    with fillfactor= 30
    on "PRIMARY"
    
    
    /*==============================================================*/
    /* Index: ctno                                                  */
    /*==============================================================*/
    create index ctno on dbo.fact_sslpdt (
    sd_ctno,
    sd_class
    )
    with fillfactor= 30
    on "PRIMARY"
    
    
    /*==============================================================*/
    /* Index: ordno                                                 */
    /*==============================================================*/
    create index ordno on dbo.fact_sslpdt (
    sd_ordno,
    sd_class
    )
    with fillfactor= 30
    on "PRIMARY"
    
    
    /*==============================================================*/
    /* Index: sdda                                                  */
    /*==============================================================*/
    create index sdda on dbo.fact_sslpdt (
    sd_date,
    sd_class
    )
    with fillfactor= 30
    on "PRIMARY"
    
    
    /*==============================================================*/
    /* Index: skda                                                  */
    /*==============================================================*/
    create index skda on dbo.fact_sslpdt (
    sd_skno,
    sd_date
    )
    with fillfactor= 30
    on "PRIMARY"
    
    
    /*==============================================================*/
    /* Index: skno                                                  */
    /*==============================================================*/
    create index skno on dbo.fact_sslpdt (
    sd_skno,
    sd_ctno,
    sd_date,
    sd_class
    )
    with fillfactor= 30
    on "PRIMARY"
    
    
    /*==============================================================*/
    /* Index: slip_fg                                               */
    /*==============================================================*/
    create index slip_fg on dbo.fact_sslpdt (
    sd_slip_fg
    )
    with fillfactor= 30
    on "PRIMARY"

    /*==============================================================*/
    /* Index: IX_sslpdt_N1                                          */
    /*==============================================================*/
    CREATE NONCLUSTERED INDEX [IX_sslpdt_N1]  ON [dbo].[Fact_sslpdt] (
    [Chg_skno_BKind],[sd_slip_fg],[Chg_sp_pdate_Year]
    )
    INCLUDE ([Chg_sales_Name],[Chg_sd_Pay_Sale],[Chg_sp_pdate_Month],[sp_sales])

    Exec uSP_Sys_Write_Log @Proc, @Msg, @Msg, @Cnt
  end try
  begin catch
    set @Cnt = -1
    set @Msg = '(¿ù»~°T®§:'+ERROR_MESSAGE()+')'

    Exec uSP_Sys_Write_Log @Proc, @Msg, @Msg, @Cnt
  end catch
end
GO
