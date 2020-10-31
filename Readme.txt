'---------------------------------------------------------------------------------------
' Module    : ADOHelper.bas
' DateTime  : 4/10/2002 14:10
' Author    : Avaneesh Dvivedi
'           : http://www.tax-publishers.com/advivedi
' Purpose   :This module demonstrates how to perform common operations
'            with ADO. This bas is essential if you want to do any
'            any database access with ADO. It shows how to perform common
'            task as well as complex task using ADO. The notes on top of
'            of each function is self explanatory.
'            The important feature is that it shows comparison with
'            DAO (the older database access technology)and instructs
'            how to perform similar task in ADO and which parameter to use.
'            A DAO bas is also included to show the comparison and to show
'            how to perform similar task.
'
'            I am a chartered accountant based in India. I program for fun
'            you can check out the latest version of this bas at my
'            web site at http://www.tax-publishers.com/advivedi
'
'The functions included are as under:
'
'DATABASE OPENING

'ADOOpenJetDatabase()
'ADOOpenJetDatabaseReadOnly()
'Sub ADOOpenJetDatabaseExclusive()
'Sub ADOSetJetDBOption()
'Sub ADOOpenDBPasswordDatabase()
'Sub ADOOpenSecuredDatabase()
'Sub ADOOpenISAMDatabase()
'Sub ADOGetCurrentDatabase()
'Sub ADOOpenJetDatabaseExclusive()

'RECORDSET OPERATIONS

'Sub ADOOpenRecordset()
'Sub ADOMoveNext()
'Sub ADOGetCurrentPosition()
'ADDING, EDITING, DELETING RECORD AND FIND,SEEK
'Sub ADOFindRecord()
'Sub ADOSeekRecord()
'Sub ADOFilterRecordset()
'Sub ADOSortRecordset()
'Sub ADOAddRecord()
'Sub ADOAddRecord2()
'Sub ADOUpdateRecord()
'Sub ADOReadMemo()
'Sub ADOUpdateBLOB()

'CREATING QUERY, USING QUERY AND PARAMETERS

'Sub ADOExecuteQuery()
'Sub ADOExecuteParamQuery()
'Sub ADOExecuteParamQuery2()
'Sub ADOExecuteBulkOpQuery()

'DATABASE MAINENANCE AND CREATING NEW DATABASE

'Sub ADOCreateDatabase()
'Sub ADOListTables()
'Sub ADOListTables2()
'Sub ADOCreateTable()
'Sub ADOCreateAttachedJetTable()
'Sub ADOCreateAttachedODBCTable()
'Sub ADOCreateAutoIncrColumn()
'Sub ADORefreshLinks()
'Sub ADOCreateIndex()
'Sub ADOCreatePrimaryKey()
'Sub ADOCreateForeignKey()
'Sub ADOCreateForeignKeyCascade()
'Sub ADOCreateQuery()
'Sub ADOCreateParameterizedQuery()
'Sub ADOModifyQuery()
'Sub ADOCreateSQLPassThrough()

'CHANGING PASSWORD, COMPACTING DATABASE AND ENCRYPTING

'Sub ADOChangePassword()
'Sub JROChangeDatabasePassword()

'CREATING WORKGROUP, USER, SETTING PERMISSIONS ETC.

'Sub ADOCreateUser()
'Sub ADOAddUserToNewGroup()
'Sub ADOSetUserObjectPermissions()
'Sub ADOSetDatabasePermissions()
'Sub ADOSetUserContainerPermissions()
'Sub ADOGetObjectOwner()

'USING JRO

'Sub JROMakeDesignMaster()
'Sub JROKeepObjectLocal()
'Sub JROCreatePartial()
'Sub JROListFilters()
'Sub JROTwoWayDirectSync()
'Sub JROInternetSync()
'Sub JROConflictTables()
'Sub ADODatabaseError()
'Sub ADOTransactions()
'Sub JROCompactDatabase()
'Sub JROEncryptDatabase()
'Sub JRORefreshCache()
'Sub ADOCreateRecordset()
'Sub ADOUseExistingDataLink()
'Sub ADOCreateEnhancedAutoIncrColumn()
'Sub JROTwoWayIndirectSync()
'Sub JROJetSQLSync()
'Sub JROMakeDesignMaster2()
'
'You can send your comments to me at a_dvivedi@hotmail.com
'---------------------------------------------------------------------------------------