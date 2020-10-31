Attribute VB_Name = "ADO"
'---------------------------------------------------------------------------------------
' Module    : ADO
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
'            how to perform similar task in ADO and which parameter to use
'            I am a chartered accountant based in India. I program for fun
'            you can check out the latest version of this bas at my
'            web site at http://www.tax-publishers.com/advivedi
'            The functions included are as under:
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
'
'---------------------------------------------------------------------------------------
Option Explicit

Sub ADOOpenJetDatabase()

   Dim cnn As New ADODB.Connection

   cnn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
      "Data Source=.\NorthWind.mdb;"
   cnn.Close

End Sub

'The following code listings show how to open (and then close) a shared, read-only database using DAO and ADO.



Sub ADOOpenJetDatabaseReadOnly()

   Dim cnn As New ADODB.Connection

   ' Open shared, read-only
   cnn.Mode = adModeRead
   cnn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
      "Data Source=.\NorthWind.mdb;"
   cnn.Close

End Sub
'Alternatively, the ADO listing could have been written in a single line of code as follows:


Sub ADOOpenJetDatabaseExclusive()

   Dim cnn As New ADODB.Connection

   cnn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
      "Data Source=.\NorthWind.mdb;Mode=" & adModeRead
   cnn.Close
'In this listing, the Mode property was specified as a part of the connection string to the Open method rather than as a property of the Connection object. In ADO, you can set connection properties as a property or string them together with other properties to create the connection string. Even provider-specific properties (prefixed by "Jet OLEDB:" for Microsoft Jet–specific properties) can be set as part of the connection string or with the Connection object's Properties collection
End Sub
'The following listings demonstrate how to override the Page Timeout setting of the engine and open a database using that setting.

Sub ADOSetJetDBOption()

   Dim cnn As New ADODB.Connection

   cnn.Provider = "Microsoft.Jet.OLEDB.4.0;"
   cnn.Open ".\NorthWind.mdb"
   cnn.Properties("Jet OLEDB:Page Timeout") = 4000
   cnn.Close


'The following table lists the values that can be set with DAO's SetOption method and the corresponding property to use with ADO.

'DAO constant ADO property
'dbPageTimeout Jet OLEDB:Page Timeout
'dbSharedAsyncDelay Jet OLEDB:Shared Async Delay
'dbExclusiveAsyncDelay Jet OLEDB:Exclusive Async Delay
'dbLockRetry Jet OLEDB:Lock Retry
'dbUserCommitSync Jet OLEDB:User Commit Sync
'dbImplicitCommitSync Jet OLEDB:Implicit Commit Sync
'dbMaxBufferSize Jet OLEDB:Max Buffer Size
'dbMaxLocksPerFile Jet OLEDB:Max Locks Per File
'dbLockDelay Jet OLEDB:Lock Delay
'dbRecycleLVs Jet OLEDB:Recycle Long-Valued Pages
'dbFlushTransactionTimeout Jet OLEDB:Flush Transaction Timeout

End Sub

Sub ADOOpenDBPasswordDatabase()

'Share-Level (Password Protected) Databases
'The following listings demonstrate how to open a Microsoft Jet database that has been secured at the share level.
'DAO

   Dim cnn As New ADODB.Connection

   cnn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
      "Data Source=.\NorthWind.mdb;" & _
      "Jet OLEDB:Database Password=password;"
   cnn.Close
'In DAO, the Connect parameter of the OpenDatabase method sets the database password when opening a database. With ADO, the Microsoft Jet Provider connection property Jet OLEDB:Database Password sets the password instead
End Sub


'Opening a Database with User-Level Security
'These next listings demonstrate how to open a database that is secured at the user level using a workgroup information file named "system.mdw".
'DAO

Sub ADOOpenSecuredDatabase()

   Dim cnn As New ADODB.Connection

   cnn.Provider = "Microsoft.Jet.OLEDB.4.0;"
   cnn.Properties("Jet OLEDB:System database") = _
      "C:\Program Files\Microsoft Office\Office\SYSTEM.MDW"

   cnn.Open "Data Source=.\NorthWind.mdb;User Id=Admin;Password=;"
   cnn.Close

End Sub


'External Databases
'The Microsoft Jet database engine can be used to access other database files, spreadsheets, and textual data stored in tabular format through installable ISAM drivers.
'The following listings demonstrate how to open a Microsoft Excel 2000 spreadsheet first using DAO, then using ADO and the Microsoft Jet provider.
'DAO

Sub ADOOpenISAMDatabase()

   Dim cnn As New ADODB.Connection

   cnn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
      "Data Source=.\Sales.xls" & _
      ";Extended Properties=Excel 8.0;"

   cnn.Close
'The DAO and ADO code for opening an external database is similar. In both examples, the name of the external file (Sales.xls) is used in place of a Microsoft Jet database file name. With both DAO and ADO you must also specify the type of external database you are opening, in this case, an Excel 2000 spreadsheet. With DAO, the database type is specified in the Connect argument of the OpenDatabase method. The database type is specified in the Extended Properties property of the Connection with ADO. The following table lists the strings to use to specify which ISAM to open.
'Database String
'dBASE III dBASE III;
'dBASE IV dBASE IV;
'dBASE 5 dBASE 5.0;
'Paradox 3.x Paradox 3.x;
'Paradox 4.x Paradox 4.x;
'Paradox 5.x Paradox 5.x;
'Excel 3.0 Excel 3.0;
'Excel 4.0 Excel 4.0;
'Excel 5.0/Excel 95 Excel 5.0;
'Excel 97 Excel 97;
'Excel 2000 Excel 8.0;
'HTML Import HTML Import;
'HTML Export HTML Export;
'Text Text;
'ODBC ODBC;
'DATABASE=database;
'UID=user;
'PWD=password;
'DSN = DataSourceName

End Sub

'The Current Microsoft Access Database
'When you open Microsoft Access, you are opening a Microsoft Jet database. When writing code within Access, you may often want to use the same connection to Microsoft Jet as Access is using. To allow you to do this, Microsoft Access 2000 exposes two mechanisms: CurrentDB() and CurrentProject.Connection allow you to get a DAO Database object and an ADO Connection object, respectively, for the database Access currently has open.zdatabase currently open in Microsoft Access.

Sub ADOGetCurrentDatabase()

   Dim cnn As ADODB.Connection

   Set cnn = CurrentProject.Connection

End Sub

'Alternatively, the ADO listing could have been written in a single line of code as follows:




Sub ADOOpenRecordset()

   Dim cnn As New ADODB.Connection
   Dim rst As New ADODB.Recordset
   Dim fld As ADODB.Field

   ' Open the connection
   cnn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
      "Data Source=.\NorthWind.mdb;"

   ' Open the forward-only,
   ' read-only recordset
   rst.Open _
      "SELECT * FROM Customers WHERE Region = 'WA'", _
      cnn, adOpenForwardOnly, adLockReadOnly

   ' Print the values for the fields in
   ' the first record in the debug window
   For Each fld In rst.Fields
      Debug.Print fld.Value & ";";
   Next

   Debug.Print

   ' Close the recordset
   rst.Close

End Sub



Sub ADOMoveNext()

   Dim cnn As New ADODB.Connection
   Dim rst As New ADODB.Recordset
   Dim fld As ADODB.Field

   ' Open the connection
   cnn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" _
      & "Data Source=.\NorthWind.mdb;"

   ' Open the forward-only,
   ' read-only recordset
   rst.Open _
      "SELECT * FROM Customers WHERE Region = 'WA'", _
      cnn, adOpenForwardOnly, adLockReadOnly

   ' Print the values for the fields in
   ' the first record in the debug window
   Do Until rst.EOF
      For Each fld In rst.Fields
         Debug.Print fld.Value & ";";
      Next
      Debug.Print
      rst.MoveNext
   Loop

   ' Close the recordset
   rst.Close

End Sub

Sub ADOGetCurrentPosition()

   Dim cnn As New ADODB.Connection
   Dim rst As New ADODB.Recordset

   ' Open the connection
   cnn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
      "Data Source=.\NorthWind.mdb;"

   ' Open the recordset
   rst.CursorLocation = adUseClient
   rst.Open "SELECT * FROM Customers", cnn, adOpenKeyset, _
      adLockOptimistic, adCmdText

   ' Print the absolute position
   Debug.Print rst.AbsolutePosition

   ' Move to the last record
   rst.MoveLast

   ' Print the absolute position
   Debug.Print rst.AbsolutePosition

   ' Close the recordset
   rst.Close

End Sub

Sub ADOFindRecord()

   Dim cnn As New ADODB.Connection
   Dim rst As New ADODB.Recordset

   ' Open the connection
   cnn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
      "Data Source=.\NorthWind.mdb;"

   ' Open the recordset
   rst.Open "Customers", cnn, adOpenKeyset, adLockOptimistic

   ' Find the first customer whose country is USA
   rst.Find "Country='USA'"

   ' Print the customer id's of all customers in the USA
   Do Until rst.EOF
      Debug.Print rst.Fields("CustomerId").Value
      rst.Find "Country='USA'", 1
   Loop

   ' Close the recordset
   rst.Close

'------------------------------------------------------------------------
'DAO includes four find methods: FindFirst, FindLast, FindNext, FindPrevious. You choose which method to use based on the point from which you want to start searching (beginning, end, or curent record) and in which direction you want to search (forward or backward).
'ADO has a single method: Find. Searching always begins from the current record. The Find method has parameters that allow you to specify the search direction as well as an offset from the current record at which to beginning searching (SkipRows). The following table shows how to map the four DAO methods to the equivalent functionality in ADO.
'DAO method ADO Find with SkipRows  ADO search direction
'FindFirst 0 adSearchForward (if not currently positioned on the first record, call MoveFirst before Find)
'FindLast 0 adSearchBackward (if not currently positioned on the last record, call MoveLast before Find)
'FindNext 1 adSearchForward
'FindPrevious 1 adSearchBackward

'DAO and ADO require a different syntax for locating records based on a Null value. In DAO if you want to find a record that has a Null value you use the following syntax:
'"ColumnName Is Null"
'or, to find a record that does not have a Null value for that column:
'"ColumnName Is Not Null"
'ADO, however, does not recognize the Is operator. You must use the = or <> operators instead. So the equivalent ADO criteria would be:
'"ColumnName = Null"
'or
'"ColumnName <> Null"
'-------------------------------------------------------------------------

End Sub



Sub ADOSeekRecord()

   Dim cnn As New ADODB.Connection
   Dim rst As New ADODB.Recordset

   ' Open the connection
   cnn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
      "Data Source=.\NorthWind.mdb;"

   ' Open the recordset
   rst.Open "Order Details", cnn, adOpenKeyset, adLockReadOnly, _
      adCmdTableDirect

   ' Select the index used to order the data in the recordset
   rst.Index = "PrimaryKey"

   ' Find the order where OrderId = 10255 and ProductId = 16
   rst.Seek Array(10255, 16), adSeekFirstEQ

   ' If a match is found print the quantity of the order
   If Not rst.EOF Then
   Debug.Print rst.Fields("Quantity").Value
   End If

   ' Close the recordset
   rst.Close

End Sub

Sub ADOFilterRecordset()

   Dim cnn As New ADODB.Connection
   Dim rst As New ADODB.Recordset

   ' Open the connection
   cnn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
      "Data Source=.\NorthWind.mdb;"

   ' Open the recordset
   rst.Open "Customers", cnn, adOpenKeyset, adLockOptimistic

   ' Filter the recordset to include only those customers in
   ' the USA that have a fax number
   rst.Filter = "Country='USA' And Fax <> Null"
   Debug.Print rst.Fields("CustomerId").Value

   ' Close the recordset
   rst.Close

'The DAO and ADO Filter properties are used slightly differently. With DAO, the Filter property specifies a filter to be applied to any subsequently opened Recordset objects based on the Recordset for which you have applied the filter. With ADO, the Filter property applies to the Recordset to which you applied the filter. The ADO Filter property allows you to create a temporary view that can be used to locate a particular record or set of records within the Recordset. When a filter is applied to the Recordset, the RecordCount property reflects just the number of records within the view. The filter can be removed by setting the Filter property to adFilterNone.

End Sub


Sub ADOSortRecordset()

   Dim cnn As New ADODB.Connection
   Dim rst As New ADODB.Recordset

   ' Open the connection
   cnn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
      "Data Source=.\NorthWind.mdb;"

   ' Open the recordset
   rst.CursorLocation = adUseClient
   rst.Open "Customers", cnn, adOpenKeyset, adLockOptimistic

   ' Sort the recordset based on Country and Region both in
   ' ascending order
   rst.Sort = "Country, Region"
   Debug.Print rst.Fields("CustomerId").Value

   ' Close the recordset
   rst.Close
'-------------------------------------------------------------------
'Like the Filter property, the DAO and ADO Sort properties differ in that the DAO Sort applies to subsequently opened Recordset objects, and for ADO it applies to the current Recordset.

'Note that the Microsoft Jet Provider does not support the OLE DB interfaces that ADO could use to filter and sort the Recordset (IViewFilter and IViewSort). In the case of Filter, ADO will perform the filter itself. However, for Sort, you must use the Cursor Service by specifying adUseClient for the CursorLocation property prior to opening the Recordset. The Cursor Service will copy all of the records in the Recordset to a cache on your local machine and will build temporary indexes in order to perform the sorting. In many cases, you may achieve better performance by re-executing the query used to open the Recordset and specifying an SQL WHERE or ORDER BY clause as appropriate.

'Also, you may not get identical results with DAO and ADO when sorting Recordset objects. Different sort algorithms can create different sequences for records that have equal values in the sorted fields. In the example above, the DAO code gives 'RANCH' as the CustomerId for the first record, while the ADO code gives 'CACTU' as the CustomerId. Both results are valid.

'-------------------------------------------------------------------
End Sub
Sub ADOAddRecord()

   Dim cnn As New ADODB.Connection
   Dim rst As New ADODB.Recordset

   ' Open the connection
   cnn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
      "Data Source=.\NorthWind.mdb;"

   ' Open the recordset
   rst.Open "SELECT * FROM Customers", _
      cnn, adOpenKeyset, adLockOptimistic

   ' Add a new record
   rst.AddNew

   ' Specify the values for the fields
   rst!CustomerId = "HENRY"
   rst!CompanyName = "Henry's Chop House"
   rst!ContactName = "Mark Henry"
   rst!ContactTitle = "Sales Representative"
   rst!Address = "40178 NE 8th Street"
   rst!City = "Bellevue"
   rst!Region = "WA"
   rst!PostalCode = "98107"
   rst!Country = "USA"
   rst!Phone = "(425) 555-9876"
   rst!Fax = "(425) 555-8908"

   ' Save the changes you made to the
   ' current record in the Recordset
   rst.Update

   ' For this example, just print out
   ' CustomerId for the new record
   Debug.Print rst!CustomerId

   ' Close the recordset
   rst.Close

End Sub

'DAO and ADO behave differently when a new record is added. With DAO, the record that was current before you used AddNew remains current. With ADO, the newly inserted record becomes the current record. Because of this, it is not necessary to explicitly reposition on the new record to get information such as the value of an auto-increment column for the new record. For this reason, in the ADO example above, there is no equivalent code to the rst.Bookmark = rst.LastModified code found in the DAO example.

'ADO also provides a shortcut syntax for adding new records. The AddNew method has two optional parameters, FieldList and Values, that take an array of field names and field values respectively. The following example demonstrates how to use the shortcut syntax.

Sub ADOAddRecord2()

   Dim cnn As New ADODB.Connection
   Dim rst As New ADODB.Recordset

   ' Open the connection
   cnn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
      "Data Source=.\NorthWind.mdb;"

   ' Open the recordset
   rst.Open "SELECT * FROM Shippers", _
      cnn, adOpenKeyset, adLockOptimistic

   ' Add a new record
   rst.AddNew Array("CompanyName", "Phone"), _
      Array("World Express", "(425) 555-7863")

   ' Save the changes you made to the
   ' current record in the Recordset
   rst.Update

   ' For this example, just print out the
   ' ShipperId for the new row.
   Debug.Print rst!ShipperId

   ' Close the recordset
   rst.Close

End Sub

Sub ADOUpdateRecord()

   Dim cnn As New ADODB.Connection
   Dim rst As New ADODB.Recordset

   ' Open the connection
   cnn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" _
      & "Data Source=.\NorthWind.mdb;"

   ' Open the recordset
   rst.Open _
      "SELECT * FROM Customers WHERE CustomerId = 'LAZYK'", _
      cnn, adOpenKeyset, adLockOptimistic

   ' Update the Contact name of the
   ' first record
   rst.Fields("ContactName").Value = "New Name"

   ' Save the changes you made to the
   ' current record in the Recordset
   rst.Update

   ' Close the recordset
   rst.Close
'
'Alternatively, in both the DAO and ADO code examples, the explicit syntax
'rst.Fields("ContactName").Value = "New Name"
'can be shortened to
'rst!ContactName = "New Name"
'
End Sub

Sub ADOReadMemo()

   Dim cnn As New ADODB.Connection
   Dim rst As New ADODB.Recordset
   Dim sNotes As String
   Dim sChunk As String
   Dim cchChunkReceived As Long
   Dim cchChunkRequested As Long

   ' Open the connection
   cnn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
      "Data Source=.\NorthWind.mdb;"

   ' Open the recordset
   rst.Open "SELECT Notes FROM Employees ", _
      cnn, adOpenKeyset, adLockOptimistic

   ' cchChunkRequested artifically set low at 16
   ' to demonstrate looping
   cchChunkRequested = 16

   ' Loop through as many chunks as it takes
   ' to read the entire BLOB into memory
   Do
      ' Temporarily store the next chunk
      sChunk = rst.Fields("Notes").GetChunk(cchChunkRequested)

      ' Check how much we got
      cchChunkReceived = Len(sChunk)

      ' If we got anything,
      ' concatenate it to the main BLOB
      If cchChunkReceived > 0 Then
         sNotes = sNotes & sChunk
      End If

   Loop While cchChunkReceived = cchChunkRequested

   ' For this example, print the value of
   ' the Notes field for just the first record
   Debug.Print sNotes

   ' Close the recordset
   rst.Close

End Sub

'The following listings demonstrate how to update binary data in an OLE object field without using the GetChunk or AppendChunk methods.
Sub ADOUpdateBLOB()

   Dim cnn As New ADODB.Connection
   Dim rst As New ADODB.Recordset
   Dim rgPhoto() As Byte

   ' Open the connection
   cnn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
      "Data Source=.\NorthWind.mdb;"

   ' Open the recordset
   rst.Open "SELECT Photo FROM Employees ", _
      cnn, adOpenKeyset, adLockOptimistic

   ' Get the first photo
   rgPhoto = rst.Fields("Photo").Value

   ' Move to the next record
   rst.MoveNext

   ' Copy the photo into the next record
   rst.Fields("Photo").Value = rgPhoto

   ' Save the changes you made to the
   ' current record in the Recordset
   rst.Update

   ' Close the recordset
   rst.Close

End Sub

'Executing a Non-Parameterized Stored Query
'A non-parameterized stored query is an SQL statement that has been saved in the database and does not require that additional variable information be specified in order to execute. The following listings demonstrate how to execute such a query.

Sub ADOExecuteQuery()

   Dim cnn As New ADODB.Connection
   Dim rst As New ADODB.Recordset
   Dim fld As ADODB.Field

   ' Open the connection
   cnn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
      "Data Source=.\NorthWind.mdb;"

   ' Open the recordset
   rst.Open "[Products Above Average Price]", _
      cnn, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc

   ' Display the records in the
   ' debug window
   Do Until rst.EOF
      For Each fld In rst.Fields
         Debug.Print fld.Value & ";";
      Next
      Debug.Print
      rst.MoveNext
   Loop

   ' Close the recordset
   rst.Close
'The code for executing a non-parameterized, row-returning query is almost identical. With ADO, if the query name contains spaces you must use square brackets ([ ]) around the name.
End Sub





Sub ADOExecuteParamQuery()

   Dim cnn As New ADODB.Connection
   Dim cat As New ADOX.Catalog
   Dim cmd As ADODB.Command
   Dim rst As New ADODB.Recordset
   Dim fld As ADODB.Field

   ' Open the connection
   cnn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
      "Data Source=.\NorthWind.mdb;"

   ' Open the catalog
   cat.ActiveConnection = cnn

   ' Get the Command object from the
   ' Procedure
   Set cmd = cat.Procedures("Sales by Year").Command

   ' Specify the parameter values
   cmd.Parameters _
      ("Forms![Sales by Year Dialog]!BeginningDate") = #8/1/1997#
   cmd.Parameters _
      ("Forms![Sales by Year Dialog]!EndingDate") = #8/31/1997#

   ' Open the recordset
   rst.Open cmd, , adOpenForwardOnly, _
      adLockReadOnly, adCmdStoredProc

   ' Display the records in the
   ' debug window
   Do Until rst.EOF
      For Each fld In rst.Fields
         Debug.Print fld.Value & ";";
      Next
      Debug.Print
      rst.MoveNext
   Loop

   ' Close the recordset
   rst.Close

End Sub

'Alternatively, the ADO example could be written more concisely by specifying the parameter values using the Parameters parameter with the Command object's Execute method. The following lines of code:
'   ' Specify the parameter values
'   cmd.Parameters _
'      ("Forms![Sales by Year Dialog]!BeginningDate") = #8/1/1997#
'   cmd.Parameters _
'      ("Forms![Sales by Year Dialog]!EndingDate") = #8/31/1997#
'   ' Open the recordset
'   rst.Open cmd, , adOpenForwardOnly, _
'      adLockReadOnly, adCmdStoredProc
'could be replaced by the single line:
'   ' Execute the Command, passing in the
'   ' values for the parameters
'   Set rst = cmd.Execute(, Array(#8/1/1997#, #8/31/1997#), _
'      adCmdStoredProc)
'
'In one more variation of the ADO code to execute a parameterized query, the example could be rewritten to not use any ADOX code.

Sub ADOExecuteParamQuery2()

   Dim cnn As New ADODB.Connection
   Dim cmd As New ADODB.Command
   Dim rst As New ADODB.Recordset
   Dim fld As ADODB.Field

   ' Open the connection
   cnn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
      "Data Source=.\NorthWind.mdb;"

   ' Create the command
   Set cmd.ActiveConnection = cnn
   cmd.CommandText = "[Sales by Year]"

   ' Execute the Command, passing in the
   ' values for the parameters
   Set rst = cmd.Execute(, Array(#8/1/1997#, #8/31/1997#), _
      adCmdStoredProc)

   ' Display the records in the
   ' debug window
   Do Until rst.EOF
      For Each fld In rst.Fields
         Debug.Print fld.Value & ";";
      Next
      Debug.Print
      rst.MoveNext
   Loop

   ' Close the recordset
   rst.Close

End Sub

'Executing Bulk Operations
'The ADO Command object's Execute method can be used for row-returning queries, as shown in the previous section, as well as for non row-returning queries—also known as bulk operations. The following code examples demonstrate how to execute a bulk operation in both DAO and ADO.

Sub ADOExecuteBulkOpQuery()

   Dim cnn As New ADODB.Connection
   Dim iAffected As Integer

   ' Open the connection
   cnn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
      "Data Source=.\NorthWind.mdb;"

   ' Execute the query
   cnn.Execute "UPDATE Customers SET Country = 'United States' " & _
      "WHERE Country = 'USA'", iAffected, adExecuteNoRecords

   Debug.Print "Records Affected = " & iAffected

   ' Close the connection
   cnn.Close

'Unlike DAO, which has two methods for executing SQL statements, OpenRecordset and Execute, ADO has a single method, Execute, that executes row-returning as well as bulk operations. In the ADO example, the constant adExecuteNoRecords indicates that the SQL statement is non row-returning. If this constant is omitted, the ADO code will still execute successfully, but you will pay a performance penalty. When adExecuteNoRecords is not specified, ADO will create a Recordset object as the return value for the Execute method. Creating this object is unnecessary overhead if the statement does not return records and should be avoided by specifying adExecuteNoRecords when you know that the statement is non row-returning.

End Sub


Sub ADOCreateDatabase()

   Dim cat As New ADOX.Catalog

   cat.Create "Provider=Microsoft.Jet.OLEDB.4.0;" & _
      "Data Source=.\New.mdb;"
'--------------------------------------------------------------------------------
'In ADO, encryption and database version information is specified by provider-specific properties. With the Microsoft Jet Provider, use the Encrypt Database and Engine Type properties, respectively. The following line of code specifies these values in the connection string to create an encrypted, version 1.1 Microsoft Jet database:
'   cat.Create "Provider=Microsoft.Jet.OLEDB.4.0;" & _
'      "Data Source=.\New.mdb;" & _
'      "Jet OLEDB:Encrypt Database=True;" & _
'      "Jet OLEDB:Engine Type=2;"
'--------------------------------------------------------------------------------
End Sub





Sub ADOListTables()

   Dim cat As New ADOX.Catalog
   Dim tbl As ADOX.Table

   ' Open the catalog
   cat.ActiveConnection = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
      "Data Source=.\NorthWind.mdb;"

   ' Loop through the tables in the database and print their name
   For Each tbl In cat.Tables
      If tbl.Type <> "VIEW" Then Debug.Print tbl.Name
   Next

'With DAO, the TableDef object represents a table in the database and the TableDefs collection contains a TableDef object for each table in the database. This is similar to ADO, in which the Table object represents a table and the Tables collection contains all the tables.
'However, unlike DAO, the ADO Tables collection may contain Table objects that aren't actual tables in your Microsoft Jet database. For example, row-returning, non-parameterized Microsoft Jet queries (considered Views in ADO) are also included in the Tables collection. To determine whether or not the Table object represents a table in the database, use the Type property. The following table lists the possible values for the Type property when using ADO with the Microsoft Jet Provider.
'Type Description
'ACCESS TABLE The Table is a Microsoft Access system table.
'LINK The Table is a linked table from a non-ODBC data source.
'PASS-THROUGH The Table is a linked table from an ODBC data source.
'SYSTEM TABLE The Table is a Microsoft Jet system table.
'TABLE The Table is a table.
'VIEW The Table is a row-returning, non-parameterized query.

End Sub

'In general, it is faster to use the OpenSchema method rather than looping through the collection, because ADOX must incur the overhead of creating objects for each element in the collection. The following code demonstrates how to use the OpenSchema method to print the same information as the previous DAO and ADOX examples.

Sub ADOListTables2()

   Dim cnn As New ADODB.Connection
   Dim rst As ADODB.Recordset

   ' Open the connection
   cnn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
      "Data Source=.\NorthWind.mdb;"

   ' Open the tables schema rowset
   Set rst = cnn.OpenSchema(adSchemaTables)

   ' Loop through the results and print
   ' the names in the debug window
   Do Until rst.EOF
      If rst.Fields("TABLE_TYPE") <> "VIEW" Then
         Debug.Print rst.Fields("TABLE_NAME")
      End If
      rst.MoveNext
   Loop

End Sub



Sub ADOCreateTable()

   Dim cat As New ADOX.Catalog
   Dim tbl As New ADOX.Table

   ' Open the catalog
   cat.ActiveConnection = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
      "Data Source=.\NorthWind.mdb;"

   ' Create a new Table object.
   With tbl
      .Name = "Contacts"
      ' Create fields and append them to the new Table
      ' object. This must be done before appending the
      ' Table object to the Tables collection of the
      ' Catalog.
      .Columns.Append "ContactName", adVarWChar
      .Columns.Append "ContactTitle", adVarWChar
      .Columns.Append "Phone", adVarWChar
      .Columns.Append "Notes", adLongVarWChar
      .Columns("Notes").Attributes = adColNullable
   End With

   ' Add the new table to the database.
   cat.Tables.Append tbl

   Set cat = Nothing


'
'The process for creating a table using DAO or ADOX is the same. First, create the object (TableDef or Table), append the columns (Field or Column objects), and finally append the table to the collection. Though the process is the same, the syntax is slightly different.
'With ADOX, it is not necessary to use a "create" method to create the column before appending it to the collection. The Append method can be used to both create and append the column.
'You 'll also notice the data type names for the columns are different between DAO and ADOX. The following table shows how the DAO data types that apply to Microsoft Jet databases map to the ADO data types.

'DAO data type ADO data type
'dbBinary adBinary
'dbBoolean adBoolean
'dbByte adUnsignedTinyInt
'dbCurrency adCurrency
'dbDate adDate
'dbDecimal adNumeric
'dbDouble adDouble
'dbGUID adGUID
'dbInteger adSmallInt
'dbLong adInteger
'dbLongBinary adLongVarBinary
'dbMemo adLongVarWChar
'dbSingle adSingle
'dbText adVarWChar

'Though not shown in this example, there are a number of other attributes of a table or column that you can set when creating the table or column, using the DAO Attributes property. The table below shows how these attributes map to ADO and Microsoft Jet Provider–specific properties.
'DAO TableDef Property Value ADOX Table Property Value
'Attributes dbAttachExclusive Jet OLEDB:Exclusive Link True
'Attributes dbAttachSavePWD Jet OLEDB:Cache Link Name/Password True
'Attributes dbAttachedTable Type "LINK"
'Attributes dbAttachedODBC Type "PASS-THROUGH"

'DAO Field Property Value ADOX Column Property Value
'Attributes dbAutoIncrField AutoIncrement True
'Attributes dbFixedField ColumnAttributes adColFixed
'Attributes dbHyperlinkField Jet OLEDB:Hyperlink True
'Attributes dbSystemField No equivalent n/a
'Attributes dbUpdatableField Attributes (Field Object) adFldUpdatable
'Attributes dbVariableField ColumnAttributes Not adColFixed

End Sub
Sub ADOCreateAttachedJetTable()

   Dim cat As New ADOX.Catalog
   Dim tbl As New ADOX.Table

   ' Open the catalog
   cat.ActiveConnection = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
      "Data Source=.\NorthWind.mdb;"

   ' Set the name and target catalog for the table
   tbl.Name = "Authors"
   Set tbl.ParentCatalog = cat

   ' Set the properties to create the link
   tbl.Properties("Jet OLEDB:Create Link") = True
   tbl.Properties("Jet OLEDB:Link Datasource") = ".\Pubs.mdb"
   tbl.Properties("Jet OLEDB:Link Provider String") = ";Pwd=password"
   tbl.Properties("Jet OLEDB:Remote Table Name") = "authors"

   ' Append the table to the collection
   cat.Tables.Append tbl

   Set cat = Nothing

End Sub

Sub ADOCreateAttachedODBCTable()

   Dim cat As New ADOX.Catalog
   Dim tbl As New ADOX.Table

   ' Open the catalog
   cat.ActiveConnection = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
      "Data Source=.\NorthWind.mdb;"

   ' Set the name and target catalog for the table
   tbl.Name = "Titles"
   Set tbl.ParentCatalog = cat

   ' Set the properties to create the link
   tbl.Properties("Jet OLEDB:Create Link") = True
   tbl.Properties("Jet OLEDB:Link Provider String") = _
      "ODBC;DSN=ADOPubs;UID=sa;PWD=;"
   tbl.Properties("Jet OLEDB:Remote Table Name") = "titles"

   ' Append the table to the collection
   cat.Tables.Append tbl

   Set cat = Nothing

End Sub

Sub ADOCreateAutoIncrColumn()

   Dim cat As New ADOX.Catalog
   Dim col As New ADOX.Column

   ' Open the catalog
   cat.ActiveConnection = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
      "Data Source=.\NorthWind.mdb;"

   ' Create the new auto increment column
   With col
      .Name = "ContactId"
      .Type = adInteger
      Set .ParentCatalog = cat
      .Properties("AutoIncrement") = True
   End With

   ' Append the column to the table
   cat.Tables("Contacts").Columns.Append col

   Set cat = Nothing

End Sub

Sub ADORefreshLinks()

   Dim cat As New ADOX.Catalog
   Dim tbl As ADOX.Table

   ' Open the catalog
   cat.ActiveConnection = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
      "Data Source=.\NorthWind.mdb;"

   For Each tbl In cat.Tables
      ' Check to make sure table is a linked table.
      If tbl.Type = "LINK" Then
         tbl.Properties("Jet OLEDB:Create Link") = False
         tbl.Properties("Jet OLEDB:Link Provider String") = _
            ";pwd=NewPassWord"
         tbl.Properties("Jet OLEDB:Link Datasource") = _
            ".\NewPubs.mdb"
         tbl.Properties("Jet OLEDB:Create Link") = True
      End If
   Next

End Sub

Sub ADOCreateIndex()

   Dim cat As New ADOX.Catalog
   Dim tbl As ADOX.Table
   Dim idx As New ADOX.Index

   ' Open the catalog
   cat.ActiveConnection = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
      "Data Source=.\NorthWind.mdb;"

   Set tbl = cat.Tables("Employees")

   ' Create Index object append table columns to it.
   idx.Name = "CountryIndex"
   idx.Columns.Append "Country"

   ' Allow Null values to be added in the index field
   idx.IndexNulls = adIndexNullsAllow

   ' Append the Index object to the Indexes collection of Table
   tbl.Indexes.Append idx

   Set cat = Nothing

End Sub

Sub ADOCreatePrimaryKey()

   Dim cat As New ADOX.Catalog
   Dim tbl As ADOX.Table
   Dim pk As New ADOX.Key

   ' Open the catalog
   cat.ActiveConnection = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
      "Data Source=.\NorthWind.mdb;"

   Set tbl = cat.Tables("Contacts")

   ' Create the Primary Key and append table columns to it.
   pk.Name = "PrimaryKey"
   pk.Type = adKeyPrimary
   pk.Columns.Append "ContactId"

   ' Append the Key object to the Keys collection of Table
   tbl.Keys.Append pk

   Set cat = Nothing
'----------------------
'Alternatively, the ADOX code to create and append the key could have been written in a single line of code. The following code:

'   ' Create the Primary Key and append table columns to it.
'   pk.Name = "PrimaryKey"
'   pk.Type = adKeyPrimary
'   pk.Columns.Append "ContactId"

'   ' Append the Key object to the Keys collection of Table
'   tbl.Keys.Append pk
'is equivalent to:

'   ' Append the Key object to the Keys collection of Table
'   tbl.Keys.Append "PrimaryKey", adKeyPrimary, "ContactId"
'----------------------
End Sub



Sub ADOCreateForeignKey()

   Dim cat As New ADOX.Catalog
   Dim tbl As ADOX.Table
   Dim fk As New ADOX.Key

   ' Open the catalog
   cat.ActiveConnection = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
      "Data Source=.\NorthWind.mdb;"

   ' Get the table for the foreign side of the relationship
   Set tbl = cat.Tables("Products")

   ' This key already exists in the Northwind database.
   ' For the purposes of this example, we're going to
   ' delete it and then recreate it
   tbl.Keys.Delete "CategoriesProducts"

   ' Create the Foreign Key
   fk.Name = "CategoriesProducts"
   fk.Type = adKeyForeign
   fk.RelatedTable = "Categories"

   ' Append column(s) in the foreign table to it
   fk.Columns.Append "CategoryId"

   ' Set RelatedColumn property to the name of the corresponding
   ' column in the primary table
   fk.Columns("CategoryId").RelatedColumn = "CategoryId"

   ' Append the Key object to the Keys collection of Table
   tbl.Keys.Append fk

   Set cat = Nothing
'----------------
'Alternatively, the ADOX code to create and append the key could have been written in a single line of code. The following code:
'   ' Create the Foreign Key
'   fk.Name = "CategoriesProducts"
'   fk.Type = adKeyForeign
'   fk.RelatedTable = "Categories"
'   ' Append column(s) in the foreign table to it
'   fk.Columns.Append "CategoryId"
'   ' Set RelatedColumn property to the name of the corresponding
'   ' column in the primary table
'   fk.Columns("CategoryId").RelatedColumn = "CategoryId"
'   ' Append the Key object to the Keys collection of Table
'   tbl.Keys.Append fk
'is equivalent to:
'   ' Append the Key object to the Keys collection of Table
'   tbl.Keys.Append "CategoriesProducts", adKeyForeign, _
'      "CategoryId", "Categories", "CategoryId"
'----------------
End Sub



Sub ADOCreateForeignKeyCascade()

   Dim cat As New ADOX.Catalog
   Dim tbl As ADOX.Table
   Dim fk As New ADOX.Key

   ' Open the catalog
   cat.ActiveConnection = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
   "Data Source=.\NorthWind.mdb;"

   ' Get the table for the foreign side of the relationship
   Set tbl = cat.Tables("Products")

   ' This key already exists in the Northwind database.
   ' For the purposes of this example, we're going to
   ' delete it and then recreate it
   tbl.Keys.Delete "CategoriesProducts"

   ' Create the Foreign Key
   fk.Name = "CategoriesProducts"
   fk.Type = adKeyForeign
   fk.RelatedTable = "Categories"

   ' Specify cascading updates and deletes
   fk.UpdateRule = adRICascade
   fk.DeleteRule = adRICascade

   ' Append column(s) in the foreign table to it
   fk.Columns.Append "CategoryId"
   ' Set RelatedColumn property to the name of the corresponding
   ' column in the primary table
   fk.Columns("CategoryId").RelatedColumn = "CategoryId"

   ' Append the Key object to the Keys collection of Table
   tbl.Keys.Append fk

   Set cat = Nothing

'--------------------------------------
'The following table shows how the values for the DAO Attributes property of a Relation object map to properties of the ADOX Key object.
'Note   The following values for the DAO Attributes property of a Relation object have no corresponding properties in ADOX: dbRelationDontEnforce, dbRelationInherited, dbRelationLeft, dbRelationRight.
'DAO Relation Object Property Value ADOX Key Object Property Value
'Attributes dbRelationUnique Type adKeyUnique
'Attributes dbRelationUpdateCascade UpdateRule adRICascade
'Attributes dbRelationDeleteCascade DeleteRule adRICascade
'--------------------------------------
End Sub



Sub ADOCreateQuery()

   Dim cat As New ADOX.Catalog
   Dim cmd As New ADODB.Command

   ' Open the catalog
   cat.ActiveConnection = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
      "Data Source=.\NorthWind.mdb;"

   ' Create the query
   cmd.CommandText = "SELECT * FROM Categories"
   cat.Views.Append "AllCategories", cmd

   Set cat = Nothing

End Sub

Sub ADOCreateParameterizedQuery()

   Dim cat As New ADOX.Catalog
   Dim cmd As New ADODB.Command

   ' Open the catalog
   cat.ActiveConnection = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
      "Data Source=.\NorthWind.mdb;"

   ' Create the Command
   cmd.CommandText = "PARAMETERS [prmRegion] TEXT(255);" & _
      "SELECT * FROM Employees WHERE Region = [prmRegion]"

   ' Create the Procedure
   cat.Procedures.Append "Employees by Region", cmd

   Set cat = Nothing

End Sub



Sub ADOModifyQuery()

   Dim cat As New ADOX.Catalog
   Dim cmd As ADODB.Command

   ' Open the catalog
   cat.ActiveConnection = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
      "Data Source=.\NorthWind.mdb;"

   ' Get the query
   Set cmd = cat.Procedures("Employees by Region").Command

   ' Update the SQL
   cmd.CommandText = "PARAMETERS [prmRegion] TEXT(255);" & _
      "SELECT * FROM Employees WHERE Region = [prmRegion] " & _
      "ORDER BY City"

   ' Save the updated query
   Set cat.Procedures("Employees by Region").Command = cmd

   Set cat = Nothing
'---------------------------------
'In the ADO code, setting the Procedure object's Command property to the modified Command object saves the changes. If this last step were not included, the changes would not have been persisted to the database. This difference results from the fact that ADO Command objects are designed as temporary queries while DAO QueryDef objects are designed as saved queries. You need to be aware of this when working with Commands, Procedures, and Views. You may think that the following ADO code examples are equivalent:

'   Set cmd = cat.Procedures("Employees by Region").Command
'   cmd.CommandText = "PARAMETERS [prmRegion] TEXT(255);" & _
'      "SELECT * FROM Employees WHERE Region = [prmRegion] " & _
'      "ORDER BY City"
'   Set cat.Procedures("Employees by Region").Command = cmd
                                                          
'and

'   cat.Procedures("Employees by Region").CommandText = _
'      "PARAMETERS [prmRegion] TEXT;" & _
'      "SELECT * FROM Employees WHERE Region = [prmRegion] " & _
'      "ORDER BY City"
                                 
'However, they are not. Both will compile, but the second piece of code will not actually update the query in the database. In the second example, ADOX will create a tear-off command object and hand it back to Visual Basic for Applications. Visual Basic for Applications will then ask ADOX to update the CommandText property, which it does. Finally, Visual Basic for Applications moves to execute the next line of code and the Command object is lost. ADOX is never asked to update the Procedure with the changes to the modified Command object
'---------------------------------
End Sub



Sub ADOCreateSQLPassThrough()

   Dim cat As New ADOX.Catalog
   Dim cmd As New ADODB.Command

   ' Open the catalog
   cat.ActiveConnection = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
      "Data Source=.\NorthWind.mdb;"

   ' Create the Command
   Set cmd.ActiveConnection = cat.ActiveConnection
   cmd.CommandText = "SELECT * FROM Titles WHERE Type = 'business'"
   cmd.Properties("Jet OLEDB:ODBC Pass-Through Statement") = True
   cmd.Properties("Jet OLEDB:Pass Through Query Connect String") = _
      "ODBC;DSN=ADOPubs;database=pubs;UID=sa;PWD=;"

   ' Create the Procedure
   cat.Procedures.Append "Business Books", cmd

   Set cat = Nothing

End Sub

Sub ADOChangePassword()

   Dim cat As New ADOX.Catalog

   ' Open the catalog, specifying the system database to use
   cat.ActiveConnection = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
      "Data Source=.\NorthWind.mdb;Jet OLEDB:System database=" & _
      "C:\Program Files\Microsoft Office\Office\SYSTEM.MDW"

   ' Change the password for the user Admin
   cat.Users("Admin").ChangePassword "", "password"

End Sub

Sub JROChangeDatabasePassword()

   Dim je As New JRO.JetEngine

   ' Make sure there isn't already a file with the
   ' name of the compacted database.
   If Dir(".\NewNorthWind.mdb") <> "" Then _
      Kill ".\NewNorthWind.mdb"

   ' Compact the database specifying the new database password
   je.CompactDatabase "Data Source=.\NorthWind.mdb;", _
      "Data Source=.\NewNorthWind.mdb;" & _
      "Jet OLEDB:Database Password=password"

   ' Delete the original database
   Kill ".\NorthWind.mdb"

   ' Rename the file back to the original name
   Name ".\NewNorthWind.mdb" As ".\NorthWind.mdb"

'Note   JRO, not ADOX, is used to change a database password at share level.
'Both DAO and JRO allow you to change the database password when compacting the database. The syntax is slightly different: in DAO, specify ";pwd=password;" in the Password parameter of CompactDatabase. In JRO, specify the provider-specific "Jet OLEDB:Database Password=password" in the destination connection parameter of CompactDatabase.

End Sub

Sub ADOCreateUser()

   Dim cat As New ADOX.Catalog

   ' Open the catalog, specifying the system database to use
   cat.ActiveConnection = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
      "Data Source=.\NorthWind.mdb;Jet OLEDB:System database=" & _
      "C:\Program Files\Microsoft Office\Office\SYSTEM.MDW;" & _
      "User Id=Admin;Password=password;"

   ' Create the new user and append it to the users collection
   cat.Users.Append "MyUser", "password"

End Sub
Sub ADOAddUserToNewGroup()

   Dim cat As New ADOX.Catalog

   ' Open the catalog
   cat.ActiveConnection = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
      "Data Source=.\NorthWind.mdb;User Id=Admin;" & _
      "Password=password;Jet OLEDB:System database=" & _
      "C:\Program Files\Microsoft Office\Office\SYSTEM.MDW"

   ' Create a new group
   cat.Groups.Append "MyGroup"

   ' Add the user to the new group
   cat.Users("MyUser").Groups.Append "MyGroup"

End Sub

Sub ADOSetUserObjectPermissions()

   Dim cat As New ADOX.Catalog

   ' Open the catalog
   cat.ActiveConnection = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
      "Data Source=.\NorthWind.mdb;User Id=Admin;" & _
      "Password=password;Jet OLEDB:System database=" & _
      "C:\Program Files\Microsoft Office\Office\SYSTEM.MDW"

   ' Set permissions for MyUser on the Customers table
   cat.Users("MyUser").SetPermissions "Customers", adPermObjTable, _
      adAccessSet, adRightRead Or adRightInsert Or adRightUpdate _
      Or adRightDelete
'------------------
'With the DAO Permissions property, which maps to the Rights parameter of the ADOX SetPermissions method, you supply a constant or combination of constants that represent the permissions to set. The table below shows how the DAO Security constants map to the ADOX Rights constants.
'DAO ADOX
'dbSecNoAccess adRightNone
'dbSecFullAccess adRightFull
'dbSecDelete adRightDrop
'dbSecReadSec adRightReadPermissions
'dbSecWriteSec adRightWritePermissions
'dbSecWriteOwner adRightWriteOwner
'dbSecCreate adRightCreate
'dbSecReadDef adRightReadDesign
'dbSecWriteDef adRightWriteDesign
'dbSecRetrieveData adRightRead
'dbSecInsertData adRightInsert
'dbSecReplaceData adRightUpdate
'dbSecDeleteData adRightDelete
'dbSecDBAdmin adRightFull
'dbSecDBCreate adRightCreate
'dbSecDBExclusive adRightExclusive
'dbSecDBOpen adRightRead
'------------------
End Sub

Sub ADOSetDatabasePermissions()

   Dim cat As New ADOX.Catalog

   ' Open the catalog
   cat.ActiveConnection = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
      "Data Source=.\NorthWind.mdb;User Id=Admin;" & _
      "Password=password;Jet OLEDB:System database=" & _
      "C:\Program Files\Microsoft Office\Office\SYSTEM.MDW"

   ' Set permissions for MyUser on the current database
   cat.Users("MyUser").SetPermissions "", adPermObjDatabase, _
      adAccessSet, adRightExclusive

End Sub

Sub ADOSetUserContainerPermissions()

   Dim cat As New ADOX.Catalog

   ' Open the catalog
   cat.ActiveConnection = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
      "Data Source=.\NorthWind.mdb;User Id=Admin;" & _
      "Password=password;Jet OLEDB:System database=" & _
      "C:\Program Files\Microsoft Office\Office\SYSTEM.MDW"

   ' Set permissions for MyUser on the Tables Container
   cat.Users("MyUser").SetPermissions Null, adPermObjTable, _
      adAccessSet, adRightRead Or adRightInsert Or adRightUpdate _
      Or adRightDelete, adInheritNone

End Sub

Sub ADOGetObjectOwner()

   Dim cat As New ADOX.Catalog

   ' Open the catalog
   cat.ActiveConnection = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
      "Data Source=.\NorthWind.mdb;User Id=Admin;" & _
      "Password=password;Jet OLEDB:System database=" & _
      "C:\Program Files\Microsoft Office\Office\SYSTEM.MDW"

   ' Print the owner of the Customers table
   Debug.Print cat.GetObjectOwner("Customers", adPermObjTable)

End Sub

Sub JROMakeDesignMaster()

   Dim repMaster As New JRO.Replica

   ' Make the Northwind database replicable.
   ' If successful, this will create a connection to the
   ' database.
   repMaster.MakeReplicable ".\NorthWind.mdb", False

   Set repMaster = Nothing

End Sub

Sub JROKeepObjectLocal()

   Dim repMaster As New JRO.Replica

   repMaster.ActiveConnection = ".\NorthWind.mdb"

   repMaster.SetObjectReplicability "Contacts", "Tables", False

   Set repMaster = Nothing

End Sub

Sub JROMakeObjectReplicable(strTable As String)

   Dim repMaster As New JRO.Replica

   repMaster.ActiveConnection = ".\NorthWind.mdb"

   repMaster.SetObjectReplicability strTable, "Tables", True

   Set repMaster = Nothing

End Sub

Function JROMakeAdditionalReplica(strReplicableDB As String, _
   strNewReplica As String) As Integer

   Dim repMaster As New JRO.Replica

   repMaster.ActiveConnection = strReplicableDB

   repMaster.CreateReplica strNewReplica, "Replica of " & _
      strReplicableDB

   Set repMaster = Nothing

End Function

Sub JROCreatePartial()

   Dim repFull As New JRO.Replica
   Dim repPartial As New JRO.Replica

   ' Create partial replica.
   repFull.ActiveConnection = ".\NorthWind.mdb"
   repFull.CreateReplica ".\FY96.mdb", "Partial Sales Replica", _
      jrRepTypePartial
   Set repFull = Nothing

   ' Create an expression based filter in the partial replica.
   ' The PopulatePartial method requires an exclusive connection
   repPartial.ActiveConnection = _
      "Provider=Microsoft.Jet.OLEDB.4.0;" & _
      "Data Source=.\FY96.mdb;Mode=Share Exclusive"
   repPartial.Filters.Append "Customers", jrFilterTypeTable, _
      "Region = 'CA'"

   ' Create a filter based on a relationship in the partial replica.
   repPartial.Filters.Append "Orders", jrFilterTypeRelationship, _
      "CustomersOrders"

   ' Repopulate the partial replica based on the filters.
   repPartial.PopulatePartial ".\NorthWind.mdb"

   Set repPartial = Nothing

End Sub

Sub JROListFilters()

   Dim repPartial As New JRO.Replica
   Dim flt As JRO.Filter
   Dim strFilterType As String

   repPartial.ActiveConnection = _
      "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=.\FY96.mdb"

   For Each flt In repPartial.Filters
      If flt.FilterType = jrFilterTypeTable Then
         strFilterType = "Table Filter"
      Else
         strFilterType = "Relationship Filter"
      End If
      Debug.Print flt.TableName & " : " & strFilterType & " : " & _
         flt.FilterCriteria
   Next

   Set repPartial = Nothing

End Sub

Sub JROTwoWayDirectSync()

   Dim repMaster As New JRO.Replica

   repMaster.ActiveConnection = ".\NorthWind.mdb"

   ' Sends changes made in each replica to the other.
   repMaster.Synchronize ".\FY96.mdb", jrSyncTypeImpExp, _
      jrSyncModeDirect

   Set repMaster = Nothing

End Sub

Sub JROInternetSync()

   Dim repMaster As New JRO.Replica

   repMaster.ActiveConnection = ".\NorthWind.mdb"

   ' Synchronize the local database with the replica on
   ' the Internet server.
   repMaster.Synchronize "sampleserver/files/Northwind.mdb", _
      jrSyncTypeImpExp, jrSyncModeInternet

   Set repMaster = Nothing

End Sub

Sub JROConflictTables()

   Dim repMaster As New JRO.Replica
   Dim rstConflicts As ADODB.Recordset

   repMaster.ActiveConnection = ".\NorthWind.mdb"

   Set rstConflicts = repMaster.ConflictTables

   If rstConflicts.BOF And rstConflicts.EOF Then
      ' There are no conflict tables so no conflicts occurred.
      Debug.Print "No conflicts."
   Else
      Do Until rstConflicts.EOF
         Debug.Print rstConflicts.Fields(0) & " had a conflict."
         rstConflicts.MoveNext
      Loop
   End If

End Sub

Sub ADODatabaseError()

   On Error GoTo ADODatabaseError_Err

   Dim cnn As New ADODB.Connection
   Dim errDB As ADODB.Error

   cnn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
      "Data Source=.\NonExistent.mdb"

   Exit Sub

ADODatabaseError_Err:
   For Each errDB In cnn.Errors
      Debug.Print "Description: " & errDB.Description
      Debug.Print "Number: " & errDB.Number & " (" & _
         Hex$(errDB.Number) & ")"
      Debug.Print "JetErr: " & errDB.SQLState
   Next

End Sub

Sub ADOTransactions()

   On Error GoTo ADOTransactions_Err

   Dim cnn As New ADODB.Connection
   Dim cat As New ADOX.Catalog
   Dim tbl As New ADOX.Table
   Dim bTrans As Boolean

   ' Open the connection
   cnn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
      "Data Source=.\NorthWind.mdb;"

   ' Begin the Transaction
   cnn.BeginTrans
   bTrans = True

   Set cat.ActiveConnection = cnn

   ' Create the Contacts table
   With tbl
      .Name = "Contacts"
      Set .ParentCatalog = cat
      .Columns.Append "ContactId", adInteger
      .Columns("ContactId").Properties("AutoIncrement") = True
      .Columns.Append "ContactName", adWChar
      .Columns.Append "ContactTitle", adWChar
      .Columns.Append "Phone", adWChar
      .Columns.Append "Notes", adLongVarWChar
      .Columns("Notes").Attributes = adColNullable
   End With
   cat.Tables.Append tbl

   ' Populate the Contacts table with information from the
   ' customers table
   cnn.Execute "INSERT INTO Contacts (ContactName, ContactTitle," & _
      "Phone) SELECT DISTINCTROW Customers.ContactName," & _
      "Customers.ContactTitle, Customers.Phone FROM Customers;"

   ' Add a ContactId field to the Customers Table
   Set tbl = cat.Tables("Customers")
   tbl.Columns.Append "ContactId", adInteger

   ' Populate the Customers table with the appropriate ContactId
   cnn.Execute "UPDATE DISTINCTROW Contacts INNER JOIN Customers " _
      & "ON Contacts.ContactName = Customers.ContactName SET " & _
      "Customers.ContactId = [Contacts].[ContactId];"

   ' Delete the ContactName, ContactTitle, and Phone columns
   ' from Customers
   tbl.Columns.Delete "ContactName"
   tbl.Columns.Delete "ContactTitle"
   tbl.Columns.Delete "Phone"

   ' Commit the transaction
   cnn.CommitTrans

   Exit Sub

ADOTransactions_Err:
   If bTrans Then cnn.RollbackTrans

   Debug.Print cnn.Errors(0).Description
   Debug.Print cnn.Errors(0).Number
   Debug.Print cnn.Errors(0).SQLState

End Sub

Sub JROCompactDatabase()

   Dim je As New JRO.JetEngine

   ' Make sure there isn't already a file with the
   ' name of the compacted database.
   If Dir(".\NewNorthWind.mdb") <> "" Then Kill ".\NewNorthWind.mdb"

   ' Compact the database
   je.CompactDatabase "Data Source=.\NorthWind.mdb;", _
      "Data Source=.\NewNorthWind.mdb;"

   ' Delete the original database
   Kill ".\NorthWind.mdb"

   ' Rename the file back to the original name
   Name ".\NewNorthWind.mdb" As ".\NorthWind.mdb"

End Sub

Sub JROEncryptDatabase()

   Dim je As New JRO.JetEngine

   ' Use compact to create a new, encrypted version of the database
   je.CompactDatabase "Data Source=.\NorthWind.mdb;", _
      "Data Source=.\NewNorthWind.mdb;" & _
      "Jet OLEDB:Encrypt Database=True"

End Sub

Sub JRORefreshCache()

   Dim cnn As New ADODB.Connection
   Dim rst As ADODB.Recordset
   Dim fld As ADODB.Field
   Dim je As New JRO.JetEngine

   ' Open the connection
   cnn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
      "Data Source=.\NorthWind.mdb;"

   ' Refresh the cache to ensure that the latest data
   ' is available.
   je.RefreshCache cnn

   ' Open a recordset and read the data
   Set rst = cnn.Execute("SELECT * FROM Shippers")
   Do Until rst.EOF
      For Each fld In rst.Fields
         Debug.Print fld.Value;
      Next
      Debug.Print
      rst.MoveNext
   Loop
   rst.Close

End Sub






Sub ADOCreateRecordset()

   Dim rst As New ADODB.Recordset

   rst.CursorLocation = adUseClient

   ' Add Some Fields
   rst.Fields.Append "dbkey", adInteger
   rst.Fields.Append "field1", adVarChar, 40, adFldIsNullable
   rst.Fields.Append "field2", adDate

   ' Create the Recordset
   rst.Open , , adOpenStatic, adLockBatchOptimistic

   ' Add Some Rows
   rst.AddNew Array("dbkey", "field1", "field2"), _
      Array(1, "string1", Date)
   rst.AddNew Array("dbkey", "field1", "field2"), _
      Array(2, "string2", #1/6/1992#)

   ' Look at the values -
   ' a value of 1 for status column = newly record
   rst.MoveFirst
   Debug.Print "Status", "dbkey", "field1", "field2"
   Do Until rst.EOF
      Debug.Print rst.Status, rst!dbkey, rst!field1, rst!field2
      rst.MoveNext
   Loop

   ' Commit the rows without ActiveConnection
   ' set resets the status bits
   rst.UpdateBatch adAffectAll

   ' Change the first of the two rows
   rst.MoveFirst
   rst!field1 = "changed"

   ' Now look at the status, first row shows 2 (modified row),
   ' second shows 8 (no modifications)
   ' Also note that the OriginalValue property shows the value
   ' before the modification
   rst.MoveFirst
   Do Until rst.EOF
      Debug.Print
      Debug.Print rst.Status, rst!dbkey, rst!field1, rst!field2
      Debug.Print , rst!dbkey.OriginalValue, _
         rst!field1.OriginalValue, rst!field2.OriginalValue
      rst.MoveNext
   Loop

End Sub



'The following code shows how to use the data link to open the connection rather than providing the connection information directly.

Sub ADOUseExistingDataLink()

   ' Opens an ADO Connection using a Data Links file (UDL)

   Dim cnn As New ADODB.Connection
   Dim rs As New ADODB.Recordset

   cnn.Open "File Name=.\NorthWind.udl;"

   rs.Open "Customers", cnn, adOpenKeyset, adLockReadOnly

   rs.MoveLast
   Debug.Print rs.RecordCount

   rs.Close
   cnn.Close

End Sub


Sub ADOCreateEnhancedAutoIncrColumn()

   Dim cat As New ADOX.Catalog
   Dim col As New ADOX.Column

   ' Open the catalog
   cat.ActiveConnection = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
      "Data Source=.\NorthWind.mdb;"

   ' Create the new auto increment column
   With col
      .Name = "ContactId"
      .Type = adInteger
      Set .ParentCatalog = cat
      .Properties("AutoIncrement") = True
      .Properties("Seed") = CLng(10)
      .Properties("Increment") = CLng(100)
   End With

   ' Append the column to the table
   cat.Tables("Contacts").Columns.Append col

   Set cat = Nothing

End Sub

'The following code demonstrates how to create a new Anonymous replica:

Function JROMakeAnonReplica(strReplicableDB As String, _
   strNewReplica As String) As Integer

   Dim repMaster As New JRO.Replica

   repMaster.ActiveConnection = strReplicableDB

   repMaster.CreateReplica strNewReplica, "Replica of " & _
      strReplicableDB, , jrRepVisibilityAnon

   Set repMaster = Nothing

End Function
'The following code demonstrates how to set the priority when creating a new replica:
Function JROMakeAdditionalReplica2(strReplicableDB As String, _
   strNewReplica As String, intPriority As Integer) As Integer

   Dim repMaster As New JRO.Replica

   repMaster.ActiveConnection = strReplicableDB

   repMaster.CreateReplica strNewReplica, "Replica of " & _
      strReplicableDB, , , intPriority

   Set repMaster = Nothing

End Function
'The following code demonstrates how to perform an indirect synchronization:
Sub JROTwoWayIndirectSync()

   Dim repMaster As New JRO.Replica

   repMaster.ActiveConnection = ".\NorthWind.mdb"

   ' Sends changes made in each replica to the other.
   repMaster.Synchronize ".\NewNorthWind.mdb", jrSyncTypeImpExp, _
      jrSyncModeIndirect

   Set repMaster = Nothing

End Sub

'The following code demonstrates how to perform a Microsoft Jet to SQL synchronization:

Sub JROJetSQLSync()

   Dim repMaster As New JRO.Replica

   repMaster.ActiveConnection = ".\Pubs.mdb"

   ' Sends changes made in each replica to the other.
   repMaster.Synchronize "", jrSyncTypeImpExp, jrSyncModeDirect

   Set repMaster = Nothing

End Sub

'The following code demonstrates how to turn on column level tracking when making a database replicable:

Sub JROMakeDesignMaster2()

   Dim repMaster As New JRO.Replica

   repMaster.MakeReplicable ".\NorthWind.mdb", True

   Set repMaster = Nothing

End Sub


