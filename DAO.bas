Attribute VB_Name = "DAO"
'---------------------------------------------------------------------------------------
' Module    : DAO
' DateTime  : 4/10/2002 14:34
' Author    : Avaneesh Dvivedi
' Purpose   :Everything you want to do with DAO. Every functions
'            are included. DAO still being used by many programmers.
'            for further details check out my web site
'            http://www.tax-publishers.com/advivedi
'---------------------------------------------------------------------------------------

Option Explicit


'The following code listings show how to open (and then close) a shared, read-only database using DAO and ADO.

'DAO

Sub DAOOpenJetDatabaseReadOnly()

   Dim db As DAO.Database

   ' Open shared, read-only.
   Set db = DBEngine.OpenDatabase(".\NorthWind.mdb", False, True)
   db.Close

End Sub
'The following listings demonstrate how to override the Page Timeout setting of the engine and open a database using that setting.

'DAO

Sub DAOSetJetDBOption()

   Dim db As DAO.Database

   Set db = DBEngine.OpenDatabase(".\NorthWind.mdb")
   DBEngine.SetOption dbPageTimeout, 4000
   db.Close
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


'Share-Level (Password Protected) Databases
'The following listings demonstrate how to open a Microsoft Jet database that has been secured at the share level.
'DAO

Sub DAOOpenDBPasswordDatabase()

   Dim db As DAO.Database

   Set db = DBEngine.OpenDatabase(".\NorthWind.mdb", _
      False, False, ";pwd=password")
   db.Close
'In DAO, the Connect parameter of the OpenDatabase method sets the database password when opening a database. With ADO, the Microsoft Jet Provider connection property Jet OLEDB:Database Password sets the password instead
End Sub
'Opening a Database with User-Level Security
'These next listings demonstrate how to open a database that is secured at the user level using a workgroup information file named "system.mdw".
'DAO
Sub DAOOpenSecuredDatabase()

   Dim wks As DAO.Workspace
   Dim db As DAO.Database

   DBEngine.SystemDB = _
      "C:\Program Files\Microsoft Office\Office\SYSTEM.MDW"

   Set wks = DBEngine.CreateWorkspace("", "Admin", "")
   Set db = wks.OpenDatabase(".\NorthWind.mdb")

   db.Close
   wks.Close

End Sub


'External Databases
'The Microsoft Jet database engine can be used to access other database files, spreadsheets, and textual data stored in tabular format through installable ISAM drivers.
'The following listings demonstrate how to open a Microsoft Excel 2000 spreadsheet first using DAO, then using ADO and the Microsoft Jet provider.
'DAO

Sub DAOOpenISAMDatabase()

   Dim db As DAO.Database

   Set db = DBEngine.OpenDatabase(".\Sales.xls", _
      False, False, "Excel 8.0;")

   db.Close
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
'When you open Microsoft Access, you are opening a Microsoft Jet database. When writing code within Access, you may often want to use the same connection to Microsoft Jet as Access is using. To allow you to do this, Microsoft Access 2000 exposes two mechanisms: CurrentDB() and CurrentProject.Connection allow you to get a DAO Database object and an ADO Connection object, respectively, for the database Access currently has open.
'The following listings demonstrate how to get a reference to the database currently open in Microsoft Access.

Sub DAOGetCurrentDatabase()

   Dim db As DAO.Database

   Set db = CurrentDb()

End Sub

'The following code demonstrates how to open a Microsoft Jet database for shared, updateable access. Then the code immediately closes the database because this code is for demonstration purposes.


Sub DAOOpenJetDatabase()

   Dim db As DAO.Database

   Set db = DBEngine.OpenDatabase(".\NorthWind.mdb")
   db.Close

End Sub
'The following code listings show how to open (and then close) a shared, read-only database using DAO and ADO.


Sub DAOOpenRecordset()

   Dim db As DAO.Database
   Dim rst As DAO.Recordset
   Dim fld As DAO.Field

   ' Open the database
   Set db = DBEngine.OpenDatabase(".\NorthWind.mdb")

   ' Open the Recordset
   Set rst = db.OpenRecordset _
      ("SELECT * FROM Customers WHERE Region = 'WA'", _
      dbOpenForwardOnly, dbReadOnly)

   ' Print the values for the fields in
   ' the first record in the debug window
   For Each fld In rst.Fields
      Debug.Print fld.Value & ";";
   Next

   Debug.Print

   ' Close the recordset
   rst.Close

End Sub

Sub DAOMoveNext()

   Dim db As DAO.Database
   Dim rst As DAO.Recordset
   Dim fld As DAO.Field

   ' Open the database
   Set db = DBEngine.OpenDatabase(".\NorthWind.mdb")

   ' Open the Recordset
   Set rst = db.OpenRecordset _
      ("SELECT * FROM Customers WHERE Region = 'WA'", _
      dbOpenForwardOnly, dbReadOnly)

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

Sub DAOGetCurrentPosition()

   Dim db As DAO.Database
   Dim rst As DAO.Recordset

   ' Open the database
   Set db = DBEngine.OpenDatabase(".\NorthWind.mdb")

   ' Open the Recordset
   Set rst = db.OpenRecordset("SELECT * FROM Customers", _
      dbOpenDynaset)

   ' Print the absolute position
   Debug.Print rst.AbsolutePosition

   ' Move to the last record
   rst.MoveLast

   ' Print the absolute position
   Debug.Print rst.AbsolutePosition

   ' Close the recordset
   rst.Close

End Sub

Sub DAOFindRecord()

   Dim db As DAO.Database
   Dim rst As DAO.Recordset

   ' Open the database
   Set db = DBEngine.OpenDatabase(".\NorthWind.mdb")

   ' Open the Recordset
   Set rst = db.OpenRecordset("Customers", dbOpenDynaset)

   ' Find the first customer whose country is USA
   rst.FindFirst "Country = 'USA'"

   ' Print the customer id's of all customers in the USA
   Do Until rst.NoMatch
      Debug.Print rst.Fields("CustomerId").Value
      rst.FindNext "Country = 'USA'"
   Loop

   ' Close the recordset
   rst.Close

End Sub

Sub DAOSeekRecord()

   Dim db As DAO.Database
   Dim rst As DAO.Recordset

   ' Open the database
   Set db = DBEngine.OpenDatabase(".\NorthWind.mdb")

   ' Open the Recordset
   Set rst = db.OpenRecordset("Order Details", dbOpenTable)

   ' Select the index used to order the data in the recordset
   rst.Index = "PrimaryKey"

   ' Find the order where OrderId = 10255 and ProductId = 16
   rst.Seek "=", 10255, 16

   ' If a match is found print the quantity of the order
   If Not rst.NoMatch Then
      Debug.Print rst.Fields("Quantity").Value
   End If

   ' Close the recordset
   rst.Close

End Sub

Sub DAOFilterRecordset()

   Dim db As DAO.Database
   Dim rst As DAO.Recordset
   Dim rstFlt As DAO.Recordset

   ' Open the database
   Set db = DBEngine.OpenDatabase(".\NorthWind.mdb")

   ' Open the Recordset
   Set rst = db.OpenRecordset("Customers", dbOpenDynaset)

   ' Set the Filter to be used for subsequent recordsets
   rst.Filter = "Country='USA' And Fax Is Not Null"

   ' Open the filtered recordset
   Set rstFlt = rst.OpenRecordset()
   Debug.Print rstFlt.Fields("CustomerId").Value

   ' Close the recordsets
   rst.Close
   rstFlt.Close

End Sub

Sub DAOSortRecordset()

   Dim db As DAO.Database
   Dim rst As DAO.Recordset
   Dim rstSort As DAO.Recordset

   ' Open the database
   Set db = DBEngine.OpenDatabase(".\NorthWind.mdb")

   ' Open the Recordset
   Set rst = db.OpenRecordset("Customers", dbOpenDynaset)

   ' Sort the recordset based on Country and Region both in
   ' ascending order
   rst.Sort = "Country, Region"

   ' Open the sorted recordset
   Set rstSort = rst.OpenRecordset()
   Debug.Print rstSort.Fields("CustomerId").Value

   ' Close the recordsets
   rst.Close
   rstSort.Close

End Sub

Sub DAOAddRecord()

   Dim db As DAO.Database
   Dim rst As DAO.Recordset

   ' Open the database
   Set db = DBEngine.OpenDatabase(".\NorthWind.mdb")

   ' Open the Recordset
   Set rst = db.OpenRecordset _
      ("SELECT * FROM Customers", dbOpenDynaset)

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
   ' Position recordset on new record
   rst.Bookmark = rst.LastModified
   Debug.Print rst!CustomerId

   ' Close the recordset
   rst.Close

End Sub

Sub DAOUpdateRecord()

   Dim db As DAO.Database
   Dim rst As DAO.Recordset

   ' Open the database
   Set db = DBEngine.OpenDatabase(".\NorthWind.mdb")

   ' Open the Recordset
   Set rst = db.OpenRecordset _
      ("SELECT * FROM Customers WHERE CustomerId = 'LAZYK'", _
      dbOpenDynaset)

   ' Put the Recordset in Edit Mode
   rst.Edit

   ' Update the Contact name of the
   ' first record
   rst.Fields("ContactName").Value = "New Name"

   ' Save the changes you made to the
   ' current record in the Recordset
   rst.Update

   ' Close the recordset
   rst.Close

End Sub

Sub DAOReadMemo()

   Dim db As DAO.Database
   Dim rst As DAO.Recordset
   Dim sNotes As String
   Dim sChunk As String
   Dim cchChunkReceived As Long
   Dim cchChunkRequested As Long
   Dim cchChunkOffset As Long

   ' Open the database
   Set db = DBEngine.OpenDatabase(".\NorthWind.mdb")

   ' Open the Recordset
   Set rst = db.OpenRecordset _
      ("SELECT Notes FROM Employees", dbOpenDynaset)

   ' Initialize offset
   cchChunkOffset = 0

   ' cchChunkRequested artifically set low at 16
   ' to demonstrate looping
   cchChunkRequested = 16

   ' Loop through as many chunks as it takes
   ' to read the entire BLOB into memory
   Do
      ' Temporarily store the next chunk
      sChunk = rst!Fields("Notes").GetChunk _
         (cchChunkOffset, cchChunkRequested)

      ' Check how much we got
      cchChunkReceived = Len(sChunk)

      ' Adjust offset for next iteration
      cchChunkOffset = cchChunkOffset + cchChunkReceived

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
Sub DAOUpdateBLOB()

   Dim db As DAO.Database
   Dim rst As DAO.Recordset
   Dim rgPhoto() As Byte

   ' Open the database
   Set db = DBEngine.OpenDatabase(".\NorthWind.mdb")

   ' Open the Recordset
   Set rst = db.OpenRecordset( _
      "SELECT Photo FROM Employees", dbOpenDynaset)

   ' Get the first photo
   rgPhoto = rst.Fields("Photo").Value

   ' Move to the next record
   rst.MoveNext

   ' Put the Recordset in Edit Mode
   rst.Edit

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

Sub DAOExecuteQuery()

   Dim db As DAO.Database
   Dim rst As DAO.Recordset
   Dim fld As DAO.Field

   ' Open the database
   Set db = DBEngine.OpenDatabase(".\NorthWind.mdb")

   ' Open the Recordset
   Set rst = db.OpenRecordset("Products Above Average Price", _
      dbOpenForwardOnly, dbReadOnly)

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


'Executing a Parameterized Stored Query
'A parameterized stored query is an SQL statement that has been saved in the database and requires that additional variable information be specified in order to execute. The code below shows how to execute such a query.

Sub DAOExecuteParamQuery()

   Dim db As DAO.Database
   Dim qdf As DAO.QueryDef
   Dim rst As DAO.Recordset
   Dim fld As DAO.Field

   ' Open the database
   Set db = DBEngine.OpenDatabase(".\NorthWind.mdb")

   ' Get the QueryDef from the
   ' QueryDefs collection
   Set qdf = db.QueryDefs("Sales by Year")

   ' Specify the parameter values
   qdf.Parameters _
      ("Forms!Sales by Year Dialog!BeginningDate") = #8/1/1997#
   qdf.Parameters _
      ("Forms!Sales by Year Dialog!EndingDate") = #8/31/1997#

   ' Open the Recordset
   Set rst = qdf.OpenRecordset(dbOpenForwardOnly, dbReadOnly)

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

Sub DAOExecuteBulkOpQuery()

   Dim db As DAO.Database

   ' Open the database
   Set db = DBEngine.OpenDatabase(".\NorthWind.mdb")

   ' Execute the query
   db.Execute "UPDATE Customers SET Country = 'United States' " & _
      "WHERE Country = 'USA'"

   Debug.Print "Records Affected = " & db.RecordsAffected

   ' Close the database
   db.Close

End Sub



'Creating a Database
'Before tables or other objects can be defined, the database itself must be created. The following code creates and opens a new Microsoft Jet database.

Sub DAOCreateDatabase()

   Dim db As DAO.Database

   Set db = DBEngine.CreateDatabase(".\New.mdb", dbLangGeneral)
'
'In the DAO code above, the Locale parameter is specified as dbLangGeneral. In the ADOX code, locale is not explicitly specified. The default locale for the Microsoft Jet Provider is equivalent to dbLangGeneral. Use the ADO Locale Identifier property to specify a different locale.
'In DAO, CreateDatabase also can take a third Options parameter, specifying information for encrytion and database version. For example, the following line is used to create an encrypted, version 1.1 Microsoft Jet database:
'
'   Set db = DBEngine.CreateDatabase(".\New.mdb", dbLangGeneral, _
'      dbEncrypt + dbVersion11)
'
End Sub


'The following code demonstrates how to print the name of every table in the database by looping through the DAO TableDefs collection and the ADOX Tables collection.

Sub DAOListTables()

   Dim db As DAO.Database
   Dim tbl As DAO.TableDef

   ' Open the database
   Set db = DBEngine.OpenDatabase(".\NorthWind.mdb")

   ' Loop through the tables in the database and print their name
   For Each tbl In db.TableDefs
      Debug.Print tbl.Name
   Next

End Sub


Sub DAOCreateTable()

   Dim db As DAO.Database
   Dim tbl As DAO.TableDef

   ' Open the database
   Set db = DBEngine.OpenDatabase(".\NorthWind.mdb")

   ' Create a new TableDef object.
   Set tbl = db.CreateTableDef("Contacts")

   With tbl
      ' Create fields and append them to the new TableDef object.
      ' This must be done before appending the TableDef object to
      ' the TableDefs collection of the Database.
      .Fields.Append .CreateField("ContactName", dbText)
      .Fields.Append .CreateField("ContactTitle", dbText)
      .Fields.Append .CreateField("Phone", dbText)
      .Fields.Append .CreateField("Notes", dbMemo)
      .Fields("Notes").Required = False
   End With

   ' Add the new table to the database.
   db.TableDefs.Append tbl

   db.Close

End Sub


Sub DAOCreateAttachedJetTable()

   Dim db As DAO.Database
   Dim tbl As DAO.TableDef

   ' Open the database
   Set db = DBEngine.OpenDatabase(".\NorthWind.mdb")

   ' Create a new TableDef object.
   Set tbl = db.CreateTableDef("Authors")

   ' Set the properties to create the link
   tbl.Connect = ";DATABASE=.\Pubs.mdb;pwd=password;"
   tbl.SourceTableName = "authors"

   ' Add the new table to the database.
   db.TableDefs.Append tbl

   db.Close

End Sub

Sub DAOCreateAttachedODBCTable()

   Dim db As DAO.Database
   Dim tbl As DAO.TableDef

   ' Open the database
   Set db = DBEngine.OpenDatabase(".\NorthWind.mdb")

   ' Create a new TableDef object.
   Set tbl = db.CreateTableDef("Titles")

   ' Set the properties to create the link
   tbl.Connect = "ODBC;DSN=ADOPubs;UID=sa;PWD=;"
   tbl.SourceTableName = "titles"

   ' Add the new table to the database.
   db.TableDefs.Append tbl

   db.Close

End Sub

Sub DAOCreateAutoIncrColumn()

   Dim db As DAO.Database
   Dim tbl As DAO.TableDef
   Dim fld As DAO.Field

   ' Open the database
   Set db = DBEngine.OpenDatabase(".\NorthWind.mdb")

   ' Get the Contacts table
   Set tbl = db.TableDefs("Contacts")

   ' Create the new auto increment column
   Set fld = tbl.CreateField("ContactId", dbLong)
   fld.Attributes = dbAutoIncrField

   ' Add the new table to the database.
   tbl.Fields.Append fld

   db.Close

End Sub

'The next example shows how to update an existing linked table to refresh the link. This involves updating the connection string for the table and then resetting the Jet OLEDB:CreateLink property to tell Microsoft Jet to reestablish the link.

Sub DAORefreshLinks()

   Dim db As DAO.Database
   Dim tbl As DAO.TableDef

   ' Open the database
   Set db = DBEngine.OpenDatabase(".\NorthWind.mdb")

   For Each tbl In db.TableDefs
      ' Check to make sure table is a linked table.
      If (tbl.Attributes And dbAttachedTable) Then
         tbl.Connect = "MS Access;PWD=NewPassWord;" & _
            "DATABASE=.\NewPubs.mdb"
         tbl.RefreshLink
      End If
   Next

End Sub


Sub DAOCreateIndex()

   Dim db As DAO.Database
   Dim tbl As DAO.TableDef
   Dim idx As DAO.Index

   ' Open the database
   Set db = DBEngine.OpenDatabase(".\NorthWind.mdb")

   Set tbl = db.TableDefs("Employees")

   ' Create Index object append Field object to the Index object.
   Set idx = tbl.CreateIndex("CountryIndex")
   idx.Fields.Append idx.CreateField("Country")

   ' Append the Index object to the
   ' Indexes collection of the TableDef.
   tbl.Indexes.Append idx

   db.Close

End Sub

Sub DAOCreatePrimaryKey()

   Dim db As DAO.Database
   Dim tbl As DAO.TableDef
   Dim idx As DAO.Index

   ' Open the database
   Set db = DBEngine.OpenDatabase(".\NorthWind.mdb")

   Set tbl = db.TableDefs("Contacts")

   ' Create the Primary Key and append table columns to it.
   Set idx = tbl.CreateIndex("PrimaryKey")
   idx.Primary = True
   idx.Fields.Append idx.CreateField("ContactId")

   ' Append the Index object to the
   ' Indexes collection of the TableDef.
   tbl.Indexes.Append idx

   db.Close

End Sub

Sub DAOCreateForeignKey()

   Dim db As DAO.Database
   Dim rel As DAO.Relation
   Dim fld As DAO.Field

   ' Open the database
   Set db = DBEngine.OpenDatabase(".\NorthWind.mdb")

   ' This key already exists in the Northwind database.
   ' For the purposes of this example, we're going to
   ' delete it and then recreate it
   db.Relations.Delete "CategoriesProducts"

   ' Create the relation
   Set rel = db.CreateRelation()
   rel.Name = "CategoriesProducts"
   rel.Table = "Categories"
   rel.ForeignTable = "Products"

   ' Create the field the tables are related on
   Set fld = rel.CreateField("CategoryId")

   ' Set ForeignName property of the field to the name of
   ' the corresponding field in the primary table
   fld.ForeignName = "CategoryId"

   rel.Fields.Append fld

   ' Append the relation to the collection
   db.Relations.Append rel

End Sub

Sub DAOCreateForeignKeyCascade()

   Dim db As DAO.Database
   Dim rel As DAO.Relation
   Dim fld As DAO.Field

   ' Open the database
   Set db = DBEngine.OpenDatabase(".\NorthWind.mdb")

   ' This key already exists in the Northwind database.
   ' For the purposes of this example, we're going to
   ' delete it and then recreate it
   db.Relations.Delete "CategoriesProducts"

   ' Create the relation
   Set rel = db.CreateRelation()
   rel.Name = "CategoriesProducts"
   rel.Table = "Categories"
   rel.ForeignTable = "Products"

   ' Specify cascading updates and deletes
   rel.Attributes = dbRelationUpdateCascade Or _
      dbRelationDeleteCascade

   ' Create the field the tables are related on
   Set fld = rel.CreateField("CategoryId")
   ' Set ForeignName property of the field to the name of
   ' the corresponding field in the primary table
   fld.ForeignName = "CategoryId"

   rel.Fields.Append fld

   ' Append the relation to the collection
   db.Relations.Append rel

End Sub

Sub DAOCreateQuery()

   Dim db As DAO.Database
   Dim qry As DAO.QueryDef

   ' Open the database
   Set db = DBEngine.OpenDatabase(".\NorthWind.mdb")

   ' Create query
   Set qry = db.CreateQueryDef("AllCategories", _
      "SELECT * FROM Categories")

   db.Close

End Sub

Sub DAOCreateParameterizedQuery()

   Dim db As DAO.Database
   Dim qry As DAO.QueryDef

   ' Open the database
   Set db = DBEngine.OpenDatabase(".\NorthWind.mdb")

   ' Create query
   Set qry = db.CreateQueryDef("Employees by Region", _
      "PARAMETERS [prmRegion] TEXT(255);" & _
      "SELECT * FROM Employees WHERE Region = [prmRegion]")

   db.Close

End Sub

Sub DAOModifyQuery()

   Dim db As DAO.Database
   Dim qry As DAO.QueryDef

   ' Open the database
   Set db = DBEngine.OpenDatabase(".\NorthWind.mdb")

   ' Get the query
   Set qry = db.QueryDefs("Employees by Region")

   ' Update the SQL and save the updated query
   qry.SQL = "PARAMETERS [prmRegion] TEXT(255);" & _
      "SELECT * FROM Employees WHERE Region = [prmRegion] " & _
      "ORDER BY City"

   db.Close

End Sub


Sub DAOCreateSQLPassThrough()

   Dim db As DAO.Database
   Dim qry As DAO.QueryDef

   ' Open the database
   Set db = DBEngine.OpenDatabase(".\NorthWind.mdb")

   ' Create query
   Set qry = db.CreateQueryDef("Business Books", _
      "SELECT * FROM Titles WHERE Type = 'business'")

   qry.Connect = "ODBC;DSN=ADOPubs;UID=sa;PWD=;"
   qry.ReturnsRecords = True

   db.Close

End Sub

Sub DAOChangePassword()

   Dim wks As Workspace
   Dim usr As DAO.User

   ' Open the workspace, specifying the system database to use

   DBEngine.SystemDB = _
      "C:\Program Files\Microsoft Office\Office\SYSTEM.MDW"
   Set wks = DBEngine.CreateWorkspace("", "Admin", "")

   ' Change the password for the user Admin
   wks.Users("Admin").NewPassword "", "password"

End Sub

'The following code shows how to change the database password for enabling security at the share level.

Sub DAOChangeDatabasePassword()

   ' Make sure there isn't already a file with the
   ' name of the compacted database.
   If Dir(".\NewNorthWind.mdb") <> "" Then _
      Kill ".\NewNorthWind.mdb"

   ' Basic compact - creating new database named newnwind
   DBEngine.CompactDatabase ".\NorthWind.mdb", _
      ".\NewNorthWind.mdb", , , ";pwd=password;"

   ' Delete the original database
   Kill ".\NorthWind.mdb"

   ' Rename the file back to the original name
   Name ".\NewNorthWind.mdb" As ".\NorthWind.mdb"

End Sub



'Alternatively, the DAO code could be rewritten to use the NewPassword method of the Database object.

Sub DAOChangeDatabasePassword2()

   Dim db As DAO.Database

   Set db = DBEngine.OpenDatabase(".\NorthWind.mdb", True)
   db.NewPassword "", "password"
   db.Close

'A similar mechanism is not currently available in JRO or ADOX. You must use the CompactDatabase method in order to change the database password.

End Sub
Sub DAOCreateUser()

   Dim wks As DAO.Workspace

   ' Open a workspace
   DBEngine.SystemDB = _
      "C:\Program Files\Microsoft Office\Office\SYSTEM.MDW"
   Set wks = DBEngine.CreateWorkspace("", "Admin", "password")

   ' Create the user and append it to the Users collection
   wks.Users.Append wks.CreateUser("MyUser", "xNewUser", "password")

End Sub

Sub ADOCreateUser2()

   Dim cmd As New ADODB.Command

   ' Create the Command
   cmd.ActiveConnection = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
      "Data Source=.\NorthWind.mdb;Jet OLEDB:System database=" & _
      "C:\Program Files\Microsoft Office\Office\SYSTEM.MDW;" & _
      "User Id=Admin;Password=password;"

   ' Execute the DDL security command
   cmd.CommandText = "CREATE USER MyUser MyPW MyPID"
   cmd.Execute

End Sub


Sub DAOAddUserToNewGroup()

   Dim wks As DAO.Workspace

   ' Open the workspace
   DBEngine.SystemDB = _
      "C:\Program Files\Microsoft Office\Office\SYSTEM.MDW"
   Set wks = DBEngine.CreateWorkspace("", "Admin", "password")

   ' Create a new group
   wks.Groups.Append wks.CreateGroup("MyGroup", "xMyGroup")

   ' Add the user to the new group
   wks.Users("MyUser").Groups.Append _
      wks.Users("MyUser").CreateGroup("MyGroup")

End Sub

Sub DAOSetUserObjectPermissions()

   Dim db As DAO.Database
   Dim wks As DAO.Workspace
   Dim doc As DAO.Document

   ' Open the database
   DBEngine.SystemDB = _
      "C:\Program Files\Microsoft Office\Office\SYSTEM.MDW"
   Set wks = DBEngine.CreateWorkspace("", "Admin", "password")
   Set db = wks.OpenDatabase(".\NorthWind.mdb")

   ' Set permissions for MyUser on the Customers table
   Set doc = db.Containers("Tables").Documents("Customers")
   doc.UserName = "MyUser"
   doc.Permissions = dbSecRetrieveData Or dbSecInsertData _
      Or dbSecReplaceData Or dbSecDeleteData
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
Sub DAOSetDatabasePermissions()

   Dim db As DAO.Database
   Dim wks As DAO.Workspace
   Dim doc As DAO.Document

   ' Open the database
   DBEngine.SystemDB = _
      "C:\Program Files\Microsoft Office\Office\SYSTEM.MDW"
   Set wks = DBEngine.CreateWorkspace("", "Admin", "password")
   Set db = wks.OpenDatabase(".\NorthWind.mdb")

   ' Set permissions for MyUser on the current database
   Set doc = db.Containers("Databases").Documents("MSysDB")
   doc.UserName = "MyUser"
   doc.Permissions = dbSecDBExclusive

End Sub

Sub DAOSetUserContainerPermissions()

   Dim db As DAO.Database
   Dim wks As DAO.Workspace
   Dim ctr As DAO.Container

   ' Open the database
   DBEngine.SystemDB = _
      "C:\Program Files\Microsoft Office\Office\SYSTEM.MDW"
   Set wks = DBEngine.CreateWorkspace("", "Admin", "password")
   Set db = wks.OpenDatabase(".\NorthWind.mdb")

   ' Set permissions for MyUser on the Tables Container
   Set ctr = db.Containers("Tables")
   ctr.UserName = "MyUser"
   ctr.Inherit = True
   ctr.Permissions = dbSecRetrieveData Or dbSecInsertData _
      Or dbSecReplaceData Or dbSecDeleteData

End Sub

Sub DAOGetObjectOwner()

   Dim db As DAO.Database
   Dim wks As DAO.Workspace

   ' Open the database
   DBEngine.SystemDB = _
      "C:\Program Files\Microsoft Office\Office\SYSTEM.MDW"
   Set wks = DBEngine.CreateWorkspace("", "Admin", "password")
   Set db = wks.OpenDatabase(".\NorthWind.mdb")

   ' Print the owner of the Customers table
   Debug.Print db.Containers("Tables").Documents("Customers").Owner

End Sub

Sub DAOMakeDesignMaster()

   Dim dbsNorthwind As DAO.Database
   Dim prpNew As DAO.Property

   ' Open database for exclusive access.
   Set dbsNorthwind = DBEngine.OpenDatabase(".\NorthWind.mdb", True)

   With dbsNorthwind
      ' If Replicable property doesn't exist, create it.

      ' Turn on error handling in case property exists.
      On Error Resume Next
      Set prpNew = .CreateProperty("Replicable", dbText, "T")
      .Properties.Append prpNew

      ' Set database Replicable property to True.
      .Properties("Replicable") = "T"

      .Close
   End With

End Sub


Sub DAOKeepObjectLocal()

   Dim dbsNorthwind As DAO.Database
   Dim docTemp As DAO.Document
   Dim prpTemp As DAO.Property

   Set dbsNorthwind = DBEngine.OpenDatabase(".\NorthWind.mdb")

   Set docTemp = _
      dbsNorthwind.Containers("Tables").Documents("Contacts")
   Set prpTemp = docTemp.CreateProperty("KeepLocal", dbText, "T")

   docTemp.Properties.Append prpTemp

   dbsNorthwind.Close

End Sub

Sub DAOMakeObjectReplicable(strTable As String)

   Dim dbsNorthwind As DAO.Database
   Dim tdfTemp As DAO.TableDef

   Set dbsNorthwind = DBEngine.OpenDatabase(".\NorthWind.mdb")
   Set tdfTemp = dbsNorthwind.TableDefs(strTable)

   On Error GoTo ErrHandler

   tdfTemp.Properties("Replicable") = "T"

   On Error GoTo 0

   dbsNorthwind.Close

   Exit Sub

ErrHandler:

   Dim prpNew As DAO.Property

   If Err.Number = 3270 Then
      Set prpNew = tdfTemp.CreateProperty("Replicable", dbText, "T")
      tdfTemp.Properties.Append prpNew
   Else
      MsgBox "Error " & Err & ": " & Error
   End If

End Sub

Function DAOMakeAdditionalReplica(strReplicableDB As String, _
   strNewReplica As String) As Integer

   Dim dbsTemp As DAO.Database

   Set dbsTemp = DBEngine.OpenDatabase(strReplicableDB)

   dbsTemp.MakeReplica strNewReplica, "Replica of " & strReplicableDB

   dbsTemp.Close

End Function

Sub DAOCreatePartial()

   Dim dbsFull As DAO.Database
   Dim dbsPartial As DAO.Database
   Dim tdfCustomers As DAO.TableDef
   Dim relCustOrders As DAO.Relation

   ' Create partial replica.
   Set dbsFull = DBEngine.OpenDatabase(".\NorthWind.mdb")
   dbsFull.MakeReplica ".\FY96.mdb", "Partial Sales Replica", _
      dbRepMakePartial
   dbsFull.Close

   ' Create an expression based filter in the partial replica.
   Set dbsPartial = DBEngine.OpenDatabase(".\FY96.mdb", True)
   Set tdfCustomers = dbsPartial.TableDefs("Customers")
   tdfCustomers.ReplicaFilter = "Region = 'CA'"

   ' Create a filter based on a relationship in the partial replica.
   Set relCustOrders = dbsPartial.Relations("CustomersOrders")
   relCustOrders.PartialReplica = True

   ' Repopulate the partial replica based on the filters.
   dbsPartial.PopulatePartial ".\NorthWind.mdb"

   dbsPartial.Close

End Sub

Sub DAOListFilters()

   Dim dbPartial As DAO.Database
   Dim tbl As DAO.TableDef
   Dim rel As DAO.Relation

   Set dbPartial = DBEngine.OpenDatabase(".\FY96.mdb")

   For Each tbl In dbPartial.TableDefs
      If tbl.ReplicaFilter <> "" Then
         Debug.Print tbl.Name & " : Table Filter : " & _
            tbl.ReplicaFilter
      End If
   Next

   For Each rel In dbPartial.Relations
      Debug.Print rel.Name & " : Relationship Filter";
      If rel.PartialReplica Then
         Debug.Print " : Partial";
      End If
      Debug.Print
   Next

   dbPartial.Close

End Sub

Sub DAOTwoWayDirectSync()

   Dim dbsNorthwind As DAO.Database

   Set dbsNorthwind = DBEngine.OpenDatabase(".\NorthWind.mdb")

   ' Sends changes made in each replica to the other.
   dbsNorthwind.Synchronize ".\FY96.mdb", dbRepImpExpChanges

   dbsNorthwind.Close

End Sub

Sub DAOInternetSync()

   Dim dbsTemp As DAO.Database

   Set dbsTemp = DBEngine.OpenDatabase(".\NewNorthWind.mdb")

   ' Synchronize the local database with the replica on
   ' the Internet server.
   dbsTemp.Synchronize "sampleserver/files/Northwind.mdb", _
      dbRepImpExpChanges + dbRepSyncInternet

   dbsTemp.Close

End Sub

Sub DAOConflictTables()

   Dim dbsNorthwind As DAO.Database
   Dim tdfTest As DAO.TableDef
   Dim bConflict As Boolean

   Set dbsNorthwind = DBEngine.OpenDatabase(".\NorthWind.mdb")

   bConflict = False

   ' Enumerate TableDefs collection and check ConflictTable
   ' property of each.
   For Each tdfTest In dbsNorthwind.TableDefs
      If tdfTest.ConflictTable <> "" Then
         ' There was a conflict with this table
         Debug.Print tdfTest.Name & " had a conflict."
         bConflict = True
      End If
   Next tdfTest

   ' If bConflict is still false then we didn't find any
   ' tables that had conflicts.
   If Not bConflict Then Debug.Print "No conflicts."

   dbsNorthwind.Close

End Sub

Sub DAODatabaseError()

   On Error GoTo DAODatabaseError_Err

   Dim db As DAO.Database
   Dim errDB As DAO.Error

   Set db = DBEngine.OpenDatabase(".\NonExistent.mdb")

   Exit Sub

DAODatabaseError_Err:
   For Each errDB In DBEngine.Errors
      Debug.Print "Description: " & errDB.Description
      Debug.Print "Number: " & errDB.Number
      Debug.Print "JetErr: " & errDB.Number
   Next

End Sub

Sub DAOTransactions()

   On Error GoTo DAOTransactions_Err

   Dim wks As DAO.Workspace
   Dim db As DAO.Database
   Dim tbl As DAO.TableDef
   Dim bTrans As Boolean

   ' Get the default workspace
   Set wks = DBEngine.Workspaces(0)

   ' Open the database
   Set db = wks.OpenDatabase(".\NorthWind.mdb")

   ' Begin the Transaction
   wks.BeginTrans
   bTrans = True

   ' Create the Contacts table.
   Set tbl = db.CreateTableDef("Contacts")
   With tbl
      ' Create fields and append them to the new TableDef object.
      ' This must be done before appending the TableDef object to
      ' the TableDefs collection of the Database.
      .Fields.Append .CreateField("ContactId", dbLong)
      .Fields("ContactId").Attributes = dbAutoIncrField
      .Fields.Append .CreateField("ContactName", dbText)
      .Fields.Append .CreateField("ContactTitle", dbText)
      .Fields.Append .CreateField("Phone", dbText)
      .Fields.Append .CreateField("Notes", dbMemo)
      .Fields("Notes").Required = False
   End With
   db.TableDefs.Append tbl

   ' Populate the Contacts table with information from the
   ' customers table
   db.Execute "INSERT INTO Contacts (ContactName, ContactTitle," & _
      "Phone) SELECT DISTINCTROW [Customers].[ContactName], " & _
      "[Customers].[ContactTitle], [Customers].[Phone] " & _
      "FROM Customers;"

   ' Add a ContactId field to the Customers Table
   Set tbl = db.TableDefs("Customers")
   tbl.Fields.Append tbl.CreateField("ContactId", dbLong)

   ' Populate the Customers table with the appropriate ContactId
   db.Execute "UPDATE DISTINCTROW Contacts INNER JOIN Customers " & _
      "ON Contacts.ContactName = Customers.ContactName SET " & _
      "Customers.ContactId = [Contacts].[ContactId];"

   ' Delete the ContactName, ContactTitle, and Phone columns from
   ' Customers
   tbl.Fields.Delete "ContactName"
   tbl.Fields.Delete "ContactTitle"
   tbl.Fields.Delete "Phone"

   ' Commit the transaction
   wks.CommitTrans

   Exit Sub

DAOTransactions_Err:
   If bTrans Then wks.Rollback

   Debug.Print DBEngine.Errors(0).Description
   Debug.Print DBEngine.Errors(0).Number

End Sub

Sub DAOCompactDatabase()

   ' Make sure there isn't already a file with the
   ' name of the compacted database.
   If Dir(".\NewNorthWind.mdb") <> "" Then Kill ".\NewNorthWind.mdb"

   ' Basic compact - creating new database named newnwind
   DBEngine.CompactDatabase ".\NorthWind.mdb", ".\NewNorthWind.mdb"

   ' Delete the original database
   Kill ".\NorthWind.mdb"

   ' Rename the file back to the original name
   Name ".\NewNorthWind.mdb" As ".\NorthWind.mdb"

End Sub

Sub DAOEncryptDatabase()

   ' Use compact to create a new, encrypted version of the database
   DBEngine.CompactDatabase ".\NorthWind.mdb", _
      ".\NewNorthWind.mdb", , dbEncrypt

End Sub

Sub DAORefreshCache()

   Dim db As DAO.Database
   Dim rst As DAO.Recordset
   Dim fld As DAO.Field

   ' Open the database
   Set db = DBEngine.OpenDatabase(".\NorthWind.mdb")

   ' Refresh the cache to ensure that the latest data
   ' is available.
   DBEngine.Idle dbRefreshCache

   Set rst = db.OpenRecordset("SELECT * FROM Shippers")
   Do Until rst.EOF
      For Each fld In rst.Fields
         Debug.Print fld.Value;
      Next
      Debug.Print
      rst.MoveNext
   Loop
   rst.Close

End Sub







