<div align="center">

## Forms to Database the fast and easy way


</div>

### Description

Have you ever had a serious crunch time for writing code and finishing the project on time? Then this may help you out a bit especially if you do a lot of work with data storage. The basis and structure of this code can also be used for many things. The main object we will be working with is a collection object,(namely the Request.Form collection). This can also be used with Arrays and Session variables and anything that can hold a collection of values. You will see how we go through each item in the collection and assign the data that that collection variable is holding to a database with fields matching the same name of the items in the collection. The For..Each..Next statement is a very powerful method that can save you a lot of time. Have fun with this and fill free to use it in any variation you see fit.
 
### More Info
 
The code is rather stupid,(the dumber the better). Which means that you can pass any thing to it no matter what it holds, with the exception of binary data. It will respond the same. The only thing that MUST be hardcoded is that you must have a table in a database with field names matching the names of the form fields.

basic ADO operations and what the Request object is. Not relating to this paragraph on a side note I woudl like to mention that I use a very structured coding method. I try to keep everything structured jsut like a VB application. Remember that just because this is web page, it still is an application

IT will write the value from the forms to the fields that match

little debugging and testing of code. If you need to knwo wht kind of data is being submitted you can still use this code, but you may need to make modifications


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[c0rrupt0](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/c0rrupt0.md)
**Level**          |Advanced
**User Rating**    |4.8 (19 globes from 4 users)
**Compatibility**  |ASP \(Active Server Pages\), HTML
**Category**       |[Databases](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases__4-5.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/c0rrupt0-forms-to-database-the-fast-and-easy-way__4-6639/archive/master.zip)

### API Declarations

I work for a company called DotAnything Inc. I wrote this bit of code for a project that I am working on for it, but this code has been modified to not reflect anything that can tie it to any given customer.


### Source Code

```
<%
Option Explicit
'CopyRight 2001 DotAnything Inc.
'~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~
'   Application description
'~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~
'~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~
'    Variables
'~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~
'database variables
Dim rs				'recordset
Dim SQL				'sql statement
Dim conn      'connection
'This will call the SUB "main" and get the app started
Main
'~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~
'   Functions and Subs
'~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~
'Main Sub where most application proccessing will take place
Sub Main()
	Dim vF		'form fields place holder
	'This si the SQL statment which will determine what kind of data is opened.
  SQL = "SELECT tbl_Main.*"
	SQL = SQL & " FROM tbl_Main"
	SQL = SQL & " WHERE (((tbl_Main.fld_uid)=" & Request.Form("fld_uid") & "))"
  'Create a connection to the database
  'We will use a DSN DB called "myDB"
  Set Server.CreateObject("ADODB.Connection")
  conn.Open "DSN=myDB"
  'now we will create and open a recordset based on a user ID
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open SQL, conn, 2, 2
	'If no records were returned we know that this user does not have a record in this table.
  'So we add one.
	If rs.EOF Then
        'Add New record
		rs.AddNew
		'rs.Fields("fld_uid") = Request.Form("fld_uid") & ""
'Here you will see hte For..Each..Next work it's
'magic.
'vF will contian the name of the collection item
'We will loop through the collection FOR..EACH
'item the collection contians.
'vF will have a new value each loop
		For Each vF In Request.Form
'Here you can see how we are writing to a
'recordset field with the data that is in the
'form collection item that has the same name
			rs.Fields(vF) = Request.Form(vF) & ""
		Next
'Now we are done adding data so we cal the update
'method to update the table
		rs.Update
'now we close and clear the rs object
		rs.Close
		Set rs = Nothing
'This just tells you what record was added
		Response.Write "Record added for User #" & Request.Form("fld_uid")
	Else
'WE know the user is in here so we just update
'there record
'This code works the same as above but since we
'have already added this user we only need to
'update her data and not create a new record.
		For Each vF In Request.Form
			rs.Fields(vF) = Request.Form(vF) & ""
		Next
		rs.Update
		rs.Close
		Set rs = Nothing
		Response.Write "Record updated for User #" & Request.Form("fld_uid")
	End IF
	'all done
End Sub
%>
```

