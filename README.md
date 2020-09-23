<div align="center">

## All\_Form\.ASP


</div>

### Description

Multi-function form for basic navigation, table editing, and recordset paging. This example includes code to dynamically build an SQL UPDATE command based on changed items on the current record.

http://adozone.cnw.com/default.htm
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Found on the World Wide Web](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/found-on-the-world-wide-web.md)
**Level**          |Beginner
**User Rating**    |4.6 (37 globes from 8 users)
**Compatibility**  |ASP \(Active Server Pages\)
**Category**       |[Controls/ Forms/ Dialogs/ Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/controls-forms-dialogs-menus__4-3.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/found-on-the-world-wide-web-all-form-asp__4-30/archive/master.zip)





### Source Code

```
<% Option Explicit %>
<% Response.Expires=0 %>
<HTML>
<HEAD></HEAD>
<BODY BGColor=White Text=Black>
<STYLE>
 	.btn {Width:100%}
</STYLE>
<%
Dim Page				' Local var for page #
Dim cn				' Connection object
Dim rs				' Recordset object
Dim Action			' Button pressed
Dim PageSize		' How far to page
Dim UpdSQL, MySQL		' String to hold SQL
Dim i					' Loop counter
Dim item, value	' Used to retrieve changed fields
Dim issueUpdate	' After Save button press, any changes to make?
Action = Request.Form("NavAction")
If Request.Form("Page") <> "" Then
 	Page = Request.Form("Page")
Else
 	Page = 1
End If
If Request.Form("PageSize") <> "" Then
 	PageSize = Request.Form("PageSize")
Else
 	PageSize = 5
End If
  		Set cn = Server.CreateObject("ADODB.Connection")
  		cn.Open Application("guestDSN")
  		' Get initial recordset
  		Set rs = Server.CreateObject("ADODB.Recordset")
  		MySQL = "SELECT * FROM AUTHORS"
rs.PageSize = PageSize
rs.Open MySQL, cn, adOpenKeyset, adLockOptimistic
  		Select Case Action
   			Case "Begin"
 	  Page = 1
   			Case "Back"
  		If (Page > 1) Then
   			Page = Page - 1
  		Else
   			Page = 1
   			  End If
  		rs.AbsolutePage = Page
   			Case "Forward"
  		If (CInt(Page) < rs.PageCount) Then
   			Page = Page + 1
  		Else
   			Page = rs.PageCount
 			  End If
  		rs.AbsolutePage = Page
   			Case "End"
  		rs.AbsolutePage = rs.PageCount
 	Case "Save"
  		' Grab the proper record, then update
  		' This routine is hard coded for AU_ID as the key field.
  		' To alter this to work with another DB Table you will need to
  		' Use the proper primary key instead of AU_ID.
  		rs.Close
  		MySQL = "SELECT * FROM AUTHORS WHERE au_id = '" & Request.Form("Au_id") & "'"
  		rs.MaxRecords = 1
  		rs.Open MySQL, cn, adOpenStatic, adLockOptimistic
  		UpdSQL = "UPDATE AUTHORS "
  		issueUpdate = False
  		For i = 0 To (rs.Fields.Count - 1)
   			item = rs.Fields(i).Name
   			value = Request.Form(item)
   			' Only update items that have changed
   			If (rs(i) <> value) Then
    				If issueUpdate = False Then
     					UpdSQL = UpdSQL & "SET "
    				Else
     					UpdSQL = UpdSQL & ","
    				End If
    				issueUpdate = True
    				Select Case VarType(rs.Fields(i))
     					' Determine datatype for proper SQL UPDATE syntax
     					' NOTE: Not all data types covered
     					Case vbString, vbDate
      						UpdSQL = UpdSQL & item & "='" & value & "'"
     					Case vbNull
     					Case vbInteger
      						UpdSQL = UpdSQL & item & "=" & value
     					Case vbBoolean
      						If value Then
       							UpdSQL = UpdSQL & item & "= 1"
      						Else
       							UpdSQL = UpdSQL & item & "= 0"
      						End If
    				End Select
   			End If
  		Next
  		UpdSQL = UpdSQL & " WHERE au_id = '" & Request.Form("Au_id") & "'"
  		If issueUpdate Then
   			cn.Execute UpdSQL
   			Set rs = cn.Execute(MySQL)
   			  End If
   			Case "New"
  		' response.write "New"
    				rs.AddNew
   			Case "Bookmark"
    				Session("myBookMark") = rs.BookMark
   			Case "Goto"
    				If Not IsNull(Session("myBookMark")) Then
     					rs.BookMark = Session("myBookMark")
    				End If
   			Case Else
   			  rs.MoveFirst
  		End Select
%>
<Center>
<!-- 2 Column Table -->
<!-- 1 Column for Data, 1 for Controls -->
<Table Align=Center border=1 BGColor=Navy
  BorderColorDark=Navy BorderColorLight=Aqua BorderColor=Blue>
<!-- Table Header -->
<th Colspan=2>
   <Font Color=White Size=+2><Center>Navigating Example</Center></Font>
</th>
<!-- Main Table Content -->
<tr><td>
<!-- Nested Table 1 -->
<!-- Author Detail -->
<Form Action=all_form.asp Method="POST">
<TABLE Align=Left BORDER=0 BGColor=Gray Text=White>
 	<%
 	For i = 0 To rs.Fields.Count - 1
  		%>
  		<TR><TD><B><%= rs.Fields(i).Name %></B></TD>
  		<TD><Input Type=text Name="<%= rs.Fields(i).Name %>" Value="<%= rs(i) %>"></TD>
</TR>
  		<%
 	Next
 	%>
</TABLE>
</td>
<td BGColor=Black Width=100>
 	<!-- Nested Form 2 -->
  		<!-- Persisted Values -->
 	  <Input Type="Hidden" Name="PageSize" Value="1">
 	  <Input Type="Hidden" Name="Page" Value="<%= Page %>">
 	<!-- Navigation Buttons -->
 	  <INPUT TYPE="Submit" Name="NavAction" Value="Begin" Class=Btn><BR>
 	  <INPUT TYPE="Submit" Name="NavAction" Value="Back" Class=Btn><BR>
 	  <INPUT TYPE="Submit" Name="NavAction" Value="Forward" Class=Btn><BR>
 	  <INPUT TYPE="Submit" Name="NavAction" Value="End" Class=Btn><P>
 	  <INPUT TYPE="Submit" Name="NavAction" Value="Save" Class=Btn><BR>
 	  <INPUT TYPE="Submit" Name="NavAction" Value="New" Class=Btn><P>
 	  <INPUT TYPE="Submit" Name="NavAction" Value="Bookmark" Class=Btn><BR>
 	  <INPUT TYPE="Submit" Name="NavAction" Value="Goto" Class=Btn><P>
</td>
</tr>
</table>
</Form>
<P>
<!-- Floating Frame -->
 	<IFRAME width=70% height=180 src="list.asp?auid=<%= rs( </include/code.asp?source=/ado/samples/list.asp?auid=<%= rs(>"au_id") %>" FrameBorder=1 Scrolling=No>
 	<FRAME width=70% height=180 src="list.asp?auid=<%= rs( </include/code.asp?source=/ado/samples/list.asp?auid=<%= rs(>"au_id") %>">
 	</IFRAME>
</Center>
</BODY>
</HTML>
```

