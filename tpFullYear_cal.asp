<%@language=Jscript%>
<%
/*
ThoughtProcess.net Jscript ASP Calendar Example
Copyright (C) 2001
Author: Matt Kaatman
Date: 6/18/2001

This program is free software; you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation; version 2 of the License.
This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.
You should have received a copy of the GNU General Public License along with this program; if not, write to the Free Software Foundation, Inc., 675 Mass Ave, Cambridge, MA 02139, USA.

I'd love to hear where and how you're using it. If you'd like to report, send e-mail to asp@ThoughtProcess.net
Code contributions can be submitted via SourceForge.com.
*/
%>
<!--#include file="tpLib_cal.inc"-->
<!--#include file="tpLib_db.inc"-->
<!--#include file="tpDB_custom.inc"-->
<%
Response.Buffer=true;
var vToday = new fCalDate(new Date());
var gaEventsARR = new Array();
var vID = 0;
var dbPath	= "";
var dbName	= "calendar.mdb";	//Data base name
var vTables = "Events|Categories";	// pipe delimited Table Names

var vTableName=fCreateTableArr(vTables);

var vSQL	= "SELECT * FROM "+vTableName+" WHERE numMonth = "+vToday.Month+";";
var dbConn	= Server.CreateObject("ADODB.Connection");
var dbRS	= fCreateRS(dbConn, dbPath, dbName, vSQL);

fGetEventsData(vTableName, vSQL);

%>
<HTML>
<HEAD>
<script language="JavaScript">
function newWindow(theURL,winName,features) {
	window.open(theURL,winName,features);
}	
</script>	
<!--#include file="tpStyle_cal.inc"-->
</HEAD>
<BODY>
<table><TR>
<TD valign="top"><%fDrawCalendar(new fCalDate(new Date("01/01/2001"), "short"), "one", "small");%></TD>
<TD valign="top"><%fDrawCalendar(new fCalDate(new Date("02/01/2001"), "short"), "one", "small");%></TD>
</TR>
<TR>
<TD valign="top"><%fDrawCalendar(new fCalDate(new Date("03/01/2001"), "short"), "one", "small");%>
<TD valign="top"><%fDrawCalendar(new fCalDate(new Date("04/01/2001"), "short"), "one", "small");%>
</TR>
<TR>
<TD valign="top"><%fDrawCalendar(new fCalDate(new Date("05/01/2001"), "short"), "one", "small");%>
<TD valign="top"><%fDrawCalendar(new fCalDate(new Date("06/01/2001"), "short"), "one", "small");%>
</TR>
<TR>
<TD valign="top"><%fDrawCalendar(new fCalDate(new Date("07/01/2001"), "short"), "one", "small");%>
<TD valign="top"><%fDrawCalendar(new fCalDate(new Date("08/01/2001"), "short"), "one", "small");%>
</TR>
<TR>
<TD valign="top"><%fDrawCalendar(new fCalDate(new Date("09/01/2001"), "short"), "one", "small");%>
<TD valign="top"><%fDrawCalendar(new fCalDate(new Date("10/01/2001"), "short"), "one", "small");%>
</TR>
<TR>
<TD valign="top"><%fDrawCalendar(new fCalDate(new Date("11/01/2001"), "short"), "one", "small");%>
<TD valign="top"><%fDrawCalendar(new fCalDate(new Date("12/01/2001"), "short"), "one", "small");%>
</TR>
</table>
</BODY>
</HTML>
