<%@ codepage="65001" language=Jscript%>
<%
/*
ThoughtProcess.net Jscript ASP Calendar Default Example
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
var vID = 0;

var vToday = new fCalDate(new Date());
var gaEventsARR = new Array();

var dbPath	= "";
var dbName	= "calendar.mdb";	//Data base name
var vTables = "Events|Categories";	// pipe delimited Table Names

var vTableName=fCreateTableArr(vTables);

var vSQL	= "SELECT * FROM "+vTableName+" WHERE numMonth = "+vToday.Month+" ORDER BY numDay;";
var dbConn	= Server.CreateObject("ADODB.Connection");
var dbRS	= fCreateRS(dbConn, dbPath, dbName, vSQL);

fGetEventsData(vTableName, vSQL);
%>
<HTML>
<HEAD>
<meta http-equiv="content-type" content="text/html;charset=utf-8">
<meta content="IE=Edge,chrome=1" http-equiv="X-UA-Compatible">
<meta content="webkit" name="renderer">
<meta content="width=device-width, initial-scale=1.0" name="viewport">
<script language="JavaScript">
function newWindow(theURL,winName,features) {
	window.open(theURL,winName,features);
}	
</script>	
<!--#include file="tpStyle_cal.inc"-->
</HEAD>
<BODY>
<h2>外呼数据收集</h2>
<div align="right"><a href="export.asp">导出</a>&nbsp&nbsp<a href="tpInput_cal.asp">去录入</a></div>
<%fDrawCalendar(vToday, "full", "large", gaEventsARR);%>
<BR><BR>
<!--<%fDrawCalendar(vToday, "one", "small", gaEventsARR);%>-->
</BODY>
</HTML>
