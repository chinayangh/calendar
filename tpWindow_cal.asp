<%@ codepage="65001" language=Jscript%>
<%
/*
ThoughtProcess.net Jscript ASP Calendar Functions
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
var vaMonthNames = new Array("January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December");
var gaEventsARR = new Array();
var vID = 0;
var vMonth = 0;
var dbPath	= "";
var dbName	= "calendar.mdb";	//Data base name
var vTables = "Events|Categories";	// pipe delimited Table Names
if(Request("ID").Count)	vID=""+Request("ID");
if(Request("Month").Count)	vMonth=""+Request("Month");

var vTableName=fCreateTableArr(vTables);

var vSQL	= "SELECT * FROM "+vTableName+" WHERE numMonth = "+parseInt(vMonth)+" ORDER BY numDay;";
var dbConn	= Server.CreateObject("ADODB.Connection");
var dbRS	= fCreateRS(dbConn, dbPath, dbName, vSQL);

fGetEventsData(vTableName, vSQL);
var vToday = new fCalDate(new Date());

if(Request("ID").Count)
	vToday=new fCalDate(new Date(gaEventsARR[0].numMonth+"/"+gaEventsARR[0].numDay+"/"+gaEventsARR[0].numYear));
%>
<HTML>
<HEAD>
<meta http-equiv="content-type" content="text/html;charset=utf-8">
<script language="JavaScript">
function newWindow(theURL,winName,features) {
	window.open(theURL,winName,features);
}	
</script>	
<!--#include file="tpStyle_cal.inc"-->
</HEAD>
<BODY>
<div align="center">
<%
var vNextID=vID;
var vPrevID=vID;

for(x=0; x<gaEventsARR.length; x++)
{
	if(vID==gaEventsARR[x].ID)
	{
	%>
	<table class="tableWindow">
	<TR>
	<TH class="thCal"><%=vaMonthNames[gaEventsARR[x].numMonth-1]%></TH>
	<%
	var vDaysAway=(daysBetween(vToday.Date,new Date())*-1);
	var vDaysAwayText=vDaysAway+" days from today.";
	if(vDaysAway==0)	vDaysAwayText="Today!";
	else if(vDaysAway<0)vDaysAwayText=(vDaysAway*-1)+" days ago.";
	%>
	
	<TH class="thCal"><%=vDaysAwayText%></TH>
	<TH class="thCal">(<%=vToday.DayName%>)</TH>
	<TH class="thCal">	
	<%
	Response.Write(gaEventsARR[x].numMonth);
	Response.Write("/");
	Response.Write(gaEventsARR[x].numDay);
	Response.Write("/");
	Response.Write(gaEventsARR[x].numYear);
	%>
	</TH>
	</TR>
	<TR>
	<TD class="dayTextLarge">分校:</TD>
	<TD class="dayTextLarge">数据总量:</TD>
	
	<!--<TD class="dayTextLarge">Start:</TD>
	<TD class="dayTextLarge">End:</TD>-->
	<TD class="dayTextLarge">项目说明:</TD>
	</TR>
	<TR>
	<TD class="tdCal"><%=gaEventsARR[x].Location%></TD>
	<TD class="tdCal"><%=gaEventsARR[x].Subject%></TD>
	
	<!--<TD class="tdCal"><%if(gaEventsARR[x].StartTimeHour=="0"){Response.Write("N/A")}else{Response.Write(gaEventsARR[x].StartTimeHour+":"+gaEventsARR[x].StartTimeMin+" "+gaEventsARR[x].StartTimeAMPM)}%></TD>
	<TD class="tdCal"><%if(gaEventsARR[x].EndTimeHour=="0"){Response.Write("N/A")}else{Response.Write(gaEventsARR[x].EndTimeHour+":"+gaEventsARR[x].EndTimeMin+" "+gaEventsARR[x].EndTimeAMPM)}%></TD>-->
	<TD class="tdCal" colspan=4><%=gaEventsARR[x].Description%></TD>
	</TR>
	<!--<TR><TD class="tdCal" colspan=4><%=gaEventsARR[x].Description%></TD></TR>-->
	</table>
	<%
		if(x!=0)
			vPrevID=gaEventsARR[x-1].ID;
		else
			vPrevID=gaEventsARR[gaEventsARR.length-1].ID;
		if(x>=gaEventsARR.length-1)
			vNextID=gaEventsARR[0].ID;
		else
			vNextID=gaEventsARR[x+1].ID;

	}
}

if(gaEventsARR.length>1)
{%>
<table class="tableWindow">
<TR>
<TH class="thCal"><a title="上一个" href="<%=Request.ServerVariables("SCRIPT_NAME")%>?Month=<%=vToday.Month%>&ID=<%=parseInt(vPrevID)%>"><</a></TH>
<!--<TH class="thCal">当月外呼数据收集汇总</TH>-->
<TH class="thCal"><a title="下一个" href="<%=Request.ServerVariables("SCRIPT_NAME")%>?Month=<%=vToday.Month%>&ID=<%=parseInt(vNextID)%>">></a></TH>
<TH class="thCal" width=5><A href="javascript:window.close()" title="Close">X</A></TH>
</TR>
</table>
<%}%>
</div>
</BODY>
</HTML>
