<%@codepage="65001" language=Jscript%>
<%Response.Charset="utf-8"%>
<%
/*
ThoughtProcess.net Jscript ASP Calendar
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
<!--#include file="tpLib_common.inc"-->
<!--#include file="tpDB_custom.inc"-->
<%
Response.Buffer=true;
var vID = 0;
var vCategoryID	=0;
var vSubject	="";
var vDescription="";
var vLocation	="";
var vnumDay		=""+new Date().getDate();
var vnumYear	=""+new Date().getYear();
var vnumMonth	=""+(new Date().getMonth()+1);
var vStartTimeHour	="";
var vEndTimeHour	="";
var vStartTimeMin	="";
var vEndTimeMin		="";
var vStartTimeAMPM	="";
var vEndTimeAMPM	="";

var vAction="";
var vCalDate="";
var vToday = new fCalDate(new Date(), "short");
var gaEventsARR = new Array();
var vaCategories = new Array();

var vRequestVars="CalDate|Action|ID|CategoryID|Subject|Description|Location|numDay|numYear|numMonth|StartTimeHour|StartTimeMin|StartTimeAMPM|EndTimeHour|EndTimeMin|EndTimeAMPM";
fProcessRequestVars(vRequestVars); // Set variables from Request

if(vCalDate=="")	vCalDate=vToday.NormalDate;

var dbPath	= "";
var dbName	= "calendar.mdb";	//Data base name
var vTables = "Events|Categories";	// pipe delimited Table Names
var vDBFields="CategoryID|Subject|Description|Location|numDay|numYear|numMonth|StartTimeHour|StartTimeMin|StartTimeAMPM|EndTimeHour|EndTimeMin|EndTimeAMPM";

vaCategories=fGetTableData("Categories", dbPath, dbName, "SELECT * FROM Categories;", false, false);

if((Request("numDay").Count)||(vID!=0))
{
	var vModifySQL="";
	vAction="modify";
	var vModifySQL="SELECT * FROM Events WHERE numDay = "+vnumDay+";";
	if(vID!=0)	vModifySQL="SELECT * FROM Events WHERE ID = "+vID+";";

	var vaModify=fGetTableData("Events", dbPath, dbName, vModifySQL, false, false);
}

var vTableName=fCreateTableArr(vTables);
//fCreateGenericForm(vDBFields);


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

function validateDate(form)
{
// Matt Kaatman 07/03/2001
	// Is date valid for month?
	if (form.numMonth.value == 2) {
		// Check for leap year
		if(form.numDay.value>28)
		{
			if ( ( (form.numYear.value%4 == 0)&&(form.numYear.value%100 != 0) ) || (form.numYear.value%400 == 0) ) { // leap year
					alert("Date is a leap year, February only has 29 days.");
					form.numDay.value=29;
					return false;
			}else{
				alert("Date is not a leap year, February only has 28 days.");
				form.numDay.value=28;
				return false;
			}
		}
	}
	if ((form.numMonth.value==4)||(form.numMonth.value==6)||(form.numMonth.value==9)||(form.numMonth.value==11)) {
		if (form.numDay.value > 30)
		{
		alert("Date only has 30 days.");
		form.numDay.value=30;
		 return false;
		 }
	}
	return true;
}

function validateForm(form)
{
	if(!validateDate(form)) return false;

if(form.Subject.value=='')
{
	alert("外呼数据量不能为空！");
	form.Subject.focus();
	return false;
}
if(form.Description.value=='')
{
	alert("项目说明不能为空！");
	form.Description.focus();
	return false;
}
if(form.Location.value=='')
{
	alert("分校不能为空！");
	form.Location.focus();
	return false;
}
if(parseInt(form.Subject.value)>'6000')
{
	alert("数据总量不能大于6000  ！");
	form.Subject.focus();
	return false;
}

/*if(document.getElementById('outno').value>'6000')
{
  alert("输入值不能大于6000  ！");
  return false;
}*/

return true;

}


</script>
<!--#include file="tpStyle_cal.inc"-->
</HEAD>
<BODY>
<h2>外呼数据收集</h2>
<!--<div align="right"><a href="tpBase_cal.asp">查看数据汇总</a></div>-->
<div align="right"><A href="tpBase_cal.asp<%if(vCalDate!="")Response.Write("?CalDate="+vCalDate)%>">查看数据汇总</A></div>
<%//fDrawCalendar(vToday, "full", "large", gaEventsARR);%>
<%fDrawCalendar(vToday, "one", "small", gaEventsARR);%>
<BR><BR>

<%

if(vAction.toLowerCase()=="modify")
{
	if(!Request("ID").Count)	Response.Write("<table class='tableWindow'>");
	for(var modX=0; modX<vaModify.length; modX++)
	{
		vID=vaModify[modX].ID;
		vCategoryID=unescape(vaModify[modX].CategoryID);
		vSubject=unescape(vaModify[modX].Subject);
		vDescription=unescape(vaModify[modX].Description);
		vLocation=unescape(vaModify[modX].Location);
		vnumMonth=vaModify[modX].numMonth;
		vnumDay=vaModify[modX].numDay;
		vnumYear=vaModify[modX].numYear;
		vStartTimeHour=vaModify[modX].StartTimeHour;
		vStartTimeMin=vaModify[modX].StartTimeMin;
		vStartTimeAMPM=vaModify[modX].StartTimeAMPM;
		vEndTimeHour=vaModify[modX].EndTimeHour;
		vEndTimeMin=vaModify[modX].EndTimeMin;
		vEndTimeAMPM=vaModify[modX].EndTimeAMPM;
		if((!Request("numDay").Count)||(vaModify.length==1))		fDisplayEditForm();
		else{
%>
<TR>
<TD class='tdCal'><A href="tpInput_cal.asp?ID=<%=vID%>"><%=vID%></A></TD>
<TD class='tdCal'><%=vLocation%></TD>
<TD class='tdCal'><%=vSubject%></TD>
<TD class='tdCal'><%=vDescription%></TD>
<TD class='tdCal'><%=vnumYear%>-<%=vnumMonth%>-<%=vnumDay%></TD>
<!--<TD class='tdCal'><a href="tpEdit_cal.asp?Del=<%=vID%>&Action=remove">删除</A></TD>-->
</TD></TR>
<%		}
	}
	if(!Request("ID").Count)	Response.Write("</TABLE>");
}else{
		fDisplayEditForm();
}
%>

<%function fDisplayEditForm(){%>
<table class='tableWindow'>
<!--<TR><TH class='thCal'>外呼数据收集: <%if(vAction=="modify"){Response.Write(vID);}else{Response.Write("");}%></TH></TR>-->
<TR><TD class='tdCal'>
<form action='tpEdit_cal.asp' method='post' onSubmit='return validateForm(this);'>
<!--<select name="CategoryID">
<option value='0'<%if(vCategoryID==""){Response.Write(" selected ");}%>>All</option>
<%
for(var x=0; x<vaCategories.length; x++)
{
	Response.Write("<option value='"+vaCategories[x].ID+"'");
	if(vCategoryID==vaCategories[x].ID){Response.Write(" selected ");}
	Response.Write(">");
	Response.Write(vaCategories[x].Category);
	Response.Write("</option>");
}
%>
</select> CategoryID<BR>-->
<span>数据总量</span> <input type='text' id='outno' name='Subject' maxlength='249' value='<%=vSubject%>'><--请输入阿拉伯数字 <BR><BR>
<span>项目说明</span>
<textarea cols="30" rows="4" name="Description"><%=vDescription%></textarea><BR><BR>
<span class="w2">分校</span> <input type='text' name='Location' maxlength='249' value='<%=vLocation%>'><--请输入相应分校城市名称<BR><BR>
<span>外呼日期</span>
<select name='numMonth'>
<%
for(var x=1; x<13; x++)
{
	Response.Write("<option value="+x);
	if(vnumMonth==x){Response.Write(" selected ");}
	Response.Write(" >");
	Response.Write(x);
	Response.Write("</option>");
}
%>
</select>
/
<select name='numDay'>
<%
for(var x=1; x<32; x++)
{
	Response.Write("<option value="+x);
	if(vnumDay==x){Response.Write(" selected ");}
	Response.Write(" >");
	Response.Write(x);
	Response.Write("</option>");
}
%>
</select>
/

<select name='numYear'>
<%
vStartYear=(parseInt(new Date().getYear())-0)

vEndYear=(parseInt(new Date().getYear())+3)
for(var x=vStartYear; x<vEndYear; x++)
{
	Response.Write("<option value="+x);
	if(vnumYear==x){Response.Write(" selected ");}
	Response.Write(" >");
	Response.Write(x);
	Response.Write("</option>");
}
%>
</select>
<!--<BR>
<select name='StartTimeHour'>
<%
for(var x=0; x<13; x++)
{
	Response.Write("<option value='"+x+"'");
	if(vStartTimeHour==x){Response.Write(" selected ");}
	Response.Write(">");
	Response.Write(x);
	Response.Write("</option>");
}
%>
</select>
<select name='StartTimeMin'>
<option value='00'<%if(vStartTimeMin=="00"){Response.Write(" selected ");}%>>00</option>
<option value='15'<%if(vStartTimeMin=="15"){Response.Write(" selected ");}%>>15</option>
<option value='30'<%if(vStartTimeMin=="30"){Response.Write(" selected ");}%>>30</option>
<option value='45'<%if(vStartTimeMin=="45"){Response.Write(" selected ");}%>>45</option>
</select>
<select name='StartTimeAMPM'>
<option value='AM'<%if(vStartTimeAMPM=="AM"){Response.Write(" selected ");}%>>AM</option>
<option value='PM'<%if(vStartTimeAMPM=="PM"){Response.Write(" selected ");}%>>PM</option>
</select> StartTime (Leave hour 0 to ignore)<BR>
<select name='EndTimeHour'>
<%
for(var x=0; x<13; x++)
{
	Response.Write("<option value='"+x+"'");
	if(vEndTimeHour==x){Response.Write(" selected ");}
	Response.Write(">");
	Response.Write(x);
	Response.Write("</option>");
}
%>
</select>
<select name='EndTimeMin'>
<option value='00'<%if(vEndTimeMin=="00"){Response.Write(" selected ");}%>>00</option>
<option value='15'<%if(vEndTimeMin=="15"){Response.Write(" selected ");}%>>15</option>
<option value='30'<%if(vEndTimeMin=="30"){Response.Write(" selected ");}%>>30</option>
<option value='45'<%if(vEndTimeMin=="45"){Response.Write(" selected ");}%>>45</option>
</select>
<select name='EndTimeAMPM'>
<option value='AM'<%if(vEndTimeAMPM=="AM"){Response.Write(" selected ");}%>>AM</option>
<option value='PM'<%if(vEndTimeAMPM=="PM"){Response.Write(" selected ");}%>>PM</option>
</select> EndTime (Leave hour 0 to ignore)<BR>-->
<BR>
<input type='hidden' name='ID' value='<%=vID%>'>
</TD>
</tr>
<%if(vAction.toLowerCase()=="modify"){%>
<TR><Td class='tdCal'>
<table><TR><TD>
	<input type='hidden' name='Action' value="modify">
	<input type='Submit' value='保存/更改'>
</TD></form><!--<form action='tpEdit_cal.asp' method='post'><TD>
	<input type='hidden' name='Del' value='<%=vID%>'>
	<input type='hidden' name='Action' value="remove">
	<input type='submit' name='Remove' value='Remove Entry'>
</TD></form>--></TR></table>
</Td></TR>
<TR><Td class='tdCal'>
	<A href="<%=Request.ServerVariables("SCRIPT_NAME")+"?CalDate="+vCalDate%>">点此新增数据</A>
</TD></TR>
<%}else{%>
<TR><Td class='tdCal'>
	<input type='hidden' name='Action' value="add">
	<input type='Submit' value='提交保存'>
</Td></form></TR>
<TR><Td class='tdCal'>
<!--To modify an existing event, select the day from the Calendar.-->
TIPS:点击相应日期查看
</Td></TR>
<%}%>
</table>
<%}%>
<BR>
<!--<div><A href="tpBase_cal.asp<%if(vCalDate!="")Response.Write("?CalDate="+vCalDate)%>">View Calendar</A></div>-->
</BODY>
</HTML>
