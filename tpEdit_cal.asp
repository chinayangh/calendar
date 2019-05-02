<%@codepage="65001" language=Jscript%>
<%Response.Charset="utf-8"%>
<%
/*
ThoughtProcess.net Jscript ASP Calendar Modification
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
<!-- #include file="tpLib_common.inc" -->
<!-- #include file="tpLib_db.inc" -->
<%
Response.Buffer=true;
var vError = "";
// ADODB Constants
//---- CursorTypeEnum Values ----
var adOpenForwardOnly = 0;
var adOpenKeyset = 1;
var adOpenDynamic = 2;
var adOpenStatic = 3;
//---- LockTypeEnum Values ----
var adLockReadOnly = 1;
var adLockPessimistic = 2;
var adLockOptimistic = 3;
var adLockBatchOptimistic = 4;
//---- CursorLocationEnum Values ----
var adUseServer = 2;
var adUseClient = 3;

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


var dbName="calendar.mdb";
var vSQL="";
var vAction="";
var vRequestVars="Action|ID|CategoryID|Subject|Description|Location|numDay|numYear|numMonth|StartTimeHour|StartTimeMin|StartTimeAMPM|EndTimeHour|EndTimeMin|EndTimeAMPM";

fProcessRequestVars(vRequestVars); // Set variables from Request
var strConnect	= "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" + Server.Mappath(dbName);


if(vAction.toLowerCase()=="modify")
	vSQL="SELECT * FROM Events where (ID = " + vID + ");";

var dbEvents = Server.CreateObject ("ADODB.Connection"); //*** Create an ADO database connection object.
var rsEvents	= Server.CreateObject ("ADODB.Recordset"); //** Create an ADO RecordSet object.
dbEvents.Open(strConnect); //*** Open the connection to the database. strconnect as found in datastore.inc

if(vAction.toLowerCase()=="remove")
	fRemoveRecords(dbEvents, "Events", "tpInput_cal.asp");

if(vAction.toLowerCase()=="modify")
{
	rsEvents.Open(vSQL, dbEvents, adOpenKeyset, adLockOptimistic); //*** Open the recordset
}else{
	rsEvents.Open("Events", dbEvents, adOpenForwardOnly, adLockPessimistic, adUseServer); //*** Open the recordset
	rsEvents.AddNew();
}
	var vDBFields="CategoryID|Subject|Description|Location|numDay|numYear|numMonth|StartTimeHour|StartTimeMin|StartTimeAMPM|EndTimeHour|EndTimeMin|EndTimeAMPM";
	fSetFields(vDBFields, rsEvents);
	rsEvents.Update();


//Response.Write(vSQL+"<BR>");
rsEvents.Close();
dbEvents.Close();
rsEvents	= null;
dbEvents = null;
Response.Redirect("tpInput_cal.asp");
%>
