<%
/*
ThoughtProcess.net Jscript ASP Calendar Database Functions
Copyright (C) 2001
Author: Matt Kaatman
Date: 6/18/2001

This program is free software; you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation; version 2 of the License.
This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.
You should have received a copy of the GNU General Public License along with this program; if not, write to the Free Software Foundation, Inc., 675 Mass Ave, Cambridge, MA 02139, USA.

I'd love to hear where and how you're using it. If you'd like to report, send e-mail to asp@ThoughtProcess.net
Code contributions can be submitted via SourceForge.com.
*/

function fEventsOBJ(ID,CategoryID,Subject,Description,Location,numDay,numYear,numMonth,StartTimeHour,StartTimeMin,StartTimeAMPM,EndTimeHour,EndTimeMin,EndTimeAMPM,strDate){
	if(ID=='null')	ID='';
	if(CategoryID=='null')	CategoryID='';
	if(Subject=='null')	Subject='';
	if(Description=='null')	Description='';
	if(Location=='null')	Location='';
	if(numDay=='null')	numDay='';
	if(numYear=='null')	numYear='';
	if(numMonth=='null')	numMonth='';
	if(StartTimeHour=='null')	StartTimeHour='';
	if(StartTimeMin=='null')	StartTimeMin='';
	if(StartTimeAMPM=='null')	StartTimeAMPM='';
	if(EndTimeHour=='null')	EndTimeHour='';
	if(EndTimeMin=='null')	EndTimeMin='';
	if(EndTimeAMPM=='null')	EndTimeAMPM='';
	if(strDate=='null')	strDate='';
	this.ID=''+ID;
	this.CategoryID=''+CategoryID;
	this.Subject=''+Subject;
	this.Description=''+Description;
	this.Location=''+Location;
	this.numDay=''+numDay;
	this.numYear=''+numYear;
	this.numMonth=''+numMonth;
	this.StartTimeHour=''+StartTimeHour;
	this.StartTimeMin=''+StartTimeMin;
	this.StartTimeAMPM=''+StartTimeAMPM;
	this.EndTimeHour=''+EndTimeHour;
	this.EndTimeMin=''+EndTimeMin;
	this.EndTimeAMPM=''+EndTimeAMPM;
	this.strDate=''+strDate;
}


function fGetEventsData(lTableName, lSQL){
	if((lSQL==null)||(lSQL==""))	lSQL= "SELECT * FROM "+lTableName+";";
	var lEventsRS = Server.CreateObject('ADODB.Recordset');
	lEventsRS = dbConn.Execute(lSQL);
	while (!lEventsRS.EOF){
		gaEventsARR[gaEventsARR.length] = new fEventsOBJ(
		""+lEventsRS('ID'),
		""+lEventsRS('CategoryID'),
		""+lEventsRS('Subject'),
		""+lEventsRS('Description'),
		""+lEventsRS('Location'),
		""+lEventsRS('numDay'),
		""+lEventsRS('numYear'),
		""+lEventsRS('numMonth'),
		""+lEventsRS('StartTimeHour'),
		""+lEventsRS('StartTimeMin'),
		""+lEventsRS('StartTimeAMPM'),
		""+lEventsRS('EndTimeHour'),
		""+lEventsRS('EndTimeMin'),
		""+lEventsRS('EndTimeAMPM'),
		""+lEventsRS('strDate')
		);
		lEventsRS.MoveNext();
	}
	lEventsRS.Close();
	lEventsRS=null;
}
%>