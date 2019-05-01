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

function fEvent(lDay, lTitle, lDesc, lCategory, lURL)
{
	// Create an event. Most likely fed from a database.
if((lDesc==null)||(lDesc==""))	lDesc="";
if((lURL==null)||(lURL==""))	lURL="";
if((lCategory==null)||(lCategory==""))	lCategory="";
	this.Category = lCategory;
	this.Day = lDay;
	this.Title = lTitle;
	this.Desc = lDesc;
	this.URL = lURL;
}


function fCalDate(lDate, lMonthNamesLength)
{
	//Create Calendar Date
	if((lDate==null)||(lDate==""))
		lDate=new Date();
	if(Request("CalDate").Count)
		lDate=new Date(Request("CalDate"));
	var laDayNames = new Array("Sunday", "Monday", "Tuesday", "Wednesday","Thursday", "Friday", "Saturday");
	if((lMonthNamesLength=="")||(lMonthNamesLength==null))
		lMonthNamesLength="long";
	else
		lMonthNamesLength="short";

	var	laMonthNames = new Array();
	if(lMonthNamesLength=="short")
	{
		laMonthNames = new Array("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec");
	}else{
		laMonthNames = new Array("January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December");
	}

	this.DoM=lDate.getDate();			// Day of Month
	this.Year=lDate.getFullYear()		// Year
	this.Month=lDate.getMonth()+1;		// Month
	this.Day=lDate.getDay();			// Day Num
	this.DayName = laDayNames[this.Day];// Day Name
	this.MonthName = laMonthNames[this.Month-1];// Month Name
	this.Date=lDate;					// Original Date
	this.MonthStartDay = (new Date(this.Month+"/1/"+this.Year)).getDay(); // Day Num that the Month Starts
	this.DaysInMonth=31;
	if(new Date(this.Month+"/31/"+this.Year).getMonth()+1!=this.Month)
		this.DaysInMonth=30;
	if(this.Month-1==1)
	{	// February
		if(((this.Year%4==0)&&(this.Year%100!=0))||(this.Year%400==0))
			this.DaysInMonth=29;	// Leap Year?
		else
			this.DaysInMonth=28;
	}
	this.NormalDate=this.Month+"/"+this.DoM+"/"+this.Year;
}

function fDrawCalendar(lToday, lLabelSelect, lCalSize, laEvents)
{
	// Function to Draw the Calendar
	// lToday = Date to Start Draw
	// lLabelSelect = Size of Day Name Labels (one|three|full)
	// lCalSize = Size of Calendar. Sets blank spacer.
	// laEvents = Object with .Day and .Desc

	var lEdit = false;
	if(((""+Request.ServerVariables("SCRIPT_NAME")).toLowerCase()).indexOf("input")>-1)
		lEdit=true;
	if((laEvents==null)||(laEvents==""))
		laEvents=new Array();

	var lSpacerText="";
	if((lCalSize==null)||(lCalSize==""))lCalSize="small";
	if(lCalSize.toLowerCase()!="small")
		lSpacerText="<br><br>&nbsp;<br>&nbsp;";
	var	laMonthNames = new Array("January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December");
	var lDayLabelsOne="S|M|T|W|T|F|S";
	var lDayLabelsThree="Sun|Mon|Tues|Wed|Thu|Fri|Sat";
	var lDayLabelsFull="Sunday|Monday|Tuesday|Wednesday|Thursday|Friday|Saturday";
	var laDayLabels=new Array();
	if((lLabelSelect==null)||(lLabelSelect==""))lLabelSelect="one";
	if(lLabelSelect.toLowerCase()=="three")		laDayLabels=lDayLabelsThree.split("|");
	else if(lLabelSelect.toLowerCase()=="full")	laDayLabels=lDayLabelsFull.split("|");
	else										laDayLabels=lDayLabelsOne.split("|");

	var lMaxWIM=6;		// Maximum weeks in a month
	var lDaysInWeek=7;	// Days in a week
	var lDaysInMonth=lToday.DaysInMonth;	// Days in current month

	var lCurrentDay=0;
	var lEventCounter=0;	// Database Counter

	Response.Write("<TABLE class=\"table"+lCalSize+"\">\n");
	Response.Write("<TR>");
	Response.Write("<TH class=\"thCalHeader\" ><A title=\"上月\" href=\""+Request.ServerVariables("SCRIPT_NAME")+"?CalDate="+mdy(lToday.Date,-1)+"\"><</A></TH>");
	Response.Write("<TH class=\"thCalHeader\">&nbsp;</TH>");
	Response.Write("<TH class=\"thCalHeader\" colspan=3>"+lToday.MonthName+" "+lToday.Year+"</TH>");
	Response.Write("<TH class=\"thCalHeader\" ><A title=\"跳转到当前月: "+laMonthNames[new Date().getMonth()]+"\" href=\""+Request.ServerVariables("SCRIPT_NAME")+"\">J</A></TH>");
	Response.Write("<TH class=\"thCalHeader\" ><A title=\"下月\" href=\""+Request.ServerVariables("SCRIPT_NAME")+"?CalDate="+mdy(lToday.Date,1)+"\">></A></A></TH>");
	Response.Write("</TR>\n")
	Response.Write("<TR>");
	for(var x=0; x<7; x++)
	{
		Response.Write("<TH class=\"thCal\">"+laDayLabels[x]+"</TH>");
	}
	Response.Write("</TR>\n");
	for (WeekInMonth=0; WeekInMonth<lMaxWIM; WeekInMonth++)
	{
		if(lCurrentDay<lDaysInMonth)
		{
			Response.Write("<TR>");
			for (x=0; x<lDaysInWeek; x++)
			{
				if(lCurrentDay>=1)	lCurrentDay++;
				Response.Write("<TD ");
/*
				if((lEventCounter<laEvents.length))	// Write Description for Table Cell mouseover
				{
					if(lCurrentDay==laEvents[lEventCounter].numDay)
						Response.Write(" TITLE=\""+laEvents[lEventCounter].Description+"\" ");
				}
*/
				if((new Date().getMonth()+1==lToday.Month)&&(lCurrentDay==(new Date()).getDate())&&(new Date().getFullYear()==lToday.Year))	// Determine if it's today
					Response.Write(" class=\"tdToday\"");
				else
					Response.Write(" class=\"tdCal\"");
				Response.Write(" width=\"14%\" valign=top>");

				if(lCurrentDay<=lDaysInMonth)
				{
					if(WeekInMonth==0)
					{
						if(lToday.MonthStartDay==x)
							lCurrentDay++;
						if(lToday.MonthStartDay>x)
							Response.Write(lSpacerText);
					}
					if((lToday.MonthStartDay<=x)||(lCurrentDay>x))
					{
						Response.Write("<div class=\"dayText"+lCalSize+"\">");
						if((lEventCounter<laEvents.length))
						{
//							Response.Write(laEvents[lEventCounter].numDay);
							if((lCurrentDay==laEvents[lEventCounter].numDay))
							{
								if(lEdit!=true)
/*									Response.Write("<A href=\"JavaScript:newWindow('tpWindow_cal.asp?Month="+vToday.Month+"&ID="+laEvents[lEventCounter].ID+"','Calendar','width=440,height=180,toolbar=no,location=no,scrollbars=yes,directories=no,resizable=yes,menubar=no,status=no');\">");*/
										Response.Write("<A href=\""+Request.ServerVariables("SCRIPT_NAME")+"?numDay="+laEvents[lEventCounter].numDay+"&CalDate="+lToday.NormalDate+"\">");
								else
									Response.Write("<A href=\""+Request.ServerVariables("SCRIPT_NAME")+"?numDay="+laEvents[lEventCounter].numDay+"&CalDate="+lToday.NormalDate+"\">");
							}
						}
						Response.Write(lCurrentDay);
						if((lEventCounter<laEvents.length))
						{
							if((lCurrentDay==laEvents[lEventCounter].numDay))
								Response.Write("</A>");
						}
						Response.Write("</div>")

						if((lEventCounter<=laEvents.length-1))

						{
							//for(xy=0; xy<laEvents.length;xy++)
							for(xy=0; xy<5;xy++)
							{
								if(lEventCounter<laEvents.length)
								{
									if(lCurrentDay==laEvents[lEventCounter].numDay)
									{
										if(lCalSize!="small")
										{
											//Response.Write("<A  TITLE=\""+laEvents[lEventCounter].Description+"\" href=\"JavaScript:newWindow('tpWindow_cal.asp?Month="+vToday.Month+"&ID="+laEvents[lEventCounter].ID+"','Calendar','width=440,height=180,toolbar=no,location=no,scrollbars=yes,directories=no,resizable=yes,menubar=no,status=no');\">");
Response.Write("<A TITLE=\""+laEvents[lEventCounter].Description+"\" href=\""+"tpInput_cal.asp"+"?numDay="+laEvents[lEventCounter].numDay+"&CalDate="+lToday.NormalDate+"\">");
											Response.Write(laEvents[lEventCounter].Subject);
											Response.Write(laEvents[lEventCounter].Location);

											//Response.Write("</A><BR>");
											Response.Write("</A> ");
										}
										lEventCounter++;

									}
									else if(xy==0)	// Only write out the spacer the first time through
										Response.Write(lSpacerText);
								}
							}
						}
						else
						Response.Write(lSpacerText);

						//if(!(lEventCounter<laEvents.length-1))
						//	Response.Write(lSpacerText);
					}
				}else{
					Response.Write(lSpacerText);
				}



					Response.Write("</TD>\n");
			}
			Response.Write("</TR>\n");
		}
	}
	Response.Write("</TABLE>\n\n");
}

// The below functions are modified versions of client side functions produced by Matt Kruse <mkruse@netexpress.net>
// WWW: http://www.mattkruse.com/

 function isDate(val)
 {
  // returns true if the string passed is a valid date.
  var sep1 = parseInt(val.indexOf("/"));
  var sep2 = parseInt(val.indexOf("/",sep1+1));
  var len = parseInt(val.length);
  var mm = parseInt(val.substr(0,sep1));
  var dd = parseInt(val.substr(sep1+1,sep2-sep1-1));
  var yy = parseInt(val.substr(sep2+1,len-sep2-1));
  if (yy<1900) yy+=2000;
  var leap = ((yy == (parseInt(yy/4) * 4)) && !(yy == (parseInt(yy/100) * 100)));
  if (!((mm >= 1) && (mm <= 12))) return false;
  if ((mm == 2) && (leap)) dom = 29;
  if ((mm == 2) && !(leap)) dom = 28;
  if ((mm == 1) || (mm == 3) || (mm == 5) || (mm == 7) || (mm == 8) || (mm == 10) || (mm == 12))
  	dom = 31;
  if ((mm == 4) || (mm == 6) || (mm == 9) || (mm == 11))
  	dom = 30;
  if (dd > dom) return false;
  return true;
 }

function whenIs(anyDate, n){
   //-- Returns the date that is n days from any date object.
   var newDate = new Date();
   newDate.setTime(anyDate.getTime()+(n*1000*60*60*24));
   return newDate;
}

function mdy(anyDate, advanceMonth) {
	// Convert to MM/DD/YYYY
	if(advanceMonth!=null)
		return anyDate.getMonth()+1+advanceMonth+"/1/"+anyDate.getFullYear();
	else
		return anyDate.getMonth()+1+"/"+anyDate.getDate()+"/"+anyDate.getFullYear();
}

function daysBetween(earlyDate,laterDate){
   //-- Returns number of days between two date objects.
   var earlySecs=earlyDate.getTime();
   var laterSecs=laterDate.getTime();
   return Math.floor ((((((laterSecs-earlySecs)/1000)/60)/60)/24));
}
%>
