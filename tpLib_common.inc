<%
/*
ThoughtProcess.net Jscript ASP Common Functions
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
<%
function fProcessRequestVars(lRequestVars)
{
	var laRequestVars = lRequestVars.split("|");
	for(var x=0; x<laRequestVars.length; x++)
		eval("if(Request('"+laRequestVars[x]+"').Count)	v"+laRequestVars[x]+"=''+Request('"+laRequestVars[x]+"');");
}

function fCreateGenericForm(lFields)
{
	var laFields = lFields.split("|");
	Response.Write("<form action='' method='get'>\n");
	for(var x=0; x<laFields.length; x++)
		Response.Write("<input type='text' name='"+laFields[x]+"' value=''> "+laFields[x]+"<BR>\n");
	Response.Write("<input type='Submit' value='Submit'>\n");
	Response.Write("</form>\n");
}
%>