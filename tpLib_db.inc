<%
/*
ThoughtProcess.net Jscript ASPDatabase Functions
Copyright (C) 2001
Author: Matt Kaatman
Date: 6/18/2001

This program is free software; you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation; version 2 of the License.
This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.
You should have received a copy of the GNU General Public License along with this program; if not, write to the Free Software Foundation, Inc., 675 Mass Ave, Cambridge, MA 02139, USA.

I'd love to hear where and how you're using it. If you'd like to report, send e-mail to asp@ThoughtProcess.net
Code contributions can be submitted via SourceForge.com.
*/

function fRemoveRecords(lConn, lTableName, lRedirPage)
{
	var vaDel = new Array();
	var vSQL="";
	if(Request("Del").Count)	vaDel	= (""+Request("Del")).split(",");
	for(var x=0; x<vaDel.length; x++)
	{
		vSQL="DELETE * FROM "+lTableName+" where (ID = " + vaDel[x] + ");";
		//Response.Write(vSQL+"<BR>");
		lConn.Execute(vSQL);
	}
	lConn.Close();
	lConn = null;
	Response.Redirect(lRedirPage);
}

function fCreateRS(lConn, lPath, lName, lSQL)
{
	//Response.Write(lSQL+"<BR>");
	lConnect = "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" + Server.Mappath(lPath+lName);
	lConn.Open(lConnect);
	var lRS		= lConn.Execute(lSQL);
	return lRS;
}

function fDestroyRS(lRS)
{
	lRS.Close();
	lRS = null;
}

function fSetFields(lFields, lRS)
{
	var laFields = lFields.split("|");
	for(var x=0; x<laFields.length; x++)
		eval("lRS('"+laFields[x]+"')=v"+laFields[x]+";");
}

function fGetTableData(lTableName, lDBPath, lDBName, lSQL, lPrintData, lBuildObject)
{
	if((lBuildObject==null)||(lBuildObject!=true))	lBuildObject=false;
	if((lPrintData==null)||(lPrintData!=false))	lPrintData=true;
	if((lSQL==null)||(lSQL==""))	lSQL= "SELECT * FROM "+lTableName+";";
	var lConn	= Server.CreateObject("ADODB.Connection");
	var lRS = Server.CreateObject('ADODB.Recordset');
	var lRS	= fCreateRS(lConn, lDBPath, lDBName, lSQL);
//	lRS = dbConn.Execute(lSQL);
	var laFieldNames = new Array();
	var enumField = new Enumerator(lRS.Fields);
	for (;!enumField.atEnd();enumField.moveNext())
	{
		//enumField.item()
		objField = enumField.item();
		//Response.Write("<td>" + objField.Name+ ": "+ objField.Value+ "</td>");
		laFieldNames[laFieldNames.length]=objField.Name;
	}
	var laTableData = new Array();
	while (!lRS.EOF)
	{
		//for(var x=0; x<laFieldNames.length;x++)
		//{
			laTableData[laTableData.length]=new fDataOBJ(laFieldNames, lRS);
/*
			for(y=0; y<laFieldNames.length;y++)
			{
				eval("laTableData[laTableData.length].ID=''");
				//eval("laTableData[laTableData.length]."+laFieldNames[y]+"=\"\"+lRS('"+laFieldNames[y]+"');");
			}
*/
		//}
		lRS.MoveNext();
	}
	lRS.Close();
	lRS=null;
	if(lPrintData!=false)
		fDisplayObj(laTableData, laFieldNames);
	if(lBuildObject!=false)
	{
		eval(fBuildObject(lTableName, laFieldNames));
		fGetData(lTableName, lSQL, laFieldNames);
	}
	return laTableData;
}

function fDataOBJ(laFieldNames, lRS)
{
	var lValue="";
//	Response.Write(laFieldNames);
	for(var x=0; x<laFieldNames.length; x++)
	{
		//Response.Write(laFieldNames[x]+"<BR>");
		lValue=""+escape(lRS(laFieldNames[x]));
		//eval("this."+laFieldNames[x]+"='';");
		eval("this."+laFieldNames[x]+"='"+lValue+"';");
//		Response.Write(x+" ");
//		eval("Response.Write(this."+laFieldNames[x]+"+\'<BR>');");
	}
}

function fDisplayObj(laTableData, laFieldNames)
{
	Response.Write("<form>");
	Response.Write("<table width='100%' border=1><tr>\n")
	//Response.Write("<TH>x</TH>");
	for(x=0; x<laFieldNames.length; x++)	Response.Write("<tH>" + laFieldNames[x] + "</tH>\n");
	Response.Write("</TR>");
	for(y=0; y<laTableData.length; y++)
	{
		Response.Write("<TR>");
		//Response.Write("<TD><input type=\"checkbox\" name=\"Del"+laTableData[y].ID+"\"></TD>");
		for(x=0; x<laFieldNames.length; x++)
		{
			Response.Write("<td>" + unescape(eval("laTableData[y]."+laFieldNames[x])) + "</td>\n");
		}
		Response.Write("</TR>\n");
	}
	Response.Write("</table>\n");
	Response.Write("<input type=\"submit\" value=\"Delete\">\n");
	Response.Write("</form>\n\n");
}

function fCreateTableArr(lTables)
{
// Set current table & create array of table names
	vaTables = lTables.split("|");
	var lTableName = "";
	if(Request("Table").Count)
		lTableName=""+Request("Table");
	else
		lTableName=vaTables[0];
	return lTableName;
}

function fBuildObject(lTableName, laFieldNames){
	lTempObjStr="function f"+lTableName+"OBJ("+laFieldNames+"){";
	for(y=0; y<laFieldNames.length; y++)
		lTempObjStr+="\n\tif("+laFieldNames[y]+"=='null')	"+laFieldNames[y]+"='';";
	for(y=0; y<laFieldNames.length; y++)
		lTempObjStr+="\n\tthis."+laFieldNames[y]+"=''+"+laFieldNames[y]+";";
	lTempObjStr+="\n}";
	Response.Write("\n\n<!--\n"+lTempObjStr+"\n");
	return lTempObjStr;
}

function fGetData(lTableName, lSQL, laFieldNames)
{
	tempStr="";
	tempStr+="\n\tif((lSQL==null)||(lSQL==\"\"))	lSQL= \"SELECT * FROM \"+lTableName+\";\";"
	tempStr+=("\n\tvar l"+lTableName+"RS = Server.CreateObject('ADODB.Recordset');");
	tempStr+=("\n\tl"+lTableName+"RS = dbConn.Execute(lSQL);");
	lTempObjStr="\tga"+lTableName+"ARR[ga"+lTableName+"ARR.length] = new f"+lTableName+"OBJ(";
	for(y=0; y<laFieldNames.length; y++){
		lTempObjStr+="\n\t\t\"\"+l"+lTableName+"RS('"+laFieldNames[y]+"')";
		if(laFieldNames.length!=y+1)
			lTempObjStr+=",";
		else
			lTempObjStr+="\n\t\t)";
	}
	tempStr+=("\n\t"+"while (!l"+lTableName+"RS.EOF){\n\t"+lTempObjStr+";\n\t\t"+"l"+lTableName+"RS.MoveNext();\n\t}");
	tempStr+=("\n\tl"+lTableName+"RS.Close();");
	tempStr+=("\n\tl"+lTableName+"RS=null;");
	Response.Write("\n\nfunction fGet"+lTableName+"Data(lTableName, lSQL){"+tempStr+"\n}\n-->\n\n");
	//eval(tempStr);
}
%>