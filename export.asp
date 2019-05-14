<%@ codepage="65001" Language=VBScript %>
<%
Option Explicit
Response.Buffer = TRUE
Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename =waihudata.xls"
%>
<html>
<head>
<meta http-equiv="content-type" content="text/html;charset=utf-8">
</head>
<body>
<%
Dim objconn,objrecord
set objconn=server.createobject("adodb.connection")
	  objconn.open "provider=microsoft.jet.oledb.4.0; data source=" & server.mappath("calendar.mdb")
	  set objrecord=server.createobject("adodb.recordset")

set objrecord=objconn.execute("select location as 分校,subject as 数据总量,description as 项目说明, Format(strdate,'yyyy-mm-dd') as 日期 from events  order by location")

'set objrecord=objconn.execute("select location as 分校,subject as 数据总量,description as 项目说明, Format(strdate,'yyyy-mm-dd') as 日期 from events  where nummonth like month(now) order by location")

if objrecord.eof then
	response.write "没有筛选到有效数据"

end if
%>
<table width="1000" border="1">
<tr>
<th width="50"> <div align="center">分校</div></th>
<th width="70"> <div align="center">数据总量</div></th>
<th width="200"> <div align="center">项目说明 </div></th>
<th width="70"> <div align="center">日期</div></th>


</tr>
<%
While Not objrecord.EOF
%>
<tr>
<td nowrap="nowrap"><div ><%=objrecord.Fields("分校").Value%></div></td>
<td nowrap="nowrap"><%=objrecord.Fields("数据总量").Value%></td>
<td nowrap="nowrap"><%=objrecord.Fields("项目说明").Value%></td>
<td nowrap="nowrap"><div ><%=objrecord.Fields("日期").Value%></div></td>

</tr>
<%
objrecord.MoveNext
Wend
%>
</table>
<%
objrecord.Close()

Set objrecord = Nothing

%>      
</body>
</html>