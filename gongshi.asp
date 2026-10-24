<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<% Session.CodePage=936 %>
<!--#include file="Connections/sjk.asp" -->
<%
Dim rs
Set rs=Server.CreateObject("ADODB.Recordset")
rs.Open "select school,qq,wanmei_name from player order by school",MM_sjk_STRING,1,1
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>张家港市高校CS联赛选手公示</title>
<style>
body{font-family:"Microsoft YaHei";background:#f5f7fa;margin:0;padding:30px 15px;}
.box{max-width:900px;margin:0 auto;background:#fff;padding:25px;border-radius:8px;box-shadow:0 1px 4px #ccc;}
h2{text-align:center;color:#1a3b70;margin-top:0;}
table{width:100%;border-collapse:collapse;margin-top:20px;}
th,td{border:#b8c8e0 1px solid;padding:10px;text-align:center;}
th{background:#234b94;color:white;}
tr:nth-child(even){background:#f2f6fc;}
</style>
</head>
<body>
<div class="box">
	<h2>张家港市高校CS联赛选手信息公示</h2>
	<table>
		<tr>
			<th>所属学校</th>
			<th>QQ号码</th>
			<th>完美世界账号名称</th>
		</tr>
		<% Do While Not rs.EOF %>
		<tr>
			<td><%=Server.HTMLEncode(rs("school"))%></td>
			<td><%=Server.HTMLEncode(rs("qq"))%></td>
			<td><%=Server.HTMLEncode(rs("wanmei_name"))%></td>
		</tr>
		<% rs.MoveNext
		Loop %>
	</table>
</div>
</body>
</html>
<%
rs.Close
Set rs=Nothing
%>
