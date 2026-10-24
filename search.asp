<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<% Session.CodePage = 936 %>
<!--#include file="Connections/sjk.asp" -->
<%
searchNum = Trim(Request.Form("searchNum"))
errMsg = ""
showResult = False

Function SQLStr(s)
    SQLStr = Replace(s, "'", "''")
End Function

Function H(s)
    If IsNull(s) Or s = "" Then
        H = "缺失"
    Else
        H = Server.HTMLEncode(CStr(s))
    End If
End Function

If searchNum <> "" Then
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT * FROM [汇总] WHERE [统一编号]='" & SQLStr(searchNum) & "'"
    rs.Open sql, MM_sjk_STRING, 1, 1
    
    If rs.EOF And rs.BOF Then
        errMsg = "未找到对应编号的记录，请核对编号是否正确"
    Else
        showResult = True
    End If
End If
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>参赛信息统一查询</title>
<style>
  :root{ --ink:#1b1b1b; --main:#12305e; --line:#c9d4e4; --soft:#f4f7fb; }
  *{box-sizing:border-box;}
  body{
    margin:0; padding:40px 16px 56px; background:#eceff5; color:var(--ink);
    font-family:'Songti SC','SimSun','Noto Serif CJK SC',serif;
    font-size:15px; line-height:1.9;
  }
  .paper{
    position:relative; max-width:780px; margin:0 auto; background:#fff;
    padding:40px 48px 36px; border-radius:4px;
    box-shadow:0 10px 34px rgba(20,45,90,.13);
  }
  .paper::before{
    content:""; position:absolute; top:0; left:0; right:0; height:6px;
    background:linear-gradient(90deg,#12305e,#2f6fc4 55%,#57b0e8);
    border-radius:4px 4px 0 0;
  }
  h2{
    font-family:'PingFang SC','Microsoft YaHei',sans-serif;
    font-size:22px; color:var(--main); text-align:center; font-weight:700;
    margin:0 0 20px; letter-spacing:1px;
  }
  .search-bar{
    margin:20px 0 30px; padding:18px; background:var(--soft);
    border:1px solid var(--line); border-radius:6px;
  }
  .search-row{
    display:flex; gap:10px; align-items:center;
  }
  .search-row input{
    flex:1; padding:8px 10px; border:1px solid var(--line);
    border-radius:4px; font-size:14px;
  }
  .search-row button{
    padding:8px 24px; background:#12305e; color:#fff;
    border:0; border-radius:4px; cursor:pointer;
    font-family:'PingFang SC','Microsoft YaHei',sans-serif;
  }
  .msg-error{
    padding:12px 16px; background:#fff2f2; border:1px solid #ffccc7;
    color:#cf1322; border-radius:6px; margin-bottom:16px;
    font-family:'PingFang SC','Microsoft YaHei',sans-serif;
  }
  .result-block{
    margin-bottom:24px;
  }
  .result-block h3{
    margin:0 0 12px; font-size:17px; color:var(--main);
    font-family:'PingFang SC','Microsoft YaHei',sans-serif;
    border-left:4px solid #2f6fc4; padding-left:10px;
  }
  .info-table{
    width:100%; border-collapse:collapse;
  }
  .info-table th,.info-table td{
    border:1px solid var(--line); padding:10px 12px;
    font-family:'PingFang SC','Microsoft YaHei',sans-serif; font-size:14px;
  }
  .info-table th{
    width:160px; background:var(--soft); color:var(--main);
    text-align:left; font-weight:600;
  }
  @media (max-width:640px){
    body{padding:20px 12px 40px;}
    .paper{padding:28px 20px 24px;}
    .search-row{flex-direction:column; align-items:stretch;}
  }
</style>
</head>
<body>

<div class="paper">
  <h2>参赛信息统一查询</h2>
  
  <% If errMsg <> "" Then %>
  <div class="msg-error"><%=errMsg%></div>
  <% End If %>

  <div class="search-bar">
    <form method="POST" action="">
      <div class="search-row">
        <input type="text" name="searchNum" placeholder="请输入统一编号 例：ZJGCS2027001" value="<%=Server.HTMLEncode(searchNum)%>" />
        <button type="submit">查询</button>
      </div>
    </form>
  </div>

  <% If showResult Then %>
  <!-- 报名信息 -->
  <div class="result-block">
    <h3>报名信息</h3>
    <table class="info-table">
      <tr><th>统一编号</th><td><%=H(rs("统一编号"))%></td></tr>
      <tr><th>完美名称</th><td><%=H(rs("选手完美名称"))%></td></tr>
      <tr><th>所在院校</th><td><%=H(rs("选手学校"))%></td></tr>
      <tr><th>Steam 64位ID</th><td><%=H(rs("选手Steam64位ID"))%></td></tr>
      <tr><th>QQ号</th><td><%=H(rs("选手QQ"))%></td></tr>
      <tr><th>完美段位</th><td><%=H(rs("选手完美段位"))%></td></tr>
    </table>
  </div>

  <!-- 承诺书信息 -->
  <div class="result-block">
    <h3>承诺书信息</h3>
    <table class="info-table">
      <tr><th>是否签署</th><td><%=H(rs("是否填写承诺书"))%></td></tr>
      <tr><th>参赛队伍</th><td><%=H(rs("选手参赛队伍"))%></td></tr>
      <tr><th>签署日期</th><td><%=H(rs("日期"))%></td></tr>
    </table>
  </div>
  
  <% rs.Close : Set rs = Nothing %>
  <% End If %>

</div>

</body>
</html>