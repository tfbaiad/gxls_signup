<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% Session.CodePage = 65001 ' 统一服务端UTF-8编码，避免中文乱码 %>
<!--#include file="Connections/sjk.asp" -->
<%
Dim MM_editAction
MM_editAction = CStr(Request.ServerVariables("SCRIPT_NAME"))
If (Request.QueryString <> "") Then
  MM_editAction = MM_editAction & "?" & Server.HTMLEncode(Request.QueryString)
End If
' boolean to abort record edit
Dim MM_abortEdit
MM_abortEdit = false
%>
<%
If (CStr(Request("MM_insert")) = "forml") Then
  If (Not MM_abortEdit) Then
    ' execute the insert
    Dim MM_editCmd
    Set MM_editCmd = Server.CreateObject ("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_sjk_STRING
    MM_editCmd.CommandText = "INSERT INTO [汇总] ([选手完美名称], [选手Steam64位ID], [选手QQ], [选手备用手机号], [选手完美段位], [选手学校]) VALUES (?, ?, ?, ?, ?, ?)" 
    MM_editCmd.Prepared = true
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 201, 1, -1, Request.Form("选手完美名称")) ' adLongVarChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param2", 201, 1, -1, Request.Form("选手Steam64位ID")) ' adLongVarChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param3", 201, 1, -1, Request.Form("选手QQ")) ' adLongVarChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param4", 201, 1, -1, Request.Form("选手备用手机号")) ' adLongVarChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param5", 201, 1, -1, Request.Form("选手完美段位")) ' adLongVarChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param6", 201, 1, -1, Request.Form("选手学校")) ' adLongVarChar
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close
    ' append the query string to the redirect URL
    Dim MM_editRedirectUrl
    MM_editRedirectUrl = "tips.html"
    If (Request.QueryString <> "") Then
      If (InStr(1, MM_editRedirectUrl, "?", vbTextCompare) = 0) Then
        MM_editRedirectUrl = MM_editRedirectUrl & "?" & Request.QueryString
      Else
        MM_editRedirectUrl = MM_editRedirectUrl & "&" & Request.QueryString
      End If
    End If
    Response.Redirect(MM_editRedirectUrl)
  End If
End If
%>
<%
' ========== 表单提交路径 ==========
MM_editAction = CStr(Request.ServerVariables("SCRIPT_NAME"))
If Request.QueryString <> "" Then
  MM_editAction = MM_editAction & "?" & Server.HTMLEncode(Request.QueryString)
End If
' ========== 终止编辑标记 ==========
MM_abortEdit = false
' ========== 表单提交插入处理 ==========
If CStr(Request("MM_insert")) = "forml" Then
  If Not MM_abortEdit Then
    Set MM_editCmd = Server.CreateObject("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_sjk_STRING
    MM_editCmd.CommandText = "INSERT INTO [汇总] ([选手完美名称], [选手Steam64位ID], [选手完美段位], [选手QQ], [选手备用手机号], [选手学校]) VALUES (?, ?, ?, ?, ?, ?)"
    MM_editCmd.Prepared = true
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 201, 1, -1, Request.Form("选手完美名称"))
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param2", 201, 1, -1, Request.Form("选手Steam64位ID"))
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param3", 201, 1, -1, Request.Form("选手完美段位"))
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param4", 201, 1, -1, Request.Form("选手QQ"))
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param5", 201, 1, -1, Request.Form("选手备用手机号"))
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param6", 201, 1, -1, Request.Form("选手学校"))
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close
    Set MM_editCmd = Nothing
    ' 提交成功跳转
    MM_editRedirectUrl = "tips.html"
    If Request.QueryString <> "" Then
      If InStr(1, MM_editRedirectUrl, "?", vbTextCompare) = 0 Then
        MM_editRedirectUrl = MM_editRedirectUrl & "?" & Request.QueryString
      Else
        MM_editRedirectUrl = MM_editRedirectUrl & "&" & Request.QueryString
      End If
    End If
    Response.Redirect(MM_editRedirectUrl)
  End If
End If
' ========== HTML转义函数 ==========
Function H(s)
    If IsNull(s) Or s = "" Then
        H = ""
    Else
        H = Server.HTMLEncode(CStr(s))
    End If
End Function
' ========== 院校列表【已新增2条】 ==========
schoolList = Array( _
    "梁丰高级中学及其国际部", "沙州中学本校区", "沙州中学北校区", _
    "暨阳高级中学", "张家港高级中学", "塘桥高级中学", _
    "崇真中学", "凤凰高级中学", "乐余高级中学", _
    "常青藤实验学校及其国际部", "外国语实验学校及其国际部", _
    "张家港第一职业学校", "张家港第二职业学校", "张家港第三职业学校", _
    "张家港开放大学", "江苏科技大学张家港校区", "沙洲工学院", _
    "张家港市艺术学校（原张家港舞蹈学校）","张家港市特殊教育学校" _
)
' ========== 段位列表 ==========
rankList = Array("D","C","C+","C++","B","B+","B++","A","A+","A++","S0-9","S10-24","S25-49","S>=50")
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>张家港市高校CS联赛选手信息填报</title>
<style>
  :root{ --ink:#1b1b1b; --main:#12305e; --line:#c9d4e4; --soft:#f4f7fb; }
  *{box-sizing:border-box;}
  body{
    margin:0; padding:40px 16px 56px; background:#eceff5; color:var(--ink);
    font-family:'Songti SC','SimSun','Noto Serif CJK SC',serif;
    font-size:15px; line-height:1.9; -webkit-font-smoothing:antialiased;
  }
  .paper{
    position:relative; max-width:820px; margin:0 auto; background:#fff;
    padding:56px 64px 44px; border-radius:4px;
    box-shadow:0 10px 34px rgba(20,45,90,.13);
  }
  .paper::before{
    content:""; position:absolute; top:0; left:0; right:0; height:6px;
    background:linear-gradient(90deg,#12305e,#2f6fc4 55%,#57b0e8);
    border-radius:4px 4px 0 0;
  }
  .doc-no{
    text-align:right; font-family:'PingFang SC','Microsoft YaHei',sans-serif;
    font-size:13px; color:#5b6779; letter-spacing:.5px; margin-bottom:20px;
  }
  h1{
    font-family:'PingFang SC','Microsoft YaHei',sans-serif;
    font-size:25px; line-height:1.55; letter-spacing:2px;
    color:var(--main); text-align:center; font-weight:700; margin:0 0 8px;
  }
  .subtitle{
    text-align:center; font-size:13px; letter-spacing:7px; color:#8b98ad;
    margin:0 0 30px; font-family:'PingFang SC','Microsoft YaHei',sans-serif;
  }
  .form-block{
    margin:24px 0 8px; padding:18px 22px;
    background:var(--soft); border:1px solid var(--line); border-radius:6px;
    font-family:'PingFang SC','Microsoft YaHei',sans-serif; font-size:14px;
  }
  .form-block p{margin:0 0 10px; font-weight:700; color:var(--main); font-size:15px;}
  .rank-list{list-style:none; padding-left:0; margin:0; columns:4; column-gap:16px;}
  .rank-list li{margin-bottom:8px; break-inside:avoid;}
  .rank-list label{margin-left:6px; cursor:pointer;}
  .school-list-ul{list-style:none; padding-left:0; margin:0; columns:2; column-gap:32px;}
  .school-list-ul li{margin-bottom:8px; break-inside:avoid;}
  .school-list-ul label{margin-left:6px; cursor:pointer;}
  table.form-table{
    width:100%; border-collapse:collapse; margin:22px 0 6px;
    font-size:14.5px; line-height:1.7;
  }
  table.form-table caption{
    caption-side:top; text-align:left;
    font-family:'PingFang SC','Microsoft YaHei',sans-serif;
    font-size:15px; font-weight:700; color:var(--main); padding:0 0 10px;
  }
  table.form-table td{border:1px solid var(--line); padding:10px 12px; vertical-align:middle;}
  table.form-table td.lbl{
    width:200px; background:var(--soft); color:#31456b;
    font-family:'PingFang SC','Microsoft YaHei',sans-serif; font-size:14px; font-weight:600;
  }
  input[type=text]{
    width:100%; max-width:420px; padding:8px 10px;
    border:1px solid var(--line); border-radius:4px;
    font-family:inherit; font-size:14.5px; line-height:1.5;
  }
  input[type=text]:focus{outline:2px solid #9cc2ee; outline-offset:1px; border-color:#2f6fc4;}
  .btn-row{margin-top:22px; text-align:center;}
  .btn-row input[type=submit]{
    font-family:'PingFang SC','Microsoft YaHei',sans-serif;
    font-size:15px; color:#fff; background:#12305e; border:0;
    padding:11px 42px; border-radius:6px; cursor:pointer; letter-spacing:2px;
  }
  .btn-row input[type=submit]:hover{background:#1b4b8f;}
  .btn-row .cancel-link{
    margin-left:24px; font-family:'PingFang SC','Microsoft YaHei',sans-serif;
    color:#5b6779; text-decoration:none; font-size:14px;
  }
  .btn-row .cancel-link:hover{color:var(--main);}
  .note{
    margin:26px 0 0; padding-top:16px; border-top:1px solid var(--line);
    font-size:13px; color:#6b788c; line-height:1.85;
    font-family:'PingFang SC','Microsoft YaHei',sans-serif;
  }
  @media (max-width:640px){
    body{font-size:14.5px; padding:16px 10px 40px;}
    .paper{padding:32px 22px 30px;}
    h1{font-size:20px; letter-spacing:1px;}
    .rank-list{columns:2;}
    .school-list-ul{columns:1;}
    table.form-table{font-size:13.5px;}
    table.form-table td{padding:7px 8px;}
    table.form-table td.lbl{width:auto;}
  }
</style>
</head>
<body>
<div class="paper">
  <div class="doc-no">报名编号：ZJG-CS-2027</div>
  <h1>张家港市高校 CS 联赛<br>选手信息收集表  </h1>
  <form name="forml" method="POST" action="<%=MM_editAction%>">
    <table class="form-table">
      <caption>基本信息</caption>
      <tr>
        <td class="lbl">选手完美名称</td>
        <td><input type="text" name="选手完美名称" id="选手完美名称" /></td>
      </tr>
      <tr>
        <td class="lbl">选手 Steam 64 位 ID</td>
        <td><input type="text" name="选手Steam64位ID" id="选手Steam64位ID"  /></td>
      </tr>
      <tr>
        <td class="lbl">选手 QQ</td>
        <td><input type="text" name="选手QQ" id="选手QQ"  /></td>
      </tr>
      <tr>
        <td class="lbl">选手备用手机号</td>
        <td><input type="text" name="选手备用手机号" id="选手备用手机号"  /></td>
      </tr>
    </table>
    <div class="form-block">
      <p>选手完美段位（历史最高）</p>
<p>
    <input type="radio" name="选手完美段位" id="radio" value="D" />
    <label for="radio">D</label>
  </p>
  <p>
  <input type="radio" name="选手完美段位" id="radio2" value="C" />
    <label for="radio2">C</label>
    <input type="radio" name="选手完美段位" id="radio3" value="C+" />
    <label for="radio3">C+</label>
    <input type="radio" name="选手完美段位" id="radio4" value="C++" />
    <label for="radio4">C++</label>
  </p>
  <p>
    <input type="radio" name="选手完美段位" id="radio5" value="B" />
    <label for="radio5">B</label>
    <input type="radio" name="选手完美段位" id="radio6" value="B+" />
    <label for="radio6">B+</label>
    <input type="radio" name="选手完美段位" id="radio7" value="B++" />
    <label for="radio7">B++</label>
  </p>
  <p>
  <input type="radio" name="选手完美段位" id="radio8" value="A" />
    <label for="radio8">A</label>
    <input type="radio" name="选手完美段位" id="radio9" value="A+" />
    <label for="radio9">A+</label>
    <input type="radio" name="选手完美段位" id="radio10" value="A++" />
    <label for="radio10">A++</label>
  </p>
  <p>
  <input type="radio" name="选手完美段位" id="radio11" value="S0-9" />
    <label for="radio11">S0-9</label>
    <input type="radio" name="选手完美段位" id="radio12" value="S10-24" />
    <label for="radio12">S10-24</label>
    <input type="radio" name="选手完美段位" id="radio13" value="S25-49" />
    <label for="radio13">S25-49</label>
    <input type="radio" name="选手完美段位" id="radio14" value="S>=50" />
    <label for="radio14">S&gt;=50</label></p></div>
    <div class="form-block">
      <p>选手所在院校</p><p>
    <input type="radio" name="选手学校" id="school1" value="梁丰高级中学及其国际部" />
    <label for="school1">梁丰高级中学及其国际部</label>
  </p>
  <p>
    <input type="radio" name="选手学校" id="school2" value="沙州中学本校区" />
    <label for="school2">沙州中学本校区</label>
  </p>
  <p>
    <input type="radio" name="选手学校" id="school3" value="沙州中学北校区" />
    <label for="school3">沙州中学北校区</label>
  </p>
  <p>
    <input type="radio" name="选手学校" id="school4" value="暨阳高级中学" />
    <label for="school4">暨阳高级中学</label>
  </p>
  <p>
    <input type="radio" name="选手学校" id="school5" value="张家港高级中学" />
    <label for="school5">张家港高级中学</label>
  </p>
  <p>
    <input type="radio" name="选手学校" id="school6" value="塘桥高级中学" />
    <label for="school6">塘桥高级中学</label>
  </p>
  <p>
    <input type="radio" name="选手学校" id="school7" value="崇真中学" />
    <label for="school7">崇真中学</label>
  </p>
  <p>
    <input type="radio" name="选手学校" id="school8" value="凤凰高级中学" />
    <label for="school8">凤凰高级中学</label>
  </p>
  <p>
    <input type="radio" name="选手学校" id="school9" value="乐余高级中学" />
    <label for="school9">乐余高级中学</label>
  </p>
  <p>
    <input type="radio" name="选手学校" id="school10" value="常青藤实验学校及其国际部" />
    <label for="school10">常青藤实验学校及其国际部</label>
  </p>
  <p>
    <input type="radio" name="选手学校" id="school11" value="外国语实验学校及其国际部" />
    <label for="school11">外国语实验学校及其国际部</label>
  </p>
  <p>
    <input type="radio" name="选手学校" id="school12" value="张家港第一职业学校" />
    <label for="school12">张家港第一职业学校</label>
  </p>
  <p>
    <input type="radio" name="选手学校" id="school13" value="张家港第二职业学校" />
    <label for="school13">张家港第二职业学校</label>
  </p>
  <p>
    <input type="radio" name="选手学校" id="school14" value="张家港第三职业学校" />
    <label for="school14">张家港第三职业学校</label>
  </p>
  <p>
    <input type="radio" name="选手学校" id="school15" value="张家港开放大学" />
    <label for="school15">张家港开放大学</label>
  </p>
  <p>
    <input type="radio" name="选手学校" id="school16" value="江苏科技大学张家港校区" />
    <label for="school16">江苏科技大学张家港校区</label>
  </p>
  <p>
    <input type="radio" name="选手学校" id="school17" value="沙洲工学院" />
    <label for="school17">沙洲工学院</label>
  </p>
  <!-- 新增2个院校单选框 -->
  <p>
    <input type="radio" name="选手学校" id="school18" value="张家港市艺术学校（原张家港舞蹈学校）" />
    <label for="school18">张家港市艺术学校（原张家港舞蹈学校）</label>
  </p>
  <p>
    <input type="radio" name="选手学校" id="school19" value="张家港市特殊教育学校" />
    <label for="school19">张家港市特殊教育学校</label>
  </p>
</div>
    <div class="btn-row">
      <input type="submit" value="提 交 信 息" />
    </div>
    <input type="hidden" name="MM_insert" value="forml">
  </form>
  <p class="note">
    注：请确保所填信息真实有效，赛事组委会将对报名信息进行核验。
    信息如有变更请及时联系组委会更新，以免影响正常参赛。
  </p>
</div>
</body>
</html>
