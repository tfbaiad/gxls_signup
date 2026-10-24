<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% Session.CodePage = 65001 %>
<!--#include file="Connections/sjk.asp" -->
<%
'==========登录逻辑==========
Dim loginUser,loginPwd
loginUser = Trim(Request.Form("username"))
loginPwd = Trim(Request.Form("password"))

' 初始化session
If IsEmpty(Session("admin")) Then
    Session("admin") = ""
End If

' 主管理员账号
Const ADMIN_USER = "mayunxuan"
Const ADMIN_PWD = "mayunxuan2010"

Dim isLogin,isSuperAdmin
isLogin = False
isSuperAdmin = False
If Session("admin") <> "" Then
    isLogin = True
    isSuperAdmin = (Session("admin")=ADMIN_USER)
End If

' 登录提交处理
If Request.Form("action") = "login" Then
    If loginUser = ADMIN_USER And loginPwd = ADMIN_PWD Then
        Session("admin") = ADMIN_USER
        isLogin = True
        isSuperAdmin = True
    Else
        Dim rsCheck,sqlCheck
        Set rsCheck = Server.CreateObject("ADODB.Recordset")
        sqlCheck = "SELECT * FROM [管理员] WHERE [账号]='" & Replace(loginUser,"'","''") & "' AND [密码]='" & Replace(loginPwd,"'","''") & "'"
        On Error Resume Next
        rsCheck.Open sqlCheck,MM_sjk_STRING,1,1
        On Error GoTo 0
        If Not rsCheck.EOF Then
            Session("admin") = loginUser
            isLogin = True
        End If
        rsCheck.Close
        Set rsCheck=Nothing
    End If
End If

' 退出登录
If Request.QueryString("logout")="1" Then
    Session.Abandon
    Response.Redirect("back.asp")
End If

'==========新增管理员（仅主账号可用）==========
Dim newUser,newPwd
If Request.Form("action")="addadmin" And isSuperAdmin Then
    newUser = Trim(Request.Form("newusername"))
    newPwd = Trim(Request.Form("newpassword"))
    If newUser<>"" And newPwd<>"" Then
        Dim sqlAddAdmin
        sqlAddAdmin = "INSERT INTO [管理员]([账号],[密码]) VALUES('" & Replace(newUser,"'","''") & "','" & Replace(newPwd,"'","''") & "')"
        Dim cmdAdd
        Set cmdAdd = Server.CreateObject("ADODB.Command")
        cmdAdd.ActiveConnection = MM_sjk_STRING
        cmdAdd.CommandText = sqlAddAdmin
        cmdAdd.Execute
        Set cmdAdd=Nothing
    End If
End If

'HTML转义函数
Function H(s)
    If IsNull(s) Or s="" Then
        H=""
    Else
        H=Server.HTMLEncode(CStr(s))
    End If
End Function
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>后台 - 选手数据查看</title>
<style>
  :root{ --ink:#1b1b1b; --main:#12305e; --line:#c9d4e4; --soft:#f4f7fb; }
  *{box-sizing:border-box;}
  body{
    margin:0; padding:40px 16px 56px; background:#eceff5; color:var(--ink);
    font-family:'PingFang SC','Microsoft YaHei','Noto Sans CJK SC',sans-serif;
    font-size:14px; line-height:1.7;
  }
  .paper{
    position:relative; max-width:1150px; margin:0 auto; background:#fff;
    padding:48px 40px; border-radius:4px;
    box-shadow:0 10px 34px rgba(20,45,90,.13);
  }
  .paper::before{
    content:""; position:absolute; top:0; left:0; right:0; height:6px;
    background:linear-gradient(90deg,#12305e,#2f6fc4 55%,#57b0e8);
    border-radius:4px 4px 0 0;
  }
  h1{
    font-size:24px; line-height:1.5; letter-spacing:2px;
    color:var(--main); font-weight:700; margin:0 0 24px; text-align:center;
  }
  table.data-table{
    width:100%; border-collapse:collapse; margin:16px 0;
  }
  table.data-table th{
    background:var(--main); color:#fff; padding:11px 8px;
    border:1px solid #10284b; text-align:center; font-weight:500;
  }
  table.data-table td{
    border:1px solid var(--line); padding:9px 8px; text-align:center;
  }
  table.data-table tr:nth-child(even){
    background:var(--soft);
  }
  .back-box{
    margin-top:24px; text-align:center;
  }
  .back-btn{
    display:inline-block; padding:11px 28px;
    background:var(--main); color:#fff; text-decoration:none;
    border-radius:6px; font-size:15px; letter-spacing:1px; margin:4px;
    border:none;cursor:pointer;
  }
  .back-btn:hover{
    background:#1b4b8f;
  }
  .login-box{max-width:420px;margin:30px auto;border:1px solid var(--line);padding:30px;border-radius:6px;background:var(--soft);}
  .login-box p{margin:12px 0;}
  .login-box input{width:100%;padding:9px;border:1px solid var(--line);border-radius:4px;font-size:14px;}
  .admin-form{margin:30px 0;padding:20px;border:1px solid var(--line);background:#f8fafe;border-radius:6px;}
  .empty-tip{
    text-align:center; padding:40px; color:#6b788c; font-size:16px;
  }
  .yes{color:#23856d;font-weight:bold;}
  .no{color:#c53030;font-weight:bold;}
  @media (max-width:768px){
    .paper{padding:24px 14px;}
    table.data-table{font-size:12px;}
    table.data-table th,table.data-table td{padding:6px 4px;}
  }
</style>
</head>
<body>
<div class="paper">
<%If Not isLogin Then%>
    <h1>后台登录</h1>
    <form class="login-box" method="post">
        <input type="hidden" name="action" value="login">
        <p>账号：<input type="text" name="username"></p>
        <p>密码：<input type="password" name="password"></p>
        <p style="text-align:center"><input class="back-btn" type="submit" value="登录"></p>
    </form>
<%Else%>
    <h1>选手报名数据后台</h1>
    <div style="text-align:right;">
        <span>当前登录：<%=H(Session("admin"))%></span>
        <a href="?logout=1" class="back-btn">退出登录</a>
    </div>

    <%If isSuperAdmin Then%>
        <div class="admin-form">
            <h3>新增管理员账户（仅主账号可用）</h3>
            <form method="post">
                <input type="hidden" name="action" value="addadmin">
                <p>新管理员账号：<input type="text" name="newusername"></p>
                <p>新管理员密码：<input type="password" name="newpassword"></p>
                <input class="back-btn" type="submit" value="添加管理员">
            </form>
        </div>
    <%End If%>

    <%
    Dim rs,sql
    Set rs = Server.CreateObject("ADODB.Recordset")
    '查询汇总表，增加【是否填写承诺书】字段
    sql = "SELECT * FROM [汇总] ORDER BY [选手完美名称]"
    rs.Open sql,MM_sjk_STRING,1,1
    %>
    <%If rs.EOF And rs.BOF Then%>
        <div class="empty-tip">暂无报名数据</div>
    <%Else%>
    <table class="data-table">
        <tr>
          <th>选手完美名称</th>
          <th>Steam64位ID</th>
          <th>完美段位</th>
          <th>QQ</th>
          <th>备用手机号</th>
          <th>学校</th>
          <th>是否填写承诺书</th>
        </tr>
        <%Do While Not rs.EOF%>
        <tr>
          <td><%=H(rs("选手完美名称"))%></td>
          <td><%=H(rs("选手Steam64位ID"))%></td>
          <td><%=H(rs("选手完美段位"))%></td>
          <td><%=H(rs("选手QQ"))%></td>
          <td><%=H(rs("选手备用手机号"))%></td>
          <td><%=H(rs("选手学校"))%></td>
          <td>
              <%If Trim(rs("是否填写承诺书"))="是" Then%>
                <span class="yes">已填写</span>
              <%Else%>
                <span class="no">未填写</span>
              <%End If%>
          </td>
        </tr>
        <%rs.MoveNext
        Loop%>
    </table>
    <%End If%>
    <%rs.Close
    Set rs=Nothing%>

    <div class="back-box">
        <a href="first.html" class="back-btn">返回导航页</a>
    </div>
<%End If%>
</div>
</body>
</html>
