<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<% Session.CodePage = 936 %>
<!--#include file="Connections/sjk.asp" -->
<%
On Error Resume Next
' ========== 管理员权限校验 ==========
Dim isAdmin
isAdmin = False
If Not IsEmpty(Session("admin")) And Session("admin") <> "" Then
    isAdmin = True
End If

' ========== 保存修改 ==========
If isAdmin And Request.Form("action") = "save" Then
    Dim matchCount, i, mId, t1, s1, t2, s2, sqlUp, cmdUp
    matchCount = CInt(Request.Form("matchcount"))
    For i = 1 To matchCount
        mId = Request.Form("id_" & i)
        t1 = Replace(Trim(Request.Form("t1_" & i)), "'", "''")
        s1 = Trim(Request.Form("s1_" & i))
        t2 = Replace(Trim(Request.Form("t2_" & i)), "'", "''")
        s2 = Trim(Request.Form("s2_" & i))
        If s1 = "" Then s1 = "0"
        If s2 = "" Then s2 = "0"
        
        sqlUp = "UPDATE [对阵表] SET [队1]='" & t1 & "',[分1]=" & s1 & ",[队2]='" & t2 & "',[分2]=" & s2 & ",[赛制]='" & Request.Form("sai_" & i) & "' WHERE ID=" & mId
        Set cmdUp = Server.CreateObject("ADODB.Command")
        cmdUp.ActiveConnection = MM_sjk_STRING
        cmdUp.CommandText = sqlUp
        cmdUp.Execute
        Set cmdUp = Nothing
    Next
End If

' ========== 读取对阵数据 ==========
Dim rs, sql
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT * FROM [对阵表] ORDER BY [ID]"
rs.Open sql, MM_sjk_STRING, 1, 3

' HTML转义
Function H(s)
    If IsNull(s) Or s = "" Then
        H = ""
    Else
        H = Server.HTMLEncode(CStr(s))
    End If
End Function

' 按分组输出比赛的函数
Sub RenderGroup(groupName)
    Dim hasMatch
    hasMatch = False
    rs.MoveFirst
    Do While Not rs.EOF
        If rs("组别") = groupName Then
            hasMatch = True
            %>
            <div class="match-card">
              <div class="match-teams">
                <% If isAdmin Then %>
                <div class="team"><input type="text" name="t1_<%=rs("ID")%>" value="<%=H(rs("队1"))%>"></div>
                <div class="score"><input type="text" name="s1_<%=rs("ID")%>" value="<%=H(rs("分1"))%>"></div>
                <div class="vs">VS</div>
                <div class="score"><input type="text" name="s2_<%=rs("ID")%>" value="<%=H(rs("分2"))%>"></div>
                <div class="team"><input type="text" name="t2_<%=rs("ID")%>" value="<%=H(rs("队2"))%>"></div>
                <input type="hidden" name="id_<%=rs("ID")%>" value="<%=rs("ID")%>">
                <% Else %>
                <div class="team <%If rs("分1")>rs("分2") Then%>win<%End If%>"><%=H(rs("队1"))%></div>
                <div class="score"><%=H(rs("分1"))%></div>
                <div class="vs">VS</div>
                <div class="score"><%=H(rs("分2"))%></div>
                <div class="team <%If rs("分2")>rs("分1") Then%>win<%End If%>"><%=H(rs("队2"))%></div>
                <% End If %>
              </div>
              <div class="bo-tag"><%=H(rs("赛制"))%></div>
            </div>
            <%
        End If
        rs.MoveNext
    Loop
    If Not hasMatch Then
        Response.Write "<div class='empty-tip'>暂无对阵</div>"
    End If
End Sub
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>16进8瑞士轮对阵表｜张家港市高校CS联赛</title>
<style>
  :root{ --ink:#1b1b1b; --main:#12305e; --line:#c9d4e4; --soft:#f4f7fb; }
  *{box-sizing:border-box;margin:0;padding:0;}
  body{
    margin:0; padding:40px 16px; background:#ffffff; color:var(--ink);
    font-family:'PingFang SC','Microsoft YaHei','Noto Sans CJK SC',sans-serif;
    font-size:14px; line-height:1.7;
  }
  .paper{
    position:relative; max-width:1200px; margin:0 auto; background:#fff;
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
    color:var(--main); font-weight:700; margin-bottom:10px;text-align:center;
  }
  .subtitle{
    text-align:center; color:#8b98ad;margin-bottom:32px;font-size:14px;
  }

  /* 对阵图容器 */
  .match-container{
    display:flex; gap:20px; align-items:stretch;
    overflow-x:auto; padding-bottom:16px;
  }
  .round-column{
    flex:1; min-width:190px; display:flex; flex-direction:column; gap:16px;
  }
  .group-box{
    flex:1;
  }
  .group-title{
    text-align:center; font-weight:700; color:var(--main);
    margin-bottom:12px; padding:6px; background:var(--soft);
    border-radius:4px; font-size:14px;
  }
  .match-card{
    border:1px solid var(--line); border-radius:6px;
    margin-bottom:12px; background:#fff; overflow:hidden;
  }
  .match-teams{
    display:flex; align-items:center; justify-content:space-between;
    padding:10px 8px; gap:6px;
  }
  .team{
    flex:1; text-align:center; font-size:13px; min-width:0;
    overflow:hidden; text-overflow:ellipsis; white-space:nowrap;
  }
  .team.win{ font-weight:700; color:#23856d; }
  .team input{
    width:100%; padding:4px 6px; border:1px solid var(--line);
    border-radius:4px; text-align:center; font-size:12px;
  }
  .score{
    width:36px; text-align:center; font-weight:700; color:var(--main);
  }
  .score input{
    width:100%; padding:4px 0; border:1px solid var(--line);
    border-radius:4px; text-align:center; font-weight:700; font-size:13px;
  }
  .vs{
    font-size:12px; color:#8b98ad; padding:0 2px;
  }
  .bo-tag{
    text-align:center; font-size:12px; color:#6b788c;
    padding:4px; background:var(--soft); border-top:1px solid var(--line);
  }

  /* 晋级标注 */
  .promote-tag{
    text-align:center; font-size:13px; color:#23856d;
    font-weight:600; margin-top:8px; padding:4px;
    background:#f0f9f4; border-radius:4px;
  }

  .btn-row{
    margin-top:28px; text-align:center;
  }
  .nav-btn{
    display:inline-block; padding:11px 28px; margin:0 6px;
    background:var(--main); color:#fff; text-decoration:none;
    border-radius:6px; font-size:15px; letter-spacing:1px;
    border:none; cursor:pointer;
  }
  .nav-btn:hover{ background:#1b4b8f; }
  .nav-btn.secondary{
    background:var(--soft); color:var(--main); border:1px solid var(--line);
  }
  .empty-tip{
    text-align:center; padding:20px; color:#8b98ad; font-size:13px;
  }
  @media (max-width:768px){
    .paper{ padding:28px 20px; }
    .round-column{ min-width:170px; }
  }
</style>
</head>
<body>
<div class="paper">
  <h1>张家港市高校 CS 联赛</h1>
  <p class="subtitle">16进8 · 瑞士轮对阵表</p>

  <% If isAdmin Then %>
  <form method="post">
    <input type="hidden" name="action" value="save">
    <input type="hidden" name="matchcount" value="<%=rs.RecordCount%>">
  <% End If %>

  <div class="match-container">
    <!-- 第1轮：0-0 -->
    <div class="round-column">
      <div class="group-box">
        <div class="group-title">0-0</div>
        <% Call RenderGroup("0-0") %>
      </div>
    </div>

    <!-- 第2轮：1-0 / 0-1 -->
    <div class="round-column">
      <div class="group-box">
        <div class="group-title">1-0</div>
        <% Call RenderGroup("1-0") %>
      </div>
      <div class="group-box">
        <div class="group-title">0-1</div>
        <% Call RenderGroup("0-1") %>
      </div>
    </div>

    <!-- 第3轮：2-0 / 1-1 / 0-2 -->
    <div class="round-column">
      <div class="group-box">
        <div class="group-title">2-0</div>
        <% Call RenderGroup("2-0") %>
        <div class="promote-tag">第一、二名</div>
      </div>
      <div class="group-box">
        <div class="group-title">1-1</div>
        <% Call RenderGroup("1-1") %>
      </div>
      <div class="group-box">
        <div class="group-title">0-2</div>
        <% Call RenderGroup("0-2") %>
      </div>
    </div>

    <!-- 第4轮：2-1 / 1-2 -->
    <div class="round-column">
      <div class="group-box">
        <div class="group-title">2-1</div>
        <% Call RenderGroup("2-1") %>
        <div class="promote-tag">第三、四、五名</div>
      </div>
      <div class="group-box">
        <div class="group-title">1-2</div>
        <% Call RenderGroup("1-2") %>
      </div>
    </div>

    <!-- 第5轮：2-2 -->
    <div class="round-column">
      <div class="group-box">
        <div class="group-title">2-2</div>
        <% Call RenderGroup("2-2") %>
        <div class="promote-tag">第六、七、八名</div>
      </div>
    </div>
  </div>

  <div class="btn-row">
    <% If isAdmin Then %>
    <input type="submit" class="nav-btn" value="保存修改">
    <% End If %>
    <a href="first.html" class="nav-btn secondary">返回导航页</a>
  </div>

  <% If isAdmin Then %>
  </form>
  <% End If %>

</div>
<%
rs.Close
Set rs = Nothing
%>
</body>
</html>
