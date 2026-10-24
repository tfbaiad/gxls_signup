<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<% Session.CodePage = 936 %>
<!--#include file="Connections/sjk.asp" -->
<%
' ========== 路径与状态 ==========
MM_editAction = CStr(Request.ServerVariables("SCRIPT_NAME"))
If Request.QueryString <> "" Then
  MM_editAction = MM_editAction & "?" & Server.HTMLEncode(Request.QueryString)
End If

MM_abortEdit = False
errMsg = ""
successMsg = ""

' ========== 接收表单数据 ==========
playerSchool = Trim(Request.Form("playerSchool"))
perfectName  = Trim(Request.Form("perfectName"))
teamName     = Trim(Request.Form("teamName"))
playerQQ     = Trim(Request.Form("playerQQ"))
confirmText  = Trim(Request.Form("confirmText"))
fillDate     = Trim(Request.Form("fillDate"))

' ========== 单引号转义 ==========
Function SQLStr(s)
    SQLStr = Replace(s, "'", "''")
End Function

' ========== HTML转义 ==========
Function H(s)
    If IsNull(s) Or s = "" Then
        H = ""
    Else
        H = Server.HTMLEncode(CStr(s))
    End If
End Function

' ========== 表单提交处理 ==========
If CStr(Request("MM_insert")) = "form1" Then
    ' 服务端非空校验
    If perfectName = "" Then
        errMsg = errMsg & "请填写选手完美名称<br>"
        MM_abortEdit = True
    End If
    If confirmText <> "是" Then
        errMsg = errMsg & "请手动输入确认文字：是<br>"
        MM_abortEdit = True
    End If
    If fillDate = "" Then
        errMsg = errMsg & "请填写填写日期<br>"
        MM_abortEdit = True
    End If

    ' 校验通过，按完美名称匹配并仅更新指定字段
    If Not MM_abortEdit Then
        On Error Resume Next
        Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT * FROM [汇总] WHERE [选手完美名称]='" & SQLStr(perfectName) & "'"
        rs.Open sql, MM_sjk_STRING, 1, 3
        
        If rs.EOF And rs.BOF Then
            errMsg = errMsg & "未找到该完美名称的报名记录，请先完成选手信息报名<br>"
            MM_abortEdit = True
            rs.Close
            Set rs = Nothing
        Else
            ' 仅更新两个字段：是否填写承诺书、日期
            rs("是否填写承诺书") = confirmText
            rs("日期") = fillDate
            rs.Update
            rs.Close
            Set rs = Nothing
            
            If Err.Number = 0 Then
                MM_editRedirectUrl = "tips.html"
                If Request.QueryString <> "" Then
                    If InStr(1, MM_editRedirectUrl, "?", vbTextCompare) = 0 Then
                        MM_editRedirectUrl = MM_editRedirectUrl & "?" & Request.QueryString
                    Else
                        MM_editRedirectUrl = MM_editRedirectUrl & "&" & Request.QueryString
                    End If
                End If
                Response.Redirect(MM_editRedirectUrl)
            Else
                errMsg = "提交失败，数据库错误：" & Err.Description
            End If
        End If
        On Error GoTo 0
    End If
End If

' ========== 院校列表【已新增2所】 ==========
schoolList = Array( _
    "梁丰高级中学及其国际部", "沙州中学本校区", "沙州中学北校区", _
    "暨阳高级中学", "张家港高级中学", "塘桥高级中学", _
    "崇真中学", "凤凰高级中学", "乐余高级中学", _
    "常青藤实验学校及其国际部", "外国语实验学校及其国际部", _
    "张家港第一职业学校", "张家港第二职业学校", "张家港第三职业学校", _
    "张家港开放大学", "江苏科技大学张家港校区", "沙洲工学院", _
    "张家港市艺术学校（原张家港舞蹈学校）", "张家港市特殊教育学校" _
)
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>张家港市高校CS联赛公平比赛承诺书</title>
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
  .lead{margin:0 0 20px; text-indent:2em; text-align:justify;}
  ol.clauses{padding-left:2.2em; margin:0;}
  ol.clauses li{margin-bottom:11px; padding-left:.35em; text-align:justify; line-height:1.9;}
  ol.clauses li::marker{color:var(--main); font-weight:700;}
  ol.clauses li b{font-weight:700;}
  .school-list{
    margin:24px 0 8px; padding:18px 22px;
    background:var(--soft); border:1px solid var(--line); border-radius:6px;
    font-family:'PingFang SC','Microsoft YaHei',sans-serif; font-size:14px;
  }
  .school-list p{margin:0 0 10px;}
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
  .btn-row{margin-top:22px;}
  .btn-row input[type=submit]{
    font-family:'PingFang SC','Microsoft YaHei',sans-serif;
    font-size:15px; color:#fff; background:#12305e; border:0;
    padding:11px 34px; border-radius:6px; cursor:pointer; letter-spacing:2px;
  }
  .btn-row input[type=submit]:hover{background:#1b4b8f;}
  .msg-box{
    padding:12px 16px; border-radius:6px; margin-bottom:20px;
    font-family:'PingFang SC','Microsoft YaHei',sans-serif; font-size:14px;
  }
  .msg-error{ background:#fff2f2; border:1px solid #ffccc7; color:#cf1322; }
  .msg-success{ background:#f6ffed; border:1px solid #b7eb8f; color:#389e0d; }
  .note{
    margin:26px 0 0; padding-top:16px; border-top:1px solid var(--line);
    font-size:13px; color:#6b788c; line-height:1.85;
    font-family:'PingFang SC','Microsoft YaHei',sans-serif;
  }
  @media (max-width:640px){
    body{font-size:14.5px; padding:16px 10px 40px;}
    .paper{padding:32px 22px 30px;}
    h1{font-size:20px; letter-spacing:1px;}
    .school-list-ul{columns:1;}
    table.form-table{font-size:13.5px;}
    table.form-table td{padding:7px 8px;}
    table.form-table td.lbl{width:auto;}
  }
</style>
<script type="text/vbscript">
' 客户端提交前校验
Function form1_onsubmit()
    If Trim(document.form1.perfectName.value) = "" Then
        MsgBox "请填写选手完美名称", 48, "提示"
        document.form1.perfectName.focus()
        form1_onsubmit = False
        Exit Function
    End If
    If Trim(document.form1.confirmText.value) <> "是" Then
        MsgBox "请手动输入确认文字：是",48,"提示"
        document.form1.confirmText.focus()
        form1_onsubmit = False
        Exit Function
    End If
    If Trim(document.form1.fillDate.value) = "" Then
        MsgBox "请填写填写日期", 48, "提示"
        document.form1.fillDate.focus()
        form1_onsubmit = False
        Exit Function
    End If
    form1_onsubmit = True
End Function

Function Trim(str)
    Trim = Replace(Replace(str, " ", ""), Chr(9), "")
End Function
</script>
</head>
<body>
<div class="paper">
  <div class="doc-no">承诺书编号：ZJG-CS-2027</div>
  <h1>张家港市高校 CS 联赛<br>公平比赛承诺书</h1>
  <p class="subtitle">公平 &middot; 公正 &middot; 公开</p>
  <!-- 错误/成功提示 -->
  <% If errMsg <> "" Then %>
  <div class="msg-box msg-error"><%=errMsg%></div>
  <% End If %>
  <% If successMsg <> "" Then %>
  <div class="msg-box msg-success"><%=successMsg%></div>
  <% End If %>
  <p class="lead">
    为维护张家港市高校 CS 联赛（以下简称&ldquo;本次联赛&rdquo;）公平、公正、公开的竞赛环境，
    营造风清气正的高校电竞氛围，本人作为本次联赛的参赛人员，郑重作出如下承诺：
  </p>
  <ol class="clauses">
    <li><b>遵守规程，服从管理。</b>本人已认真阅读并完全理解《张家港市高校 CS 联赛竞赛规程》及组委会发布的各项通知与补充规定，自愿报名参赛并严格遵守，服从组委会、裁判组的统一管理与赛程安排。</li>
    <li><b>按时参赛，恪守纪律。</b>按赛程规定时间到达比赛场地（线上赛按通知提前进入指定服务器或房间），不无故迟到、缺席、中途退赛；确需变更或请假的，按规程提前向组委会报备。</li>
    <li><b>杜绝作弊，公平竞技。</b>保证不使用任何外挂、自瞄、透视、脚本、宏命令、修改器、注入程序及其他第三方作弊软硬件；不利用游戏漏洞（Bug）或其他非正常手段获取比赛优势，发现漏洞第一时间向裁判报告。</li>
    <li><b>真实身份，严禁代打。</b>保证报名信息及参赛身份真实有效，与本人所在高校学籍一致；不请他人代打、替赛，不冒用他人账号，不出租、出借、转让参赛账号与参赛资格。</li>
    <li><b>规范设备，不擅改动。</b>使用组委会允许并审核报备的设备与外设，不擅自安装软件、修改系统或游戏设置，不篡改比赛数据、比赛录像与比赛结果。</li>
    <li><b>尊重对手，文明参赛。</b>尊重对手、裁判、解说与观众，不使用侮辱、歧视、挑衅性语言与行为，不在游戏内公屏、语音频道及社交平台上进行人身攻击、地域攻击或恶意炒作。</li>
    <li><b>尊重判罚，理性申诉。</b>服从裁判现场判罚；对判罚有异议的，由本队领队按规程在赛后规定时限内以书面形式提出申诉，不擅自中断比赛、不聚众滋事、不通过网络舆论向组委会施压。</li>
    <li><b>廉洁自律，拒绝交易。</b>不参与任何形式的假赛、消极比赛、串通比赛结果、赌赛博彩及利益交换；不接受任何可能影响比赛公正性的财物、宴请或其他好处。</li>
    <li><b>保守秘密，维护赛事。</b>不泄露、不传播涉及赛事未公开的战术资料、赛程安排、裁判信息等内部信息；自觉维护赛事与所在学校的声誉，不损坏比赛场地、设备及公共财物。</li>
    <li><b>如实申报，自担风险。</b>如实填写报名信息与健康信息，确认本人身体状况适合参加本次比赛；比赛期间因个人原因造成的人身伤害及财产损失，由本人依法承担。</li>
    <li><b>接受监督，配合调查。</b>自觉接受组委会、裁判组及参赛各方的监督，积极配合赛风赛纪调查，如实提供相关证据材料，不隐瞒、不伪造、不串供。</li>
    <li><b>违约担责，后果自负。</b>如违反上述任何一项承诺，本人自愿接受组委会依据竞赛规程作出的处罚，包括但不限于：口头警告、判负、扣分、取消当场或全部比赛成绩、取消参赛资格、禁赛、通报所在学校等；造成其他损失的，依法予以赔偿；涉嫌违法犯罪的，移交有关部门处理。</li>
  </ol>
  <form name="form1" action="<%=MM_editAction%>" method="POST" onsubmit="return form1_onsubmit()">
    <div class="school-list">
      <p><b>选手学校（单选）：</b></p>
      <ul class="school-list-ul">
      <%
      For i = 0 To UBound(schoolList)
          s = schoolList(i)
          If playerSchool = s Then
              checkedAttr = " checked=""checked"""
          Else
              checkedAttr = ""
          End If
      %>
        <li>
          <input type="radio" name="playerSchool" id="school<%=i%>" value="<%=H(s)%>"<%=checkedAttr%> />
          <label for="school<%=i%>"><%=H(s)%></label>
        </li>
      <%
      Next
      %>
      </ul>
    </div>
    <table class="form-table">
      <caption>信息填写区域</caption>
      <tr>
        <td class="lbl">选手本人完美名称</td>
        <td><input type="text" name="perfectName" id="perfectName" value="<%=H(perfectName)%>" /></td>
      </tr>
      <tr>
        <td class="lbl">选手参赛队伍</td>
        <td><input type="text" name="teamName" id="teamName" value="<%=H(teamName)%>" /></td>
      </tr>
      <tr>
        <td class="lbl">选手QQ</td>
        <td><input type="text" name="playerQQ" id="playerQQ" value="<%=H(playerQQ)%>" /></td>
      </tr>
      <tr>
        <td class="lbl">阅读确认<br><span style="font-weight:400;color:#7b8798;">手动输入：是</span></td>
        <td><input type="text" name="confirmText" id="confirmText" value="<%=H(confirmText)%>" /></td>
      </tr>
      <tr>
        <td class="lbl">填写时日期<br /><span style="font-weight:400;color:#7b8798;">（例：2026/09/01）</span></td>
        <td><input type="text" name="fillDate" id="fillDate" value="<%=H(fillDate)%>" /></td>
      </tr>
    </table>
    <div class="btn-row">
      <input type="submit" value="提 交 承 诺 书" />
    </div>
    <input type="hidden" name="MM_insert" value="form1">
  </form>
  <p class="note">
    注：本承诺书一式两份，承诺人与联赛组委会各留存一份。
    本承诺书自签署之日起生效，至本次联赛全部赛程结束时自动终止。
  </p>
</div>
</body>
</html>