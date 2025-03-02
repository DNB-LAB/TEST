<%@Language="VBScript" CODEPAGE="65001" %>
<% 
Response.CharSet="utf-8" 
Session.codepage="65001" 
Response.codepage="65001" 
Response.ContentType="text/html;charset=utf-8" 
%>

<!DOCTYPE HTML>

<!-- #include file="inc.asp" -->
<%
if session("id")="" then  
'로그인하여 얻은 세션(id)가 없으면 로그인으로 돌려 보내고 있으면 리스트를 보여준다.
%>

<html>
<head>
<body leftmargin="0" topmargin="0" bgcolor="#D7F1FA">
<script language="javascript">
		alert("로그인이 필요합니다. 로그인 하세요! \n\n\혹은 로그인 됐더라도 오래되어 종료되었습니다.. \n\n\재 로그인이 필요합니다.  login please !!!");
	window.open('../../Log_in.asp','end','width=310,height=190,top=270, left=350');
</script>

<% else %>

<%
'삭제할 글번호
sid = Request.QueryString("sid")

' 파일만 삭제인지를 전송받음
file_mode = Request.QueryString("file_mode")

Var3 = Var2 & "&sid=" & sid
%>

<html>
<head>
<title시장출하 적부 판정 기록 삭제하기</title>
<link rel="stylesheet" href="basic.css">
<meta http-equiv="content-type" content="text/html; charset=utf-8">
</head>
<body bgcolor="#D7F1FA">
<center>

<script language="Javascript">
<!--
function Send() 
{
	var vP = document.form.Pass.value;
	if (vP == "") {
		alert("암호를 입력하세요.\n");
		document.form.Pass.focus();
		return false;
}
return true;
} // end function
//  -->
</script>

<br>
  <table cellspacing=0 cellpadding=2 border=0 width=720>
    <tr height="30"> 
      <td ><b>▶시장출하 적부 판정 기록 삭제하기</b></td>
    </tr>
  </table>
  <table Cellspacing=0 cellpadding=2 width="720" border="0">
    <tr> 
     
       <tr> 
      <td   bgcolor="#D7F1FA"><FONT COLOR=RED><B>&nbsp;└ 시장출하 적부 판정 기록 삭제시 등록된 물류 샘플링 검사 결과도 같이 삭제되니 유의하세요. </td>
       </tr>
       <tr> 
       <td   bgcolor="#D7F1FA"> <% if file_mode = "file1" then%>
     <font color=red>[ 첨부파일 1  삭제하기 ]</font>
     <% elseif file_mode = "file2" then%>
     <font color=red>[ 첨부파일 2  삭제하기 ]</font>
      <% elseif file_mode = "file3" then%>
     <font color=red>[ 첨부파일 3  삭제하기 ]</font>
      <% else%>
      <font color=red>[ 전체 삭제하기 ]</font>
     <% end if %>
     </td>
        
       <td align="right"><a href="javascript:history.go(-1)"><img src="images/back.gif"  border="0" valign=bottom align="top"></a>
      &nbsp;<a href="list.asp"><img src="images/list.gif"  border="0" ></a></td>
    </tr>
  </table>
  <table width="720" border=1 cellpadding=2 cellspacing="0" bgcolor="#FFFFFF">
    <form method="post" name="form" action="delete_ok.asp?<%=Var3%>&mode=<%=mode%>" onSubmit="return Send()">
      <tr height=80> 
        <td width="100" align="center" bgcolor="#CCCC"><font color="black"><b>암&nbsp;&nbsp;&nbsp;호</b></font></td>
        <td bgcolor="#FFFFFF">
          &nbsp;<input type="password" NAME="Pass" SIZE="15" MAXLANGTH=15>
          * 처음에 입력한 암호가 필요합니다. &nbsp;&nbsp;
          <input type="submit" value="삭  제" name="submit"></td>
       </tr>
       </td>
        </tr>
      </table>
    </form>
</td>
</tr>
</table>
<br><br><br><br><br>
  
  
  
  
</center>
</body>
</html>
 <% end if %>