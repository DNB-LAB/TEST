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
<body >
<body leftmargin="0" topmargin="0" bgcolor="#D7F1FA">
<script language="javascript">
		alert("로그인이 필요합니다. 로그인 하세요! \n\n\혹은 로그인 됐더라도 오래되어 종료되었습니다.. \n\n\재 로그인이 필요합니다.  login please !!!");
	window.open('../../../Log_in.asp','end','width=310,height=190,top=270, left=350');
</script>

<% else %>

<!-- #include file="inc_Warwhouse.asp" -->
<html>
<head>
<title>물류 샘플링 검사 결과 삭제하기</title>
<meta http-equiv="content-type" content="text/html; charset=utf-8">
<link rel="stylesheet" href="basic.css" type="text/css">
</head>

<%

'삭제할 글번호
Sid = Request.QueryString("Sid")

' 파일만 삭제인지를 전송받음
file_mode = Request.QueryString("file_mode")

Var3 = Var2 & "&sid=" & sid



'내용을 볼 레코드의 각 필드값을 가져온다

Set DB = Server.Createobject("ADODB.connection")
DB.open Connstring


 '내용을 볼 레코드의 각 필드값을 가져온다
SQL = "select * from  AL_022_Judge_Warehouse"
SQL = SQL & " WHERE Original_sid =" & Sid
Set RS = DB.Execute(SQL)


'한꺼번에 모두 가져와서 변수에 대입

Sfile1 = rs("Sfile1")
Sfile2 = rs("Sfile2")
Sfile3 = rs("Sfile3")

W_Registor = rs("W_Registor")
Pass= rs("Pass")
%> 
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
<body bgcolor="#D7F1FA">
<center>
<form method="post" name="form" action="Warehouse_Delete_ok.asp?<%=Var3%>&file_mode=<%=file_mode%>" onSubmit="return Send()">
 
   <table Cellspacing=0 cellpadding=2 width="880" border="0" align=center  style="table-layout:fixed;">   
    <tr height=35> 
      <td width="500" align="left" bgcolor=#D7F1FA >
     <b>▶&nbsp;첨부 파일 포함
      &nbsp;&nbsp;<% if file_mode = "file1" then%>
     <font color=red> 첨부파일 1  </font>
       <% elseif file_mode = "file2" then%>
     <font color=red> 첨부파일 2 </font>
        <% elseif file_mode = "file3" then%>
     <font color=red> 첨부파일 3  </font>
     
        
     
        <% else%>
      <font color=red>[ 전 체 ]</font>
     <% end if %>
     
     삭제</b></td>
 </b></td>
  <td align="right" bgcolor=#D7F1FA ><a href="javascript:history.go(-1)"><img src="../images/back.gif"  border="0" valign=bottom align="top"></a>
        &nbsp;
      <a href="../list.asp?<%=Var3%>"><img src=../images/list.gif border=0></a></td>
    </tr>
</table>
 
<table width="880" border=1 cellpadding=0 cellspacing="0" align=center  style="table-layout:fixed;">

   <tr height=60> 
        <td  width=120 bgcolor="#F0B6B6" style="text-align:center; text-indent:0; margin:0; padding-top:5px; padding-right:0px; padding-bottom:5px; padding-left:0px;">
      <b>암&nbsp;&nbsp;&nbsp;호</td>
        
        <td width="760" bgcolor="#FFFFFF">
         <table border=0 width="720" cellpadding=0 cellspacing="0"  style="table-layout:fixed;">
          <tr>
      <td  width="600"  bgcolor="#FFFFFF" style="text-align:left; text-indent:0; margin:0; padding-top:3px; padding-right:0px; padding-bottom:3px; padding-left:20px;">            
          <input type="password" name="Pass" size="10" maxlength="10">  ※ 입력시 등록한 암호&nbsp;&nbsp;&nbsp;&nbsp;
           <% if session("mname") = W_Registor  then %>
              <font color=red><b>암호 : <%=Pass%></b></font>
               (작성자 본인 로그인시만 보임)
              <% else %>
              <font color=red><b>암호 : *****</font></b>
              (작성자 본인 로그인시만 보임)
              <% end if%></td>
         <td width="120" style="text-align:left; text-indent:0; margin:0; padding-top:8px; padding-right:0px; padding-bottom:3px; padding-left:0px;">            
         <input type="image" img src="../images/del.gif"  border="0">
        </td>
        </tr>
      </table>
  </form>

</td>
</tr>
</table>
<br><br><br><br><br>
</body>
</html>

 <% end if %>
 
 