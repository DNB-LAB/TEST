<!-- #include file="inc.asp" -->
<%
if session("id")="" or session("Aceess_RR")<>"P" then  
'로그인하여 얻은 세션(id)가 없으면 로그인으로 돌려 보내고 있으면 리스트를 보여준다.
%>
<html>
<head>
<body  bgcolor="#D7F1FA">
<script language="javascript">
		alert("조회 권한이 있는 사번으로 로그인 하세요! \n\n\혹은 로그인 됐더라도 오래되어 종료되었습니다.. \n\n\재 로그인이 필요합니다.  login please !!!");
	window.open('../../Log_in_B.asp','end','width=310,height=190,top=270, left=350');
</script>
<% else %>

<%
to_year = year(date())

start_year = Cint(to_year)-20
end_year = Cint(to_year)+1

Fungus_date=date()

%>

<html>
<head>
<title>수탁제품, 외주가공 진균  시험결과 조회 및 인쇄</title>
<link rel="stylesheet" href="basic.css">
<meta name="author" content="이청희">
<body bgcolor="#D7F1FA">
<script language="javascript">
function Check() {
form.target = 'newW'; //타겟지정, 아래에는 url을 적지않는다
window.open ;
}
</script>

<style type="text/css">
<!--
.style4 {font-size: 12px}
-->
</style>
</head>

    <table border="1" width="800" align="center" cellspacing="0">
        <form method=get action="05_Fungus_result_ok.asp?<%=Var5%>"  name="form" onSubmit="Check();return true;">
        <tr height="30"> 
            <td width="600"  bgcolor="#EEEEEE" align="center">
          <font face="돋움"><b> 진균 시험 판정 결과 등록일</td>
            <td width="200" height="25" bgcolor="#EEEEEE" align="center">
               <font face="돋움"><b>조 회 </td>
        </tr>
        <tr  height="50">
            <td  align=center>
&nbsp;
<select name="Fungus_year" id="Fungus_year" style="width:80">
                <% for k=start_year to end_year %>
                <option value="<%=k%>"  <%if k=year(Fungus_date) then%> selected <% end if %>><%=k%>년</option>
								<% next%>
      </select>
              <select name="Fungus_month" id="Fungus_month"  style="width:60">
							  <% for k=1 to 12 %>
                <option value="<%=k%>"  <%if k=month(Fungus_date) then%> selected <% end if %>><%=k%>월</option>
								<% next%>
              </select>
              <select name="Fungus_day" id="Fungus_day"  style="width:60">
							  <% for k=1 to 31 %>
                <option value="<%=k%>"  <%if k=day(Fungus_date) then%> selected <% end if %>><%=k%>일</option>
	<% next%> </select> 
	<select name="Microbial_judge">
                                    <option value="2">진균</option></select></td>
            <td align=center>
                <span style="font-size:10pt;" ><input type="submit" value="조회"></span></td>
        </tr>
</table>
</body>
 <% end if %>
</form>
</html>
