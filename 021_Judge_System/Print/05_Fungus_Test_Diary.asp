<!-- #include file="inc.asp" -->
<%
if session("id")="" or session("Aceess_RR")<>"P" then  
'�α����Ͽ� ���� ����(id)�� ������ �α������� ���� ������ ������ ����Ʈ�� �����ش�.
%>
<html>
<head>
<body  bgcolor="#D7F1FA">
<script language="javascript">
		alert("��ȸ ������ �ִ� ������� �α��� �ϼ���! \n\n\Ȥ�� �α��� �ƴ��� �����Ǿ� ����Ǿ����ϴ�.. \n\n\�� �α����� �ʿ��մϴ�.  login please !!!");
	window.open('../../Log_in_B.asp','end','width=310,height=190,top=270, left=350');
</script>
<% else %>

<%
to_year = year(date())

start_year = Cint(to_year)-20
end_year = Cint(to_year)+1

sdate=date()

%>

<html>
<head>
<title>��Ź��ǰ, ���ְ�����ǰ/��Ÿ �̻��� ���� ����  ��ȸ �� ����ϱ�</title>
<link rel="stylesheet" href="basic.css">
<meta name="author" content="��û��">

<script language="javascript">
function Check() {
form.target = 'newW'; //Ÿ������, �Ʒ����� url�� �����ʴ´�
window.open ;
}
</script>

</head>
<body bgcolor="#D7F1FA">
    <table border="1" width="800" align="center" cellspacing="0">
        <form method=get action="05_Fungus_Test_Diary_ok.asp?<%=Var5%>"  name="form" onSubmit="Check();return true;">
        <tr height="30">
            <td width="600"  bgcolor="#E8A15E" align="center">
          <font face="����"><b>��,��Ź��ǰ �� ���ְ��� ��ǰ/��Ÿ ���� ���� ���� ��ȸ �� ���</td>
            <td width="200" bgcolor="#E8A15E"  align="center">
                <font face="����"><b>�� ȸ</td>
        </tr>
        <tr height="50">
            <td  align=center>
&nbsp;
<select name="Syear" id="Syear" style="width:80">
                <% for k=start_year to end_year %>
                <option value="<%=k%>"  <%if k=year(sdate) then%> selected <% end if %>><%=k%>��</option>
								<% next%>
      </select>
              <select name="Smonth" id="Smonth"  style="width:60">
							  <% for k=1 to 12 %>
                <option value="<%=k%>"  <%if k=month(sdate) then%> selected <% end if %>><%=k%>��</option>
								<% next%>
              </select>
              <select name="Sday" id="Sday"  style="width:60">
							  <% for k=1 to 31 %>
                <option value="<%=k%>"  <%if k=day(sdate) then%> selected <% end if %>><%=k%>��</option>
	<% next%> </select>
	
                              <select name="Microbial_judge">
                                    <option value="2">����</option></select></td>
                                     
                                     
            <td  align=center>
                <span style="font-size:10pt;" ><input type="submit" value="��ȸ">
                </span> </td>
        </tr>
</table>
</body>
 <% end if %>
</form>
</html>
