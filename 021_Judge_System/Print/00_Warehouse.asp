<!DOCTYPE html>
<% @LANGUAGE='VBSCRIPT' CODEPAGE='949' %> 
<% session.codepage = 949 %>
<% Response.ChaRset = "euc-kr" %>

<!-- #include file="inc.asp" -->

<%
to_year = year(date())

start_year = Cint(to_year)-20
end_year = Cint(to_year)+1

sdate=date()

%>

<html>
<head>
<title>�������� ���� ���� ���</title>
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
        <form method=get action="00_Warehouse_Ok.asp?<%=Var5%>"  name="form" onSubmit="Check();return true;">
        <tr height="30">
            <td width="600"  bgcolor="#FF9999" align="center">
          <font face="����"><b>�������� ���� ���� ��� ����</td>
            <td width="200" bgcolor="#FF9999"  align="center">
                <font face="����"><b>�� ȸ </td>
        </tr>
        <tr height="50">
            <td  align=center  bgcolor="#FFFFFF">
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
	&nbsp;&nbsp;&nbsp;&nbsp;�������: &nbsp;
	<select name="Judge_Result">
                                     <option value="">��ü</option>
                                     <option value="����">����</option>
                                     <option value="������">������</option>
                                     <option value="����">����</option>
                                     <option value="���԰�">���԰�</option>
                                     </select></td>
            <td  align=center  bgcolor="#FFFFFF">
                  <span style="font-size:10pt;" ><input type="submit" value="�� ȸ"></span></td>
        </tr>
</table>
 
</body>
</form>
</html>
