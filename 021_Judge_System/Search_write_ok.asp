<!-- #include file="inc.asp" -->
<%
if session("id")="" or session("Aceess_RR")<>"P" then  
'�α����Ͽ� ���� ����(id)�� ������ �α������� ���� ������ ������ ����Ʈ�� �����ش�.
%>

<html>
<head>
<body bgcolor="#D7F1FA">
<script language="javascript">
		alert("��ȸ ������ �ִ� ������� �α��� �ϼ���! \n\n\Ȥ�� �α��� �ƴ��� �����Ǿ� ����Ǿ����ϴ�.. \n\n\�� �α����� �ʿ��մϴ�.  login please !!!");
	window.open('../Log_in_B.asp','end','width=310,height=190,top=270, left=350');
</script>

<% else %>

<!--#include file="inc_Outsourcing_Good.asp"-->

<html>
<head>
<title>��ǰ �ڵ� �˻� ���</title>


<%
P_Code=request("P_Code")

set DB=server.createobject("adodb.connection")
DB.open connstring

sql="select * from AL_038_Outsourcing_Good where P_Code LIKE '%" & P_Code & "%'"
set RS=DB.execute(sql)

'�������� �°� �����Ѵ�

SQL = SQL & " ORDER BY P_Code DESC,sid DESC"

%>

<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link rel="stylesheet" href="basic.css" type="text/css">
</head>
<br>
<br style="line-height:10px;">
<body topmargin=0 marginheight=0 leftmargin="0"  bgcolor="#D7F1FA">

<table width=720 border=0 cellspacing=0 cellpadding=0  align="center">
<tr height="35"> 
<td valign=middle> <b><font color="#0066FF">�� <u>���� ��ǰ Db �˻� ���</u>  </font></b></td>
 <td height=25 valign="bottom" align=right >
           <a href="javascript:history.go(-1)"><img src="images/back.gif"  border="0" valign=bottom align="top"></a>&nbsp;
           <a href="list.asp"><img src=images/list.gif border=0></a></td>
      </tr>
</table>
<table width="720"  border="0" cellpadding="0" cellspacing="0" align=center bgcolor=#FFFFFF> 
       <tr> <td height="1" colspan="6" align="center" bgcolor=#BBBBBB></td>
              </tr>
              <tr height="24"> 
                <td width="40" align="center" valign="middle" bgcolor="#E5B2B2"><b>��ȣ</td>
                <td width="120" align="center" bgcolor="#E5B2B2"><b>��ǰ�ڵ�</td>
                <td width="360" align="center" bgcolor="#E5B2B2"><b>��ǰ��</td>
                <td width="70" align="center" bgcolor="#E5B2B2"><b>ERP No</td>
                <td width="60" align="center" bgcolor="#E5B2B2"><b>�����</td>
                <td width="70" align="center" bgcolor="#E5B2B2"><b>�����</td>
              </tr>
              <tr> 
                <td height="1" colspan="6" align="center" bgcolor=#BBBBBB></td>
              </tr>
<%
IF (RS.BOF and RS.EOF) Then
	Response.Write "<tr height=50> <td colspan=6 align=center>"
	Response.Write "���� ��ǰ Db �˻� ����� �����ϴ�."
	Response.Write "</td></tr>"
	
Else
 TotRecord = RS.RecordCount
	Rs.pagesize = 200 '���������� ������ ���ڵ尳��(inc.asp���� ����)
	TotPage   = RS.PageCount
	
  S_number=0 '���� �ʱ� ��
	
	RCount = RS.pageSize
	
	imsiNO=totrecord-(ipage-1)*(Rcount) '���ڵ��ȣ�� ����� �ӽù�ȣ
	
	Do while (NOT RS.EOF) and (RCount > 0 )

'�� ���������� ����� �ʵ��� ���� ���� �Ѳ����� �����ͼ� ������ ������ �д�.
'�̷��� �ѹ��� �������� �� ���� �����̴�.


	Sid=RS("sid")
	P_code=RS("P_code")
	Registor =RS("Registor")
	P_name=RS("P_name")
	ERP_No=RS("ERP_No")
	Visit=RS("visit")
  STime=RS("Stime")
	
	'�Ϸù�ȣ �ű��
  S_number = S_number+1
	%> 
	
	<tr onMouseOut="this.style.background='#FFFFFF'" onMouseOver="this.style.background='#ffdee9'" >
                <!--��� ��ȣ�� ����Ѵ�-->
                <td style="text-align:center; text-indent:0; margin:0; padding-top:5px; padding-right:0px; padding-bottom:3px; padding-left:5px;" valign="middle">
                  <%=S_number%></td>
                <!--������ ����Ѵ�-->
                <td  style="text-align:left; text-indent:0; margin:0; padding-top:5px; padding-right:0px; padding-bottom:3px; padding-left:5px;" valign="middle">
                  <a href="product_detail.asp?<%=Var2%>&sid=<%=Sid%>"><strong><%=P_code%></strong></a></td>
                <td  style="text-align:left; text-indent:0; margin:0; padding-top:5px; padding-right:0px; padding-bottom:3px; padding-left:5px;" valign="middle">
                   <a href="Write_Outsourcing_Good.asp?<%=Var2%>&sid=<%=Sid%>"><%=P_name%></a></td>
                   <td style="text-align:center; text-indent:0; margin:0; padding-top:5px; padding-right:0px; padding-bottom:3px; padding-left:5px;" valign="middle">
              <%=ERP_No%></td>
                  <td style="text-align:center; text-indent:0; margin:0; padding-top:5px; padding-right:0px; padding-bottom:3px; padding-left:5px;" valign="middle">
              <%=Registor%></td>
                 <td style="text-align:center; text-indent:0; margin:0; padding-top:5px; padding-right:0px; padding-bottom:3px; padding-left:5px;" valign="middle">
              <%=left(STime,10)%></td>
              </tr>
<%
	RS.MoveNext
	RCount = RCount -1
	Loop

End if

RS.Close
Set RS=nothing

%>
             <tr> 
                <td height="1" colspan="6" align="center" bgcolor=#BBBBBB></td>
              </tr>
</table>





<br>
<br style="line-height:30px;">


<!--#include file="inc_Packing_Info.asp"-->

<%
P_Code=request("P_Code")

set DB=server.createobject("adodb.connection")
DB.open connstring

sql="select * from AL_034_Packing_Info where P_Code LIKE '%" & P_Code & "%'"
set RS=DB.execute(sql)

'�������� �°� �����Ѵ�

SQL = SQL & " ORDER BY P_Code DESC,sid DESC"

%>



<table width=720 border=0 cellspacing=0 cellpadding=0  align="center">
<tr height="35"> 
<td valign=middle> <b><font color="#0066FF">�� <u>���� ���� ��ǰ Db �˻� ���</u>  </font></b></td>
 <td height=25 valign="bottom" align=right >
           <a href="javascript:history.go(-1)"><img src="images/back.gif"  border="0" valign=bottom align="top"></a>&nbsp;
           <a href="list.asp"><img src=images/list.gif border=0></a></td>
      </tr>
</table>
<table width="720"  border="0" cellpadding="0" cellspacing="0" align=center bgcolor=#FFFFFF> 
       <tr> <td height="1" colspan="6" align="center" bgcolor=#BBBBBB></td>
              </tr>
              <tr height="28"> 
                <td width="40" align="center" valign="middle" bgcolor="#008080"><b>��ȣ</td>
                <td width="120" align="center" bgcolor="#008080"><b>��ǰ�ڵ�</td>
                <td width="360" align="center" bgcolor="#008080"><b>��ǰ��</td>
                <td width="70" align="center" bgcolor="#008080"><b>ERP No</td>
                <td width="60" align="center" bgcolor="#008080"><b>�����</td>
                <td width="70" align="center" bgcolor="#008080"><b>�����</td>
              </tr>
              <tr> 
                <td height="1" colspan="6" align="center" bgcolor=#BBBBBB></td>
              </tr>
<%
IF (RS.BOF and RS.EOF) Then
	Response.Write "<tr height=50> <td colspan=6 align=center>"
	Response.Write "���� ���� ��ǰ Db �˻� ����� �����ϴ�."
	Response.Write "</td></tr>"
	
Else
 TotRecord = RS.RecordCount
	Rs.pagesize = 200 '���������� ������ ���ڵ尳��(inc.asp���� ����)
	TotPage   = RS.PageCount
	
  S_number=0 '���� �ʱ� ��
	
	RCount = RS.pageSize
	
	imsiNO=totrecord-(ipage-1)*(Rcount) '���ڵ��ȣ�� ����� �ӽù�ȣ
	
	Do while (NOT RS.EOF) and (RCount > 0 )

'�� ���������� ����� �ʵ��� ���� ���� �Ѳ����� �����ͼ� ������ ������ �д�.
'�̷��� �ѹ��� �������� �� ���� �����̴�.


	Sid=RS("sid")
	P_code=RS("P_code")
	Registor =RS("Registor")
	P_name=RS("P_name")
	ERP_No=RS("ERP_No")
	Visit=RS("visit")
  STime=RS("Stime")
	
	'�Ϸù�ȣ �ű��
  S_number = S_number+1
	%> 
	
	<tr onMouseOut="this.style.background='#FFFFFF'" onMouseOver="this.style.background='#ffdee9'" height="28" >
                <!--��� ��ȣ�� ����Ѵ�-->
                  <td style="text-align:center; text-indent:0; margin:0; padding-top:5px; padding-right:0px; padding-bottom:3px; padding-left:5px;" valign="middle">
              <%=S_number%></td>
                <!--������ ����Ѵ�-->
               <td  style="text-align:left; text-indent:0; margin:0; padding-top:5px; padding-right:0px; padding-bottom:3px; padding-left:5px;" valign="middle">
                  <a href="product_detail.asp?<%=Var2%>&sid=<%=Sid%>"><strong><%=P_code%></strong></a></td>
               <td  style="text-align:left; text-indent:0; margin:0; padding-top:5px; padding-right:0px; padding-bottom:3px; padding-left:5px;" valign="middle">
                   <a href="Write_Packing_Good.asp?<%=Var2%>&sid=<%=Sid%>"><%=P_name%></a></td>
                   <td style="text-align:center; text-indent:0; margin:0; padding-top:5px; padding-right:0px; padding-bottom:3px; padding-left:5px;" valign="middle">
              <%=ERP_No%></td>
                  <td style="text-align:center; text-indent:0; margin:0; padding-top:5px; padding-right:0px; padding-bottom:3px; padding-left:5px;" valign="middle">
              <%=Registor%></td>
                 <td style="text-align:center; text-indent:0; margin:0; padding-top:5px; padding-right:0px; padding-bottom:3px; padding-left:5px;" valign="middle">
              <%=left(STime,10)%></td>
              </tr>
<%
	RS.MoveNext
	RCount = RCount -1
	Loop

End if

RS.Close
Set RS=nothing

%>
             <tr> 
                <td height="1" colspan="6" align="center" bgcolor=#BBBBBB></td>
              </tr>
</table>

<br>
<br style="line-height:50px;">
</body>
</html>
       <% end if %>   
