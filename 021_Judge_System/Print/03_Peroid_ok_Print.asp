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

Set RS = SERVER.CreateObject("ADODB.Recordset")
RS.CursorType = 3

Syear = request("Syear")
Smonth = request("Smonth")
Sday = request("Sday")
Microbial_judge = request("Microbial_judge")

Fyear = request("Fyear")
Fmonth = request("Fmonth")
Fday = request("Fday")

Sdate = syear & "-" & smonth & "-" & sday
Sdate = CDATE(Sdate) 

Fdate = Fyear & "-" & Fmonth & "-" & Fday
Fdate = CDATE(Fdate) 


if Microbial_judge = "1" then

SQL = "SELECT * FROM AL_046_OEM_Product_2018"
SQL = SQL & " WHERE Stime_request  between '"& Sdate & "' and '" & Fdate & "'"  
SQL = SQL & " and lev =  " &  0     '�Ƿڼ��� ��µǰ�

elseif Microbial_judge = "2" then

SQL = "SELECT * FROM AL_046_OEM_Product_2018"
SQL = SQL & " WHERE Stime_request  between '"& Sdate & "' and '" & Fdate & "'"  
SQL = SQL & " and Microbial_judge =  " &  2  '�̻�����
SQL = SQL & " and lev =  " &  0     '�Ƿڼ��� ��µǰ�

elseif Microbial_judge = "3" then

SQL = "SELECT * FROM AL_046_OEM_Product_2018"
SQL = SQL & " WHERE Stime_request  between '"& Sdate & "' and '" & Fdate & "'"  
SQL = SQL & " and Microbial_judge =  " &  1  ' �̻��� ����
SQL = SQL & " and lev =  " &  0  

end if 

'�������� �°� �����Ѵ�
SQL = SQL & " ORDER BY Sid ASC"

RS.Open SQL, ConnString

IF (RS.BOF and RS.EOF) Then
	TotRecord = 0 
	TotPage   = 0
Else
	TotRecord = RS.RecordCount
	Rs.pagesize=10000 '���������� 10000���� �����ش�
	TotPage   = RS.PageCount
End if
%>

<html>
<head>
<title>�Ⱓ�� ��Ź ODM  ��ǰ �����Ƿ� ��� �μ��ϱ�</title>
<link rel="stylesheet" href="basic.css">
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr"><script language="javascript">
function printWindow() {
factory.printing.header = ""
factory.printing.footer = ""   
factory.printing.portrait = true    
factory.printing.leftMargin = 0.0   
factory.printing.topMargin = 15.0    
factory.printing.rightMargin = 0.0
factory.printing.bottomMargin = 0.0
factory.printing.Print(true, window)
}
</script>


</head>
<body id="b1" onload="javascript:printWindow();">

<object id="factory" style="display:none" classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814" 
codebase="http://www.meadroid.com/scriptx/ScriptX.cab#Version=6,1,429,14">
</object>

<center><span style="font-size:11pt;">
<font face="����">
<b><%=Sdate%>���� <%=Fdate%>���� 

<% if Microbial_judge=1 then %>
          ��Ź ��ǰ �����Ƿڼ�<br></b></font></span>
<% elseif Microbial_judge=2 then %>          
          ��Ź ��ǰ �̻��� �����Ƿڼ�<br></b></font></span>
 <% elseif Microbial_judge=3 then %>          
          ��Ź ��ǰ �����Ƿڼ�(�̻��� ����)<br></b></font></span>         
  <% end if %>  

<table style="border-collapse:collapse;" align="center" cellspacing="0" width="720">
    <tr>
        <td width="720" style="border-width:0%; border-color:white; border-style:none;" align=right>&nbsp;
        </td>
    </tr>
</table>


<table style="border-collapse:collapse;" align="center" cellspacing="0" width="720" height="93%">
    <tr>
        <td width="720" style="border-width:0%; border-color:white; border-style:none;" valign="top">
<table border=1 cellspacing="0" cellpadding="0" width="720" align="center" bgcolor="#FAFDF6"  style='table-layout:fixed;'>
      <tr height="35">
        <td align="center" width="40"  bgcolor="#CCCCCC"><b>��ȣ</b></td>
        <td align="center" width="320" bgcolor="#CCCCCC"><b>�� ǰ ��</b></td>
        <td align="center" width="90"  bgcolor="#CCCCCC"><b>������</b></td>
        <td align="center" width="60"  bgcolor="#CCCCCC"><b>Lot No</b></td>
        <td align="center" width="60"  bgcolor="#CCCCCC"><b>����</b></td>
        <td align="center" width="50"  bgcolor="#CCCCCC"><b>�Ƿ���</b></td>
        <td align="center" width="100" bgcolor="#CCCCCC"><b>�� &nbsp;��</b></td>
      </tr> 
<center>


    <%
IF (RS.BOF and RS.EOF) Then
	Response.Write "<tr height=40> <td colspan=9 align=center>"
	Response.Write "��ȸ ��¥�� ��ϵ� �Ƿڰ� �����ϴ�."
	Response.Write "</td></tr>"

  S_number=0 '���� �ʱ� ��
  
 Else
	RS.AbsolutePage = IPage '�ش� �������� ù��° ���ڵ�� �̵��Ѵ�
	RCount = RS.pageSize
	Do while (NOT RS.EOF) and (RCount > 0 )

  Sid=RS("sid")
	Grp=RS("grp")
	Seq=RS("seq")
	Lev=RS("lev")
	Subject=RS("Subject")
	P_code=RS("P_code")
	ERP_No=RS("ERP_No")
	P_division=RS("P_division")
	
	Lot_number=RS("Lot_number")
	Mf_amount=RS("Mf_amount")
	Mf_unit=RS("Mf_unit")
	
	Mark_amount=RS("Mark_amount")
	Mark_unit=RS("Mark_unit")
	
	good_class=RS("good_class")
	
	Microbial_judge=RS("Microbial_judge")
	Registor_name=RS("Registor_name")
	Syear=RS("Syear")
	smonth=RS("smonth")
	sday=RS("sday")
	Remarks=RS("Remarks")
	  
	STime=RS("stime")
	STime_request=RS("STime_request")
	UTime=RS("utime")
	Visit=RS("visit")
	
	S_number=S_number+1
	
	 Remarks = replace (Remarks,"&","&amp;")
Remarks = replace (Remarks,"<","&lt;")
Remarks = replace (Remarks,">","&gt;")
Remarks = replace (Remarks,Chr(32),"&nbsp;") '����(�����̽�)
Remarks = replace (Remarks,Chr(13),"<br>") '�ٹٲ�
	
%> 
      <tr>
        <td style="text-align:center; text-indent:0; margin:0; padding-top:8px; padding-right:0px; padding-bottom:2px; padding-left:0px;">
        <%=S_number%></td>
        <td style="text-align:left; text-indent:0; margin:0; padding-top:8px; padding-right:5px; padding-bottom:2px; padding-left:5px;">
            <%=Subject%> &nbsp;[ <%=Mark_amount%>  <%=Mark_unit%> ]
        <% if Microbial_judge=2 then %>  
        <font color=red>��</font></b>
        <% else %>
        <% end if %>
        </td>
        <td style="text-align:center; text-indent:0; margin:0; padding-top:8px; padding-right:0px; padding-bottom:2px; padding-left:0px;">
        <%=Syear%>/<%=Smonth%>/<%=Sday%></td>
       <td style="text-align:center; text-indent:0; margin:0; padding-top:8px; padding-right:0px; padding-bottom:2px; padding-left:0px;">
        <%=Lot_number%></td>
        <td style="text-align:center; text-indent:0; margin:0; padding-top:8px; padding-right:0px; padding-bottom:2px; padding-left:0px;">
        
         <% =good_class%></td>
       <td style="text-align:center; text-indent:0; margin:0; padding-top:8px; padding-right:0px; padding-bottom:2px; padding-left:0px;">
        <%=Registor_name%></td>
        <td style="text-align:left; text-indent:0; margin:0; padding-top:8px; padding-right:5px; padding-bottom:2px; padding-left:5px;">
           <%=Remarks%>&nbsp;</td>
      </tr>  
  
    <%
	RS.MoveNext
	RCount = RCount -1
	Loop
End if


RS.Close
Set RS=nothing

%> 

  </table>
        </td>
    </tr>
    <tr>
        <td width="720" style="border-width:0%; border-color:white; border-style:none;" height="100">
<iframe src="../../../sign01.asp" border='0' frameBorder='0' scrolling='no' width='720' marginhieght='0' marginWidth='0' onload="this.style.height=this.contentWindow.document.body.scrollHeight"></iframe>

        </td>
    </tr>
</table>
<table width="720" cellspacing=0 cellpadding=0 border="0"  height="10">
<tr>
  <td align=left>QI-20-10-06</td>
  <td align=right>(��)�����Ѻ�</td>
</tr>
</table>
 <% end if %>
</body>
</html>
