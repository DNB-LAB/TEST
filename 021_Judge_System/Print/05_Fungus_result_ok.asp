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

Fungus_year = request("Fungus_year")
Fungus_month = request("Fungus_month")
Fungus_day = request("Fungus_day")


Print_date = Fungus_year & "-" & Fungus_month & "-" & Fungus_day
Print_date = CDATE(Print_date) 

sql = " select * from AL_046_OEM_Product_Fungus_2018 INNER JOIN AL_046_OEM_Product_2018 ON AL_046_OEM_Product_Fungus_2018.original_sid=AL_046_OEM_Product_2018.sid "
SQL = SQL & " WHERE Fungus_date  = '" & Print_date & "' "  '������������ ����� ��ȸ���ڿ� ������ڰ� ������ �ڷḸ ����




'�������� �°� �����Ѵ�
SQL = SQL & " ORDER BY AL_046_OEM_Product_Fungus_2018.Sid ASC"

RS.Open SQL, ConnString

IF (RS.BOF and RS.EOF) Then
	TotRecord = 0 
	TotPage   = 0
Else
	TotRecord = RS.RecordCount
	Rs.pagesize=100 '���������� 100���� �����ش�
	TotPage   = RS.PageCount
End if



%>
<html>
<head>
<title>��,��Ź ODM �� ���ְ��� ��ǰ/��Ÿ ���� ������ �μ��ϱ�</title>
<link rel="stylesheet" href="basic.css">
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">


<center><span style="font-size:14pt;">
<font face="����">
<b><%=Print_date%>

<%if weekday(Print_date)=1 then %>(��)
          <%elseif weekday(Print_date)=2 then %>(��)
          <%elseif weekday(Print_date)=3 then %>(ȭ)
          <%elseif weekday(Print_date)=4 then %>(��)
          <%elseif weekday(Print_date)=5 then %>(��)
          <%elseif weekday(Print_date)=6 then %>(��)
          <%elseif weekday(Print_date)=7 then %>(��)<% end if %> 
&nbsp;&nbsp;<b>��Ź ODM �� ���ְ��� ��ǰ/��Ÿ ���� ������<br></b></font></span>
<table style="border-collapse:collapse;" align="center" cellspacing="0" width="720">
   <tr>
       <td width="60"  align=left>
       <a href="05_Fungus_result_ok_Print.asp?Fungus_year=<%=Fungus_year%>&Fungus_month=<%=Fungus_month%>&Fungus_day=<%=Fungus_day%>">
       <img src="../images/print.gif" border=0></a></td>
        <td width="360"  align=right>&nbsp;</td>
        <td width="300"  align=right>��ȸ�ð�: <%=now%>&nbsp;</td>
    </tr>
     <tr>
      <td  style="text-align:left; text-indent:0; margin:0; padding-top:20px; padding-right:5px; padding-bottom:0px; padding-left:0px;">            
      <form method="post" action="05_Fungus_result_ok_Print_Control.asp?sid=<%=Sid%>" name="form" OnSubmit="return Send()">
       <input type="image"   img src="../images/print.gif" border=0></td>
       <td  style="text-align:left; text-indent:0; margin:0; padding-top:15px; padding-right:5px; padding-bottom:0px; padding-left:0px;">            
        (��½ð�, ����� ���� ��, ISO ��ȣ, ȸ��� ���� ���)</a>
        <input type=hidden name=Fungus_year value="<%=Fungus_year%>">
        <input type=hidden name=Fungus_month value="<%=Fungus_month%>">
        <input type=hidden name=Fungus_day value="<%=Fungus_day%>">
        <input type=hidden name=Microbial_judge value="1"></td>
      
         <td   style="text-align:right; text-indent:0; margin:0; padding-top:5px; padding-right:5px; padding-bottom:0px; padding-left:0px;">            
  <input type="text" name="New_Now" size="40" style="text-align:right;" value="��½ð� : <%=Fungus_year%>-<%=Fungus_month%>-<%=Fungus_day%>&nbsp;<%=RIGHT(now,11)%>" ></td>
              </tr>
</table>


<table style="border-collapse:collapse;" align="center" cellspacing="0" width="720" height="93%">
    <tr>
        <td width="720" style="border-width:0%; border-color:white; border-style:none;" valign="top">
<table border=1 cellspacing="0" cellpadding="0" width="720" align="center" bgcolor="#FAFDF6">
        <tr height="28">    
        <td align="center" width="30" bgcolor="#CCCCCC"   ><b>��ȣ</b></center></td>
        <td align="center" width="40" bgcolor="#CCCCCC"   ><b>�Ƿ���</b></td>
        <td align="center" width="220" bgcolor="#CCCCCC"   ><b><font color="black">�� ǰ ��</font></b></td>
        <td align="center" width="70" bgcolor="#CCCCCC"   ><b>������</b></td>
        <td align="center" width="70" bgcolor="#CCCCCC"   ><b><font color="black">Lot No</font></b></td>
        <td align="center" width="60" bgcolor="#CCCCCC" ><font color="black"><b>���귮</b></font></td>
        <td align="center" width="60" bgcolor="#CCCCCC"  ><font color="black"><b>�� ��</b></font></td>
        <td align="center" width="50" bgcolor="#CCCCCC"  ><font color="black"><b>����</b></font></td>
        <td align="center" width="50" bgcolor="#CCCCCC" ><font color="black"><b>������</b></font></td>
        <td align="center" width="70" bgcolor="#CCCCCC"  ><b><font color="black">�� &nbsp;��</font></b></td>
      
      </tr> 
<center>


    <%
IF (RS.BOF and RS.EOF) Then
	Response.Write "<tr height=40> <td colspan=10 align=center>"
	Response.Write "��ȸ ��¥�� ��ϵ� ��Ź, ���ְ��� ��ǰ ���� ��������� �����ϴ�."
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
	
	
	Fungus_Registor=RS("Fungus_Registor")
	Fungus_result=RS("Fungus_result")
	Fungus_date=RS("Fungus_date")
	Fungus_remarks=RS("Fungus_remarks")
	
	
	S_number=S_number+1
	
	
%> 
    <tr>
        <td align="center" width="" bgcolor="white"  height="40"><p align="center"><%=S_number%></td>
        <td align="center" width="" bgcolor="white"><p align="center"><%=mid(left(Stime_request,10),6)%></td>
        <td style="text-align:left; text-indent:0; margin:0; padding-top:3px; padding-right:5px; padding-bottom:3px; padding-left:5px;">
            
         <%=Subject%>&nbsp;[ <%=Mark_amount%>  <%=Mark_unit%> ]</td>
        <td align="center" width="" bgcolor="white"><%=Syear%>/<%=Smonth%>/<%=Sday%></td>
        <td align="center" width="" bgcolor="white"><%=Lot_number%></td>
        <td align="center" width="" bgcolor="white"><%=formatnumber(Mf_amount,0)%> <%=Mf_unit%></td>
        <td align="center" width="" bgcolor="white">  <% if good_class="��Ź��ǰ" then %>
        ODM
        <% else %>
        <%=good_class%><% end if %></td>
        <td align="center" width="" bgcolor="white"><%=Fungus_result%></td>
        <td align="center" width="" bgcolor="white"><%=Fungus_Registor%></td>
        <td align="left" width="" bgcolor="white">&nbsp;<%=Micro_remarks%></td>
      
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
        <td   align=right style="border-width:0%; border-color:white; border-style:none;" height="100">

       ����� <select name="Approval" size="1" style="width:150px">
                    <option value="New">��¥ ��</option>
                    <option value="Old">��¥ ��</option>
                    </select></td>
    </tr>
    <tr>
     
        <td width="720" style="border-width:0%; border-color:white; border-style:none;" height="100">

<table width="720" cellspacing=0 cellpadding=0 border="0"  height="10">
<tr>
  <td align=left><input type="text" name="New_QI_No" size="30" value="QI-20-10-05"></td>
  <td align=right><select name="New_Company" size="1" style="width:150px">
                    <option value="(��)�����Ѻ�">(��)�����Ѻ�</option>
                    <option value="�Ѻ�ȭ��ǰ(��)">�Ѻ�ȭ��ǰ(��)</option>
                    </select></td>
</tr>

</table>

</body>
 <% end if %>
</html>
