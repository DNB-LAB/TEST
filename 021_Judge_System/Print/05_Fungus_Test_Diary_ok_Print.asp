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

Print_date = syear & "-" & smonth & "-" & sday
Print_date = CDATE(Print_date) 

sql = " select ERP_No,P_division, Mark_amount, Mf_unit, Mf_amount, Mark_unit, Syear,smonth,sday, STime_request,Micro_name, AL_046_OEM_Product_Microorganism_2018.Micro_result, AL_046_OEM_Product_2018.Sid,AL_046_OEM_Product_2018.grp, AL_046_OEM_Product_2018.lev,P_code, AL_046_OEM_Product_2018.Seq,AL_046_OEM_Product_2018.Visit,AL_046_OEM_Product_2018.Mf_amount, AL_046_OEM_Product_2018.Mf_unit, AL_046_OEM_Product_2018.STime, AL_046_OEM_Product_2018.uTime, Microbial_judge, Subject, Lot_number, Good_class, Registor_name, Produce_date,  Fungus_result, Remarks from AL_046_OEM_Product_2018 left outer JOIN AL_046_OEM_Product_Fungus_2018 ON AL_046_OEM_Product_2018.sid=AL_046_OEM_Product_Fungus_2018.original_sid   left outer join     AL_046_OEM_Product_Microorganism_2018 ON AL_046_OEM_Product_Microorganism_2018.original_sid=AL_046_OEM_Product_2018.sid      "
SQL = SQL & " WHERE left(AL_046_OEM_Product_2018.STime, 10)  = '" & Print_date & "' "  





'�������� �°� �����Ѵ�
SQL = SQL & " ORDER BY AL_046_OEM_Product_Fungus_2018.Sid ASC"

RS.Open SQL, ConnString

IF (RS.BOF and RS.EOF) Then
	TotRecord = 0 
	TotPage   = 0
Else
	TotRecord = RS.RecordCount
	Rs.pagesize=1000 '���������� 100���� �����ش�
	TotPage   = RS.PageCount
End if

%>


<html>
<head>
<title>��Ź ODM �� ���ְ��� ��ǰ/��Ÿ ���� �������� �μ��ϱ�</title>
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

<center><span style="font-size:14pt;"><b>
<font face="����"><%=Print_date%>

<%if weekday(Print_date)=1 then %>(��)
          <%elseif weekday(Print_date)=2 then %>(��)
          <%elseif weekday(Print_date)=3 then %>(ȭ)
          <%elseif weekday(Print_date)=4 then %>(��)
          <%elseif weekday(Print_date)=5 then %>(��)
          <%elseif weekday(Print_date)=6 then %>(��)
          <%elseif weekday(Print_date)=7 then %>(��)<% end if %>
   
    ��Ź ODM �� ���ְ��� ��ǰ/��Ÿ ���� ��������<br></b></font></span>
             
<table style="border-collapse:collapse;" align="center" cellspacing="0" width="720">
    <tr>
        <td width="720" style="border-width:0%; border-color:white; border-style:none;" align=right>������ : 
        <u>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</u>
        </td>
    </tr>
</table>


<table style="border-collapse:collapse;" align="center" cellspacing="0" width="720" height="93%">
    <tr>
        <td width="720" style="border-width:0%; border-color:white; border-style:none;" valign="top">
<table border=1 cellspacing="0" cellpadding="0" width="720" align="center" bgcolor="#FAFDF6" style='table-layout:fixed;'>
      <tr height="22">
        <td align="center" width="30" bgcolor="#CCCCCC"  rowspan=2><b>��ȣ</b></center></td>
        <td align="center" width="260" bgcolor="#CCCCCC" rowspan=2  ><b><font color="black">��ǰ��</font></b></td>
        <td align="center" width="70" bgcolor="#CCCCCC"  rowspan=2><b>������</b></td>
        <td align="center" width="70" bgcolor="#CCCCCC"  rowspan=2 ><b><font color="black">Lot No</font></b></td>
        <td align="center" width="150" bgcolor="#CCCCCC"  colspan=5><font color="black"><b>���(��)</b></font></td>
        <td align="center" width="60" bgcolor="#CCCCCC"  rowspan=2><font color="black"><b>����</b></font></td>
        <td align="center" width="40" bgcolor="#CCCCCC"  rowspan=2><b><font color="black">���</font></b></td>
      </tr>  
      <tr height="22"> 
      <td align="center" width="30" bgcolor="#CCCCCC" ><font color="black"><b>1</b></font></td>
      <td align="center" width="30" bgcolor="#CCCCCC" ><font color="black"><b>2</b></font></td>
      <td align="center" width="30" bgcolor="#CCCCCC"><font color="black"><b>3</b></font></td>
      <td align="center" width="30" bgcolor="#CCCCCC" ><font color="black"><b>4</b></font></td>
      <td align="center" width="30" bgcolor="#CCCCCC"><font color="black"><b>5</b></font></td>
      </tr>  


    <%
IF (RS.BOF and RS.EOF) Then
	Response.Write "<tr height=40> <td colspan=11 align=center>"
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
	
	
%> 
    <tr>
        <td style="text-align:center; text-indent:0; margin:0; padding-top:5px; padding-right:0px; padding-bottom:3px; padding-left:0px;">            
        <%=S_number%></td>
        
        <td style="text-align:left; text-indent:0; margin:0; padding-top:5px; padding-right:0px; padding-bottom:3px; padding-left:5px;">            
       <%=Subject%> &nbsp;[ <%=Mark_amount%>  <%=Mark_unit%> ]
        
        </td>
        <td style="text-align:center; text-indent:0; margin:0; padding-top:5px; padding-right:0px; padding-bottom:3px; padding-left:0px;">            
        <%=Syear%>/<%=Smonth%>/<%=Sday%></td>
        <td style="text-align:center; text-indent:0; margin:0; padding-top:5px; padding-right:0px; padding-bottom:3px; padding-left:0px;">            
        <%=Lot_number%></td>
       <td style="text-align:center; text-indent:0; margin:0; padding-top:5px; padding-right:0px; padding-bottom:3px; padding-left:0px;">            
        &nbsp;</td>
        <td style="text-align:center; text-indent:0; margin:0; padding-top:5px; padding-right:0px; padding-bottom:3px; padding-left:0px;">            
        &nbsp;</td>
        <td style="text-align:center; text-indent:0; margin:0; padding-top:5px; padding-right:0px; padding-bottom:3px; padding-left:0px;">            
        &nbsp;</td>
        <td style="text-align:center; text-indent:0; margin:0; padding-top:5px; padding-right:0px; padding-bottom:3px; padding-left:0px;">            
        &nbsp;</td>
        <td style="text-align:center; text-indent:0; margin:0; padding-top:5px; padding-right:0px; padding-bottom:3px; padding-left:0px;">            
        &nbsp;</td>
         <td style="text-align:center; text-indent:0; margin:0; padding-top:5px; padding-right:0px; padding-bottom:3px; padding-left:0px;">            
        &nbsp;</td>
        <td style="text-align:center; text-indent:0; margin:0; padding-top:5px; padding-right:0px; padding-bottom:3px; padding-left:0px;">            
        &nbsp;</td>
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
      <td colspan=11 style="text-align:left; text-indent:0; margin:0; padding-top:5px; padding-right:3px; padding-bottom:3px; padding-left:5px;">            
      [ ��� : (+) ���� �缺 / (-) ���� ����] </td>
      </tr>
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
  <td align=left></td>
  <td align=right>(��)�����Ѻ�</td>
</tr>
</table>
 <% end if %>
</body>
</html>
