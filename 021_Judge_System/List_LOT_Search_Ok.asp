<% @LANGUAGE='VBSCRIPT' CODEPAGE='949' %> 
<% session.codepage = 949 %>
<% Response.ChaRset = "euc-kr" %>

<!-- #include file="inc.asp" -->

<%

Set RS = SERVER.CreateObject("ADODB.Recordset")
RS.CursorType = 3

LOT_NO_All = request("LOT_NO_All")


sql="SELECT AL_021_Judge_System.sid, Product_Code,Product_Name_DZ, Delivery_Amount,Lot_number_01,Lot_number_02, Lot_number_03, Lot_number_04, Lot_number_05, Lot_number_06, Lot_number_07, Lot_number_08, Expiration_Date_01,Expiration_Date_02, Expiration_Date_03, Expiration_Date_04, Expiration_Date_05, Expiration_Date_06, Expiration_Date_07, Expiration_Date_08, Lot_No_Divide, Good_class,  Judge_Result, Supplier, Warehouse, Manage_No, COA_Obtain, Warehouse_Date, Remarks, Registor_name, Visit,Sdate, W_Registor, W_Judge_Method, W_Judge_Result, W_Remarks  from AL_021_Judge_System left outer join  AL_022_Judge_Warehouse ON AL_021_Judge_System.sid=AL_022_Judge_Warehouse.original_sid "
SQL = SQL & " WHERE Lot_number_01  = '" & LOT_NO_All & "' "  '������������ ����� ��Ʈ��ȣ�� ������ �ڷḸ ����
SQL = SQL & " or Lot_number_02 =  '" & LOT_NO_All & "' "   
SQL = SQL & " or Lot_number_03 =  '" & LOT_NO_All & "' " 
SQL = SQL & " or Lot_number_04 =  '" & LOT_NO_All & "' " 
SQL = SQL & " or Lot_number_05 =  '" & LOT_NO_All & "' " 
SQL = SQL & " or Lot_number_06 =  '" & LOT_NO_All & "' " 
SQL = SQL & " or Lot_number_07 =  '" & LOT_NO_All & "' " 
SQL = SQL & " or Lot_number_08 =  '" & LOT_NO_All & "' " 




'�������� �°� �����Ѵ�
SQL = SQL & " ORDER BY Warehouse_Date ASC, sid"

RS.Open SQL, ConnString

IF (RS.BOF and RS.EOF) Then
	TotRecord = 0 
	TotPage   = 0
Else
	TotRecord = RS.RecordCount
	Rs.pagesize=100000 '���������� 10000���� �����ش�
	TotPage   = RS.PageCount
End if



%>
<html>
<head>
<title>1. �������� ���� ���� ��� �Ⱓ�� ��ȸ ���[��Ʈ��ȣ ��ü]</title>
<link rel="stylesheet" href="basic.css">
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<body bgcolor="#D7F1FA">


<table style="border-collapse:collapse;" align="center" cellspacing="0" width="1024">
   <tr height="28">
       <td w align=left>    
         ��Ʈ��ȣ ��ü����  <b><font color=red>[ <%=LOT_NO_All%> ]</font></b>��(��) �˻��� ���</td>
       
        <td width="260"  align=right>��ȸ�ð�: <%=now%>&nbsp;</td>
    </tr>
</table>



  <table border=1 cellspacing="0" cellpadding="0" width="1024" align="center" bgcolor="#FFFFFF"  style='table-layout:fixed;'>
        <!---�Խ��� ������ �̸� ���--->
         
       <tr height="42">
        <td align="center" width="40"  bgcolor="#CCCCCC"><b>��ȣ</b></td>
        <td align="center" width="60"  bgcolor="#CCCCCC"><b><b>�԰���</b></td>
        <td align="center" width="90" bgcolor="#CCCCCC"><b>ǰ���ڵ�</b></td>
        <td align="center" width="214" bgcolor="#CCCCCC"><b>�� ǰ ��</b></td>
        <td align="center" width="60"  bgcolor="#CCCCCC"><b><b>�԰����</b></td>
        <td align="center" width="70"  bgcolor="#CCCCCC"><b>Lot No</b></td>
        <td align="center" width="80"  bgcolor="#CCCCCC"><b>������</b></td>
        <td align="center" width="60"  bgcolor="#CCCCCC"><b>����ǰ<br>����</b></td>
        <td align="center" width="60"  bgcolor="#CCCCCC"><b>����<br>QC</b></td>
        <td align="center" width="150"  bgcolor="#CCCCCC"><b>�԰�ó<br>���ó</b></td>
        <td align="center" width="50"  bgcolor="#CCCCCC"><b>����<br>��ȣ</b></td>
        <td align="center" width="90" bgcolor="#CCCCCC"><b>��&nbsp;��</b></td>
      </tr> 
          <%
IF (RS.BOF and RS.EOF) Then
	Response.Write "<tr> <td colspan=12  align=center height=50 bgcolor=white><font color=red>"
	Response.Write "�˻��� ��Ʈ��ȣ�� ��Ȯ�� ��ġ�ϴ� �ڷᰡ�����ϴ�.&nbsp;&nbsp;&nbsp;&nbsp;�ٽ� �˻��غ��� �ٶ��ϴ�."
	Response.Write "</td></tr>"

S_number=0 '���� �ʱ� ��

Else

	
	RS.AbsolutePage = IPage '�ش� �������� ù��° ���ڵ�� �̵��Ѵ�
	RCount = RS.pageSize
	
	imsiNO=totrecord-(ipage-1)*(Rcount) '���ڵ��ȣ�� ����� �ӽù�ȣ
	
	Do while (NOT RS.EOF) and (RCount > 0 )

'�� ���������� ����� �ʵ��� ���� ���� �Ѳ����� �����ͼ� ������ ������ �д�.
'�̷��� �ѹ��� �������� �� ���� �����̴�.

	Sid=RS("sid")
	
	Warehouse_Date=RS("Warehouse_Date")
	Manage_No=RS("Manage_No")
	Product_Code=RS("Product_Code")
	Product_Name_DZ=RS("Product_Name_DZ")
	Delivery_Amount=RS("Delivery_Amount")
	Lot_No_Divide=RS("Lot_No_Divide")
	Good_class=RS("Good_class")
	Judge_Result=RS("Judge_Result")
	Supplier=RS("Supplier")
	Warehouse=RS("Warehouse")
	
	
	
	
	Lot_number_01=RS("Lot_number_01")
	Lot_number_02=RS("Lot_number_02")
	Lot_number_03=RS("Lot_number_03")
	Lot_number_04=RS("Lot_number_04")
	Lot_number_05=RS("Lot_number_05")
	Lot_number_06=RS("Lot_number_06")
	Lot_number_07=RS("Lot_number_07")
	Lot_number_08=RS("Lot_number_08")
	
	Expiration_Date_01=RS("Expiration_Date_01")
	Expiration_Date_02=RS("Expiration_Date_02")
	Expiration_Date_03=RS("Expiration_Date_03")
	Expiration_Date_04=RS("Expiration_Date_04")
	Expiration_Date_05=RS("Expiration_Date_05")
	Expiration_Date_06=RS("Expiration_Date_06")
	Expiration_Date_07=RS("Expiration_Date_07")
	Expiration_Date_08=RS("Expiration_Date_08")
	
	Remarks=RS("Remarks")
	Registor_name=RS("Registor_name")
	
	W_Registor=RS("W_Registor")
	W_Judge_Method=RS("W_Judge_Method")
	W_Judge_Result=RS("W_Judge_Result")
	W_Remarks=RS("W_Remarks")
	
 
S_number=S_number+1
 
 
 'Remarks = replace (Remarks,"&","&amp;")
 'Remarks = replace (Remarks,"<","&lt;")
 'Remarks = replace (Remarks,">","&gt;")
 'Remarks = replace (Remarks,Chr(32),"&nbsp;") '����(�����̽�)
 'Remarks = replace (Remarks,Chr(13),"<br>") '�ٹٲ�
   %>
 <tr>
        <td style="text-align:center; text-indent:0; margin:0; padding-top:5px; padding-right:0px; padding-bottom:1px; padding-left:0px;">
          <%=S_number%></td>
        <td style="text-align:center; text-indent:0; margin:0; padding-top:5px; padding-right:0px; padding-bottom:1px; padding-left:0px;">
         <%=mid(Warehouse_Date,3)%></td>
        <td style="text-align:left; text-indent:0; margin:0; padding-top:5px; padding-right:5px; padding-bottom:1px; padding-left:5px;">
           <%=Product_Code%></td>
        <td style="text-align:left; text-indent:0; margin:0; padding-top:3px; padding-right:5px; padding-bottom:1px; padding-left:5px;">
           <%=Product_Name_DZ%>
        </td>
       <td style="text-align:right; text-indent:0; margin:0; padding-top:3px; padding-right:5px; padding-bottom:1px; padding-left:0px;">
           <%=formatnumber(Delivery_Amount,0)%></td>
        <td style="text-align:center; text-indent:0; margin:0; padding-top:5px; padding-right:0px; padding-bottom:1px; padding-left:0px;">
          <%=Lot_number_01%>
          <% if Lot_number_02 <>"" then%><br><%=Lot_number_02%><% else %><% end if%>
          <% if Lot_number_03 <>"" then%><br><%=Lot_number_03%><% else %><% end if%>
          <% if Lot_number_04 <>"" then%><br><%=Lot_number_04%><% else %><% end if%>
          <% if Lot_number_05 <>"" then%><br><%=Lot_number_05%><% else %><% end if%>
          <% if Lot_number_06 <>"" then%><br><%=Lot_number_06%><% else %><% end if%>
          <% if Lot_number_07 <>"" then%><br><%=Lot_number_07%><% else %><% end if%>
          <% if Lot_number_08 <>"" then%><br><%=Lot_number_08%><% else %><% end if%></td>
        <td style="text-align:center; text-indent:0; margin:0; padding-top:5px; padding-right:0px; padding-bottom:1px; padding-left:0px;">
         <%=Expiration_Date_01%>
          <% if Expiration_Date_02 <>"" then%><br><%=Expiration_Date_02%><% else %><% end if%>
          <% if Expiration_Date_03 <>"" then%><br><%=Expiration_Date_03%><% else %><% end if%>
          <% if Expiration_Date_04 <>"" then%><br><%=Expiration_Date_04%><% else %><% end if%>
          <% if Expiration_Date_05 <>"" then%><br><%=Expiration_Date_05%><% else %><% end if%>
          <% if Expiration_Date_06 <>"" then%><br><%=Expiration_Date_06%><% else %><% end if%>
          <% if Expiration_Date_07 <>"" then%><br><%=Expiration_Date_07%><% else %><% end if%>
          <% if Expiration_Date_08 <>"" then%><br><%=Expiration_Date_08%><% else %><% end if%></td>
        <td style="text-align:center; text-indent:0; margin:0; padding-top:5px; padding-right:0px; padding-bottom:1px; padding-left:0px;">
        <%=Judge_Result%></td>
        <td style="text-align:center; text-indent:0; margin:0; padding-top:5px; padding-right:0px; padding-bottom:1px; padding-left:0px;">
        <% if W_Judge_Result<>"" then %><%=W_Judge_Result%><% else %>&nbsp;<% end if %></td>
       
       
        <td style="text-align:center; text-indent:0; margin:0; padding-top:5px; padding-right:0px; padding-bottom:1px; padding-left:0px;">
        <%=Supplier%><br><br style="line-height:3pt;"> <%=Warehouse%></td>
       
       
       
        <td style="text-align:center; text-indent:0; margin:0; padding-top:5px; padding-right:0px; padding-bottom:1px; padding-left:0px;">
           <% if Manage_No<> "" then %><%=Manage_No%><% else %>&nbsp;<% end if %></td>

         <td style="text-align:left; text-indent:0; margin:0; padding-top:3px; padding-right:0px; padding-bottom:1px; padding-left:5px;">
       <%=Remarks%>&nbsp;</td>
      </tr>  

          <%
	RS.MoveNext

	imsiNO=imsiNO-1
	RCount = RCount -1
	Loop
End if


RS.Close
Set RS=nothing

%>

        </table>
           <table align="center" width=1024 align="center"  style='table-layout:fixed;'>
         <tr height=50>
             <td  align=center>
             <a href= "javascript:window.open('about:blank', '_self').close();">
             <img src="images/close.gif" border="0"></a></td>
          </tr>
        </table>
      
</body>
</html>



