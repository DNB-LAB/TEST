<% @LANGUAGE='VBSCRIPT' CODEPAGE='949' %> 
<% session.codepage = 949 %>
<% Response.ChaRset = "euc-kr" %>

<!-- #include file="inc.asp" -->

<html>
<title>�������� ���� ���� ��� �ڷ��</title>
<link rel="stylesheet" href="basic.css">
<body bgcolor="#D7F1FA">
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
</head>

<center>

<%
Set RS = SERVER.CreateObject("ADODB.Recordset")
RS.CursorType = 3

SQL="SELECT AL_021_Judge_System.Sid as HD_sid, AL_021_Judge_System.Grp as HD_Grp,AL_021_Judge_System.Seq as HD_Seq, Product_Code,Product_Name_DZ, Delivery_Amount,Lot_number_01,Lot_number_02,Lot_number_03,Lot_number_04,Lot_number_05,Lot_number_06,Lot_number_07,Lot_number_08,Judge_Result,Supplier,Good_class,COA_Obtain,Registor_name, Warehouse,Visit,Sdate, Manage_No, W_Judge_Result,Warehouse_Date from AL_021_Judge_System left outer join  AL_022_Judge_Warehouse ON AL_021_Judge_System.sid=AL_022_Judge_Warehouse.original_sid "


'�˻��̳� �ƴϳĿ� ����
if Str<>"" then 
SQL = SQL & " WHERE " & Field
SQL = SQL & " LIKE '%" & Str & "%'"

end if

'������/�������� ���� ���ļ����� �ٸ��� �Ѵ�
IF SMode = "qa" Then
	SQL = SQL & " ORDER BY HD_sid DESC"
Else 
	SQL = SQL & " ORDER BY HD_Grp DESC,HD_Seq"
END IF 

RS.Open SQL, ConnString




IF (RS.BOF and RS.EOF) Then
	TotRecord = 0 
	TotPage   = 0
Else
	TotRecord = RS.RecordCount
	Rs.pagesize = Pagesize '���������� ������ ���ڵ尳��(inc.asp���� ����)
	TotPage   = RS.PageCount
End if
%>
  <table cellspacing=0 cellpadding=0 border=0 width="1024" align="center" valign=top>
  <tr>
  <td align="center" bgcolor="#D7F1FA">

      <table cellspacing=0 cellpadding=2 border=0 width="1024"  align=center  style='table-layout:fixed;'>
          <tr height="40"> 
            <td width=424 align="left">
                           <b>�� 
              <%if Str<>"" then%>
              �˻����:</b>&nbsp;
                <% if field="Product_Name_DZ"  then %>����ǰ������
                <% elseif field="Product_Code"  then %>��ǰ �ڵ��
                <% elseif field="Lot_number_All"  then %>��Ʈ��ȣ ��ü��
                <% elseif field="Lot_number_01" then %>��Ʈ��ȣ 1��
                <% elseif field="Lot_number_02" then %>��Ʈ��ȣ 2��
                <% elseif field="Lot_number_03" then %>��Ʈ��ȣ 3����
                <% elseif field="Lot_number_04" then %>��Ʈ��ȣ 4��
                <% elseif field="Lot_number_05" then %>��Ʈ��ȣ 5��
                <% elseif field="Lot_number_06" then %>��Ʈ��ȣ 6����
                <% elseif field="Lot_number_07" then %>��Ʈ��ȣ 7��
                <% elseif field="Lot_number_08" then %>��Ʈ��ȣ 8��
                <% elseif field="Good_class" then %>��ǰ ������
                <% elseif field="Judge_Result" then %>���� �����
                <% elseif field="COA_Obtain" then %>������ �Լ���
                <% elseif field="Supplier" then %>������
                <% elseif field="Warehouse" then %>�԰�ó��
                <% elseif field="Manage_No" then %>������ȣ��
                <% elseif field="Registor_name" then %>�����ڷ�
                
                <% end if %>
              <b><font color=#000088><%=Str%></font></b>(��)�� �˻��� ���
              <% else %>
            �������� ���� ���� ���
             
              <%end if %>
            </td>
           
           
              <td  width=200 align="left" width="360"">&nbsp;&nbsp;
             Total : <%=TotRecord%> &nbsp;<%=IPage%>/<%=TotPage%> Pages</td>
             
             <td width="400" align="right"> 
             <a href="Search_Code.asp?<%=Var1%>"><img src="images/Code_Serch.gif"  border="0" valign=bottom align="top"></a>
             &nbsp;&nbsp;&nbsp;&nbsp;
             <a href="Search_Name.asp?<%=Var1%>"><img src="images/Name_Serch.gif"  border="0" valign=bottom align="top"></a>
             &nbsp;&nbsp;&nbsp;&nbsp;
           
            <a href="list.asp?<%=Var1%>"><img src="images/list.gif"  border="0"></a></td>
            </tr>
            </table>
      
      
        <table cellspacing=0 width="1024" cellpadding="3" align=center bgcolor=#FFFFFF align=center  style='table-layout:fixed;'>
          <!---�Խ��� ������ �̸� ���--->
          <tr  height=2> 
          <td bgcolor=#336600></td>
          </tr>
           </table>
      
          <table cellspacing=0 width="1024" cellpadding="3" align=center bgcolor=#FFFFFF align=center  style='table-layout:fixed;'>
          <tr height="50" align="center"> 
      <td width="50"  bgcolor="#CCFFCC"><b>��ȣ</td>
      <td width="100" bgcolor="#CCFFCC"><b>����ǰ�ڵ�</td>
      <td width="319" bgcolor="#CCFFCC"><b>�� ǰ ��</td>
      <td width="50"  bgcolor="#CCFFCC"><b>����</td>
      <td width="90"  bgcolor="#CCFFCC"><b>LOT</td>
      <td width="50"  bgcolor="#CCFFCC"><b>����ǰ</td>
      <td width="55"  bgcolor="#CCFFCC"><b>����QC</td>
      
      <td width="150"  bgcolor="#CCFFCC"  align="center"><b>������<br><br style="line-height:3pt;"> �԰�ó</td>
      <td width="50"  bgcolor="#CCFFCC"  align="center"><b>������</td>
      <td width="35"  bgcolor="#CCFFCC"><b>��ȸ</td>
      <td width="75"  bgcolor="#CCFFCC"><b>�԰���<br><br style="line-height:3pt;"> �����</td>
    </tr>
    <tr  height=1> 
          <td bgcolor=#336600 colspan="11"></td>
          </tr>
          <%
IF (RS.BOF and RS.EOF) Then
	Response.Write "<tr height=50> <td colspan=11 align=center height=30><font color=red>"
	Response.Write "�˻� ���ǿ� �ش�Ǵ� �ڷᰡ�����ϴ�.&nbsp;&nbsp;&nbsp;&nbsp;�ٸ� �������� �˻��غ��� �ٶ��ϴ�."
	Response.Write "</td></tr>"


Else

	
	RS.AbsolutePage = IPage '�ش� �������� ù��° ���ڵ�� �̵��Ѵ�
	RCount = RS.pageSize
	
	imsiNO=totrecord-(ipage-1)*(Rcount) '���ڵ��ȣ�� ����� �ӽù�ȣ
	
	Do while (NOT RS.EOF) and (RCount > 0 )

'�� ���������� ����� �ʵ��� ���� ���� �Ѳ����� �����ͼ� ������ ������ �д�.
'�̷��� �ѹ��� �������� �� ���� �����̴�.

	HD_sid=RS("HD_sid")
	HD_Grp=RS("HD_Grp")
	HD_Seq=RS("HD_Seq")
	
	
	Product_Code=RS("Product_Code")
	Product_Name_DZ=RS("Product_Name_DZ")
	Delivery_Amount=RS("Delivery_Amount")
	
	Lot_number_01=RS("Lot_number_01")
	Lot_number_02=RS("Lot_number_02")
	Lot_number_03=RS("Lot_number_03")
	Lot_number_04=RS("Lot_number_04")
	Lot_number_05=RS("Lot_number_05")
	Lot_number_06=RS("Lot_number_06")
	Lot_number_07=RS("Lot_number_07")
	Lot_number_08=RS("Lot_number_08")
	
	Judge_Result=RS("Judge_Result")
	COA_Obtain=RS("COA_Obtain")
	Manage_No=RS("Manage_No")
	COA_Obtain=RS("COA_Obtain")
	
	
	
	Supplier=RS("Supplier")
	Good_class=RS("Good_class")
	Warehouse=RS("Warehouse")
	
	Visit=RS("Visit")
	Sdate=RS("Sdate")
  Registor_name=RS("Registor_name")

	W_Judge_Result=RS("W_Judge_Result")
 
  Warehouse_Date=RS("Warehouse_Date")
 
 
    '����ǰ���� ���̰� �ʹ� ��� ���ٷ� �Ѿ�°� �����ϱ� ���� ���� ���
   If len(Product_Name_DZ)>50 then
      str1=".."
      else
      str1=""
     end if

        
 
    '�ۼ����� 24�ð��̳���� new!��� �޽����� �غ��Ѵ�
    Ndate=date()        '���� ����
       IF datediff("h",Stime,Ndate) < 24 Then       
		  Msg2="<font color=red>��</font>"
       Else
		  Msg2=""
       End if
 
       
  '�̻��� ���� ���ο� ���� �ٸ� �������� list�� �ѷ��ش�    
     IF Micro_result="����" Then     '������ �ƴ϶��
   		  Msg3="<font color=blue>����</font>"
     elseif Micro_result="����" Then
       Msg3="<font color=red>����</font>"
      elseif  Micro_result="������" Then                 
	   	  Msg3="<font color=red>������</font>"
	   else                  
	   	  Msg3=""
       end If
 
    %>
     <tr onMouseOut="this.style.background='#FFFFFF'" onMouseOver="this.style.background='#ffff00'" >
  
         
           <!--��ȣ�� ����Ѵ�-->
            <td  style="text-align:center; text-indent:0; margin:0; padding-top:8px; padding-right:0px; padding-bottom:3px; padding-left:0px;">
            <%=imsiNO%></td>
              <!--����ǰ �ڵ带 ����Ѵ�-->
            <td  style="text-align:left; text-indent:0; margin:0; padding-top:8px; padding-right:0px; padding-bottom:3px; padding-left:5px;">
          <% if Good_class="ȭ��ǰ" then %>
          <font color=blue><%=Product_Code%></font>
          <% else %>
         <font color=green><%=Product_Code%></a>
          <% end if %></td>
            <td  style="text-align:left; text-indent:0; margin:0; padding-top:8px; padding-right:0px; padding-bottom:3px; padding-left:5px;">
          <font color="black">
              <%  
	           IF SMode = "qa" Then
	           For I=1 To Lev
			       Response.Write("&nbsp;")
		         next
             End if 
              %>
             <%=Msg1%>
            <a href="view.asp?<%=Var2%>&sid=<%=HD_sid%>">
             <%=Product_Name_DZ%></a></td>
   
    
    <td style="text-align:right; text-indent:0; margin:0; padding-top:8px; padding-right:5px; padding-bottom:3px; padding-left:0px;">
          <%=formatnumber(Delivery_Amount,0)%></td>
     
      <!--��Ʈ�ѹ��� ����Ѵ�--> 
     <td  style="text-align:center; text-indent:0; margin:0; padding-top:8px; padding-right:0px; padding-bottom:3px; padding-left:0px;">
     <%=Lot_number_01%></td>
     
    <!--����ǰ ��������� ����Ѵ�-->
     <!--���� ��������� ����Ѵ�-->
     <td  style="text-align:center; text-indent:0; margin:0; padding-top:8px; padding-right:0px; padding-bottom:3px; padding-left:0px;">
      <% if Judge_Result="����"  then %>
            <img src="images/OK.gif">
         <% elseif Judge_Result="������"   then %>
          <img src="images/No.gif">
          <% elseif Judge_Result="����"  then %>
          <img src="images/Hold.gif">
      <% elseif Judge_Result="���԰�"  then %>
          <img src="images/Aready.gif">
    
    

          <% else %>
          <img src="images/Not_Regist.gif">
          
          <% end if %></td>
          
         <td  style="text-align:center; text-indent:0; margin:0; padding-top:8px; padding-right:0px; padding-bottom:3px; padding-left:0px;">
     
          
           <% if W_Judge_Result="����"  then %>
            <img src="images/OK.gif">
         <% elseif W_Judge_Result="������"   then %>
          <img src="images/No.gif">
          <% elseif W_Judge_Result="����"  then %>
          <img src="images/Hold.gif">
    
          <% else %>
          &nbsp;
          
          <% end if %></td>
     <!--������ ����Ѵ�-->
    <td  style="text-align:center; text-indent:0; margin:0; padding-top:8px; padding-right:0px; padding-bottom:3px; padding-left:0px;">
       <%=Supplier%><br><br style="line-height:3pt;">
       <%=Warehouse%></td>
       
       <td  style="text-align:center; text-indent:0; margin:0; padding-top:8px; padding-right:0px; padding-bottom:3px; padding-left:0px;">
        <% if COA_Obtain="�Լ�"  then %>
            <img src="images/icon_accept.gif">
         <% elseif COA_Obtain="���Լ�"   then %>
          <img src="images/action_stop.gif">
          <% elseif COA_Obtain="���Լ�"  then %>
          <img src="images/icon_package_get.gif">
    
          <% else %>
          &nbsp;
          
          <% end if %> </td>
       
          <!--�԰��� ����Ѵ�-->
     <td  style="text-align:center; text-indent:0; margin:0; padding-top:8px; padding-right:0px; padding-bottom:3px; padding-left:0px;">
     <%=Visit%></td>
   <td  style="text-align:center; text-indent:0; margin:0; padding-top:8px; padding-right:0px; padding-bottom:3px; padding-left:0px;">
   <%=mid(Warehouse_Date,3)%>
   
    <br><br style="line-height:3pt;">
     <%=mid(Sdate,3)%></td>
  </tr>
          <tr>
            <td colspan="11" height="1" bgcolor="#BBBBBB"></td>
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
          <tr> 
            <td colspan="11" height=1></td>
          </tr>
          <tr  height=1> 
          <td bgcolor=#336600 colspan="11"></td>
          </tr>
        </table>
        <table width="1024" cellpadding="2" cellspacing="0" border="0" align=center>
          <tr> 
            <td height="50" align="center"> 
<%
TAr = Var1 & "&Page=1"
%>
              [<a href="list.asp?<%=Tar%>">&lt;</a>] 
<%
	Gsize = Groupsize 
	PreGNum = (IPage - 1) \ GSize
	EndPNum  = PreGNum * GSize
		Tar = Var1 & "&Page=" & EndPNum
	IF ( EndPNum > 0 ) Then
%>
              [<a href="list.asp?<%=Tar%>">�� ��</a>] 

<% 
End IF
%>
[<%
	LCount = GSize
	intI = EndPNum + 1
	Do While (LCount > 0) and (intI <= TotPage)

    '���� �������� ���� ǥ��	 
	if intI = Ipage then
	   intI2 = "<font color=tomato>" & intI & "</font>"
    else
	   intI2 = intI
	end if
		Tar = Var1 & "&Page=" & intI
%> &nbsp;<a href="list.asp? <%=TAr%>" ><%=intI2%></a>
<%
	intI = intI + 1
	LCount = LCount - 1
	Loop
%>
&nbsp;]
<%
	intI = EndPNum + (GSize + 1)
		TAr = Var1 & "&Page=" & intI
	If (intI <= TotPage ) Then
%>
[<a href="list.asp?<%=TAr%>">����</a>] 
<% 
End if 
TAr = Var1 & "&Page=" & Totpage
%>
              [<a href="list.asp?<%=Tar%>">&gt;</a>] </td>
          </tr>
        </table>
              
        <table width="1024" cellpadding="2" cellspacing="0" border="0" align=center>
          <form method="Get" Action="list.asp">
            <tr> 
              <td align="right">
               <input  type="hidden" name="Table" value="<%=Stable%>">
                <select name="Field" style="width:110">
                  <option value="Product_Name_DZ" <% if field="Product_Name_DZ" then %> selected <% end if %>>����ǰ��
                  <option value="Product_Code"    <% if field="Product_Code" then %> selected <% end if %>>��ǰ �ڵ�
                  
                  
                  <option value="Lot_number_01"  <% if field="Lot_number_01" then %> selected <% end if %>>��Ʈ��ȣ1
                  <option value="Lot_number_02"  <% if field="Lot_number_02" then %> selected <% end if %>>��Ʈ��ȣ2
                  <option value="Lot_number_03"  <% if field="Lot_number_03" then %> selected <% end if %>>��Ʈ��ȣ3
                  <option value="Lot_number_04"  <% if field="Lot_number_04" then %> selected <% end if %>>��Ʈ��ȣ4
                  <option value="Lot_number_05"  <% if field="Lot_number_05" then %> selected <% end if %>>��Ʈ��ȣ5
                  <option value="Lot_number_06"  <% if field="Lot_number_06" then %> selected <% end if %>>��Ʈ��ȣ6
                  <option value="Lot_number_07"  <% if field="Lot_number_07" then %> selected <% end if %>>��Ʈ��ȣ7
                  <option value="Lot_number_08"  <% if field="Lot_number_08" then %> selected <% end if %>>��Ʈ��ȣ8
                                    
                  <option value="Good_class"     <% if field="Good_class" then %> selected <% end if %>>��ǰ ����
                  <option value="Judge_Result"   <% if field="Judge_Result" then %> selected <% end if %>>���� ���
                  <option value="COA_Obtain"     <% if field="COA_Obtain" then %> selected <% end if %>>������ �Լ�
                  <option value="Supplier"       <% if field="Supplier" then %> selected <% end if %>>������
                  <option value="Warehouse"      <% if field="Warehouse" then %> selected <% end if %>>�԰�ó
                  <option value="Manage_No"      <% if field="Manage_No" then %> selected <% end if %>>������ȣ
                  <option value="Registor_name"  <% if field="Registor_name" then %> selected <% end if %>>�����</select> 
                  
                  <input type="text" name="Str" value="<%=Str%>" style="width:400">
                  <input  type="submit" value="�˻�"> 
                  <% if str<>"" then %>
                 <input  type="button" onClick="document.location.href='list.asp?Table=<%=Stable%>'" value="���">
                <%end if%></td>
              <td align="right" width="170"><a href="Search/Search_condition.asp"><img src="images/inquiry.gif"  border="0"></a></td>
              <td align="center" width="80"><a href="Print/Print_condition.asp"><img src="images/icon_calendar.png"  border="0"></a></td>
            </tr>
          </form>
        </table>
        
        




<script language="javascript">
function Send() {
	var vA = document.form.LOT_NO_All.value;

	if (vA == "") {
		alert("��Ʈ��ȣ�� �Է��ϼ���.\n");
		document.form.LOT_NO_All.focus();
		return false;
		}



return true;
} // end function
//  -->
</script>


        <form method=get action="List_LOT_Search_Ok.asp?<%=Var5%>"  name="form" onSubmit="return Send()" target="_blank">
        <br style="line-height:8pt;"> 
        <table width="1024" cellpadding="2" cellspacing="0" border="0" align=center>
          <tr> 
              <td align="right">
              <select name="" style="width:150"> 
              <option value="">��Ʈ��ȣ ��ü �˻�</option></td>
              <td align="left" width="190"><input type="text" name="LOT_NO_All" size="25" maxlength="20"></td>
             <td align="left" width="50"><input type="image" img src="images/btn_search.gif" border="0"></td>
              <td width="450" style="text-align:left; text-indent:0; margin:0; padding-top:8px; padding-right:0px; padding-bottom:3px; padding-left:5px;">
              �� ��Ʈ��ȣ ��ü �˻��� ��Ȯ�� ��ġ�ϴ� ��ȣ�� �˻��˴ϴ�!(���� ����)</td>
            </tr>
             </form>
         </table>
        
	  </td>
    </tr>
   
  </table>
<br><br style="line-height:50pt;"> 

</center>
</body>
</html>
