<!DOCTYPE html>
<% @LANGUAGE='VBSCRIPT' CODEPAGE='949' %> 
<% session.codepage = 949 %>
<% Response.ChaRset = "euc-kr" %>

<!-- #include file="inc.asp" -->

<html>
<head>
<title>시장출하 적부 판정 기록 조회</title>
<body bgcolor="#D7F1FA">

<br>
<table width="800" style="border-collapse:collapse;" cellspacing="0" align="center">
      <tr>
            <td height="40"   align="center">
      <span style="font-size:11pt;"><b><font face="돋움">▶ 일일 시장출하 적부 판정 기록 조회</font></td>
  </tr>
  
  <tr>
  <td align="right" >
       <a href="javascript:history.go(-1)"><img src=../images/back.gif border=0></a>&nbsp;
      <a href="../list.asp"><img src=../images/list.gif border=0></a></td>
    </tr>
</table>

    <table border="0" width="810" align="center">
      <tr>
         <td  valign="top" align="center">
         <iframe src="00_Warehouse.asp" frameborder="0" width="810" height=100 marginwidth="1" marginheight="1" name="1" scrolling="no" allowtransparency="true"></iframe>
        </td>
    </tr>
     
    
</table>
</body>
</form>
</html>
