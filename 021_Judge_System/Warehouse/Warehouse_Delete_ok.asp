<%@Language="VBScript" CODEPAGE="65001" %>
<% 
Response.CharSet="utf-8" 
Session.codepage="65001" 
Response.codepage="65001" 
Response.ContentType="text/html;charset=utf-8" 
%>
<!-- #include file="inc.asp" -->
<%
if session("id")="" then  
'로그인하여 얻은 세션(id)가 없으면 로그인으로 돌려 보내고 있으면 리스트를 보여준다.
%>

<html>
<head>
<body >
<body leftmargin="0" topmargin="0" bgcolor="#D7F1FA">
<script language="javascript">
		alert("로그인이 필요합니다. 로그인 하세요! \n\n\혹은 로그인 됐더라도 오래되어 종료되었습니다.. \n\n\재 로그인이 필요합니다.  login please !!!");
	window.open('../../../Log_in.asp','end','width=310,height=190,top=270, left=350');
</script>

<% else %>

<!-- #include file="inc_Warwhouse.asp" -->

<%
'삭제할 글번호를 전송받는다
Sid = Request.QueryString("sid")
file_mode = Request.QueryString("file_mode")

'사용자가 입력한 비밀번호를 전송받는다
Pass = Request.Form("pass")

'데이타베이스의 비밀번호를 가져온다
Set DB = Server.CreateObject("ADODB.Connection")
DB.open ConnString


SQL = "select * from  AL_022_Judge_Warehouse"
SQL = SQL & " WHERE Original_sid =" & Sid
Set RS = DB.Execute(SQL)

Oldfile1 = rs("Sfile1")
Oldfile2 = rs("Sfile2")
Oldfile3 = rs("Sfile3")

OldPass = Rs("pass")




'두개의 비밀번호를 확인한다
IF (Pass = OldPass) or (Pass=adminpass) then  'adminpass는 관리자용==> inc.asp에 정의********************************

    Set DB = Server.Createobject("ADODB.Connection")
    DB.open ConnString
   
   
    if file_mode = "file1" then '첫번째 파일만 삭제 경우
      SQL = "UPDATE AL_022_Judge_Warehouse " 
      SQL = SQL & " Set sfile1='' WHERE Original_sid =" & sid
  
   elseif file_mode = "file2" then '두번째 파일만 삭제 경우
      SQL = "UPDATE AL_022_Judge_Warehouse " 
      SQL = SQL & " Set sfile2='' WHERE Original_sid =" & sid

   elseif file_mode = "file3" then '세번째 파일만 삭제 경우
      SQL = "UPDATE AL_022_Judge_Warehouse " 
      SQL = SQL & " Set sfile3='' WHERE Original_sid =" & sid
   

   else
      
      SQL = "DELETE  AL_022_Judge_Warehouse " 
      SQL = SQL & " WHERE Original_sid =" & sid
    end if
    DB.Execute(SQL)

    '첫번째 파일만 삭제 경우
    If file_mode = "file1" then '첫번째 파일만 삭제 경우

      Spath=Server.MapPath(".") & "\Upload_01\"  '상대경로
      OldfilePath1=Spath & OldFile1
      On Error Resume Next
      Set FS=Server.CreateObject("Scripting.FileSystemObject")
      FS.DeleteFile(oldfilepath1)
      Set FS=Nothing
  
  '2번째 파일만 삭제 경우
    elseif file_mode = "file2" then '2번째 파일만 삭제 경우

      Spath=Server.MapPath(".") & "\Upload_02\"  '상대경로
      OldfilePath2=Spath & OldFile2
      On Error Resume Next
      Set FS=Server.CreateObject("Scripting.FileSystemObject")
      FS.DeleteFile(oldfilepath2)
      Set FS=Nothing
 
  '3번째 파일만 삭제 경우
    elseif file_mode = "file3" then '3번째 파일만 삭제 경우

      Spath=Server.MapPath(".") & "\Upload_03\"  '상대경로
      OldfilePath3=Spath & OldFile3
      On Error Resume Next
      Set FS=Server.CreateObject("Scripting.FileSystemObject")
      FS.DeleteFile(oldfilepath3)
      Set FS=Nothing
  
      
   else
      Spath=Server.MapPath(".") & "\Upload_01\"  '상대경로
      OldfilePath1=Spath & OldFile1
      On Error Resume Next
      Set FS=Server.CreateObject("Scripting.FileSystemObject")
      FS.DeleteFile(oldfilepath1)
      Set FS=Nothing
      
      Spath=Server.MapPath(".") & "\Upload_02\"  '상대경로
      OldfilePath2=Spath & OldFile2
      On Error Resume Next
      Set FS=Server.CreateObject("Scripting.FileSystemObject")
      FS.DeleteFile(oldfilepath2)
      Set FS=Nothing
      
      Spath=Server.MapPath(".") & "\Upload_03\"  '상대경로
      OldfilePath3=Spath & OldFile3
      On Error Resume Next
      Set FS=Server.CreateObject("Scripting.FileSystemObject")
      FS.DeleteFile(oldfilepath3)
      Set FS=Nothing

      
      
     End if   

     
     
	URL = "../View.asp?sid=" & sid
	Response.Redirect URL


Else
%>
<body leftmargin="0" topmargin="0" bgcolor="#D7F1FA">
<script language="javascript">
alert("입력한 암호가 틀립니다! ");
history.back();
</script>
<%
End if
DB.Close
Set DB=nothing
%>

<% end if %>