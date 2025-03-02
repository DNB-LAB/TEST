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
'수정할 글번호를 전송받는다
Original_sid = Request.Querystring("Original_sid")

Set Upload = Server.CreateObject("TABSUpload4.Upload") 

'최대업로드용량
Upload.MaxBytesToAbort = Maxsize * 1024 * 1024 'Maxsize 는 inc.asp에서 정의

'업로드된 파일을 저장할 서버의 폴더
Spath = Server.MapPath(".") & "\Upload_temp" '임시로 저장할 폴더
Upload.Start Spath

W_Judge_Method = Upload.Form("W_Judge_Method")
W_Judge_Method = trim(Replace(W_Judge_Method,"'","''"))

W_Judge_Result = Upload.Form("W_Judge_Result")
W_Judge_Result = trim(Replace(W_Judge_Result,"'","''"))


W_Remarks = Upload.Form("W_Remarks")
W_Remarks = trim(Replace(W_Remarks,"'","''"))

Pass = Upload.Form("Pass")
Pass = trim(Replace(Pass,"'","''"))

W_Utime = now() '글수정 시각을 구한다

Set DB =Server.Createobject("ADODB.Connection")
DB.open ConnString


'등록자가 맞는지 확인한다
SQL = "SELECT Pass, sFile1,sFile2, sFile3  FROM " & STable
SQL = SQL & " WHERE Original_sid = " & Original_sid
Set RS = DB.Execute(SQL)


Oldfile1=RS("sFile1")
Oldfile2=RS("sFile2")
Oldfile3=RS("sFile3")


Oldpass=RS("pass")

RS.close
set RS=nothing


'db에 저장된 패스워드와 사용자가 입력한 패스워드를 비교함
IF (Pass = OldPass) or (Pass=adminpass) or (Pass=adminpass01) or (Pass=adminpass02) then  'adminpass는 관리자용==> inc.asp에 정의********************************

  '일치할 경우 수정한다
'/////////


'//첨부파일1이 새로 업되면 저장한다

If Upload.Form("sFile1").FileSize <> 0 Then
'기존 첨부파일을 삭제한다
If Oldfile1<>"" then

Set FS=Server.CreateObject("Scripting.FileSystemObject")
OldfilePath1=Spath & "\" & Oldfile1
On Error Resume Next
FS.DeleteFile(OldfilePath1)
set FS=nothing
End if

'새로 업된 첨부파일1을  저장한다
Upload.Form("sfile1").Save("D:\000_LNP_Db\021_Judge_System\Warehouse\Upload_01") '// 중복시 자동 시리얼이 붙음
 sFile1 = Upload.Form("sfile1").SaveName '// 저장된 전체경로와 파일이름 
 sFile1 = Mid(sFile1,instrrev(sFile1,"\")+1)'// 저장된 파일이름 

sFile1path = Spath & "\" & sFile1
on error resume next
Upload.Form("sFile1").SaveAs(sFile1path)
end if

'//첨부파일2이 새로 업되면 저장한다
If Upload.Form("sFile2").FileSize <> 0 Then
'기존 첨부파일2를 삭제한다
If Oldfile2<>"" then
Set FS=Server.CreateObject("Scripting.FileSystemObject")
OldfilePath2=Spath & "\" & Oldfile2
On Error Resume Next
FS.DeleteFile(OldfilePath1)
set FS=nothing
End if

'새로 업된 첨부파일2를  저장한다
Upload.Form("sFile2").Save("D:\000_LNP_Db\000_LNP_Db\021_Judge_System\Warehouse\Upload_02") '// 중복시 자동 시리얼이 붙음
 sFile2 = Upload.Form("sFile2").SaveName '// 저장된 전체경로와 파일이름 
 sFile2 = Mid(sFile2,instrrev(sFile2,"\")+1)'// 저장된 파일이름 

sFile2path = Spath & "\" & sFile2
on error resume next
Upload.Form("sFile2").SaveAs(sFile2path)
end if


'//첨부파일3이 새로 업되면 저장한다
If Upload.Form("sFile3").FileSize <> 0 Then
'기존 첨부파일3를 삭제한다
If Oldfile3<>"" then
Set FS=Server.CreateObject("Scripting.FileSystemObject")
OldfilePath3=Spath & "\" & Oldfile3
On Error Resume Next
FS.DeleteFile(OldfilePath1)
set FS=nothing
End if

'새로 업된 첨부파일3를  저장한다
Upload.Form("sFile3").Save("D:\000_LNP_Db\000_LNP_Db\021_Judge_System\Warehouse\Upload_03") '// 중복시 자동 시리얼이 붙음
 sFile3 = Upload.Form("sFile3").SaveName '// 저장된 전체경로와 파일이름 
 sFile3 = Mid(sFile3,instrrev(sFile3,"\")+1)'// 저장된 파일이름 

sFile3path = Spath & "\" & sFile3
on error resume next
Upload.Form("sFile3").SaveAs(sFile3path)
end if




'테이블에 저장한다

  SQL = " UPDATE " & STable
   SQL = SQL & " SET W_Judge_Method='" & W_Judge_Method & "'"
   SQL = SQL & ", W_Judge_Result='" & W_Judge_Result & "'"
   SQL = SQL & ", W_Remarks='" & W_Remarks & "'"
      SQL = SQL & ", W_Utime='" & W_Utime & "'"
    
  
	 
	if sFile1<>"" then '// 파일이름은 업된 경우만 수정한다
   SQL = SQL & ", sFile1='" & sFile1 & "'"
   end if
   
   if sFile2<>"" then '// 파일이름은 업된 경우만 수정한다
   SQL = SQL & ", sFile2='" & sFile2 & "'"
   end if
   
   if sFile3<>"" then '// 파일이름은 업된 경우만 수정한다
   SQL = SQL & ", sFile3='" & sFile3 & "'"
   end if
   
   
   
  
   SQL = SQL & " WHERE Original_sid = " & Original_sid
   
	 

 DB.Execute(SQL)

 DB.close
 set DB=nothing

	
	URL = "../View.asp?sid=" & Original_sid
	Response.Redirect URL

Else
	
	
%>

<body bgcolor="#D7F1FA">
<script language="javascript">
alert("암호가 틀립니다.-----");
history.back();
</script>


<%

DB.Close
Set DB = Nothing

Set Upload = Nothing
%>
<% end if %>
<% end if %>
