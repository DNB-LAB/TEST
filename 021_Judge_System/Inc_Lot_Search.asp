<%

  ConnString = "Provider=SQLOLEDB;Data Source=(Local);Initial Catalog=LNP_COSMETICS;User ID=Sa;Password=1656;"  



STable=Request.QueryString("Table")  '게시판 링크화일에서 URL로 전송된 값 

'테이블 이름이 없을 경우 디폴트 지정
If Stable="" then
Stable="AL_023_Judge_Lot_List"
end if


SMode=Request.QueryString("Mode")    '목록보기화일에서 순서형/응답형으로 보기에서 전송된 값 
If (SMode = "") Then
    SMode = "qa"
End If

Var1 = "Table=" & STable

'검색할 필드와 단어를 전송받는다
Field = Request.QueryString("Field")
Str = Request.QueryString("Str")

Var1 = "Table=" & STable


SPage = Request.QueryString("Page") '페이지 정보가 필요한 모든 파일에서 전송된 값
If (SPage = "") Then
    SPage = "1"
End If

IPage = CInt(SPage)              'SPage변수의 값을 INT형으로 변환
Var2 = Var1 & "&Page=" & SPage

'검색할 필드와 단어를 전송받는다
Field = Request.QueryString("Field")
Str = Request.QueryString("Str")
Str = server.htmlencode(Str)

Var3 = Var2 & "&Field=" & Field
Var3 = Var3 & "&Str=" & Str




'==========  사용시 수정 대상  =================

Pagesize = 10 '출력할 레코드개수
Groupsize= 10 '출력할 페이지개수

adminpass="1656" '관리자용 수정삭제암호

Maxsize = 20 '최대허용 업로드 용량
'============================================
%>
