<%

  ConnString = "Provider=SQLOLEDB;Data Source=(Local);Initial Catalog=LNP_COSMETICS;User ID=Sa;Password=1656;"  



STable=Request.QueryString("Table")  '�Խ��� ��ũȭ�Ͽ��� URL�� ���۵� �� 

'���̺� �̸��� ���� ��� ����Ʈ ����
If Stable="" then
Stable="AL_023_Judge_Lot_List"
end if


SMode=Request.QueryString("Mode")    '��Ϻ���ȭ�Ͽ��� ������/���������� ���⿡�� ���۵� �� 
If (SMode = "") Then
    SMode = "qa"
End If

Var1 = "Table=" & STable

'�˻��� �ʵ�� �ܾ ���۹޴´�
Field = Request.QueryString("Field")
Str = Request.QueryString("Str")

Var1 = "Table=" & STable


SPage = Request.QueryString("Page") '������ ������ �ʿ��� ��� ���Ͽ��� ���۵� ��
If (SPage = "") Then
    SPage = "1"
End If

IPage = CInt(SPage)              'SPage������ ���� INT������ ��ȯ
Var2 = Var1 & "&Page=" & SPage

'�˻��� �ʵ�� �ܾ ���۹޴´�
Field = Request.QueryString("Field")
Str = Request.QueryString("Str")
Str = server.htmlencode(Str)

Var3 = Var2 & "&Field=" & Field
Var3 = Var3 & "&Str=" & Str




'==========  ���� ���� ���  =================

Pagesize = 10 '����� ���ڵ尳��
Groupsize= 10 '����� ����������

adminpass="1656" '�����ڿ� ����������ȣ

Maxsize = 20 '�ִ���� ���ε� �뷮
'============================================
%>
