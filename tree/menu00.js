//소스수정자 : 이청희
  foldersTree = gFld("<span style=font-size:9pt;><b><font color=#000000>&nbsp;코스메틱 오케스트라라", "index.html")
		 
	aux1 = insFld(foldersTree, gFld("<span style=font-size:9pt;><font color=#000000>제품 정보 자료실", ""))	
	 insDoc(foldersTree, gLnk(2, "<span style=font-size:9pt;><font color=#000000>벌크 제품 Db(화장품)", "031_Bulk_Goods.html"))
	 insDoc(foldersTree, gLnk(2, "<span style=font-size:9pt;><font color=#000000>완제품 Db(화장품)", "032_Finsh_Goods.html"))
	 insDoc(foldersTree, gLnk(2, "<span style=font-size:9pt;><font color=#000000>벌크 제품 Db(기타)", "033_Bulk_Goods_Other.html"))
	 insDoc(foldersTree, gLnk(2, "<span style=font-size:9pt;><font color=#000000>완제품 Db(기타)", "034_Finsh_Goods_Other.html"))
   
   
   
 aux1 = insFld(foldersTree, gFld("<span style=font-size:9pt;><font color=#000000>제품 안정도 자료실", ""))	
	 insDoc(foldersTree, gLnk(2, "<span style=font-size:9pt;><font color=#000000>신제품 안정도 점검", "321_New_Product_Stability.html"))       
   insDoc(foldersTree, gLnk(2, "<span style=font-size:9pt;><font color=#000000>개발 제품(원료/소재) 안정도 점검", "322_Rnd_Product_Stability.html"))     
   
	aux1 = insFld(foldersTree, gFld("<span style=font-size:9pt;><font color=#000000>이상 발생 자료실", ""))	
	 insDoc(foldersTree, gLnk(2, "<span style=font-size:9pt;><font color=#000000>벌크제품 이상 발생 자료실", "402_Bulk_Claims.html"))       
   insDoc(foldersTree, gLnk(2, "<span style=font-size:9pt;><font color=#000000>완제품 이상 발생 자료실", "401_Claims.html"))       
  
  
   
 aux1 = insFld(foldersTree, gFld("<span style=font-size:9pt;><font color=#000000>이력 관련 자료실", ""))	  
 	insDoc(foldersTree, gLnk(2, "<span style=font-size:9pt;><font color=#000000>자재 테스트 자료실", "503_Sub_Material_Test.html")) 
	insDoc(foldersTree, gLnk(2, "<span style=font-size:9pt;><font color=#000000>원료 분석 Data", "702_HYDRODEEPER_Data.html")) 
  
  
 aux1 = insFld(foldersTree, gFld("<span style=font-size:9pt;><font color=#000000>각종 자료실", ""))	  
   insDoc(foldersTree, gLnk(2, "<span style=font-size:9pt;><font color=#000000>아이디어 제안 자료실", "801_Iead_Suggestion.html")) 
   insDoc(foldersTree, gLnk(2, "<span style=font-size:9pt;><font color=#000000>법령/시행규칙 관련 자료실", "901_Law_Regulation.html")) 
   insDoc(foldersTree, gLnk(2, "<span style=font-size:9pt;><font color=#000000>시장 조사 자료실", "911_Market_Research.html")) 
   insDoc(foldersTree, gLnk(2, "<span style=font-size:9pt;><font color=#000000>외부 교육/방문 세미나 자료실", "912_External_Education.html")) 
   
 aux1 = insFld(foldersTree, gFld("<span style=font-size:9pt;><font color=#000000>정보공유시스템 접근", ""))	
   insDoc(foldersTree, gLnk(2, "<span style=font-size:9pt;><font color=#000000>접근 권한 생성/수정", "001_LNP_Staff.html"))
