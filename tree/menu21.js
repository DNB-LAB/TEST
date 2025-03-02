//소스수정자 : 이 청 희
foldersTree = gFld("<span style=font-size:9pt;><b><font color=#000000>&nbsp;종합연구원", "index.html")


	insDoc(foldersTree, gLnk(2, "<span style=font-size:9pt;><font color=#000000><b>업무일지</font>", "51_Duty_diary.html"))
	insDoc(foldersTree, gLnk(2, "<span style=font-size:9pt;><font color=#DF208B><b><u>新야(특)근일지</font>", "51_OT_Work.html"))
	insDoc(foldersTree, gLnk(2, "<span style=font-size:9pt;><font color=#000099><b>통합 샘플 송부 Db</font>", "52_OEM_Sample.html"))
	
	insDoc(foldersTree, gLnk(2, "<span style=font-size:9pt;><font color=#000099>샘플 송부 Db(3연 기존자료)</font>", "30_Samples.asp"))
	
	
	insDoc(foldersTree, gLnk(2, "<span style=font-size:9pt;><font color=#000000><b>행낭발송 자료실</font>", "54_Daily_Pouch.html"))
	
	
	
	insDoc(foldersTree, gLnk(2, "<span style=font-size:9pt;><font color=#000000>신제품 개발 담당 연구원</font>", "53_OEM_Resercher.html"))
	insDoc(foldersTree, gLnk(2, "<span style=font-size:9pt;><font color=#000000>월간일정", "01_Monthly_schedule.html"))
	insDoc(foldersTree, gLnk(2, "<span style=font-size:9pt;><font color=#000000>주간실적 및 계획", "06_Week_Schedule.html"))
	insDoc(foldersTree, gLnk(2, "<span style=font-size:9pt;><font color=#DF208B><b><u>新주간실적 및 계획", "06_Week_Schedule_New.html"))
  insDoc(foldersTree, gLnk(2, "<span style=font-size:9pt;><font color=#000000>연구소 공지사항", "02_RnD_board.html"))
	insDoc(foldersTree, gLnk(2, "<span style=font-size:9pt;><font color=#000000>원료 정보 변경 공지", "02_RnD_board_Raw_Material.html"))
	
	
	
	aux1 = insFld(foldersTree, gFld("<span style=font-size:9pt;><font color=#336600><b>견적 산출 및 통보 Db", ""))	
				insDoc(foldersTree, gLnk(2, "<span style=font-size:9pt;><font color=#336600>반제품 견적 자료실", "71_Raw_Cost_Quotation.html"))
				insDoc(foldersTree, gLnk(2, "<span style=font-size:9pt;><font color=#336600>기술영업부 통보 이력</font>", "72_Raw_Cost_OEM_Team_Notify.html"))
				insDoc(foldersTree, gLnk(2, "<span style=font-size:9pt;><font color=#336600>재무팀 통보 이력</font>", "73_Raw_Cost_Finance_Team_Notify.html"))
				
	aux1 = insFld(foldersTree, gFld("<span style=font-size:9pt;><font color=#000000>원료/처방 관련 Db", ""))	
					insDoc(foldersTree, gLnk(2, "<span style=font-size:9pt;><font color=#000000>화장품 성분 사전", "10_Cosmetic_Ingredient_Name.html"))
					insDoc(foldersTree, gLnk(2, "<span style=font-size:9pt;><font color=#000000>잇츠한불 원료 사전</font>", "09_HB_Raw_Dictionary.html"))
					insDoc(foldersTree, gLnk(2, "<span style=font-size:9pt;><font color=#000000>신중국허용원료사전(2015)</font>", "../Db_04_RnD_Center_China_Dic/08_China_New_Dictionary_2015.html"))
					
					insDoc(foldersTree, gLnk(2, "<span style=font-size:9pt;><font color=#000000>신중국허용원료사전(2014)</font>", "../Db_04_RnD_Center_China_Dic/08_China_New_Dictionary_2014.html"))
					insDoc(foldersTree, gLnk(2, "<span style=font-size:9pt;><font color=#000000>중국 허용 원료 사전</font>", "../Db_04_RnD_Center_China_Dic/08_China_Raw_Dictionary_2004.html"))
					insDoc(foldersTree, gLnk(2, "<span style=font-size:9pt;><font color=#000000>원료 Db", "11_Raw_Material.html"))
					insDoc(foldersTree, gLnk(2, "<span style=font-size:9pt;><font color=#8B0000>후저우 원료 Db", "11_Huzhou_Raw.html"))
					insDoc(foldersTree, gLnk(2, "<span style=font-size:9pt;><font color=#000000>자재팀 원료 Db", "11_Raw_Material_Team.html"))
					insDoc(foldersTree, gLnk(2, "<span style=font-size:9pt;><font color=#000000>처방 등록/조회/수정", "12_Product_formula.html")) 
					insDoc(foldersTree, gLnk(2, "<span style=font-size:9pt;><font color=#000000>원료업체 면담 자료", "05_Offer_Meeting.html"))       
          
		aux1 = insFld(foldersTree, gFld("<span style=font-size:9pt;><font color=#000000>제품 충전/포장 Db", ""))	
					insDoc(foldersTree, gLnk(2, "<span style=font-size:9pt;><font color=#000000>시작품 포장 사전 점검", "91_Paking_Pre_Inspection.html"))
					insDoc(foldersTree, gLnk(2, "<span style=font-size:9pt;><font color=#000000>시작품 포장보고서</font>", "92_Paking_Report.html"))
				
	
	
	aux1 = insFld(foldersTree, gFld("<span style=font-size:9pt;><font color=#000000>전산처방 관련 Db", ""))	
					insDoc(foldersTree, gLnk(2, "<span style=font-size:9pt;><font color=#000000>반제품 전산/조회/수정", "13_E_formula.html"))
					insDoc(foldersTree, gLnk(2, "<span style=font-size:9pt;><font color=#000000>BASE 등록/조회/수정", "14_Base_E_formula.html"))
					insDoc(foldersTree, gLnk(2, "<span style=font-size:9pt;><font color=#000000>반제품/BASE 전산 변경", "15_E_formula_change.html"))       
					insDoc(foldersTree, gLnk(2, "<span style=font-size:9pt;><font color=#000000>자재 ERP 등록/조회/수정", "16_Material_ERP_No.html"))  
					insDoc(foldersTree, gLnk(2, "<span style=font-size:9pt;><font color=#000000>칭량용 처방 등록/수정", "16_Weighing_Recipe.html"))  
					insDoc(foldersTree, gLnk(2, "<span style=font-size:9pt;><font color=#000000>기존제품 현재 원자재비", "16_Price_calculation_from_Weighing_Recipe.html"))  
					
					
		
	 		
	aux1 = insFld(foldersTree, gFld("<span style=font-size:9pt;><font color=#000000>평가 관련 Db", ""))
	        insDoc(aux1, gLnk(0, "<span style=font-size:9pt;><font color=#000000>타사원료 평가결과", "21_Other_RM_Evaluation.html"))	
					insDoc(aux1, gLnk(0, "<span style=font-size:9pt;><font color=#0064FF>新타사원료 평가결과", "21_New_Other_RM_Evaluation.html"))	
          insDoc(aux1, gLnk(0, "<span style=font-size:9pt;><font color=#000000>자사원료 평가결과", "22_Our_RM_Evaluation.html"))	
          insDoc(aux1, gLnk(0, "<span style=font-size:9pt;><font color=#000000>신소재팀 전용 자료", "23_New_MT_Only.html"))	
          insDoc(aux1, gLnk(0, "<span style=font-size:9pt;><font color=#000000>제품 효능 평가결과", "24_Product_Efficacy_Evaluation.html"))	
          insDoc(aux1, gLnk(0, "<span style=font-size:9pt;><font color=#000000>방부력 평가결과", "25_Preservation_Evaluation.html"))	
          insDoc(aux1, gLnk(0, "<span style=font-size:9pt;><font color=#006400>新방부력(의뢰/결과)", "26_New_Preservation.html"))	
    
    
  
					
					      
 
    aux1 = insFld(foldersTree, gFld("<span style=font-size:9pt;><font color=#000000>제품분석 Db", ""))
	        insDoc(aux1, gLnk(0, "<span style=font-size:9pt;><font color=#000000>기초제품류", "31_SK_analysis.html"))	
					insDoc(aux1, gLnk(0, "<span style=font-size:9pt;><font color=#000000>립스틱/립글로스류", "32_LS_analysis.html"))	
          insDoc(aux1, gLnk(0, "<span style=font-size:9pt;><font color=#000000>MB/FB/CS/SC류", "33_MB_analysis.html"))	
          insDoc(aux1, gLnk(0, "<span style=font-size:9pt;><font color=#000000>파우더류", "34_Powder_analysis.html"))	
          insDoc(aux1, gLnk(0, "<span style=font-size:9pt;><font color=#000000>MS/EL류", "35_MS_analysis.html"))	
          
          insDoc(aux1, gLnk(0, "<span style=font-size:9pt;><font color=#000000>기초제품류(3연 기존자료)", "30_SC_analysis.asp"))
          insDoc(aux1, gLnk(0, "<span style=font-size:9pt;><font color=#000000>색조제품류(3연 기존자료)", "30_MC_analysis.asp"))


   aux1 = insFld(foldersTree, gFld("<span style=font-size:9pt;><font color=#000000>자료실", ""))	
					insDoc(aux1, gLnk(0, "<span style=font-size:9pt;><font color=#000000>세미나 및 학회자료", "41_Seminar_Society.html"))
					insDoc(aux1, gLnk(0, "<span style=font-size:9pt;><font color=#000000>정기간행물 요약", "42_Periodical_Publications_Summary.html"))
					insDoc(aux1, gLnk(0, "<span style=font-size:9pt;><font color=#000000>양식/서식/회사규정", "43_Form_of_all_sort.html"))
					insDoc(aux1, gLnk(0, "<span style=font-size:9pt;><font color=#000000>실험기기 자료", "44_Machine_History.html"))
					insDoc(aux1, gLnk(0, "<span style=font-size:9pt;><font color=#000000>도서 자료 Db", "45_Book_Db.html"))
					insDoc(aux1, gLnk(0, "<span style=font-size:9pt;><font color=#000000>저널 복사 자료", "46_Jr_Copy.html"))
					insDoc(aux1, gLnk(0, "<span style=font-size:9pt;><font color=#000000>특허 관련 자료", "47_Patents.html"))
		
	
	 aux1 = insFld(foldersTree, gFld("<span style=font-size:9pt;><font color=#000000>품질보증팀 공유 자료실", ""))	
          insDoc(foldersTree, gLnk(2, "<span style=font-size:9pt;><font color=#000000>식약처 행정 처분 자료실", "81_KFDA_Measure.html")) 
          insDoc(foldersTree, gLnk(2, "<span style=font-size:9pt;><font color=#000000>벌크 안정도 자료실", "82_Bulk_Stability.html"))    
          insDoc(foldersTree, gLnk(2, "<span style=font-size:9pt;><font color=#000000>분석실 시약 자료실</font>", "62_QA_Reagent.html"))			
	        insDoc(foldersTree, gLnk(0, "<span style=font-size:9pt;><font color=#000000>분석실 실험 의뢰/결과</font>", "63_QA_Anlaysis.html"))
		
	
		aux1 = insFld(foldersTree, gFld("<span style=font-size:9pt;><font color=#000000>시약 관련 자료", ""))	
					insDoc(foldersTree, gLnk(2, "<span style=font-size:9pt;><font color=#000000>신소재 시약 자료실", "61_Reagent.html"))
					insDoc(foldersTree, gLnk(0, "<span style=font-size:9pt;><font color=#000000>정부과제 자료실", "64_Goverment_Project.html"))
						
		
		
					
					
		aux1 = insFld(foldersTree, gFld("<span style=font-size:9pt;><font color=#000000>공지사항/게시판", ""))	
					insDoc(aux1, gLnk(2, "<span style=font-size:9pt;><font color=#000000>연구소 공지사항", "02_RnD_board.html"))
					insDoc(aux1, gLnk(2, "<span style=font-size:9pt;><font color=#000000>연구소 앨범", "18_Photo_Album.html"))
					
	                    
