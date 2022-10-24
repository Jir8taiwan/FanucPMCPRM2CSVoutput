# FanucPMCPRM2CSVoutput
FANUC controller PMC parameter backup file converts to CSV and EXCEL format to read easily.

 Version. 2022.10.24

 Using PYTHON3 language to open and convert at FANUC controller system of parameter backup.
 Please copy "PMC1_PRM.TXT" in program folder together.
 It will output CSV and EXCEL files for studying in a formatted data.

 ****************************************************************
 If this small code is helping any needed case, it can donate
 to me for encourage as following address:
 1. BTC - 3M4wWghm4MxmrSfXmHMEeCFNwP8Lxxqjzk
 2. BCH - bitcoincash:qq6ghvdmyusnse9735rd5q09ensacl8z8qzrlwf49q
 3. LTC - MR6HaFkfkmsfifX3jWu7xz33dULGotVUWB
 4. DOGE- DGEFd3AAfJrBuaUwc4P6R2ZT754Jon9fQ7

If any usage advise or bug, I am happy to do improvement.
Thank you a lot.

中文說明內容：
- PMC1_PRM.TXT檔案是使用0iMF、0iMF+控制器的備份內容做開發，舊版的控制器尚未完整測試過，理論上也能處理輸出。
- 使用方式為執行FanucPMCPRM2CSVoutput.exe(python FanucPMCPRM2CSVoutput.py)時，需要放FANUC參數備份檔案CNC-PARA.TXT在同一個目錄下執行。
- 檔案處理後，會輸出四個檔案。分別為：
'1_PMC1_PRM.CSV'         #CSV輸出，無篩選。
'2_PMC1_PRM_trimed.CSV'  #CSV輸出，篩選空資料。
'3_PMC1_PRM.xlsx'        #EXCEL輸出，無篩選。
'4_PMC1_PRM_trimed.xlsx' #EXCEL輸出，篩選空資料。
- 請選擇需要使用的檔案讀取，EXCEL格式可以完整顯示8個BIT資料。

如果使用上有什麼問題，懇請建議。感謝。
