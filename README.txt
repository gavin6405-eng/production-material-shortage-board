Streamlit Cloud 修正版

1. 將 app.py 與 requirements.txt 上傳到 GitHub 儲存庫根目錄。
2. Streamlit Main file path 填 app.py。
3. 網頁開啟後，上傳製令缺料 Excel。

修正內容：未附預設 Excel 時不再發生 KeyError，而是提示使用者上傳檔案。
