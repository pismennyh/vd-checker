import streamlit as st
import pandas as pd
import requests
import urllib.parse
import time

st.set_page_config(page_title="Проверка запчастей VolgaDiesel", layout="wide")
st.title("Проверка наличия запчастей на vd-dizel.ru")
st.markdown("Загрузите Excel-файл. Программа найдет столбец с артикулами, проверит их наличие на сайте и выдаст обновленный файл.")

uploaded_file = st.file_uploader("Загрузите файл Excel (.xlsx)", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    
    # Автоматический поиск столбца с артикулами
    article_col = None
    for col in df.columns:
        if "артикул" in str(col).lower() or "код" in str(col).lower() or "наименование" in str(col).lower():
            article_col = col
            break
    
    if not article_col:
        article_col = df.columns[0]
        
    st.info(f"Для проверки будет использован столбец: **{article_col}**")

    if st.button("Начать проверку"):
        progress_bar = st.progress(0)
        status_text = st.empty()
        results = []
        
        headers = {
            "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/115.0.0.0 Safari/537.36"
        }
        
        total_rows = len(df)
        
        for index, row in df.iterrows():
            article = str(row[article_col]).strip()
            status_text.text(f"Проверка: {article} ({index + 1} из {total_rows})")
            
            if not article or article.lower() == "nan":
                results.append("Пусто")
                progress_bar.progress((index + 1) / total_rows)
                continue
                
            search_url = f"https://vd-dizel.ru/?dispatch=products.search&q={urllib.parse.quote(article)}"
            
            try:
                response = requests.get(search_url, headers=headers, timeout=5)
                if response.status_code == 200 and article in response.text:
                    results.append("✅ Есть")
                else:
                    results.append("❌ Нет")
            except Exception:
                results.append("⚠️ Ошибка сети")
                
            time.sleep(0.3) # Задержка для снижения нагрузки на сайт
            progress_bar.progress((index + 1) / total_rows)
            
        df["Статус на сайте"] = results
        status_text.text("Готово!")
        
        # Формирование файла для скачивания
        output = "проверенные_заявки.xlsx"
        df.to_excel(output, index=False)
        
        with open(output, "rb") as f:
            st.download_button(
                label="📥 Скачать готовый Excel",
                data=f,
                file_name=output,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )