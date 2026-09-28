import streamlit as st
import pandas as pd
import requests
import urllib.parse
import time
import re

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
        links = []
        
        headers = {
            "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/115.0.0.0 Safari/537.36",
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
            "Accept-Language": "ru-RU,ru;q=0.9,en-US;q=0.8,en;q=0.7"
        }
        
        total_rows = len(df)
        
        for index, row in df.iterrows():
            article = str(row[article_col]).strip()
            status_text.text(f"Проверка: {article} ({index + 1} из {total_rows})")
            
            if not article or article.lower() == "nan":
                results.append("Пусто")
                links.append("")
                progress_bar.progress((index + 1) / total_rows)
                continue
            
            # Используем Bing поиск
            search_query = f'site:vd-dizel.ru {article}'
            bing_url = f"https://www.bing.com/search?q={urllib.parse.quote(search_query)}"
            
            try:
                response = requests.get(bing_url, headers=headers, timeout=10)
                
                if response.status_code == 200:
                    # Проверяем наличие результатов от vd-dizel.ru
                    has_results = 'vd-dizel.ru' in response.text
                    
                    if has_results:
                        # Пытаемся извлечь ссылку
                        link_match = re.search(r'href="(https://vd-dizel\.ru/[^"]+)"', response.text)
                        if link_match:
                            results.append("✅ Есть")
                            links.append(link_match.group(1))
                        else:
                            results.append("✅ Есть")
                            links.append("https://vd-dizel.ru")
                    else:
                        results.append("❌ Нет")
                        links.append("")
                else:
                    results.append("⚠️ Ошибка сети")
                    links.append("")
                    
            except Exception as e:
                results.append("⚠️ Ошибка сети")
                links.append("")
                
            time.sleep(1)
            progress_bar.progress((index + 1) / total_rows)
            
        df["Статус на сайте"] = results
        df["Ссылка"] = links
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
