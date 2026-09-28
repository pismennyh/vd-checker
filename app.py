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
        
        # Создаем сессию для сохранения cookies
        session = requests.Session()
        
        headers = {
            "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/115.0.0.0 Safari/537.36",
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8",
            "Accept-Language": "ru-RU,ru;q=0.9,en-US;q=0.8,en;q=0.7",
            "Accept-Encoding": "gzip, deflate, br",
            "Connection": "keep-alive",
            "Upgrade-Insecure-Requests": "1",
            "Sec-Fetch-Dest": "document",
            "Sec-Fetch-Mode": "navigate",
            "Sec-Fetch-Site": "none",
            "Cache-Control": "max-age=0"
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
            
            # Прямой запрос к поиску на сайте
            search_url = f"https://vd-dizel.ru/?dispatch=products.search&q={urllib.parse.quote(article)}"
            
            try:
                response = session.get(search_url, headers=headers, timeout=15, allow_redirects=True)
                
                if response.status_code == 200:
                    html_content = response.text.lower()
                    
                    # Проверяем несколько признаков наличия товара
                    has_nothing = "ничего не найдено" in html_content or "товаров не найдено" in html_content
                    has_results = (
                        article.lower() in html_content and 
                        "результаты поиска" in html_content and
                        not has_nothing
                    )
                    
                    # Ищем ссылку на товар
                    link_pattern = rf'href="(https://vd-dizel\.ru/[^"]*{re.escape(article.replace(".", "-").replace("_", "-"))}[^"]*)"'
                    link_match = re.search(link_pattern, response.text)
                    
                    if has_results or link_match:
                        results.append("✅ Есть")
                        if link_match:
                            links.append(link_match.group(1))
                        else:
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
                
            time.sleep(2)
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
