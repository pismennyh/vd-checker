import streamlit as st
import pandas as pd
import requests
import urllib.parse
import time
import re

st.set_page_config(page_title="Проверка запчастей VolgaDiesel", layout="wide")
st.title("Проверка наличия запчастей на vd-dizel.ru")
st.markdown("Загрузите Excel-файл. Программа проверит артикулы и выдаст готовый файл.")

uploaded_file = st.file_uploader("Загрузите файл Excel (.xlsx)", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    
    article_col = None
    for col in df.columns:
        if "артикул" in str(col).lower() or "код" in str(col).lower() or "наименование" in str(col).lower():
            article_col = col
            break
    if not article_col:
        article_col = df.columns[0]
        
    st.info(f"Будем проверять столбец: **{article_col}**")

    if st.button("Начать проверку"):
        progress_bar = st.progress(0)
        status_text = st.empty()
        results = []
        links = []
        
        session = requests.Session()
        headers = {
            "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
            "Accept-Language": "ru-RU,ru;q=0.9",
        }
        
        # Заходим на главную, чтобы получить куки сессии
        session.get("https://vd-dizel.ru/", headers=headers, timeout=10)
        
        total_rows = len(df)
        
        for index, row in df.iterrows():
            article = str(row[article_col]).strip()
            status_text.text(f"Проверка: {article} ({index + 1} из {total_rows})")
            
            if not article or article.lower() == "nan":
                results.append("Пусто")
                links.append("")
                progress_bar.progress((index + 1) / total_rows)
                continue
            
            # ГЛАВНОЕ ИСПРАВЛЕНИЕ: добавляем search_performed=Y
            search_url = f"https://vd-dizel.ru/?dispatch=products.search&search_performed=Y&q={urllib.parse.quote(article)}"
            
            try:
                response = session.get(search_url, headers=headers, timeout=15)
                text = response.text.lower()
                
                # Проверка 1: ищем явные признаки отсутствия
                not_found_markers = ["ничего не найдено", "товары не найдены", "по вашему запросу ничего"]
                is_not_found = any(marker in text for marker in not_found_markers)
                
                # Проверка 2: если маркеров нет, проверяем, встречается ли артикул в тексте больше 1 раза
                # (1 раз он всегда есть в поле поиска, 2+ раз значит он есть в карточке товара)
                article_count = text.count(article.lower())
                
                if not is_not_found and article_count > 1:
                    results.append("✅ Есть")
                    # Пытаемся вытащить ссылку
                    link_match = re.search(rf'href="(https://vd-dizel\.ru/[^"]*{re.escape(article.replace(".", "-"))}[^"]*)"', response.text)
                    if link_match:
                        links.append(link_match.group(1))
                    else:
                        links.append("https://vd-dizel.ru")
                else:
                    results.append("❌ Нет")
                    links.append("")
                    
            except Exception as e:
                results.append("⚠️ Ошибка")
                links.append("")
                
            time.sleep(0.8) # Оптимальная задержка для большого списка
            progress_bar.progress((index + 1) / total_rows)
            
        df["Статус на сайте"] = results
        df["Ссылка"] = links
        status_text.text("Готово! Файл сформирован.")
        
        output = "проверенные_заявки.xlsx"
        df.to_excel(output, index=False)
        
        with open(output, "rb") as f:
            st.download_button(
                label="📥 Скачать готовый Excel",
                data=f,
                file_name=output,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
