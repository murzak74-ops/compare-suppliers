import streamlit as st
import pandas as pd

# --------------------------
# 1. Авторизация по e-mail
# --------------------------
AUTHORIZED_EMAILS = [
    "rab.org@bk.ru",
    "rab-organ@yandex.ru",
    "ooo.rab.org@gmail.com"
]

if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if not st.session_state.authenticated:
    st.title("🔒 Доступ только для своих")

    email = st.text_input("Введите ваш e-mail")

    if st.button("Войти"):
        if email.strip().lower() in [e.lower() for e in AUTHORIZED_EMAILS]:
            st.session_state.authenticated = True
            st.success("Добро пожаловать ✅")
        else:
            st.error("❌ У вас нет доступа к этому приложению")

    st.stop()

# ---- Стилизация ----
st.markdown(
    """
    <style>
    /* Общий фон */
    .main {
        background-color: #f8f9fa;
    }
    /* Заголовки */
    h1, h2, h3 {
        font-family: 'Arial', sans-serif;
    }
    /* Карточка инструкции */
    .instruction-box {
        background-color: #ffffff;
        border: 1px solid #dcdcdc;
        border-radius: 10px;
        padding: 15px 20px;
        margin-bottom: 20px;
        box-shadow: 2px 2px 6px rgba(0,0,0,0.05);
    }
    /* Таблицы */
    .stDataFrame tbody tr td {
        text-align: center;
    }
    /* Кнопка скачивания */
    .stDownloadButton button {
        background-color: #004080;
        color: white;
        border-radius: 8px;
        padding: 8px 16px;
        border: none;
    }
    .stDownloadButton button:hover {
        background-color: #0066cc;
        color: white;
    }
    </style>
    """,
    unsafe_allow_html=True
)

# ---- Шапка с логотипом ----
col1, col2 = st.columns([1, 3])
with col1:
    st.image("Логонов2.png", width=200)  # путь к файлу логотипа
with col2:
    st.markdown(
        """
        <h1 style='color:#004080; font-size:38px; margin-bottom:0px;'>Сравнение цен поставщиков</h1>
        <h3 style='color:gray; font-weight:normal; margin-top:5px;'>Компания «Рабочий Орган»</h3>
        """,
        unsafe_allow_html=True
    )

# ---- Инструкция ----
st.markdown(
    """
    <div class="instruction-box">
    <b>Загрузите Excel-файл с колонками:</b><br>
    - <b>Артикул</b><br>
    - <b>Кол-во</b><br>
    - пары колонок для каждого поставщика: <code>Цена_* </code> и <code>Производитель_* </code><br><br>
    <i>Нулевые и пустые цены учитываются только для отображения, но не влияют на выбор лучших.</i>
    </div>
    """,
    unsafe_allow_html=True
)

# ---- Функционал ----
uploaded_file = st.file_uploader("Загрузите Excel", type=["xlsx", "xls"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)

        if not {"Артикул", "Кол-во"}.issubset(df.columns):
            st.error("Файл должен содержать колонки 'Артикул' и 'Кол-во'.")
        else:
            # Определяем поставщиков
            suppliers = {}
            for col in df.columns:
                if col.startswith("Цена_"):
                    supplier_name = col.replace("Цена_", "")
                    prod_col = f"Производитель_{supplier_name}"
                    if prod_col in df.columns:
                        suppliers[supplier_name] = (col, prod_col)

            if not suppliers:
                st.error("Не найдены пары колонок вида 'Цена_*' и 'Производитель_*'.")
            else:
                # Преобразуем данные в длинный формат
                records = []
                for _, row in df.iterrows():
                    for supplier, (price_col, prod_col) in suppliers.items():
                        price = row[price_col]
                        producer = row[prod_col]
                        records.append({
                            "Артикул": row["Артикул"],
                            "Кол-во": row["Кол-во"],
                            "Поставщик": supplier,
                            "Цена": price if pd.notna(price) and price > 0 else None,
                            "Производитель": producer
                        })

                long_df = pd.DataFrame(records)

                mode = st.radio(
                    "Выберите режим отображения:",
                    ["Лучший поставщик", "ТОП-3 поставщиков"]
                )

                if mode == "Лучший поставщик":
                    best_records = []
                    for artikel, group in long_df.groupby("Артикул"):
                        valid = group[group["Цена"].notna()]
                        if not valid.empty:
                            idx = valid["Цена"].idxmin()
                            best_records.append(long_df.loc[idx])
                        else:
                            row = group.iloc[0].copy()
                            row["Цена"] = None
                            row["Поставщик"] = None
                            row["Производитель"] = None
                            best_records.append(row)
                    result = pd.DataFrame(best_records, columns=["Артикул", "Кол-во", "Производитель", "Поставщик", "Цена"])
                    st.subheader("Лучшие цены по каждому артикулу")
                    st.dataframe(result.reset_index(drop=True))

                else:  # ТОП-3 поставщиков в строку
                    def top_n_wide(group, n=3):
                        temp = group[group["Цена"].notna()].sort_values("Цена").head(n)
                        row = {
                            "Артикул": group["Артикул"].iloc[0],
                            "Кол-во": group["Кол-во"].iloc[0]
                        }
                        for i, (_, r) in enumerate(temp.iterrows(), 1):
                            row[f"Цена_{i}"] = r["Цена"]
                            row[f"Поставщик_{i}"] = r["Поставщик"]
                            row[f"Производитель_{i}"] = r["Производитель"]
                        for j in range(len(temp)+1, n+1):
                            row[f"Цена_{j}"] = None
                            row[f"Поставщик_{j}"] = None
                            row[f"Производитель_{j}"] = None
                        return pd.Series(row)

                    result = long_df.groupby("Артикул").apply(top_n_wide).reset_index(drop=True)
                    st.subheader("ТОП-3 предложений по каждому артикулу")
                    st.dataframe(result)

                # Экспорт
                @st.cache_data
                def convert_to_excel(df):
                    from io import BytesIO
                    output = BytesIO()
                    with pd.ExcelWriter(output, engine="openpyxl") as writer:
                        df.to_excel(writer, index=False, sheet_name="Результаты")
                    return output.getvalue()

                excel_data = convert_to_excel(result)

                st.download_button(
                    label="Скачать результат в Excel",
                    data=excel_data,
                    file_name="best_prices.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

    except Exception as e:
        st.error(f"Ошибка при обработке файла: {e}")