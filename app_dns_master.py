import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(
    page_title="DNS • Направление «Обеспечение»",
    page_icon="📊",
    layout="wide"
)

st.markdown("""
    <style>
    .main { background-color: #f8fafc; }
    .dns-header {
        background: linear-gradient(135deg, #0f172a 0%, #1e293b 100%);
        border-top: 6px solid #FF6600;
        padding: 34px 42px;
        border-radius: 16px;
        color: white;
        margin-bottom: 25px;
        box-shadow: 0 10px 30px rgba(15, 23, 42, 0.15);
    }
    .metric-card {
        background: white;
        padding: 20px;
        border-radius: 12px;
        border: 1px solid #e2e8f0;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05);
    }
    .stTabs [data-baseweb="tab-list"] { gap: 12px; }
    .stTabs [data-baseweb="tab"] {
        background-color: #f1f5f9;
        border-radius: 8px 8px 0px 0px;
        padding: 10px 22px;
        font-weight: 700;
        color: #334155;
    }
    .stTabs [aria-selected="true"] {
        background-color: #FF6600 !important;
        color: white !important;
    }
    </style>
""", unsafe_allow_html=True)

st.markdown("""
    <div class="dns-header">
        <h1 style="margin:0; font-size: 32px; font-weight: 800; letter-spacing: -0.5px;">DNS • Направление «Обеспечение»</h1>
        <p style="margin:8px 0 0 0; opacity: 0.9; font-size: 16px;">Операционный аналитический пульт управления филиалами (Июль — Сентябрь 2026)</p>
    </div>
""", unsafe_allow_html=True)

def safe_float(val):
    if pd.isna(val):
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip()
    if s in ['', '-', 'nan', 'None', 'N/A']:
        return 0.0
    try:
        return float(s.replace(',', '.'))
    except ValueError:
        return 0.0

@st.cache_data
def load_all_data():
    costs_df = pd.read_excel("Цифры для дашборда.xlsx", sheet_name="база данных")
    # Clean month formatting to ensure exact matching and sorting
    costs_df['Месяц'] = costs_df['Месяц'].astype(str).str.strip().str.capitalize()
    
    grades_file = "Оценки за три месяца.xlsx"
    xls = pd.ExcelFile(grades_file)
    grade_sheets = [s for s in xls.sheet_names if 'критерий' not in s.lower()]
    
    crit_sheet_name = next((s for s in xls.sheet_names if 'критерий' in s.lower()), grade_sheets[0])
    criteria_df = pd.read_excel(grades_file, sheet_name=crit_sheet_name)
    
    all_branches = set()
    for sheet in grade_sheets:
        df = pd.read_excel(grades_file, sheet_name=sheet)
        for b in df[df.columns[0]].dropna():
            b_str = str(b).strip()
            if b_str and b_str != 'nan':
                all_branches.add(b_str)
                
    for b in costs_df['Филиал'].dropna():
        b_str = str(b).strip()
        if b_str and b_str != 'nan':
            all_branches.add(b_str)
            
    master_branches = sorted(list(all_branches))
    
    grades_data = {}
    for sheet in grade_sheets:
        df = pd.read_excel(grades_file, sheet_name=sheet)
        col_branch = df.columns[0]
        col_score = df.columns[-1]
        
        sheet_dict = {}
        for idx, row in df.iterrows():
            b = str(row[col_branch]).strip()
            if b and b != 'nan':
                score = safe_float(row[col_score])
                sheet_dict[b] = {
                    "Списания в срок": safe_float(row.get('списание в срок', 0)),
                    "Комментарии": safe_float(row.get('верность комментарий', 0)),
                    "Статья затрат": safe_float(row.get('правильно выбранная статья затрат', 0)),
                    "Самозакуп": safe_float(row.get('Контроль самозакупа', 0)),
                    "План затрат": safe_float(row.get('Проверка плана по затратам', 0)),
                    "Итоговый балл": round(score, 2),
                    "Есть данные": True
                }
                
        month_rows = []
        clean_sheet_name = sheet.replace('г.', '').strip().capitalize()
        for b in master_branches:
            if b in sheet_dict:
                month_rows.append({"Филиал": b, "Месяц": clean_sheet_name, **sheet_dict[b]})
            else:
                month_rows.append({
                    "Филиал": b, "Месяц": clean_sheet_name,
                    "Списания в срок": 0.0, "Комментарии": 0.0, "Статья затрат": 0.0,
                    "Самозакуп": 0.0, "План затрат": 0.0, "Итоговый балл": 0.0,
                    "Есть данные": False
                })
        grades_data[clean_sheet_name] = pd.DataFrame(month_rows)
        
    return grade_sheets, grades_data, costs_df, criteria_df

grade_sheets_raw, grades_data, costs_df, criteria_df = load_all_data()
clean_sheets = [s.replace('г.', '').strip().capitalize() for s in grade_sheets_raw]

st.sidebar.markdown("### 🎛️ Навигация по пультам")
dashboard_choice = st.sidebar.radio(
    "Выберите дашборд:",
    [
        "⭐ Дашборд №1: Оценки и качество филиалов",
        "💰 Дашборд №2: Управление затратами"
    ]
)

st.sidebar.markdown("---")
st.sidebar.info("💡 **Подсказка:** Используйте фильтры месяца и города/филиала в верхней части каждого дашборда для детального просмотра.")

if "Оценки" in dashboard_choice:
    st.markdown("## ⭐ Дашборд №1: Мониторинг качества, оценок и критериев филиалов")
    
    col_f1, col_f2, col_f3 = st.columns(3)
    with col_f1:
        selected_grade_month = st.selectbox(
            "📅 Выберите месяц:",
            clean_sheets,
            key="g_m_sel"
        )
    with col_f2:
        all_grade_branches = sorted(grades_data[selected_grade_month]['Филиал'].unique())
        selected_grade_branch = st.selectbox(
            "🏢 Выберите город / филиал:",
            ["Все филиалы"] + all_grade_branches,
            key="g_b_sel"
        )
    with col_f3:
        status_filter = st.selectbox(
            "🎯 Категория филиалов:",
            [
                "Все филиалы",
                "Отличники (≥ 9.0)",
                "Норма (7.0 - 8.9)",
                "Критичные (< 7.0)"
            ],
            key="g_stat_sel"
        )
        
    df_g_current = grades_data[selected_grade_month].copy()
    
    if selected_grade_branch != "Все филиалы":
        df_g_current = df_g_current[df_g_current['Филиал'] == selected_grade_branch]
        
    if status_filter == "Отличники (≥ 9.0)":
        df_g_current = df_g_current[(df_g_current["Есть данные"]) & (df_g_current["Итоговый балл"] >= 9.0)]
    elif status_filter == "Норма (7.0 - 8.9)":
        df_g_current = df_g_current[(df_g_current["Есть данные"]) & (df_g_current["Итоговый балл"] >= 7.0) & (df_g_current["Итоговый балл"] < 9.0)]
    elif status_filter == "Критичные (< 7.0)":
        df_g_current = df_g_current[(df_g_current["Есть данные"]) & (df_g_current["Итоговый балл"] < 7.0)]
        
    active_g = df_g_current[df_g_current["Есть данные"]]
    avg_s = active_g["Итоговый балл"].mean() if len(active_g) > 0 else 0
    
    mg1, mg2, mg3 = st.columns(3)
    with mg1:
        st.metric("Средний балл выборки", f"{avg_s:.2f} / 10")
    with mg2:
        st.metric("Филиалов в выборке", len(df_g_current))
    with mg3:
        st.metric("В зоне риска (< 7.0)", len(active_g[active_g["Итоговый балл"] < 7.0]))
        
    st.markdown("---")
    
    tab_g_main, tab_g_prog, tab_g_crit = st.tabs([
        "📊 Рейтинг и критерии филиала",
        "📈 Прогрессия из месяца в месяц",
        "📋 Официальные критерии оценивания"
    ])
    
    with tab_g_main:
        sub_tg1, sub_tg2 = st.columns([7, 5])
        with sub_tg1:
            st.markdown(f"##### Стилизованный рейтинг филиалов ({selected_grade_month})")
            chart_grades = df_g_current[df_g_current["Есть данные"]].sort_values(by="Итоговый балл", ascending=True)
            if len(chart_grades) > 0:
                fig_g = px.bar(
                    chart_grades,
                    x="Итоговый балл",
                    y="Филиал",
                    orientation="h",
                    text="Итоговый балл",
                    color="Итоговый балл",
                    color_continuous_scale=["#ef4444", "#f59e0b", "#10b981"],
                    range_x=[0, 10]
                )
                fig_g.update_layout(
                    plot_bgcolor="white",
                    paper_bgcolor="white",
                    height=520,
                    margin=dict(l=0, r=0, t=10, b=10),
                    coloraxis_showscale=False,
                    font=dict(family="sans-serif", size=12, color="#334155")
                )
                fig_g.update_traces(
                    texttemplate='%{text:.2f}',
                    textposition='outside',
                    marker_line_width=0,
                    opacity=0.9
                )
                st.plotly_chart(fig_g, use_container_width=True)
            else:
                st.info("Нет данных по заданным фильтрам.")
                
        with sub_tg2:
            st.markdown("##### Детализация критериев выбранного филиала")
            if len(df_g_current[df_g_current["Есть данные"]]) > 0:
                sel_branch_crit = st.selectbox(
                    "Детальный филиал:",
                    sorted(df_g_current[df_g_current["Есть данные"]]['Филиал'].unique()),
                    key="crit_sel_box"
                )
                row_crit = df_g_current[df_g_current['Филиал'] == sel_branch_crit].iloc[0]
                
                crit_df_bar = pd.DataFrame({
                    "Критерий": [
                        "Списания в срок",
                        "Комментарии",
                        "Статья затрат",
                        "Самозакуп",
                        "План затрат"
                    ],
                    "Балл": [
                        row_crit['Списания в срок'],
                        row_crit['Комментарии'],
                        row_crit['Статья затрат'],
                        row_crit['Самозакуп'],
                        row_crit['План затрат']
                    ]
                })
                fig_crit = px.bar(
                    crit_df_bar,
                    x="Критерий",
                    y="Балл",
                    text="Балл",
                    color="Балл",
                    range_y=[0, 11],
                    color_continuous_scale=["#f97316", "#fb923c", "#fdba74"]
                )
                fig_crit.update_layout(
                    plot_bgcolor="white",
                    paper_bgcolor="white",
                    height=420,
                    coloraxis_showscale=False,
                    font=dict(family="sans-serif", size=12, color="#334155")
                )
                fig_crit.update_traces(texttemplate='%{text}', textposition='outside')
                st.plotly_chart(fig_crit, use_container_width=True)
            else:
                st.info("Выберите филиал с доступными данными.")

    with tab_g_prog:
        st.markdown("##### 📈 Прогрессия среднего балла сети и филиалов по месяцам (Июль — Сентябрь 2026)")
        all_mega = pd.concat([grades_data[m] for m in clean_sheets], ignore_index=True)
        
        prog_branch = st.selectbox(
            "Выберите филиал для персональной прогрессии:",
            ["Вся сеть (среднее)"] + sorted(all_mega['Филиал'].unique())
        )
        
        if prog_branch == "Вся сеть (среднее)":
            net_prog = all_mega[all_mega["Есть данные"]].groupby('Месяц')['Итоговый балл'].mean().reset_index()
            month_order_map = {'Июль 2026': 1, 'Август 2026': 2, 'Сентябрь 2026': 3}
            net_prog['order'] = net_prog['Месяц'].map(month_order_map)
            net_prog = net_prog.sort_values('order').drop(columns=['order'])
            
            fig_net_prog = px.line(
                net_prog,
                x='Месяц',
                y='Итоговый балл',
                markers=True,
                text='Итоговый балл',
                range_y=[0, 10],
                color_discrete_sequence=["#FF6600"]
            )
            fig_net_prog.update_layout(
                plot_bgcolor="white",
                paper_bgcolor="white",
                height=400,
                font=dict(size=13)
            )
            fig_net_prog.update_traces(
                textposition="top center",
                texttemplate='%{text:.2f}',
                marker=dict(size=12)
            )
            st.plotly_chart(fig_net_prog, use_container_width=True)
        else:
            b_prog = all_mega[all_mega['Филиал'] == prog_branch]
            fig_b_prog = px.bar(
                b_prog,
                x='Месяц',
                y='Итоговый балл',
                text='Итоговый балл',
                range_y=[0, 10],
                color='Итоговый балл',
                color_continuous_scale="Greens"
            )
            fig_b_prog.update_layout(
                plot_bgcolor="white",
                paper_bgcolor="white",
                height=400,
                coloraxis_showscale=False
            )
            fig_b_prog.update_traces(texttemplate='%{text}', textposition='outside')
            st.plotly_chart(fig_b_prog, use_container_width=True)
            
        st.markdown("##### Сводная матрица оценок всех филиалов")
        pivot_g = all_mega.pivot(index='Филиал', columns='Месяц', values='Итоговый балл').fillna(0)
        st.dataframe(pivot_g, use_container_width=True, height=400)

    with tab_g_crit:
        st.markdown("### 📋 Официальные критерии оценивания и шкала штрафов")
        st.markdown("Регламентные требования и нормативы контроля направления «Обеспечение»:")
        st.dataframe(criteria_df, use_container_width=True, height=450)

else:
    st.markdown("## 💰 Дашборд №2: Управление и глубокий анализ затрат")
    
    cost_months = sorted(costs_df['Месяц'].unique())
    all_arts = sorted(costs_df['Статья затрат'].dropna().unique())
    all_cost_branches = sorted(costs_df['Филиал'].dropna().unique())
    
    col_c1, col_c2, col_c3 = st.columns(3)
    with col_c1:
        selected_cost_month = st.selectbox(
            "📅 Выберите месяц затрат:",
            cost_months,
            key="c_m_sel"
        )
    with col_c2:
        selected_arts = st.multiselect(
            "🏷️ Статьи затрат (мультивыбор):",
            all_arts,
            default=all_arts,
            key="c_art_sel"
        )
    with col_c3:
        selected_cost_branch = st.selectbox(
            "🏢 Выберите город / филиал:",
            ["Все филиалы"] + all_cost_branches,
            key="c_b_sel"
        )
        
    df_c_current = costs_df[costs_df['Месяц'] == selected_cost_month].copy()
    if selected_arts:
        df_c_current = df_c_current[df_c_current['Статья затрат'].isin(selected_arts)]
    if selected_cost_branch != "Все филиалы":
        df_c_current = df_c_current[df_c_current['Филиал'] == selected_cost_branch]
        
    total_c = df_c_current['Сумма'].sum()
    mc1, mc2 = st.columns(2)
    with mc1:
        st.metric("Общая сумма расходов по фильтру", f"{total_c:,.0f} ₸".replace(",", " "))
    with mc2:
        st.metric("Количество записей в реестре", len(df_c_current))
        
    st.markdown("---")
    
    tab_c_main, tab_c_prog = st.tabs([
        "📊 Анализ и структура расходов",
        "📈 Прогрессия затрат по месяцам"
    ])
    
    with tab_c_main:
        col_cc1, col_cc2 = st.columns(2)
        with col_cc1:
            st.markdown("##### 📊 Топ филиалов по расходам")
            b_costs_sum = df_c_current.groupby('Филиал')['Сумма'].sum().reset_index().sort_values(by='Сумма', ascending=True)
            if len(b_costs_sum) > 0:
                fig_bc = px.bar(
                    b_costs_sum.tail(15),
                    x='Сумма',
                    y='Филиал',
                    orientation='h',
                    text='Сумма',
                    color='Сумма',
                    color_continuous_scale="Oranges"
                )
                fig_bc.update_layout(
                    plot_bgcolor="white",
                    paper_bgcolor="white",
                    height=500,
                    margin=dict(l=0, r=0, t=10, b=10),
                    coloraxis_showscale=False,
                    font=dict(family="sans-serif", size=12, color="#334155")
                )
                fig_bc.update_traces(
                    texttemplate='%{text:,.0f}',
                    textposition='outside',
                    opacity=0.9
                )
                st.plotly_chart(fig_bc, use_container_width=True)
            else:
                st.info("Нет данных")
                
        with col_cc2:
            st.markdown("##### 🍩 Структура расходов по статьям")
            art_sum = df_c_current.groupby('Статья затрат')['Сумма'].sum().reset_index()
            if len(art_sum) > 0:
                fig_pie = px.pie(
                    art_sum,
                    names='Статья затрат',
                    values='Сумма',
                    hole=0.5,
                    color_discrete_sequence=px.colors.qualitative.Bold
                )
                fig_pie.update_layout(
                    paper_bgcolor="white",
                    height=500,
                    margin=dict(l=0, r=0, t=10, b=10),
                    legend=dict(
                        orientation="h",
                        yanchor="bottom",
                        y=-0.35,
                        xanchor="center",
                        x=0.5
                    ),
                    font=dict(family="sans-serif", size=11, color="#334155")
                )
                st.plotly_chart(fig_pie, use_container_width=True)
            else:
                st.info("Нет данных")

        st.markdown("##### 📋 Детальный реестр транзакций затрат")
        st.dataframe(df_c_current, use_container_width=Thread if False else True, height=400)

    with tab_c_prog:
        st.markdown("##### 📈 Прогрессия и динамика общих затрат сети по месяцам (с учетом выбранных статей затрат)")
        
        # FIX: Filter progression data by selected articles so progression works interactively with filters!
        df_c_prog_filtered = costs_df.copy()
        if selected_arts:
            df_c_prog_filtered = df_c_prog_filtered[df_c_prog_filtered['Статья затрат'].isin(selected_arts)]
        if selected_cost_branch != "Все филиалы":
            df_c_prog_filtered = df_c_prog_filtered[df_c_prog_filtered['Филиал'] == selected_cost_branch]
            
        c_trend = df_c_prog_filtered.groupby('Месяц')['Сумма'].sum().reset_index()
        cost_order_map = {'Июль 2026': 1, 'Август 2026': 2, 'Сентябрь 2026': 3}
        c_trend['order'] = c_trend['Месяц'].map(cost_order_map)
        c_trend = c_trend.sort_values('order').drop(columns=['order'])
        
        if len(c_trend) > 0:
            fig_c_prog = px.bar(
                c_trend,
                x='Месяц',
                y='Сумма',
                text='Сумма',
                color='Сумма',
                color_continuous_scale="Sunsetdark"
            )
            fig_c_prog.update_layout(
                plot_bgcolor="white",
                paper_bgcolor="white",
                height=420,
                coloraxis_showscale=False
            )
            fig_c_prog.update_traces(texttemplate='%{text:,.0f}', textposition='outside')
            st.plotly_chart(fig_c_prog, use_container_width=True)
        else:
            st.info("Нет данных для отображения прогрессии по выбранным фильтрам.")
        
        st.markdown("##### Сводная матрица затрат филиалов по месяцам")
        pivot_c = df_c_prog_filtered.pivot_table(index='Филиал', columns='Месяц', values='Сумма', aggfunc='sum').fillna(0)
        st.dataframe(pivot_c, use_container_width=True, height=400)
