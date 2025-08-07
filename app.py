import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, timedelta
import numpy as np
from io import BytesIO
import base64
from reportlab.lib.pagesizes import letter, A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib import colors
import tempfile
import os
from utils import OrderAnalyzer, generate_mock_data

# Настройка страницы
st.set_page_config(
    page_title="Order Delivery Analyzer",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Заголовок приложения
st.title("📦 Order Delivery Analyzer")
st.markdown("### Инструмент для анализа доставки заказов")

# Боковая панель
st.sidebar.header("⚙️ Настройки")

# Контактная информация
st.sidebar.markdown("---")
st.sidebar.markdown("📞 **Контакты**")
st.sidebar.markdown("""
<a href="https://t.me/alkash_slayer" target="_blank">
    <svg width="20" height="20" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
        <path d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm4.64 6.8c-.15 1.58-.8 5.42-1.13 7.19-.14.75-.42 1-.68 1.03-.58.05-1.02-.38-1.58-.75-.88-.58-1.38-.94-2.23-1.5-.99-.65-.35-1.01.22-1.59.15-.15 2.71-2.48 2.76-2.69a.2.2 0 00-.05-.18c-.06-.05-.14-.03-.21-.02-.09.02-1.49.95-4.22 2.79-.4.27-.76.41-1.08.4-.36-.01-1.04-.2-1.55-.37-.63-.2-1.12-.31-1.08-.66.02-.18.27-.36.74-.55 2.92-1.27 4.86-2.11 5.83-2.51 2.78-1.16 3.35-1.36 3.73-1.36.08 0 .27.02.39.12.1.08.13.19.14.27-.01.06.01.24 0 .38z" fill="#0088cc"/>
    </svg>
    Telegram
</a>
""", unsafe_allow_html=True)
st.sidebar.markdown("---")

# Выбор источника данных
data_source = st.sidebar.radio(
    "Источник данных:",
    ["Загрузить CSV файл", "Использовать тестовые данные"]
)

# Инициализация данных
df = None

if data_source == "Загрузить CSV файл":
    uploaded_file = st.sidebar.file_uploader(
        "Выберите CSV файл",
        type=['csv'],
        help="Файл должен содержать поля: Номер заказа, Номер отправления, Принят в обработку, Дата отгрузки, Статус, Дата доставки, Фактическая дата передачи в доставку, Сумма отправления, Наименование товара, OZON id, Количество"
    )
    
    if uploaded_file is not None:
        try:
            df = pd.read_csv(uploaded_file, sep=';', encoding='utf-8')
            st.sidebar.success(f"✅ Файл загружен: {len(df)} записей")
        except Exception as e:
            st.sidebar.error(f"❌ Ошибка загрузки файла: {str(e)}")
else:
    # Параметры для генерации тестовых данных
    st.sidebar.subheader("Параметры генерации данных")
    n_orders = st.sidebar.slider("Количество заказов", 50, 500, 100, 25)
    
    if st.sidebar.button("🎲 Сгенерировать данные"):
        with st.spinner("Генерация данных..."):
            df = generate_mock_data(n_orders)
            st.sidebar.success(f"✅ Данные сгенерированы: {len(df)} записей")

# Основной интерфейс
if df is not None:
    # Проверка структуры данных
    required_columns = [
        'Номер заказа', 'Номер отправления', 'Принят в обработку', 'Дата отгрузки',
        'Статус', 'Дата доставки', 'Фактическая дата передачи в доставку',
        'Сумма отправления', 'Наименование товара', 'OZON id', 'Количество'
    ]
    
    missing_columns = [col for col in required_columns if col not in df.columns]
    
    if missing_columns:
        st.error(f"❌ Отсутствуют обязательные поля: {', '.join(missing_columns)}")
        st.stop()
    
    # Создание анализатора
    analyzer = OrderAnalyzer(df)
    
    # Вкладки
    tab1, tab2, tab3, tab4 = st.tabs(["📊 Обзор данных", "📈 Анализ доставки", "🔍 Фильтры", "📄 Отчет"])
    
    with tab1:
        st.header("📊 Обзор данных")
        
        # Основные метрики
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            st.metric("Всего заказов", len(df))
        
        with col2:
            delivered = len(df[df['Статус'] == 'Доставлен'])
            st.metric("Доставлено", delivered)
        
        with col3:
            avg_sum = analyzer.df['Сумма отправления'].mean()
            st.metric("Средняя сумма", f"{avg_sum:.2f} ₽")
        
        with col4:
            total_sum = analyzer.df['Сумма отправления'].sum()
            st.metric("Общая сумма", f"{total_sum:.2f} ₽")
        
        # Распределение статусов
        col1, col2 = st.columns(2)
        
        with col1:
            st.subheader("Распределение статусов")
            status_dist = df['Статус'].value_counts()
            fig_status = px.pie(values=status_dist.values, names=status_dist.index, 
                               title="Статусы заказов")
            st.plotly_chart(fig_status, use_container_width=True)
        
        with col2:
            st.subheader("Топ-10 товаров по количеству")
            top_products = analyzer.df.groupby('Наименование товара')['Количество'].sum().sort_values(ascending=False).head(10)
            fig_products = px.bar(x=top_products.values, y=top_products.index, 
                                orientation='h', title="Топ товары")
            st.plotly_chart(fig_products, use_container_width=True)
        
        # Таблица с данными
        st.subheader("Просмотр данных")
        st.dataframe(df.head(50), use_container_width=True)
    
    with tab2:
        st.header("📈 Анализ доставки")
        
        # Временные метрики
        metrics = analyzer.calculate_delivery_metrics()
        
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            st.metric(
                "Среднее время доставки",
                f"{metrics['avg_delivery_time']:.1f} дней"
            )
        
        with col2:
            st.metric(
                "Медианное время доставки",
                f"{metrics['median_delivery_time']:.1f} дней"
            )
        
        with col3:
            st.metric(
                "Время до передачи в доставку",
                f"{metrics['avg_processing_time']:.1f} дней"
            )
        
        with col4:
            st.metric(
                "Время в доставке",
                f"{metrics['avg_shipping_time']:.1f} дней"
            )
        
        # Графики временных рядов
        col1, col2 = st.columns(2)
        
        with col1:
            st.subheader("Заказы по дням")
            daily_orders = analyzer.get_daily_orders()
            fig_daily = px.line(daily_orders, x='date', y='count', 
                               title="Количество заказов по дням")
            st.plotly_chart(fig_daily, use_container_width=True)
        
        with col2:
            st.subheader("Топ-10 товаров по сумме")
            top_revenue = analyzer.df.groupby('Наименование товара')['Сумма отправления'].sum().sort_values(ascending=False).head(10)
            fig_revenue = px.bar(x=top_revenue.values, y=top_revenue.index, 
                               orientation='h', title="Топ товары по выручке")
            st.plotly_chart(fig_revenue, use_container_width=True)
        
        # График задержек
        st.subheader("Анализ задержек доставки")
        delays = analyzer.get_delivery_delays()
        if len(delays) > 0:
            fig_delays = px.histogram(delays, x='delay_days', 
                                    title="Распределение задержек (дни)")
            st.plotly_chart(fig_delays, use_container_width=True)
        else:
            st.info("Задержек доставки не обнаружено")
    
    with tab3:
        st.header("🔍 Фильтры и детальный анализ")
        
        # Фильтры
        col1, col2, col3 = st.columns(3)
        
        with col1:
            # Фильтр по дате
            # Используем обработанные данные из analyzer.df вместо исходного df
            valid_dates = analyzer.df['Принят в обработку'].dropna()
            if len(valid_dates) > 0:
                date_range = st.date_input(
                    "Период",
                    value=(valid_dates.min().date(), 
                          valid_dates.max().date()),
                    min_value=valid_dates.min().date(),
                    max_value=valid_dates.max().date()
                )
            else:
                st.error("Нет корректных дат в столбце 'Принят в обработку'")
                date_range = []
        
        with col2:
            # Фильтр по статусу
            selected_statuses = st.multiselect(
                "Статусы",
                options=analyzer.df['Статус'].unique(),
                default=analyzer.df['Статус'].unique()
            )
        
        with col3:
            # Фильтр по товарам
            selected_products = st.multiselect(
                "Товары",
                options=analyzer.df['Наименование товара'].unique(),
                default=analyzer.df['Наименование товара'].unique()[:10]  # Первые 10 для удобства
            )
        
        # Применение фильтров
        if len(date_range) == 2:
            start_date, end_date = date_range
            # Фильтруем только строки с корректными датами
            filtered_df = analyzer.df[
                (analyzer.df['Принят в обработку'].notna()) &
                (analyzer.df['Принят в обработку'].dt.date >= start_date) &
                (analyzer.df['Принят в обработку'].dt.date <= end_date) &
                (analyzer.df['Статус'].isin(selected_statuses)) &
                (analyzer.df['Наименование товара'].isin(selected_products))
            ]
        else:
            filtered_df = analyzer.df[
                (analyzer.df['Статус'].isin(selected_statuses)) &
                (analyzer.df['Наименование товара'].isin(selected_products))
            ]
        
        if len(filtered_df) == 0:
            st.warning("⚠️ Нет данных для выбранных фильтров")
        else:
            st.success(f"✅ Найдено {len(filtered_df)} заказов")
            
            # Метрики для отфильтрованных данных
            filtered_analyzer = OrderAnalyzer(filtered_df)
            filtered_metrics = filtered_analyzer.calculate_delivery_metrics()
            
            col1, col2, col3 = st.columns(3)
            
            with col1:
                st.metric(
                    "Среднее время доставки (фильтр)",
                    f"{filtered_metrics['avg_delivery_time']:.1f} дней"
                )
            
            with col2:
                filtered_analyzer = OrderAnalyzer(filtered_df)
                avg_sum_filtered = filtered_analyzer.df['Сумма отправления'].mean()
                st.metric(
                    "Средняя сумма (фильтр)",
                    f"{avg_sum_filtered:.2f} ₽"
                )
            
            with col3:
                delivered_filtered = len(filtered_df[filtered_df['Статус'] == 'Доставлен'])
                delivery_rate = (delivered_filtered / len(filtered_df)) * 100
                st.metric(
                    "Процент доставки",
                    f"{delivery_rate:.1f}%"
                )
            
            # Детальная таблица
            st.subheader("Отфильтрованные данные")
            st.dataframe(filtered_df, use_container_width=True)
    
    with tab4:
        st.header("📄 Генерация отчета")
        
        st.markdown("""
        Этот раздел позволяет сгенерировать Excel-отчет с основными метриками и аналитикой.
        Excel формат полностью поддерживает кириллицу и позволяет удобно работать с данными.
        """)
        
        if st.button("📊 Сгенерировать Excel отчет"):
            with st.spinner("Генерация отчета..."):
                try:
                    # Генерация Excel отчета
                    excel_buffer = analyzer.generate_excel_report()
                    
                    # Создание ссылки для скачивания
                    b64_excel = base64.b64encode(excel_buffer.getvalue()).decode()
                    href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64_excel}" download="order_delivery_report.xlsx">📥 Скачать Excel отчет</a>'
                    st.markdown(href, unsafe_allow_html=True)
                    
                    st.success("✅ Excel отчет успешно сгенерирован!")
                    st.info("💡 Excel отчет содержит 5 листов: Основные метрики, Топ товары (количество), Топ товары (выручка), Статусы заказов, Анализ задержек")
                    
                except Exception as e:
                    st.error(f"❌ Ошибка генерации отчета: {str(e)}")
        

        
        # Предварительный просмотр метрик для отчета
        st.subheader("Предварительный просмотр метрик")
        
        metrics = analyzer.calculate_delivery_metrics()
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.write("**Временные метрики:**")
            st.write(f"• Среднее время доставки: {metrics['avg_delivery_time']:.1f} дней")
            st.write(f"• Медианное время доставки: {metrics['median_delivery_time']:.1f} дней")
            st.write(f"• Время обработки: {metrics['avg_processing_time']:.1f} дней")
            st.write(f"• Время в доставке: {metrics['avg_shipping_time']:.1f} дней")
        
        with col2:
            st.write("**Общие метрики:**")
            st.write(f"• Всего заказов: {len(df)}")
            st.write(f"• Доставлено: {len(df[df['Статус'] == 'Доставлен'])}")
            st.write(f"• Средняя сумма: {analyzer.df['Сумма отправления'].mean():.2f} ₽")
            st.write(f"• Общая выручка: {analyzer.df['Сумма отправления'].sum():.2f} ₽")

else:
    st.info("👆 Загрузите CSV файл или сгенерируйте тестовые данные для начала анализа")
    
    # Показать пример структуры данных
    st.subheader("📋 Ожидаемая структура CSV файла")
    
    example_data = {
        'Номер заказа': ['ORD001', 'ORD002'],
        'Номер отправления': ['SHIP001', 'SHIP002'],
        'Принят в обработку': ['2024-01-01 10:00:00', '2024-01-01 11:00:00'],
        'Дата отгрузки': ['2024-01-02 09:00:00', '2024-01-02 10:00:00'],
        'Статус': ['Доставлен', 'В пути'],
        'Дата доставки': ['2024-01-05 14:00:00', ''],
        'Фактическая дата передачи в доставку': ['2024-01-02 15:00:00', '2024-01-02 16:00:00'],
        'Сумма отправления': [1500.00, 2300.50],
        'Наименование товара': ['Товар А', 'Товар Б'],
        'OZON id': ['OZ123', 'OZ124'],
        'Количество': [2, 1]
    }
    
    example_df = pd.DataFrame(example_data)
    st.dataframe(example_df, use_container_width=True)
    
    st.markdown("""
    **Требования к файлу:**
    - Разделитель: `;` (точка с запятой)
    - Кодировка: UTF-8
    - Формат дат: YYYY-MM-DD HH:MM:SS
    """)