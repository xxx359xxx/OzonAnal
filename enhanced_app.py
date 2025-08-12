import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import base64
from datetime import datetime, timedelta
import numpy as np
from enhanced_analyzer import EnhancedOrderAnalyzer
from utils import generate_mock_data, create_sample_csv

# Настройка страницы
st.set_page_config(
    page_title="OzonStream Enhanced Analytics",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Боковая панель
with st.sidebar:
    st.markdown("### 📊 OzonStream Enhanced")
    st.markdown("**Расширенная аналитика заказов Ozon**")
    st.markdown("---")
    st.markdown("💬 [Связаться с разработчиком](https://t.me/alkash_slayer)")
    st.markdown("---")
    
    # Выбор источника данных
    data_source = st.radio(
        "Источник данных:",
        ["📁 Загрузить CSV файл", "🎲 Сгенерировать тестовые данные"]
    )

# Заголовок
st.title("📊 OzonStream Enhanced Analytics")
st.markdown("**Расширенная система анализа заказов Ozon с поддержкой 25 полей**")
st.markdown("---")

# Загрузка данных
df = None

if data_source == "📁 Загрузить CSV файл":
    uploaded_file = st.file_uploader(
        "Выберите CSV файл с заказами Ozon",
        type=['csv'],
        help="Поддерживается новый формат CSV с 25 колонками"
    )
    
    if uploaded_file is not None:
        try:
            # Читаем CSV с правильными параметрами для нового формата
            df = pd.read_csv(
                uploaded_file,
                sep=';',
                encoding='utf-8',
                quotechar='"'
            )
            
            st.success(f"✅ Файл успешно загружен! Найдено {len(df)} записей")
            
            # Проверяем наличие ключевых колонок
            required_columns = [
                'Номер заказа', 'Принят в обработку', 'Статус',
                'Наименование товара', 'Сумма отправления'
            ]
            
            missing_columns = [col for col in required_columns if col not in df.columns]
            
            if missing_columns:
                st.error(f"❌ Отсутствуют обязательные колонки: {', '.join(missing_columns)}")
                st.info("📋 Доступные колонки в файле:")
                st.write(list(df.columns))
                df = None
            else:
                st.info(f"📊 Обнаружено {len(df.columns)} колонок в файле")
                
        except Exception as e:
            st.error(f"❌ Ошибка при чтении файла: {str(e)}")
            st.info("💡 Убедитесь, что файл имеет правильный формат (разделитель ';', кодировка UTF-8)")

else:  # Генерация тестовых данных
    if st.button("🎲 Сгенерировать тестовые данные", type="primary"):
        with st.spinner("Генерация тестовых данных..."):
            # Адаптируем генератор под новый формат
            mock_data = generate_mock_data(200)  # Генерируем больше данных
            
            # Преобразуем в новый формат
            df_new_format = []
            for i, order in enumerate(mock_data):
                df_new_format.append({
                    'Номер заказа': f"{order['order_id']}-{i:04d}",
                    'Номер отправления': f"{order['order_id']}-{i:04d}-1",
                    'Принят в обработку': order['order_date'].strftime('%Y-%m-%d %H:%M:%S'),
                    'Дата отгрузки': (order['order_date'] + timedelta(hours=12)).strftime('%Y-%m-%d %H:%M:%S'),
                    'Статус': order['status'],
                    'Дата доставки': order['delivery_date'].strftime('%Y-%m-%d %H:%M:%S') if order['delivery_date'] else '',
                    'Фактическая дата передачи в доставку': (order['order_date'] + timedelta(hours=8)).strftime('%Y-%m-%d %H:%M:%S'),
                    'Сумма отправления': order['total_amount'],
                    'Код валюты отправления': 'RUB',
                    'Наименование товара': order['product_name'],
                    'OZON id': f"176{np.random.randint(1000000, 9999999)}",
                    'Артикул': f"VMS-{np.random.choice(['TS', 'SW', 'SC', 'TY'])}-{np.random.randint(100, 999)}-{np.random.randint(100, 999)}",
                    'Ваша цена': order['total_amount'],
                    'Код валюты товара': 'RUB',
                    'Оплачено покупателем': order['total_amount'] * 0.7,
                    'Код валюты покупателя': 'RUB',
                    'Количество': order['quantity'],
                    'Стоимость доставки': 0,
                    'Связанные отправления': '',
                    'Выкуп товара': '',
                    'Цена товара до скидок': order['total_amount'] * 1.5,
                    'Скидка %': f"{np.random.randint(10, 80)}%",
                    'Скидка руб': order['total_amount'] * 0.3,
                    'Акции': 'Системная виртуальная скидка селлера Россия (RUB)',
                    'Объемный вес товаров, кг': np.random.uniform(0.1, 3.0)
                })
            
            df = pd.DataFrame(df_new_format)
            st.success(f"✅ Сгенерировано {len(df)} тестовых записей")

# Основной анализ
if df is not None:
    try:
        # Создаем экземпляр расширенного анализатора
        analyzer = EnhancedOrderAnalyzer(df)
        
        # Вкладки для разных видов анализа
        tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
            "📊 Обзор данных", 
            "💰 Финансовый анализ", 
            "🎯 Анализ скидок", 
            "🚚 Анализ доставки", 
            "📦 Логистика", 
            "📅 Месячный анализ",
            "📈 Отчеты"
        ])
        
        with tab1:
            st.header("📊 Обзор данных")
            
            # Основные метрики
            metrics = analyzer.get_basic_metrics()
            
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("Всего заказов", metrics['total_orders'])
            with col2:
                st.metric("Доставлено", metrics['delivered_orders'], 
                         f"{metrics['delivery_rate']:.1f}%")
            with col3:
                st.metric("Отменено", metrics['cancelled_orders'], 
                         f"{metrics['cancellation_rate']:.1f}%")
            with col4:
                st.metric("В доставке", metrics['in_delivery'])
            
            # Распределение статусов
            col1, col2 = st.columns(2)
            
            with col1:
                status_counts = df['Статус'].value_counts()
                fig_status = px.pie(
                    values=status_counts.values,
                    names=status_counts.index,
                    title="Распределение статусов заказов",
                    color_discrete_sequence=px.colors.qualitative.Set3
                )
                st.plotly_chart(fig_status, use_container_width=True)
            
            with col2:
                # Временной ряд заказов
                time_series = analyzer.get_time_series_analysis()
                if isinstance(time_series, pd.DataFrame) and not time_series.empty:
                    fig_time = px.line(
                        time_series,
                        x='date',
                        y='orders_count',
                        title="Динамика заказов по дням",
                        markers=True
                    )
                    fig_time.update_layout(xaxis_title="Дата", yaxis_title="Количество заказов")
                    st.plotly_chart(fig_time, use_container_width=True)
            
            # Топ товаров
            st.subheader("🏆 Топ-10 товаров по выручке")
            product_stats = analyzer.analyze_product_categories().head(10)
            
            fig_products = px.bar(
                x=product_stats['total_revenue'],
                y=[name[:40] + '...' if len(name) > 40 else name for name in product_stats.index],
                orientation='h',
                title="Топ товаров по выручке",
                labels={'x': 'Выручка (₽)', 'y': 'Товар'}
            )
            fig_products.update_layout(height=500)
            st.plotly_chart(fig_products, use_container_width=True)
        
        with tab2:
            st.header("💰 Финансовый анализ")
            
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Общая выручка", f"{metrics['total_revenue']:,.2f} ₽")
            with col2:
                st.metric("Средний чек", f"{metrics['avg_order_value']:,.2f} ₽")
            with col3:
                st.metric("Общая сумма скидок", f"{metrics['total_discount_amount']:,.2f} ₽")
            
            # Анализ по регионам (если есть данные)
            regional_stats = analyzer.get_regional_analysis()
            if isinstance(regional_stats, pd.DataFrame) and not regional_stats.empty:
                st.subheader("🗺️ Анализ по регионам")
                
                col1, col2 = st.columns(2)
                with col1:
                    fig_region_orders = px.bar(
                        x=regional_stats.index,
                        y=regional_stats['order_count'],
                        title="Количество заказов по регионам",
                        labels={'x': 'Регион', 'y': 'Количество заказов'}
                    )
                    st.plotly_chart(fig_region_orders, use_container_width=True)
                
                with col2:
                    fig_region_revenue = px.bar(
                        x=regional_stats.index,
                        y=regional_stats['total_revenue'],
                        title="Выручка по регионам",
                        labels={'x': 'Регион', 'y': 'Выручка (₽)'}
                    )
                    st.plotly_chart(fig_region_revenue, use_container_width=True)
            
            # Временной анализ выручки
            time_series = analyzer.get_time_series_analysis()
            if isinstance(time_series, pd.DataFrame) and not time_series.empty:
                st.subheader("📈 Динамика выручки")
                
                fig_revenue_time = px.line(
                    time_series,
                    x='date',
                    y='daily_revenue',
                    title="Ежедневная выручка",
                    markers=True
                )
                fig_revenue_time.update_layout(xaxis_title="Дата", yaxis_title="Выручка (₽)")
                st.plotly_chart(fig_revenue_time, use_container_width=True)
        
        with tab3:
            st.header("🎯 Анализ скидок и акций")
            
            discount_analysis = analyzer.analyze_discounts()
            
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("Заказов со скидкой", discount_analysis['orders_with_discount'])
            with col2:
                st.metric("% заказов со скидкой", f"{discount_analysis['discount_percentage']:.1f}%")
            with col3:
                st.metric("Средняя скидка", f"{discount_analysis['avg_discount_amount']:.2f} ₽")
            with col4:
                st.metric("Общая экономия", f"{discount_analysis['total_savings']:,.2f} ₽")
            
            col1, col2 = st.columns(2)
            
            with col1:
                # Распределение размеров скидок
                if 'Скидка руб' in df.columns:
                    discount_data = df[df['Скидка руб'] > 0]['Скидка руб']
                    if len(discount_data) > 0:
                        fig_discount_dist = px.histogram(
                            x=discount_data,
                            title="Распределение размеров скидок",
                            labels={'x': 'Размер скидки (₽)', 'y': 'Количество заказов'},
                            nbins=20
                        )
                        st.plotly_chart(fig_discount_dist, use_container_width=True)
            
            with col2:
                # Анализ акций
                if 'Акции' in df.columns:
                    promo_counts = df['Акции'].value_counts().head(10)
                    if len(promo_counts) > 0:
                        fig_promos = px.bar(
                            x=promo_counts.values,
                            y=[name[:30] + '...' if len(name) > 30 else name for name in promo_counts.index],
                            orientation='h',
                            title="Топ акций по количеству заказов",
                            labels={'x': 'Количество заказов', 'y': 'Акция'}
                        )
                        st.plotly_chart(fig_promos, use_container_width=True)
        
        with tab4:
            st.header("🚚 Анализ доставки")
            
            delivery_analysis = analyzer.analyze_delivery_performance()
            
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("Среднее время доставки", f"{delivery_analysis['avg_delivery_time']:.1f} дней")
            with col2:
                st.metric("Медианное время", f"{delivery_analysis['median_delivery_time']:.1f} дней")
            with col3:
                st.metric("Доставка вовремя", f"{delivery_analysis['on_time_delivery_rate']:.1f}%")
            with col4:
                st.metric("Всего доставлено", delivery_analysis['total_delivered'])
            
            # Анализ времени доставки
            delivered_orders = df[df['Статус'] == 'Доставлен'].copy()
            if len(delivered_orders) > 0 and 'Дата доставки' in df.columns and 'Принят в обработку' in df.columns:
                # Конвертируем даты
                delivered_orders['Принят в обработку'] = pd.to_datetime(delivered_orders['Принят в обработку'], errors='coerce')
                delivered_orders['Дата доставки'] = pd.to_datetime(delivered_orders['Дата доставки'], errors='coerce')
                
                # Вычисляем время доставки
                delivered_orders['delivery_time'] = (
                    delivered_orders['Дата доставки'] - delivered_orders['Принят в обработку']
                ).dt.days
                
                # Фильтруем корректные значения
                valid_delivery_times = delivered_orders[
                    (delivered_orders['delivery_time'] >= 0) & 
                    (delivered_orders['delivery_time'] <= 30)
                ]['delivery_time']
                
                if len(valid_delivery_times) > 0:
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        fig_delivery_hist = px.histogram(
                            x=valid_delivery_times,
                            title="Распределение времени доставки",
                            labels={'x': 'Время доставки (дни)', 'y': 'Количество заказов'},
                            nbins=15
                        )
                        st.plotly_chart(fig_delivery_hist, use_container_width=True)
                    
                    with col2:
                        # Категории времени доставки
                        delivery_categories = pd.cut(
                            valid_delivery_times,
                            bins=[0, 1, 3, 5, 10, 30],
                            labels=['1 день', '2-3 дня', '4-5 дней', '6-10 дней', '11+ дней']
                        ).value_counts()
                        
                        fig_delivery_cat = px.pie(
                            values=delivery_categories.values,
                            names=delivery_categories.index,
                            title="Категории времени доставки"
                        )
                        st.plotly_chart(fig_delivery_cat, use_container_width=True)
        
        with tab5:
            st.header("📦 Логистический анализ")
            
            weight_analysis = analyzer.analyze_weight_logistics()
            
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("Общий вес", f"{weight_analysis['total_weight_kg']:.2f} кг")
            with col2:
                st.metric("Средний вес заказа", f"{weight_analysis['avg_weight_per_order']:.0f} г")
            with col3:
                st.metric("Тяжелых заказов", f"{weight_analysis['heavy_orders_count']} ({weight_analysis['heavy_orders_percentage']:.1f}%)")
            with col4:
                st.metric("Легких заказов", f"{weight_analysis['light_orders_count']} ({weight_analysis['light_orders_percentage']:.1f}%)")
            
            # Анализ веса товаров
            if 'Объемный вес товаров, кг' in df.columns:
                weight_data = df['Объемный вес товаров, кг'].dropna()
                if len(weight_data) > 0:
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        fig_weight_hist = px.histogram(
                            x=weight_data,
                            title="Распределение веса товаров",
                            labels={'x': 'Вес (кг)', 'y': 'Количество заказов'},
                            nbins=20
                        )
                        st.plotly_chart(fig_weight_hist, use_container_width=True)
                    
                    with col2:
                        # Категории веса
                        weight_categories = pd.cut(
                            weight_data,
                            bins=[0, 0.5, 1.0, 2.0, 5.0, float('inf')],
                            labels=['≤0.5кг', '0.5-1кг', '1-2кг', '2-5кг', '>5кг']
                        ).value_counts()
                        
                        fig_weight_cat = px.pie(
                            values=weight_categories.values,
                            names=weight_categories.index,
                            title="Категории веса товаров"
                        )
                        st.plotly_chart(fig_weight_cat, use_container_width=True)
        
        with tab6:
            st.header("📅 Месячный анализ успешности")
            
            # Получаем месячный анализ
            monthly_analysis = analyzer.get_monthly_analysis()
            
            if not monthly_analysis.empty:
                st.subheader("📊 Ключевые метрики по месяцам")
                
                # Топ-3 самых успешных месяца
                top_months = monthly_analysis.nlargest(3, 'success_rating')
                
                col1, col2, col3 = st.columns(3)
                for i, (idx, month_data) in enumerate(top_months.iterrows()):
                    with [col1, col2, col3][i]:
                        st.metric(
                            f"🏆 #{i+1} {month_data['month']}",
                            f"{month_data['success_rating']:.1f} баллов",
                            f"{month_data['total_revenue']:,.0f} ₽"
                        )
                
                st.markdown("---")
                
                # Графики анализа
                col1, col2 = st.columns(2)
                
                with col1:
                    # График рейтинга успешности по месяцам (сортируем по рейтингу)
                    monthly_analysis_rating = monthly_analysis.sort_values('success_rating', ascending=False)
                    fig_rating = px.bar(
                        monthly_analysis_rating,
                        x='month',
                        y='success_rating',
                        title="Рейтинг успешности по месяцам",
                        labels={'month': 'Месяц', 'success_rating': 'Рейтинг успешности'},
                        color='success_rating',
                        color_continuous_scale='RdYlGn'
                    )
                    fig_rating.update_layout(xaxis_tickangle=-45)
                    st.plotly_chart(fig_rating, use_container_width=True)
                
                with col2:
                    # График выручки по месяцам (данные уже отсортированы по дате)
                    fig_revenue = px.line(
                        monthly_analysis,
                        x='month',
                        y='total_revenue',
                        title="Динамика выручки по месяцам",
                        labels={'month': 'Месяц', 'total_revenue': 'Выручка (₽)'},
                        markers=True
                    )
                    fig_revenue.update_layout(xaxis_tickangle=-45)
                    st.plotly_chart(fig_revenue, use_container_width=True)
                
                # Детальная таблица
                st.subheader("📋 Детальная статистика по месяцам")
                
                # Форматируем данные для отображения
                display_data = monthly_analysis.copy()
                display_data['total_revenue'] = display_data['total_revenue'].apply(lambda x: f"{x:,.0f} ₽")
                display_data['avg_order_value'] = display_data['avg_order_value'].apply(lambda x: f"{x:,.0f} ₽")
                display_data['total_discount'] = display_data['total_discount'].apply(lambda x: f"{x:,.0f} ₽")
                display_data['avg_weight_kg'] = display_data['avg_weight_kg'].apply(lambda x: f"{x*1000:.0f} г" if pd.notna(x) else "N/A")
                display_data['total_weight_kg'] = display_data['total_weight_kg'].apply(lambda x: f"{x:.1f} кг" if pd.notna(x) else "N/A")
                display_data['revenue_per_order'] = display_data['revenue_per_order'].apply(lambda x: f"{x:,.0f} ₽")
                display_data['discount_rate'] = display_data['discount_rate'].apply(lambda x: f"{x:.1f}%")
                display_data['success_rating'] = display_data['success_rating'].apply(lambda x: f"{x:.1f}")
                
                # Переименовываем колонки для отображения
                display_data.columns = [
                    'Месяц-год', 'Заказы', 'Выручка', 'Средний чек', 'Скидки', 
                    'Товары', 'Средний вес', 'Общий вес', 'Выручка/заказ', 
                    'Скидка %', 'Товары/заказ', 'Рейтинг успешности', 'Месяц'
                ]
                
                st.dataframe(display_data, use_container_width=True)
                
                # Пояснение к рейтингу
                st.info("""
                **📊 Методика расчета рейтинга успешности:**
                - 40% - выручка (нормализованная)
                - 30% - объем заказов × средний чек (нормализованные)
                - 20% - эффективность скидок (обратная зависимость от % скидок)
                - 10% - товарооборот (количество товаров)
                
                Максимальный рейтинг: 100 баллов
                """)
            else:
                st.warning("⚠️ Недостаточно данных для месячного анализа")
        
        with tab7:
            st.header("📈 Отчеты и инсайты")
            
            # Ключевые инсайты
            st.subheader("🔍 Ключевые инсайты")
            insights = analyzer.get_summary_insights()
            for insight in insights:
                st.info(insight)
            
            st.markdown("---")
            
            # Генерация Excel отчета
            st.subheader("📊 Скачать расширенный Excel отчет")
            
            if st.button("🔄 Сгенерировать отчет", type="primary"):
                with st.spinner("Генерация отчета..."):
                    try:
                        excel_buffer = analyzer.generate_enhanced_excel_report()
                        
                        # Кодируем в base64 для скачивания
                        excel_data = excel_buffer.getvalue()
                        b64 = base64.b64encode(excel_data).decode()
                        
                        # Создаем ссылку для скачивания
                        filename = f"ozon_enhanced_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                        href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="{filename}">📥 Скачать Excel отчет</a>'
                        
                        st.markdown(href, unsafe_allow_html=True)
                        st.success("✅ Отчет успешно сгенерирован!")
                        
                        # Показываем содержимое отчета
                        st.info("""
                        📋 **Содержимое отчета:**
                        - **Основные метрики**: общая статистика заказов
                        - **Анализ скидок**: детальная информация о скидках и акциях
                        - **Анализ товаров**: топ-20 товаров по выручке
                        - **Анализ доставки**: производительность доставки
                        - **Логистика по весу**: анализ веса товаров
                        """)
                        
                    except Exception as e:
                        st.error(f"❌ Ошибка при генерации отчета: {str(e)}")
            
            # Статистика по данным
            st.subheader("📊 Статистика по данным")
            col1, col2 = st.columns(2)
            
            with col1:
                st.write("**Основная информация:**")
                st.write(f"- Всего записей: {len(df)}")
                st.write(f"- Колонок в файле: {len(df.columns)}")
                st.write(f"- Период данных: {df['Принят в обработку'].min()} - {df['Принят в обработку'].max()}")
            
            with col2:
                st.write("**Качество данных:**")
                missing_data = df.isnull().sum().sum()
                total_cells = len(df) * len(df.columns)
                completeness = ((total_cells - missing_data) / total_cells * 100)
                st.write(f"- Полнота данных: {completeness:.1f}%")
                st.write(f"- Пустых значений: {missing_data}")
                st.write(f"- Уникальных заказов: {df['Номер заказа'].nunique()}")
    
    except Exception as e:
        st.error(f"❌ Ошибка при анализе данных: {str(e)}")
        st.info("💡 Проверьте формат загруженного файла")

else:
    # Инструкции по использованию
    st.info("""
    ### 🚀 Добро пожаловать в OzonStream Enhanced Analytics!
    
    **Новые возможности:**
    - ✅ Поддержка нового формата CSV с 25 колонками
    - 📊 Расширенный анализ скидок и акций
    - 🚚 Детальная аналитика доставки
    - 📦 Логистический анализ по весу товаров
    - 🗺️ Региональная аналитика
    - 📈 Временные тренды и инсайты
    
    **Для начала работы:**
    1. Загрузите CSV файл с заказами Ozon или сгенерируйте тестовые данные
    2. Изучите аналитику во вкладках
    3. Скачайте детальный Excel отчет
    
    **Формат файла:**
    - Разделитель: точка с запятой (;)
    - Кодировка: UTF-8
    - Поддерживается новый формат с 25 колонками
    """)
    
    # Пример структуры нового CSV
    st.subheader("📋 Пример структуры нового CSV файла")
    example_columns = [
        "Номер заказа", "Номер отправления", "Принят в обработку", "Дата отгрузки",
        "Статус", "Дата доставки", "Сумма отправления", "Наименование товара",
        "Ваша цена", "Количество", "Цена товара до скидок", "Скидка %",
        "Скидка руб", "Акции", "Объемный вес товаров, кг", "и другие..."
    ]
    
    st.code(";".join(example_columns[:8]) + ";...")
    st.caption("Всего поддерживается 25 колонок с детальной информацией о заказах")