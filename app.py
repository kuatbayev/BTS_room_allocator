import os
from io import BytesIO
from pathlib import Path

import pandas as pd
import streamlit as st

from allocator import generate_outputs, read_excel_from_upload


st.set_page_config(page_title='BTS Room Allocator', page_icon='🏫', layout='wide')

st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Manrope:wght@400;700;800&display=swap');
    .stApp {
        font-family: 'Manrope', sans-serif;
        background: radial-gradient(1200px 500px at 10% -10%, #ffe8b6 0%, rgba(255,232,182,0) 60%),
                    radial-gradient(900px 500px at 100% 0%, #d8f3ff 0%, rgba(216,243,255,0) 55%),
                    #f6f7fb;
    }
    .hero {
        padding: 1.2rem 1.4rem;
        border-radius: 18px;
        background: linear-gradient(125deg, #0f1b40 0%, #1e3a8a 45%, #0ea5a4 100%);
        color: white;
        box-shadow: 0 12px 30px rgba(16, 24, 40, 0.18);
        margin-bottom: 1rem;
    }
    .hero h1 {
        margin: 0;
        font-size: 2rem;
        font-weight: 800;
        letter-spacing: 0.4px;
    }
    .hero p {
        margin-top: 0.45rem;
        margin-bottom: 0;
        opacity: 0.92;
        font-size: 1rem;
    }
    .stat {
        padding: 0.8rem 1rem;
        border-radius: 14px;
        background: #ffffff;
        border: 1px solid #e5e7eb;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


def df_to_excel_bytes(df: pd.DataFrame) -> bytes:
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    buffer.seek(0)
    return buffer.getvalue()


rooms_template_df = pd.DataFrame({'Кабинет': ['A101', 'A102', 'D201']})
students_template_df = pd.DataFrame(
    {
        'ИИН': ['123456789012', '234567890123', '345678901234'],
        'Сыныбы': ['7A', '7A', '8B'],
        'Тегі': ['Иванов', 'Сейітов', 'Ахметова'],
        'Аты': ['Алексей', 'Нұржан', 'Алина'],
    }
)

st.markdown(
    """
    <div class='hero'>
      <h1>Оқушыларды кабинеттерге бөлу</h1>
      <p>Excel файлдарын жүктеңіз, жүйе автоматты түрде оқушыларды кабинетке бөледі.</p>
    </div>
    """,
    unsafe_allow_html=True,
)

st.markdown('### Excel үлгілері')
t1, t2 = st.columns(2)
with t1:
    st.download_button(
        label='Кабинет үлгісін жүктеу',
        data=df_to_excel_bytes(rooms_template_df),
        file_name='Шаблон_Кабинет.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        use_container_width=True,
    )
with t2:
    st.download_button(
        label='Оқушылар үлгісін жүктеу',
        data=df_to_excel_bytes(students_template_df),
        file_name='Шаблон_Оқушылар.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        use_container_width=True,
    )

st.caption('Алдымен үлгілерді жүктеп алыңыз, деректерді толтырыңыз, содан кейін төмендегі өрістерге жүктеңіз.')

col_left, col_right = st.columns(2)
with col_left:
    rooms_file = st.file_uploader('Кабинет файлы (.xlsx)', type=['xlsx'])
with col_right:
    students_file = st.file_uploader('Оқушылар файлы (.xlsx)', type=['xlsx'])

st.markdown('### Бөлу баптаулары')
s1, s2, s3 = st.columns(3)
with s1:
    max_per_room = st.number_input(
        'Кабинеттегі ең көп оқушы саны',
        min_value=1,
        max_value=100,
        value=23,
        step=1,
    )
with s2:
    max_per_class_in_room = st.number_input(
        'Бір сыныптан кабинеттегі ең көп саны',
        min_value=1,
        max_value=30,
        value=3,
        step=1,
    )
with s3:
    attempts = st.number_input(
        'Іріктеу әрекеттері саны',
        min_value=50,
        max_value=2000,
        value=400,
        step=50,
    )

run_clicked = st.button('Распределить', type='primary', use_container_width=True)

if run_clicked:
    if not rooms_file or not students_file:
        st.error('Екі файлды да жүктеңіз: кабинет және оқушылар файлы.')
    else:
        try:
            rooms_df = read_excel_from_upload(rooms_file)
            students_df = read_excel_from_upload(students_file)
            result = generate_outputs(
                rooms_df,
                students_df,
                max_per_room=int(max_per_room),
                max_per_class_in_room=int(max_per_class_in_room),
                attempts=int(attempts),
            )

            st.success('Распределение завершено успешно.')

            c1, c2, c3 = st.columns(3)
            c1.markdown(f"<div class='stat'><b>Барлығы:</b> {result['total_count']}</div>", unsafe_allow_html=True)
            c2.markdown(f"<div class='stat'><b>Орналасты:</b> {result['assigned_count']}</div>", unsafe_allow_html=True)
            c3.markdown(f"<div class='stat'><b>Орналаспаған:</b> {result['unassigned_count']}</div>", unsafe_allow_html=True)

            st.markdown('### Жүктеу')
            ready_path = Path(result['ready_name'])
            reference_path = Path(result['reference_name'])

            with open(ready_path, 'rb') as f:
                st.download_button(
                    label='Скачать дайын тізім',
                    data=f.read(),
                    file_name=ready_path.name,
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    use_container_width=True,
                )

            with open(reference_path, 'rb') as f:
                st.download_button(
                    label='Скачать анықтамаға іліп қою үшін',
                    data=f.read(),
                    file_name=reference_path.name,
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    use_container_width=True,
                )

            if result['unassigned_name']:
                unassigned_path = Path(result['unassigned_name'])
                with open(unassigned_path, 'rb') as f:
                    st.download_button(
                        label='Скачать орналаспағандар тізімі',
                        data=f.read(),
                        file_name=unassigned_path.name,
                        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                        use_container_width=True,
                    )

            for fname in [result['ready_name'], result['reference_name'], result['unassigned_name']]:
                if fname and os.path.exists(fname):
                    os.remove(fname)

        except Exception as e:
            st.error(f'Қате: {e}')
