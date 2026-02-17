# Deploy Guide

## 1) Local check
```bash
py -3 -m pip install -r requirements.txt
py -3 -m streamlit run app.py
```

## 2) Streamlit Community Cloud (быстрее всего)
1. Загрузите проект в GitHub (файлы: `app.py`, `allocator.py`, `requirements.txt`).
2. Откройте https://share.streamlit.io/
3. Нажмите **New app**.
4. Выберите репозиторий и ветку.
5. Main file path: `app.py`
6. Deploy.

## 3) Render
В проект уже добавлен `render.yaml`.

1. Загрузите проект в GitHub.
2. В Render нажмите **New +** -> **Blueprint**.
3. Подключите репозиторий.
4. Render прочитает `render.yaml` и поднимет сервис автоматически.

Альтернатива без Blueprint:
- Build Command: `pip install -r requirements.txt`
- Start Command: `streamlit run app.py --server.port=$PORT --server.address=0.0.0.0 --server.headless=true`

## Важно
- Не публикуйте личные Excel-файлы в репозитории.
- Если в репо есть `students.xlsx` и `Кабинет.xlsx`, удалите их перед публичным деплоем.
