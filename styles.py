def load_styles():
    return """
    <link href="https://fonts.googleapis.com/css2?family=Montserrat:wght@400;500;600&display=swap" rel="stylesheet">

    <style>
    html, body, [class*="css"] {
        font-family: 'Montserrat', sans-serif;
        background-color: #0e1117;
        color: #ffffff;
    }

    /* кнопки */
    .stButton>button {
        background-color: #ff4b4b;
        color: white;
        border-radius: 10px;
        font-weight: 600;
        border: none;
        padding: 10px 20px;
    }

    /* інпути */
    input, textarea, select {
        background-color: #1c1f26 !important;
        color: white !important;
        border-radius: 8px !important;
        border: 1px solid #2a2d35 !important;
    }

    /* блоки */
    .card {
        background: #1c1f26;
        padding: 15px;
        border-radius: 12px;
        margin-bottom: 10px;
        border: 1px solid #2a2d35;
    }

    /* оцінки */
    .score-good {
        background: #4caf50;
        padding: 10px;
        border-radius: 8px;
        margin-bottom: 6px;
    }

    .score-mid {
        background: #2a2d35;
        padding: 10px;
        border-radius: 8px;
        margin-bottom: 6px;
    }

    .score-bad {
        background: #ff4b4b;
        padding: 10px;
        border-radius: 8px;
        margin-bottom: 6px;
    }

    /* прибрати підсвітку коду */
    code {
        color: white !important;
        background: none !important;
    }
    </style>
    """
