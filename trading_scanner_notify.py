#!/usr/bin/env python3
"""
trading_scanner_notify.py
Автоматическая система скоринга и email-уведомлений по Торговому плану 5.6.1
Версия: 6.0 (изолированная копия)
Дата: 07.06.2026

Полностью изолирована от оригинального trading_scanner_5.6.py
Запуск: 6 раз в сутки (00:10, 04:10, 08:10, 12:10, 16:10, 20:10 МСК) через GitHub Actions / Task Scheduler

Требования:
    pip install ccxt pandas numpy python-dotenv pytz

Настройка: через переменные окружения (.env или GitHub Secrets)
"""

import os
import sys
import logging
from logging.handlers import RotatingFileHandler
import atexit
import time
from datetime import datetime
import pytz
from dotenv import load_dotenv

import pandas as pd
import numpy as np
import ccxt
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# ==================== ЗАГРУЗКА КОНФИГУРАЦИИ ====================
load_dotenv()

# === ПУТИ (ОБЯЗАТЕЛЬНО УНИКАЛЬНЫЕ - ИЗОЛЯЦИЯ) ===
EXCEL_FILE = os.getenv('EXCEL_FILE', os.path.abspath('./scoring_notify.xlsx'))
LOG_DIR = os.getenv('LOG_DIR', os.path.abspath('./logs_notify'))
CACHE_DIR = os.getenv('CACHE_DIR', os.path.abspath('./cache_notify'))

os.makedirs(LOG_DIR, exist_ok=True)
os.makedirs(CACHE_DIR, exist_ok=True)

# === ПОРОГИ СКОРИНГА (из Плана 5.6.1) ===
LONG_THRESHOLD = int(os.getenv('LONG_THRESHOLD', 26))
SHORT_THRESHOLD = int(os.getenv('SHORT_THRESHOLD', 24))

# === BYBIT API (ОТДЕЛЬНЫЕ КЛЮЧИ - ОБЯЗАТЕЛЬНО!) ===
BYBIT_API_KEY = os.getenv('BYBIT_API_KEY', '')
BYBIT_API_SECRET = os.getenv('BYBIT_API_SECRET', '')

# === SMTP ДЛЯ EMAIL-УВЕДОМЛЕНИЙ ===
SMTP_SERVER = os.getenv('SMTP_SERVER', 'smtp.gmail.com')
SMTP_PORT = int(os.getenv('SMTP_PORT', 587))
SMTP_USER = os.getenv('SMTP_USER', '')
SMTP_PASSWORD = os.getenv('SMTP_PASSWORD', '')
EMAIL_FROM = os.getenv('EMAIL_FROM', SMTP_USER)
EMAIL_TO = os.getenv('EMAIL_TO', '')  # можно несколько через запятую

# ==================== ПРОВЕРКИ ИЗОЛЯЦИИ (из ТЗ 2.3.3) ====================
def check_isolation():
    """Строгие проверки, чтобы случайно не затронуть оригинальную систему"""
    ORIGINAL_EXCEL = r"D:\Денис\Desktop\ИИ\Trading_Plan_5.6_Scoring.xlsx"
    if os.path.abspath(EXCEL_FILE) == os.path.abspath(ORIGINAL_EXCEL):
        raise Exception(
            "ОШИБКА КРИТИЧЕСКАЯ: Путь к Excel-файлу совпадает с оригиналом!\n"
            "Измените EXCEL_FILE в .env или переменных окружения."
        )

    ORIGINAL_LOG_DIR = r"D:\Денис\Desktop\ИИ\logs"
    if LOG_DIR and os.path.abspath(LOG_DIR) == os.path.abspath(ORIGINAL_LOG_DIR):
        raise Exception(
            "ОШИБКА КРИТИЧЕСКАЯ: Каталог логов совпадает с оригинальным!\n"
            "Измените LOG_DIR в .env или переменных окружения."
        )

    # Дополнительная проверка на cache (если используется)
    ORIGINAL_CACHE = r"D:\Денис\Desktop\ИИ\cache"
    if CACHE_DIR and os.path.abspath(CACHE_DIR) == os.path.abspath(ORIGINAL_CACHE):
        raise Exception(
            "ОШИБКА КРИТИЧЕСКАЯ: Каталог кэша совпадает с оригинальным!\n"
            "Измените CACHE_DIR."
        )

    print("✅ Проверка изоляции пройдена успешно. Работаем только с копией.")


# ==================== ФАЙЛОВАЯ БЛОКИРОВКА (из ТЗ 2.3.5) ====================
LOCKFILE = os.path.join(LOG_DIR, "trading_scanner_notify.lock")

def acquire_lock():
    """Предотвращает одновременный запуск нескольких экземпляров"""
    if os.path.exists(LOCKFILE):
        print("⚠️  Предыдущий запуск ещё не завершён (lock-файл существует). Выход.")
        sys.exit(0)
    with open(LOCKFILE, 'w', encoding='utf-8') as f:
        f.write(f"PID: {os.getpid()}\nStarted: {datetime.now().isoformat()}")
    atexit.register(lambda: os.remove(LOCKFILE) if os.path.exists(LOCKFILE) else None)
    print(f"🔒 Lock-файл создан: {LOCKFILE}")


# ==================== ЛОГИРОВАНИЕ (из ТЗ 2.3.7) ====================
def setup_logging():
    """Настройка ротации логов в ./logs_notify/"""
    log_file = os.path.join(LOG_DIR, 'trading_scanner_notify.log')
    handler = RotatingFileHandler(
        log_file,
        maxBytes=10 * 1024 * 1024,  # 10 MB
        backupCount=7,
        encoding='utf-8'
    )
    formatter = logging.Formatter(
        '%(asctime)s [%(levelname)s] %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )
    handler.setFormatter(formatter)

    logger = logging.getLogger()
    logger.setLevel(logging.INFO)
    logger.handlers.clear()  # очистка на случай перезапуска
    logger.addHandler(handler)

    # Дублирование в консоль
    console = logging.StreamHandler(sys.stdout)
    console.setFormatter(formatter)
    logger.addHandler(console)

    return logger


# ==================== ИНДИКАТОРЫ (базовые реализации) ====================
def calculate_ema(series: pd.Series, period: int) -> pd.Series:
    return series.ewm(span=period, adjust=False).mean()

def calculate_rsi(series: pd.Series, period: int = 14) -> pd.Series:
    delta = series.diff()
    gain = delta.where(delta > 0, 0.0).rolling(window=period).mean()
    loss = (-delta.where(delta < 0, 0.0)).rolling(window=period).mean()
    rs = gain / loss
    return 100 - (100 / (1 + rs))

def calculate_macd(series: pd.Series, fast: int = 12, slow: int = 26, signal: int = 9):
    ema_fast = calculate_ema(series, fast)
    ema_slow = calculate_ema(series, slow)
    macd_line = ema_fast - ema_slow
    signal_line = calculate_ema(macd_line, signal)
    histogram = macd_line - signal_line
    return macd_line, signal_line, histogram

def calculate_atr(df: pd.DataFrame, period: int = 14) -> pd.Series:
    high = df['high']
    low = df['low']
    close = df['close']
    tr1 = high - low
    tr2 = (high - close.shift(1)).abs()
    tr3 = (low - close.shift(1)).abs()
    tr = pd.concat([tr1, tr2, tr3], axis=1).max(axis=1)
    return tr.rolling(window=period).mean()


# ==================== ЗАГРУЗКА ДАННЫХ ====================
def fetch_ohlcv(exchange, symbol: str, timeframe: str, limit: int = 200):
    """Загрузка свечей с защитой от ошибок"""
    try:
        ohlcv = exchange.fetch_ohlcv(symbol, timeframe=timeframe, limit=limit)
        if not ohlcv:
            return None
        df = pd.DataFrame(ohlcv, columns=['timestamp', 'open', 'high', 'low', 'close', 'volume'])
        df['timestamp'] = pd.to_datetime(df['timestamp'], unit='ms')
        df.set_index('timestamp', inplace=True)
        df = df.dropna()
        return df
    except Exception as e:
        logging.error(f"Ошибка загрузки {symbol} {timeframe}: {e}")
        return None


def get_trading_symbols(exchange, max_symbols: int = 18):
    """Список популярных USDT-перпетуалов на Bybit (можно расширить)"""
    try:
        markets = exchange.load_markets()
        # Фильтруем только активные линейные перпетуалы USDT
        symbols = [
            s for s in markets.keys()
            if s.endswith(':USDT') and markets[s].get('active', False)
            and markets[s].get('type') in ('swap', 'future')
        ]
        # Приоритетные монеты (можно заменить на динамический топ по объёму)
        priority = [
            'BTC/USDT:USDT', 'ETH/USDT:USDT', 'SOL/USDT:USDT', 'XRP/USDT:USDT',
            'DOGE/USDT:USDT', 'TON/USDT:USDT', 'ADA/USDT:USDT', 'AVAX/USDT:USDT',
            'LINK/USDT:USDT', 'LTC/USDT:USDT', 'NEAR/USDT:USDT', 'DOT/USDT:USDT',
            'TRX/USDT:USDT', 'UNI/USDT:USDT', 'AAVE/USDT:USDT', 'SUI/USDT:USDT',
            'APT/USDT:USDT', 'ARB/USDT:USDT'
        ]
        final_symbols = [s for s in priority if s in symbols][:max_symbols]
        if not final_symbols:
            final_symbols = symbols[:max_symbols]
        logging.info(f"Выбрано {len(final_symbols)} символов для сканирования")
        return final_symbols
    except Exception as e:
        logging.error(f"Ошибка загрузки рынков: {e}")
        return ['BTC/USDT:USDT', 'ETH/USDT:USDT', 'SOL/USDT:USDT']


# ==================== СКОРИНГ (ЗАГЛУШКА - ЗАМЕНИТЬ НА РЕАЛЬНУЮ ЛОГИКУ ИЗ ОРИГИНАЛА) ====================
def compute_score_and_recommendation(symbol: str, df_1d: pd.DataFrame, df_4h: pd.DataFrame):
    """
    ВАЖНО: Здесь должна быть ТОЧНАЯ копия логики скоринга из trading_scanner_5.6.py
    и правил из "ПЛАН 5.6.1 УСИЛЕННЫМИ ПРАВИЛАМИ ДЛЯ ПЯТНИЦЫ И ВЫХОДНЫХ.docx"

    Текущая версия — демонстрационная заглушка на основе обязательных индикаторов.
    Замените эту функцию на реальную при переносе кода из оригинала.
    """
    if df_1d is None or df_4h is None or len(df_1d) < 200 or len(df_4h) < 60:
        return None

    close_1d = df_1d['close']
    close_4h = df_4h['close']

    # Обязательные индикаторы по ТЗ
    ema50_1d = calculate_ema(close_1d, 50).iloc[-1]
    ema200_1d = calculate_ema(close_1d, 200).iloc[-1]
    ema20_4h = calculate_ema(close_4h, 20).iloc[-1]

    rsi_4h = calculate_rsi(close_4h, 14).iloc[-1]
    _, _, macd_hist_4h = calculate_macd(close_4h)
    macd_hist_4h = macd_hist_4h.iloc[-1]

    atr_4h = calculate_atr(df_4h, 14).iloc[-1]
    atr_1d = calculate_atr(df_1d, 14).iloc[-1]

    volume_1d = df_1d['volume'].iloc[-1]
    volume_4h = df_4h['volume'].iloc[-1]
    current_price = close_4h.iloc[-1]

    # ==================== ПРИМЕР ПРОСТОГО СКОРИНГА (ЗАМЕНИТЬ!) ====================
    # В реальном плане 5.6.1 используется сложная система баллов за:
    # - MTF alignment (1D + 4H)
    # - EMA пересечения и положение цены
    # - MACD / RSI / Volume confirmation
    # - ATR volatility filter
    # - Специальные правила для пятницы/выходных и азиатской сессии
    # Здесь — упрощённая демонстрация

    score = 0.0

    # Тренд 1D (обязательно)
    if ema50_1d > ema200_1d:
        score += 12
    elif ema50_1d < ema200_1d:
        score -= 8

    # Положение цены относительно EMA20 на 4H
    if current_price > ema20_4h:
        score += 6
    else:
        score -= 4

    # MACD гистограмма 4H
    if macd_hist_4h > 0:
        score += 5
    else:
        score -= 3

    # RSI не в перекупленности/перепроданности (зона 35-65)
    if 35 < rsi_4h < 65:
        score += 4
    elif rsi_4h > 75 or rsi_4h < 25:
        score -= 5

    # Объём выше среднего (20 периодов)
    vol_ma_4h = df_4h['volume'].rolling(20).mean().iloc[-1]
    if volume_4h > vol_ma_4h * 1.2:
        score += 4
    elif volume_4h < vol_ma_4h * 0.6:
        score -= 3

    # Дополнительный бонус за сильный тренд
    if abs(ema50_1d - ema200_1d) / ema200_1d > 0.03:
        score += 3

    # Ограничение диапазона
    score = max(0.0, min(35.0, score))

    # Определение направления и рекомендации
    if score >= LONG_THRESHOLD:
        recommendation = "ПРОХОДИТ"
        direction = "LONG"
    elif score >= SHORT_THRESHOLD:
        recommendation = "ПРОХОДИТ"
        direction = "SHORT"
    else:
        recommendation = "НЕ ПРОХОДИТ"
        direction = None

    msk_time = datetime.now(pytz.timezone('Europe/Moscow')).strftime('%Y-%m-%d %H:%M:%S MSK')

    return {
        'Symbol': symbol,
        'Direction': direction or '-',
        'Total_Score': round(score, 1),
        'Recommendation': recommendation,
        'Price': round(current_price, 4),
        'EMA50_1D': round(ema50_1d, 4),
        'EMA200_1D': round(ema200_1d, 4),
        'EMA20_4H': round(ema20_4h, 4),
        'RSI_4H': round(rsi_4h, 1),
        'MACD_Hist_4H': round(macd_hist_4h, 4),
        'ATR_4H': round(atr_4h, 4),
        'ATR_1D': round(atr_1d, 4),
        'Volume_1D': int(volume_1d),
        'Volume_4H': int(volume_4h),
        'MTF_ok': 'Да' if (ema50_1d > ema200_1d and current_price > ema20_4h) else 'Нет',
        'Confirmed_4H': 'Да' if macd_hist_4h > 0 else 'Нет',
        'R/R': round(2.8, 1),  # placeholder (реальный считается по ATR)
        'Timestamp_MSK': msk_time
    }


# ==================== ОСНОВНОЙ ПРОЦЕСС СКОРИНГА ====================
def run_scoring():
    logger = logging.getLogger()
    msk = pytz.timezone('Europe/Moscow')
    now_msk = datetime.now(msk)
    logger.info("=" * 60)
    logger.info(f"🚀 ЗАПУСК АВТОМАТИЧЕСКОГО СКОРИНГА (копия notify) | {now_msk.strftime('%d.%m.%Y %H:%M:%S MSK')}")
    logger.info("=" * 60)

    # Инициализация Bybit (анонимно если ключей нет)
    exchange_params = {'enableRateLimit': True}
    if BYBIT_API_KEY and BYBIT_API_SECRET:
        exchange_params.update({'apiKey': BYBIT_API_KEY, 'secret': BYBIT_API_SECRET})
    else:
        logger.warning("⚠️  API-ключи Bybit не заданы. Работаем в анонимном режиме (риск лимитов API).")

    exchange = ccxt.bybit(exchange_params)

    symbols = get_trading_symbols(exchange)
    results = []

    for i, symbol in enumerate(symbols, 1):
        logger.info(f"[{i}/{len(symbols)}] Сканирование {symbol}...")
        df_1d = fetch_ohlcv(exchange, symbol, '1d', limit=250)
        df_4h = fetch_ohlcv(exchange, symbol, '4h', limit=120)

        if df_1d is None or df_4h is None:
            logger.warning(f"Пропуск {symbol} — нет данных")
            continue

        data = compute_score_and_recommendation(symbol, df_1d, df_4h)
        if data:
            results.append(data)
            logger.info(f"  → Score: {data['Total_Score']} | Rec: {data['Recommendation']} | Dir: {data['Direction']}")

        time.sleep(0.6)  # уважение к rate limits Bybit

    if not results:
        logger.warning("Нет результатов скоринга.")
        return pd.DataFrame()

    df = pd.DataFrame(results)
    # Сортировка по скору
    df = df.sort_values('Total_Score', ascending=False).reset_index(drop=True)

    # Сохранение полного результата (как в оригинале)
    try:
        with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Scoring_5.6.1')
        logger.info(f"✅ Полные результаты сохранены: {EXCEL_FILE}")
    except Exception as e:
        logger.error(f"Ошибка записи Excel: {e}")

    return df


# ==================== ФИЛЬТРАЦИЯ СИГНАЛОВ ====================
def get_passing_signals(df: pd.DataFrame):
    if df.empty:
        return pd.DataFrame()
    # Только те, у кого Recommendation == "ПРОХОДИТ" (т.е. score >= порога и есть направление)
    passing = df[df['Recommendation'] == 'ПРОХОДИТ'].copy()
    return passing


# ==================== ОТПРАВКА EMAIL (из ТЗ 2.3.6) ====================
def send_email_notification(signals_df: pd.DataFrame):
    if signals_df.empty:
        return False

    if not SMTP_USER or not SMTP_PASSWORD or not EMAIL_TO:
        logging.warning("SMTP-параметры не настроены. Email не отправлен.")
        return False

    try:
        msk = pytz.timezone('Europe/Moscow')
        now_str = datetime.now(msk).strftime('%d.%m.%Y %H:%M:%S MSK')

        msg = MIMEMultipart('alternative')
        msg['Subject'] = f"🚨 СИГНАЛЫ 5.6.1 | Автоскор {now_str}"
        msg['From'] = EMAIL_FROM
        msg['To'] = EMAIL_TO

        # Красивый HTML
        html_table = signals_df.to_html(index=False, escape=False, border=1, justify='center')

        html_body = f"""
        <html>
          <head>
            <style>
              body {{ font-family: Arial, sans-serif; }}
              table {{ border-collapse: collapse; width: 100%; }}
              th, td {{ border: 1px solid #ddd; padding: 8px; text-align: center; }}
              th {{ background-color: #1a73e8; color: white; }}
              .long {{ background-color: #d4edda; }}
              .short {{ background-color: #f8d7da; }}
            </style>
          </head>
          <body>
            <h2>📈 Автоматические торговые сигналы — План 5.6.1</h2>
            <p><b>Время сканирования:</b> {now_str}</p>
            <p><b>Пороги:</b> LONG ≥ {LONG_THRESHOLD} | SHORT ≥ {SHORT_THRESHOLD}</p>
            <p><b>Найдено сигналов:</b> {len(signals_df)}</p>
            <br>
            {html_table}
            <br>
            <p><i>⚠️ Это автоматическое уведомление. Все решения принимайте самостоятельно. 
            Риск-менеджмент обязателен (SL по ATR, частичная фиксация при +1R).</i></p>
            <p>Полный отчёт: <b>scoring_notify.xlsx</b></p>
          </body>
        </html>
        """

        msg.attach(MIMEText(html_body, 'html', 'utf-8'))

        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(SMTP_USER, SMTP_PASSWORD)
            server.sendmail(EMAIL_FROM, [e.strip() for e in EMAIL_TO.split(',')], msg.as_string())

        logging.info(f"✅ Email успешно отправлен на: {EMAIL_TO} ({len(signals_df)} сигналов)")
        return True

    except Exception as e:
        logging.error(f"❌ Ошибка отправки email: {e}")
        return False


# ==================== ГЛАВНАЯ ФУНКЦИЯ ====================
def main():
    logger = setup_logging()

    try:
        check_isolation()
        acquire_lock()

        df_results = run_scoring()
        passing_signals = get_passing_signals(df_results)

        if not passing_signals.empty:
            logger.info(f"🎯 Найдено {len(passing_signals)} ПРОХОДЯЩИХ сигналов!")
            send_email_notification(passing_signals)
        else:
            logger.info("ℹ️  Проходящих сигналов нет. Email не отправлен (экономия).")

        logger.info("🏁 Работа скрипта завершена успешно.")

    except Exception as e:
        logging.critical(f"💥 КРИТИЧЕСКАЯ ОШИБКА: {e}", exc_info=True)
        sys.exit(1)


if __name__ == "__main__":
    main()
