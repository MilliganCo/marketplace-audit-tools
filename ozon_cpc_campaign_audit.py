"""Аудит CPC-кампаний Ozon по выгрузке статистики.
GUI, опции и расширения — см. IMPROVEMENTS ниже.
"""
from __future__ import annotations

import argparse
import math
import re
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Optional, Tuple

import pandas as pd


REQUIRED_COLUMNS = {
    "ID кампании",
    "Название кампании",
    "Статус",
    "Расход, ₽",
    "Показы",
    "Клики",
    "CTR, %",
    "Средняя стоимость клика, ₽",
    "Заказы, шт.",
    "Продажи, ₽",
    "ДРР, %",
    "Стратегия",
    "В корзину",
}


@dataclass
class Benchmarks:
    active_ctr_median: float
    active_ctr_p25: float
    active_click_to_cart_median: float
    active_click_to_cart_p25: float
    active_click_to_order_median: float
    active_click_to_order_p25: float
    active_cart_to_order_median: float
    active_cart_to_order_p25: float
    active_cpc_median: float
    active_orders_median: float
    inactive_drr_median: float


def pick_file_dialog() -> Optional[Path]:
    try:
        import tkinter as tk
        from tkinter import filedialog
    except Exception:
        return None

    root = tk.Tk()
    root.withdraw()
    root.update()
    file_path = filedialog.askopenfilename(
        title="Выберите отчёт Ozon CPC",
        filetypes=[("Excel files", "*.xlsx *.xls")],
    )
    root.destroy()
    return Path(file_path) if file_path else None


def run_gui(initial_file: Optional[Path] = None) -> int:
    """Запуск графического интерфейса аудита."""
    import tkinter as tk
    from tkinter import filedialog, messagebox, scrolledtext

    root = tk.Tk()
    root.title("Ozon CPC — Аудит кампаний")
    root.minsize(520, 420)
    root.geometry("560x480")

    file_path_var = tk.StringVar(value=str(initial_file) if initial_file else "")
    include_inactive_var = tk.BooleanVar(value=True)
    suggest_strategy_var = tk.BooleanVar(value=True)
    target_drr_var = tk.StringVar(value="0.30")

    def choose_file():
        path = filedialog.askopenfilename(
            title="Выберите отчёт Ozon CPC",
            filetypes=[("Excel", "*.xlsx *.xls"), ("Все файлы", "*.*")],
        )
        if path:
            file_path_var.set(path)

    def run_audit():
        path_str = file_path_var.get().strip()
        if not path_str:
            messagebox.showerror("Ошибка", "Выберите файл для аудита.")
            return
        file_path = Path(path_str)
        if not file_path.exists():
            messagebox.showerror("Ошибка", f"Файл не найден: {file_path}")
            return
        try:
            target_drr = float(target_drr_var.get().replace(",", "."))
        except ValueError:
            messagebox.showerror("Ошибка", "Целевой ДРР должен быть числом (например, 0.30).")
            return
        log_area.delete("1.0", tk.END)
        log_area.insert(tk.END, "Запуск аудита...\n", "info")
        root.update()
        try:
            xlsx_out, txt_out = analyze(
                file_path,
                target_drr=target_drr,
                include_inactive=include_inactive_var.get(),
                suggest_strategy=suggest_strategy_var.get(),
            )
            log_area.insert(tk.END, f"Готово.\nExcel: {xlsx_out}\nTXT:   {txt_out}\n", "success")
            messagebox.showinfo("Готово", f"Аудит завершён.\n\nExcel: {xlsx_out.name}\nTXT: {txt_out.name}")
        except Exception as exc:
            log_area.insert(tk.END, f"Ошибка: {exc}\n", "error")
            messagebox.showerror("Ошибка", str(exc))

    main = tk.Frame(root, padx=12, pady=12)
    main.pack(fill=tk.BOTH, expand=True)

    f_row = tk.Frame(main)
    f_row.pack(fill=tk.X, pady=(0, 8))
    tk.Label(f_row, text="Файл:").pack(side=tk.LEFT, padx=(0, 6))
    tk.Entry(f_row, textvariable=file_path_var, width=45).pack(side=tk.LEFT, fill=tk.X, expand=True)
    tk.Button(f_row, text="Обзор…", command=choose_file).pack(side=tk.LEFT, padx=(6, 0))

    opt_frame = tk.LabelFrame(main, text="Опции", padx=8, pady=8)
    opt_frame.pack(fill=tk.X, pady=(0, 8))
    tk.Checkbutton(
        opt_frame,
        text="Учитывать неактивные кампании (бенчмарки и поиск похожих)",
        variable=include_inactive_var,
    ).pack(anchor=tk.W)
    tk.Checkbutton(
        opt_frame,
        text="Предлагать смену стратегии и CPC (иначе только «оставить» / «выключить»)",
        variable=suggest_strategy_var,
    ).pack(anchor=tk.W)

    drr_row = tk.Frame(main)
    drr_row.pack(fill=tk.X, pady=(0, 8))
    tk.Label(drr_row, text="Целевой ДРР (доля, напр. 0.30):").pack(side=tk.LEFT, padx=(0, 6))
    tk.Entry(drr_row, textvariable=target_drr_var, width=8).pack(side=tk.LEFT)

    tk.Button(main, text="Запустить аудит", command=run_audit).pack(pady=(8, 4))

    log_label = tk.Label(main, text="Лог:")
    log_label.pack(anchor=tk.W, pady=(8, 2))
    log_area = scrolledtext.ScrolledText(main, height=10, wrap=tk.WORD, state=tk.NORMAL)
    log_area.pack(fill=tk.BOTH, expand=True)
    log_area.tag_config("info", foreground="gray")
    log_area.tag_config("success", foreground="green")
    log_area.tag_config("error", foreground="red")

    root.mainloop()
    return 0


def find_header_row_and_sheet(file_path: Path) -> Tuple[str, int]:
    with pd.ExcelFile(file_path) as xls:
        for sheet_name in xls.sheet_names:
            raw = pd.read_excel(xls, sheet_name=sheet_name, header=None)
            max_rows = min(len(raw), 15)
            for i in range(max_rows):
                row_values = {str(v).strip() for v in raw.iloc[i].tolist() if pd.notna(v)}
                if len(REQUIRED_COLUMNS & row_values) >= 8:
                    return sheet_name, i
    raise ValueError("Не удалось найти строку заголовков. Проверьте структуру отчёта Ozon.")


def to_float(value) -> float:
    if pd.isna(value):
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace("\xa0", " ")
    text = text.replace("₽", "").replace("%", "")
    text = text.replace(" ", "")
    if text.count(",") == 1 and text.count(".") == 0:
        text = text.replace(",", ".")
    elif text.count(",") > 0 and text.count(".") > 0:
        text = text.replace(" ", "").replace(",", "")
    try:
        return float(text)
    except Exception:
        return 0.0


def safe_div(a: float, b: float) -> float:
    if not b:
        return 0.0
    return a / b


def normalize_rate(series: pd.Series) -> pd.Series:
    numeric = series.apply(to_float)
    # Ozon exports rates either as fractions (0.129 = 12.9%) or as displayed percentages (12.9).
    # We treat the column as percentages only when most values look like whole percents.
    clean = numeric.dropna()
    if clean.empty:
        return numeric
    share_above_one = float((clean > 1).mean())
    p75 = float(clean.quantile(0.75))
    should_divide = share_above_one > 0.5 or p75 > 1.5
    return numeric / 100.0 if should_divide else numeric


AUDIT_OUTPUT_ALIASES = {
    "Средний CPC, ₽": "Средняя стоимость клика, ₽",
}


def load_report(file_path: Path) -> pd.DataFrame:
    sheet_name, header_row = find_header_row_and_sheet(file_path)
    df = pd.read_excel(file_path, sheet_name=sheet_name, header=header_row)
    df = df.dropna(how="all").copy()

    for alias, canonical in AUDIT_OUTPUT_ALIASES.items():
        if alias in df.columns and canonical not in df.columns:
            df = df.rename(columns={alias: canonical})

    missing = REQUIRED_COLUMNS - set(df.columns)
    if missing:
        raise ValueError(f"В файле не хватает столбцов: {', '.join(sorted(missing))}")

    numeric_cols = [
        "Расход, ₽",
        "Показы",
        "Клики",
        "Средняя стоимость клика, ₽",
        "Заказы, шт.",
        "Продажи, ₽",
        "Бюджет, ₽",
        "В корзину",
    ]
    for col in numeric_cols:
        if col in df.columns:
            df[col] = df[col].apply(to_float)

    df["CTR_frac"] = normalize_rate(df["CTR, %"])
    df["DRR_frac"] = normalize_rate(df["ДРР, %"])

    df["CTR, %"] = df["CTR_frac"] * 100
    df["ДРР, %"] = df["DRR_frac"] * 100

    df["Расход_calc"] = df["Расход, ₽"]
    df["Показы_calc"] = df["Показы"]
    df["Клики_calc"] = df["Клики"]
    df["Заказы_calc"] = df["Заказы, шт."]
    df["Продажи_calc"] = df["Продажи, ₽"]
    df["Корзина_calc"] = df["В корзину"]
    df["CPC_calc"] = df["Средняя стоимость клика, ₽"].where(
        df["Средняя стоимость клика, ₽"] > 0,
        df["Расход, ₽"] / df["Клики"].replace(0, pd.NA)
    ).fillna(0)
    df["Клик→корзина, %"] = safe_div_series(df["Корзина_calc"], df["Клики_calc"]) * 100
    df["Клик→заказ, %"] = safe_div_series(df["Заказы_calc"], df["Клики_calc"]) * 100
    df["Корзина→заказ, %"] = safe_div_series(df["Заказы_calc"], df["Корзина_calc"]) * 100
    df["Выручка/клик, ₽"] = safe_div_series(df["Продажи_calc"], df["Клики_calc"])
    df["Средний чек, ₽"] = safe_div_series(df["Продажи_calc"], df["Заказы_calc"])

    return df


def safe_div_series(a: pd.Series, b: pd.Series) -> pd.Series:
    denom = b.where(b != 0, other=float("nan"))
    return a.div(denom).fillna(0.0)


def status_is_active(status: str) -> bool:
    status = str(status).strip().lower()
    return status.startswith("актив")


def build_benchmarks(active_df: pd.DataFrame, inactive_df: pd.DataFrame) -> Benchmarks:
    def q(series: pd.Series, quantile: float) -> float:
        clean = series.replace([math.inf, -math.inf], pd.NA).dropna()
        if clean.empty:
            return 0.0
        return float(clean.quantile(quantile))

    return Benchmarks(
        active_ctr_median=q(active_df["CTR_frac"], 0.50),
        active_ctr_p25=q(active_df["CTR_frac"], 0.25),
        active_click_to_cart_median=q(active_df["Клик→корзина, %"] / 100, 0.50),
        active_click_to_cart_p25=q(active_df["Клик→корзина, %"] / 100, 0.25),
        active_click_to_order_median=q(active_df["Клик→заказ, %"] / 100, 0.50),
        active_click_to_order_p25=q(active_df["Клик→заказ, %"] / 100, 0.25),
        active_cart_to_order_median=q(active_df["Корзина→заказ, %"] / 100, 0.50),
        active_cart_to_order_p25=q(active_df["Корзина→заказ, %"] / 100, 0.25),
        active_cpc_median=q(active_df["CPC_calc"], 0.50),
        active_orders_median=q(active_df["Заказы_calc"], 0.50),
        inactive_drr_median=q(inactive_df["DRR_frac"], 0.50) if not inactive_df.empty else 0.0,
    )


def normalize_name(text: str) -> str:
    text = str(text).lower()
    text = re.sub(r"[\/_\-]+", " ", text)
    text = re.sub(r"[^a-zа-я0-9 ]+", " ", text, flags=re.IGNORECASE)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def token_set(text: str) -> set[str]:
    words = normalize_name(text).split()
    stop = {"м", "ж", "с", "г", "и", "по", "за", "с"}
    return {w for w in words if len(w) > 2 and w not in stop}


def best_inactive_match(row: pd.Series, inactive_df: pd.DataFrame) -> Optional[pd.Series]:
    if inactive_df.empty:
        return None

    base = normalize_name(row["Название кампании"])
    base_tokens = token_set(base)
    best_idx = None
    best_score = 0.0

    for idx, candidate in inactive_df.iterrows():
        cand_name = normalize_name(candidate["Название кампании"])
        seq = similar_ratio(base, cand_name)
        cand_tokens = token_set(cand_name)
        inter = len(base_tokens & cand_tokens)
        union = len(base_tokens | cand_tokens) or 1
        jaccard = inter / union
        score = 0.65 * seq + 0.35 * jaccard
        if score > best_score:
            best_score = score
            best_idx = idx

    if best_idx is None or best_score < 0.84:
        return None
    return inactive_df.loc[best_idx]


def similar_ratio(a: str, b: str) -> float:
    try:
        from difflib import SequenceMatcher
        return SequenceMatcher(None, a, b).ratio()
    except Exception:
        return 0.0


def fmt_money(v: float) -> str:
    return f"{v:,.0f} ₽".replace(",", " ")


def fmt_pct(frac: float) -> str:
    return f"{frac * 100:.1f}%"


def action_and_comment(
    row: pd.Series,
    bm: Benchmarks,
    target_drr: float,
    inactive_df: pd.DataFrame,
    *,
    suggest_strategy: bool = True,
) -> tuple[str, str, str, str, str]:
    spend = float(row["Расход_calc"])
    clicks = float(row["Клики_calc"])
    orders = float(row["Заказы_calc"])
    sales = float(row["Продажи_calc"])
    carts = float(row["Корзина_calc"])
    ctr = float(row["CTR_frac"])
    drr = float(row["DRR_frac"])
    cpc = float(row["CPC_calc"])
    c2c = safe_div(carts, clicks)
    c2o = safe_div(orders, clicks)
    cart2o = safe_div(orders, carts)
    rpc = safe_div(sales, clicks)
    max_cpc = rpc * target_drr

    reasons: list[str] = []
    advice: list[str] = []
    score = 0

    if sales <= 0 and spend > 0:
        action = "Поставить на паузу"
        reasons.append("Есть расход, но нет продаж")
        advice.extend([
            "остановить кампанию",
            "проверить цену, остатки, фото и релевантность SKU",
        ])
        return action, build_comment(reasons, advice), "; ".join(reasons), "", ""

    if drr <= target_drr * 0.75:
        score += 3
        reasons.append(f"ДРР {fmt_pct(drr)} лучше целевого {fmt_pct(target_drr)}")
    elif drr <= target_drr:
        score += 2
        reasons.append(f"ДРР {fmt_pct(drr)} в пределах целевого {fmt_pct(target_drr)}")
    elif drr <= target_drr * 1.3:
        reasons.append(f"ДРР {fmt_pct(drr)} слегка выше цели {fmt_pct(target_drr)}")
    elif drr <= target_drr * 1.8:
        score -= 2
        reasons.append(f"ДРР {fmt_pct(drr)} заметно выше цели {fmt_pct(target_drr)}")
    else:
        score -= 4
        reasons.append(f"ДРР {fmt_pct(drr)} критично выше цели {fmt_pct(target_drr)}")

    if ctr >= bm.active_ctr_median * 1.10:
        score += 1
        reasons.append(f"CTR {fmt_pct(ctr)} выше медианы активных {fmt_pct(bm.active_ctr_median)}")
    elif ctr < bm.active_ctr_p25 and clicks > 0:
        score -= 1
        reasons.append(f"CTR {fmt_pct(ctr)} ниже нижнего квартиля активных {fmt_pct(bm.active_ctr_p25)}")
        advice.append("обновить обложку, цену в оффере и заголовок карточки")

    if c2c >= bm.active_click_to_cart_median * 1.10:
        score += 1
        reasons.append(f"Клик→корзина {fmt_pct(c2c)} выше медианы активных {fmt_pct(bm.active_click_to_cart_median)}")
    elif c2c < bm.active_click_to_cart_p25 and clicks > 0:
        score -= 1
        reasons.append(f"Клик→корзина {fmt_pct(c2c)} слабый относительно активных")
        advice.append("усилить первый экран карточки и ценностное предложение")

    if cart2o >= bm.active_cart_to_order_median * 1.10:
        score += 2
        reasons.append(f"Корзина→заказ {fmt_pct(cart2o)} выше медианы активных {fmt_pct(bm.active_cart_to_order_median)}")
    elif cart2o < bm.active_cart_to_order_p25 and carts > 0:
        score -= 2
        reasons.append(f"Корзина→заказ {fmt_pct(cart2o)} слабый относительно активных")
        advice.append("проверить цену, скидку, сроки доставки, размерный ряд и отзывы")

    if cpc <= max_cpc * 0.90 and max_cpc > 0:
        score += 1
        reasons.append(f"CPC {cpc:.2f} ₽ вписывается в допустимый порог {max_cpc:.2f} ₽")
    elif cpc > max_cpc * 1.10 and max_cpc > 0:
        score -= 2
        reasons.append(f"CPC {cpc:.2f} ₽ выше допустимого порога {max_cpc:.2f} ₽")
        if suggest_strategy:
            if str(row.get("Стратегия", "")).strip().lower().startswith("авто"):
                advice.append("снизить агрессию автостратегии или перейти в стратегию со средней стоимостью клика")
            else:
                advice.append("снизить ставку по клику")

    if orders >= max(bm.active_orders_median * 1.2, 20):
        score += 1
        reasons.append(f"Заказов {orders:.0f}, кампания уже набрала статистику")
    elif orders < 3 and clicks >= 300:
        score -= 2
        reasons.append(f"Мало заказов ({orders:.0f}) при заметном трафике")

    matched_name = ""
    matched_comment = ""
    matched = best_inactive_match(row, inactive_df)
    if matched is not None:
        matched_name = str(matched["Название кампании"])
        matched_drr = float(matched["DRR_frac"])
        matched_c2o = safe_div(float(matched["Заказы_calc"]), float(matched["Клики_calc"]))
        if matched_drr > 0 and drr > matched_drr * 1.5:
            score -= 1
            matched_comment = (
                f"Похожая неактивная кампания «{matched_name}» ранее имела ДРР {fmt_pct(matched_drr)} "
                f"против текущих {fmt_pct(drr)}"
            )
            advice.append("сверить карточку и настройки с исторически более сильной кампанией")
        if matched_c2o > 0 and c2o < matched_c2o * 0.7:
            matched_comment = (
                matched_comment + "; " if matched_comment else ""
            ) + f"конверсия клик→заказ сейчас ниже исторической по похожему SKU"

    if drr > target_drr * 2.0 or (orders == 0 and clicks > 150):
        action = "Поставить на паузу"
    elif score >= 4:
        action = "Масштабировать"
    elif score >= 1:
        action = "Оставить и наблюдать"
    elif score >= -1:
        action = "Снизить CPC / доработать карточку"
    else:
        action = "Поставить на паузу и пересобрать"

    if not suggest_strategy:
        if action in ("Масштабировать", "Снизить CPC / доработать карточку"):
            action = "Оставить и наблюдать"
        elif action == "Поставить на паузу и пересобрать":
            action = "Поставить на паузу"
        advice = [a for a in advice if "стратеги" not in a.lower() and "ставк" not in a.lower()]

    if not advice:
        if action == "Масштабировать":
            advice.append("осторожно увеличить недельный бюджет или ставку, не ломая текущий ДРР")
        elif action == "Оставить и наблюдать":
            advice.append("сохранять кампанию в работе и еженедельно контролировать CPC и ДРР")
        elif action == "Снизить CPC / доработать карточку":
            advice.append("снизить стоимость трафика и улучшить конверсию карточки")
        else:
            advice.append("после паузы запускать заново только после правок карточки и предложения")

    comment = build_comment(reasons, advice)
    return action, comment, "; ".join(reasons), matched_name, matched_comment


def build_comment(reasons: list[str], advice: list[str]) -> str:
    reasons = unique_keep_order(reasons)
    advice = unique_keep_order(advice)
    lead = " | ".join(reasons[:3]) if reasons else "Недостаточно данных для уверенной интерпретации"
    tail = "; ".join(advice[:3]) if advice else "наблюдать за кампанией"
    return f"{lead}. Рекомендуется: {tail}."


def unique_keep_order(items: list[str]) -> list[str]:
    result = []
    seen = set()
    for item in items:
        key = item.strip().lower()
        if key and key not in seen:
            seen.add(key)
            result.append(item)
    return result


def build_summary_text(active_df: pd.DataFrame, target_drr: float) -> str:
    total_spend = active_df["Расход_calc"].sum()
    total_sales = active_df["Продажи_calc"].sum()
    total_orders = active_df["Заказы_calc"].sum()
    total_clicks = active_df["Клики_calc"].sum()
    total_impr = active_df["Показы_calc"].sum()
    total_carts = active_df["Корзина_calc"].sum()

    ctr = safe_div(total_clicks, total_impr)
    drr = safe_div(total_spend, total_sales)
    cpc = safe_div(total_spend, total_clicks)
    c2c = safe_div(total_carts, total_clicks)
    c2o = safe_div(total_orders, total_clicks)
    cart2o = safe_div(total_orders, total_carts)

    return (
        "Сводка по активным кампаниям\n"
        f"- Кампаний: {len(active_df)}\n"
        f"- Расход: {fmt_money(total_spend)}\n"
        f"- Продажи: {fmt_money(total_sales)}\n"
        f"- Заказы: {total_orders:,.0f}\n"
        f"- CTR: {fmt_pct(ctr)}\n"
        f"- CPC: {cpc:.2f} ₽\n"
        f"- Клик→корзина: {fmt_pct(c2c)}\n"
        f"- Клик→заказ: {fmt_pct(c2o)}\n"
        f"- Корзина→заказ: {fmt_pct(cart2o)}\n"
        f"- ДРР: {fmt_pct(drr)} при целевом {fmt_pct(target_drr)}\n"
    )


def analyze(
    file_path: Path,
    target_drr: float,
    output_prefix: Optional[Path] = None,
    *,
    include_inactive: bool = True,
    suggest_strategy: bool = True,
) -> tuple[Path, Path]:
    df = load_report(file_path)
    active_df = df[df["Статус"].apply(status_is_active)].copy()
    inactive_df = df[~df["Статус"].apply(status_is_active)].copy() if include_inactive else pd.DataFrame()

    if active_df.empty:
        raise ValueError("В отчёте нет активных кампаний.")

    bm = build_benchmarks(active_df, inactive_df)

    results = []
    for _, row in active_df.iterrows():
        action, comment, reasons, matched_name, matched_comment = action_and_comment(
            row, bm, target_drr, inactive_df, suggest_strategy=suggest_strategy
        )
        rpc = safe_div(float(row["Продажи_calc"]), float(row["Клики_calc"]))
        max_cpc = rpc * target_drr
        results.append(
            {
                "ID кампании": row["ID кампании"],
                "Название кампании": row["Название кампании"],
                "Статус": row["Статус"],
                "Стратегия": row.get("Стратегия", ""),
                "Расход, ₽": row["Расход_calc"],
                "Показы": row["Показы_calc"],
                "Клики": row["Клики_calc"],
                "CTR, %": row["CTR, %"],
                "Средний CPC, ₽": row["CPC_calc"],
                "Допустимый CPC, ₽": max_cpc,
                "Заказы, шт.": row["Заказы_calc"],
                "Продажи, ₽": row["Продажи_calc"],
                "ДРР, %": row["ДРР, %"],
                "В корзину": row["Корзина_calc"],
                "Клик→корзина, %": row["Клик→корзина, %"],
                "Клик→заказ, %": row["Клик→заказ, %"],
                "Корзина→заказ, %": row["Корзина→заказ, %"],
                "Средний чек, ₽": row["Средний чек, ₽"],
                "Рекомендованное действие": action,
                "Причины": reasons,
                "Комментарий": comment,
                "Похожая неактивная кампания": matched_name,
                "Историческое замечание": matched_comment,
            }
        )

    result_df = pd.DataFrame(results)
    action_order = {
        "Масштабировать": 0,
        "Оставить и наблюдать": 1,
        "Снизить CPC / доработать карточку": 2,
        "Поставить на паузу и пересобрать": 3,
        "Поставить на паузу": 4,
    }
    result_df["__sort"] = result_df["Рекомендованное действие"].map(action_order).fillna(99)
    result_df = result_df.sort_values(["__sort", "ДРР, %", "Продажи, ₽"], ascending=[True, True, False]).drop(columns="__sort")

    if output_prefix is None:
        output_prefix = file_path.with_suffix("")
    xlsx_out = output_prefix.parent / f"{output_prefix.name}_active_audit.xlsx"
    txt_out = output_prefix.parent / f"{output_prefix.name}_active_audit.txt"

    with pd.ExcelWriter(xlsx_out, engine="openpyxl") as writer:
        result_df.to_excel(writer, sheet_name="Активные кампании", index=False)
        summary_df = pd.DataFrame({"Сводка": build_summary_text(active_df, target_drr).splitlines()})
        summary_df.to_excel(writer, sheet_name="Сводка", index=False)

        ws = writer.sheets["Активные кампании"]
        widths = {
            "A": 14, "B": 42, "C": 12, "D": 18, "E": 14, "F": 14, "G": 12,
            "H": 10, "I": 14, "J": 16, "K": 12, "L": 14, "M": 10, "N": 11,
            "O": 16, "P": 14, "Q": 16, "R": 15, "S": 28, "T": 55, "U": 80,
            "V": 40, "W": 55,
        }
        for col, width in widths.items():
            ws.column_dimensions[col].width = width
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

    lines = [build_summary_text(active_df, target_drr), "", "Рекомендации по активным кампаниям:"]
    for _, row in result_df.iterrows():
        lines.extend([
            "",
            f"- {row['Название кампании']}",
            f"  Действие: {row['Рекомендованное действие']}",
            f"  Комментарий: {row['Комментарий']}",
        ])
        if row["Историческое замечание"]:
            lines.append(f"  Исторический ориентир: {row['Историческое замечание']}")

    txt_out.write_text("\n".join(lines), encoding="utf-8")
    return xlsx_out, txt_out


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Оценка активных рекламных кампаний Ozon CPC по выгруженному отчёту."
    )
    parser.add_argument("file", nargs="?", help="Путь к Excel-отчёту Ozon")
    parser.add_argument(
        "--target-drr",
        type=float,
        default=0.30,
        help="Целевой ДРР в долях. Например, 0.30 = 30%%",
    )
    parser.add_argument(
        "--output-prefix",
        help="Префикс выходных файлов без расширения. По умолчанию рядом с отчётом.",
    )
    parser.add_argument(
        "--no-inactive",
        action="store_true",
        help="Не учитывать неактивные кампании при расчёте бенчмарков.",
    )
    parser.add_argument(
        "--only-on-off",
        action="store_true",
        help="Только «оставить» / «выключить» без смены стратегии и CPC.",
    )
    parser.add_argument(
        "--gui",
        action="store_true",
        help="Запустить графический интерфейс (по умолчанию при запуске без аргументов).",
    )
    parser.add_argument(
        "--cli",
        action="store_true",
        help="Режим командной строки (диалог выбора файла вместо GUI).",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()

    if args.gui or (not args.file and not args.cli):
        initial = Path(args.file) if args.file else None
        return run_gui(initial_file=initial)

    file_path = Path(args.file) if args.file else pick_file_dialog()
    if not file_path:
        print("Файл не выбран.")
        return 1
    if not file_path.exists():
        print(f"Файл не найден: {file_path}")
        return 1

    try:
        output_prefix = Path(args.output_prefix) if args.output_prefix else None
        xlsx_out, txt_out = analyze(
            file_path,
            target_drr=args.target_drr,
            output_prefix=output_prefix,
            include_inactive=not args.no_inactive,
            suggest_strategy=not args.only_on_off,
        )
    except Exception as exc:
        print(f"Ошибка: {exc}")
        return 2

    print(f"Готово. Excel: {xlsx_out}")
    print(f"Готово. TXT:   {txt_out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

# --- IMPROVEMENTS: идеи для развития скрипта ---
# 1. Выбор папки для сохранения отчётов (по умолчанию рядом с исходным файлом).
# 2. Экспорт в CSV помимо Excel/TXT — для дальнейшей обработки в BI.
# 3. Фильтрация по категориям/типам кампаний (поисковая/каталог), если Ozon добавляет столбец.
# 4. Настраиваемые пороги (ДРР, CTR, CPC) в GUI вместо фиксированного target_drr.
# 5. Сравнение аудитов по периодам: загрузить два отчёта и показать дельту по ДРР, CPC.
# 6. Подсветка «горячих» кампаний: быстрый рост расхода при падающем ДРР.
# 7. Прогноз бюджета на следующий период по тренду.
# 8. Интеграция с Ozon Seller API для автоматической загрузки отчётов (нужен API-ключ).