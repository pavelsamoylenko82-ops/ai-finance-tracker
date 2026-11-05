import io
from typing import Any, Dict, List
import matplotlib
matplotlib.use("Agg")
from matplotlib.figure import Figure


def render_minimalist_stats_chart(
    year_month: str,
    income_byn: float,
    expense_byn: float,
    net_byn: float,
    expense_categories: List[Dict[str, Any]],
    currency: str = "BYN",
) -> bytes:
    cats = expense_categories[:8]
    rows_count = max(len(cats), 1)
    fig_height = max(3.6, 1.8 + rows_count * 0.55)

    fig = Figure(figsize=(8.2, fig_height), dpi=150, facecolor="#0F172A")
    ax = fig.add_subplot(111)
    ax.set_facecolor("#0F172A")

    net_sign = "+" if net_byn >= 0 else ""
    header_title = f"ФИНАНСОВЫЙ ОТЧЕТ  •  {year_month}"
    summary_line = (
        f"Доход: +{income_byn:.2f} {currency}    |    "
        f"Расход: -{expense_byn:.2f} {currency}    |    "
        f"Сальдо: {net_sign}{net_byn:.2f} {currency}"
    )

    fig.text(
        0.04, 0.92, header_title,
        color="#F8FAFC", fontsize=13, fontweight="bold", family="DejaVu Sans"
    )
    fig.text(
        0.04, 0.84, summary_line,
        color="#94A3B8", fontsize=10, family="DejaVu Sans"
    )

    if not cats or expense_byn <= 0:
        ax.text(
            0.5, 0.45, "Расходов в этом месяце нет",
            color="#64748B", fontsize=12, ha="center", va="center",
            transform=ax.transAxes
        )
        ax.axis("off")
    else:
        labels = [c["category"] for c in reversed(cats)]
        values = [float(c["total"]) for c in reversed(cats)]
        max_val = max(values) if values else 1.0

        y_pos = range(len(labels))
        ax.barh(
            y_pos,
            [max_val * 1.38] * len(labels),
            height=0.56,
            color="#1E293B",
            edgecolor="none",
            zorder=1,
        )
        bars = ax.barh(
            y_pos,
            values,
            height=0.56,
            color="#38BDF8",
            edgecolor="none",
            zorder=2,
        )

        ax.set_yticks(list(y_pos))
        ax.set_yticklabels(labels, color="#E2E8F0", fontsize=10, family="DejaVu Sans")
        ax.set_xlim(0, max_val * 1.42)
        ax.xaxis.set_visible(False)

        for spine in ax.spines.values():
            spine.set_visible(False)
        ax.tick_params(axis="y", length=0, pad=8)

        for bar, val in zip(bars, values):
            pct = (val / expense_byn * 100.0) if expense_byn > 0 else 0.0
            label_text = f" {val:.2f} {currency}  ({pct:.0f}%)"
            ax.text(
                val + (max_val * 0.02),
                bar.get_y() + bar.get_height() / 2.0,
                label_text,
                va="center",
                ha="left",
                color="#F8FAFC",
                fontsize=9.5,
                fontweight="bold",
                family="DejaVu Sans",
                zorder=3,
            )

        fig.subplots_adjust(left=0.28, right=0.96, top=0.76, bottom=0.08)

    buf = io.BytesIO()
    fig.savefig(buf, format="png", facecolor=fig.get_facecolor())
    fig.clf()
    buf.seek(0)
    return buf.getvalue()