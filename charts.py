"""
노션 페이지에 붙일 정적 차트 이미지를 만드는 유틸리티.

노션 API는 네이티브 차트 블록이 없어서, matplotlib으로 PNG를 그려
image 블록으로 삽입한다 (dataviz 스킬의 색상/마크 원칙을 따름 —
얇은 선, 단일 계열은 범례 없이, sequential/categorical 구분).
"""

import os

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import matplotlib.font_manager as fm

# 실행 환경에 한글 폰트가 없으면 텍스트가 빈 박스로 깨지므로, 시스템 폰트에 의존하지
# 않고 저장소에 폰트 파일을 함께 배포해서 매 실행마다 명시적으로 등록한다.
_FONT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "fonts")
_KOREAN_FONT_FAMILY = "DejaVu Sans"
for _fname in ("NanumGothic.ttf", "NanumGothicBold.ttf"):
    _fpath = os.path.join(_FONT_DIR, _fname)
    if os.path.exists(_fpath):
        fm.fontManager.addfont(_fpath)
        _KOREAN_FONT_FAMILY = fm.FontProperties(fname=_fpath).get_name()
plt.rcParams["font.family"] = _KOREAN_FONT_FAMILY
plt.rcParams["axes.unicode_minus"] = False

# dataviz 스킬 reference palette (light mode) 값
SURFACE = "#fcfcfb"
TEXT_PRIMARY = "#0b0b0b"
TEXT_SECONDARY = "#52514e"
SERIES_BLUE = "#2a78d6"
SERIES_ORANGE = "#eb6834"
SERIES_AQUA = "#1baf7a"


def _base_style(ax, fig):
    fig.patch.set_facecolor(SURFACE)
    ax.set_facecolor(SURFACE)
    for spine in ["top", "right", "left"]:
        ax.spines[spine].set_visible(False)
    ax.spines["bottom"].set_color(TEXT_SECONDARY)
    ax.spines["bottom"].set_alpha(0.3)
    ax.tick_params(colors=TEXT_SECONDARY, labelsize=9)
    ax.yaxis.grid(True, color=TEXT_SECONDARY, alpha=0.12, linewidth=0.8)
    ax.set_axisbelow(True)


def line_chart(dates, series, title, out_path, ylabel=""):
    """series: {"label": [values...]} — 단일 계열이면 범례 없이, 2개 이상이면 범례 표시."""
    fig, ax = plt.subplots(figsize=(7, 3), dpi=150)
    _base_style(ax, fig)

    colors = [SERIES_BLUE, SERIES_ORANGE, SERIES_AQUA]
    for i, (label, values) in enumerate(series.items()):
        ax.plot(
            dates, values,
            color=colors[i % len(colors)], linewidth=2, solid_capstyle="round",
            marker="o", markersize=4, label=label,
        )

    ax.set_title(title, color=TEXT_PRIMARY, fontsize=12, fontweight="bold", loc="left", pad=12)
    if ylabel:
        ax.set_ylabel(ylabel, color=TEXT_SECONDARY, fontsize=9)
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%m/%d"))
    fig.autofmt_xdate(rotation=0, ha="center")

    if len(series) > 1:
        legend = ax.legend(loc="upper left", frameon=False, fontsize=9, labelcolor=TEXT_SECONDARY)

    fig.tight_layout()
    fig.savefig(out_path, facecolor=SURFACE)
    plt.close(fig)
    return out_path


def bar_chart(labels, values, title, out_path, ylabel="", color=SERIES_BLUE):
    fig, ax = plt.subplots(figsize=(7, 3), dpi=150)
    _base_style(ax, fig)
    ax.xaxis.grid(False)

    bars = ax.bar(labels, values, color=color, width=0.6)
    for bar in bars:
        h = bar.get_height()
        ax.annotate(
            f"{h:,.0f}", (bar.get_x() + bar.get_width() / 2, h),
            textcoords="offset points", xytext=(0, 3), ha="center",
            fontsize=8, color=TEXT_PRIMARY,
        )

    ax.set_title(title, color=TEXT_PRIMARY, fontsize=12, fontweight="bold", loc="left", pad=12)
    if ylabel:
        ax.set_ylabel(ylabel, color=TEXT_SECONDARY, fontsize=9)
    ax.tick_params(axis="x", colors=TEXT_PRIMARY, labelsize=9)

    fig.tight_layout()
    fig.savefig(out_path, facecolor=SURFACE)
    plt.close(fig)
    return out_path
