# -*- coding: utf-8 -*-
"""
紙箱実測データのジッタープロット作成スクリプト

必要なライブラリ:
    pip install pandas numpy matplotlib openpyxl

入力: コンタック個装箱.xlsx  (列: ロット / 時刻 / 質量)
出力: 紙箱実測.png
"""

import numpy as np
import pandas as pd
import matplotlib
import matplotlib.pyplot as plt

# ============================================================
# 0. 日本語フォントの設定
#    Windows なら "Yu Gothic" or "MS Gothic"
#    Mac なら "Hiragino Sans"
#    Linux なら "Noto Sans CJK JP" / "IPAexGothic"
# ============================================================
plt.rcParams["font.family"] = "Yu Gothic"      # 環境に合わせて変更
plt.rcParams["axes.unicode_minus"] = False      # マイナス記号の文字化け防止


# ============================================================
# 1. データ読み込み
# ============================================================
df = pd.read_excel("コンタック個装箱.xlsx")
df.columns = ["ロット", "時刻", "質量"]          # 列名を統一
df = df.dropna(subset=["質量"])

# 時刻を "HH:MM" の文字列に整える
df["時刻"] = df["時刻"].astype(str).str[:5]

# ロットの出現順を保持（sortすると順番が狂うため）
order = list(dict.fromkeys(df["ロット"]))


# ============================================================
# 2. 箱ごとの平均を計算
# ============================================================
means = np.array([df.loc[df["ロット"] == lot, "質量"].mean() for lot in order])
rng_g = means.max() - means.min()               # 箱平均レンジ


# ============================================================
# 3. 作図
# ============================================================
fig, ax = plt.subplots(figsize=(9.5, 5.4))

# --- レンジ帯（最重箱の平均〜最軽箱の平均を薄いオレンジで塗る）---
ax.axhspan(means.min(), means.max(), color="#EF9F27", alpha=0.08)

rs = np.random.default_rng(7)                   # 乱数シード固定 → 毎回同じ図になる

for i, lot in enumerate(order):
    w = df.loc[df["ロット"] == lot, "質量"].values

    # --- 実測点（ジッター付き散布）---
    # x座標に ±0.08 の乱数を加えて横に散らす（点の重なり回避）
    jitter = rs.uniform(-0.08, 0.08, len(w))
    ax.plot(np.full(len(w), i) + jitter, w,
            "o", ms=5, alpha=0.55, color="#2a78d6")

    # --- 箱平均（オレンジの横棒）---
    ax.plot([i - 0.24, i + 0.24], [w.mean()] * 2,
            color="#D95926", lw=3)

    # --- 平均値のラベル ---
    ax.text(i, w.mean() + 0.02, f"{w.mean():.3f}",
            ha="center", fontsize=9, color="#D95926", fontweight="bold")

# --- レンジを示す両端矢印 ---
x_arrow = len(order) - 0.55
ax.annotate("", xy=(x_arrow, means.max()), xytext=(x_arrow, means.min()),
            arrowprops=dict(arrowstyle="<->", color="#C00000", lw=1.5))
ax.text(x_arrow - 0.07, means.mean(),
        f"箱平均レンジ\n{rng_g:.3f} g\n＝大函 {rng_g*100:.0f} g",
        ha="right", va="center",
        fontsize=10, color="#C00000", fontweight="bold")

# --- 軸の設定 ---
ax.set_xticks(range(len(order)))
ax.set_xticklabels(
    [f"{lot}\n({df.loc[df['ロット']==lot,'時刻'].iloc[0]})" for lot in order],
    fontsize=9)
ax.set_ylabel("紙箱1枚の重量 (g)")
ax.set_title("空紙箱の実測（5箱×10枚、12:20～13:40の80分内に測定）", fontsize=11.5)
ax.grid(alpha=0.3, axis="y")                    # 横方向のグリッドのみ

plt.tight_layout()
plt.savefig("紙箱実測.png", dpi=150)
print(f"保存しました。箱平均レンジ = {rng_g:.3f} g（大函換算 {rng_g*100:.0f} g）")
