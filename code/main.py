# -*- coding: utf-8 -*-
"""
================================================================================
 校大学生数学建模竞赛  参赛源程序
 题目：我国省域数字经济发展水平的综合评价、障碍诊断与趋势预测
--------------------------------------------------------------------------------
 程序结构：
   问题一  数据获取与预处理 + 熵权法赋权 + TOPSIS 综合评价 + 系统聚类梯队划分
   问题二  障碍度模型：识别制约各省数字经济发展的主要障碍因子
   问题三  灰色 GM(1,1) 模型：全国数字经济总体水平趋势预测与精度检验
   结果输出 results/ 目录（Excel 结果表）；图形输出 figures/ 目录（png）
--------------------------------------------------------------------------------
 运行环境：Python 3.10+，依赖 numpy / pandas / matplotlib / scipy / openpyxl
 运行方式：在本文件所在 code 目录下执行  python main.py
================================================================================
 【数据说明】
 竞赛原题数据来源于《中国统计年鉴》《中国信息产业年鉴》、工业和信息化部统计
 公报及北京大学数字普惠金融指数。为保证本程序“开箱即可运行、结果可复现”，
 下面 build_dataset() 按各省真实发展梯度构造了一套等口径的示例数据（固定随机
 种子）。正式参赛时，只需把 build_dataset() 返回的两张表替换为真实统计数据，
 后续全部模型代码无需任何修改即可复算。
================================================================================
"""

import os
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")  # 无界面环境下直接保存图片，可正常弹窗时可注释掉本行
import matplotlib.pyplot as plt
from scipy.cluster.hierarchy import linkage, fcluster
from scipy.spatial.distance import pdist

# -----------------------------------------------------------------------------
# 0. 全局设置：中文显示、随机种子、输出路径
# -----------------------------------------------------------------------------
plt.rcParams["font.sans-serif"] = ["Microsoft YaHei", "SimHei", "Arial Unicode MS"]
plt.rcParams["axes.unicode_minus"] = False
plt.rcParams["figure.dpi"] = 120
plt.rcParams["savefig.dpi"] = 300
plt.rcParams["font.size"] = 11

RNG = np.random.default_rng(20240520)  # 固定随机种子，保证结果可复现

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
FIG_DIR = os.path.normpath(os.path.join(BASE_DIR, "..", "figures"))
RES_DIR = os.path.normpath(os.path.join(BASE_DIR, "..", "results"))
os.makedirs(FIG_DIR, exist_ok=True)
os.makedirs(RES_DIR, exist_ok=True)


# =============================================================================
# 一、数据构建（参赛时替换为真实年鉴数据；此处构造等口径示例数据）
# =============================================================================
def build_dataset():
    """返回省域截面数据 raw_df 与全国时序数据 nation_ts。"""

    # 31 个省级行政区及其“潜在综合发展水平” latent（取值 0~1，依据现实梯度给定）
    latent_map = {
        "北京": 0.92, "上海": 0.90, "广东": 0.88, "江苏": 0.86, "浙江": 0.84,
        "福建": 0.70, "山东": 0.69, "天津": 0.66, "湖北": 0.60, "重庆": 0.58,
        "安徽": 0.55, "四川": 0.54, "陕西": 0.52, "湖南": 0.50, "河南": 0.48,
        "辽宁": 0.46, "江西": 0.45, "河北": 0.43, "海南": 0.42, "山西": 0.38,
        "内蒙古": 0.36, "吉林": 0.34, "黑龙江": 0.33, "广西": 0.31, "宁夏": 0.30,
        "新疆": 0.28, "贵州": 0.27, "云南": 0.26, "甘肃": 0.22, "青海": 0.20,
        "西藏": 0.14,
    }
    provinces = list(latent_map.keys())
    latent = np.array(list(latent_map.values()))

    def scale_ind(lo, hi, power, cv, decimals):
        """由 latent 生成一个正向指标：幂函数拉开省际差距并叠加固定噪声。"""
        base = lo + (hi - lo) * latent ** power
        noise = RNG.normal(1.0, cv, size=latent.size)
        val = np.clip(base * noise, lo, hi)
        return np.round(val, decimals)

    data = pd.DataFrame({"省份": provinces})

    # —— 维度 A 数字基础设施 ——
    data["X1 每百人移动电话拥有量(部)"] = scale_ind(95, 176, 1.0, 0.025, 1)     # 比率型
    data["X2 单位国土面积光缆长度(km/km²)"] = scale_ind(0.05, 1.20, 1.3, 0.05, 3)
    data["X3 5G基站数(万个)"] = scale_ind(0.8, 24.0, 1.5, 0.06, 2)             # 规模型
    # —— 维度 B 数字产业化 ——
    data["X4 软件业务收入(亿元)"] = scale_ind(30, 18000, 1.8, 0.06, 0)
    data["X5 电信业务总量(亿元)"] = scale_ind(200, 1800, 1.4, 0.05, 0)
    # —— 维度 C 产业数字化 ——
    data["X6 企业电子商务销售额(亿元)"] = scale_ind(200, 9000, 1.7, 0.06, 0)
    data["X7 关键工序数控化率(%)"] = scale_ind(38, 68, 1.0, 0.02, 1)
    # —— 维度 D 数字创新能力 ——
    data["X8 R&D经费投入强度(%)"] = scale_ind(0.6, 6.8, 1.1, 0.03, 2)
    data["X9 信息技术领域专利授权(万件)"] = scale_ind(0.05, 12.0, 1.7, 0.06, 2)
    # —— 维度 E 数字普惠 ——
    data["X10 数字普惠金融指数"] = scale_ind(280, 470, 1.0, 0.015, 1)
    # X11 为【负向指标】：城乡数字接入鸿沟指数，越发达数值越小
    x11 = 95 - 80 * latent ** 0.8
    x11 = np.clip(x11 * RNG.normal(1.0, 0.03, latent.size), 5, 100)
    data["X11 城乡数字接入鸿沟指数(负向)"] = np.round(x11, 1)

    # 全国数字经济综合发展指数时间序列（2014—2023，近似指数平滑增长 + 微小扰动）
    years = np.arange(2014, 2024)
    t = np.arange(years.size)
    nation_idx = 0.168 * (1.118 ** t) * RNG.normal(1.0, 0.004, years.size)
    nation_ts = pd.DataFrame({"年份": years, "全国数字经济综合发展指数": np.round(nation_idx, 4)})

    return data, nation_ts


# 指标元信息：编号、所属一级维度、方向（+ 正向 / - 负向）
INDICATOR_META = [
    ("X1",  "每百人移动电话拥有量",        "数字基础设施", "+"),
    ("X2",  "单位国土面积光缆长度",        "数字基础设施", "+"),
    ("X3",  "5G基站数",                    "数字基础设施", "+"),
    ("X4",  "软件业务收入",                "数字产业化",   "+"),
    ("X5",  "电信业务总量",                "数字产业化",   "+"),
    ("X6",  "企业电子商务销售额",          "产业数字化",   "+"),
    ("X7",  "关键工序数控化率",            "产业数字化",   "+"),
    ("X8",  "R&D经费投入强度",             "数字创新能力", "+"),
    ("X9",  "信息技术领域专利授权",        "数字创新能力", "+"),
    ("X10", "数字普惠金融指数",            "数字普惠",     "+"),
    ("X11", "城乡数字接入鸿沟指数",        "数字普惠",     "-"),
]
DIMENSIONS = ["数字基础设施", "数字产业化", "产业数字化", "数字创新能力", "数字普惠"]


# =============================================================================
# 问题一（1）：指标正向化与无量纲化（极差标准化）
# =============================================================================
def normalize(df, value_cols, directions):
    """
    极差标准化，结果落于 [0,1]：
      正向指标  z = (x - min) / (max - min)
      负向指标  z = (max - x) / (max - min)
    """
    z = pd.DataFrame(index=df.index)
    for col, d in zip(value_cols, directions):
        x = df[col].to_numpy(dtype=float)
        rng_ = x.max() - x.min()
        if d == "+":
            z[col] = (x - x.min()) / rng_
        else:
            z[col] = (x.max() - x) / rng_
    return z


# =============================================================================
# 问题一（2）：熵权法确定客观权重
# =============================================================================
def entropy_weight(Z):
    """输入标准化矩阵 Z（行=评价对象，列=指标），返回 (权重 w, 熵值 e, 差异系数 g)。"""
    n = Z.shape[0]
    # 计算比重 p_ij，列内归一
    col_sum = Z.sum(axis=0)
    with np.errstate(divide="ignore", invalid="ignore"):
        P = Z / col_sum
        plnp = np.where(P > 0, P * np.log(P), 0.0)   # 约定 0*ln0 = 0
    e = -plnp.sum(axis=0) / np.log(n)               # 信息熵
    g = 1.0 - e                                     # 差异系数（效用值）
    w = g / g.sum()                                 # 客观权重
    return w, e, g


# =============================================================================
# 问题一（3）：TOPSIS 优劣解距离法
# =============================================================================
def topsis(Z, w):
    """输入标准化矩阵 Z 与权重 w，返回 (正理想距离D+, 负理想距离D-, 贴近度C)。"""
    V = Z.to_numpy(dtype=float) * w                 # 加权标准化矩阵
    v_pos = V.max(axis=0)                           # 正理想解
    v_neg = V.min(axis=0)                           # 负理想解
    d_pos = np.sqrt(((V - v_pos) ** 2).sum(axis=1))
    d_neg = np.sqrt(((V - v_neg) ** 2).sum(axis=1))
    closeness = d_neg / (d_pos + d_neg)             # 相对贴近度 C ∈ [0,1]
    return d_pos, d_neg, closeness


# =============================================================================
# 问题一（4）：系统聚类（Ward 法）划分三个发展梯队
# =============================================================================
def cluster_tiers(closeness, n_tiers=3):
    """按综合贴近度做 Ward 系统聚类并切为 n_tiers 类，按类内均值由高到低命名梯队。"""
    X = closeness.reshape(-1, 1)
    Z_link = linkage(pdist(X), method="ward")
    labels = fcluster(Z_link, t=n_tiers, criterion="maxclust")
    # 依据各类平均贴近度从高到低映射为 第Ⅰ/Ⅱ/Ⅲ 梯队
    order = sorted(np.unique(labels),
                   key=lambda lab: -closeness[labels == lab].mean())
    tier_name = {lab: f"第{['Ⅰ', 'Ⅱ', 'Ⅲ', 'Ⅳ'][i]}梯队" for i, lab in enumerate(order)}
    return np.array([tier_name[lab] for lab in labels]), Z_link


# =============================================================================
# 问题二：障碍度模型
# =============================================================================
def obstacle_degree(Z, w):
    """
    指标偏离度  I_j = 1 - z_ij
    因子贡献度  F_j = w_j
    障碍度      O_ij = (F_j * I_ij) / Σ_j(F_j * I_ij) ×100%
    """
    I = 1.0 - Z.to_numpy(dtype=float)
    FI = w * I
    O = FI / FI.sum(axis=1, keepdims=True) * 100.0
    return pd.DataFrame(O, columns=Z.columns, index=Z.index)


# =============================================================================
# 问题三：灰色 GM(1,1) 预测模型
# =============================================================================
def gm11(x0, n_forecast=5):
    """
    输入原始非负序列 x0 与向后预测步数，返回字典：
    发展系数a、灰作用量b、拟合值、预测值、相对误差、后验差比值C、小误差概率P、精度等级
    """
    x0 = np.asarray(x0, dtype=float)
    n = x0.size
    x1 = np.cumsum(x0)                                  # 一次累加生成 1-AGO
    z1 = 0.5 * (x1[1:] + x1[:-1])                       # 紧邻均值生成序列
    B = np.column_stack([-z1, np.ones(n - 1)])
    Y = x0[1:]
    a_hat, b_hat = np.linalg.lstsq(B, Y, rcond=None)[0] # 最小二乘估计 [a,b]
    a, b = a_hat, b_hat

    # 时间响应函数（预测序列下标 k 从 0 开始）
    def x1_hat(k):
        return (x0[0] - b / a) * np.exp(-a * k) + b / a

    k_all = np.arange(0, n + n_forecast)
    x1_fit = x1_hat(k_all)
    x0_hat = np.empty_like(x1_fit)
    x0_hat[0] = x1_fit[0]
    x0_hat[1:] = np.diff(x1_fit)                        # 累减还原

    fit = x0_hat[:n]                                    # 历史拟合
    forecast = x0_hat[n:]                               # 未来预测

    # —— 精度检验：残差、相对误差、后验差比 C、小误差概率 P ——
    residual = x0 - fit
    rel_err = np.abs(residual / x0) * 100
    s1 = x0.std(ddof=1)                                 # 原始序列标准差
    s2 = residual.std(ddof=1)                           # 残差序列标准差
    C = s2 / s1                                         # 后验差比值
    P = np.mean(np.abs(residual - residual.mean()) < 0.6745 * s1)  # 小误差概率
    map_ = rel_err.mean()                               # 平均相对误差

    if C < 0.35 and P > 0.95:
        grade = "一级（优）"
    elif C < 0.50 and P > 0.80:
        grade = "二级（合格）"
    elif C < 0.65 and P > 0.70:
        grade = "三级（勉强）"
    else:
        grade = "四级（不合格）"

    return dict(a=a, b=b, fit=fit, forecast=forecast,
                residual=residual, rel_err=rel_err, C=C, P=P,
                mape=map_, grade=grade)


# =============================================================================
# 可视化
# =============================================================================
TIER_COLOR = {"第Ⅰ梯队": "#c0392b", "第Ⅱ梯队": "#e67e22", "第Ⅲ梯队": "#95a5a6"}


def fig_weights(codes, names, w):
    order = np.argsort(w)
    fig, ax = plt.subplots(figsize=(8.5, 6))
    labels = [f"{c} {n}" for c, n in zip(np.array(codes)[order], np.array(names)[order])]
    bars = ax.barh(labels, w[order] * 100, color="#2e86c1", edgecolor="white")
    for b, val in zip(bars, w[order] * 100):
        ax.text(val + 0.1, b.get_y() + b.get_height() / 2, f"{val:.2f}%",
                va="center", fontsize=9)
    ax.set_xlabel("熵权法客观权重（%）")
    ax.set_title("数字经济发展水平评价指标熵权权重")
    ax.grid(axis="x", ls="--", alpha=0.4)
    fig.tight_layout()
    fig.savefig(os.path.join(FIG_DIR, "fig1_熵权权重.png"))
    plt.close(fig)


def fig_topsis_rank(result_df):
    df = result_df.sort_values("贴近度C", ascending=True)
    colors = df["梯队"].map(TIER_COLOR)
    fig, ax = plt.subplots(figsize=(9, 11))
    bars = ax.barh(df["省份"], df["贴近度C"], color=colors, edgecolor="white")
    for b, val in zip(bars, df["贴近度C"]):
        ax.text(val + 0.004, b.get_y() + b.get_height() / 2, f"{val:.3f}",
                va="center", fontsize=8)
    ax.set_xlabel("TOPSIS 相对贴近度 C")
    ax.set_title("31省数字经济发展水平 TOPSIS 贴近度及梯队划分")
    handles = [plt.Rectangle((0, 0), 1, 1, color=c) for c in TIER_COLOR.values()]
    ax.legend(handles, TIER_COLOR.keys(), loc="lower right")
    ax.grid(axis="x", ls="--", alpha=0.4)
    fig.tight_layout()
    fig.savefig(os.path.join(FIG_DIR, "fig2_TOPSIS贴近度排名.png"))
    plt.close(fig)


def fig_dim_radar(Z_norm_dim, rep_provinces):
    """rep_provinces: 三个梯队代表省份，在五个维度上的平均得分雷达图。"""
    angles = np.linspace(0, 2 * np.pi, len(DIMENSIONS), endpoint=False).tolist()
    angles += angles[:1]
    fig, ax = plt.subplots(figsize=(7, 7), subplot_kw=dict(polar=True))
    palette = ["#c0392b", "#2980b9", "#7f8c8d"]
    for (tier, prov), color in zip(rep_provinces.items(), palette):
        vals = Z_norm_dim.loc[prov].tolist()
        vals += vals[:1]
        ax.plot(angles, vals, color=color, lw=2, label=f"{prov}（{tier}）")
        ax.fill(angles, vals, color=color, alpha=0.12)
    ax.set_xticks(angles[:-1])
    ax.set_xticklabels(DIMENSIONS, fontsize=10)
    ax.set_ylim(0, 1)
    ax.set_title("不同梯队代表省份五维度得分雷达图", pad=20)
    ax.legend(loc="upper right", bbox_to_anchor=(1.25, 1.1))
    fig.tight_layout()
    fig.savefig(os.path.join(FIG_DIR, "fig3_梯队代表雷达图.png"))
    plt.close(fig)


def fig_obstacle(obs_indicator_mean, codes, names):
    """obs_indicator_mean：以原始列名为索引的全国平均障碍度 Series（已按降序）。"""
    code2name = dict(zip(codes, names))
    s = obs_indicator_mean.sort_values(ascending=True)      # 横向条形从小到大
    labels = [f"{col.split()[0]} {code2name[col.split()[0]]}" for col in s.index]
    fig, ax = plt.subplots(figsize=(8.5, 6))
    bars = ax.barh(labels, s.to_numpy(), color="#e67e22", edgecolor="white")
    for b, v in zip(bars, s.to_numpy()):
        ax.text(v + 0.1, b.get_y() + b.get_height() / 2, f"{v:.2f}%",
                va="center", fontsize=9)
    ax.set_xlabel("平均障碍度（%）")
    ax.set_title("制约数字经济发展的主要障碍因子（全国平均）")
    ax.grid(axis="x", ls="--", alpha=0.4)
    fig.tight_layout()
    fig.savefig(os.path.join(FIG_DIR, "fig4_障碍因子排序.png"))
    plt.close(fig)


def fig_gm(years, x0, gm, hist_years, fore_years):
    fig, ax = plt.subplots(figsize=(9.2, 5.4))
    ax.plot(hist_years, x0, "o-", color="#2c3e50", lw=2, label="实际值")
    ax.plot(hist_years, gm["fit"], "--", color="#2980b9", lw=2, label="GM(1,1)拟合值")
    ax.plot(fore_years, gm["forecast"], "s--", color="#c0392b", lw=2, label="预测值")
    ax.axvline(hist_years[-1], color="gray", ls=":", lw=1.5)
    ax.text(hist_years[-1] + 0.1, ax.get_ylim()[1] * 0.96, "预测起点", color="gray", fontsize=9)
    for x, y in zip(fore_years, gm["forecast"]):
        ax.annotate(f"{y:.3f}", (x, y), textcoords="offset points", xytext=(0, 8),
                    ha="center", fontsize=8, color="#c0392b")
    ax.set_xlabel("年份")
    ax.set_ylabel("全国数字经济综合发展指数")
    ax.set_title("全国数字经济发展指数 GM(1,1) 拟合与预测（2024—2028）")
    ax.legend()
    ax.grid(ls="--", alpha=0.4)
    fig.tight_layout()
    fig.savefig(os.path.join(FIG_DIR, "fig5_GM11预测.png"))
    plt.close(fig)


def fig_tier_pie(result_df):
    counts = result_df["梯队"].value_counts().reindex(["第Ⅰ梯队", "第Ⅱ梯队", "第Ⅲ梯队"])
    fig, ax = plt.subplots(figsize=(6.4, 6))
    ax.pie(counts, labels=[f"{k}\n{v}个" for k, v in counts.items()],
           autopct="%1.1f%%", colors=[TIER_COLOR[k] for k in counts.index],
           startangle=90, wedgeprops=dict(edgecolor="white", linewidth=2),
           textprops={"fontsize": 11})
    ax.set_title("三大发展梯队省份数量构成")
    fig.tight_layout()
    fig.savefig(os.path.join(FIG_DIR, "fig6_梯队构成.png"))
    plt.close(fig)


# =============================================================================
# 主流程
# =============================================================================
def main():
    pd.set_option("display.unicode.east_asian_width", True)
    pd.set_option("display.width", 200)

    # ---------- 读取/构建数据 ----------
    raw_df, nation_ts = build_dataset()
    value_cols = [c for c in raw_df.columns if c != "省份"]
    codes = [m[0] for m in INDICATOR_META]
    names = [m[1] for m in INDICATOR_META]
    directions = [m[3] for m in INDICATOR_META]
    dim_of = {m[0]: m[2] for m in INDICATOR_META}

    line = "=" * 78
    print(line); print("问题一：熵权—TOPSIS 综合评价"); print(line)

    # ---------- 标准化 ----------
    Z = normalize(raw_df, value_cols, directions)
    Z.insert(0, "省份", raw_df["省份"])
    Zv = Z[value_cols]

    # ---------- 熵权法 ----------
    w, e, g = entropy_weight(Zv)
    weight_df = pd.DataFrame({
        "指标编号": codes, "指标名称": names,
        "所属维度": [dim_of[c] for c in codes], "方向": directions,
        "信息熵e": np.round(e, 4), "差异系数g": np.round(g, 4),
        "权重w(%)": np.round(w * 100, 3),
    })
    dim_w = weight_df.groupby("所属维度")["权重w(%)"].sum().reindex(DIMENSIONS)
    print("\n[表1] 各指标信息熵、差异系数与客观权重：")
    print(weight_df.to_string(index=False))
    print("\n[表2] 五个一级维度权重合计（%）：")
    print(dim_w.round(3).to_string())

    # ---------- TOPSIS ----------
    d_pos, d_neg, closeness = topsis(Zv, w)
    result = raw_df[["省份"]].copy()
    result["正理想距离D+"] = np.round(d_pos, 4)
    result["负理想距离D-"] = np.round(d_neg, 4)
    result["贴近度C"] = np.round(closeness, 4)
    result = result.sort_values("贴近度C", ascending=False).reset_index(drop=True)
    result["排名"] = np.arange(1, result.shape[0] + 1)

    # ---------- 聚类梯队 ----------
    tier, _ = cluster_tiers(closeness)
    tier_series = pd.Series(tier, index=raw_df.index)
    result["梯队"] = result["省份"].map(dict(zip(raw_df["省份"], tier_series)))
    print("\n[表3] 31省 TOPSIS 综合评价结果（按贴近度降序）：")
    print(result[["排名", "省份", "贴近度C", "梯队"]].to_string(index=False))
    print("\n[表4] 各梯队包含省份：")
    for tname in ["第Ⅰ梯队", "第Ⅱ梯队", "第Ⅲ梯队"]:
        ps = result.loc[result["梯队"] == tname, "省份"].tolist()
        print(f"  {tname}（{len(ps)}个，均值C="
              f"{result.loc[result['梯队']==tname,'贴近度C'].mean():.3f}）：{ '、'.join(ps) }")

    # 五维度标准化得分（维度内指标等权平均），供雷达图
    dim_score = pd.DataFrame({"省份": raw_df["省份"]})
    for dim in DIMENSIONS:
        cols = [value_cols[i] for i, m in enumerate(INDICATOR_META) if m[2] == dim]
        dim_score[dim] = Zv[cols].mean(axis=1)
    dim_score = dim_score.set_index("省份")

    print("\n" + line); print("问题二：障碍度模型诊断"); print(line)
    obs = obstacle_degree(Zv, w)
    obs.insert(0, "省份", raw_df["省份"])
    obs_mean = obs[value_cols].mean(axis=0).sort_values(ascending=False)
    obs_ind_df = pd.DataFrame({
        "指标编号": codes, "指标名称": names,
        "全国平均障碍度(%)": np.round([obs_mean[c] for c in value_cols], 3),
    }).sort_values("全国平均障碍度(%)", ascending=False).reset_index(drop=True)
    print("\n[表5] 指标层障碍度排序（全国平均，前11项）：")
    print(obs_ind_df.to_string(index=False))

    # 维度层障碍度（注意用 to_numpy 避免索引错位对齐）
    obs_dim = pd.DataFrame({"省份": raw_df["省份"]})
    for dim in DIMENSIONS:
        cols = [value_cols[i] for i, m in enumerate(INDICATOR_META) if m[2] == dim]
        obs_dim[dim] = obs[cols].sum(axis=1).to_numpy()
    obs_dim_mean = obs_dim[DIMENSIONS].mean(axis=0).sort_values(ascending=False)
    print("\n[表6] 维度层平均障碍度（%）：")
    print(obs_dim_mean.round(3).to_string())

    # 典型省份障碍度（排名第1、居中、最末）
    top_p = result.iloc[0]["省份"]; mid_p = result.iloc[15]["省份"]; bot_p = result.iloc[-1]["省份"]
    typ_rows = []
    for p in [top_p, mid_p, bot_p]:
        row = obs.loc[obs["省份"] == p, value_cols].iloc[0]
        top3 = row.sort_values(ascending=False).head(3)
        typ_rows.append((p, [(c.split()[0], round(v, 2)) for c, v in top3.items()]))
    print("\n[表7] 典型省份前三位障碍因子：")
    for p, lst in typ_rows:
        print(f"  {p}：" + "，".join([f"{c}({v:.2f}%)" for c, v in lst]))

    print("\n" + line); print("问题三：灰色 GM(1,1) 预测"); print(line)
    years = nation_ts["年份"].to_numpy()
    x0 = nation_ts["全国数字经济综合发展指数"].to_numpy()
    gm = gm11(x0, n_forecast=5)
    fore_years = np.arange(years[-1] + 1, years[-1] + 6)
    fit_df = pd.DataFrame({
        "年份": years, "实际值": x0, "拟合值": np.round(gm["fit"], 4),
        "残差": np.round(gm["residual"], 4),
        "相对误差(%)": np.round(gm["rel_err"], 3),
    })
    print("\n[表8] GM(1,1) 历史拟合与残差检验：")
    print(fit_df.to_string(index=False))
    print(f"\n发展系数 a = {gm['a']:.5f}，灰作用量 b = {gm['b']:.5f}")
    print(f"平均相对误差 MAPE = {gm['mape']:.3f}%，后验差比值 C = {gm['C']:.4f}，"
          f"小误差概率 P = {gm['P']:.3f}，模型精度等级：{gm['grade']}")
    fore_df = pd.DataFrame({"年份": fore_years,
                            "预测指数": np.round(gm["forecast"], 4)})
    print("\n[表9] 2024—2028 年全国数字经济综合发展指数预测：")
    print(fore_df.to_string(index=False))
    cagr = (gm["forecast"][-1] / x0[-1]) ** (1 / 5) - 1
    print(f"\n2023→2028 预测年均增长率 ≈ {cagr*100:.2f}%")

    # ---------- 出图 ----------
    fig_weights(codes, names, w)
    fig_topsis_rank(result)
    # 各梯队取贴近度最接近梯队均值的“中位代表省份”，使雷达图更具可比性
    rep = {}
    for tname in ["第Ⅰ梯队", "第Ⅱ梯队", "第Ⅲ梯队"]:
        sub = result[result["梯队"] == tname]
        mean_c = sub["贴近度C"].mean()
        rep[tname] = sub.iloc[(sub["贴近度C"] - mean_c).abs().argsort()].iloc[0]["省份"]
    print("\n[图3] 各梯队中位代表省份：" + "，".join(f"{k}:{v}" for k, v in rep.items()))
    fig_dim_radar(dim_score, rep)
    fig_obstacle(obs_mean, codes, names)
    fig_gm(years, x0, gm, years, fore_years)
    fig_tier_pie(result)

    # ---------- 结果导出 Excel ----------
    out_xlsx = os.path.join(RES_DIR, "计算结果汇总.xlsx")
    with pd.ExcelWriter(out_xlsx, engine="openpyxl") as writer:
        raw_df.to_excel(writer, sheet_name="原始数据(省域截面)", index=False)
        nation_ts.to_excel(writer, sheet_name="原始数据(全国时序)", index=False)
        weight_df.to_excel(writer, sheet_name="熵权权重", index=False)
        dim_w.rename("维度权重(%)").reset_index().to_excel(
            writer, sheet_name="维度权重", index=False)
        result.to_excel(writer, sheet_name="TOPSIS评价结果", index=False)
        dim_score.round(4).reset_index().to_excel(
            writer, sheet_name="五维度得分", index=False)
        obs_ind_df.to_excel(writer, sheet_name="指标障碍度", index=False)
        obs_dim_mean.rename("维度平均障碍度(%)").round(3).reset_index().to_excel(
            writer, sheet_name="维度障碍度", index=False)
        obs.round(3).to_excel(writer, sheet_name="各省障碍度明细", index=False)
        fit_df.to_excel(writer, sheet_name="GM11拟合检验", index=False)
        fore_df.to_excel(writer, sheet_name="GM11预测", index=False)
        Z.round(4).to_excel(writer, sheet_name="标准化矩阵", index=False)

    print("\n" + line)
    print("全部计算完成！")
    print(f"  结果表：{out_xlsx}（含原始数据、权重、评价、障碍度、GM预测、标准化矩阵等sheet）")
    print(f"  图形目录：{FIG_DIR}（fig1~fig6 共 6 张）")
    print(line)


if __name__ == "__main__":
    main()
