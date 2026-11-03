# -*- coding: utf-8 -*-
"""_data.py 模板 —— 数值的唯一来源。

复制本文件为项目的 `_data.py`，然后：
    1. 只改「① 输入区」——口径常量与原始记录
    2. 运行 `python _data.py`，确认打印出的中间量与判定符合预期
    3. 构建脚本 `import _data as D`，用 D.xxx 引用，**不要自己再算一遍**

本文件自带一组演示数据（HPLC 外标法），直接运行即可看到完整输出。
删掉演示数据、填入真实口径后就是你的项目。

结构
----
    ① 输入区        口径常量 + 原始记录（人改这里）
    ② 统计工具      mean / sd / rsd / 回归 / 格式化（n<2 返回 None）
    ③ 计算模型      按实验类型选用，其余删掉
    ④ 自检打印      把全部中间量、判定、诊断打出来
"""

from __future__ import annotations

import math

# ==================================================================
# ① 输入区
# ==================================================================
DRUG = "示例片"
METHOD = "高效液相色谱法（外标法）"

# ---- 取样方案 ----
N_TABLETS = 20              # 取本品 N 片
V_FLASK = 50.0              # mL   供试品溶液定容体积
W_SAMPLE_G = 0.2742         # g    每份称样量
W_BAR_G = 0.329             # g/片 平均片重
W_THEOR_MG = 25.0           # mg   称样量"约相当于主药"的量（相当量）

# ---- 对照品 ----
C_STD = 0.5058              # mg/mL 对照品溶液浓度
C_THEOR = W_THEOR_MG / V_FLASK        # mg/mL 供试品理论浓度 = 0.500

# ---- 标示量与限度（★ 引药典/课件原文） ----
B_LABEL_MG = 30.0           # mg/片 标示量
CONTENT_LOWER, CONTENT_UPPER = 93.0, 107.0     # %
RSD_MAX = 2.0               # %
R_MIN = 0.999               # 线性相关系数下限

# ---- 数据状态 ----
DATA_IS_PLACEHOLDER = False   # True → 演示数据；False → 已换入实测
DATA_OBS_FILLED = True        # False → 观察结果待填，正文显示「（待填）」

# ---- 原始记录（逐格转录；划改处取改后值，并在表注交代） ----
# 线性：[(进样量 μL, 峰面积), ...]
LINEAR: list[tuple[float, float]] = [
    (4.0, 205600.0), (8.0, 412300.0), (12.0, 622400.0),
    (16.0, 825100.0), (20.0, 1010000.0),
]
# 精密度：同一对照液连续进样 6 针的峰面积
PRECISION_AREAS: list[float] = [512340.0, 519870.0, 508110.0, 521450.0, 515600.0, 517230.0]
# 稳定性：0 / 0.5 / 1 / 2 / 4 h 的峰面积
STABILITY_AREAS: list[float] = [489120.0, 490350.0, 488670.0, 491240.0, 487960.0]
# 重复性 / 含量：平行 6 份供试液的峰面积
SAMPLE_AREAS: list[float] = [487210.0, 491340.0, 488760.0, 490150.0, 489020.0, 490270.0]

# ==================================================================
# ② 统计工具（不要改）
# ==================================================================
def mean(xs):
    xs = list(xs)
    return sum(xs) / len(xs)


def sd(xs):
    """样本标准差（n−1）。★ n < 2 返回 None —— 不编造数值。"""
    xs = list(xs)
    n = len(xs)
    if n < 2:
        return None
    m = mean(xs)
    return math.sqrt(sum((x - m) ** 2 for x in xs) / (n - 1))


def rsd(xs):
    """相对标准偏差 %。n < 2 返回 None。"""
    s = sd(xs)
    return None if s is None else s / mean(xs) * 100.0


def rel_dev(a: float, b: float) -> float:
    """两份平行的相对平均偏差 %。"""
    return abs(a - b) / ((a + b) / 2) * 100.0


def linreg(xs, ys) -> dict:
    """最小二乘：斜率 / 截距 / r / r² / 预测值 / 逐点相对偏差 / 斜率标准误。"""
    xs, ys = list(xs), list(ys)
    n = len(xs)
    mx, my = mean(xs), mean(ys)
    sxx = sum((x - mx) ** 2 for x in xs)
    syy = sum((y - my) ** 2 for y in ys)
    sxy = sum((x - mx) * (y - my) for x, y in zip(xs, ys))
    slope = sxy / sxx
    intercept = my - slope * mx
    pred = [slope * x + intercept for x in xs]
    dev = [(y - p) / p * 100.0 for y, p in zip(ys, pred)]
    sb = math.sqrt(sum((y - (slope * x + intercept)) ** 2 for x, y in zip(xs, ys))
                   / (n - 2) / sxx)
    r = sxy / math.sqrt(sxx * syy)
    return {"n": n, "slope": slope, "intercept": intercept,
            "r": r, "r2": r * r,
            "pred": pred, "dev_pct": dev, "dev_abs_max": max(abs(d) for d in dev),
            "slope_sd": sb}


def fmt(x, n: int = 2) -> str:
    """安全格式化：None → 「—」（单份测定不报 s/RSD 时的显示）。"""
    return "—" if x is None else f"{x:.{n}f}"


def in_range(x: float, lo: float, hi: float) -> bool:
    return lo <= x <= hi


def verdict(ok: bool) -> str:
    return "符合规定" if ok else "不符合规定"


# ==================================================================
# ③ 计算模型（按实验类型保留需要的那个）
# ==================================================================

# --- 模型 A：限量检查  L(%) = V·c / S × 100% ---
def limit_pct(v_ml: float, c_ug_ml: float, s_g: float) -> float:
    """标准液体积 × 浓度（μg/mL）÷ 取样量（g，换算成 μg）。"""
    return v_ml * c_ug_ml / (s_g * 1e6) * 100.0


# --- 模型 B：外标一点法 ---
def precision() -> dict:
    m, s = mean(PRECISION_AREAS), sd(PRECISION_AREAS)
    r = rsd(PRECISION_AREAS)
    return {"n": len(PRECISION_AREAS), "mean": m, "s": s, "rsd": r,
            "a_std": m, "pass": (r is not None and r < RSD_MAX)}


def stability() -> dict:
    r = rsd(STABILITY_AREAS)
    return {"n": len(STABILITY_AREAS), "mean": mean(STABILITY_AREAS),
            "s": sd(STABILITY_AREAS), "rsd": r,
            "pass": (r is not None and r < RSD_MAX)}


def content() -> dict:
    """C样 = C对 × A样/A对；含量% = C样×V ÷ 相当量 × 100%。"""
    a_std = precision()["a_std"]
    a_sam = mean(SAMPLE_AREAS)
    c_sam = C_STD * a_sam / a_std
    m_mg = c_sam * V_FLASK
    contents = [C_STD * (a / a_std) * V_FLASK / W_THEOR_MG * 100.0 for a in SAMPLE_AREAS]
    c_mean = mean(contents)
    s = sd(contents)
    return {"n": len(contents), "a_std": a_std, "a_sam": a_sam,
            "ratio": a_sam / a_std, "c_sam": c_sam, "m_mg": m_mg,
            "mean": c_mean, "s": s,
            "rsd": (None if s is None else s / c_mean * 100.0),
            "per_tablet_mg": c_mean / 100.0 * B_LABEL_MG,
            "pass": in_range(c_mean, CONTENT_LOWER, CONTENT_UPPER)}


def methodology() -> dict:
    """方法学五项汇总（含量与重复性共用数据）。"""
    L = linreg([p[0] for p in LINEAR], [p[1] for p in LINEAR])
    P, S, C = precision(), stability(), content()
    items = [
        {"item": "线性关系", "index": f"r = {L['r']:.4f}", "req": f"r > {R_MIN}",
         "pass": L["r"] > R_MIN},
        {"item": "精密度", "index": f"RSD = {fmt(P['rsd'])}%", "req": f"RSD < {RSD_MAX:.0f}%",
         "pass": P["pass"]},
        {"item": "稳定性", "index": f"RSD = {fmt(S['rsd'])}%", "req": f"RSD < {RSD_MAX:.0f}%",
         "pass": S["pass"]},
        {"item": "重复性", "index": f"RSD = {fmt(C['rsd'])}%", "req": f"RSD < {RSD_MAX:.0f}%",
         "pass": (C["rsd"] is not None and C["rsd"] < RSD_MAX)},
        {"item": "含量测定", "index": f"{C['mean']:.2f}%",
         "req": f"{CONTENT_LOWER:.1f}%~{CONTENT_UPPER:.1f}%", "pass": C["pass"]},
    ]
    return {"linear": L, "items": items,
            "n_pass": sum(1 for i in items if i["pass"]), "n_total": len(items)}


# --- 模型 C：数据一致性诊断（含量异常时必做） ---
def diagnose() -> dict:
    C = content()
    a_expected = C["a_std"] * C_THEOR / C_STD       # 若溶液为理论浓度，应得的峰面积
    return {"ratio_pct": C["ratio"] * 100.0,
            "c_sam": C["c_sam"], "c_theor": C_THEOR,
            "a_expected": a_expected, "a_sam": C["a_sam"],
            "ratio_expected": C["a_sam"] / a_expected}


# ==================================================================
# ④ 自检打印
# ==================================================================
def main() -> int:
    print("=" * 74)
    print(f"{DRUG} —— 数据模型自检（{METHOD}）")
    print("=" * 74)
    print(f"数据状态：{'演示数据' if DATA_IS_PLACEHOLDER else '实测数据'}；"
          f"观察结果：{'已填' if DATA_OBS_FILLED else '（待填）'}")
    print(f"取样 {N_TABLETS} 片；定容 {V_FLASK:.0f} mL；称样 {W_SAMPLE_G} g；"
          f"平均片重 {W_BAR_G} g/片；相当量 {W_THEOR_MG} mg")
    print(f"C对 = {C_STD:.4f} mg/mL；C理 = {C_THEOR:.3f} mg/mL；"
          f"标示量 {B_LABEL_MG} mg/片；限度 {CONTENT_LOWER:.1f}%~{CONTENT_UPPER:.1f}%")

    M = methodology()
    L = M["linear"]
    print("\n--- 线性关系 ---")
    for (v, a), p, d in zip(LINEAR, L["pred"], L["dev_pct"]):
        print(f"  {v:5.1f} μL  A = {a:10.0f}  回归 {p:10.0f}  偏差 {d:+6.2f}%")
    print(f"  A = {L['slope']:.0f}·V {'−' if L['intercept'] < 0 else '+'} "
          f"{abs(L['intercept']):.0f}   r = {L['r']:.5f}   "
          f"r > {R_MIN} ? {'是' if L['r'] > R_MIN else '否'}")
    print(f"  最大相对偏差 {L['dev_abs_max']:.2f}%；斜率标准误 {L['slope_sd']:.1f}")

    P, S, C = precision(), stability(), content()
    print("\n--- 精密度 ---")
    print(f"  n = {P['n']}  均值 = {P['mean']:.2f}  s = {fmt(P['s'])}  "
          f"RSD = {fmt(P['rsd'])}%  → {'通过' if P['pass'] else '未通过'}")
    print("\n--- 稳定性 ---")
    print(f"  n = {S['n']}  均值 = {S['mean']:.2f}  s = {fmt(S['s'])}  "
          f"RSD = {fmt(S['rsd'])}%  → {'通过' if S['pass'] else '未通过'}")
    print("\n--- 含量测定 ---")
    print(f"  A对 = {C['a_std']:.2f}   A样 = {C['a_sam']:.2f}   "
          f"A样/A对 = {C['ratio']:.4f}")
    print(f"  C样 = {C['c_sam']:.4f} mg/mL → m = {C['m_mg']:.2f} mg"
          f"（理论 {W_THEOR_MG:.1f} mg）")
    print(f"  含量 = {C['mean']:.2f}%   每片 {C['per_tablet_mg']:.2f} mg   "
          f"RSD = {fmt(C['rsd'])}%")
    print(f"  判定：{CONTENT_LOWER:.1f}%~{CONTENT_UPPER:.1f}% → {verdict(C['pass'])}")

    D = diagnose()
    print("\n--- 数据一致性诊断 ---")
    print(f"  A样/A对 = {D['ratio_pct']:.2f}%   折算浓度 {D['c_sam']:.4f} mg/mL"
          f"（理论 {D['c_theor']:.3f}）")
    print(f"  应有峰面积 {D['a_expected']:.0f}，实测 {D['a_sam']:.0f}"
          f"（{D['ratio_expected'] * 100:.1f}%）")

    print("\n--- 方法学汇总 ---")
    for it in M["items"]:
        print(f"  {it['item']:6s} {it['index']:24s} 要求 {it['req']:14s} "
              f"{'通过' if it['pass'] else '未通过'}")
    print(f"  合计：{M['n_pass']}/{M['n_total']} 项通过")
    print("=" * 74)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
