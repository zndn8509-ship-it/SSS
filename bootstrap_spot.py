"""
회사채 YTM으로부터 스팟 레이트(spot rate) 커브를 부트스트래핑으로 산출한다.

한국 관행: 3개월 복리(분기 복리), Actual/365.

입력 CSV 컬럼:
    maturity_years : 잔존만기 (년, 실수)
    coupon_rate    : 연 표면금리 (소수, 예: 4% -> 0.04)
    ytm            : 시장 YTM     (소수, 분기 복리 기준)
    face           : (옵션) 액면가, 기본 100

사용 예:
    python bootstrap_spot.py sample_bonds.csv -o spot_curve.csv --plot spot_curve.png

KOFIA (https://www.kofiabond.or.kr/) 데이터 활용:
    공식 공개 API는 제공되지 않는다. 웹사이트에서 '시가평가 수익률' 또는
    '민평금리'를 엑셀/CSV로 다운로드한 뒤, 본 스크립트의 입력 스키마에
    맞춰 maturity_years / coupon_rate / ytm 컬럼을 정리해 사용한다.
"""

from __future__ import annotations

import argparse
from dataclasses import dataclass

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from scipy.interpolate import interp1d
from scipy.optimize import brentq


@dataclass
class Bond:
    maturity_years: float
    coupon_rate: float
    ytm: float
    face: float = 100.0


def coupon_times(maturity_years: float, freq: int) -> np.ndarray:
    """만기까지의 쿠폰 지급 시점(년 단위) 배열. 마지막 원소는 정확히 만기."""
    n = int(round(maturity_years * freq))
    if n < 1:
        return np.array([maturity_years], dtype=float)
    times = np.arange(1, n + 1, dtype=float) / freq
    times[-1] = maturity_years
    return times


def price_from_ytm(bond: Bond, freq: int = 4) -> float:
    times = coupon_times(bond.maturity_years, freq)
    coupon = bond.coupon_rate * bond.face / freq
    cashflows = np.full_like(times, coupon)
    cashflows[-1] += bond.face
    discount = (1.0 + bond.ytm / freq) ** (freq * times)
    return float(np.sum(cashflows / discount))


def bootstrap_spot_curve(bonds: list[Bond], freq: int = 4) -> pd.DataFrame:
    """만기 오름차순으로 정렬한 뒤, 각 채권의 이론가격과 일치하도록
    해당 만기의 스팟 레이트를 순차적으로 해를 구한다.

    중간 쿠폰 시점의 스팟은 이미 부트스트랩된 점들을 선형보간해 사용한다.
    """
    bonds = sorted(bonds, key=lambda b: b.maturity_years)
    known_mats: list[float] = []
    known_rates: list[float] = []

    for bond in bonds:
        T = bond.maturity_years
        target_price = price_from_ytm(bond, freq)
        coupon = bond.coupon_rate * bond.face / freq
        times = coupon_times(T, freq)

        def pv_given_sT(sT: float) -> float:
            mats = known_mats + [T]
            rates = known_rates + [sT]
            if len(mats) == 1:
                def curve(_t: float) -> float:
                    return rates[0]
            else:
                f = interp1d(mats, rates, kind="linear", fill_value="extrapolate")

                def curve(t: float) -> float:
                    return float(f(t))

            pv = 0.0
            for i, t in enumerate(times):
                cf = coupon + (bond.face if i == len(times) - 1 else 0.0)
                s = curve(t)
                pv += cf / (1.0 + s / freq) ** (freq * t)
            return pv

        solution = brentq(lambda s: pv_given_sT(s) - target_price, -0.5, 1.0, xtol=1e-10)
        known_mats.append(T)
        known_rates.append(solution)

    return pd.DataFrame({"maturity_years": known_mats, "spot_rate": known_rates})


def load_bonds_csv(path: str) -> list[Bond]:
    df = pd.read_csv(path)
    required = {"maturity_years", "coupon_rate", "ytm"}
    missing = required - set(df.columns)
    if missing:
        raise ValueError(f"CSV에 다음 컬럼이 없습니다: {sorted(missing)}")
    bonds: list[Bond] = []
    for _, row in df.iterrows():
        bonds.append(
            Bond(
                maturity_years=float(row["maturity_years"]),
                coupon_rate=float(row["coupon_rate"]),
                ytm=float(row["ytm"]),
                face=float(row["face"]) if "face" in df.columns else 100.0,
            )
        )
    return bonds


def plot_curve(
    spot_df: pd.DataFrame,
    bonds: list[Bond] | None = None,
    path: str | None = None,
) -> plt.Figure:
    fig, ax = plt.subplots(figsize=(9, 5))
    ax.plot(
        spot_df["maturity_years"],
        spot_df["spot_rate"] * 100,
        "o-",
        label="Spot rate (bootstrapped)",
    )
    if bonds:
        ax.plot(
            [b.maturity_years for b in bonds],
            [b.ytm * 100 for b in bonds],
            "x--",
            alpha=0.6,
            label="YTM (input)",
        )
    ax.set_xlabel("Maturity (years)")
    ax.set_ylabel("Rate (%)")
    ax.set_title("Bootstrapped Spot Curve (quarterly comp., Act/365)")
    ax.grid(True, alpha=0.3)
    ax.legend()
    fig.tight_layout()
    if path:
        fig.savefig(path, dpi=150)
    return fig


def build_spot_function(spot_df: pd.DataFrame):
    """임의 만기 t(년)에 대해 spot(t)를 반환하는 선형보간 함수."""
    f = interp1d(
        spot_df["maturity_years"].values,
        spot_df["spot_rate"].values,
        kind="linear",
        fill_value="extrapolate",
    )
    return lambda t: float(f(t))


def main() -> None:
    parser = argparse.ArgumentParser(
        description="회사채 YTM으로부터 스팟 레이트 커브를 부트스트래핑한다."
    )
    parser.add_argument("input_csv", help="입력 CSV 경로 (maturity_years, coupon_rate, ytm)")
    parser.add_argument("-o", "--output", default="spot_curve.csv", help="결과 CSV 경로")
    parser.add_argument("--plot", default="spot_curve.png", help="결과 그래프 경로")
    parser.add_argument(
        "--freq", type=int, default=4, help="연간 복리 주기 (기본 4 = 분기)"
    )
    args = parser.parse_args()

    bonds = load_bonds_csv(args.input_csv)
    spot_df = bootstrap_spot_curve(bonds, freq=args.freq)
    spot_df.to_csv(args.output, index=False)

    print(spot_df.to_string(index=False, float_format=lambda x: f"{x:.6f}"))
    plot_curve(spot_df, bonds, path=args.plot)
    print(f"\nSaved: {args.output}, {args.plot}")


if __name__ == "__main__":
    main()
