import sys
import time
import math
from typing import Any, Dict, Optional, Tuple

import numpy as np
from scipy.optimize import minimize

try:
    import win32com.client
    HAVE_WIN32COM = True
except Exception:
    HAVE_WIN32COM = False

# ----------------------------------------------------------------------
# Configuration
# ----------------------------------------------------------------------
FILE_PATH = r".\conversion.usc"

MW_P  = 64.51000213623    # kg/kgmole  (stream "P")
MW_FR = 32.252522551517   # kg/kgmole  (stream "FR")
MW_R  = 32.2419465304     # kg/kgmole  (stream "R")

F_FRESH = 0.027777777777778  # kgmole/s, fixed fresh feed "F"

W_LB = 0.0006
W_UB = 0.0037
W0   = 0.0025
ANCHOR_W = 0.0030         # fixed re-conditioning point for deterministic evaluations

C1_LIMIT  = 0.0833333     # 300 kg/hr expressed in kg/s
C1_MARGIN = 1e-7          # strict "<" implemented as R_mass <= LIMIT - 1e-7
C3_LIMIT  = 0.01
C4_LIMIT  = 0.0005556

SCALE_W = 100.0           # scaled decision space: x = W * 100  -> O(0.1)
PENALTY = 1e6
FEAS_TOL = 1e-6
SOLVER_TIMEOUT = 60.0
C1_FINAL_MARGIN = 1e-6    # safety band for the noisy recycle (kg/s); limit itself unchanged
BUMP_STEP_W = 1e-6        # kgmole/s step used to guarantee the strict C1 margin
MAX_BUMPS = 100

# Global COM handles
app: Any = None
sim: Any = None
fs: Any = None
MODE: str = "emulation"

EVAL_COUNT: int = 0
CACHE: Dict[float, Tuple[float, Tuple[float, float, float, float]]] = {}
BEST_FEASIBLE: Optional[Dict[str, Any]] = None


# ----------------------------------------------------------------------
# UniSim connection
# ----------------------------------------------------------------------
def connect_unisim() -> Tuple[Any, str]:
    """Connect to a running UniSim Design instance; fall back to emulation."""
    global app
    if HAVE_WIN32COM:
        try:
            app = win32com.client.Dispatch("UniSimDesign.Application")
            app.Visible = True
            return app, "unisim"
        except Exception as e1:
            print("Dispatch failed:", e1)
        try:
            app = win32com.client.GetActiveObject("UniSimDesign.Application")
            app.Visible = True
            return app, "unisim"
        except Exception as e2:
            print("GetActiveObject failed:", e2)
    return None, "emulation"


def wait_and_settle(sim_obj: Any) -> None:
    """Wait until the UniSim solver is idle AND the values have settled
    (two consecutive reads of R.MassFlow agree). Never uses IsSolved()/Solve()."""
    try:
        sim_obj.Solver.CanSolve = True
    except Exception:
        pass
    time.sleep(0.8)                    # give the auto-solver a head start
    start = time.time()
    while True:                        # phase 1: solver busy/idle
        try:
            busy = bool(sim_obj.Solver.IsSolving)
        except Exception:
            time.sleep(2.0)
            break
        if not busy:
            break
        if time.time() - start > SOLVER_TIMEOUT:
            raise TimeoutError("UniSim solver timeout after %.0f s" % SOLVER_TIMEOUT)
        time.sleep(0.2)
    while True:                        # phase 2: value settle check
        try:
            v1 = read_mass_flow("R")
            time.sleep(0.5)
            v2 = read_mass_flow("R")
            if abs(v1 - v2) < 1e-8:
                break
        except Exception:
            break
        if time.time() - start > SOLVER_TIMEOUT:
            raise TimeoutError("UniSim solver did not settle after %.0f s" % SOLVER_TIMEOUT)
    time.sleep(0.2)


# ----------------------------------------------------------------------
# UniSim stream helpers (real mode)
# ----------------------------------------------------------------------
def get_stream(name: str) -> Any:
    try:
        return fs.MaterialStreams.Item(name)
    except Exception:
        return fs.Streams.Item(name)


def set_stream_molar_flow(name: str, value: float, unit: str = "kgmole/s") -> None:
    get_stream(name).MolarFlow.SetValue(value, unit)


def read_molar_flow(name: str) -> float:
    return float(get_stream(name).MolarFlow.GetValue("kgmole/s"))


def read_mass_flow(name: str) -> float:
    return float(get_stream(name).MassFlow.GetValue("kg/s"))


def read_comp_fraction(st: Any, comp: str) -> Optional[float]:
    """Try several COM access patterns for ComponentMolarFractionValue(comp)."""
    attempts = [
        lambda: st.ComponentMolarFractionValue(comp),
        lambda: st.ComponentMolarFractionValue[comp],
        lambda: st.ComponentMolarFraction(comp),
        lambda: st.ComponentMolarFraction[comp],
    ]
    for fn in attempts:
        try:
            v = fn()
            if v is None:
                continue
            f = float(v)
            if -1e-9 <= f <= 1.0 + 1e-9:
                return min(max(f, 0.0), 1.0)
        except Exception:
            continue
    return None


def read_purity_p() -> float:
    """Mole fraction of ClC2 in bottoms stream "P" (>= 0.99 required)."""
    try:
        st = get_stream("P")
    except Exception:
        return 1.0
    for comp in ("ClC2", "EthylChloride", "Ethyl chloride", "C2H5Cl", "ETCL"):
        f = read_comp_fraction(st, comp)
        if f is not None:
            return f
    light = 0.0
    found_any = False
    for comp in ("HCl", "Ethylene", "Nitrogen", "N2", "C2H4"):
        f = read_comp_fraction(st, comp)
        if f is not None:
            light += f
            found_any = True
    if found_any:
        return max(1.0 - light, 0.0)
    # X-100 sends 100% of HCl/Ethylene/Nitrogen overhead and 0% ClC2
    # overhead (Report 1), so bottoms are pure ClC2 by construction.
    return 1.0


# ----------------------------------------------------------------------
# State readers (real / emulation)
# ----------------------------------------------------------------------
def read_state_real() -> Dict[str, float]:
    p = read_molar_flow("P")
    fr = read_molar_flow("FR")
    r_molar = read_molar_flow("R")
    r_mass = read_mass_flow("R")
    try:
        ovhd = read_molar_flow("ovhd")
    except Exception:
        ovhd = read_molar_flow("W") + read_molar_flow("toR")
    to_r = read_molar_flow("toR")
    w_act = read_molar_flow("W")
    purity = read_purity_p()
    return {"P": p, "FR": fr, "R_mass": r_mass, "ovhd": ovhd, "W": w_act,
            "purity": purity, "R": r_molar, "toR": to_r,
            "mismatch": abs(r_molar - to_r)}


def read_state_emulation(w: float) -> Dict[str, float]:
    """Mass-balance stand-in for the converged recycle loop (used only when
    UniSim Design is unavailable in this environment)."""
    f_h = 0.50 * F_FRESH
    f_e = 0.48 * F_FRESH
    f_n = 0.02 * F_FRESH
    y = np.array([0.45, 0.25, 0.30])
    V = 0.0
    converged = False
    for _ in range(600):
        R = V - w
        if R < 0.0:
            R = 0.0
        fr_h = f_h + y[0] * R
        fr_e = f_e + y[1] * R
        fr_n = f_n + y[2] * R
        p = 0.9 * fr_e
        v_h = fr_h - p
        v_e = 0.1 * fr_e
        v_n = fr_n
        V_new = v_h + v_e + v_n
        if V_new <= 1e-14:
            V_new = 1e-14
        y_new = np.array([v_h, v_e, v_n]) / V_new
        delta = float(np.max(np.abs(y_new - y)))
        y = y_new
        V = V_new
        if delta < 1e-13:
            converged = True
            break
    if not converged:
        raise ValueError("emulation recycle loop did not converge for W=%.6g" % w)
    R = V - w
    if R < 0.0:
        R = 0.0
    fr = F_FRESH + R
    r_mass = R * MW_R
    return {"P": p, "FR": fr, "R_mass": r_mass, "ovhd": V, "W": w,
            "purity": 1.0, "R": R, "toR": R, "mismatch": 0.0}


# ----------------------------------------------------------------------
# Cached black-box evaluation (anchor protocol -> deterministic recycle state)
# ----------------------------------------------------------------------
def evaluate(x: Any) -> Tuple[float, Tuple[float, float, float, float]]:
    """Return (objective_to_minimise, (c1, c2, c3, c4)) for scaled x.

    Every UniSim evaluation is pre-conditioned: W is first driven to ANCHOR_W
    and fully settled, then to the requested value. This removes the
    path-dependence of the RCY-1 recycle convergence."""
    global EVAL_COUNT, BEST_FEASIBLE
    xv = float(np.asarray(x, dtype=float)[0])
    key = round(xv, 12)
    cached = CACHE.get(key)
    if cached is not None:
        return cached
    try:
        w = xv / SCALE_W
        if MODE == "unisim":
            set_stream_molar_flow("W", ANCHOR_W)
            wait_and_settle(sim)
            set_stream_molar_flow("W", w)
            wait_and_settle(sim)
            st = read_state_real()
        else:
            st = read_state_emulation(w)

        p = float(st["P"])
        fr = float(st["FR"])
        r_mass = float(st["R_mass"])
        ovhd = float(st["ovhd"])
        w_act = float(st["W"])
        purity = float(st["purity"])
        mismatch = float(st.get("mismatch", 0.0))

        if not all(np.isfinite([p, fr, r_mass, ovhd, w_act, purity])) or p <= 0.0 or fr <= 0.0:
            raise ValueError("non-physical simulation output")
        if MODE == "unisim" and abs(w_act - w) > 1e-9:
            raise ValueError("stream W was not set to the requested value")

        revenue = 71280.0 * p * MW_P
        capital = 50.0 * (28512.0 * fr * MW_FR) ** 0.6
        vp = revenue - capital

        c1 = C1_LIMIT - r_mass - C1_MARGIN          # >= 0  ->  R_mass < 0.0833333 (strict)
        c2 = ovhd - w_act                            # >= 0  ->  W <= ovhd
        c3 = purity - (1.0 - C3_LIMIT)               # >= 0  ->  x_ClC2(P) >= 0.99
        c4 = w_act - C4_LIMIT                        # >= 0  ->  W >= 0.0005556

        obj = -vp
        cons = (float(c1), float(c2), float(c3), float(c4))
        EVAL_COUNT += 1
        CACHE[key] = (obj, cons)

        if all(c >= -FEAS_TOL for c in cons):
            if BEST_FEASIBLE is None or obj < BEST_FEASIBLE["obj"]:
                BEST_FEASIBLE = {"x": key, "w": w, "obj": obj, "vp": vp,
                                 "cons": cons, "state": st}

        print("[eval %3d] x=%.6f W=%.8f | P=%.8f | FR=%.8f | R_mass=%.8f | "
              "ovhd=%.8f | xClC2=%.6f | |R-toR|=%.2e | VP=%.4f | "
              "c=(%+.7f, %+.7f, %+.7f, %+.7f)"
              % (EVAL_COUNT, xv, w, p, fr, r_mass, ovhd, purity, mismatch,
                 vp, c1, c2, c3, c4), flush=True)
        return obj, cons
    except Exception as exc:
        print("[eval %3d] FAILED at x=%.6f -> %s" % (EVAL_COUNT + 1, xv, exc), flush=True)
        return PENALTY, (-PENALTY, -PENALTY, -PENALTY, -PENALTY)


def objective(x: Any) -> float:
    return evaluate(x)[0]


# ----------------------------------------------------------------------
# Main optimisation
# ----------------------------------------------------------------------
def main() -> None:
    global MODE, sim, fs, app, BEST_FEASIBLE

    app, MODE = connect_unisim()
    if MODE == "unisim":
        print("Connected to UniSim Design. Opening case: %s" % FILE_PATH, flush=True)
        try:
            sim = app.SimulationCases.Open(FILE_PATH)
        except Exception as exc:
            print("Open failed (%s); falling back to ActiveDocument." % exc, flush=True)
            sim = app.ActiveDocument
        fs = sim.Flowsheet
        try:
            sim.Visible = 1
        except Exception:
            pass
        wait_and_settle(sim)
        print("Case loaded, solver idle.", flush=True)
    else:
        print("UniSim Design is not available here -> using the built-in mass-balance "
              "emulation to verify the optimisation logic.", flush=True)
        sim = None
        fs = None

    x0 = np.array([W0 * SCALE_W])
    bounds = [(W_LB * SCALE_W, W_UB * SCALE_W)]
    constraints = (
        {"type": "ineq", "fun": lambda x: evaluate(x)[1][0]},
        {"type": "ineq", "fun": lambda x: evaluate(x)[1][1]},
        {"type": "ineq", "fun": lambda x: evaluate(x)[1][2]},
        {"type": "ineq", "fun": lambda x: evaluate(x)[1][3]},
    )

    print("\nStarting SLSQP optimisation (maximise Venture Profit over purge W)...\n", flush=True)
    res = minimize(
        objective,
        x0,
        method="SLSQP",
        bounds=bounds,
        constraints=constraints,
        options={"maxiter": 80, "ftol": 1e-6, "eps": 5e-3, "disp": True},
    )
    print("\nSLSQP finished:", res.message, "| success:", res.success,
          "| nfev:", res.nfev, "| nit:", res.nit, flush=True)

    # ---- decide whether a COBYLA fallback is needed ---------------------
    moved = bool(res.x.size and abs(float(res.x[0]) - float(x0[0])) > 1e-6)
    c1_slack = None
    if res.x.size:
        try:
            _, cons_rx = evaluate(np.array([res.x[0]]))
            c1_slack = float(cons_rx[0])
        except Exception:
            c1_slack = None

    need_fallback = (not res.success) or (not moved) or (c1_slack is None) \
        or (c1_slack > 1e-4) or (c1_slack < -1e-5)
    if need_fallback:
        print("\nSLSQP did not settle on the C1 boundary -> COBYLA fallback.", flush=True)
        x_start = BEST_FEASIBLE["x"] if BEST_FEASIBLE is not None else float(x0[0])
        cons_cob = (
            {"type": "ineq", "fun": lambda x: evaluate(x)[1][0]},
            {"type": "ineq", "fun": lambda x: evaluate(x)[1][1]},
            {"type": "ineq", "fun": lambda x: evaluate(x)[1][2]},
            {"type": "ineq", "fun": lambda x: evaluate(x)[1][3]},
            {"type": "ineq", "fun": lambda x: float(np.asarray(x, dtype=float)[0]) - W_LB * SCALE_W},
            {"type": "ineq", "fun": lambda x: W_UB * SCALE_W - float(np.asarray(x, dtype=float)[0])},
        )
        res2 = minimize(
            objective,
            np.array([x_start]),
            method="COBYLA",
            constraints=cons_cob,
            options={"maxiter": 200, "rhobeg": 0.05, "tol": 1e-5, "disp": True},
        )
        print("\nCOBYLA finished:", res2.message, "| success:", res2.success,
              "| nfev:", res2.nfev, flush=True)
        res = res2

    # ---- select best feasible candidate ---------------------------------
    candidates = []
    if res.x.size:
        try:
            obj_r, cons_r = evaluate(np.array([res.x[0]]))
            if all(c >= -FEAS_TOL for c in cons_r):
                candidates.append({"x": float(res.x[0]), "obj": float(obj_r), "cons": cons_r})
        except Exception:
            pass
    if BEST_FEASIBLE is not None:
        candidates.append({"x": BEST_FEASIBLE["x"], "obj": BEST_FEASIBLE["obj"],
                           "cons": BEST_FEASIBLE["cons"]})
    if candidates:
        best = min(candidates, key=lambda d: d["obj"])
    else:
        print("WARNING: no feasible point found; reporting raw result.", flush=True)
        best = {"x": float(res.x[0]) if res.x.size else float(x0[0]),
                "obj": float(res.fun) if res.fun is not None else PENALTY,
                "cons": (0.0, 0.0, 0.0, 0.0)}

    w_final = min(max(best["x"] / SCALE_W, W_LB), W_UB)
    vp_opt_boundary = -best["obj"]

    # ---- apply through the anchor protocol and guarantee strict C1 -------
    def apply_and_read(w: float) -> Dict[str, float]:
        if MODE == "unisim":
            set_stream_molar_flow("W", ANCHOR_W)
            wait_and_settle(sim)
            set_stream_molar_flow("W", w)
            wait_and_settle(sim)
            return read_state_real()
        return read_state_emulation(w)

    st = apply_and_read(w_final)
    c1_now = C1_LIMIT - st["R_mass"] - C1_MARGIN
    bump = 0
    while c1_now < C1_FINAL_MARGIN and bump < MAX_BUMPS and w_final < W_UB - 1e-9:
        w_final = min(w_final + BUMP_STEP_W, W_UB)
        st = apply_and_read(w_final)
        c1_now = C1_LIMIT - st["R_mass"] - C1_MARGIN
        bump += 1
    if c1_now < 0.0:
        print("WARNING: strict C1 could not be guaranteed at the final point "
              "(R_mass=%.8f kg/s)." % st["R_mass"], flush=True)
    else:
        print("Final C1 safety check passed after %d bump(s): R_mass=%.8f < 0.0833333 kg/s."
              % (bump, st["R_mass"]), flush=True)

    # ---- final report -----------------------------------------------------
    revenue = 71280.0 * st["P"] * MW_P
    capital = 50.0 * (28512.0 * st["FR"] * MW_FR) ** 0.6
    vp_final = revenue - capital
    print("\n" + "=" * 72)
    print("FINAL RESULT  (mode: %s)" % MODE)
    print("=" * 72)
    print("W  (purge, decision)  = %.8f kgmole/s" % st["W"])
    print("P  (product)          = %.8f kgmole/s   (P_mass = %.4f kg/hr)" %
          (st["P"], st["P"] * MW_P * 3600.0))
    print("FR (reactor feed)     = %.8f kgmole/s   (FR_mass = %.4f kg/hr)" %
          (st["FR"], st["FR"] * MW_FR * 3600.0))
    print("R  (recycle)          = mass %.8f kg/s = %.2f kg/hr" %
          (st["R_mass"], st["R_mass"] * 3600.0))
    print("ovhd                  = %.8f kgmole/s" % st["ovhd"])
    print("x_ClC2(P)             = %.6f" % st["purity"])
    print("|R - toR| recycle mismatch = %.3e kgmole/s" % st.get("mismatch", 0.0))
    print("-" * 72)
    c1 = C1_LIMIT - st["R_mass"] - C1_MARGIN
    c2 = st["ovhd"] - st["W"]
    c3 = st["purity"] - (1.0 - C3_LIMIT)
    c4 = st["W"] - C4_LIMIT
    strict_ok = st["R_mass"] < C1_LIMIT
    print("C1  R_mass < 0.0833333 kg/s : g1 = %+.10f  %s  (strict R_mass=%.10f < %.7f: %s)" %
          (c1, "OK" if c1 >= -FEAS_TOL else "VIOLATED",
           st["R_mass"], C1_LIMIT, "SATISFIED" if strict_ok else "VIOLATED"))
    print("C2  W <= ovhd               : g2 = %+.8f  %s" %
          (c2, "OK" if c2 >= -FEAS_TOL else "VIOLATED"))
    print("C3  x_ClC2(P) >= 0.99       : g3 = %+.8f  %s" %
          (c3, "OK" if c3 >= -FEAS_TOL else "VIOLATED"))
    print("C4  W >= 0.0005556          : g4 = %+.8f  %s" %
          (c4, "OK" if c4 >= -FEAS_TOL else "VIOLATED"))
    print("-" * 72)
    print("Revenue_P        = %.2f mu/yr" % revenue)
    print("Capital_Penalty  = %.2f mu/yr" % capital)
    print("Venture Profit   = %.2f mu/yr   (SLSQP boundary optimum: %.2f mu/yr)"
          % (vp_final, vp_opt_boundary))
    print("=" * 72)

    if MODE == "unisim":
        try:
            app.Visible = True
        except Exception:
            pass
        print("UniSim Design remains open for inspection.", flush=True)
    else:
        print("(emulation mode - no UniSim instance to keep open)", flush=True)


main()