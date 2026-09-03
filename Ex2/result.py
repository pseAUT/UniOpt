"""
Three-stage nitrogen compression train optimisation (UniSim Design COM automation).

Goal:
    minimise  Total_Compressor_Duty = |W_E1| + |W_E2| + |W_E3|  [kW]
    by choosing the intermediate discharge pressures of compressors
    K-100 (stream "Mid1") and K-101 (stream "Mid2").

Fixed process data (Report 1):
    Feed1  : 100 kPa, 25 C, pure N2
    Output : 1000 kPa (fixed)
    E-100/E-101 coolers : 25 C outlet, 0 kPa pressure drop

Decision variables (scaled by SCALE_P = 1000):
    x[0] = P_Mid1 / 1000   (101 <= P_Mid1 <= 900 kPa)
    x[1] = P_Mid2 / 1000   (102 <= P_Mid2 <= 999 kPa)

Constraints (SLSQP ineq, must be >= 0):
    C1: P_Mid1 - 101        >= 0
    C2: P_Mid2 - P_Mid1 - 1 >= 0
    C3: 1000 - P_Mid2 - 1   >= 0
"""

import sys
import time
import math
from typing import Any, Optional, Tuple

import numpy as np
import win32com.client
from scipy.optimize import minimize

# ----------------------------------------------------------------------------
# Configuration
# ----------------------------------------------------------------------------
UNISIM_FILE = r".\Comp.usc"

STREAM_FEED = "Feed1"
STREAM_MID1 = "Mid1"
STREAM_MID2 = "Mid2"
STREAM_OUT = "Output"
ENERGY_STREAMS = ("E1", "E2", "E3")

P_FEED = 100.0   # kPa, fixed
P_OUT = 1000.0   # kPa, fixed

P_MID1_LB = 101.0
P_MID1_UB = 900.0
P_MID2_LB = 102.0
P_MID2_UB = 999.0
P_MID1_NOM = 300.0   # Report 1 base case
P_MID2_NOM = 600.0   # Report 1 base case

SCALE_P = 1000.0         # kPa -> scaled space
PENALTY = 1.0e6
SOLVER_TIMEOUT = 30.0    # seconds
SLSQP_MAXITER = 60
COBYLA_MAXITER = 80

# ----------------------------------------------------------------------------
# Globals (COM objects are untyped by design)
# ----------------------------------------------------------------------------
app: Any = None
sim: Any = None
_eval_cache: dict = {}
_eval_count = 0
_stream_obj_cache: dict = {}
_energy_obj_cache: dict = {}


def remove_ipython_f_arg() -> None:
    """Ignore the -f argument injected by Jupyter/IPython %run."""
    if "-f" in sys.argv:
        sys.argv.remove("-f")


# ----------------------------------------------------------------------------
# Helper functions
# ----------------------------------------------------------------------------
def open_or_reuse_case(app_obj: Any, path: str) -> Any:
    """Return the simulation case, reusing an already-open instance if possible."""
    norm = path.replace("/", "\\").lower()
    try:
        n_cases = int(app_obj.SimulationCases.Count)
    except Exception:
        n_cases = 0
    for i in range(n_cases):
        try:
            case = app_obj.SimulationCases.Item(i)
            if str(case.FullName).replace("/", "\\").lower() == norm:
                return case
        except Exception:
            continue
    return app_obj.SimulationCases.Open(path)


def wait_for_solver(case: Any, timeout: float = SOLVER_TIMEOUT) -> None:
    """
    Wait until UniSim's solver is idle.
    CRITICAL RULE: never use .Solve() or .IsSolved - only the IsSolving loop.
    """
    try:
        case.Solver.CanSolve = True
    except Exception:
        pass
    time.sleep(0.2)  # let the solver kick off after the last change

    # Detect whether IsSolving is available at all
    try:
        _ = bool(case.Solver.IsSolving)
        is_solving_available = True
    except Exception:
        is_solving_available = False

    if not is_solving_available:
        # Fallback specified in the instructions
        time.sleep(2.0)
        return

    start = time.time()
    while True:
        try:
            solving = bool(case.Solver.IsSolving)
        except Exception:
            solving = False
        if not solving:
            break
        if time.time() - start > timeout:
            raise TimeoutError(f"UniSim solver did not finish within {timeout:.1f} s")
        time.sleep(0.1)
    time.sleep(0.2)  # settling time


def get_material_stream(name: str) -> Any:
    if name not in _stream_obj_cache:
        _stream_obj_cache[name] = sim.Flowsheet.MaterialStreams.Item(name)
    return _stream_obj_cache[name]


def get_energy_stream(name: str) -> Any:
    if name not in _energy_obj_cache:
        _energy_obj_cache[name] = sim.Flowsheet.EnergyStreams.Item(name)
    return _energy_obj_cache[name]


def set_stream_pressure(name: str, value: float, unit: str = "kPa") -> None:
    """Set a material stream pressure directly (CRITICAL RULE)."""
    stream = get_material_stream(name)
    last_exc: Optional[Exception] = None
    for _attempt in range(5):
        try:
            stream.Pressure.SetValue(float(value), unit)
            return
        except Exception as exc:  # COM busy retries
            last_exc = exc
            time.sleep(0.5)
    raise RuntimeError(f"Could not set pressure of '{name}': {last_exc}")


def read_stream_pressure(name: str) -> float:
    stream = get_material_stream(name)
    for _attempt in range(5):
        try:
            return float(stream.Pressure.GetValue("kPa"))
        except Exception:
            time.sleep(0.5)
    return float(stream.PressureValue)  # internal unit (kPa)


def read_duty(name: str) -> float:
    """Read an energy-stream duty in kW (signed, as reported by UniSim)."""
    es = get_energy_stream(name)
    for _attempt in range(5):
        try:
            # Preferred: explicit unit conversion
            for var_name in ("Power", "HeatFlow"):
                try:
                    var = getattr(es, var_name)
                    v = float(var.GetValue("kW"))
                    if math.isfinite(v):
                        return v
                except Exception:
                    continue
            # Fallback: internal HeatFlowValue (assumed kJ/h) -> kW
            v = float(es.HeatFlowValue)
            if math.isfinite(v):
                return v / 3600.0
        except Exception:
            time.sleep(0.5)
    raise RuntimeError(f"Could not read duty of energy stream '{name}'")


# ----------------------------------------------------------------------------
# Cached evaluation function
# ----------------------------------------------------------------------------
def evaluate(x: Any) -> Tuple[float, np.ndarray]:
    """Return (objective, constraint_array) for a point in SCALED space."""
    global _eval_count
    xa = np.asarray(x, dtype=float).ravel()
    p1 = float(np.clip(xa[0] * SCALE_P, P_MID1_LB, P_MID1_UB))
    p2 = float(np.clip(xa[1] * SCALE_P, P_MID2_LB, P_MID2_UB))
    key = (round(p1, 3), round(p2, 3))

    cached = _eval_cache.get(key, None)
    if cached is not None:
        return cached

    try:
        set_stream_pressure(STREAM_MID1, p1)
        set_stream_pressure(STREAM_MID2, p2)
        wait_for_solver(sim)

        duties = [read_duty(n) for n in ENERGY_STREAMS]
        if not all(math.isfinite(d) for d in duties):
            raise ValueError(f"non-finite duty values: {duties}")

        # Objective = total compressor shaft work (absolute values make the
        # result independent of the energy-stream sign convention).
        obj = float(sum(abs(d) for d in duties))

        cons = np.array([
            p1 - P_MID1_LB,       # C1 : P_Mid1 >= 101
            p2 - p1 - 1.0,        # C2 : P_Mid2 - P_Mid1 >= 1
            P_OUT - 1.0 - p2,     # C3 : 1000 - P_Mid2 >= 1
        ], dtype=float)

        result = (obj, cons)
        _eval_count += 1
        signed = ", ".join(f"{n}={d:+.4f}" for n, d in zip(ENERGY_STREAMS, duties))
        print(f"[eval {_eval_count:3d}] P_Mid1={p1:8.3f} kPa  P_Mid2={p2:8.3f} kPa | "
              f"duties [kW]: {signed} | total={obj:.4f} kW | cons={np.round(cons, 3)}")
    except Exception as exc:
        print(f"[eval ERROR] P_Mid1={p1:.3f} kPa, P_Mid2={p2:.3f} kPa -> {exc}")
        result = (PENALTY, np.array([-1.0, -1.0, -1.0], dtype=float))

    _eval_cache[key] = result
    return result


def objective(x: Any) -> float:
    return float(evaluate(x)[0])


def constraint_vec(x: Any) -> np.ndarray:
    return evaluate(x)[1]


def constraint_c1(x: Any) -> float:
    return float(evaluate(x)[1][0])


def constraint_c2(x: Any) -> float:
    return float(evaluate(x)[1][1])


def constraint_c3(x: Any) -> float:
    return float(evaluate(x)[1][2])


# ----------------------------------------------------------------------------
# Main
# ----------------------------------------------------------------------------
def main() -> None:
    global app, sim
    remove_ipython_f_arg()

    print("=" * 78)
    print("Three-stage N2 compression train optimisation (UniSim COM automation)")
    print("=" * 78)

    print("Starting UniSim Design ...")
    app = win32com.client.Dispatch("UniSimDesign.Application")
    app.Visible = True

    print("Opening case:", UNISIM_FILE)
    sim = open_or_reuse_case(app, UNISIM_FILE)
    try:
        sim.Visible = 1
    except Exception:
        pass
    time.sleep(1.0)
    wait_for_solver(sim)

    # ----- Sanity check: stream names and base-case state -----
    print("\nCurrent stream pressures (kPa):")
    for nm in (STREAM_FEED, STREAM_MID1, STREAM_MID2, STREAM_OUT):
        try:
            print(f"  {nm:8s} = {read_stream_pressure(nm):10.3f}")
        except Exception as exc:
            print(f"  {nm:8s} = UNAVAILABLE ({exc})")

    try:
        d0 = [read_duty(n) for n in ENERGY_STREAMS]
        print("Base-case duties (kW):",
              ", ".join(f"{n}={v:+.4f}" for n, v in zip(ENERGY_STREAMS, d0)))
        print(f"Base-case total compressor duty = {sum(abs(v) for v in d0):.6f} kW")
    except Exception as exc:
        print(f"WARNING: could not read base-case duties: {exc}")

    # ----- Optimisation setup (scaled space) -----
    bounds = [
        (P_MID1_LB / SCALE_P, P_MID1_UB / SCALE_P),
        (P_MID2_LB / SCALE_P, P_MID2_UB / SCALE_P),
    ]
    x0 = np.array([P_MID1_NOM / SCALE_P, P_MID2_NOM / SCALE_P])
    print(f"\nScaled initial guess x0 = {x0}  (physical: {x0 * SCALE_P} kPa)")
    print("Scaled bounds:", bounds)

    cons_slsqp = [{"type": "ineq", "fun": constraint_vec}]
    options_slsqp = {
        "maxiter": SLSQP_MAXITER,
        "ftol": 1.0e-6,
        "eps": 1.0e-3,   # finite-difference step = 1 kPa in physical units
        "disp": True,
    }

    print("\nRunning SLSQP ...")
    res = minimize(
        objective, x0, method="SLSQP",
        bounds=bounds, constraints=cons_slsqp, options=options_slsqp,
    )

    # ----- Robustness fallback: COBYLA (derivative-free) -----
    if (not res.success) or (float(res.fun) >= PENALTY / 2.0):
        print("\nSLSQP failed or hit the penalty region -> falling back to COBYLA ...")
        cons_cobyla = [
            {"type": "ineq", "fun": constraint_c1},
            {"type": "ineq", "fun": constraint_c2},
            {"type": "ineq", "fun": constraint_c3},
        ]
        res2 = minimize(
            objective, x0, method="COBYLA",
            bounds=bounds, constraints=cons_cobyla,
            options={"maxiter": COBYLA_MAXITER, "rhobeg": 0.01, "tol": 1.0e-6, "disp": True},
        )
        if float(res2.fun) < float(res.fun):
            res = res2

    # ----- Report -----
    print("\n" + "=" * 78)
    print("OPTIMISATION REPORT")
    print("=" * 78)
    print(f"  success      : {res.success}")
    print(f"  message      : {res.message}")
    print(f"  iterations   : {getattr(res, 'nit', 'n/a')}")
    print(f"  objective    : {res.fun:.6f} kW (reported)")
    p1_opt = float(res.x[0] * SCALE_P)
    p2_opt = float(res.x[1] * SCALE_P)
    print(f"  P_Mid1 opt   : {p1_opt:.6f} kPa")
    print(f"  P_Mid2 opt   : {p2_opt:.6f} kPa")

    # ----- Apply optimum and verify in UniSim -----
    print("\nApplying the optimal point to UniSim ...")
    set_stream_pressure(STREAM_MID1, p1_opt)
    set_stream_pressure(STREAM_MID2, p2_opt)
    wait_for_solver(sim)

    duties_opt = [read_duty(n) for n in ENERGY_STREAMS]
    total_opt = sum(abs(d) for d in duties_opt)
    print("Optimal duties (kW):",
          ", ".join(f"{n}={v:+.6f}" for n, v in zip(ENERGY_STREAMS, duties_opt)))
    print(f"Optimal total compressor duty = {total_opt:.6f} kW")

    c1 = p1_opt - P_MID1_LB
    c2 = p2_opt - p1_opt - 1.0
    c3 = P_OUT - 1.0 - p2_opt
    print(f"Constraint values (must be >= 0):  C1={c1:+.6f}  C2={c2:+.6f}  C3={c3:+.6f}")

    for nm in (STREAM_FEED, STREAM_MID1, STREAM_MID2, STREAM_OUT):
        try:
            print(f"  {nm:8s} pressure = {read_stream_pressure(nm):10.3f} kPa")
        except Exception as exc:
            print(f"  {nm:8s} pressure = UNAVAILABLE ({exc})")

    print("\nOptimisation finished. UniSim remains open with the optimal case.")
    app.Visible = True
    try:
        sim.Visible = 1
    except Exception:
        pass


# Support both plain-script and Jupyter execution.
if __name__ == "__main__" or "ipykernel" in sys.modules or "-f" in sys.argv:
    main()