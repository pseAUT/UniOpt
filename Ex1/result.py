import math
import time
from typing import Any, Dict, List, Optional, Tuple

import numpy as np
from scipy.optimize import minimize

import win32com.client

try:
    import pythoncom
    pythoncom.CoInitialize()
except Exception:
    pass

# ----------------------------------------------------------------------
# Configuration
# ----------------------------------------------------------------------
UNISIM_FILE = r"F:\Iliya_Project\AI\autonomos_optimization\hex\hex.usc"

SCALE_T = 10.0          # scaled decision variables -> degrees C
SOLVER_TIMEOUT = 30.0   # seconds per flowsheet solve

SLSQP_OPTIONS = {"maxiter": 80, "ftol": 1e-6, "eps": 0.05, "disp": True}
COBYLA_OPTIONS = {"maxiter": 200, "tol": 1e-5, "rhobeg": 0.05, "catol": 0.05, "disp": True}

# ----------------------------------------------------------------------
# Globals
# ----------------------------------------------------------------------
app: Any = None
sim: Any = None
fs: Any = None

MID1: Any = None
MID2: Any = None
OUT: Any = None
INFLOW: Any = None
E1_HANDLE: Any = None
E2: Any = None
E3: Any = None
OP100: Any = None
OP101: Any = None
OP102: Any = None

EVAL_CACHE: Dict[Tuple[float, float, float], Optional[Tuple[float, np.ndarray]]] = {}
EVAL_COUNT = 0
BEST: Dict[str, Any] = {"x": None, "obj": float("inf"), "cons": None}

COOLING_MODE = "unknown"
H_REF_25: Optional[float] = None
MASS_FLOW: Optional[float] = None
E1_STORED: Optional[float] = None


# ----------------------------------------------------------------------
# Helpers
# ----------------------------------------------------------------------
def wait_for_solver(sim_obj: Any) -> None:
    """Wait until the UniSim solver is idle (never calls .Solve())."""
    try:
        sim_obj.Solver.CanSolve = True
    except Exception:
        pass
    try:
        time.sleep(0.3)
        start = time.time()
        while True:
            try:
                solving = bool(sim_obj.Solver.IsSolving)
            except Exception:
                time.sleep(2.0)
                return
            if not solving and (time.time() - start) > 0.5:
                return
            if time.time() - start > SOLVER_TIMEOUT:
                raise TimeoutError(f"UniSim solver did not finish within {SOLVER_TIMEOUT} s")
            time.sleep(0.1)
    except TimeoutError:
        raise
    except Exception:
        time.sleep(2.0)
        return


def safe_names(coll: Any) -> List[str]:
    out: List[str] = []
    try:
        for i in range(1, int(coll.Count) + 1):
            try:
                out.append(str(coll.Item(i).name))
            except Exception:
                pass
    except Exception:
        pass
    return out


def stream_attached(s: Any) -> List[str]:
    out: List[str] = []
    try:
        for i in range(1, int(s.AttachedOpers.Count) + 1):
            out.append(str(s.AttachedOpers.Item(i).name))
    except Exception:
        pass
    return out


def read_temperature(handle: Any) -> float:
    return float(handle.Temperature.GetValue("C"))


def set_temperature(handle: Any, value: float) -> None:
    handle.Temperature.SetValue(float(value), "C")


def read_energy_duty(handle: Any) -> float:
    try:
        return float(handle.HeatFlow.GetValue("kJ/h")) / 3600.0
    except Exception:
        return float(handle.HeatFlowValue) / 3600.0


def init_cooling_model() -> None:
    """Establish how the cooler (E-100) cooling load is obtained.

    Spec objective:  Total = Duty_E101 + Duty_E102 - Duty_E100
    In this case file the energy stream 'E1' stores the cooling load with a
    POSITIVE sign (verified live: +0.5398 kJ/s at T_Mid1 = 20 C and ~0 at
    25 C), i.e. E1 = -Duty_E100.  Therefore:

        Total = Duty(E-2) + Duty(E3) + E1

    Priority:
      1) live energy stream 'E1' (exact literal name from Report 1),
      2) dynamic enthalpy balance  m * (h(25 C) - h(T_Mid1))  as fallback.
    """
    global COOLING_MODE, H_REF_25, MASS_FLOW, E1_HANDLE, E1_STORED

    e1h: Any = None
    try:
        e1h = fs.Streams.Item("E1")
    except Exception:
        e1h = None
    E1_HANDLE = e1h

    q_before: Optional[float] = None
    q_at25: Optional[float] = None
    if e1h is not None:
        try:
            q_before = float(e1h.HeatFlow.GetValue("kJ/h")) / 3600.0
        except Exception:
            q_before = None
        try:
            t0 = read_temperature(MID1)
            set_temperature(MID1, 25.0)
            wait_for_solver(sim)
            q_at25 = float(e1h.HeatFlow.GetValue("kJ/h")) / 3600.0
            set_temperature(MID1, t0)
            wait_for_solver(sim)
        except Exception:
            q_at25 = None
        E1_STORED = q_before
        print(f"[cooling-model] E1 heat flow: at base T_Mid1 = {q_before}; "
              f"at T_Mid1 = 25 C = {q_at25}", flush=True)
        try:
            print(f"[cooling-model] E1 AttachedOpers: {stream_attached(e1h)}", flush=True)
        except Exception:
            pass

    if (e1h is not None and q_before is not None and q_at25 is not None
            and abs(q_before - q_at25) > 1e-4):
        d_e2 = read_energy_duty(E2)
        if abs(q_at25 - d_e2) < 1e-3:
            print("[cooling-model] E1 duplicates heater duty E-2; ignored, "
                  "using enthalpy-balance cooling load instead.", flush=True)
        else:
            COOLING_MODE = "live-E1"
            print("[cooling-model] Using live energy stream 'E1' as cooling load.", flush=True)
            return

    # Fallback: dynamic enthalpy balance across the cooler (feed at 25 C)
    MASS_FLOW = float(MID1.MassFlow.GetValue("kg/s"))
    t0 = read_temperature(MID1)
    set_temperature(MID1, 25.0)
    wait_for_solver(sim)
    H_REF_25 = float(MID1.MassEnthalpy.GetValue("kJ/kg"))
    set_temperature(MID1, t0)
    wait_for_solver(sim)
    COOLING_MODE = "enthalpy-balance"
    print(f"[cooling-model] Using enthalpy-balance cooling load: "
          f"m={MASS_FLOW:.6f} kg/s, h(25 C)={H_REF_25:.6f} kJ/kg", flush=True)


def cooling_load() -> float:
    """Cooling load of E-100 as a non-negative energy term, kJ/s."""
    if COOLING_MODE == "live-E1":
        return read_energy_duty(E1_HANDLE)
    if COOLING_MODE == "enthalpy-balance":
        h = float(MID1.MassEnthalpy.GetValue("kJ/kg"))
        return MASS_FLOW * (H_REF_25 - h)
    return 0.0


def evaluate(x: np.ndarray) -> Tuple[float, np.ndarray]:
    global EVAL_COUNT
    key = (round(float(x[0]), 6), round(float(x[1]), 6), round(float(x[2]), 6))
    cached = EVAL_CACHE.get(key, "MISSING")
    if cached != "MISSING" and cached is not None:
        return cached[0], cached[1]

    EVAL_COUNT += 1
    try:
        t_mid1 = float(x[0]) * SCALE_T
        t_mid2 = float(x[1]) * SCALE_T
        t_out = float(x[2]) * SCALE_T

        set_temperature(MID1, t_mid1)
        set_temperature(MID2, t_mid2)
        set_temperature(OUT, t_out)
        wait_for_solver(sim)

        d_cool = cooling_load()          # E1 (cooling load, positive)
        d_e101 = read_energy_duty(E2)    # E-2 first heater duty
        d_e102 = read_energy_duty(E3)    # E3 second heater duty
        tm1 = read_temperature(MID1)
        tm2 = read_temperature(MID2)
        to = read_temperature(OUT)

        vals = [d_cool, d_e101, d_e102, tm1, tm2, to]
        if not all(math.isfinite(v) for v in vals):
            raise ValueError("non-finite simulator values")

        # Total = Duty_E101 + Duty_E102 - Duty_E100
        #       = Duty(E-2) + Duty(E3) + E1   (E1 = -Duty_E100 in this file)
        obj = d_e101 + d_e102 + d_cool

        cons = np.array([
            to - 80.0,     # C1: T_OutFlow >= 80
            25.0 - tm1,    # C2: T_Mid1 <= 25
            tm2 - tm1,     # C3: T_Mid2 >= T_Mid1
            to - tm2,      # C4: T_OutFlow >= T_Mid2
            95.0 - to,     # C5: T_OutFlow <= 95
        ], dtype=float)
        result: Tuple[float, np.ndarray] = (float(obj), cons)
        EVAL_CACHE[key] = result
        return result
    except Exception as exc:
        print(f"[evaluate] FAILED at x={np.round(x, 6)} : {exc}", flush=True)
        EVAL_CACHE[key] = None
        return (1e6, np.full(5, 1e3))


def f_obj(x: np.ndarray) -> float:
    return evaluate(x)[0]


def f_cons(x: np.ndarray) -> np.ndarray:
    return evaluate(x)[1]


def f_cons_jac(x: np.ndarray) -> np.ndarray:
    """Analytic Jacobian of the constraint vector w.r.t. scaled x (T = SCALE_T*x)."""
    return SCALE_T * np.array([
        [0.0, 0.0, 1.0],   # d(to - 80)/dx
        [-1.0, 0.0, 0.0],  # d(25 - tm1)/dx
        [-1.0, 1.0, 0.0],  # d(tm2 - tm1)/dx
        [0.0, -1.0, 1.0],  # d(to - tm2)/dx
        [0.0, 0.0, -1.0],  # d(95 - to)/dx
    ])


def feasible(x: np.ndarray, tol: float = 0.05) -> bool:
    _, cons = evaluate(x)
    return bool(np.all(cons >= -tol))


def cb(xk: np.ndarray) -> None:
    obj, cons = evaluate(xk)
    t = xk * SCALE_T
    print(f"[iter {EVAL_COUNT:3d}] T=({t[0]:6.3f}, {t[1]:6.3f}, {t[2]:6.3f}) C "
          f"obj={obj:8.4f} kJ/s  min_cons={float(cons.min()):8.4f}", flush=True)
    if np.all(cons >= -0.05) and obj < BEST["obj"]:
        BEST["x"] = xk.copy()
        BEST["obj"] = float(obj)
        BEST["cons"] = cons.copy()


def apply_final(x: np.ndarray) -> None:
    t_mid1 = float(x[0]) * SCALE_T
    t_mid2 = float(x[1]) * SCALE_T
    t_out = float(x[2]) * SCALE_T
    print("\n--- Applying optimal operating point to UniSim ---", flush=True)
    set_temperature(MID1, t_mid1)
    set_temperature(MID2, t_mid2)
    set_temperature(OUT, t_out)
    wait_for_solver(sim)
    tm1 = read_temperature(MID1)
    tm2 = read_temperature(MID2)
    to = read_temperature(OUT)
    d_cool = cooling_load()
    d1 = read_energy_duty(E2)
    d2 = read_energy_duty(E3)
    obj = d1 + d2 + d_cool
    print(f"  Cooling-load source: {COOLING_MODE}", flush=True)
    print(f"  T(Mid1)    = {tm1:.4f} C", flush=True)
    print(f"  T(Mid2)    = {tm2:.4f} C", flush=True)
    print(f"  T(OutFlow) = {to:.4f} C", flush=True)
    print(f"  Cooling load E1 (=-Duty_E100) = {d_cool:.6f} kJ/s", flush=True)
    print(f"  Duty E-2 (heater E-101)       = {d1:.6f} kJ/s", flush=True)
    print(f"  Duty E3  (heater E-102)       = {d2:.6f} kJ/s", flush=True)
    print(f"  TOTAL ENERGY (E-2 + E3 + E1)  = {obj:.6f} kJ/s", flush=True)
    print("  Constraint checks: "
          f"T_out>=80:{to >= 79.95}, T_mid1<=25:{tm1 <= 25.05}, "
          f"T_mid2>=T_mid1:{tm2 >= tm1 - 0.05}, T_out>=T_mid2:{to >= tm2 - 0.05}, "
          f"T_out<=95:{to <= 95.05}", flush=True)
    print(f"  Total simulator evaluations: {EVAL_COUNT}", flush=True)
    app.Visible = True
    try:
        sim.Visible = 1
    except Exception:
        pass
    print("\nOptimisation complete. UniSim remains open with the optimal point applied.", flush=True)


# ----------------------------------------------------------------------
# Main
# ----------------------------------------------------------------------
def main() -> None:
    global app, sim, fs, MID1, MID2, OUT, INFLOW, E2, E3, OP100, OP101, OP102

    print("=" * 70, flush=True)
    print("UniSim HEN optimisation: minimize (Duty_E101 + Duty_E102 - Duty_E100)", flush=True)
    print("=" * 70, flush=True)

    app = win32com.client.Dispatch("UniSimDesign.Application")
    app.Visible = True

    sim = None
    try:
        n = int(app.SimulationCases.Count)
    except Exception:
        n = 0
    if n > 0:
        for i in range(1, n + 1):
            try:
                c = app.SimulationCases.Item(i)
                p = str(c.Path).lower()
                nm = str(c.name).lower()
                if "hex.usc" in p or "hex" in nm:
                    sim = c
                    break
            except Exception:
                continue
        if sim is None:
            try:
                sim = app.ActiveDocument
            except Exception:
                sim = None
    if sim is None:
        try:
            sim = app.SimulationCases.Open(UNISIM_FILE)
        except Exception as exc:
            print(f"[warn] Open() failed: {exc}; trying ActiveDocument", flush=True)
            try:
                sim = app.ActiveDocument
            except Exception:
                sim = None
    if sim is None:
        raise RuntimeError("Unable to open or locate the UniSim case")

    try:
        sim.Visible = 1
    except Exception:
        pass
    fs = sim.Flowsheet
    wait_for_solver(sim)
    time.sleep(1.0)

    # ------------------- topology discovery -------------------
    print(f"MaterialStreams : {safe_names(fs.MaterialStreams)}", flush=True)
    print(f"EnergyStreams   : {safe_names(fs.EnergyStreams)}", flush=True)
    print(f"Streams (all)   : {safe_names(fs.Streams)}", flush=True)
    print(f"Operations      : {safe_names(fs.Operations)}", flush=True)

    _MAT: Dict[str, Any] = {}
    _EN: Dict[str, Any] = {}
    _OP: Dict[str, Any] = {}
    for coll, table in ((fs.MaterialStreams, _MAT), (fs.EnergyStreams, _EN), (fs.Operations, _OP)):
        try:
            for i in range(1, int(coll.Count) + 1):
                try:
                    obj = coll.Item(i)
                    table[str(obj.name)] = obj
                except Exception:
                    pass
        except Exception:
            pass

    MID1 = _MAT.get("Mid1")
    MID2 = _MAT.get("Mid2")
    OUT = _MAT.get("OutFlow")
    INFLOW = _MAT.get("InFlow")
    E2 = _EN.get("E-2")
    E3 = _EN.get("E3")
    OP100 = _OP.get("E-100")
    OP101 = _OP.get("E-101")
    OP102 = _OP.get("E-102")

    print("Resolved handles:", flush=True)
    for nm, h in (("Mid1", MID1), ("Mid2", MID2), ("OutFlow", OUT), ("InFlow", INFLOW),
                  ("E-2", E2), ("E3", E3),
                  ("E-100", OP100), ("E-101", OP101), ("E-102", OP102)):
        print(f"  {nm:10s} -> {None if h is None else str(h.name)}", flush=True)

    if MID1 is None or MID2 is None or OUT is None or E2 is None or E3 is None:
        raise RuntimeError("Required streams missing in case")

    for nm, h in (("Mid1", MID1), ("Mid2", MID2), ("OutFlow", OUT)):
        try:
            print(f"  Stream '{nm}' AttachedOpers: {stream_attached(h)}", flush=True)
        except Exception as exc:
            print(f"  Stream '{nm}' AttachedOpers ERR {exc}", flush=True)
    try:
        e1_probe = fs.Streams.Item("E1")
        print(f"  Stream 'E1' TypeName={e1_probe.TypeName} "
              f"AttachedOpers={stream_attached(e1_probe)}", flush=True)
    except Exception as exc:
        print(f"  Stream 'E1' probe ERR {exc}", flush=True)

    # restore nominal base-case operating point (Report 1)
    print("\nRestoring nominal base case: Mid1=20 C, Mid2=50 C, OutFlow=60 C", flush=True)
    set_temperature(MID1, 20.0)
    set_temperature(MID2, 50.0)
    set_temperature(OUT, 60.0)
    wait_for_solver(sim)

    init_cooling_model()

    # re-assert base point after the probing inside init_cooling_model
    set_temperature(MID1, 20.0)
    set_temperature(MID2, 50.0)
    set_temperature(OUT, 60.0)
    wait_for_solver(sim)

    print("\n--- Base-case values (Report 1 nominal) ---", flush=True)
    print(f"  T(Mid1)={read_temperature(MID1):.4f} C, T(Mid2)={read_temperature(MID2):.4f} C, "
          f"T(OutFlow)={read_temperature(OUT):.4f} C", flush=True)
    d_base_2 = read_energy_duty(E2)
    d_base_3 = read_energy_duty(E3)
    d_base_0 = cooling_load()
    print(f"  Duty(E-2)={d_base_2:.6f} kJ/s, Duty(E3)={d_base_3:.6f} kJ/s, "
          f"Cooling load E1={d_base_0:.6f} kJ/s", flush=True)
    print(f"  Base-case TOTAL ENERGY = {d_base_2 + d_base_3 + d_base_0:.6f} kJ/s "
          f"(Report 1 reference: 4.86 kJ/s)", flush=True)

    # ------------------- optimisation -------------------
    LB = np.array([15.0, 15.0, 80.0]) / SCALE_T
    UB = np.array([25.0, 95.0, 95.0]) / SCALE_T
    bounds = list(zip(LB, UB))
    x0 = np.array([20.0, 50.0, 80.0]) / SCALE_T

    obj0, cons0 = evaluate(x0)
    print(f"\nInitial point (C): Mid1=20, Mid2=50, OutFlow=80", flush=True)
    print(f"Initial objective = {obj0:.6f} kJ/s   min constraint = {float(cons0.min()):.4f}", flush=True)

    constraints_slsqp = [{"type": "ineq", "fun": f_cons, "jac": f_cons_jac}]

    print("\n--- SLSQP optimisation ---", flush=True)
    res = None
    try:
        res = minimize(f_obj, x0, method="SLSQP", bounds=bounds,
                       constraints=constraints_slsqp, options=SLSQP_OPTIONS,
                       callback=cb)
    except Exception as exc:
        print(f"[warn] SLSQP raised: {exc}", flush=True)

    if res is not None:
        print(f"\nSLSQP finished: success={res.success}, status={res.status}, "
              f"message={str(res.message).strip()}", flush=True)
        if res.x is not None:
            obj_r, cons_r = evaluate(res.x)
            print(f"  x (C) = {np.round(res.x * SCALE_T, 4)}", flush=True)
            print(f"  obj   = {obj_r:.6f} kJ/s", flush=True)
            print(f"  cons  = {np.round(cons_r, 6)}", flush=True)

    need_fallback = (res is None) or (res.x is None) or (not feasible(res.x)) or (not res.success)

    if need_fallback:
        print("\n--- COBYLA fallback optimisation ---", flush=True)
        start_x = res.x.copy() if (res is not None and res.x is not None) else x0.copy()
        cobyla_cons = [{"type": "ineq", "fun": (lambda xx, i=i: evaluate(xx)[1][i])} for i in range(5)]
        cobyla_cons += [{"type": "ineq", "fun": (lambda xx, i=i: xx[i] - LB[i])} for i in range(3)]
        cobyla_cons += [{"type": "ineq", "fun": (lambda xx, i=i: UB[i] - xx[i])} for i in range(3)]
        try:
            res2 = minimize(f_obj, start_x, method="COBYLA",
                            constraints=cobyla_cons, options=COBYLA_OPTIONS,
                            callback=cb)
        except Exception as exc:
            print(f"[warn] COBYLA raised: {exc}", flush=True)
            res2 = None
        if res2 is not None:
            print(f"\nCOBYLA finished: success={res2.success}, status={res2.status}, "
                  f"message={str(res2.message).strip()}", flush=True)
            if res2.x is not None:
                obj_r, cons_r = evaluate(res2.x)
                print(f"  x (C) = {np.round(res2.x * SCALE_T, 4)}", flush=True)
                print(f"  obj   = {obj_r:.6f} kJ/s", flush=True)
                print(f"  cons  = {np.round(cons_r, 6)}", flush=True)
            res = res2

    x_best = None
    if res is not None and res.x is not None and feasible(res.x):
        x_best = res.x
    elif BEST["x"] is not None:
        x_best = BEST["x"]
        print("Using best feasible point recorded during optimisation.", flush=True)
    if x_best is None:
        x_best = x0
        print("WARNING: no feasible point found; applying initial guess.", flush=True)

    apply_final(x_best)


if __name__ == "__main__":
    main()