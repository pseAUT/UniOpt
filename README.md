# UniOpt

**An Experimental Agentic AI Framework Employing Large Language Models for Automated Chemical Process Optimization**

UniOpt is a multi-agent LLM pipeline that takes a converged steady-state **UniSim Design** simulation (XML export + PFD screenshot) and a natural-language optimization request, then automatically:

1. Interprets the process (textual + visual)
2. Formulates a continuous-variable mathematical optimization problem
3. Generates executable Python code that drives the original `.usc` file via the COM interface

---

## Framework Overview

```
┌─────────────────┐     ┌─────────────────┐
│  XML Interpreter│     │  PFD Interpreter│
│  (DeepSeek-V4-  │     │  (Qwen3.5-397B) │
│   Flash)        │     │                 │
└────────┬────────┘     └────────┬────────┘
         │                       │
         └───────────┬───────────┘
                     ▼
          ┌─────────────────────┐
          │ Optimization        │
          │ Formulator Agent    │
          │ (DeepSeek-V4-Pro,   │
          │  max reasoning)     │
          └──────────┬──────────┘
                     ▼
          ┌─────────────────────┐
          │ Coder Agent         │
          │ (DeepSeek-V4-Pro)   │
          └──────────┬──────────┘
                     ▼
              Python + COM
              
```

### Agents

| Agent                  | Model                  | Role                                                                 |
|------------------------|------------------------|----------------------------------------------------------------------|
| XML Interpreter        | DeepSeek-V4-Flash      | Distills UniSim XML → process report (equipment, streams, topology) |
| PFD Interpreter        | Qwen3.5-397B-A17B      | Vision-based cross-check of equipment ordering and connectivity     |
| Optimization Formulator| DeepSeek-V4-Pro (max)  | Builds objective, decision variables, constraints, solver settings  |
| Coder                  | DeepSeek-V4-Pro (max)  | Generates and debugs Python/COM optimization code                   |

All instruction prompts are **static system prompts**.

---

## Case Studies

Three steady-state examples are provided:

1. **Heat-exchangers in series**  
   Minimize total energy duty while enforcing T= 80° C.  
   The cooler duty is correctly driven to zero.

2. **Three-stage nitrogen compression**  
   Intermediate pressures compared with the classical equal-pressure-ratio rule (Edgar et al., 2001).  
   Gap < 1 %.

3. **Ethyl chloride manufacturing (with recycle)**  
   Maximize venture profit by adjusting purge rate.  
   Result matches Seider et al. (2017) within 0.11 %.


   ## Requirements

- Python 3.10+
- Windows (UniSim Design COM interface)
- Honeywell UniSim Design (licensed)
- OpenAI-compatible API access to:
  - DeepSeek-V4-Flash
  - DeepSeek-V4-Pro
  - Qwen3.5-397B-A17B (via SiliconFlow or equivalent)
