# macro/

Each script in this folder is a standalone macro data pull, named `<topic>_pull.py`, writing outputs to `output/<topic>/`.

For example, `yield_curve_pull.py` writes to `output/yield_curve/`. Companion scripts on the same topic (e.g. `yield_curve_chart.py`) write to the same `output/<topic>/` folder.

Run a script directly from the repo root:

```
python macro/yield_curve_pull.py
```

Outputs are gitignored — only `.gitkeep` markers are tracked.
