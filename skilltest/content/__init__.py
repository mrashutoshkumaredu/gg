"""All question papers and drills."""
import importlib.util
import sys

from . import mock1, mock2, mock3, mock4, mock5, variants

_MODULES = {"mock1": mock1, "mock2": mock2, "mock3": mock3, "mock4": mock4, "mock5": mock5}
PAPERS = [m.PAPER for m in _MODULES.values()]


def build_paper(paper_id, seed=0):
    """The paper with fresh names and numbers for this seed (seed 0 = the original paper).
    The module is loaded again as a separate copy, so its marking always uses its own numbers."""
    if not seed:
        return next(p for p in PAPERS if p.id == paper_id)
    mod = _MODULES[paper_id]
    name = f"{__name__}.{paper_id}_v{seed}"
    if name in sys.modules:
        return sys.modules[name].PAPER
    spec = importlib.util.spec_from_file_location(name, mod.__file__)
    copy = importlib.util.module_from_spec(spec)
    copy.__package__ = __name__
    variants.set_seed(seed)
    try:
        spec.loader.exec_module(copy)
    finally:
        variants.set_seed(0)
    sys.modules[name] = copy
    return copy.PAPER
