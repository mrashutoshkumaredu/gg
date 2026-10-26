"""Fresh names and numbers for every retake.

Attempt 1 of a test always uses the original data (seed 0). Later attempts use another seed, so names and numbers
change while the tasks stay the same. Expected answers are always worked out from the data, so marking follows.
"""
import random

SEED = 0
_rnd = random.Random(0)

FIRST_NAMES = ["Aarav", "Bhavna", "Chetan", "Divya", "Eshan", "Farah", "Gopal", "Harini", "Imran", "Jyoti", "Kiran",
               "Lata", "Manoj", "Nisha", "Omkar", "Pallavi", "Qasim", "Rohit", "Sneha", "Tanvi", "Uday", "Varsha",
               "Waseem", "Yamini", "Zubin", "Arjun", "Bindu", "Charu", "Dhruv", "Ekta", "Gauri", "Hemant", "Ishaan",
               "Kavya", "Lokesh", "Meena", "Naveen", "Pooja", "Rahul", "Swati", "Tarun", "Usha", "Vivek"]
SURNAMES = ["Kumar", "Sharma", "Das", "Nair", "Babu", "Rao", "Singh", "Iyer", "Patel", "Menon", "Reddy", "Joshi"]


def set_seed(seed):
    global SEED, _rnd
    SEED = seed or 0
    _rnd = random.Random(SEED)


def active():
    return SEED != 0


def num(n, spread=0.25, step=1, low=None, high=None):
    """n itself on attempt 1, otherwise a nearby number."""
    if not SEED:
        return n
    lo = n * (1 - spread)
    hi = n * (1 + spread)
    v = _rnd.uniform(lo, hi)
    v = int(round(v / step) * step) if isinstance(n, int) else round(v, 2)
    if low is not None:
        v = max(low, v)
    if high is not None:
        v = min(high, v)
    return v


def choice(options, default):
    return default if not SEED else _rnd.choice(options)


def names(default, full=False):
    """Same count of unique names; first names only unless full=True."""
    if not SEED:
        return list(default)
    pool = FIRST_NAMES[:]
    _rnd.shuffle(pool)
    out = pool[:len(default)]
    if full:
        out = [f"{n} {_rnd.choice(SURNAMES)}" for n in out]
    return out


def shuffled(items):
    items = list(items)
    if SEED:
        _rnd.shuffle(items)
    return items


def unique_totals(make, key, tries=50):
    """Call make() until key(rows) gives no duplicates (needed for sorting tasks)."""
    rows = make()
    for _ in range(tries):
        vals = [key(r) for r in rows]
        if len(set(vals)) == len(vals):
            return rows
        rows = make()
    return rows
