"""Build the editable deck from per-slide specs in slides/sNNN.py.

usage: python3 build.py OUT.pptx [first last]
"""
import importlib.util
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from lib import Slide, new_deck  # noqa: E402

HERE = os.path.dirname(os.path.abspath(__file__))


def main():
    out = sys.argv[1]
    first, last = (int(sys.argv[2]), int(sys.argv[3])) if len(sys.argv) > 3 else (1, 999)
    prs = new_deck()
    for n in range(first, last + 1):
        p = os.path.join(HERE, "slides", "s%03d.py" % n)
        if not os.path.exists(p):
            continue
        spec = importlib.util.spec_from_file_location("s%03d" % n, p)
        mod = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(mod)
        mod.build(Slide(prs, n))
    prs.save(out)
    print("saved", out, len(prs.slides._sldIdLst), "slides")


if __name__ == "__main__":
    main()
