import warnings
warnings.filterwarnings("ignore", category=SyntaxWarning, module=r"^ofdparser(\.|$)")
warnings.filterwarnings("ignore", category=SyntaxWarning, module=r"^easyofd(\.|$)")

import argparse
import logging

import pipeline
from gui import run_gui

logger = logging.getLogger(__name__)


def main() -> None:
    logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
    ap = argparse.ArgumentParser(
        description="发票排版与打印：批量 2-up（上下）合成 PDF/OFD/XML，--gui 进入图形界面"
    )
    ap.add_argument("-i", "--input")
    ap.add_argument("-o", "--output")
    ap.add_argument("--no-print", action="store_true")
    ap.add_argument("--gui", action="store_true")
    args = ap.parse_args()
    if args.gui or not args.input:
        run_gui()
        return
    # 批处理编排（读取 → 2-up 合成 → 写出 → 打印）在 pipeline.process
    pipeline.process(args.input, args.output, not args.no_print)


if __name__ == "__main__":
    main()
