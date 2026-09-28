import warnings
warnings.filterwarnings("ignore", category=SyntaxWarning, module=r"^ofdparser(\.|$)")
warnings.filterwarnings("ignore", category=SyntaxWarning, module=r"^easyofd(\.|$)")

import argparse
import logging
import os

from readInvoice import collect_pdfs, read_document
from layoutInvoice import two_up_vertical, write_writer
from printInvoice import print_pdf
from gui import run_gui

logger = logging.getLogger(__name__)

def process(input_path: str, output_dir: str | None, do_print: bool) -> None:
    pdfs = collect_pdfs(input_path)
    for src in pdfs:
        reader = read_document(src)
        writer = two_up_vertical(reader)
        name = os.path.splitext(os.path.basename(src))[0] + "_2up.pdf"
        out_dir = output_dir or os.path.dirname(src)
        out_path = os.path.join(out_dir, name)
        write_writer(writer, out_path)
        logger.info("已生成 %s", out_path)
        if do_print:
            try:
                print_pdf(out_path)
            except Exception as e:
                logger.error("打印失败 %s: %s", out_path, e)

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
    process(args.input, args.output, not args.no_print)

if __name__ == "__main__":
    main()
