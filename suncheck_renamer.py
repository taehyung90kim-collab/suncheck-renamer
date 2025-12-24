import json
import re
import sys
import logging
from pathlib import Path

from pypdf import PdfReader
from tkinterdnd2 import TkinterDnD, DND_FILES
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter.scrolledtext import ScrolledText


# =========================
# exe/py 공통: "프로그램 폴더" 기준 경로
# - py 실행: 이 파일이 있는 폴더
# - exe 실행(PyInstaller --onefile): exe가 있는 폴더
# =========================
def app_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


BASE_DIR = app_dir()

# 기본 Input/Output (프로그램 폴더 기준)
DEFAULT_INPUT_DIR = BASE_DIR / "Input"
DEFAULT_OUTPUT_DIR = BASE_DIR / "Output"

# 설정 파일(출력폴더 기억)
CONFIG_PATH = BASE_DIR / "config.json"

# =========================
# 로그 파일 설정 (exe/py 공통)
# =========================
LOG_PATH = BASE_DIR / "suncheck_renamer.log"
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(LOG_PATH, encoding="utf-8"),
        logging.StreamHandler(),  # 터미널 실행 시에도 보이도록
    ],
)
logger = logging.getLogger("suncheck_renamer")


# =========================
# (아이콘) 코드로 그리기
# =========================
def draw_pdf_icon(c: tk.Canvas, x: int, y: int, scale=1.0, color="#b8c2dc"):
    w, h = int(70 * scale), int(90 * scale)
    left, top = x - w // 2, y - h // 2
    right, bottom = left + w, top + h

    c.create_rectangle(left, top, right, bottom, outline=color, width=2, fill="#f7f9ff")

    fold = int(18 * scale)
    c.create_polygon(
        right - fold, top,
        right, top,
        right, top + fold,
        fill="#e6ebf7", outline=color, width=2
    )

    for i in range(3):
        yy = top + int((28 + i * 12) * scale)
        c.create_line(
            left + int(12 * scale), yy,
            right - int(12 * scale), yy,
            fill="#dfe6f6", width=2
        )

    c.create_rectangle(
        left + int(10 * scale), bottom - int(28 * scale),
        right - int(10 * scale), bottom - int(10 * scale),
        fill="#e1e7f6", outline=""
    )
    c.create_text(
        x, bottom - int(19 * scale),
        text="PDF",
        font=("Helvetica", int(14 * scale), "bold"),
        fill="#7a86a8"
    )


def draw_folder_pdf_icon(c: tk.Canvas, x: int, y: int, scale=1.0, color="#b8c2dc"):
    fw, fh = int(110 * scale), int(70 * scale)
    left, top = x - fw // 2, y - fh // 2
    right, bottom = left + fw, top + fh

    c.create_rectangle(
        left + int(10 * scale), top - int(18 * scale),
        left + int(54 * scale), top,
        outline=color, width=2, fill="#f7f9ff"
    )

    c.create_rectangle(left, top, right, bottom, outline=color, width=2, fill="#f7f9ff")

    draw_pdf_icon(c, x, y + int(8 * scale), scale=0.55 * scale, color=color)


# =========================
# 설정 로드/저장
# =========================
def load_config():
    if CONFIG_PATH.exists():
        try:
            cfg = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
            if isinstance(cfg, dict):
                return cfg
        except Exception:
            pass
    return {"output_dir": str(DEFAULT_OUTPUT_DIR)}


def save_config(cfg: dict):
    CONFIG_PATH.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")


CFG = load_config()


# =========================
# PDF 텍스트 추출 / 파싱
# =========================
PID_RE = re.compile(r"Patient ID:\s*([A-Za-z]*\d{5,9})")
NAME_RE = re.compile(r"Patient Name:\s*([^\n\r]+)")
PLAN_RE = re.compile(r"Plan Name:\s*([^\n\r]+)")
GAMMA_RE = re.compile(r"Diff\s*\(%\)\s*:\s*(\d+)\s*Dist\s*\(mm\)\s*:\s*(\d+)", re.IGNORECASE)


def extract_text(pdf: Path) -> str:
    reader = PdfReader(str(pdf))
    return "\n".join((p.extract_text() or "") for p in reader.pages)


def clean(s: str) -> str:
    return re.sub(r'[\\/:*?"<>|]', "", s.strip())


def make_unique(path: Path) -> Path:
    if not path.exists():
        return path
    stem = path.stem
    for i in range(1, 10000):
        candidate = path.with_name(f"{stem}({i}){path.suffix}")
        if not candidate.exists():
            return candidate
    return path


def process_pdf(pdf: Path, output_dir: Path) -> Path:
    text = extract_text(pdf)

    pid_m = PID_RE.search(text)
    name_m = NAME_RE.search(text)
    plan_m = PLAN_RE.search(text)
    gamma_m = GAMMA_RE.search(text)

    if not pid_m:
        raise ValueError("Patient ID를 찾지 못함")
    if not name_m:
        raise ValueError("Patient Name을 찾지 못함")
    if not plan_m:
        raise ValueError("Plan Name을 찾지 못함")
    if not gamma_m:
        raise ValueError("Diff(%)/Dist(mm)를 찾지 못함")

    pid = pid_m.group(1)
    name = name_m.group(1)
    plan = plan_m.group(1)

    diff = gamma_m.group(1)
    dist = gamma_m.group(2)
    gamma = f"_{diff}%{dist}mm"

    new_name = f"{clean(pid)}_{clean(name)}_{clean(plan)}{gamma}.pdf"

    output_dir.mkdir(parents=True, exist_ok=True)
    out_path = make_unique(output_dir / new_name)
    out_path.write_bytes(pdf.read_bytes())
    return out_path


# =========================
# GUI
# =========================
class App(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()
        self.title("SunCHECK PDF Renamer")
        self.geometry("920x420")

        # 폴더 자동 생성
        DEFAULT_INPUT_DIR.mkdir(exist_ok=True)
        DEFAULT_OUTPUT_DIR.mkdir(exist_ok=True)

        # output_dir 로드(없으면 기본 Output)
        self.output_dir = Path(CFG.get("output_dir", str(DEFAULT_OUTPUT_DIR)))
        self.output_dir.mkdir(exist_ok=True)

        # ----- 상단 -----
        top = tk.Frame(self)
        top.pack(fill="x", padx=10, pady=10)

        tk.Label(top, text="Output folder: ").pack(side="left")

        self.out_label = tk.Label(
            top,
            text=self.output_dir_str(),
            fg="blue",
            cursor="hand2",
            anchor="w"
        )
        self.out_label.pack(side="left", fill="x", expand=True)
        self.out_label.bind("<Button-1>", lambda e: self.choose_output())

        tk.Button(top, text="Change...", command=self.choose_output).pack(side="right")

        tk.Label(
            self,
            text="PDF 또는 폴더를 아래 영역에 드래그 & 드롭하세요",
            justify="left"
        ).pack(fill="x", padx=10)

        # ----- 드롭존 -----
        self.drop_canvas = tk.Canvas(self, height=240, highlightthickness=0, bg="#f3f6fb")
        self.drop_canvas.pack(fill="x", padx=10, pady=10)

        draw_pdf_icon(self.drop_canvas, 460 - 90, 100, scale=1.0)
        draw_folder_pdf_icon(self.drop_canvas, 460 + 90, 100, scale=1.0)

        self.drop_canvas.create_text(
            460, 175,
            text="Drop PDF or Folder Here",
            font=("Helvetica", 22, "bold"),
            fill="#c0c8dd"
        )
        self.drop_canvas.create_text(
            460, 205,
            text="PDF / 폴더 모두 지원",
            font=("Helvetica", 13),
            fill="#c0c8dd"
        )

        self.drop_canvas.drop_target_register(DND_FILES)
        self.drop_canvas.dnd_bind("<<Drop>>", self.on_drop)

        # ----- 로그 -----
        self.log = ScrolledText(self, height=8)
        self.log.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        # 시작 로그
        self.log_line(f"[INFO] App folder   : {BASE_DIR}")
        self.log_line(f"[INFO] Input folder : {DEFAULT_INPUT_DIR}")
        self.log_line(f"[INFO] Output folder: {self.output_dir}")
        self.log_line(f"[INFO] Log file     : {LOG_PATH}")

        logger.info("===== SunCHECK PDF Renamer started =====")
        logger.info(f"sys.version   : {sys.version}")
        logger.info(f"sys.executable: {sys.executable}")
        logger.info(f"BASE_DIR      : {BASE_DIR}")
        logger.info(f"INPUT_DIR     : {DEFAULT_INPUT_DIR}")
        logger.info(f"OUTPUT_DIR    : {self.output_dir}")
        logger.info(f"LOG_PATH      : {LOG_PATH}")

    def output_dir_str(self):
        return str(self.output_dir) if self.output_dir else "(click to set)"

    def log_line(self, msg: str):
        self.log.insert("end", msg + "\n")
        self.log.see("end")
        # 파일에도 같이 남김
        if msg.startswith("[ERR]"):
            logger.error(msg)
        else:
            logger.info(msg)

    def choose_output(self):
        d = filedialog.askdirectory(title="Output 폴더 선택")
        if not d:
            return
        self.output_dir = Path(d)
        self.output_dir.mkdir(parents=True, exist_ok=True)

        CFG["output_dir"] = str(self.output_dir)
        save_config(CFG)

        self.out_label.config(text=self.output_dir_str())
        self.log_line(f"[SET] Output 폴더: {self.output_dir}")

    def on_drop(self, event):
        try:
            if self.output_dir is None:
                self.log_line("[ERR] Output 폴더가 지정되지 않았습니다.")
                return

            items = self.tk.splitlist(event.data)
            self.log_line(f"[INFO] Drop items: {len(items)}")

            for item in items:
                path = Path(item)

                if path.is_dir():
                    pdfs = sorted(path.glob("*.pdf"))
                    self.log_line(f"[DIR] {path} ({len(pdfs)} PDFs)")
                    for pdf in pdfs:
                        self.handle_pdf(pdf)
                else:
                    self.handle_pdf(path)

        except Exception as e:
            self.log_line(f"[ERR] Drop 처리 중 예외: {e}")
            logger.exception("Drop 처리 중 예외(Traceback 포함)")

    def handle_pdf(self, pdf: Path):
        if pdf.suffix.lower() != ".pdf":
            self.log_line(f"[SKIP] PDF 아님: {pdf.name}")
            return
        try:
            out = process_pdf(pdf, self.output_dir)
            self.log_line(f"[OK] {pdf.name} → {out.name}")
        except Exception as e:
            self.log_line(f"[ERR] {pdf.name}: {e}")
            logger.exception(f"PDF 처리 실패: {pdf}")


if __name__ == "__main__":
    try:
        App().mainloop()
    except Exception:
        logger.exception("프로그램 전체 크래시 (최상위 예외)")
        raise