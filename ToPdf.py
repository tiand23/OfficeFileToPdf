from __future__ import annotations

import argparse
import importlib
import json
import logging
import os
import shutil
import subprocess
import sys
import tempfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import List, Literal, Optional, Sequence, Tuple

np = None
fitz = None
win32 = None

FileKind = Literal["pdf", "excel", "word", "powerpoint"]

EXCEL_EXTENSIONS = {
    ".csv",
    ".ods",
    ".xls",
    ".xlsb",
    ".xlsm",
    ".xlsx",
}
WORD_EXTENSIONS = {
    ".doc",
    ".docm",
    ".docx",
    ".odt",
    ".rtf",
}
POWERPOINT_EXTENSIONS = {
    ".odp",
    ".ppt",
    ".pptm",
    ".pptx",
}
IMAGE_EXTENSIONS = {
    ".bmp",
    ".gif",
    ".heic",
    ".ico",
    ".jpeg",
    ".jpg",
    ".png",
    ".svg",
    ".tif",
    ".tiff",
    ".webp",
}
TEXT_EXTENSIONS = {
    ".cfg",
    ".ini",
    ".json",
    ".log",
    ".md",
    ".rst",
    ".text",
    ".toml",
    ".txt",
    ".xml",
    ".yaml",
    ".yml",
}
TOOL_EXTENSIONS = {
    ".bat",
    ".cmd",
    ".dll",
    ".exe",
    ".ps1",
    ".py",
    ".pyc",
    ".pyd",
    ".sh",
}
SKIP_DIR_NAMES = {
    ".git",
    ".idea",
    ".venv",
    ".vscode",
    "__pycache__",
    "res",
}
@dataclass(frozen=True, slots=True)
class CropConfig:
    white_thresh: int = 250
    margin_ratio: float = 0.02
    bbox_dpi: int = 150

    def __post_init__(self) -> None:
        if not (0 <= self.white_thresh <= 255):
            raise ValueError("error: white_thresh")
        if not (0.0 <= self.margin_ratio <= 1.0):
            raise ValueError("error: margin_ratio")
        if not (72 <= self.bbox_dpi <= 600):
            raise ValueError("error: bbox_dpi")


@dataclass(frozen=True, slots=True)
class RuntimeConfig:
    source_path: str = "."
    output_dir_name: str = "res"
    auto_install_missing_deps: bool = True


@dataclass(frozen=True, slots=True)
class AppConfig:
    crop: CropConfig = field(default_factory=CropConfig)
    runtime: RuntimeConfig = field(default_factory=RuntimeConfig)


@dataclass(frozen=True, slots=True)
class FileTask:
    source_path: str
    relative_path: str
    output_pdf_path: str
    kind: FileKind
    collision_note: str = ""


@dataclass(frozen=True, slots=True)
class MappingRecord:
    source_path: str
    source_relative_path: str
    output_pdf_path: str
    kind: FileKind


@dataclass(slots=True)
class ProcessSummary:
    converted: List[str] = field(default_factory=list)
    copied_without_crop: List[str] = field(default_factory=list)
    skipped: List[str] = field(default_factory=list)
    failed: List[str] = field(default_factory=list)
    collisions: List[str] = field(default_factory=list)
    mappings: List[MappingRecord] = field(default_factory=list)
    mapping_json_path: str = ""


class ToPdfError(Exception):
    pass


class DependencyError(ToPdfError):
    pass


class ExportError(ToPdfError):
    pass


class CropError(ToPdfError):
    pass


def setup_logging() -> None:
    logging.basicConfig(
        level=logging.INFO,
        format="%(message)s",
    )


def _install_python_packages(packages: Sequence[str]) -> None:
    if not packages:
        return
    cmd = [sys.executable, "-m", "pip", "install", *packages]
    subprocess.run(cmd, check=True)


def _try_import_numpy():
    global np
    if np is not None:
        return np
    try:
        np = importlib.import_module("numpy")
    except ImportError:
        np = None
    return np


def _try_import_fitz():
    global fitz
    if fitz is not None:
        return fitz

    for module_name in ("fitz", "pymupdf"):
        try:
            fitz = importlib.import_module(module_name)
            break
        except ImportError:
            fitz = None
    return fitz


def _try_import_win32():
    global win32
    if win32 is not None:
        return win32
    try:
        win32 = importlib.import_module("win32com.client")
    except ImportError:
        win32 = None
    return win32


def ensure_numpy_available(auto_install: bool) -> object:
    module = _try_import_numpy()
    if module is None and auto_install:
        _install_python_packages(["numpy"])
        module = _try_import_numpy()
    if module is None:
        raise DependencyError("error: numpy not available")
    return module


def ensure_fitz_available(auto_install: bool) -> object:
    module = _try_import_fitz()
    if module is None and auto_install:
        _install_python_packages(["PyMuPDF"])
        module = _try_import_fitz()
    if module is None:
        raise DependencyError("error: pymupdf not available")
    return module


def ensure_win32_available(auto_install: bool) -> object:
    module = _try_import_win32()
    if module is None and auto_install and sys.platform.startswith("win"):
        _install_python_packages(["pywin32"])
        module = _try_import_win32()
    if module is None:
        raise DependencyError("error: pywin32 not available")
    return module


def assert_nonempty_file(path: str, stage_name: str) -> None:
    if not os.path.isfile(path):
        raise ToPdfError(f"error: {stage_name} failed: file not found {path}")
    if os.path.getsize(path) <= 0:
        raise ToPdfError(f"error: {stage_name} failed: file is empty {path}")


def _pixmap_to_gray_array(pix: object) -> object:
    ensure_numpy_available(auto_install=False)
    h, w = pix.height, pix.width
    n = pix.n
    arr = np.frombuffer(pix.samples, dtype=np.uint8).reshape(h, w, n)
    if n >= 3:
        return arr.max(axis=2)
    return arr.squeeze()


def _content_mask(gray: object, white_thresh: int) -> Tuple[object, object]:
    nonwhite = gray < white_thresh
    return nonwhite.any(axis=1), nonwhite.any(axis=0)


def _pixel_bounds_from_mask(rows: object, cols: object, height: int, width: int) -> Tuple[int, int, int, int]:
    if not rows.any() or not cols.any():
        return 0, 0, width, height
    y0 = int(np.argmax(rows))
    y1 = int(height - 1 - np.argmax(rows[::-1]))
    x0 = int(np.argmax(cols))
    x1 = int(width - 1 - np.argmax(cols[::-1]))
    return x0, y0, x1, y1


def _apply_margin_pixels(
    x0: int,
    y0: int,
    x1: int,
    y1: int,
    width: int,
    height: int,
    margin_ratio: float,
) -> Tuple[int, int, int, int]:
    mw = max(1, int(width * margin_ratio))
    mh = max(1, int(height * margin_ratio))
    x0 = max(0, x0 - mw)
    y0 = max(0, y0 - mh)
    x1 = min(width, x1 + mw)
    y1 = min(height, y1 + mh)
    return x0, y0, x1, y1


def content_bbox_from_pixmap(
    pix: object,
    white_thresh: int = 250,
    margin_ratio: float = 0.02,
) -> Tuple[int, int, int, int]:
    h, w = pix.height, pix.width
    gray = _pixmap_to_gray_array(pix)
    rows, cols = _content_mask(gray, white_thresh)
    x0, y0, x1, y1 = _pixel_bounds_from_mask(rows, cols, h, w)
    return _apply_margin_pixels(x0, y0, x1, y1, w, h, margin_ratio)


def pixel_bbox_to_pdf_rect(
    px0: int,
    py0: int,
    px1: int,
    py1: int,
    page_rect: object,
    pix_w: int,
    pix_h: int,
) -> object:
    ensure_fitz_available(auto_install=False)
    sx = page_rect.width / pix_w
    sy = page_rect.height / pix_h
    return fitz.Rect(
        page_rect.x0 + px0 * sx,
        page_rect.y0 + py0 * sy,
        page_rect.x0 + px1 * sx,
        page_rect.y0 + py1 * sy,
    )


def crop_pdf_pages(
    input_pdf: str,
    output_pdf: str,
    config: CropConfig,
    *,
    auto_install: bool,
) -> None:
    ensure_numpy_available(auto_install)
    ensure_fitz_available(auto_install)

    if not os.path.isfile(input_pdf):
        raise FileNotFoundError(f"error: {input_pdf}")

    doc = fitz.open(input_pdf)
    out = fitz.open()

    try:
        for index in range(len(doc)):
            page = doc[index]
            rect = page.rect
            pix = page.get_pixmap(dpi=config.bbox_dpi)
            px0, py0, px1, py1 = content_bbox_from_pixmap(
                pix,
                white_thresh=config.white_thresh,
                margin_ratio=config.margin_ratio,
            )
            crop_rect = pixel_bbox_to_pdf_rect(
                px0,
                py0,
                px1,
                py1,
                rect,
                pix.width,
                pix.height,
            )
            page.set_cropbox(crop_rect)
            out.insert_pdf(doc, from_page=index, to_page=index)
        out.save(output_pdf, deflate=True)
    except Exception as exc:
        raise CropError(f"error: pdf crop failed: {exc}") from exc
    finally:
        out.close()
        doc.close()


def _dispatch_com_application(progid: str, auto_install: bool) -> object:
    client = ensure_win32_available(auto_install)
    try:
        return client.DispatchEx(progid)
    except Exception as exc:
        raise DependencyError(
            f"error: unable to start {progid}; Microsoft Office may be missing or unavailable"
        ) from exc


def _apply_page_setup_fit_to_one_page(workbook: object) -> None:
    for worksheet in workbook.Worksheets:
        page_setup = worksheet.PageSetup
        page_setup.Zoom = False
        page_setup.FitToPagesWide = 1
        page_setup.FitToPagesTall = 1


def export_excel_to_pdf_one_page_per_sheet(
    xlsx_path: str,
    output_pdf: str,
    *,
    excel_app: Optional[object] = None,
    auto_install: bool,
) -> str:
    if not os.path.isfile(xlsx_path):
        raise FileNotFoundError(f"error: {xlsx_path}")

    input_abs = os.path.abspath(xlsx_path)
    output_abs = os.path.abspath(output_pdf)
    excel = excel_app
    workbook = None
    created_locally = False

    try:
        if excel is None:
            excel = _dispatch_com_application("Excel.Application", auto_install)
            created_locally = True
        excel.Visible = False
        excel.DisplayAlerts = False
        workbook = excel.Workbooks.Open(input_abs)
        _apply_page_setup_fit_to_one_page(workbook)
        workbook.ExportAsFixedFormat(Type=0, Filename=output_abs)
        assert_nonempty_file(output_abs, "Excel 导出")
        return output_abs
    except Exception as exc:
        raise ExportError(f"error: excel export pdf failed: {exc}") from exc
    finally:
        if workbook is not None:
            try:
                workbook.Close(SaveChanges=False)
            except Exception:
                pass
        if created_locally and excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass


def export_word_to_pdf(
    doc_path: str,
    output_pdf: str,
    *,
    word_app: Optional[object] = None,
    auto_install: bool,
) -> str:
    if not os.path.isfile(doc_path):
        raise FileNotFoundError(f"error: {doc_path}")

    input_abs = os.path.abspath(doc_path)
    output_abs = os.path.abspath(output_pdf)
    word = word_app
    document = None
    created_locally = False

    try:
        if word is None:
            word = _dispatch_com_application("Word.Application", auto_install)
            created_locally = True
        word.Visible = False
        word.DisplayAlerts = 0
        document = word.Documents.Open(input_abs, ReadOnly=True, AddToRecentFiles=False)
        document.ExportAsFixedFormat(OutputFileName=output_abs, ExportFormat=17)
        assert_nonempty_file(output_abs, "Word 导出")
        return output_abs
    except Exception as exc:
        raise ExportError(f"error: word export pdf failed: {exc}") from exc
    finally:
        if document is not None:
            try:
                document.Close(SaveChanges=False)
            except Exception:
                pass
        if created_locally and word is not None:
            try:
                word.Quit()
            except Exception:
                pass


def export_powerpoint_to_pdf(
    ppt_path: str,
    output_pdf: str,
    *,
    powerpoint_app: Optional[object] = None,
    auto_install: bool,
) -> str:
    if not os.path.isfile(ppt_path):
        raise FileNotFoundError(f"error: {ppt_path}")

    input_abs = os.path.abspath(ppt_path)
    output_abs = os.path.abspath(output_pdf)
    powerpoint = powerpoint_app
    presentation = None
    created_locally = False

    try:
        if powerpoint is None:
            powerpoint = _dispatch_com_application("PowerPoint.Application", auto_install)
            created_locally = True
        try:
            powerpoint.Visible = True
        except Exception:
            pass
        presentation = powerpoint.Presentations.Open(
            FileName=input_abs,
            ReadOnly=True,
            Untitled=False,
            WithWindow=False,
        )
        presentation.SaveAs(output_abs, 32)
        assert_nonempty_file(output_abs, "PowerPoint 导出")
        return output_abs
    except Exception as exc:
        raise ExportError(f"error: powerpoint export pdf failed: {exc}") from exc
    finally:
        if presentation is not None:
            try:
                presentation.Close()
            except Exception:
                pass
        if created_locally and powerpoint is not None:
            try:
                powerpoint.Quit()
            except Exception:
                pass


@dataclass(slots=True)
class OfficeExporter:
    auto_install_missing_deps: bool
    excel_app: object = field(default=None)
    word_app: object = field(default=None)
    powerpoint_app: object = field(default=None)

    def close(self) -> None:
        for app in (self.excel_app, self.word_app, self.powerpoint_app):
            if app is None:
                continue
            try:
                app.Quit()
            except Exception:
                pass
        self.excel_app = None
        self.word_app = None
        self.powerpoint_app = None

    def export(self, source_path: str, kind: FileKind, output_pdf: str) -> str:
        if kind == "excel":
            return self._export_excel(source_path, output_pdf)
        if kind == "word":
            return self._export_word(source_path, output_pdf)
        if kind == "powerpoint":
            return self._export_powerpoint(source_path, output_pdf)
        raise ExportError(f"error: unsupported export kind {kind}")

    def _export_excel(self, source_path: str, output_pdf: str) -> str:
        if self.excel_app is None:
            self.excel_app = _dispatch_com_application(
                "Excel.Application",
                self.auto_install_missing_deps,
            )
            self.excel_app.Visible = False
            self.excel_app.DisplayAlerts = False
        return export_excel_to_pdf_one_page_per_sheet(
            source_path,
            output_pdf,
            excel_app=self.excel_app,
            auto_install=self.auto_install_missing_deps,
        )

    def _export_word(self, source_path: str, output_pdf: str) -> str:
        if self.word_app is None:
            self.word_app = _dispatch_com_application(
                "Word.Application",
                self.auto_install_missing_deps,
            )
            self.word_app.Visible = False
            self.word_app.DisplayAlerts = 0
        return export_word_to_pdf(
            source_path,
            output_pdf,
            word_app=self.word_app,
            auto_install=self.auto_install_missing_deps,
        )

    def _export_powerpoint(self, source_path: str, output_pdf: str) -> str:
        if self.powerpoint_app is None:
            self.powerpoint_app = _dispatch_com_application(
                "PowerPoint.Application",
                self.auto_install_missing_deps,
            )
        return export_powerpoint_to_pdf(
            source_path,
            output_pdf,
            powerpoint_app=self.powerpoint_app,
            auto_install=self.auto_install_missing_deps,
        )


def _should_skip_directory(path: Path) -> bool:
    name = path.name
    if name in SKIP_DIR_NAMES:
        return True
    return name.startswith(".")


def classify_input_file(path: Path) -> Tuple[Optional[FileKind], Optional[str]]:
    suffix = path.suffix.lower()

    if suffix == ".pdf":
        return "pdf", None
    if suffix in EXCEL_EXTENSIONS:
        return "excel", None
    if suffix in WORD_EXTENSIONS:
        return "word", None
    if suffix in POWERPOINT_EXTENSIONS:
        return "powerpoint", None
    if suffix in IMAGE_EXTENSIONS:
        return None, "图片文件已跳过"
    if suffix in TEXT_EXTENSIONS:
        return None, "文本文件已跳过"
    if suffix in TOOL_EXTENSIONS:
        return None, "工具文件已跳过"
    if not suffix:
        return None, "无扩展名文件已跳过"
    return None, f"暂不支持的文件类型: {suffix}"


def _unique_output_relative_path(relative_path: Path, used: set[str]) -> Tuple[Path, str]:
    base_candidate = relative_path.with_suffix(".pdf")
    candidate = base_candidate
    collision_note = ""
    source_suffix = relative_path.suffix.lower().lstrip(".") or "file"
    counter = 1

    while candidate.as_posix().lower() in used:
        suffix = f"__from_{source_suffix}"
        if counter > 1:
            suffix = f"{suffix}_{counter}"
        candidate = base_candidate.with_name(f"{base_candidate.stem}{suffix}.pdf")
        counter += 1

    used.add(candidate.as_posix().lower())
    if candidate != base_candidate:
        collision_note = (
            f"{relative_path.as_posix()} -> {candidate.as_posix()} "
            "(因同名 PDF 冲突自动改名)"
        )
    return candidate, collision_note


def discover_file_tasks(source_path: str, output_dir_name: str) -> Tuple[List[FileTask], List[str], List[str]]:
    source = Path(source_path).resolve()
    if not source.exists():
        raise FileNotFoundError(f"error: {source}")

    used_output_paths: set[str] = set()
    tasks: List[FileTask] = []
    skipped: List[str] = []
    collisions: List[str] = []

    def _register_file(file_path: Path, root_path: Path) -> None:
        kind, skip_reason = classify_input_file(file_path)
        relative_path = file_path.relative_to(root_path)

        if kind is None:
            skipped.append(f"{relative_path.as_posix()} | {skip_reason}")
            return

        output_relative, collision_note = _unique_output_relative_path(relative_path, used_output_paths)
        output_pdf_path = root_path / output_dir_name / output_relative
        if collision_note:
            collisions.append(collision_note)
        tasks.append(
            FileTask(
                source_path=str(file_path),
                relative_path=relative_path.as_posix(),
                output_pdf_path=str(output_pdf_path),
                kind=kind,
                collision_note=collision_note,
            )
        )

    if source.is_file():
        _register_file(source, source.parent)
        return tasks, skipped, collisions

    for current_root, dir_names, file_names in os.walk(source):
        current_path = Path(current_root)
        dir_names[:] = [
            name
            for name in sorted(dir_names)
            if not _should_skip_directory(current_path / name)
        ]
        for file_name in sorted(file_names):
            file_path = current_path / file_name
            if file_path.name.startswith("."):
                skipped.append(f"{file_path.relative_to(source).as_posix()} | 隐藏文件已跳过")
                continue
            _register_file(file_path, source)

    return tasks, skipped, collisions


def build_temp_pdf_path(source_path: str) -> str:
    stem = Path(source_path).stem
    with tempfile.NamedTemporaryFile(prefix=f"{stem}_", suffix=".pdf", delete=False) as temp_file:
        return temp_file.name


def safe_crop_or_copy(
    input_pdf: str,
    output_pdf: str,
    crop_config: CropConfig,
    *,
    auto_install: bool,
) -> bool:
    os.makedirs(os.path.dirname(output_pdf), exist_ok=True)
    try:
        crop_pdf_pages(
            input_pdf,
            output_pdf,
            crop_config,
            auto_install=auto_install,
        )
        return True
    except Exception as exc:
        shutil.copy2(input_pdf, output_pdf)
        logging.getLogger(__name__).warning(
            "裁边失败，已回退为直接复制 PDF: %s | %s",
            input_pdf,
            exc,
        )
        return False


def process_file_task(task: FileTask, config: AppConfig, exporter: OfficeExporter, summary: ProcessSummary) -> None:
    logger = logging.getLogger(__name__)
    os.makedirs(os.path.dirname(task.output_pdf_path), exist_ok=True)

    logger.info("处理中: %s", task.relative_path)
    if task.collision_note:
        logger.warning("输出文件名冲突，已改名: %s", task.collision_note)

    if task.kind == "pdf":
        cropped = safe_crop_or_copy(
            task.source_path,
            task.output_pdf_path,
            config.crop,
            auto_install=config.runtime.auto_install_missing_deps,
        )
        summary.converted.append(task.output_pdf_path)
        summary.mappings.append(
            MappingRecord(
                source_path=os.path.abspath(task.source_path),
                source_relative_path=task.relative_path,
                output_pdf_path=os.path.abspath(task.output_pdf_path),
                kind=task.kind,
            )
        )
        if not cropped:
            summary.copied_without_crop.append(task.output_pdf_path)
        return

    temp_pdf = build_temp_pdf_path(task.source_path)
    try:
        exporter.export(task.source_path, task.kind, temp_pdf)
        cropped = safe_crop_or_copy(
            temp_pdf,
            task.output_pdf_path,
            config.crop,
            auto_install=config.runtime.auto_install_missing_deps,
        )
        summary.converted.append(task.output_pdf_path)
        summary.mappings.append(
            MappingRecord(
                source_path=os.path.abspath(task.source_path),
                source_relative_path=task.relative_path,
                output_pdf_path=os.path.abspath(task.output_pdf_path),
                kind=task.kind,
            )
        )
        if not cropped:
            summary.copied_without_crop.append(task.output_pdf_path)
    finally:
        if os.path.exists(temp_pdf):
            try:
                os.remove(temp_pdf)
            except Exception:
                pass


def write_mapping_json(summary: ProcessSummary, source_root: str, output_root: str) -> str:
    mapping_json_path = os.path.join(output_root, "source_to_pdf_map.json")
    payload = {
        "source_root": os.path.abspath(source_root),
        "output_root": os.path.abspath(output_root),
        "mappings": [
            {
                "source_path": record.source_path,
                "source_relative_path": record.source_relative_path,
                "output_pdf_path": record.output_pdf_path,
                "output_relative_path": Path(record.output_pdf_path).resolve().relative_to(Path(output_root).resolve()).as_posix(),
                "kind": record.kind,
            }
            for record in summary.mappings
        ],
    }
    with open(mapping_json_path, "w", encoding="utf-8") as handle:
        json.dump(payload, handle, ensure_ascii=False, indent=2)
    summary.mapping_json_path = mapping_json_path
    return mapping_json_path


def process_directory(config: AppConfig) -> ProcessSummary:
    logger = logging.getLogger(__name__)
    source = Path(config.runtime.source_path).resolve()
    source_path = str(source)
    output_base = source.parent if source.is_file() else source
    output_root = str(output_base / config.runtime.output_dir_name)
    os.makedirs(output_root, exist_ok=True)

    tasks, skipped, collisions = discover_file_tasks(
        source_path,
        config.runtime.output_dir_name,
    )

    summary = ProcessSummary(
        skipped=skipped,
        collisions=collisions,
    )

    logger.info("源目录: %s", source_path)
    logger.info("输出目录: %s", output_root)
    logger.info("待处理文件: %d", len(tasks))
    logger.info("跳过文件: %d", len(skipped))

    exporter = OfficeExporter(auto_install_missing_deps=config.runtime.auto_install_missing_deps)
    try:
        for index, task in enumerate(tasks, start=1):
            logger.info("[%d/%d]", index, len(tasks))
            try:
                process_file_task(task, config, exporter, summary)
            except Exception as exc:
                message = f"{task.relative_path} | {exc}"
                summary.failed.append(message)
                logger.exception("处理失败: %s", task.relative_path)
    finally:
        exporter.close()

    write_mapping_json(summary, source_path, output_root)
    return summary


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="递归扫描目录，将支持的 Office/PDF 文件输出到 res 目录并统一转换为 PDF。",
    )
    parser.add_argument(
        "source",
        nargs="?",
        default=".",
        help="要处理的根目录；默认是当前目录。",
    )
    parser.add_argument(
        "--output-dir",
        default="res",
        help="输出目录名称，默认 res。",
    )
    parser.add_argument(
        "--white-thresh",
        type=int,
        default=250,
        help="低于此灰度值视为内容，默认 250。",
    )
    parser.add_argument(
        "--margin",
        type=float,
        default=0.02,
        help="保留白边比例，默认 0.02。",
    )
    parser.add_argument(
        "--dpi",
        type=int,
        default=150,
        help="检测内容边界时使用的渲染 DPI，默认 150。",
    )
    parser.add_argument(
        "--no-auto-install",
        action="store_true",
        help="关闭缺失 Python 依赖的自动安装。",
    )
    return parser


def build_config_from_args(args: argparse.Namespace) -> AppConfig:
    return AppConfig(
        crop=CropConfig(
            white_thresh=args.white_thresh,
            margin_ratio=args.margin,
            bbox_dpi=args.dpi,
        ),
        runtime=RuntimeConfig(
            source_path=args.source,
            output_dir_name=args.output_dir,
            auto_install_missing_deps=not args.no_auto_install,
        ),
    )


def log_summary(summary: ProcessSummary) -> None:
    logger = logging.getLogger(__name__)
    logger.info("")
    logger.info("处理完成")
    logger.info("成功输出 PDF: %d", len(summary.converted))
    logger.info("直接复制 PDF 回退次数: %d", len(summary.copied_without_crop))
    logger.info("跳过文件: %d", len(summary.skipped))
    logger.info("失败文件: %d", len(summary.failed))
    if summary.mapping_json_path:
        logger.info("映射 JSON: %s", summary.mapping_json_path)

    if summary.collisions:
        logger.info("")
        logger.info("命名冲突自动改名:")
        for item in summary.collisions:
            logger.info("  %s", item)

    if summary.skipped:
        logger.info("")
        logger.info("已跳过文件:")
        for item in summary.skipped:
            logger.info("  %s", item)

    if summary.failed:
        logger.info("")
        logger.info("处理失败文件:")
        for item in summary.failed:
            logger.info("  %s", item)


def main(argv: Optional[Sequence[str]] = None) -> int:
    setup_logging()
    parser = build_arg_parser()
    args = parser.parse_args(argv)
    config = build_config_from_args(args)

    try:
        summary = process_directory(config)
        log_summary(summary)
    except Exception as exc:
        logging.getLogger(__name__).exception("执行失败: %s", exc)
        return 1

    if summary.failed:
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
