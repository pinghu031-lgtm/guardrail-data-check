# -*- coding: utf-8 -*-
"""云端上传批处理服务：校验、隔离处理并返回下载文件。"""

from dataclasses import dataclass
from io import BytesIO
from pathlib import Path
from tempfile import TemporaryDirectory
from typing import Any, Dict, List, Sequence
from zipfile import BadZipFile, ZIP_DEFLATED, ZipFile
import hashlib
import os
import re

from guardrail_core import (
    APP_VERSION,
    InspectionProcessor,
    IssueGroup,
    export_groups_to_excel,
    export_modified_workbooks,
)

SUPPORTED_EXTENSIONS = {".xlsx", ".xlsm"}
DEFAULT_MAX_FILES = 20
DEFAULT_MAX_FILE_BYTES = 50 * 1024 * 1024
DEFAULT_MAX_TOTAL_BYTES = 200 * 1024 * 1024
DEFAULT_MAX_ARCHIVE_ENTRIES = 10_000
DEFAULT_MAX_UNCOMPRESSED_BYTES = 250 * 1024 * 1024
DEFAULT_MAX_TOTAL_UNCOMPRESSED_BYTES = 500 * 1024 * 1024


class UploadValidationError(ValueError):
    """上传文件不符合公开服务安全限制。"""


@dataclass(frozen=True)
class UploadedFileData:
    name: str
    content: bytes


@dataclass(frozen=True)
class ProcessResult:
    result_excel: bytes
    modified_zip: bytes
    result_filename: str
    modified_filename: str
    preview_rows: List[Dict[str, Any]]
    stats: Dict[str, Any]
    modification_stats: Dict[str, Any]
    logs: List[str]


class ListLogger:
    def __init__(self) -> None:
        self.lines: List[str] = []

    def __call__(self, message: str) -> None:
        self.lines.append(message)


def upload_fingerprint(uploads: Sequence[UploadedFileData]) -> str:
    """生成当前上传批次的稳定指纹，用于避免展示与新文件不匹配的旧结果。"""
    digest = hashlib.sha256()
    for upload in uploads:
        digest.update(str(upload.name).encode("utf-8", errors="surrogatepass"))
        digest.update(b"\0")
        digest.update(upload.content)
        digest.update(b"\0")
    return digest.hexdigest()


def safe_filename(name: str) -> str:
    """删除客户端路径和控制字符，只保留一个可写入临时目录的文件名。"""
    leaf = str(name or "").replace("\\", "/").split("/")[-1].strip()
    leaf = re.sub(r"[\x00-\x1f<>:\"/\\|?*]", "_", leaf)
    leaf = leaf.rstrip(". ")
    if leaf in {"", ".", ".."}:
        raise UploadValidationError("上传文件名无效。")
    return leaf


def _deduplicate_name(name: str, used_names: set[str]) -> str:
    stem, suffix = os.path.splitext(name)
    candidate = name
    index = 2
    while candidate.casefold() in used_names:
        candidate = f"{stem}_{index}{suffix}"
        index += 1
    used_names.add(candidate.casefold())
    return candidate


def validate_upload_metadata(
    file_sizes: Sequence[int],
    *,
    max_files: int = DEFAULT_MAX_FILES,
    max_file_bytes: int = DEFAULT_MAX_FILE_BYTES,
    max_total_bytes: int = DEFAULT_MAX_TOTAL_BYTES,
) -> None:
    """仅根据上传对象元数据预检，调用方可在读取文件内容前执行。"""
    if not file_sizes:
        raise UploadValidationError("请至少上传一个 Excel 文件。")
    if len(file_sizes) > max_files:
        raise UploadValidationError(f"单次最多上传 {max_files} 个文件。")
    if any(size > max_file_bytes for size in file_sizes):
        limit_mib = max_file_bytes // (1024 * 1024)
        raise UploadValidationError(f"单文件不能超过 {limit_mib} MiB。")
    if sum(file_sizes) > max_total_bytes:
        limit_mib = max_total_bytes // (1024 * 1024)
        raise UploadValidationError(f"上传文件总大小不能超过 {limit_mib} MiB。")


def _validate_excel_archive(name: str, content: bytes) -> int:
    """验证 Office ZIP 并返回累计解压字节数。"""
    try:
        with ZipFile(BytesIO(content)) as archive:
            entries = archive.infolist()
            entry_names = {entry.filename for entry in entries}
            if "[Content_Types].xml" not in entry_names or "xl/workbook.xml" not in entry_names:
                raise UploadValidationError(f"{name}：不是有效的 Excel 工作簿。")
            if len(entries) > DEFAULT_MAX_ARCHIVE_ENTRIES:
                raise UploadValidationError(f"{name}：工作簿内部文件数量异常。")
            uncompressed_size = sum(entry.file_size for entry in entries)
            if uncompressed_size > DEFAULT_MAX_UNCOMPRESSED_BYTES:
                raise UploadValidationError(f"{name}：工作簿解压后体积超过安全限制。")
            if any(entry.flag_bits & 0x1 for entry in entries):
                raise UploadValidationError(f"{name}：不支持加密的 Excel 工作簿。")
            return uncompressed_size
    except BadZipFile as exc:
        raise UploadValidationError(f"{name}：不是有效的 Excel 工作簿。") from exc


def validate_uploads(
    uploads: Sequence[UploadedFileData],
    *,
    max_files: int = DEFAULT_MAX_FILES,
    max_file_bytes: int = DEFAULT_MAX_FILE_BYTES,
    max_total_bytes: int = DEFAULT_MAX_TOTAL_BYTES,
    max_total_uncompressed_bytes: int = DEFAULT_MAX_TOTAL_UNCOMPRESSED_BYTES,
) -> List[UploadedFileData]:
    """验证扩展名、数量和大小，并返回文件名安全且唯一的不可变副本。"""
    if not uploads:
        raise UploadValidationError("请至少上传一个 Excel 文件。")
    if len(uploads) > max_files:
        raise UploadValidationError(f"单次最多上传 {max_files} 个文件。")

    validated: List[UploadedFileData] = []
    used_names: set[str] = set()
    total_bytes = 0
    total_uncompressed_bytes = 0
    for upload in uploads:
        name = safe_filename(upload.name)
        extension = Path(name).suffix.lower()
        if extension not in SUPPORTED_EXTENSIONS:
            raise UploadValidationError(f"{name}：仅支持 .xlsx 和 .xlsm 文件。")
        content = bytes(upload.content)
        if not content:
            raise UploadValidationError(f"{name}：文件内容为空。")
        if len(content) > max_file_bytes:
            limit_mib = max_file_bytes // (1024 * 1024)
            raise UploadValidationError(f"{name}：单文件不能超过 {limit_mib} MiB。")
        total_bytes += len(content)
        if total_bytes > max_total_bytes:
            limit_mib = max_total_bytes // (1024 * 1024)
            raise UploadValidationError(f"上传文件总大小不能超过 {limit_mib} MiB。")
        total_uncompressed_bytes += _validate_excel_archive(name, content)
        if total_uncompressed_bytes > max_total_uncompressed_bytes:
            limit_mib = max_total_uncompressed_bytes // (1024 * 1024)
            raise UploadValidationError(f"工作簿累计解压体积不能超过 {limit_mib} MiB。")
        validated.append(UploadedFileData(_deduplicate_name(name, used_names), content))
    return validated


def issue_groups_to_preview(groups: Sequence[IssueGroup]) -> List[Dict[str, Any]]:
    rows: List[Dict[str, Any]] = []
    for sequence, group in enumerate(groups, start=1):
        end_pile = "-" if group.row_count <= 1 else (group.end_pile or "-")
        rows.append(
            {
                "序号": sequence,
                "数据类型": group.category,
                "地市": group.city,
                "路线编号": group.route,
                "方向": group.direction,
                "起始桩号": group.start_pile,
                "结束桩号": end_pile,
                "异常类型": group.issue_type,
            }
        )
    return rows


def _zip_modified_directory(root: Path) -> bytes:
    archive_path = root / "modified-output.zip"
    modified_root = root / "已修改数据"
    with ZipFile(archive_path, "w", compression=ZIP_DEFLATED) as archive:
        for path in sorted(modified_root.rglob("*")):
            if path.is_file():
                archive.write(path, path.relative_to(root).as_posix())
    return archive_path.read_bytes()


def process_uploads(uploads: Sequence[UploadedFileData]) -> ProcessResult:
    """在请求专属临时目录内执行检查和修改，返回可直接下载的内存数据。"""
    validated = validate_uploads(uploads)
    logger = ListLogger()

    with TemporaryDirectory(prefix="guardrail_cloud_") as temporary_directory:
        workdir = Path(temporary_directory)
        input_dir = workdir / "input"
        output_dir = workdir / "output"
        input_dir.mkdir()
        output_dir.mkdir()

        input_paths: List[str] = []
        for upload in validated:
            path = input_dir / upload.name
            path.write_bytes(upload.content)
            input_paths.append(str(path))

        groups, stats = InspectionProcessor(logger=logger).process_files(input_paths)
        if stats.get("files_success", 0) == 0:
            raise UploadValidationError("所有文件处理失败，请确认工作簿未损坏且格式符合要求。")
        processed_sheet_count = stats.get("height_sheets_processed", 0) + stats.get(
            "bolt_sheets_processed", 0
        )
        if processed_sheet_count == 0:
            raise UploadValidationError("未识别到护栏高度或螺栓缺失业务工作表。")
        result_path = output_dir / f"护栏数据异常检查结果_{APP_VERSION}.xlsx"
        export_groups_to_excel(groups, str(result_path))

        _modified_folder, modification_stats = export_modified_workbooks(
            input_paths,
            str(output_dir),
            logger=logger,
        )
        modified_zip = _zip_modified_directory(output_dir)

        return ProcessResult(
            result_excel=result_path.read_bytes(),
            modified_zip=modified_zip,
            result_filename=result_path.name,
            modified_filename=f"已修改数据_{APP_VERSION}.zip",
            preview_rows=issue_groups_to_preview(groups),
            stats=stats,
            modification_stats=modification_stats,
            logs=list(logger.lines),
        )
