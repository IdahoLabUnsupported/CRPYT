# Copyright 2026, Battelle Energy Alliance, LLC. All Rights Reserved.
from __future__ import annotations

import argparse
import hashlib
import math
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile


REPO_ROOT = Path(__file__).resolve().parents[1]
DIST_DIR = REPO_ROOT / "dist"
OUT_PATH = DIST_DIR / "crpyt-release.zip"
CHECKSUM_PATH = DIST_DIR / "crpyt-release.zip.sha256"
ASSEMBLY_PATH = DIST_DIR / "crpyt-release-assemble.txt"
MAX_RELEASE_BYTES = 180 * 1024 * 1024
DEFAULT_CHUNK_SIZE_MIB = 64
COPY_BUFFER_BYTES = 4 * 1024 * 1024


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Build the CRPYT release archive.")
    parser.add_argument(
        "--chunk-index",
        type=int,
        help="Write only this 1-based archive chunk for CI artifact upload.",
    )
    parser.add_argument(
        "--chunk-size-mib",
        type=int,
        default=DEFAULT_CHUNK_SIZE_MIB,
        help=f"Chunk size in MiB when --chunk-index is used (default: {DEFAULT_CHUNK_SIZE_MIB}).",
    )
    parser.add_argument(
        "--min-chunk-count",
        type=int,
        default=0,
        help="Advertise at least this many chunks, writing empty trailing chunks if needed.",
    )
    args = parser.parse_args()
    if args.chunk_index is not None and args.chunk_index <= 0:
        parser.error("--chunk-index must be a positive integer.")
    if args.chunk_size_mib <= 0:
        parser.error("--chunk-size-mib must be a positive integer.")
    if args.min_chunk_count < 0:
        parser.error("--min-chunk-count cannot be negative.")
    return args


def iter_release_files() -> list[Path]:
    # The PyInstaller one-file executable already embeds tools/data and the
    # required legacy cache; duplicating those files makes CI artifact uploads
    # much larger without changing the supported binary release.
    return [
        REPO_ROOT / "README.md",
        REPO_ROOT / "CHANGELOG.md",
        REPO_ROOT / "LICENSE",
        REPO_ROOT / "bin" / "CRPYT.exe",
    ]


def validate_release_files(release_files: list[Path]) -> None:
    missing = [path.relative_to(REPO_ROOT) for path in release_files if not path.is_file()]
    if missing:
        joined = "\n".join(f"- {path}" for path in missing)
        raise RuntimeError(f"Release package is missing required files:\n{joined}")


def clean_release_outputs() -> None:
    for pattern in (
        OUT_PATH.name,
        CHECKSUM_PATH.name,
        f"{OUT_PATH.name}.part-*",
        ASSEMBLY_PATH.name,
    ):
        for path in DIST_DIR.glob(pattern):
            if path.is_file():
                path.unlink()


def write_release_archive(release_files: list[Path]) -> None:
    with ZipFile(OUT_PATH, "w", ZIP_DEFLATED) as archive:
        for path in release_files:
            archive.write(path, path.relative_to(REPO_ROOT))


def file_sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as source:
        for chunk in iter(lambda: source.read(COPY_BUFFER_BYTES), b""):
            digest.update(chunk)
    return digest.hexdigest()


def write_checksum(path: Path) -> str:
    archive_sha = file_sha256(path)
    CHECKSUM_PATH.write_text(f"{archive_sha}  {path.name}\n", encoding="utf-8")
    return archive_sha


def write_release_chunk(
    archive_path: Path,
    chunk_index: int,
    chunk_size_bytes: int,
    min_chunk_count: int,
) -> tuple[Path, int, int]:
    release_size = archive_path.stat().st_size
    actual_chunk_count = max(1, math.ceil(release_size / chunk_size_bytes))
    advertised_chunk_count = max(actual_chunk_count, min_chunk_count)
    if chunk_index > advertised_chunk_count:
        raise RuntimeError(
            f"Requested chunk {chunk_index}, but the release only needs "
            f"{advertised_chunk_count} chunk(s)."
        )

    chunk_path = DIST_DIR / f"{archive_path.name}.part-{chunk_index:02d}"
    start = (chunk_index - 1) * chunk_size_bytes
    remaining = max(0, min(chunk_size_bytes, release_size - start))

    with archive_path.open("rb") as source, chunk_path.open("wb") as target:
        if remaining > 0:
            source.seek(start)
        while remaining > 0:
            data = source.read(min(COPY_BUFFER_BYTES, remaining))
            if not data:
                break
            target.write(data)
            remaining -= len(data)

    return chunk_path, actual_chunk_count, advertised_chunk_count


def write_assembly_notes(
    archive_sha: str,
    release_size: int,
    actual_chunk_count: int,
    advertised_chunk_count: int,
    chunk_size_bytes: int,
) -> None:
    lines = [
        "CRPYT release archive assembly",
        "",
        f"Archive: {OUT_PATH.name}",
        f"SHA256: {archive_sha}",
        f"Archive size: {release_size / (1024 * 1024):.1f} MiB",
        f"Chunk size: {chunk_size_bytes / (1024 * 1024):.0f} MiB",
        f"Release chunks: {actual_chunk_count}",
        f"Uploaded chunk slots: {advertised_chunk_count}",
        "",
        "Download every crpyt-release.zip.part-* artifact into one folder.",
        "Trailing empty part files are safe to include.",
        "",
        "Linux, macOS, or Git Bash:",
        f"  cat {OUT_PATH.name}.part-* > {OUT_PATH.name}",
        f"  sha256sum -c {CHECKSUM_PATH.name}",
        "",
        "Windows PowerShell:",
        f"  $out = [System.IO.File]::Create(\"{OUT_PATH.name}\")",
        f"  Get-ChildItem \"{OUT_PATH.name}.part-*\" | Sort-Object Name | ForEach-Object {{",
        "    $in = [System.IO.File]::OpenRead($_.FullName)",
        "    $in.CopyTo($out)",
        "    $in.Dispose()",
        "  }",
        "  $out.Dispose()",
        f"  Get-FileHash \"{OUT_PATH.name}\" -Algorithm SHA256",
    ]
    ASSEMBLY_PATH.write_text("\n".join(lines) + "\n", encoding="utf-8")


def main() -> None:
    args = parse_args()
    DIST_DIR.mkdir(exist_ok=True)
    clean_release_outputs()
    release_files = iter_release_files()
    validate_release_files(release_files)
    write_release_archive(release_files)

    release_size = OUT_PATH.stat().st_size
    if release_size > MAX_RELEASE_BYTES:
        raise RuntimeError(
            f"Release archive is {release_size / (1024 * 1024):.1f} MiB, "
            f"above the {MAX_RELEASE_BYTES / (1024 * 1024):.0f} MiB CI artifact budget."
        )

    archive_sha = write_checksum(OUT_PATH)
    print(f"Wrote {OUT_PATH}")
    print(f"Packaged {len(release_files)} files.")
    print(f"Release archive size: {release_size / (1024 * 1024):.1f} MiB.")
    print(f"Release archive SHA256: {archive_sha}")

    if args.chunk_index is not None:
        chunk_size_bytes = args.chunk_size_mib * 1024 * 1024
        chunk_path, actual_chunk_count, advertised_chunk_count = write_release_chunk(
            OUT_PATH,
            args.chunk_index,
            chunk_size_bytes,
            args.min_chunk_count,
        )
        write_assembly_notes(
            archive_sha,
            release_size,
            actual_chunk_count,
            advertised_chunk_count,
            chunk_size_bytes,
        )
        print(
            f"Wrote release chunk {args.chunk_index} of {advertised_chunk_count}: "
            f"{chunk_path} ({chunk_path.stat().st_size / (1024 * 1024):.1f} MiB)."
        )


if __name__ == "__main__":
    main()
