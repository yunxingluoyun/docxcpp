#!/usr/bin/env python3

from __future__ import annotations

import argparse
import io
import zipfile
from pathlib import Path


def build_archive(template_dir: Path) -> bytes:
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for path in sorted(template_dir.rglob("*")):
            if not path.is_file():
                continue
            if path.name == ".DS_Store":
                continue

            info = zipfile.ZipInfo(path.relative_to(template_dir).as_posix())
            info.compress_type = zipfile.ZIP_DEFLATED
            info.date_time = (2020, 1, 1, 0, 0, 0)
            archive.writestr(info, path.read_bytes())
    return buffer.getvalue()


def render_cpp(archive_bytes: bytes) -> str:
    lines = [
        '#include <cstdint>',
        '#include <vector>',
        '',
        '#include "docxcpp/opc_package.hpp"',
        '#include "internal/default_template_support.hpp"',
        '',
        'namespace docxcpp {',
        '',
        'OpcPackage load_default_template_package() {',
        '  static const std::uint8_t kDefaultTemplateArchive[] = {',
    ]

    for index in range(0, len(archive_bytes), 12):
        chunk = archive_bytes[index:index + 12]
        values = ", ".join(f"0x{byte:02x}" for byte in chunk)
        lines.append(f"    {values},")

    lines.extend(
        [
            '  };',
            '  return OpcPackage::from_archive_bytes(std::vector<std::uint8_t>(',
            '      kDefaultTemplateArchive,',
            '      kDefaultTemplateArchive + sizeof(kDefaultTemplateArchive)));',
            '}',
            '',
            '}  // namespace docxcpp',
            '',
        ]
    )
    return "\n".join(lines)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--template-dir", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()

    archive_bytes = build_archive(args.template_dir.resolve())
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(render_cpp(archive_bytes), encoding="utf-8")


if __name__ == "__main__":
    main()
