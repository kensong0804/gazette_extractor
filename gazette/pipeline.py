# gazette/pipeline.py
import os
from pathlib import Path
import tempfile
from zipfile import ZipFile
from typing import List, Tuple, Optional

from django.core.files.uploadedfile import UploadedFile

from .utils import extract_to_excel, extract_pdf_to_excel


def save_uploaded_to_temp(uploaded_file: UploadedFile, suffix: str) -> Path:
    """把上傳檔案存成暫存檔，回傳路徑。"""
    with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as tmp:
        for chunk in uploaded_file.chunks():
            tmp.write(chunk)
        return Path(tmp.name)


def process_single_file(uploaded: UploadedFile) -> Tuple[Path, str]:
    """
    處理單一上傳檔案，回傳 (output_path, download_name)。
    只支援 XML / ZIP(內含 XML) / PDF。
    """
    filename = uploaded.name
    lower_name = filename.lower()

    # --- XML ---
    if lower_name.endswith(".xml"):
        tmp_xml_path = save_uploaded_to_temp(uploaded, ".xml")
        output_path = tmp_xml_path.with_suffix(".xlsx")
        extract_to_excel(tmp_xml_path, output_path)
        download_name = f"{Path(filename).stem}.xlsx"
        return output_path, download_name

    # --- ZIP（內含 XML） ---
    if lower_name.endswith(".zip"):
        tmp_zip_path = save_uploaded_to_temp(uploaded, ".zip")

        with ZipFile(tmp_zip_path, "r") as zf:
            xml_names = [
                name for name in zf.namelist()
                if name.lower().endswith(".xml")
            ]

            if not xml_names:
                raise ValueError("壓縮檔內找不到任何 XML 檔案。")

            # 目前你的實作是只取第一個 XML，維持不變
            xml_in_zip = xml_names[0]

            with zf.open(xml_in_zip) as xml_file, \
                    tempfile.NamedTemporaryFile(suffix=".xml", delete=False) as tmp_xml:
                tmp_xml.write(xml_file.read())
                tmp_xml_path = Path(tmp_xml.name)

        output_path = tmp_xml_path.with_suffix(".xlsx")
        extract_to_excel(tmp_xml_path, output_path)
        download_name = f"{Path(xml_in_zip).stem}.xlsx"
        return output_path, download_name

    # --- PDF ---
    if lower_name.endswith(".pdf"):
        tmp_pdf_path = save_uploaded_to_temp(uploaded, ".pdf")
        output_path = tmp_pdf_path.with_suffix(".xlsx")
        extract_pdf_to_excel(tmp_pdf_path, output_path, original_name=filename)
        download_name = f"{Path(filename).stem}_pdf.xlsx"
        return output_path, download_name

    raise ValueError("只支援 .xml、.zip、.pdf 檔案。")


def process_uploaded_files(
    files: List[UploadedFile],
) -> Tuple[Optional[Path], Optional[str], List[str]]:
    """
    ⭐「完整資料管線」入口⭐

    - 1 個檔案：直接呼叫 process_single_file，回傳那個 Excel。
    - 多個檔案：每個跑 process_single_file，成功的統統塞進一個 ZIP。
    - errors：紀錄哪些檔案失敗（但只要有成功就會回 ZIP）。
    - 若全部失敗：output_path 會是 None。
    """
    if not files:
        return None, None, ["沒有收到任何檔案。"]

    # 單檔：直接走原本邏輯
    if len(files) == 1:
        out_path, out_name = process_single_file(files[0])
        return out_path, out_name, []

    # 多檔：建立一個暫存 ZIP
    with tempfile.NamedTemporaryFile(suffix=".zip", delete=False) as tmp_zip:
        zip_path = Path(tmp_zip.name)

    errors: List[str] = []
    seen_names = set()

    with ZipFile(zip_path, "w") as zf:
        for f in files:
            try:
                out_path, out_name = process_single_file(f)

                # 處理重複檔名：自動加 _2, _3 ...
                base, ext = os.path.splitext(out_name)
                final_name = out_name
                idx = 2
                while final_name in seen_names:
                    final_name = f"{base}_{idx}{ext}"
                    idx += 1

                zf.write(out_path, arcname=final_name)
                seen_names.add(final_name)

            except ValueError as ve:
                errors.append(f"{f.name}：{ve}")
            except Exception as exc:  # pylint: disable=broad-exception-caught
                errors.append(f"{f.name}：處理失敗（{exc}）")

    # 全部失敗：不回 ZIP
    if errors and len(errors) == len(files):
        return None, None, errors

    # 至少一個成功：回 ZIP
    return zip_path, "gazette_results.zip", errors
