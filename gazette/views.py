import os
from pathlib import Path
import tempfile
from zipfile import ZipFile

from django.http import FileResponse
from django.shortcuts import render

from .utils import extract_to_excel, extract_pdf_to_excel

def save_uploaded_to_temp(uploaded_file, suffix: str) -> Path:
    """把上傳檔案存成暫存檔，回傳路徑。"""
    with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as tmp:
        for chunk in uploaded_file.chunks():
            tmp.write(chunk)
        return Path(tmp.name)

def process_single_file(uploaded):
    """
    處理單一上傳檔案，回傳 (output_path, download_name)。
    可能拋出 ValueError（副檔名不支援）或其他例外。
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

def upload_xml(request):
    """
    顯示上傳表單，接收 XML / ZIP / PDF（單檔或多檔），
    單檔：直接回傳 Excel
    多檔：打包成一個 zip 回傳。
    """
    if request.method == "POST":
        # ✅ 同時收「檔案 input」和「資料夾 input」兩邊的檔案
        files = []
        files.extend(request.FILES.getlist("files"))         # 來自檔案選擇
        files.extend(request.FILES.getlist("folder_files"))  # 來自資料夾選擇

        if not files:
            return render(request, "gazette/upload_xml.html", {"error": "請選擇至少一個檔案。"})

        try:
            # 只有一個檔案，就走原本邏輯
            if len(files) == 1:
                output_path, download_name = process_single_file(files[0])
                response = FileResponse(
                    open(output_path, "rb"),
                    as_attachment=True,
                    filename=download_name,
                )
                return response

            # 多檔案：全部處理後打包成 zip
            with tempfile.NamedTemporaryFile(suffix=".zip", delete=False) as tmp_zip:
                zip_path = Path(tmp_zip.name)

            errors = []
            with ZipFile(zip_path, "w") as zf:
                seen_names = set()  # 用來追蹤 zip 內已經使用過的檔名

                for f in files:
                    try:
                        out_path, out_name = process_single_file(f)

                        # 如果 out_name 已經存在，就自動加 _2, _3 ... 避免重複
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

            # 如果全部都失敗，回畫面顯示錯誤
            if errors and len(errors) == len(files):
                return render(
                    request,
                    "gazette/upload_xml.html",
                    {"error": "所有檔案皆處理失敗：\n" + "\n".join(errors)},
                )

            # 至少有一個成功，就回傳打包好的 zip
            response = FileResponse(
                open(zip_path, "rb"),
                as_attachment=True,
                filename="gazette_results.zip",
            )
            return response

        except Exception as exc:  # pylint: disable=broad-exception-caught
            return render(
                request,
                "gazette/upload_xml.html",
                {"error": f"處理檔案時發生錯誤：{exc}"},
            )

    # GET 或初次進入
    return render(request, "gazette/upload_xml.html")
