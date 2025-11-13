# gazette/views.py
from typing import List

from django.http import FileResponse, HttpRequest
from django.shortcuts import render
from django.core.files.uploadedfile import UploadedFile

from .pipeline import process_uploaded_files


def upload_xml(request: HttpRequest):
    """
    顯示上傳表單，接收 XML / ZIP / PDF（單檔或多檔），
    單檔：直接回傳 Excel
    多檔：打包成一個 zip 回傳。
    """
    if request.method == "POST":
        # ✅ 同時收「檔案 input」和「資料夾 input」兩邊的檔案（維持你的作法）
        files: List[UploadedFile] = []
        files.extend(request.FILES.getlist("files"))         # 來自檔案選擇
        files.extend(request.FILES.getlist("folder_files"))  # 來自資料夾選擇

        if not files:
            return render(
                request,
                "gazette/upload_xml.html",
                {"error": "請選擇至少一個檔案。"},
            )

        output_path, download_name, errors = process_uploaded_files(files)

        # 全失敗：回頁面顯示錯誤
        if output_path is None or download_name is None:
            return render(
                request,
                "gazette/upload_xml.html",
                {
                    "error": "所有檔案皆處理失敗：\n" + "\n".join(errors),
                },
            )

        # 至少有一個成功：直接回傳檔案（可能是 Excel 或 ZIP）
        response = FileResponse(
            open(output_path, "rb"),
            as_attachment=True,
            filename=download_name,
        )
        return response

    # GET 或初次進入
    return render(request, "gazette/upload_xml.html")
