import os
import json
import datetime
import re
import threading
import queue
import tempfile
import io
import random
import locale
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, colorchooser
from dataclasses import dataclass, asdict
from typing import List, Dict, Any, Optional, Tuple

import pythoncom
import win32com.client
from office_backend import application, close_document
from pypdf import PdfWriter, PdfReader
from PIL import Image

from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase.cidfonts import UnicodeCIDFont

from tkinterdnd2 import DND_FILES, TkinterDnD

try:
    import winreg
except ImportError:
    winreg = None

CONFIG_FILE = "pdf_pro_config_v4.json"
WM_TEMPLATE_FILE = "watermark_templates.txt"
NM_TEMPLATE_FILE = "naming_templates.txt"


# 言語データの定義
# text=self._("******")へ置き換え
# self.root.title(self._("win_title"))とか
# 言語データの定義
I18N = {
    "ja": {
        "win_title": "Office2PDF v5.0",
        "file_list": "変換ファイル(ドロップで登録・ダブルクリックでページ指定)",
        "col_type": "種別",
        "col_name": "ファイル名",
        "col_range": "範囲",
        "col_out": "出力先",
        "btn_up": "上へ",
        "btn_down": "下へ",
        "btn_remove": "削除",
        "btn_clear": "全消去",
        "btn_clear_list": "リスト全クリア",
        "btn_add_folder": "フォルダ追加",
        "frame_wm": "透かし・ページ番号設定",
        "lbl_font": "フォント:",
        "lbl_size": "サイズ:",
        "lbl_alpha": "不透明度:",
        "btn_start": "PDF変換開始",
        "btn_cancel": "キャンセル",
        "msg_no_file": "ファイルがありません。",
        "err_save_config": "設定保存エラー:",
        "lbl_presets": "設定プリセット:",
        "btn_load": "読込",
        "btn_save": "保存",
        "btn_delete": "削除",
        "btn_preview_wm": "選択ファイルの1頁目をプレビュー",
        "pos_page_center": "中央下",
        "lbl_watermark": "透かし",
        "wm_label": "透かし",
        "pos_none": "なし",
        "pos_diag_center": "中央斜め",
        "pos_large_center": "中央大",
        "pos_top_left": "左上",
        "pos_top_center": "上中央",
        "pos_top_right": "右上",
        "pos_bottom_left": "左下",
        "pos_bottom_center": "下中央",
        "pos_bottom_right": "右下",
        "pos_page_center": "中央下(ページ番号用)",
        "lbl_page_num": "ページ番号:",
        "frame_detail": "出力・分割詳細設定",
        "chk_merge": "【全結合】1つのPDFにまとめる",
        "chk_split_page": "ページ毎分割",
        "chk_include_ppt": "(PPTも)",
        "chk_by_sheet": "シート毎",
        "chk_by_all_pages": "全ページ毎",
        "lbl_excel_opt": "--- Excelオプション (印刷範囲優先) ---",
        "chk_fit_width": "横幅1ページに収める",
        "chk_fit_height": "縦幅1ページに収める",
        "frame_exec": "保存設定・実行",
        "lbl_naming": "ファイル名・命名ルール:",
        "btn_tag_help": "タグ説明",
        "opt_same_dir": "元と同じ場所",
        "opt_custom_dir": "カスタム:",
        "btn_browse": "参照",
        "lbl_password": "パスワード:",
        "chk_meta_clear": "メタ削除",
        "chk_compress": "PDF軽量化",
        "chk_open_done": "完了後開く",
        "chk_open_folder": "フォルダ開く",
        "chk_clear_after": "リストクリア",
        "st_ready": "待機中...",
        "log_font_loaded": "フォント一覧の読み込みが完了しました。",
        "log_font_err": "フォント取得エラー:",
        "log_font_using": "使用フォント:",
        "log_font_fail": "フォント登録失敗:",
        "title_warn": "警告",
        "msg_no_files": "ファイルがありません",
        "st_converting": "変換中...",
        "st_conv_file": "変換中:",
        "st_finalizing": "最終処理中...",
        "log_fatal": "致命的エラー:",
        "val_all_pages": "全ページ",
        "log_ppt_err": "PPT変換エラー",
        "title_info": "情報",
        "msg_no_preview": "プレビューするファイルがありません。",
        "st_preview_gen": "プレビュー生成中:",
        "log_conv_fail": "変換に失敗しました:",
        "msg_no_output": "処理対象のファイルが生成されなかったため、終了します。",
        "msg_preview_ok": "プレビューを表示しました。",
        "msg_preview_fail": "プレビュー失敗:",
        "title_tag_help": "命名ルールのタグ説明",
        "title_overwrite": "上書き確認",
        "msg_exists": "存在します:",
        "btn_overwrite": "上書き",
        "btn_seq": "連番",
        "btn_abort": "中止",
        "msg_all_done": "すべての処理が完了しました。",
        "log_preset_load": "プリセットを読込。",
        "lbl_preset_name": "プリセット名:",
        "title_confirm": "確認",
        "msg_ask_delete": "削除しますか？",
        "title_range": "範囲編集",
        "pg_pos_bc": "中央下",
        "pg_pos_br": "右下",
        "msg_lang_restart": "言語を変更しました。アプリを再起動して適用しますか？",
        "lbl_output_example": "出力例",
        "help_tags": (
            "【利用可能なタグ】\n\n"
            "{name} : 元のファイル名\n"
            "{sheet} : Excelシート名\n"
            "{parent} : 親フォルダの名前\n"
            "{seq} : 全体の通し番号\n"
            "{fseq} : ファイル毎の番号\n"
            "{pseq} : ページ毎の番号\n"
            "{total} : 全ファイル数\n"
            "{ptotal} : ファイル内の総ページ数\n"
            "{username} : PCユーザー名\n"
            "{rand} : 4桁のランダム数字\n\n"
            "【日付・時刻】\n"
            "{date:yyyy-mm-dd} -> 2024-02-06\n"
            "※HH:時, MM:分, SS:秒"
        ),
    },
    "en": {
        "win_title": "Office2PDF v5.0",
        "file_list": "Files (Drag & Drop to add / Double-click to set range)",
        "col_type": "Type",
        "col_name": "File Name",
        "col_range": "Range",
        "col_out": "Output Preview",
        "btn_up": "Up",
        "btn_down": "Down",
        "btn_remove": "Remove",
        "btn_clear": "Clear List",
        "btn_clear_list": "Clear List",
        "btn_add_folder": "Add Folder",
        "frame_wm": "Watermark & Page Number",
        "lbl_font": "Font:",
        "lbl_size": "Size:",
        "lbl_alpha": "Opacity:",
        "btn_start": "Start Conversion",
        "btn_cancel": "Cancel",
        "msg_no_file": "No files selected.",
        "err_save_config": "Error saving settings:",
        "lbl_presets": "Presets:",
        "btn_load": "Load",
        "btn_save": "Save",
        "btn_delete": "Delete",
        "btn_preview_wm": "Preview 1st page of selected file",
        "pos_page_center": "Bottom Center (Mid)",
        "lbl_watermark": "Watermark",
        "wm_label": "Watermark",
        "pos_none": "None",
        "pos_diag_center": "Diagonal Center",
        "pos_large_center": "Large Center",
        "pos_top_left": "Top Left",
        "pos_top_center": "Top Center",
        "pos_top_right": "Top Right",
        "pos_bottom_left": "Bottom Left",
        "pos_bottom_center": "Bottom Center",
        "pos_bottom_right": "Bottom Right",
        "pos_page_center": "Bottom Center (Page)",
        "lbl_page_num": "Page Numbers:",
        "frame_detail": "Output & Split Settings",
        "chk_merge": "[Merge] Combine into a single PDF",
        "chk_split_page": "Split by Page",
        "chk_include_ppt": "(Include PPT)",
        "chk_by_sheet": "By Sheet",
        "chk_by_all_pages": "By All Pages",
        "lbl_excel_opt": "--- Excel Options (Prioritize Print Area) ---",
        "chk_fit_width": "Fit width to 1 page",
        "chk_fit_height": "Fit height to 1 page",
        "frame_exec": "Export Settings & Run",
        "lbl_naming": "Naming Rules:",
        "btn_tag_help": "Tag Guide",
        "opt_same_dir": "Same as source",
        "opt_custom_dir": "Custom:",
        "btn_browse": "Browse...",
        "lbl_password": "Password:",
        "chk_meta_clear": "Strip Metadata",
        "chk_compress": "Compress PDF",
        "chk_open_done": "Open when done",
        "chk_open_folder": "Open folder",
        "chk_clear_after": "Clear list",
        "st_ready": "Ready...",
        "log_font_loaded": "Font list loaded successfully.",
        "log_font_err": "Error fetching fonts:",
        "log_font_using": "Font used:",
        "log_font_fail": "Failed to register font:",
        "title_warn": "Warning",
        "msg_no_files": "No files selected",
        "st_converting": "Converting...",
        "st_conv_file": "Converting:",
        "st_finalizing": "Finalizing...",
        "log_fatal": "Critical Error:",
        "val_all_pages": "All Pages",
        "log_ppt_err": "PPT Conversion Error",
        "title_info": "Info",
        "msg_no_preview": "No file to preview.",
        "st_preview_gen": "Generating preview:",
        "log_conv_fail": "Conversion failed:",
        "msg_no_output": "No output files were generated. Process aborted.",
        "msg_preview_ok": "Preview displayed successfully.",
        "msg_preview_fail": "Preview failed:",
        "title_tag_help": "Naming Rule Tag Guide",
        "title_overwrite": "Confirm Overwrite",
        "msg_exists": "File already exists:",
        "btn_overwrite": "Overwrite",
        "btn_seq": "Add Seq Number",
        "btn_abort": "Abort",
        "msg_all_done": "All processes completed successfully.",
        "log_preset_load": "Preset loaded.",
        "lbl_preset_name": "Preset Name:",
        "title_confirm": "Confirm",
        "msg_ask_delete": "Are you sure you want to delete?",
        "title_range": "Edit Range",
        "pg_pos_bc": "Bottom Center",
        "pg_pos_br": "Bottom Right",
        "msg_lang_restart": "Language changed. Restart the app to apply?",
        "lbl_output_example": "Output example",
        "help_tags": (
            "[Available Tags]\n\n"
            "{name} : Original filename\n"
            "{sheet} : Excel sheet name\n"
            "{parent} : Parent folder name\n"
            "{seq} : Global sequence number\n"
            "{fseq} : File sequence number\n"
            "{pseq} : Page sequence number\n"
            "{total} : Total file count\n"
            "{ptotal} : Total pages in file\n"
            "{username} : PC username\n"
            "{rand} : 4-digit random number\n\n"
            "[Date & Time]\n"
            "{date:yyyy-mm-dd} -> 2024-02-06\n"
            "* HH:Hour, MM:Min, SS:Sec"
        ),
    },
    "zh_cn": {
        "win_title": "Office2PDF v5.0",
        "file_list": "转换文件（拖拽添加 / 双击设置页码范围）",
        "col_type": "类型",
        "col_name": "文件名",
        "col_range": "范围",
        "col_out": "输出预览",
        "btn_up": "上移",
        "btn_down": "下移",
        "btn_remove": "移除",
        "btn_clear": "全部清除",
        "btn_clear_list": "清空列表",
        "btn_add_folder": "添加文件夹",
        "frame_wm": "水印与页码设置",
        "lbl_font": "字体:",
        "lbl_size": "大小:",
        "lbl_alpha": "不透明度:",
        "btn_start": "开始转换 PDF",
        "btn_cancel": "取消",
        "msg_no_file": "没有选择文件。",
        "err_save_config": "设置保存错误:",
        "lbl_presets": "设置预设:",
        "btn_load": "读取",
        "btn_save": "保存",
        "btn_delete": "删除",
        "btn_preview_wm": "预览所选文件第 1 页",
        "pos_page_center": "下方居中",
        "lbl_watermark": "水印",
        "wm_label": "水印",
        "pos_none": "无",
        "pos_diag_center": "居中斜排",
        "pos_large_center": "居中大字",
        "pos_top_left": "左上",
        "pos_top_center": "顶部居中",
        "pos_top_right": "右上",
        "pos_bottom_left": "左下",
        "pos_bottom_center": "底部居中",
        "pos_bottom_right": "右下",
        "pos_page_center": "下方居中（页码用）",
        "lbl_page_num": "页码:",
        "frame_detail": "输出与拆分详细设置",
        "chk_merge": "【全部合并】合并为一个 PDF",
        "chk_split_page": "按页拆分",
        "chk_include_ppt": "（含 PPT）",
        "chk_by_sheet": "按工作表",
        "chk_by_all_pages": "按全部页面",
        "lbl_excel_opt": "--- Excel 选项（优先打印区域） ---",
        "chk_fit_width": "宽度适配一页",
        "chk_fit_height": "高度适配一页",
        "frame_exec": "保存设置与执行",
        "lbl_naming": "文件名命名规则:",
        "btn_tag_help": "标签说明",
        "opt_same_dir": "与原文件相同",
        "opt_custom_dir": "自定义:",
        "btn_browse": "浏览...",
        "lbl_password": "密码:",
        "chk_meta_clear": "删除元数据",
        "chk_compress": "PDF 瘦身",
        "chk_open_done": "完成后打开",
        "chk_open_folder": "打开文件夹",
        "chk_clear_after": "清空列表",
        "st_ready": "待机中...",
        "log_font_loaded": "字体列表加载完成。",
        "log_font_err": "获取字体出错:",
        "log_font_using": "使用字体:",
        "log_font_fail": "字体注册失败:",
        "title_warn": "警告",
        "msg_no_files": "没有选择文件",
        "st_converting": "转换中...",
        "st_conv_file": "转换中:",
        "st_finalizing": "正在完成最后处理...",
        "log_fatal": "致命错误:",
        "val_all_pages": "全部页面",
        "log_ppt_err": "PPT 转换错误",
        "title_info": "信息",
        "msg_no_preview": "没有可预览的文件。",
        "st_preview_gen": "正在生成预览:",
        "log_conv_fail": "转换失败:",
        "msg_no_output": "未生成任何输出文件，已终止。",
        "msg_preview_ok": "预览已显示。",
        "msg_preview_fail": "预览失败:",
        "title_tag_help": "命名规则标签说明",
        "title_overwrite": "覆盖确认",
        "msg_exists": "已存在:",
        "btn_overwrite": "覆盖",
        "btn_seq": "加序号",
        "btn_abort": "中止",
        "msg_all_done": "全部处理完成。",
        "log_preset_load": "预设已读取。",
        "lbl_preset_name": "预设名称:",
        "title_confirm": "确认",
        "msg_ask_delete": "确定要删除吗？",
        "title_range": "编辑范围",
        "pg_pos_bc": "下方居中",
        "pg_pos_br": "右下",
        "msg_lang_restart": "语言已切换。重启应用以应用新语言吗？",
        "lbl_output_example": "输出示例",
        "help_tags": (
            "【可用标签】\n\n"
            "{name} : 原文件名\n"
            "{sheet} : Excel 工作表名\n"
            "{parent} : 上级文件夹名\n"
            "{seq} : 全局序号\n"
            "{fseq} : 文件内序号\n"
            "{pseq} : 页面序号\n"
            "{total} : 文件总数\n"
            "{ptotal} : 文件内总页数\n"
            "{username} : 电脑用户名\n"
            "{rand} : 4 位随机数字\n\n"
            "【日期与时间】\n"
            "{date:yyyy-mm-dd} -> 2024-02-06\n"
            "※HH:时, MM:分, SS:秒"
        ),
    },
    "zh_tw": {
        "win_title": "Office2PDF v5.0",
        "file_list": "轉換檔案（拖曳加入 / 雙擊設定頁碼範圍）",
        "col_type": "類型",
        "col_name": "檔案名稱",
        "col_range": "範圍",
        "col_out": "輸出位置",
        "btn_up": "上移",
        "btn_down": "下移",
        "btn_remove": "移除",
        "btn_clear": "全部清除",
        "btn_clear_list": "清空清單",
        "btn_add_folder": "加入資料夾",
        "frame_wm": "浮水印與頁碼設定",
        "lbl_font": "字型:",
        "lbl_size": "大小:",
        "lbl_alpha": "不透明度:",
        "btn_start": "開始轉換 PDF",
        "btn_cancel": "取消",
        "msg_no_file": "尚未選擇檔案。",
        "err_save_config": "設定儲存錯誤:",
        "lbl_presets": "設定預設:",
        "btn_load": "讀取",
        "btn_save": "儲存",
        "btn_delete": "刪除",
        "btn_preview_wm": "預覽所選檔案第 1 頁",
        "pos_page_center": "下方置中",
        "lbl_watermark": "浮水印",
        "wm_label": "浮水印",
        "pos_none": "無",
        "pos_diag_center": "置中斜排",
        "pos_large_center": "置中大字",
        "pos_top_left": "左上",
        "pos_top_center": "上方置中",
        "pos_top_right": "右上",
        "pos_bottom_left": "左下",
        "pos_bottom_center": "下方置中",
        "pos_bottom_right": "右下",
        "pos_page_center": "下方置中（頁碼用）",
        "lbl_page_num": "頁碼:",
        "frame_detail": "輸出與分割詳細設定",
        "chk_merge": "【全部合併】合併為單一 PDF",
        "chk_split_page": "按頁分割",
        "chk_include_ppt": "（含 PPT）",
        "chk_by_sheet": "按工作表",
        "chk_by_all_pages": "按全部頁面",
        "lbl_excel_opt": "--- Excel 選項（優先列印範圍） ---",
        "chk_fit_width": "寬度適應單頁",
        "chk_fit_height": "高度適應單頁",
        "frame_exec": "儲存設定與執行",
        "lbl_naming": "檔案命名規則:",
        "btn_tag_help": "標籤說明",
        "opt_same_dir": "與原始檔相同",
        "opt_custom_dir": "自訂:",
        "btn_browse": "瀏覽...",
        "lbl_password": "密碼:",
        "chk_meta_clear": "移除中繼資料",
        "chk_compress": "PDF 瘦身",
        "chk_open_done": "完成後開啟",
        "chk_open_folder": "開啟資料夾",
        "chk_clear_after": "清空清單",
        "st_ready": "待機中...",
        "log_font_loaded": "字型清單載入完成。",
        "log_font_err": "取得字型錯誤:",
        "log_font_using": "使用字型:",
        "log_font_fail": "字型註冊失敗:",
        "title_warn": "警告",
        "msg_no_files": "尚未選擇檔案",
        "st_converting": "轉換中...",
        "st_conv_file": "轉換中:",
        "st_finalizing": "正在完成最後處理...",
        "log_fatal": "嚴重錯誤:",
        "val_all_pages": "全部頁面",
        "log_ppt_err": "PPT 轉換錯誤",
        "title_info": "資訊",
        "msg_no_preview": "沒有可預覽的檔案。",
        "st_preview_gen": "正在產生預覽:",
        "log_conv_fail": "轉換失敗:",
        "msg_no_output": "未產生任何輸出檔案，已中止。",
        "msg_preview_ok": "預覽已顯示。",
        "msg_preview_fail": "預覽失敗:",
        "title_tag_help": "命名規則標籤說明",
        "title_overwrite": "覆蓋確認",
        "msg_exists": "已存在:",
        "btn_overwrite": "覆蓋",
        "btn_seq": "加編號",
        "btn_abort": "中止",
        "msg_all_done": "全部處理完成。",
        "log_preset_load": "預設已讀取。",
        "lbl_preset_name": "預設名稱:",
        "title_confirm": "確認",
        "msg_ask_delete": "確定要刪除嗎？",
        "title_range": "編輯範圍",
        "pg_pos_bc": "下方置中",
        "pg_pos_br": "右下",
        "msg_lang_restart": "語言已切換。重新啟動應用程式以套用新語言嗎？",
        "lbl_output_example": "輸出範例",
        "help_tags": (
            "【可用標籤】\n\n"
            "{name} : 原始檔案名稱\n"
            "{sheet} : Excel 工作表名稱\n"
            "{parent} : 上層資料夾名稱\n"
            "{seq} : 全域流水號\n"
            "{fseq} : 檔案內流水號\n"
            "{pseq} : 頁面流水號\n"
            "{total} : 檔案總數\n"
            "{ptotal} : 檔案內總頁數\n"
            "{username} : 電腦使用者名稱\n"
            "{rand} : 4 位隨機數字\n\n"
            "【日期與時間】\n"
            "{date:yyyy-mm-dd} -> 2024-02-06\n"
            "※HH:時, MM:分, SS:秒"
        ),
    },
}
# 位置の内部IDと翻訳キーの対応表
POS_MAP = [
    ("None", "pos_none"),
    ("diag", "pos_diag_center"),
    ("large", "pos_large_center"),
    ("tl", "pos_top_left"),
    ("tc", "pos_top_center"),
    ("tr", "pos_top_right"),
    ("bl", "pos_bottom_left"),
    ("bc", "pos_bottom_center"),
    ("br", "pos_bottom_right"),
]


@dataclass
class AppConfig:
    engine: str = "auto"
    lang: str = ""  # 手动选择的语言（空 = 自动检测）
    output_dir: str = ""
    out_mode: str = "original"
    naming_tpl: str = "{name}"
    auto_open: bool = True
    open_folder: bool = False
    clear_after: bool = False
    compress_pdf: bool = False  # 追加

    wm1_text: str = ""
    wm1_pos: str = "None"
    wm2_text: str = ""
    wm2_pos: str = "None"

    wm_font: str = ""
    wm_size: int = 60
    wm_color: str = "#C0C0C0"
    wm_alpha: float = 0.3

    pg_enabled: bool = False
    pg_pos: str = "bc"
    pg_format: str = "- {n} / {total} -"

    merge_all: bool = False
    split_word_page: bool = False
    split_ppt_page: bool = False
    split_pdf_page: bool = False
    split_excel_sheet: bool = False
    split_excel_page: bool = False

    password: str = ""
    excel_fit: bool = False
    excel_fit_tall: bool = False
    clear_metadata: bool = False

    def __post_init__(self):
        if not self.output_dir:
            self.output_dir = os.path.expanduser(r"~\Desktop")


class PDFUltimateApp:
    def __init__(self, root):
        self.root = root

        # OSの言語を取得（'ja_JP' なら 'ja'、'zh_CN' なら 'zh_cn'、それ以外なら 'en' にする）
        system_lang = locale.getdefaultlocale()[0] or ""
        if system_lang.startswith("ja"):
            self.lang = "ja"
        elif system_lang.startswith("zh"):
            # zh_TW / zh_HK / zh_MO → 繁体字、それ以外（zh_CN / zh_SG）→ 簡体字
            self.lang = "zh_tw" if system_lang in ("zh_TW", "zh_HK", "zh_MO") else "zh_cn"
        else:
            self.lang = "en"

        # 翻訳用ヘルパー関数
        self._ = lambda key: I18N[self.lang].get(key, key)

        # 以下、UI構築でこの関数を使う
        self.root.title(self._("win_title"))
        import ctypes
        from ctypes import wintypes
        work_area = wintypes.RECT(0, 0, self.root.winfo_screenwidth(), self.root.winfo_screenheight())
        ctypes.windll.user32.SystemParametersInfoW(0x0030, 0, ctypes.byref(work_area), 0)
        width = min(1000, work_area.right - work_area.left - 80)
        height = min(860, work_area.bottom - work_area.top - 80)
        self.root.geometry(f"{width}x{height}+{work_area.left + 30}+{work_area.top + 30}")
        self.root.minsize(min(560, width), min(420, height))

        self.files: List[Dict[str, Any]] = []
        self.presets: Dict[str, Any] = {}
        self.font_map: Dict[str, Tuple[str, int]] = {}
        self.processing = False
        self.cancel_flag = threading.Event()
        self.progress_queue = queue.Queue()

        self.load_config()
        # 手动选择的语言优先于系统自动检测
        saved_lang = getattr(self.config, "lang", "")
        if saved_lang in I18N:
            self.lang = saved_lang
        self.init_templates()
        self.setup_ui()
        self.apply_config_to_ui()

        self.root.after(100, self.check_progress_queue)
        self.root.after(600, self.load_fonts_delayed)

    # --- Config & Templates ---
    def load_config(self):
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                    data = json.load(f)
                self.config = AppConfig(**data.get("current", {}))
                self.presets = data.get("presets", {})
            except:
                self.config = AppConfig()
        else:
            self.config = AppConfig()

    def save_config(self):
        try:
            self.update_config_from_ui()
            data = {"current": asdict(self.config), "presets": self.presets}
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=4)
        except Exception as e:
            self.queue_log(f"{self._('err_save_config')} {e}")

    def init_templates(self):
        if self.lang == "ja":
            wm_content = "社外秘\nコピー厳禁\n取扱注意\nSAMPLE\nCONFIDENTIAL\nDRAFT\nCOPY\nIMPORTANT\n{date:yyyy/mm/dd}\n{date:yyyy/mm/dd hh:mm}"
        elif self.lang == "zh_cn":
            wm_content = "机密\n禁止复制\n注意处理\nSAMPLE\nCONFIDENTIAL\nDRAFT\nCOPY\nIMPORTANT\n{date:yyyy/mm/dd}\n{date:yyyy/mm/dd hh:mm}"
        elif self.lang == "zh_tw":
            wm_content = "機密\n禁止複製\n注意處理\nSAMPLE\nCONFIDENTIAL\nDRAFT\nCOPY\nIMPORTANT\n{date:yyyy/mm/dd}\n{date:yyyy/mm/dd hh:mm}"
        else:
            wm_content = "CONFIDENTIAL\nDO NOT COPY\nHANDLE WITH CARE\nSAMPLE\nDRAFT\nCOPY\nIMPORTANT\n{date:yyyy/mm/dd}\n{date:yyyy/mm/dd hh:mm}"
        nm_content = "{name}\n{name}_{sheet}\n{seq}_{name}\n{date:yyyy-mm-dd}_{name}"
        for file, content in [(WM_TEMPLATE_FILE, wm_content), (NM_TEMPLATE_FILE, nm_content)]:
            if not os.path.exists(file):
                with open(file, "w", encoding="utf-8") as f:
                    f.write(content)
        with open(WM_TEMPLATE_FILE, "r", encoding="utf-8") as f:
            self.wm_templates = [l.strip() for l in f if l.strip()]
        with open(NM_TEMPLATE_FILE, "r", encoding="utf-8") as f:
            self.nm_templates = [l.strip() for l in f if l.strip()]

    def update_config_from_ui(self):
        self.config.engine = ("auto", "office", "wps")[max(0, self.engine_combo.current())]
        self.config.wm1_text = self.wm1_val.get()
        self.config.wm1_pos = self.wm1_pos_var.get()
        self.config.wm2_text = self.wm2_val.get()
        self.config.wm2_pos = self.wm2_pos_var.get()
        # 透かし位置：表示名 → 内部ID
        d1 = self.wm1_pos_var.get()
        d2 = self.wm2_pos_var.get()
        self.config.wm1_pos = self.pos_display_to_id.get(d1, "None")
        self.config.wm2_pos = self.pos_display_to_id.get(d2, "None")
        self.config.wm_font = self.wm_font_combo.get()
        self.config.wm_alpha = float(self.wm_alpha_scale.get())
        self.config.wm_size = int(self.wm_size_spin.get())
        self.config.wm_color = self.wm_color_btn.cget("bg")
        self.config.pg_enabled = self.pg_en_var.get()
        self.config.pg_pos = self.pg_pos_var.get()
        self.config.pg_format = self.pg_fmt_var.get()
        # ページ番号位置：表示名 → 内部ID
        disp = (self.pg_pos_var.get() or "").strip()
        self.config.pg_pos = self.pg_display_to_id.get(disp, "bc")
        self.config.naming_tpl = self.naming_var.get()
        self.config.out_mode = self.out_mode_var.get()
        self.config.output_dir = self.out_dir_var.get()
        self.config.merge_all = self.merge_var.get()
        self.config.split_word_page = self.sp_word_var.get()
        self.config.split_ppt_page = self.sp_ppt_var.get()
        self.config.split_pdf_page = self.sp_pdf_var.get()
        self.config.split_excel_sheet = self.sp_ex_sheet_var.get()
        self.config.split_excel_page = self.sp_ex_page_var.get()
        self.config.auto_open = self.open_var.get()
        self.config.open_folder = self.folder_var.get()
        self.config.clear_after = self.clear_after_var.get()
        self.config.password = self.pw_var.get()
        self.config.excel_fit = self.excel_fit_var.get()
        self.config.excel_fit_tall = self.excel_fit_tall_var.get()
        self.config.clear_metadata = self.meta_var.get()
        self.config.compress_pdf = self.compress_var.get()

    def apply_config_to_ui(self):
        self.engine_combo.current(("auto", "office", "wps").index(
            self.config.engine if self.config.engine in ("auto", "office", "wps") else "auto"))
        # 任意：UI変数に内部IDが入っていたら補正
        if self.pg_pos_var.get() in ("bc", "br"):
            self.pg_pos_var.set(self.pg_id_to_display.get(self.pg_pos_var.get(), self._("pg_pos_bc")))
        self.wm1_val.set(self.config.wm1_text)
        self.wm1_pos_var.set(self.config.wm1_pos)
        self.wm2_val.set(self.config.wm2_text)
        # 透かし位置：内部ID → 表示名
        self.wm1_pos_var.set(self.pos_id_to_display.get(self.config.wm1_pos, self.pos_id_to_display["None"]))
        self.wm2_pos_var.set(self.pos_id_to_display.get(self.config.wm2_pos, self.pos_id_to_display["None"]))
        self.wm2_pos_var.set(self.config.wm2_pos)
        self.wm_font_combo.set(self.config.wm_font)
        self.wm_alpha_scale.set(self.config.wm_alpha)
        self.wm_size_spin.delete(0, tk.END)
        self.wm_size_spin.insert(0, self.config.wm_size)
        self.wm_color_btn.config(bg=self.config.wm_color)
        self.pg_en_var.set(self.config.pg_enabled)
        self.pg_fmt_var.set(self.config.pg_format)
        # ページ番号位置：内部ID → 表示名

        disp = self.pg_id_to_display.get(self.config.pg_pos, self._("pg_pos_bc"))
        self.pg_pos_var.set(disp)
        self.pg_pos_var.set(self.pg_id_to_display.get(self.config.pg_pos, "pg_pos_bc"))
        self.pg_en_var.set(self.config.pg_enabled)
        self.pg_fmt_var.set(self.config.pg_format)
        self.naming_var.set(self.config.naming_tpl)
        self.out_mode_var.set(self.config.out_mode)
        self.out_dir_var.set(self.config.output_dir)
        self.merge_var.set(self.config.merge_all)
        self.sp_word_var.set(self.config.split_word_page)
        self.sp_ppt_var.set(self.config.split_ppt_page)
        self.sp_pdf_var.set(self.config.split_pdf_page)
        self.sp_ex_sheet_var.set(self.config.split_excel_sheet)
        self.sp_ex_page_var.set(self.config.split_excel_page)
        self.open_var.set(self.config.auto_open)
        self.folder_var.set(self.config.open_folder)
        self.clear_after_var.set(self.config.clear_after)
        self.pw_var.set(self.config.password)
        self.excel_fit_var.set(self.config.excel_fit)
        self.excel_fit_tall_var.set(self.config.excel_fit_tall)
        self.meta_var.set(self.config.clear_metadata)
        self.compress_var.set(self.config.compress_pdf)
        self.update_output_preview()

    # --- UI Setup ---
    def setup_ui(self):
        # 言語切替メニュー
        menubar = tk.Menu(self.root)
        lang_menu = tk.Menu(menubar, tearoff=0)
        self._lang_names = {"ja": "日本語", "en": "English", "zh_cn": "简体中文", "zh_tw": "繁體中文"}
        self._lang_var = tk.StringVar(value=self.lang)
        for code in ("ja", "en", "zh_cn", "zh_tw"):
            lang_menu.add_radiobutton(
                label=self._lang_names[code],
                variable=self._lang_var,
                value=code,
                command=lambda c=code: self.change_language(c),
            )
        menubar.add_cascade(label="Language", menu=lang_menu)
        self.root.config(menu=menubar)

        self.root.rowconfigure(0, weight=1)
        self.root.columnconfigure(0, weight=1)
        shell = tk.Frame(self.root, padx=10, pady=5)
        shell.grid(row=0, column=0, sticky="nsew")
        shell.rowconfigure(0, weight=1)
        shell.columnconfigure(0, weight=1)
        panes = ttk.Panedwindow(shell, orient=tk.VERTICAL)
        panes.grid(row=0, column=0, sticky="nsew")
        files_host = tk.Frame(panes)
        panes.add(files_host, weight=1)
        settings_host = tk.Frame(panes)
        panes.add(settings_host, weight=3)
        settings_host.rowconfigure(0, weight=1)
        settings_host.columnconfigure(0, weight=1)
        self.settings_canvas = tk.Canvas(settings_host, highlightthickness=0)
        self.settings_canvas.grid(row=0, column=0, sticky="nsew")
        vbar = ttk.Scrollbar(settings_host, orient=tk.VERTICAL, command=self.settings_canvas.yview)
        vbar.grid(row=0, column=1, sticky="ns")
        self.settings_canvas.configure(yscrollcommand=vbar.set)
        hbar = ttk.Scrollbar(settings_host, orient=tk.HORIZONTAL, command=self.settings_canvas.xview)
        hbar.grid(row=1, column=0, sticky="ew")
        self.settings_canvas.configure(xscrollcommand=hbar.set)
        main_container = tk.Frame(self.settings_canvas, padx=2, pady=2)
        settings_window = self.settings_canvas.create_window(0, 0, anchor="nw", window=main_container)
        main_container.bind("<Configure>", lambda e: self.settings_canvas.configure(
            scrollregion=self.settings_canvas.bbox("all")))
        self.settings_canvas.bind("<Configure>", lambda e: self.settings_canvas.itemconfigure(
            settings_window, width=max(e.width, main_container.winfo_reqwidth())))
        def scroll_settings(event):
            widget = event.widget
            while widget is not None:
                if widget == main_container or widget == self.settings_canvas:
                    self.settings_canvas.yview_scroll(-int(event.delta / 120), "units")
                    return "break"
                widget = getattr(widget, "master", None)
        self.root.bind("<MouseWheel>", scroll_settings, add="+")
        def reveal_focus(event):
            widget = event.widget
            parent = widget
            while parent is not None and parent != main_container:
                parent = getattr(parent, "master", None)
            if parent is None:
                return
            for axis in ("x", "y"):
                position = (widget.winfo_rootx() - main_container.winfo_rootx() if axis == "x"
                            else widget.winfo_rooty() - main_container.winfo_rooty())
                extent = main_container.winfo_width() if axis == "x" else main_container.winfo_height()
                size = widget.winfo_width() if axis == "x" else widget.winfo_height()
                view = self.settings_canvas.xview() if axis == "x" else self.settings_canvas.yview()
                start, end = view[0] * extent, view[1] * extent
                target = position if position < start else position + size - (end - start)
                if position < start or position + size > end:
                    move = self.settings_canvas.xview_moveto if axis == "x" else self.settings_canvas.yview_moveto
                    move(max(0, target) / max(1, extent))
        self.root.bind("<FocusIn>", reveal_focus, add="+")
        footer = tk.Frame(shell)
        footer.grid(row=1, column=0, sticky="ew")

        # File List
        file_frame = tk.LabelFrame(files_host, text=self._("file_list"), padx=5, pady=5)
        file_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        tree_host = tk.Frame(file_frame)
        tree_host.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        tree_host.rowconfigure(0, weight=1)
        tree_host.columnconfigure(0, weight=1)
        self.tree = ttk.Treeview(tree_host, columns=("Type", "Name", "Range", "Out"), show="headings", height=5)
        for col, head, w in [
            ("Type", self._("col_type"), 70),
            ("Name", self._("col_name"), 300),
            ("Range", self._("col_range"), 150),
            ("Out", self._("col_out"), 300),
        ]:
            self.tree.heading(col, text=head)
            self.tree.column(col, width=w)
        self.tree.grid(row=0, column=0, sticky="nsew")
        tree_y = ttk.Scrollbar(tree_host, orient=tk.VERTICAL, command=self.tree.yview)
        tree_y.grid(row=0, column=1, sticky="ns")
        tree_x = ttk.Scrollbar(tree_host, orient=tk.HORIZONTAL, command=self.tree.xview)
        tree_x.grid(row=1, column=0, sticky="ew")
        self.tree.configure(yscrollcommand=tree_y.set, xscrollcommand=tree_x.set)
        self.tree.drop_target_register(DND_FILES)
        self.tree.dnd_bind("<<Drop>>", self.handle_drop)
        self.tree.bind("<Double-1>", self.on_list_double_click)

        btn_f = tk.Frame(file_frame)
        btn_f.pack(fill=tk.X, padx=5, pady=2)
        tk.Button(btn_f, text=self._("btn_up"), command=lambda: self.move_file(-1)).pack(side=tk.LEFT, padx=2)
        tk.Button(btn_f, text=self._("btn_down"), command=lambda: self.move_file(1)).pack(side=tk.LEFT, padx=2)
        tk.Button(btn_f, text=self._("btn_remove"), command=self.remove_file).pack(side=tk.LEFT, padx=2)
        tk.Button(btn_f, text=self._("btn_clear_list"), command=self.clear_list).pack(side=tk.LEFT, padx=2)
        tk.Button(btn_f, text=self._("btn_add_folder"), command=self.add_folder).pack(side=tk.LEFT, padx=2)
        # Preset Area
        pre_f = tk.Frame(main_container)
        pre_f.pack(fill=tk.X, pady=5)
        tk.Label(pre_f, text=self._("lbl_presets")).pack(side=tk.LEFT)
        self.preset_combo = ttk.Combobox(pre_f, values=list(self.presets.keys()), width=25)
        self.preset_combo.pack(side=tk.LEFT, padx=5)
        tk.Button(pre_f, text=self._("btn_load"), command=self.load_preset).pack(side=tk.LEFT, padx=2)
        tk.Button(pre_f, text=self._("btn_save"), command=self.save_preset).pack(side=tk.LEFT, padx=2)
        tk.Button(pre_f, text=self._("btn_delete"), command=self.delete_preset).pack(side=tk.LEFT, padx=2)

        # Mid Columns
        mid_frame = tk.Frame(main_container)
        mid_frame.pack(fill=tk.X, pady=5)

        # 翻訳された表示名と内部IDを紐付ける辞書を作成
        self.pos_display_to_id = {self._(v): k for k, v in POS_MAP}
        self.pos_id_to_display = {k: self._(v) for k, v in POS_MAP}
        pos_options = list(self.pos_display_to_id.keys())

        # Left: Watermark (UI構築部分)
        wm_frame = tk.LabelFrame(mid_frame, text=self._("frame_wm"), padx=10, pady=5)
        wm_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))

        f_row = tk.Frame(wm_frame)
        f_row.pack(fill=tk.X, pady=2)
        tk.Label(f_row, text=self._("lbl_font")).pack(side=tk.LEFT)
        self.wm_font_combo = ttk.Combobox(f_row, width=25)
        self.wm_font_combo.pack(side=tk.LEFT, padx=5)
        tk.Label(f_row, text=self._("lbl_size")).pack(side=tk.LEFT)
        self.wm_size_spin = tk.Spinbox(f_row, from_=10, to=300, width=5)
        self.wm_size_spin.pack(side=tk.LEFT, padx=5)
        self.wm_color_btn = tk.Button(f_row, bg="#C0C0C0", width=2, command=self.choose_color)
        self.wm_color_btn.pack(side=tk.LEFT, padx=5)

        a_row = tk.Frame(wm_frame)
        a_row.pack(fill=tk.X, pady=2)
        tk.Label(a_row, text=self._("lbl_alpha")).pack(side=tk.LEFT)
        self.wm_alpha_scale = tk.Scale(a_row, from_=0.0, to=1.0, resolution=0.05, orient=tk.HORIZONTAL, length=120)
        self.wm_alpha_scale.pack(side=tk.LEFT, padx=5)
        # プレビューボタン追加
        tk.Button(a_row, text=self._("btn_preview_wm"), bg="#e1e1e1", command=self.preview_watermark).pack(
            side=tk.RIGHT
        )

        for i in [1, 2]:
            r = tk.Frame(wm_frame)
            r.pack(fill=tk.X, pady=2)
            tk.Label(r, text=f"{self._('wm_label')}{i}:", width=9).pack(side=tk.LEFT)
            val = ttk.Combobox(r, values=self.wm_templates, width=20)
            val.pack(side=tk.LEFT, padx=2)
            setattr(self, f"wm{i}_val", val)

            p_var = tk.StringVar(value=self.pos_id_to_display["None"])
            cb = ttk.Combobox(r, textvariable=p_var, values=pos_options, width=15, state="readonly")
            cb.pack(side=tk.LEFT, padx=2)
            setattr(self, f"wm{i}_pos_var", p_var)

        pg_row = tk.Frame(wm_frame)
        pg_row.pack(fill=tk.X, pady=5)

        # --- page number position display <-> id mapping ---
        # ページ番号位置（表示名 I18N 対応）
        self.pg_id_to_display = {
            "bc": self._("pg_pos_bc"),
            "br": self._("pg_pos_br"),
        }
        self.pg_display_to_id = {v: k for k, v in self.pg_id_to_display.items()}

        self.pg_en_var = tk.BooleanVar()
        tk.Checkbutton(pg_row, text=self._("lbl_page_num"), variable=self.pg_en_var).pack(side=tk.LEFT)
        self.pg_fmt_var = tk.StringVar()
        tk.Entry(pg_row, textvariable=self.pg_fmt_var, width=18).pack(side=tk.LEFT, padx=2)
        self.pg_pos_var = tk.StringVar(value=self._("pg_pos_bc"))

        ttk.Combobox(
            pg_row,
            textvariable=self.pg_pos_var,
            values=list(self.pg_display_to_id.keys()),
            width=18,
            state="readonly",
        ).pack(side=tk.LEFT, padx=2)

        # Right: Split Settings
        split_frame = tk.LabelFrame(mid_frame, text=self._("frame_detail"), padx=10, pady=5)
        split_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.merge_var = tk.BooleanVar()
        tk.Checkbutton(
            split_frame,
            text=self._("chk_merge"),
            variable=self.merge_var,
            font=("", 9, "bold"),
            fg="blue",
            command=self.toggle_merge_all,
        ).pack(anchor="w")

        f1 = tk.Frame(split_frame)
        f1.pack(fill=tk.X, pady=(5, 0))
        tk.Label(f1, text="Word/PPT:", width=10, anchor="w").pack(side=tk.LEFT)
        self.sp_word_var, self.sp_ppt_var = tk.BooleanVar(), tk.BooleanVar()
        tk.Checkbutton(
            f1, text=self._("chk_split_page"), variable=self.sp_word_var, command=self.toggle_split_box
        ).pack(side=tk.LEFT)
        tk.Checkbutton(
            f1, text=self._("chk_include_ppt"), variable=self.sp_ppt_var, command=self.toggle_split_box
        ).pack(side=tk.LEFT)

        f2 = tk.Frame(split_frame)
        f2.pack(fill=tk.X)
        tk.Label(f2, text="PDF:", width=10, anchor="w").pack(side=tk.LEFT)
        self.sp_pdf_var = tk.BooleanVar()
        tk.Checkbutton(
            f2, text=self._("chk_by_all_pages"), variable=self.sp_pdf_var, command=self.toggle_split_box
        ).pack(side=tk.LEFT)

        f3 = tk.Frame(split_frame)
        f3.pack(fill=tk.X)
        tk.Label(f3, text="Excel:", width=10, anchor="w").pack(side=tk.LEFT)
        self.sp_ex_sheet_var, self.sp_ex_page_var = tk.BooleanVar(), tk.BooleanVar()
        tk.Checkbutton(
            f3, text=self._("chk_by_sheet"), variable=self.sp_ex_sheet_var, command=self.toggle_excel_sheet
        ).pack(side=tk.LEFT)
        tk.Checkbutton(
            f3, text=self._("chk_by_all_pages"), variable=self.sp_ex_page_var, command=self.toggle_excel_page
        ).pack(side=tk.LEFT)

        tk.Label(split_frame, text=self._("lbl_excel_opt"), fg="gray", font=("", 8)).pack(anchor="w", pady=(5, 0))
        f4 = tk.Frame(split_frame)
        f4.pack(fill=tk.X)
        self.excel_fit_var, self.excel_fit_tall_var = tk.BooleanVar(), tk.BooleanVar()
        tk.Checkbutton(f4, text=self._("chk_fit_width"), variable=self.excel_fit_var).pack(side=tk.LEFT)
        tk.Checkbutton(f4, text=self._("chk_fit_height"), variable=self.excel_fit_tall_var).pack(side=tk.LEFT)

        # Bottom
        bottom_frame = tk.LabelFrame(main_container, text=self._("frame_exec"), padx=10, pady=5)
        bottom_frame.pack(fill=tk.X, pady=5)

        n_row = tk.Frame(bottom_frame)
        n_row.pack(fill=tk.X, pady=2)
        tk.Label(n_row, text=self._("lbl_naming")).pack(side=tk.LEFT)
        self.naming_var = tk.StringVar()
        self.naming_combo = ttk.Combobox(n_row, textvariable=self.naming_var, values=self.nm_templates, width=40)
        self.naming_combo.pack(side=tk.LEFT, padx=5)

        # ヘルプボタンを追加
        tk.Button(n_row, text=self._("btn_tag_help"), command=self.show_naming_help, bg="#f0f0f0").pack(side=tk.LEFT)

        self.final_name_label = tk.Label(n_row, text="", fg="#0056b3", font=("", 9, "bold"))
        self.final_name_label.pack(side=tk.LEFT, padx=10)

        o_row = tk.Frame(bottom_frame)
        o_row.pack(fill=tk.X, pady=2)
        self.out_mode_var = tk.StringVar()
        tk.Radiobutton(
            o_row,
            text=self._("opt_same_dir"),
            variable=self.out_mode_var,
            value="original",
            command=self.update_output_preview,
        ).pack(side=tk.LEFT)
        tk.Radiobutton(
            o_row,
            text=self._("opt_custom_dir"),
            variable=self.out_mode_var,
            value="custom",
            command=self.update_output_preview,
        ).pack(side=tk.LEFT)
        self.out_dir_var = tk.StringVar()
        tk.Entry(o_row, textvariable=self.out_dir_var, width=40).pack(side=tk.LEFT, padx=5)
        tk.Button(o_row, text=self._("btn_browse"), command=self.browse_dir).pack(side=tk.LEFT)

        s_row = tk.Frame(bottom_frame)
        s_row.pack(fill=tk.X, pady=2)
        self.pw_var = tk.StringVar()
        tk.Label(s_row, text=self._("lbl_password")).pack(side=tk.LEFT)
        tk.Entry(s_row, textvariable=self.pw_var, width=12).pack(side=tk.LEFT, padx=5)
        self.meta_var, self.compress_var = tk.BooleanVar(), tk.BooleanVar()
        tk.Checkbutton(s_row, text=self._("chk_meta_clear"), variable=self.meta_var).pack(side=tk.LEFT, padx=5)
        tk.Checkbutton(s_row, text=self._("chk_compress"), variable=self.compress_var).pack(side=tk.LEFT, padx=5)
        self.open_var, self.folder_var, self.clear_after_var = tk.BooleanVar(), tk.BooleanVar(), tk.BooleanVar()
        tk.Checkbutton(s_row, text=self._("chk_open_done"), variable=self.open_var).pack(side=tk.LEFT, padx=5)
        tk.Checkbutton(s_row, text=self._("chk_open_folder"), variable=self.folder_var).pack(side=tk.LEFT, padx=5)
        tk.Checkbutton(s_row, text=self._("chk_clear_after"), variable=self.clear_after_var).pack(side=tk.LEFT, padx=5)
        engine_row = tk.Frame(main_container)
        engine_row.pack(fill=tk.X, pady=5, before=pre_f)
        labels = {"zh_cn": ("转换引擎：", "自动（优先 Office）", "显示日志"),
                  "zh_tw": ("轉換引擎：", "自動（優先 Office）", "顯示日誌"),
                  "ja": ("変換エンジン:", "自動（Office 優先）", "ログ表示"),
                  "en": ("Engine:", "Auto (Office first)", "Show log")}
        engine_label, auto_label, log_label = labels[self.lang]
        tk.Label(engine_row, text=engine_label).pack(side=tk.LEFT)
        self.engine_combo = ttk.Combobox(engine_row, state="readonly", width=25,
                                         values=(auto_label, "Microsoft Office", "WPS Office"))
        self.engine_combo.current(0)
        self.engine_combo.pack(side=tk.LEFT)
        self.engine_combo.bind("<<ComboboxSelected>>", lambda e: self.update_config_from_ui())
        def reflow(event):
            stacked = event.width < 880
            wm_frame.pack_configure(side=tk.TOP if stacked else tk.LEFT, fill=tk.X if stacked else tk.BOTH)
            split_frame.pack_configure(side=tk.TOP if stacked else tk.LEFT, fill=tk.X if stacked else tk.BOTH)
        shell.bind("<Configure>", reflow)
        exec_f = tk.Frame(footer)
        exec_f.pack(fill=tk.X, pady=10)
        self.btn_convert = tk.Button(
            exec_f,
            text=self._("btn_start"),
            bg="#28a745",
            fg="white",
            font=("", 12, "bold"),
            height=2,
            command=self.start_thread,
        )
        self.btn_convert.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=2)
        self.btn_cancel = tk.Button(
            exec_f,
            text=self._("btn_cancel"),
            bg="#dc3545",
            fg="white",
            font=("", 12, "bold"),
            height=2,
            command=self.cancel_process,
            state=tk.DISABLED,
        )
        self.btn_cancel.pack(side=tk.LEFT, padx=2)

        self.progress_label = tk.Label(footer, text=self._("st_ready"), anchor="w")
        self.progress_label.pack(fill=tk.X)
        self.progress = ttk.Progressbar(footer, orient=tk.HORIZONTAL, mode="determinate")
        self.progress.pack(fill=tk.X)

        log_f = tk.Frame(footer)
        self.show_log_var = tk.BooleanVar(value=False)
        self.log_toggle = ttk.Checkbutton(footer, text=log_label, variable=self.show_log_var,
                        command=lambda: log_f.pack(fill=tk.X, pady=5) if self.show_log_var.get()
                        else log_f.pack_forget())
        self.log_toggle.pack(anchor="w")
        self.log_text = tk.Text(log_f, height=3, bg="#1e1e1e", fg="#00ff00", font=("Consolas", 9))
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb = tk.Scrollbar(log_f, command=self.log_text.yview)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.config(yscrollcommand=sb.set, state=tk.DISABLED)

        self.naming_var.trace_add("write", lambda *_: self.update_output_preview())
        self.out_dir_var.trace_add("write", lambda *_: self.update_output_preview())

    # --- Mutual Exclusion ---
    def toggle_merge_all(self):
        if self.merge_var.get():
            for v in [self.sp_word_var, self.sp_ppt_var, self.sp_pdf_var, self.sp_ex_sheet_var, self.sp_ex_page_var]:
                v.set(False)
        self.update_output_preview()

    def toggle_split_box(self):
        if any([self.sp_word_var.get(), self.sp_ppt_var.get(), self.sp_pdf_var.get()]):
            self.merge_var.set(False)
        self.update_output_preview()

    def toggle_excel_sheet(self):
        if self.sp_ex_sheet_var.get():
            self.merge_var.set(False)
            self.sp_ex_page_var.set(False)
        self.update_output_preview()

    def toggle_excel_page(self):
        if self.sp_ex_page_var.get():
            self.merge_var.set(False)
            self.sp_ex_sheet_var.set(False)
        self.update_output_preview()

    # --- Font Logic (Fixed for TTC) ---
    def build_registry_font_items(self):
        items_for_combo = []
        self.font_map = {}
        if winreg is None:
            return []
        fonts_dir = os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts")
        roots = [
            (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"),
            (winreg.HKEY_CURRENT_USER, r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"),
        ]
        seen_names = set()
        for root, keypath in roots:
            try:
                with winreg.OpenKey(root, keypath) as k:
                    i = 0
                    while True:
                        try:
                            name_raw, val, _ = winreg.EnumValue(k, i)
                            i += 1
                            if not isinstance(val, str):
                                continue
                            file_name = re.split(r"[,&]", val)[0].strip()
                            full_path = os.path.join(fonts_dir, file_name)
                            if not os.path.exists(full_path):
                                continue
                            clean_raw = re.sub(r"\s*\(.*?\)", "", name_raw).replace(";", "")
                            sub_names = [n.strip() for n in clean_raw.split("&")]
                            for idx, sub_name in enumerate(sub_names):
                                if sub_name and sub_name not in seen_names:
                                    seen_names.add(sub_name)
                                    self.font_map[sub_name] = (full_path, idx)
                                    items_for_combo.append(sub_name)
                        except OSError:
                            break
            except:
                continue
        items_for_combo.sort()
        return items_for_combo

    def load_fonts_delayed(self):
        """言語に合わせた最適な初期フォントを選択する"""
        try:
            names = self.build_registry_font_items()
            self.wm_font_combo["values"] = names

            # 言語別推奨フォントリスト（上から順に探す）
            if self.lang == "ja":
                target_fonts = [
                    "BIZ UDPGothic",
                    "BIZ UDPゴシック",
                    "MS Gothic",
                    "ＭＳ ゴシック",
                    "Yu Gothic",
                    "游ゴシック",
                ]
            elif self.lang in ("zh_cn", "zh_tw"):
                target_fonts = [
                    "Microsoft YaHei",
                    "微软雅黑",
                    "SimSun",
                    "宋体",
                    "PMingLiU",
                    "新細明體",
                    "MingLiU",
                ]
            else:
                target_fonts = ["Arial", "Calibri", "Segoe UI", "Verdana"]

            selected_font = ""
            for tf in target_fonts:
                if tf in names:
                    selected_font = tf
                    break

            if not selected_font and names:
                selected_font = names[0]  # 見つからなければリストの先頭

            if self.config.wm_font in names:
                self.wm_font_combo.set(self.config.wm_font)
            else:
                self.wm_font_combo.set(selected_font)

            self.queue_log(self._("log_font_loaded"))
        except Exception as e:
            self.queue_log(f"{self._('log_font_err')}: {e}")

    def _register_reportlab_font(self, chosen_font: str) -> Tuple[str, str]:
        chosen = chosen_font or ""
        if chosen in self.font_map:
            p, idx = self.font_map[chosen]
            internal_name = f"WM_{chosen.replace(' ', '_')}_{idx}"
            try:
                if internal_name not in pdfmetrics.getRegisteredFontNames():
                    if p.lower().endswith(".ttc"):
                        pdfmetrics.registerFont(TTFont(internal_name, p, subfontIndex=idx))
                    else:
                        pdfmetrics.registerFont(TTFont(internal_name, p))
                return internal_name, f"{self._('log_font_using')}: {chosen}"
            except Exception as e:
                self.queue_log(f"{self._('log_font_fail')}: {e}")

        try:
            pdfmetrics.registerFont(UnicodeCIDFont("HeiseiKakuGo-W5"))
            return "HeiseiKakuGo-W5", f"{self._('log_font_using')}: HeiseiKakuGo-W5 (CID)"
        except:
            return "Helvetica", f"{self._('log_font_using')}: Helvetica"

    # --- Core Processing ---
    def start_thread(self):
        if not self.files:
            messagebox.showwarning(self._("title_warn"), self._("msg_no_files"))
            return

        # UIスレッドで確定値を config に落とす
        self.update_config_from_ui()
        self.save_config()

        # ワーカースレッド用にスナップショットを保持（Tk参照禁止にする）
        cfg = AppConfig(**asdict(self.config))

        self.processing = True
        self.cancel_flag.clear()
        self.btn_convert.config(state=tk.DISABLED)
        self.btn_cancel.config(state=tk.NORMAL)

        threading.Thread(target=self.main_process, args=(cfg,), daemon=True).start()

    def main_process(self, cfg: AppConfig):
        pythoncom.CoInitialize()
        with tempfile.TemporaryDirectory() as tmp_dir:
            try:
                temp_units = []
                self.queue_progress(max=len(self.files), progress=0, label=self._("st_converting"))

                for i, f in enumerate(self.files):
                    if self.cancel_flag.is_set():
                        break

                    self.queue_progress(
                        progress=i + 1,
                        label=f"{self._('st_conv_file')} {os.path.basename(f['path'])}",
                    )

                    if f["type"] == "Excel":
                        units = self.cv_excel_units(f, tmp_dir, cfg)
                        for u_path, s_name in units:
                            temp_units.append({"path": u_path, "orig": f, "sheet": s_name, "fseq": i + 1})
                    else:
                        tmp_p = os.path.join(tmp_dir, f"conv_{i}.pdf")
                        ok = False

                        if f["type"] == "Word":
                            ok = self.cv_word(f, tmp_p, cfg)
                        elif f["type"] == "PowerPoint":
                            ok = self.cv_ppt(f, tmp_p, cfg)
                        elif f["type"] == "Image":
                            ok = self.cv_img(f, tmp_p)
                        elif f["type"] == "PDF":
                            ok = self.cv_pdf(f, tmp_p)

                        if ok:
                            # --- F: range指定をここでPDFへ反映 ---
                            r_spec = f.get("range", "")
                            if not self._is_all_range(r_spec):
                                tmp_r = os.path.join(tmp_dir, f"range_{i}.pdf")
                                if self.apply_range_to_pdf(tmp_p, r_spec, tmp_r):
                                    tmp_p = tmp_r
                                else:
                                    self.queue_log(
                                        f"{self._('log_conv_fail')} {os.path.basename(f['path'])} (range empty)"
                                    )
                                    continue

                            temp_units.append({"path": tmp_p, "orig": f, "sheet": "", "fseq": i + 1})

                if not temp_units or self.cancel_flag.is_set():
                    return

                self.queue_progress(label=self._("st_finalizing"))
                global_seq = 1
                dest_to_open = ""

                if cfg.merge_all:
                    dest = self.get_final_dest(temp_units[0], 1, 1, 1)
                    if dest:
                        self.finalize_pdfs([u["path"] for u in temp_units], dest, temp_units, cfg)
                        dest_to_open = dest
                else:
                    for i, f_orig in enumerate(self.files):
                        if self.cancel_flag.is_set():
                            break

                        u_list = [u for u in temp_units if u["orig"] == f_orig]
                        if not u_list:
                            continue

                        t = f_orig["type"]
                        do_pg = (
                            (t == "Word" and cfg.split_word_page)
                            or (t == "PowerPoint" and cfg.split_ppt_page)
                            or (t == "PDF" and cfg.split_pdf_page)
                            or (t == "Excel" and cfg.split_excel_page)
                        )
                        do_sh = t == "Excel" and cfg.split_excel_sheet

                        if do_pg:
                            for u in u_list:
                                reader = PdfReader(u["path"])
                                p_total = len(reader.pages)
                                for p_idx in range(p_total):
                                    dest = self.get_final_dest(u, global_seq, i + 1, p_idx + 1, p_total)
                                    if dest:
                                        writer = PdfWriter()
                                        writer.add_page(reader.pages[p_idx])
                                        tmp_s = os.path.join(tmp_dir, "split.pdf")
                                        with open(tmp_s, "wb") as fs:
                                            writer.write(fs)
                                        self.finalize_pdfs([tmp_s], dest, [u], cfg, p_idx + 1, p_total)
                                        global_seq += 1
                                        dest_to_open = dest

                        elif do_sh:
                            for u in u_list:
                                dest = self.get_final_dest(u, global_seq, i + 1, 1)
                                if dest:
                                    self.finalize_pdfs([u["path"]], dest, [u], cfg)
                                    global_seq += 1
                                    dest_to_open = dest
                        else:
                            dest = self.get_final_dest(u_list[0], global_seq, i + 1, 1)
                            if dest:
                                self.finalize_pdfs([u["path"] for u in u_list], dest, u_list, cfg)
                                global_seq += 1
                                dest_to_open = dest

                self.finish_action(dest_to_open, cfg)

            except Exception as e:
                self.queue_log(f"{self._('log_fatal')}: {e}")
            finally:
                self.processing = False
                self.root.after(0, lambda: self.btn_convert.config(state=tk.NORMAL))
                self.root.after(0, lambda: self.btn_cancel.config(state=tk.DISABLED))
                pythoncom.CoUninitialize()

    # --- Office Converters (with COM cleanup & Excel Print Area) ---
    def validate_export(self, path):
        if not os.path.isfile(path) or os.path.getsize(path) == 0:
            raise RuntimeError(f"PDF export missing: {path}")
        if not PdfReader(path).pages:
            raise RuntimeError(f"PDF export has no pages: {path}")

    def cv_excel_units(self, f: dict, tmp_dir: str, cfg: AppConfig) -> List[Tuple[str, str]]:
        units = []
        try:
            with application("Excel", cfg.engine, self.queue_log) as session:
                excel = session.app
                wb = None
                try:
                    excel.DisplayAlerts = False
                    wb = excel.Workbooks.Open(os.path.abspath(f["path"]), ReadOnly=True)
                    session.observe_document(wb)
                    names = [s.strip() for s in f.get("range", "").split(",") if s.strip()]
                    if not names or self._is_all_range(f.get("range", "")):
                        names = [s.Name for s in wb.Worksheets]
                    for name in names:
                        if self.cancel_flag.is_set():
                            break
                        try:
                            ws = wb.Worksheets(name)
                            if ws.Visible != -1:
                                continue
                            if cfg.excel_fit or cfg.excel_fit_tall:
                                ws.PageSetup.Zoom = False
                                if cfg.excel_fit:
                                    ws.PageSetup.FitToPagesWide = 1
                                if cfg.excel_fit_tall:
                                    ws.PageSetup.FitToPagesTall = 1
                            fd, target = tempfile.mkstemp(prefix="sheet_", suffix=".pdf", dir=tmp_dir)
                            os.close(fd)
                            ws.ExportAsFixedFormat(0, os.path.abspath(target))
                            self.validate_export(target)
                            units.append((target, name))
                        except Exception as exc:
                            self.queue_log(f"{self._('log_conv_fail')} {f['path']} [{name}]: {exc}")
                finally:
                    close_document(wb, self.queue_log)
        except Exception as exc:
            self.queue_log(f"{self._('log_conv_fail')} {f['path']}: {exc}")
        return units

    def cv_document(self, kind, f, out, cfg):
        try:
            with application(kind, cfg.engine, self.queue_log) as session:
                app = session.app
                doc = None
                try:
                    if kind == "Word":
                        doc = app.Documents.Open(os.path.abspath(f["path"]), ReadOnly=True)
                        session.observe_document(doc)
                        doc.ExportAsFixedFormat(os.path.abspath(out), 17)
                    else:
                        doc = app.Presentations.Open(os.path.abspath(f["path"]), True, False, False)
                        session.observe_document(doc)
                        doc.ExportAsFixedFormat(os.path.abspath(out), 2, PrintRange=None)
                    self.validate_export(out)
                    return True
                finally:
                    close_document(doc, self.queue_log, presentation=kind == "PowerPoint")
        except Exception as exc:
            self.queue_log(f"{self._('log_conv_fail')} {f['path']}: {exc}")
            return False

    def cv_word(self, f, out, cfg):
        return self.cv_document("Word", f, out, cfg)

    def cv_ppt(self, f, out, cfg):
        return self.cv_document("PowerPoint", f, out, cfg)

    def cv_img(self, f, out):
        try:
            with Image.open(f["path"]) as img:
                img.convert("RGB").save(out, "PDF")
            return True
        except:
            return False

    def cv_pdf(self, f, out):
        try:
            reader = PdfReader(f["path"])
            writer = PdfWriter()
            for p in reader.pages:
                writer.add_page(p)
            with open(out, "wb") as fs:
                writer.write(fs)
            return True
        except:
            return False

    # --- Watermark & Finalize ---
    def finalize_pdfs(
        self,
        src_list: List[str],
        dest: str,
        units: List[dict],
        cfg: AppConfig,
        page_offset: int = 1,
        total_override: int = 0,
    ):
        writer = PdfWriter()
        font_name, _ = self._register_reportlab_font(cfg.wm_font)

        readers = [PdfReader(s) for s in src_list]
        total_p = total_override if total_override > 0 else sum(len(r.pages) for r in readers)
        curr_p = page_offset

        # 透かし有無
        has_wm = any([cfg.wm1_text and cfg.wm1_pos != "None", cfg.wm2_text and cfg.wm2_pos != "None"])
        has_pg = cfg.pg_enabled

        for r in readers:
            for page in r.pages:
                page.transfer_rotation_to_content()

                if has_wm or has_pg:
                    w, h = float(page.mediabox.width), float(page.mediabox.height)
                    packet = io.BytesIO()
                    c = canvas.Canvas(packet, pagesize=(w, h))

                    # ---- Watermark (1/2) ----
                    for txt_raw, pos_id in [(cfg.wm1_text, cfg.wm1_pos), (cfg.wm2_text, cfg.wm2_pos)]:
                        if pos_id == "None" or not txt_raw:
                            continue

                        txt = self.apply_tags(txt_raw, units[0], curr_p, units[0].get("fseq", 1), curr_p, total_p)
                        c.saveState()

                        f_size = int(cfg.wm_size)
                        c.setFont(font_name, f_size)

                        rgb = [int(cfg.wm_color.lstrip("#")[j : j + 2], 16) / 255 for j in (0, 2, 4)]
                        c.setFillColorRGB(*rgb, alpha=float(cfg.wm_alpha))

                        if pos_id == "diag":
                            c.translate(w / 2, h / 2)
                            c.rotate(45)
                            c.drawCentredString(0, 0, txt)
                        elif pos_id == "large":
                            c.drawCentredString(w / 2, h / 2, txt)
                        else:
                            tw = c.stringWidth(txt, font_name, f_size)

                            if "l" in pos_id:
                                tx = 20
                            elif "r" in pos_id:
                                tx = w - tw - 20
                            else:
                                tx = (w - tw) / 2

                            if "t" in pos_id:
                                ty = h - f_size - 20
                            elif "b" in pos_id:
                                ty = 20
                            else:
                                ty = h / 2

                            c.drawString(tx, ty, txt)

                        c.restoreState()

                    # ---- Page number ----

                    if has_pg:
                        pg_txt = cfg.pg_format.replace("{n}", str(curr_p)).replace("{total}", str(total_p))
                        pg_txt = self.apply_tags(pg_txt, units[0], curr_p, units[0].get("fseq", 1), curr_p, total_p)

                        c.saveState()

                        # ★固定：10.5pt / 黒（必ずここで定義）
                        pg_size = 10.5
                        c.setFont(font_name, pg_size)
                        c.setFillColorRGB(0, 0, 0)  # 黒固定

                        tw = c.stringWidth(pg_txt, font_name, pg_size)

                        # bc=中央下, br=右下
                        margin_x = 20
                        margin_y = 24
                        if cfg.pg_pos == "br":
                            x = w - tw - margin_x
                            y = margin_y
                        else:  # "bc" default
                            x = (w - tw) / 2
                            y = margin_y

                        c.drawString(x, y, pg_txt)
                        c.restoreState()

                    c.showPage()
                    c.save()
                    packet.seek(0)
                    page.merge_page(PdfReader(packet).pages[0])

                writer.add_page(page)
                curr_p += 1

        if cfg.clear_metadata:
            writer.add_metadata({})
        if cfg.password:
            writer.encrypt(cfg.password)
        if cfg.compress_pdf:
            if hasattr(writer, "compress_contents"):
                writer.compress_contents()
            elif hasattr(writer, "compress_content_streams"):
                writer.compress_content_streams()

        with open(dest, "wb") as f:
            writer.write(f)

    # --- Preview Feature ---
    def preview_watermark(self):
        if not self.files:
            messagebox.showinfo(self._("title_info"), self._("msg_no_preview"))
            return

        # UIスレッドで確定値をconfigへ
        self.update_config_from_ui()
        cfg = AppConfig(**asdict(self.config))

        # できれば選択中、なければ先頭
        sel = self.tree.selection()
        f = self.files[self.tree.index(sel[0])] if sel else self.files[0]

        def _task(cfg_snapshot: AppConfig, f_info: dict):
            pythoncom.CoInitialize()
            with tempfile.TemporaryDirectory() as tmp_dir:
                try:
                    self.queue_log(f"{self._('st_preview_gen')} {os.path.basename(f_info['path'])}")
                    tmp_pdf = os.path.join(tmp_dir, "preview_base.pdf")
                    ok = False

                    if f_info["type"] == "Excel":
                        res = self.cv_excel_units(f_info, tmp_dir, cfg_snapshot)
                        if not res:
                            return
                        tmp_pdf = res[0][0]
                        ok = True
                    else:
                        if f_info["type"] == "Word":
                            ok = self.cv_word(f_info, tmp_pdf, cfg_snapshot)
                        elif f_info["type"] == "PowerPoint":
                            ok = self.cv_ppt(f_info, tmp_pdf, cfg_snapshot)
                        elif f_info["type"] == "Image":
                            ok = self.cv_img(f_info, tmp_pdf)
                        elif f_info["type"] == "PDF":
                            ok = self.cv_pdf(f_info, tmp_pdf)

                    if not ok:
                        self.queue_log(f"{self._('log_conv_fail')} {os.path.basename(f_info['path'])}")
                        return

                    # range指定を反映してから1ページ目を作る
                    r_spec = f_info.get("range", "")
                    if not self._is_all_range(r_spec):
                        tmp_r = os.path.join(tmp_dir, "preview_range.pdf")
                        if self.apply_range_to_pdf(tmp_pdf, r_spec, tmp_r):
                            tmp_pdf = tmp_r
                        else:
                            self.queue_log(
                                f"{self._('log_conv_fail')} {os.path.basename(f_info['path'])} (range empty)"
                            )
                            return

                    reader = PdfReader(tmp_pdf)
                    if not reader.pages:
                        return

                    writer = PdfWriter()
                    writer.add_page(reader.pages[0])
                    tmp_one = os.path.join(tmp_dir, "one.pdf")
                    with open(tmp_one, "wb") as fs:
                        writer.write(fs)

                    out_p = os.path.join(tempfile.gettempdir(), "PDFPro_Preview.pdf")
                    unit = {"orig": f_info, "sheet": "Preview", "fseq": 1}
                    self.finalize_pdfs([tmp_one], out_p, [unit], cfg_snapshot, 1, 1)

                    os.startfile(out_p)
                    self.queue_log(self._("msg_preview_ok"))

                except Exception as e:
                    self.queue_log(f"{self._('msg_preview_fail')} {e}")
                finally:
                    pythoncom.CoUninitialize()

        threading.Thread(target=_task, args=(cfg, f), daemon=True).start()

    def _is_all_range(self, s: str) -> bool:
        s = (s or "").strip()
        if not s:
            return True
        return s in ("全ページ", "All Pages", self._("val_all_pages"))

    def parse_page_spec(self, spec: str, total_pages: int) -> List[int]:
        """
        spec例: "1-3,5,8-" / "2" / "1-" / "-3"（-3は1-3扱い）
        戻り値: 0-based page indices（重複排除、昇順）
        """
        if self._is_all_range(spec):
            return list(range(total_pages))

        spec = spec.replace(" ", "")
        out = set()

        for token in [t for t in spec.split(",") if t]:
            m = re.fullmatch(r"(\d+)?-(\d+)?", token)
            if m:
                a, b = m.group(1), m.group(2)
                start = int(a) if a else 1
                end = int(b) if b else total_pages
                start = max(1, start)
                end = min(total_pages, end)
                if start <= end:
                    for p in range(start, end + 1):
                        out.add(p - 1)
                continue

            if token.isdigit():
                p = int(token)
                if 1 <= p <= total_pages:
                    out.add(p - 1)

        return sorted(out)

    def apply_range_to_pdf(self, src_pdf: str, range_spec: str, dst_pdf: str) -> bool:
        """
        src_pdf を range_spec に従って抽出して dst_pdf へ。
        range_spec が全ページなら単純コピー（読み書き）する。
        """
        try:
            r = PdfReader(src_pdf)
            total = len(r.pages)
            idxs = self.parse_page_spec(range_spec, total)
            if not idxs:
                return False

            w = PdfWriter()
            for i in idxs:
                w.add_page(r.pages[i])

            with open(dst_pdf, "wb") as f:
                w.write(f)
            return True
        except:
            return False

    # --- Utils ---
    def apply_tags(self, tpl: str, u_info: dict, seq: int, fseq: int, pseq: int, ptotal: int = 1) -> str:
        now = datetime.datetime.now()
        path = u_info["orig"]["path"]
        name = os.path.splitext(os.path.basename(path))[0]
        sheet = u_info.get("sheet", "")
        parent = os.path.basename(os.path.dirname(path))  # 親フォルダ名

        res = tpl
        res = res.replace("{name}", name)
        res = res.replace("{sheet}", sheet)
        res = res.replace("{parent}", parent)
        res = res.replace("{seq}", str(seq))
        res = res.replace("{fseq}", str(fseq))
        res = res.replace("{pseq}", str(pseq))
        res = res.replace("{total}", str(len(self.files)))
        res = res.replace("{ptotal}", str(ptotal))
        res = res.replace("{username}", os.getlogin())
        res = res.replace("{rand}", f"{random.randint(0, 9999):04d}")

        # 日付・時刻タグの処理 {date:yyyy-mm-dd HH:MM:SS} 等
        def _repl(m):
            fmt = m.group(1)
            # Pythonのstrftime形式に変換
            fmt = fmt.replace("yyyy", "%Y").replace("mm", "%m").replace("dd", "%d")
            fmt = fmt.replace("HH", "%H").replace("MM", "%M").replace("SS", "%S")
            return now.strftime(fmt)

        res = re.sub(r"{date:(.*?)}", _repl, res)
        return res

    def show_naming_help(self):
        messagebox.showinfo(self._("title_tag_help"), self._("help_tags"))

    def get_final_dest(self, u_info, seq, fseq, pseq, ptotal=1):
        out_name = self.apply_tags(self.naming_var.get(), u_info, seq, fseq, pseq, ptotal)
        base = (
            os.path.dirname(u_info["orig"]["path"]) if self.out_mode_var.get() == "original" else self.out_dir_var.get()
        )
        if not os.path.exists(base):
            os.makedirs(base, exist_ok=True)
        dest = os.path.join(base, re.sub(r'[\\/:*?"<>|]+', "_", out_name) + ".pdf")
        return self.confirm_overwrite_or_rename(dest)

    def confirm_overwrite_or_rename(self, dest):
        if not os.path.exists(dest):
            return dest
        ans = self._ui_call_sync(
            messagebox.askyesnocancel,
            self._("title_overwrite"),
            f"{self._('msg_exists')} {os.path.basename(dest)}\nYes: {self._('btn_overwrite')}, No: {self._('btn_seq')}, Cancel: {self._('btn_abort')}",
        )
        if ans is None:
            return None
        if ans:
            return dest
        b, e = os.path.splitext(dest)
        i = 1
        while os.path.exists(f"{b}_{i}{e}"):
            i += 1
        return f"{b}_{i}{e}"

    def _ui_call_sync(self, func, *args, **kwargs):
        ev = threading.Event()
        slot = {"res": None}
        self.root.after(0, lambda: [slot.update({"res": func(*args, **kwargs)}), ev.set()])
        ev.wait()
        return slot["res"]

    def queue_log(self, msg: str):
        self.progress_queue.put({"log": msg})

    def queue_progress(self, **kwargs):
        self.progress_queue.put(kwargs)

    def check_progress_queue(self):
        try:
            while True:
                m = self.progress_queue.get_nowait()
                if "log" in m:
                    self.log_text.config(state=tk.NORMAL)
                    self.log_text.insert(tk.END, f"[{datetime.datetime.now().strftime('%H:%M:%S')}] {m['log']}\n")
                    self.log_text.see(tk.END)
                    self.log_text.config(state=tk.DISABLED)
                if "progress" in m:
                    self.progress["value"] = m["progress"]
                if "max" in m:
                    self.progress["maximum"] = m["max"]
                if "label" in m:
                    self.progress_label.config(text=m["label"])
        except queue.Empty:
            pass
        self.root.after(100, self.check_progress_queue)

    def finish_action(self, path: str, cfg: AppConfig):
        self.queue_log(self._("msg_all_done"))

        # os.startfile はワーカースレッドでも大抵動くが、気になるなら root.after へ移してもOK
        if cfg.auto_open and path and os.path.exists(path):
            os.startfile(path)
        if cfg.open_folder and path:
            os.startfile(os.path.dirname(path))
        if cfg.clear_after:
            self.root.after(0, self.clear_list)

    # --- File Handlers ---
    def handle_drop(self, event):
        paths = self.root.tk.splitlist(event.data)
        threading.Thread(target=self.add_files_worker, args=(paths,), daemon=True).start()

    def add_files_worker(self, paths):
        for p in paths:
            ext = os.path.splitext(p)[1].lower()  # 拡張子を小文字で取得
            if ext in [".docx", ".doc", ".xlsx", ".xls", ".xlsm", ".pdf", ".pptx", ".ppt", ".jpg", ".png", ".jpeg"]:
                if any(f["path"] == p for f in self.files):
                    continue

                # 文字列検索ではなく、拡張子で正確に判定する
                if ext in [".docx", ".doc"]:
                    t = "Word"
                elif ext in [".xlsx", ".xls", ".xlsm"]:
                    t = "Excel"
                elif ext in [".pptx", ".ppt"]:
                    t = "PowerPoint"
                elif ext in [".pdf"]:
                    t = "PDF"
                else:
                    t = "Image"

                info = {"path": p, "type": t, "range": self._("val_all_pages"), "sheets": []}
                if t == "Excel":
                    info["sheets"] = self.get_excel_sheets(p)
                self.files.append(info)
        self.root.after(0, self.update_tree)

    def get_excel_sheets(self, p):
        try:
            with application("Excel", self.config.engine, self.queue_log) as session:
                ex = session.app
                wb = None
                try:
                    wb = ex.Workbooks.Open(os.path.abspath(p), ReadOnly=True)
                    session.observe_document(wb)
                    return [s.Name for s in wb.Worksheets]
                finally:
                    close_document(wb, self.queue_log)
        except Exception as exc:
            self.queue_log(f"{self._('log_conv_fail')} {p}: {exc}")
            return []

    def update_tree(self):
        self.tree.delete(*self.tree.get_children())
        for f in self.files:
            idx = self.files.index(f) + 1
            u = {"orig": f, "sheet": f["sheets"][0] if f["sheets"] else "", "fseq": idx}
            prev = f"{self.apply_tags(self.naming_var.get(), u, idx, idx, 1)}.pdf"
            self.tree.insert("", tk.END, values=(f["type"], os.path.basename(f["path"]), f["range"], prev))

    def update_output_preview(self):
        if not self.files:
            self.final_name_label.config(text="")
            self.update_tree()
            return
        f0 = self.files[0]
        u0 = {"orig": f0, "sheet": f0["sheets"][0] if f0["sheets"] else "Sheet1", "fseq": 1}
        p_name = self.apply_tags(self.naming_var.get(), u0, 1, 1, 1)
        self.final_name_label.config(text=f"{self._('lbl_output_example')}: {p_name}.pdf")
        self.update_tree()

    def load_preset(self):
        name = self.preset_combo.get()
        if name in self.presets:
            self.config = AppConfig(**self.presets[name])
            self.apply_config_to_ui()
            self.queue_log(f"{self._('log_preset_load')}: {name}")

    def save_preset(self):
        import tkinter.simpledialog as sd

        name = sd.askstring(self._("btn_save"), self._("lbl_preset_name"))
        if name:
            self.update_config_from_ui()
            self.presets[name] = asdict(self.config)
            self.preset_combo["values"] = list(self.presets.keys())
            self.save_config()

    def delete_preset(self):
        name = self.preset_combo.get()
        if name in self.presets and messagebox.askyesno(self._("title_confirm"), self._("msg_ask_delete")):
            del self.presets[name]
            self.preset_combo["values"] = list(self.presets.keys())
            self.save_config()

    def choose_color(self):
        c = colorchooser.askcolor()[1]
        if c:
            self.wm_color_btn.config(bg=c)

    def browse_dir(self):
        d = filedialog.askdirectory()
        if d:
            self.out_dir_var.set(d)

    def move_file(self, d):
        sel = self.tree.selection()
        if not sel:
            return
        idx = self.tree.index(sel[0])
        n = idx + d
        if 0 <= n < len(self.files):
            self.files[idx], self.files[n] = self.files[n], self.files[idx]
            self.update_output_preview()

    def remove_file(self):
        for s in self.tree.selection():
            del self.files[self.tree.index(s)]
        self.update_output_preview()

    def clear_list(self):
        self.files = []
        self.update_output_preview()

    def add_folder(self):
        d = filedialog.askdirectory()
        if d:
            self.add_files_worker([os.path.join(d, f) for f in os.listdir(d)])

    def on_list_double_click(self, e):
        item = self.tree.identify_row(e.y)
        if not item:
            return
        idx = self.tree.index(item)
        f = self.files[idx]
        win = tk.Toplevel(self.root)
        win.title(self._("title_range"))
        win.geometry("300x400")
        if f["type"] == "Excel":
            if not f["sheets"]:
                f["sheets"] = self.get_excel_sheets(f["path"])
            lb = tk.Listbox(win, selectmode=tk.MULTIPLE)
            lb.pack(fill=tk.BOTH, expand=True)
            for s in f["sheets"]:
                lb.insert(tk.END, s)
            tk.Button(
                win,
                text=self._("btn_save"),
                command=lambda: [
                    f.update({"range": ",".join([lb.get(i) for i in lb.curselection()])}),
                    self.update_output_preview(),
                    win.destroy(),
                ],
            ).pack()
        else:
            ent = tk.Entry(win)
            ent.insert(0, f["range"])
            ent.pack(pady=20)
            tk.Button(
                win,
                text=self._("btn_save"),
                command=lambda: [f.update({"range": ent.get()}), self.update_output_preview(), win.destroy()],
            ).pack()

    def cancel_process(self):
        self.cancel_flag.set()

    # --- Language Switch ---
    def change_language(self, lang_code: str):
        if lang_code == self.lang:
            return
        self.lang = lang_code
        self.config.lang = lang_code
        # 設定を即座に保存（UI 値の読み取りに失敗しても言語選択は保持する）
        try:
            self.update_config_from_ui()
            self.config.lang = lang_code
            data = {"current": asdict(self.config), "presets": self.presets}
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=4)
        except Exception:
            pass
        if messagebox.askyesno(self._("title_info"), self._("msg_lang_restart")):
            self.root.destroy()
            restart_app()


def restart_app():
    root = TkinterDnD.Tk()
    PDFUltimateApp(root)
    root.mainloop()


if __name__ == "__main__":
    restart_app()
