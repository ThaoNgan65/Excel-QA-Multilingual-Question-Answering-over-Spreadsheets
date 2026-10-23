#!/usr/bin/env python3
"""
excelqa.py - Hỏi đáp Excel đa ngôn ngữ (query + sheet EN/VI/ZH/JA/KO)

Luồng:
  1) Đọc xlsx → Sheet (merge, header path, multi-table)
  2) Dump → LLM sinh code
  3) Sandbox chạy code
  4) Vote + judge khi lệch
  5) postprocess theo format QUERY (làm tròn đa ngôn ngữ)

Chạy: --xlsx-dir / --queries / --output  hoặc --input / --output
Env: OPENROUTER_API_KEY, OPENROUTER_MODEL, WORKERS, VOTES, JUDGE, ...
"""
# ==================== TỔNG QUAN (chú thích tiếng Việt) ====================
# Chương trình trả lời câu hỏi trên file Excel. Với mỗi câu hỏi, LLM KHÔNG tự tính nhẩm
# mà chỉ viết một đoạn code Python ngắn; đoạn code đó chạy trên dữ liệu thật trong "sandbox",
# biến `answer` chính là đáp án. Luồng xử lý:
#   1) Đọc file .xlsx -> đối tượng Sheet (xử lý ô gộp, tiêu đề nhiều tầng, nhiều bảng, công thức chưa tính).
#   2) Dump (xuất) dữ liệu sheet thành văn bản đưa cho LLM.
#   3) LLM sinh code -> sandbox chạy code -> ra đáp án (làm nhiều lần = nhiều "phiếu").
#   4) Bỏ phiếu đa số; nếu các phiếu lệch nhau thì nhờ LLM "trọng tài" (judge) chọn.
#   5) Hậu xử lý (postprocess): ép định dạng đáp án theo yêu cầu trong câu hỏi
#      (làm tròn, định dạng ngày, giữ đúng nhãn gốc, Yes/No ...) vì chấm điểm theo Exact Match.
#   Đường tắt: câu hỏi chỉ hỏi "công thức trong ô X" thì đọc thẳng từ file, không cần LLM.
# Đầu vào: thư mục xlsx + queries.json. Đầu ra: submission.csv (cột id, answer).
# ===========================================================================
# Phần import bên dưới: các thư viện chuẩn (tham số dòng lệnh, CSV, JSON, đa luồng, Unicode...)
import argparse, csv, datetime as dt, json, math, os, re, statistics, sys, threading, time, traceback, unicodedata
from concurrent.futures import ThreadPoolExecutor
from collections import Counter
# Decimal: làm tròn số thập phân CHÍNH XÁC (tránh sai số của float); các hằng ROUND_* là các kiểu làm tròn.
from decimal import (Decimal, ROUND_HALF_UP, ROUND_HALF_EVEN, ROUND_UP, ROUND_DOWN, ROUND_CEILING,
                     ROUND_FLOOR, InvalidOperation)
# lru_cache: ghi nhớ kết quả của hàm để không phải tính lại với cùng đầu vào.
from functools import lru_cache
# openpyxl: thư viện đọc file Excel (.xlsx); các hàm tiện ích bên dưới để đổi tên cột <-> chỉ số cột, đọc vùng ô.
import openpyxl
from openpyxl.utils import get_column_letter, column_index_from_string, range_boundaries
from openpyxl.utils.cell import coordinate_from_string
# requests: gọi API HTTP (ở đây là OpenRouter để dùng mô hình LLM).
import requests
# Mốc thời gian bắt đầu chương trình - dùng để tính hạn chót (deadline) cho toàn bộ quá trình.
START = time.time()
# ===== CẤU HÌNH: đọc từ biến môi trường; nếu không có thì dùng giá trị mặc định =====
API_URL = os.environ.get("OPENROUTER_API_URL", "https://openrouter.ai/api/v1/chat/completions")  # địa chỉ API chat của OpenRouter
MODEL = os.environ.get("OPENROUTER_MODEL", "qwen/qwen3.7-flash")  # tên mô hình LLM mặc định
WORKERS = int(os.environ.get("WORKERS", "6"))  # số câu hỏi được xử lý song song
TIME_BUDGET = float(os.environ.get("TIME_BUDGET_SEC", "2100"))  # tổng thời gian tối đa (giây), 2100 giây = 35 phút
PRICE_IN = float(os.environ.get("PRICE_IN_PER_M", "0.03"))  # đơn giá token đầu vào (USD / 1 triệu token)
PRICE_OUT = float(os.environ.get("PRICE_OUT_PER_M", "0.13"))  # đơn giá token đầu ra (USD / 1 triệu token)
MAX_COST_PER_Q = float(os.environ.get("MAX_COST_PER_Q", "0.08"))  # trần chi phí (USD) cho MỖI câu hỏi
MAX_TOKENS = int(os.environ.get("LLM_MAX_TOKENS", "3000"))  # số token tối đa LLM được trả về trong một lần gọi
REASONING_OFF = os.environ.get("REASONING", "off").lower() == "off"  # True nếu tắt chế độ 'suy luận' (reasoning) của mô hình
DEBUG_LOG = os.environ.get("DEBUG_LOG", "debug_log.jsonl")  # tên file ghi log chi tiết (định dạng JSON Lines)
VOTES = max(1, int(os.environ.get("VOTES", "3")))  # số phiếu (số lần sinh code độc lập) cho mỗi câu hỏi
VOTE_TEMP = float(os.environ.get("VOTE_TEMP", "0.3"))  # nhiệt độ lấy mẫu cho các phiếu sau phiếu đầu (để đa dạng kết quả)
JUDGE_ON = os.environ.get("JUDGE", "1") != "0"  # bật/tắt trọng tài khi các phiếu lệch nhau
PARALLEL_VOTES = os.environ.get("PARALLEL_VOTES", "1") != "0"  # chạy các phiếu đầu song song để nhanh hơn
MULTI_TABLE = os.environ.get("MULTI_TABLE", "1") != "0"  # bật tính năng tách nhiều bảng xếp chồng trong một sheet
ROUTE_MIN_CHARS = int(os.environ.get("ROUTE_MIN_CHARS", "12000"))  # chỉ chọn sheet (routing) khi bản dump dài hơn số ký tự này
ROUTING_ON = os.environ.get("ROUTING", "1") != "0"  # bật/tắt việc nhờ LLM chọn sheet liên quan
# --- ngưỡng/tham số (mặc định = giá trị cũ; ghi đè bằng biến môi trường để A/B) ---
ROUTE_SHRINK_RATIO = float(os.environ.get("ROUTE_SHRINK_RATIO", "0.8"))  # chỉ dùng routing khi dump giảm dưới tỉ lệ này
JUDGE_TIME_MARGIN = float(os.environ.get("JUDGE_TIME_MARGIN", "15"))     # giây còn lại tối thiểu để gọi judge
JUDGE_COST_RATIO = float(os.environ.get("JUDGE_COST_RATIO", "0.85"))     # chỉ gọi judge khi chi phí < tỉ lệ này của MAX_COST_PER_Q
SANDBOX_TIMEOUT = float(os.environ.get("SANDBOX_TIMEOUT", "15"))  # thời gian tối đa (giây) cho code do LLM sinh ra chạy
MAX_HEADER_ROWS = int(os.environ.get("MAX_HEADER_ROWS", "12"))           # số dòng tối đa được coi là tiêu đề
WIDE_MERGE_MIN_COLS = int(os.environ.get("WIDE_MERGE_MIN_COLS", "4"))    # ô gộp rộng (tiêu đề) tối thiểu số cột
WIDE_MERGE_RATIO = float(os.environ.get("WIDE_MERGE_RATIO", "0.8"))      # ... hoặc tỉ lệ bề rộng bảng
LLM_RETRIES = int(os.environ.get("LLM_RETRIES", "4"))  # số lần thử lại khi gọi API bị lỗi
HTTP_TIMEOUT = float(os.environ.get("HTTP_TIMEOUT", "90"))  # thời gian chờ tối đa (giây) cho mỗi yêu cầu HTTP
VOTE_ATTEMPT_FACTOR = int(os.environ.get("VOTE_ATTEMPT_FACTOR", "2"))    # tối đa VOTES*factor lượt sinh phiếu
# --- cờ hành vi (đặt =0 để tắt) ---
REJECT_ERROR_ANSWERS = os.environ.get("REJECT_ERROR_ANSWERS", "1") != "0"  # chuỗi "Error..." không được tính là đáp án
FULL_DUMP_ON_RETRY = os.environ.get("FULL_DUMP_ON_RETRY", "1") != "0"      # repair/direct dùng dump đầy đủ, bỏ routing
LOG_CANDS = os.environ.get("LOG_CANDS", "1") != "0"                        # ghi mọi phiếu (answer+code) vào debug_log
DUMP_HEADER_CELLS = os.environ.get("DUMP_HEADER_CELLS", "1") != "0"        # dump thêm các dòng tiêu đề dạng ô thô (giữ địa chỉ ô)
CELL_FALLBACK = os.environ.get("CELL_FALLBACK", "1") != "0"                # tra cứu ô theo địa chỉ + đọc công thức thô + xử lý ô trống
JUDGE_OVERRIDE = os.environ.get("JUDGE_OVERRIDE", "0") == "1"

# Khoá (lock) để nhiều luồng không in log chồng chéo lên nhau.
_log_lock = threading.Lock()

# log(*a): in thông điệp ra stderr (an toàn đa luồng nhờ khoá), flush ngay để thấy tức thì.
def log(*a):
    with _log_lock:
        print(*a, file=sys.stderr, flush=True)

# =============================================================================
# 1. TIỆN ÍCH: số, chuỗi, ngày + làm sạch giá trị 
# =============================================================================
# --- markers & number parse (đa dạng dash / fullwidth / kế toán) ---
# Các loại dấu gạch ngang khác nhau: ASCII '-', en-dash, em-dash, dấu trừ Unicode, dạng full-width...
_DASH_CHARS = "-\u2010\u2011\u2012\u2013\u2014\u2015\u2212\ufe58\ufe63\uff0d"
# MARKERS: tập các chuỗi được coi là 'KHÔNG CÓ DỮ LIỆU' (trống, N/A, null, dấu gạch, dấu chấm lửng, ký hiệu chú
# thích thống kê...).
MARKERS = {
    "", "undefine", "undefined", "#", "n/a", "na", "n.a.", "n.a", "null", "none", "nan",
    "-", "--", "---", "?", "ø", "..", "...", "…", "‥", "x", "X", "*", "†", "‡",
    "(d)", "(D)", "s", "S", "u", "U", "nd", "ND", "np", "NP",
}
# Thêm mỗi loại dấu gạch ở dạng 1, 2 và 3 ký tự (-, --, ---) vào MARKERS.
for _d in _DASH_CHARS:
    MARKERS.add(_d)
    MARKERS.add(_d * 2)
    MARKERS.add(_d * 3)
# Các mã lỗi chuẩn của Excel.
ERRORS = {"#N/A", "#VALUE!", "#REF!", "#DIV/0!", "#NAME?", "#NUM!", "#NULL!", "#N/A!", "N/A"}
# Các marker phải GIỮ NGUYÊN (có thể là đáp án thật): nd, np, n/a... cùng các mã lỗi Excel (viết thường).
KEEP_MARKERS = {"nd", "np", "n/a", "na", "n.a", "n.a."} | {e.lower() for e in ERRORS}
# Regex kiểm tra chuỗi 'trông giống số': cho phép ký hiệu tiền tệ và dấu +/- ở đầu, ngoặc kế toán,
# chữ số (cả full-width), dấu phân cách , . và khoảng trắng, đuôi % hoặc dấu - ở cuối.
_NUM_RE = re.compile(
    r"^[\s\$\u20ac\u00a3\u00a5+\-\u2212\u2013\u2014]*"
    r"\(?[\s\$\u20ac\u00a3\u00a5+\-\u2212\u2013\u2014]*"
    r"[\d\uff10-\uff19]"
    r"[\d\uff10-\uff19,\.\s\u00a0\u202f\u2009]*"
    r"\)?\s*%?\s*[\-\u2212]?$"
)

# _to_ascii_digits(s): đổi chữ số full-width (０-９) sang ASCII và các dấu gạch Unicode sang '-'.
# Ví dụ: '１２３' -> '123'.
def _to_ascii_digits(s):
    out = []
    for ch in s:
        o = ord(ch)
        if 0xFF10 <= o <= 0xFF19:  # ký tự là chữ số full-width -> đổi sang '0'-'9'
            out.append(chr(o - 0xFF10 + ord("0")))
        elif ch in "\u2212\u2013\u2014\u2010\u2011":  # ký tự là dấu gạch/dấu trừ Unicode -> đổi thành '-'
            out.append("-")
        else:
            out.append(ch)
    return "".join(out)

# num(x): chuyển một ô thành số thực (float), hoặc None nếu KHÔNG phải số.
# Nguyên tắc quan trọng: dữ liệu thiếu (-, n/a, trống...) trả None, KHÔNG BAO GIỜ coi là 0.
# Ví dụ: '1.234,56' -> 1234.56 | '(1,200)' -> -1200.0 | '12,5%' -> 12.5 | '-' -> None | True -> None
def num(x):
    """Ô → float | None. Marker thiếu → None (không = 0)."""
    # Rỗng hoặc bool (bool là con của int nhưng không phải số thật) -> không phải số.
    if x is None or isinstance(x, bool):
        return None
    # Đã là số: loại NaN/Inf (trả None), còn lại đổi sang float.
    if isinstance(x, (int, float)):
        return None if (isinstance(x, float) and (math.isnan(x) or math.isinf(x))) else float(x)
    # Kiểu khác (ngày, đối tượng khác...) -> không phải số.
    if not isinstance(x, str):
        return None
    # Chuẩn hoá chữ số full-width và các loại khoảng trắng đặc biệt (NBSP...) thành khoảng trắng thường.
    s = _to_ascii_digits(x.strip())
    s = s.replace("\u00a0", " ").replace("\u202f", " ").replace("\u2009", " ")
    # Chuỗi rỗng, marker thiếu dữ liệu hoặc mã lỗi Excel -> None.
    if not s or s.lower() in MARKERS or s.upper() in ERRORS:
        return None
    # Chuỗi chỉ gồm dấu gạch/dấu chấm/khoảng trắng -> None.
    if re.fullmatch(r"[\s\-–—−…\.]+", s):
        return None
    # Không khớp dạng số -> None.
    if not _NUM_RE.match(s):
        return None
    # Xác định số ÂM: ngoặc kế toán (123), dấu '-' ở đầu (sau ký hiệu tiền tệ), hoặc dấu '-' ở cuối.
    neg = (
        (s.startswith("(") and s.rstrip("%").rstrip().endswith(")"))
        or s.lstrip("$€£¥ ").startswith("-")
        or s.rstrip("%").rstrip().endswith("-")
    )
    # Chỉ giữ lại chữ số, dấu phẩy và dấu chấm để xử lý dấu thập phân.
    t = re.sub(r"[^\d,\.]", "", s)
    if not t:
        return None
    # Có cả ',' và '.': dấu nào xuất hiện SAU CÙNG là dấu thập phân, dấu còn lại là phân cách hàng nghìn.
    # (1.234,56 -> 1234.56 ; 1,234.56 -> 1234.56)
    if "," in t and "." in t:
        if t.rfind(",") > t.rfind("."):
            t = t.replace(".", "").replace(",", ".")
        else:
            t = t.replace(",", "")
    # Chỉ có ',': nếu có đúng 2 phần và phần sau KHÔNG đủ 3 chữ số -> dấu thập phân (3,5 -> 3.5);
    # ngược lại là phân cách nghìn (1,234 -> 1234).
    elif "," in t:
        parts = t.split(",")
        if len(parts) == 2 and len(parts[1]) != 3:
            t = t.replace(",", ".")
        else:
            t = t.replace(",", "")
    # Có nhiều dấu '.' -> là phân cách nghìn kiểu châu Âu (1.234.567 -> 1234567).
    elif t.count(".") > 1:
        t = t.replace(".", "")
    # Đổi sang float (lỗi -> None) rồi áp dấu âm nếu có.
    try:
        v = float(t)
    except ValueError:
        return None
    return -v if neg else v

# norm(s): chuẩn hoá chuỗi để SO KHỚP NHÃN (không phân biệt hoa/thường, dấu, ký tự lạ).
# Các bước: NFKC -> xoá chú thích (1) -> bỏ dấu của chữ Latin/tiếng Việt -> đ->d -> hạ chữ thường -> bỏ ký tự lạ.
# Chữ Hán/Nhật/Hàn được giữ nguyên. Ví dụ: 'Đà Nẵng (1)' -> 'da nang'
def norm(s):
    """Chuẩn hoá để so khớp nhãn: bỏ dấu Latin/tiếng Việt, bỏ chú thích (1), bỏ ký tự lạ, hạ chữ thường.
    Giữ nguyên chữ Hán/Nhật/Hàn và các bảng chữ cái khác (trước đây bị xoá sạch)."""
    if s is None:
        return ""
    s = unicodedata.normalize("NFKC", str(s))  # gộp dấu tách rời + ký tự full-width (ＡＢＣ１２ -> ABC12)
    # Xoá chú thích dạng (1), (2), (x) ở trong nhãn.
    s = re.sub(r"\(\s*(\d+|x)\s*\)", " ", s, flags=re.I)
    # Duyệt từng ký tự: chỉ bỏ dấu với chữ Latin (kể cả tiếng Việt); chữ Hán/Nhật/Hàn giữ nguyên.
    out = []
    for ch in s:
        d = unicodedata.normalize("NFD", ch)
        if len(d) > 1 and unicodedata.name(d[0], "").startswith("LATIN"):
            ch = "".join(c for c in d if unicodedata.category(c) != "Mn")  # chỉ bỏ dấu của chữ Latin
        out.append(ch)
    # Đổi đ/Đ thành d/D (vì 'đ' không tách được dấu bằng NFD).
    s = "".join(out).replace("\u0111", "d").replace("\u0110", "D")
    # Hạ chữ thường, thay mọi ký tự không phải chữ/số bằng một khoảng trắng.
    s = re.sub(r"[\W_]+", " ", s.lower())
    return s.strip()


# to_date(x): chuyển giá trị thành datetime, hoặc None nếu không phải ngày.
# Chuỗi dd/mm/yyyy được hiểu theo kiểu NGÀY-TRƯỚC (thử lần lượt các định dạng trong _DATE_FORMATS).
def to_date(x):
    """datetime hoặc None. Chuỗi dd/mm/yyyy được hiểu theo kiểu ngày-trước."""
    if x is None:
        return None
    # datetime giữ nguyên; date đổi thành datetime lúc 00:00; chuỗi thì thử các định dạng; kiểu khác -> None.
    if isinstance(x, dt.datetime):
        return x
    if isinstance(x, dt.date):
        return dt.datetime(x.year, x.month, x.day)
    if not isinstance(x, str):
        return None
    # Gom khoảng trắng thừa rồi thử parse theo từng định dạng ngày.
    s = re.sub(r"\s+", " ", x.strip())

    # nên gọi to_date() với một CHUỖI sẽ gây NameError. Cần khai báo trước hàm, ví dụ:
    # _DATE_FORMATS = ["%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y", "%Y/%m/%d", "%d-%m-%Y"]
    '''for f in _DATE_FORMATS:
        try:
            return dt.datetime.strptime(s, f)
        except ValueError:
            pass'''
    return None

# =============================================================================
# 2. ĐỌC EXCEL → Sheet 
# =============================================================================
# Các từ/tiền tố mà dấu gạch nối phải được GIỮ khi nối lại từ bị ngắt dòng (self-employed, non-profit...).
_KEEP_HYPHEN = {"self", "non", "full", "part", "semi", "anti", "ex", "pre", "post", "multi", "co", "re", "two",
                "three", "one", "year", "short", "long", "low", "high", "middle", "e", "x"}

# clean_cell(v): làm sạch giá trị một ô khi nạp.
# - Chuỗi: bỏ \r, đổi NBSP, nối lại từ bị ngắt dòng bằng dấu gạch, gom khoảng trắng.
# - Marker cần giữ (KEEP_MARKERS) giữ nguyên; marker thường/mã lỗi -> None (coi như ô trống).
# - Giá trị không phải chuỗi (số, ngày...) giữ nguyên.
def clean_cell(v):
    if isinstance(v, str):
        # Bỏ ký tự \r, đổi khoảng trắng cứng (NBSP) thành khoảng trắng thường.
        s = v.replace("\r", "").replace("\u00a0", " ")
        # Nối từ bị ngắt dòng bằng dấu gạch: 'exam-\nple' -> 'example'
        # (giữ lại dấu gạch nếu phần đầu nằm trong _KEEP_HYPHEN, vd 'self-\nemployed' -> 'self-employed').
        s = re.sub(r"([A-Za-z]+)-[ \t]*\n[ \t]*(?=[a-z])",
                   lambda m: m.group(1) + ("-" if m.group(1).lower() in _KEEP_HYPHEN else ""), s)
        # Các xuống dòng còn lại -> một khoảng trắng.
        s = re.sub(r"\s*\n\s*", " ", s)
        # Gom nhiều khoảng trắng/tab liên tiếp thành một và cắt hai đầu.
        s = re.sub(r"[ \t]{2,}", " ", s).strip()
        # Marker cần giữ -> trả nguyên văn.
        if s.lower() in KEEP_MARKERS:
            return s
        # Marker thường hoặc mã lỗi Excel -> coi là ô trống (None).
        if s.lower() in MARKERS or s.upper() in ERRORS:
            return None
        # Chuỗi rỗng -> None, ngược lại trả chuỗi đã làm sạch.
        return s if s != "" else None
    return v

# fmt_val(v): đổi giá trị ô thành văn bản để đưa vào bản dump cho LLM.
# None -> '' | bool -> TRUE/FALSE | datetime -> YYYY-MM-DD (kèm giờ nếu có) | float -> %.12g (tránh
# 0.30000000000000004).
def fmt_val(v):
    if v is None:
        return ""
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, dt.datetime):
        return v.strftime("%Y-%m-%d") if (v.hour, v.minute, v.second) == (0, 0, 0) else v.strftime("%Y-%m-%d %H:%M:%S")
    if isinstance(v, dt.date):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, dt.time):
        return v.strftime("%H:%M:%S")
    if isinstance(v, float):
        return f"{v:.12g}"
    return str(v)

# _is_year(v): True nếu v là số nguyên trong khoảng 1900-2100 (coi là 'năm', không phải dữ liệu số).
def _is_year(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool) and float(v).is_integer() and 1900 <= v <= 2100

# _data_like(row): hàng này là hàng DỮ LIỆU hay hàng TIÊU ĐỀ?
# - Đếm số ô số/ngày (không tính 'năm') và số ô chữ.
# - Không có số nào mà có >= 5 ô chữ -> là tiêu đề (False).
# - Còn lại: là dữ liệu nếu có ít nhất 2 ô số/ngày.
def _data_like(row):
    # n: số ô dạng dữ liệu | nnum: số ô là số | ntxt: số ô là chữ.
    n = nnum = ntxt = 0
    for v in row:
        if v is None:
            continue
        if isinstance(v, (dt.datetime, dt.date)):
            n += 1
        elif isinstance(v, (int, float)) and not isinstance(v, bool):
            if not _is_year(v):
                n += 1
                nnum += 1
        elif isinstance(v, str):
            if num(v) is not None and not _is_year(num(v)):
                n += 1
                nnum += 1
            elif num(v) is None:
                ntxt += 1
    if nnum == 0 and ntxt >= 5:  # nhiều nhãn chữ, không có số nào → dòng tiêu đề (kể cả khi các cột được đặt tên bằng ngày)
        return False
    return n >= 2  # dòng dữ liệu nếu có ít nhất 2 ô số/ngày

# _mini_eval(...): bộ tính công thức ĐƠN GIẢN (SUM, tham chiếu ô/vùng) khi ô chưa có giá trị tính sẵn.
# Cách làm: biến công thức Excel thành biểu thức Python rồi eval với môi trường bị giới hạn.
# Đệ quy tối đa 25 tầng. Trả None nếu có ô phụ thuộc không tính được.
def _mini_eval(ws_f, grid, coord, depth=0):
    """Đánh giá công thức đơn giản khi ô chưa có giá trị tính sẵn (SUM, tham chiếu ô/vùng)."""
    try:
        cell = ws_f[coord]
    except Exception:
        return None
    f = cell.value
    if not (isinstance(f, str) and f.startswith("=")) or depth > 25:
        return None
    # Danh sách các ô phụ thuộc là công thức nhưng chưa tính được (nếu còn thì kết quả không đáng tin -> None).
    unresolved = []
    # ref(c): lấy giá trị số của một ô; nếu ô trống mà là công thức thì tính đệ quy; ô không có -> 0.
    def ref(c):
        c = c.replace("$", "").upper()
        try:
            col, row = coordinate_from_string(c)
            r, ci = row - 1, column_index_from_string(col) - 1
            if 0 <= r < len(grid) and 0 <= ci < len(grid[r]):
                v = grid[r][ci]
                if isinstance(v, (int, float)) and not isinstance(v, bool):
                    return v
                if v is None:
                    try:
                        is_f = isinstance(ws_f[c].value, str) and ws_f[c].value.startswith("=")
                    except Exception:
                        is_f = False
                    sub = _mini_eval(ws_f, grid, c, depth + 1)
                    if sub is None and is_f:
                        unresolved.append(c)
                    return sub or 0
            return 0
        except Exception:
            return 0
    # nums(vals): lọc ra các giá trị là số (bỏ bool và chữ).
    def nums(vals):
        return [x for x in vals if isinstance(x, (int, float)) and not isinstance(x, bool)]
    # rng(a): mở một vùng như 'A1:B3' thành danh sách giá trị các ô trong vùng.
    def rng(a):
        try:
            c1, r1, c2, r2 = range_boundaries(a.replace("$", ""))
            return [ref(f"{get_column_letter(c)}{r}") for r in range(r1, r2 + 1) for c in range(c1, c2 + 1)]
        except Exception:
            return []
    # SUM(*args): cộng các số; đối số có thể là một danh sách (vùng) hoặc một số đơn lẻ.
    def SUM(*args):
        tot = 0.0
        for a in args:
            if isinstance(a, list):
                tot += sum(nums(a))
            elif isinstance(a, (int, float)):
                tot += a
        return tot
    # Bỏ dấu '=' ở đầu công thức.
    expr = f[1:]
    # Biến vùng A1:B3 thành lời gọi rng('A1:B3').
    expr = re.sub(r"\$?([A-Z]{1,3})\$?(\d+):\$?([A-Z]{1,3})\$?(\d+)",
                  lambda m: f"rng('{m.group(0)}')", expr, flags=re.I)
    # Biến ô đơn A1 thành lời gọi ref('A1').
    expr = re.sub(r"(?<![A-Za-z'\"])\$?([A-Z]{1,3})\$?(\d+)(?![\d('])",
                  lambda m: f"ref('{m.group(0)}')", expr, flags=re.I)
    # Tính biểu thức với builtins RỖNG (chỉ có ref, rng, SUM) -> an toàn; lỗi -> None; còn ô chưa giải quyết ->
    # None.
    try:
        res = eval(expr, {"__builtins__": {}}, {"ref": ref, "rng": rng, "SUM": SUM, "sum": SUM})
    except Exception:
        return None
    return None if unresolved else res

# _eval_sum(formula, grid): phương án dự phòng cuối cùng, CHỈ xử lý công thức dạng =SUM(...).
# Không tìm thấy ô số nào thì trả None (không bịa ra số 0).
def _eval_sum(formula, grid):
    """Fallback: chỉ xử lý =SUM(...) khi mini-eval không dùng được."""
    # Chỉ nhận đúng dạng =SUM(...), dạng khác -> None.
    m = re.fullmatch(r"=\s*SUM\((.*)\)\s*", formula, flags=re.I)
    if not m:
        return None
    # total: tổng cộng dồn | found: đã gặp ít nhất một ô số hay chưa.
    total, found = 0.0, False
    # Mỗi đối số của SUM là một ô (B2) hoặc một vùng (B2:B9).
    for part in m.group(1).split(","):
        part = part.strip().replace("$", "")
        mm = re.fullmatch(r"([A-Za-z]+)(\d+)(?::([A-Za-z]+)(\d+))?", part)
        if not mm:
            return None
        c1, r1 = column_index_from_string(mm.group(1).upper()) - 1, int(mm.group(2)) - 1
        c2, r2 = (column_index_from_string(mm.group(3).upper()) - 1, int(mm.group(4)) - 1) if mm.group(3) else (c1, r1)
        for r in range(min(r1, r2), max(r1, r2) + 1):
            for c in range(min(c1, c2), max(c1, c2) + 1):
                if r < len(grid) and c < len(grid[r]):
                    v = grid[r][c]
                    if isinstance(v, (int, float)) and not isinstance(v, bool):
                        total += v
                        found = True
    return total if found else None  # không có ô số nào: không bịa số 0

# =============================================================================
# 2b. BỘ TÍNH CÔNG THỨC (xuyên sheet + cache liên kết ngoài) cho ô công thức không có giá trị tính sẵn
# =============================================================================
import operator
from collections import namedtuple
from openpyxl.formula.tokenizer import Tokenizer, Token

# XlErr: lỗi Excel THẬT (#N/A, #DIV/0!...). Hàm IFERROR/IFNA bắt được loại lỗi này.
class XlErr(Exception):
    """Lỗi Excel thật (#N/A, #DIV/0!...): IFERROR bắt được."""
    def __init__(self, code="#VALUE!"):
        super().__init__(code)
        self.code = code

# XlUnsupported: giới hạn của bộ tính (hàm lạ, liên kết ngoài không có cache, vòng tham chiếu...).
# IFERROR KHÔNG được nuốt lỗi này (để không trả đáp án sai một cách im lặng).
class XlUnsupported(Exception):
    """Giới hạn của bộ tính (hàm lạ, liên kết ngoài không cache...): IFERROR KHÔNG được nuốt."""
# XFAIL: giá trị đặc biệt đánh dấu 'không tính được'.
XFAIL = object()
# XRange: mô tả một vùng ô = (tên sheet, hàng đầu, cột đầu, hàng cuối, cột cuối).
XRange = namedtuple("XRange", "sheet r1 c1 r2 c2")
# Tập các mã lỗi Excel hợp lệ.
_ERR_CODES = {"#N/A", "#VALUE!", "#REF!", "#DIV/0!", "#NAME?", "#NUM!", "#NULL!"}
# Độ ưu tiên toán tử (số càng lớn càng ưu tiên): so sánh < nối chuỗi & < cộng/trừ < nhân/chia < lũy thừa.
_PREC = {"=": 1, "<>": 1, "<": 1, ">": 1, "<=": 1, ">=": 1, "&": 2, "+": 3, "-": 3, "*": 4, "/": 4, "^": 5}
# Ánh xạ toán tử so sánh của Excel -> hàm so sánh của Python.
_CMP = {"=": operator.eq, "<>": operator.ne, "<": operator.lt, ">": operator.gt, "<=": operator.le, ">=": operator.ge}
# Regex nhận dạng tham chiếu ô/vùng, có thể kèm tên sheet (có hoặc không có nháy đơn), cột trọn (A:A), hàng trọn
# (1:1).
_REF_RE = re.compile(r"^(?:(?:'((?:[^']|'')+)'|([^'!]+))!)?(\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?"
                     r"|\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}|\$?\d+:\$?\d+)$")
# Nhận dạng tên sheet thuộc liên kết ngoài, dạng [1]TênSheet.
_EXT_RE = re.compile(r"^\[(\d+)\](.+)$")
# Các hàm chuỗi được hỗ trợ: LEN, TRIM, UPPER, LOWER, LEFT, RIGHT, MID.
_STR_FN = {"LEN": lambda s: float(len(s)), "TRIM": lambda s: re.sub(r" +", " ", s).strip(), "UPPER": lambda s: s.upper(),
           "LOWER": lambda s: s.lower(), "LEFT": lambda s, n=1: s[:int(n)],
           "RIGHT": lambda s, n=1: s[-int(n):] if int(n) > 0 else "",
           "MID": lambda s, a, n: s[max(int(a) - 1, 0): max(int(a) - 1, 0) + int(n)]}

# _parse(formula): phân tích công thức thành CÂY CÚ PHÁP (AST) dạng tuple lồng nhau.
# Các loại nút: ("val", giá trị) | ("ref", văn bản tham chiếu) | ("fn", tên hàm, [đối số]) |
#               ("un", dấu, x) | ("bin", toán tử, trái, phải) | ("pct", x) | ("err", mã lỗi).
# Dùng bộ Tokenizer của openpyxl rồi phân tích đệ quy xuống (recursive descent).
def _parse(formula):
    # Tách công thức thành các token, bỏ token khoảng trắng.
    t = [x for x in Tokenizer(formula).items if x.type != Token.WSPACE]
    # Con trỏ vị trí token hiện tại (dùng list để hàm lồng bên dưới sửa được giá trị).
    i = [0]
    # peek(): xem token hiện tại (không di chuyển con trỏ).
    def peek():
        return t[i[0]] if i[0] < len(t) else None
    # take(): lấy token hiện tại và đẩy con trỏ lên một bước.
    def take():
        x = peek()
        i[0] += 1
        return x
    # expr(minp): xử lý toán tử hai ngôi theo độ ưu tiên (precedence climbing). Ví dụ 1+2*3 ->
    # ('bin','+',1,('bin','*',2,3)).
    def expr(minp=1):
        left = unary()
        while True:
            p = peek()
            if p is None or p.type != Token.OP_IN or _PREC.get(p.value, 0) < minp:
                return left
            take()
            left = ("bin", p.value, left, expr(_PREC[p.value] + 1))
    # unary(): xử lý dấu đứng trước (-x, +x) và dấu % đứng sau (5%).
    def unary():
        p = peek()
        if p is not None and p.type == Token.OP_PRE:
            take()
            return ("un", p.value, unary())
        n = primary()
        while peek() is not None and peek().type == Token.OP_POST:
            take()
            n = ("pct", n)
        return n
    # primary(): xử lý thành phần cơ bản: số, chuỗi, logic, lỗi, tham chiếu, lời gọi hàm (đọc danh sách đối số),
    # biểu thức trong ngoặc.
    def primary():
        p = take()
        if p is None:
            raise XlUnsupported("unexpected end")
        if p.type == Token.OPERAND:
            if p.subtype == Token.NUMBER:
                return ("val", float(p.value))
            if p.subtype == Token.TEXT:
                return ("val", p.value[1:-1].replace('""', '"'))
            if p.subtype == Token.LOGICAL:
                return ("val", p.value.upper() == "TRUE")
            return ("err", p.value) if p.subtype == Token.ERROR else ("ref", p.value)
        if p.type == Token.FUNC and p.subtype == Token.OPEN:
            args = []
            if peek() is not None and peek().type == Token.FUNC and peek().subtype == Token.CLOSE:
                take()
                return ("fn", p.value[:-1].upper().replace("_XLFN.", ""), args)
            while True:
                q = peek()
                args.append(("val", None) if q is not None and (q.type == Token.SEP or q.subtype == Token.CLOSE) else expr())
                q = take()
                if q is None:
                    raise XlUnsupported("unclosed function")
                if q.type == Token.FUNC and q.subtype == Token.CLOSE:
                    return ("fn", p.value[:-1].upper().replace("_XLFN.", ""), args)
                if q.type != Token.SEP:
                    raise XlUnsupported("bad arguments")
        if p.type == Token.PAREN and p.subtype == Token.OPEN:
            n = expr()
            q = take()
            if q is None or q.type != Token.PAREN:
                raise XlUnsupported("unclosed paren")
            return n
        raise XlUnsupported(f"token {p.value!r}")
    # Phân tích toàn bộ công thức; nếu còn token thừa thì coi là không hỗ trợ.
    n = expr()
    if peek() is not None:
        raise XlUnsupported("trailing tokens")
    return n

# _serial(d): đổi ngày sang SỐ SERIAL của Excel (số ngày tính từ 1899-12-30).
def _serial(d):
    d = d if isinstance(d, dt.datetime) else dt.datetime(d.year, d.month, d.day)
    return (d - dt.datetime(1899, 12, 30)).total_seconds() / 86400.0

# _xl_str(v): đổi giá trị sang chuỗi theo kiểu Excel (3.0 -> '3', bool -> TRUE/FALSE, ngày -> YYYY-MM-DD).
def _xl_str(v):
    if v is None:
        return ""
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, float):
        return str(int(v)) if v == int(v) and abs(v) < 1e15 else f"{v:.15g}"
    return fmt_val(v) if isinstance(v, (dt.datetime, dt.date)) else str(v)

# _cmp(op, a, b): so sánh hai giá trị theo luật Excel (chuỗi không phân biệt hoa/thường; thứ tự: số < chuỗi <
# bool).
def _cmp(op, a, b):
    # norm_: ô trống được ép theo kiểu của vế bên kia (so với chuỗi -> '', so với số -> 0); ngày đổi sang số
    # serial.
    def norm_(v, other):
        if v is None:
            return "" if isinstance(other, str) else (False if isinstance(other, bool) else 0.0)
        return _serial(v) if isinstance(v, (dt.datetime, dt.date)) else v
    a, b = norm_(a, b), norm_(b, a)
    key = lambda v: (1, v.lower()) if isinstance(v, str) else ((2, v) if isinstance(v, bool) else (0, v))  # noqa: E731
    return _CMP[op](key(a), key(b))

# _crit(c): tạo hàm lọc cho SUMIF/COUNTIF từ tiêu chí. Hỗ trợ >,<,>=,<=,<>,= và ký tự đại diện * ?
# Ví dụ: '>=5' lọc số >= 5; 'app*' khớp mọi chuỗi bắt đầu bằng 'app'.
def _crit(c):
    """Tiêu chí SUMIF/COUNTIF: số hoặc chuỗi (hỗ trợ * ?), kèm toán tử >,<,>=,<=,<>,=."""
    op, rest = "=", c
    m = re.match(r"^(<=|>=|<>|<|>|=)(.*)$", c) if isinstance(c, str) else None
    if m:
        op, rest = m.groups()
    try:
        rn = None if isinstance(rest, bool) else float(rest)
    except (TypeError, ValueError):
        rn = None
    pat = str(rest).lower()
    rx = re.compile("^" + re.escape(pat).replace(r"\*", ".*").replace(r"\?", ".") + "$", re.S)
    def pred(v):
        if rn is not None:
            try:
                return _CMP[op](float(v), rn)
            except (TypeError, ValueError):
                return op == "<>"
        ok = bool(rx.match(_xl_str(v).lower()))
        return (not ok) if op == "<>" else (ok if op == "=" else False)
    return pred

# FormulaEngine: bộ TÍNH CÔNG THỨC xuyên sheet (cả liên kết ngoài có cache) cho các ô công thức chưa có giá trị
# tính sẵn.
# File Excel do thư viện khác ghi ra thường không lưu giá trị đã tính, nên cần tự tính lại.
# Hỗ trợ: SUM, AVERAGE, MIN, MAX, COUNT(A), IF, IFERROR, IFNA, AND, OR, NOT, IS*, ROUND*, ABS, INT, CONCAT,
# LEN/TRIM/UPPER/LOWER/LEFT/RIGHT/MID, SUMIF(S), COUNTIF(S), VLOOKUP, MATCH, INDEX.
class FormulaEngine:
    # Khởi tạo: wb_v = workbook giá trị tính sẵn, wb_f = workbook công thức; chuẩn bị các bộ nhớ đệm (vals, fmls,
    # memo, busy, ast).
    def __init__(self, wb_v, wb_f):
        self.wb_v, self.wb_f = wb_v, wb_f
        self.names = {n.lower(): n for n in wb_v.sheetnames}
        self.vals, self.fmls, self.memo, self.busy, self.ast = {}, {}, {}, set(), {}
        self.ext = self._load_ext(wb_f)
    # _load_ext: đọc GIÁ TRỊ CACHE của liên kết ngoài ([1], [2]...) lưu trong metadata của workbook.
    # Trả về danh sách theo thứ tự liên kết; mỗi phần tử là {tên sheet thường: ({(hàng,cột): giá trị}, số hàng tối
    # đa, số cột tối đa)}.
    @staticmethod
    def _load_ext(wb):
        """Cache liên kết ngoài: [ {tên sheet thường: ({(r,c): giá trị}, maxr, maxc)} ... ] theo thứ tự [1], [2]..."""
        out = []
        for link in (getattr(wb, "_external_links", None) or []):
            sheets = {}
            try:
                book = link.externalBook
                names = list(book.sheetNames.sheetName)
                for sd in (book.sheetDataSet.sheetData if book.sheetDataSet is not None else []):
                    cells = {}
                    for row in sd.row:
                        for c in row.cell:
                            if c.v is None:
                                continue
                            try:
                                v = c.v if c.t == "str" else (c.v in ("1", "true", "TRUE") if c.t == "b" else
                                                              XlErr(c.v) if c.t == "e" else float(c.v))
                                col, r = coordinate_from_string(c.r)
                                cells[(r, column_index_from_string(col))] = v
                            except Exception:
                                continue
                    mr = max((k[0] for k in cells), default=0)
                    mc = max((k[1] for k in cells), default=0)
                    sheets[names[int(sd.sheetId)].lower()] = (cells, mr, mc)
            except Exception:
                pass
            out.append(sheets)
        return out
    # _g: nạp một sheet thành lưới 2 chiều (list các list) và lưu cache trong `store`.
    def _g(self, store, wb, sheet):
        if sheet not in store:
            try:
                store[sheet] = [list(r) for r in wb[sheet].iter_rows(values_only=True)]
            except Exception:
                store[sheet] = []
        return store[sheet]
    # _ext_sheet: lấy dữ liệu cache của một sheet ngoài; không có thì ném XlUnsupported.
    def _ext_sheet(self, key):
        try:
            return self.ext[key[1] - 1][key[2]]
        except (IndexError, KeyError):
            raise XlUnsupported("external sheet not cached")
    # _dims: trả (số hàng, số cột) của sheet (trong workbook hoặc liên kết ngoài).
    def _dims(self, sheet):
        if isinstance(sheet, tuple):
            return self._ext_sheet(sheet)[1:]
        g = self._g(self.vals, self.wb_v, sheet)
        return len(g), max((len(r) for r in g), default=0)
    # _cell: lấy giá trị một ô. Ưu tiên giá trị tính sẵn; nếu trống mà có công thức thì TÍNH ĐỆ QUY (có ghi nhớ
    # kết quả - memo).
    def _cell(self, sheet, r, c):
        """Giá trị ô: giá trị tính sẵn nếu có, không thì tính từ công thức (đệ quy + memo)."""
        # sheet là tuple -> thuộc liên kết ngoài, lấy từ cache liên kết ngoài.
        if isinstance(sheet, tuple):
            return self._ext_sheet(sheet)[0].get((r, c))
        key = (sheet, r, c)
        # Đã tính rồi -> dùng lại kết quả trong memo (nếu là XFAIL thì báo không hỗ trợ).
        if key in self.memo:
            if self.memo[key] is XFAIL:
                raise XlUnsupported("unresolved")
            return self.memo[key]
        # at: lấy phần tử (r, c) của lưới; nằm ngoài vùng -> None.
        at = lambda g: g[r - 1][c - 1] if 0 < r <= len(g) and 0 < c <= len(g[r - 1]) else None  # noqa: E731
        # Ưu tiên giá trị đã tính sẵn trong file.
        v = at(self._g(self.vals, self.wb_v, sheet))
        # Ô không có giá trị: nếu có công thức thì tự tính.
        if v is None:
            f = at(self._g(self.fmls, self.wb_f, sheet))
            f = getattr(f, "text", f)
            if isinstance(f, str) and len(f) > 1 and f.startswith("="):
                # Ô này đang được tính dở (đang trong chuỗi đệ quy) -> vòng tham chiếu.
                if key in self.busy:
                    raise XlUnsupported("circular")
                self.busy.add(key)
                # Tính công thức: phân tích (AST có cache) -> duyệt cây -> ép về giá trị đơn.
                try:
                    v = self._sc(self._ev(self.ast.setdefault(f, _parse(f)), key), key)
                except XlErr as e:
                    v = XlErr(e.code)
                except XlUnsupported:
                    self.memo[key] = XFAIL
                    raise
                finally:
                    self.busy.discard(key)
        # Ô chứa chuỗi là mã lỗi Excel -> đổi thành đối tượng XlErr.
        elif isinstance(v, str) and v.strip().upper() in _ERR_CODES:
            v = XlErr(v.strip().upper())
        # Lưu kết quả vào bộ nhớ đệm để lần sau không phải tính lại.
        self.memo[key] = v
        return v
    # value: giao diện công khai cho lớp Sheet. Trả XFAIL nếu không tính được; mã lỗi (vd '#DIV/0!') nếu là lỗi
    # Excel thật; None nếu rỗng.
    def value(self, sheet, r, c):
        """Giá trị ô đã tính; XFAIL nếu chưa hỗ trợ; None nếu rỗng/lỗi Excel thật."""
        try:
            v = self._cell(sheet, r, c)
        except (XlUnsupported, RecursionError, XlErr):
            return XFAIL
        if isinstance(v, XlErr):
            return v.code  # lỗi Excel thật (#DIV/0!...) là giá trị hiển thị của ô
        return None if v == "" else (XFAIL if isinstance(v, XRange) else v)
    # _ref: phân giải một tham chiếu thành XRange. Xử lý tên sheet, liên kết ngoài [1]Sheet!A1, cột trọn A:A, hàng
    # trọn 1:1, ô/vùng thường.
    def _ref(self, text, ctx):
        m = _REF_RE.match(text)
        if not m:
            raise XlUnsupported(f"name/ref {text!r}")
        name = m.group(1) if m.group(1) is not None else m.group(2)
        if name is None:
            sheet = ctx[0]
        else:
            name = name.replace("''", "'")
            e = _EXT_RE.match(name)
            sheet = ("ext", int(e.group(1)), e.group(2).lower()) if e else self.names.get(name.lower())
            if sheet is None:
                raise XlUnsupported("3D reference") if ":" in name else XlErr("#REF!")
        a = m.group(3).replace("$", "").upper()
        nr, nc = self._dims(sheet)
        if re.fullmatch(r"[A-Z]{1,3}:[A-Z]{1,3}", a):
            x, y = a.split(":")
            return XRange(sheet, 1, column_index_from_string(x), nr or 1, column_index_from_string(y))
        if re.fullmatch(r"\d+:\d+", a):
            x, y = a.split(":")
            return XRange(sheet, int(x), 1, int(y), nc or 1)
        c1, r1, c2, r2 = range_boundaries(a)
        return XRange(sheet, r1, c1, r2, c2)
    # _sc: ép về GIÁ TRỊ ĐƠN. Vùng một hàng/một cột được lấy theo 'giao điểm ngầm định' với ô đang tính; lỗi thì
    # ném ra.
    def _sc(self, v, ctx):
        """Về giá trị đơn: giao điểm ngầm định của vùng; lỗi thì ném."""
        if isinstance(v, XRange):
            if v.r1 == v.r2 and v.c1 <= (ctx[2] if v.c1 != v.c2 else v.c1) <= v.c2:
                v = self._cell(v.sheet, v.r1, ctx[2] if v.c1 != v.c2 else v.c1)
            elif v.c1 == v.c2 and v.r1 <= ctx[1] <= v.r2:
                v = self._cell(v.sheet, ctx[1], v.c1)
            else:
                raise XlErr("#VALUE!")
        if isinstance(v, XlErr):
            raise XlErr(v.code)
        return v
    # _cells: sinh ra (hàng, cột, giá trị) cho mọi ô trong vùng (cắt theo kích thước thật của sheet).
    def _cells(self, rg):
        mr, mc = self._dims(rg.sheet)
        for r in range(rg.r1, min(rg.r2, mr) + 1):
            for c in range(rg.c1, min(rg.c2, mc) + 1):
                yield r, c, self._cell(rg.sheet, r, c)
    # _rng: tính một nút AST và BẮT BUỘC kết quả là một vùng, nếu không thì lỗi #VALUE!.
    def _rng(self, node, ctx):
        v = self._ev(node, ctx)
        if not isinstance(v, XRange):
            raise XlErr("#VALUE!")
        return v
    # _n(v): ép giá trị về số theo luật Excel (None/'' -> 0, bool/số -> float, ngày -> serial, chuỗi số -> float,
    # còn lại -> #VALUE!).
    @staticmethod
    def _n(v):
        if v is None or v == "":
            return 0.0
        if isinstance(v, (bool, int, float)):
            return float(v)
        if isinstance(v, (dt.datetime, dt.date)):
            return _serial(v)
        try:
            return float(str(v).strip().replace(",", ""))
        except ValueError:
            raise XlErr("#VALUE!")
    # _flat: làm phẳng một vùng thành danh sách giá trị (hoặc [giá trị đơn]).
    def _flat(self, v, ctx):
        return [x for _, _, x in self._cells(v)] if isinstance(v, XRange) else [self._sc(v, ctx)]
    # _nums: gom các SỐ từ danh sách đối số cho SUM/AVERAGE/...: ô trong vùng chỉ lấy số (bỏ chữ, bool), gặp lỗi
    # thì ném.
    def _nums(self, args, ctx):
        out = []
        for a in args:
            v = self._ev(a, ctx)
            if isinstance(v, XRange):
                for x in self._flat(v, ctx):
                    if isinstance(x, XlErr):
                        raise XlErr(x.code)
                    if isinstance(x, (int, float)) and not isinstance(x, bool):
                        out.append(float(x))
            elif self._sc(v, ctx) is not None:
                out.append(self._n(self._sc(v, ctx)))
        return out
    # _ev: DUYỆT CÂY CÚ PHÁP và tính giá trị: hằng, lỗi, tham chiếu, hàm, dấu một ngôi, %, toán tử hai ngôi (+ - *
    # / ^ & và so sánh).
    def _ev(self, n, ctx):
        k = n[0]
        if k == "val":
            return n[1]
        if k == "err":
            raise XlErr(n[1])
        if k == "ref":
            return self._ref(n[1], ctx)
        if k == "fn":
            return self._fn(n[1], n[2], ctx)
        if k == "un":
            v = self._sc(self._ev(n[2], ctx), ctx)
            return v if n[1] == "+" else -self._n(v)  # '+' đơn ngôi giữ nguyên giá trị (cả chữ/ngày)
        if k == "pct":
            return self._n(self._sc(self._ev(n[1], ctx), ctx)) / 100.0
        if k == "bin":
            op, a, b = n[1], self._sc(self._ev(n[2], ctx), ctx), self._sc(self._ev(n[3], ctx), ctx)
            if op == "&":
                return _xl_str(a) + _xl_str(b)
            if op in _CMP:
                return _cmp(op, a, b)
            x, y = self._n(a), self._n(b)
            if op == "/" and y == 0:
                raise XlErr("#DIV/0!")
            try:
                r = {"+": x + y, "-": x - y, "*": x * y}[op] if op in "+-*" else (x / y if op == "/" else x ** y)
            except (ZeroDivisionError, OverflowError):
                raise XlErr("#NUM!")
            if isinstance(r, complex):
                raise XlErr("#NUM!")
            return r
        raise XlUnsupported(k)
    # _fn: THỰC THI một hàm Excel. Hàm chưa hỗ trợ -> ném XlUnsupported.
    def _fn(self, name, args, ctx):
        # ev(i): tính đối số thứ i và ép về giá trị đơn.
        ev = lambda i: self._sc(self._ev(args[i], ctx), ctx)  # noqa: E731
        # Nhóm hàm thống kê trên các số: SUM, AVERAGE (rỗng -> #DIV/0!), MIN, MAX, COUNT.
        if name in ("SUM", "AVERAGE", "MIN", "MAX", "COUNT"):
            v = self._nums(args, ctx)
            if name == "AVERAGE" and not v:
                raise XlErr("#DIV/0!")
            return {"SUM": sum, "COUNT": lambda x: float(len(x)), "AVERAGE": lambda x: sum(x) / len(x),
                    "MIN": lambda x: min(x, default=0.0), "MAX": lambda x: max(x, default=0.0)}[name](v)
        # COUNTA: đếm các ô KHÔNG trống.
        if name == "COUNTA":
            return float(sum(1 for a in args for x in self._flat(self._ev(a, ctx), ctx) if x is not None and x != ""))
        # IF(điều kiện, đúng, sai): điều kiện là chuỗi -> #VALUE!; thiếu nhánh sai -> FALSE.
        if name == "IF":
            c = ev(0)
            if isinstance(c, str):
                raise XlErr("#VALUE!")
            i = 1 if self._n(c) else 2
            return self._ev(args[i], ctx) if len(args) > i else bool(i == 1)
        # IFERROR/IFNA: nếu biểu thức gặp lỗi Excel THẬT thì trả giá trị dự phòng (IFNA chỉ bắt #N/A).
        if name in ("IFERROR", "IFNA"):
            try:
                return ev(0)
            except XlErr as e:
                if name == "IFNA" and e.code != "#N/A":
                    raise
                return self._ev(args[1], ctx)
        # AND/OR: gom giá trị logic từ các đối số (bỏ ô trống và chuỗi); không có giá trị nào -> #VALUE!.
        if name in ("AND", "OR"):
            vals = []
            for a in args:
                for x in self._flat(self._ev(a, ctx), ctx):
                    if isinstance(x, XlErr):
                        raise XlErr(x.code)
                    if x is not None and not isinstance(x, str):
                        vals.append(bool(self._n(x)))
            if not vals:
                raise XlErr("#VALUE!")
            return all(vals) if name == "AND" else any(vals)
        # NOT: phủ định logic.
        if name == "NOT":
            return not self._n(ev(0))
        # Nhóm hàm kiểm tra kiểu: ISNUMBER, ISTEXT, ISBLANK, ISERROR, ISNA.
        if name in ("ISNUMBER", "ISTEXT", "ISBLANK", "ISERROR", "ISNA"):
            try:
                v = ev(0)
            except XlErr as e:
                return name == "ISERROR" or (name == "ISNA" and e.code == "#N/A")
            return {"ISNUMBER": isinstance(v, (int, float, dt.datetime, dt.date)) and not isinstance(v, bool),
                    "ISTEXT": isinstance(v, str) and v != "", "ISBLANK": v is None}.get(name, False)
        # ROUND/ROUNDUP/ROUNDDOWN: làm tròn bằng Decimal; ROUND là half-up (giống Excel), khác round() của Python
        # (làm tròn chẵn).
        if name in ("ROUND", "ROUNDUP", "ROUNDDOWN"):
            d = int(self._n(ev(1))) if len(args) > 1 else 0
            mode = {"ROUND": ROUND_HALF_UP, "ROUNDUP": ROUND_UP, "ROUNDDOWN": ROUND_DOWN}[name]
            try:
                return float(Decimal(repr(self._n(ev(0)))).quantize(Decimal(1).scaleb(-d), rounding=mode))
            except (InvalidOperation, ValueError):
                raise XlErr("#NUM!")
        # ABS: giá trị tuyệt đối.
        if name == "ABS":
            return abs(self._n(ev(0)))
        # INT: làm tròn XUỐNG số nguyên gần nhất (floor).
        if name == "INT":
            return float(math.floor(self._n(ev(0))))
        # CONCAT/CONCATENATE: nối chuỗi các đối số (bỏ qua giá trị lỗi).
        if name in ("CONCATENATE", "CONCAT"):
            return "".join(_xl_str(x) for a in args for x in self._flat(self._ev(a, ctx), ctx) if not isinstance(x, XlErr))
        # Các hàm chuỗi: LEN, TRIM, UPPER, LOWER, LEFT, RIGHT, MID.
        if name in _STR_FN:
            return _STR_FN[name](_xl_str(ev(0)), *[self._n(ev(i)) for i in range(1, len(args))])
        # SUMIF/COUNTIF/SUMIFS/COUNTIFS: tính tổng/đếm có điều kiện (xử lý trong _cond).
        if name in ("SUMIF", "COUNTIF", "SUMIFS", "COUNTIFS"):
            return self._cond(name, args, ctx)
        # VLOOKUP(giá trị, vùng, cột, [khớp đúng?]): tra cứu theo cột đầu của vùng.
        if name == "VLOOKUP":
            look, rg, idx = ev(0), self._rng(args[1], ctx), int(self._n(ev(2)))
            if not 1 <= idx <= rg.c2 - rg.c1 + 1:
                raise XlErr("#REF!")
            keys = [self._cell(rg.sheet, r, rg.c1) for r in range(rg.r1, min(rg.r2, self._dims(rg.sheet)[0]) + 1)]
            exact = len(args) > 3 and not self._n(ev(3))
            v = self._cell(rg.sheet, rg.r1 + self._find(look, keys, 0 if exact else 1), rg.c1 + idx - 1)
            if isinstance(v, XlErr):
                raise XlErr(v.code)
            return v
        # MATCH(giá trị, vùng, [kiểu]): trả VỊ TRÍ (bắt đầu từ 1) của giá trị trong vùng.
        if name == "MATCH":
            return float(self._find(ev(0), self._flat(self._ev(args[1], ctx), ctx), int(self._n(ev(2))) if len(args) > 2 else 1) + 1)
        # INDEX(vùng, hàng, [cột]): lấy ô tại vị trí (hàng, cột) trong vùng.
        if name == "INDEX":
            rg = self._rng(args[0], ctx)
            r = int(self._n(ev(1))) if len(args) > 1 else 1
            c = int(self._n(ev(2))) if len(args) > 2 else 1
            if rg.r1 == rg.r2 and len(args) == 2:
                r, c = 1, r
            rr, cc = rg.r1 + (r or 1) - 1, rg.c1 + (c or 1) - 1
            if not (rg.r1 <= rr <= rg.r2 and rg.c1 <= cc <= rg.c2):
                raise XlErr("#REF!")
            return self._cell(rg.sheet, rr, cc)
        # Hàm chưa được hỗ trợ -> báo không hỗ trợ (để lớp Sheet dùng phương án dự phòng).
        raise XlUnsupported(f"function {name}")
    # _find(look, keys, mt): tìm vị trí (từ 0) cho MATCH/VLOOKUP.
    # mt=0: khớp đúng | mt>0: giá trị lớn nhất <= look (dữ liệu sắp tăng) | mt<0: giá trị nhỏ nhất >= look. Không
    # có -> #N/A.
    @staticmethod
    def _find(look, keys, mt):
        """Vị trí (0-based) cho MATCH/VLOOKUP: mt=0 khớp đúng; mt>0 lớn nhất <= look; mt<0 nhỏ nhất >= look."""
        best = None
        for i, k in enumerate(keys):
            if k is None or isinstance(k, XlErr) or isinstance(k, str) != isinstance(look, str):
                continue
            if mt == 0:
                if _cmp("=", k, look):
                    return i
            elif _cmp("<=" if mt > 0 else ">=", k, look):
                best = i
            else:
                break
        if best is None:
            raise XlErr("#N/A")
        return best
    # _cond: lõi của SUMIF/SUMIFS/COUNTIF/COUNTIFS. Ghép từng cặp (vùng, tiêu chí); với mỗi vị trí, nếu TẤT CẢ
    # tiêu chí thoả thì đếm hoặc cộng ô tương ứng của vùng tổng.
    def _cond(self, name, args, ctx):
        tgt, rest = None, args
        if name == "SUMIFS":
            tgt, rest = self._rng(args[0], ctx), args[1:]
        elif name == "SUMIF":
            rest, tgt = args[:2], self._rng(args[2] if len(args) > 2 else args[0], ctx)
        pairs = [(self._rng(rest[i], ctx), _crit(self._sc(self._ev(rest[i + 1], ctx), ctx))) for i in range(0, len(rest) - 1, 2)]
        base, total, cnt = pairs[0][0], 0.0, 0
        for r, c, _ in self._cells(base):
            dr, dc = r - base.r1, c - base.c1
            vals = [self._cell(g.sheet, g.r1 + dr, g.c1 + dc) for g, _ in pairs]
            if all(not isinstance(v, XlErr) and p(v) for v, (_, p) in zip(vals, pairs)):
                cnt += 1
                if tgt is not None:
                    x = self._cell(tgt.sheet, tgt.r1 + dr, tgt.c1 + dc)
                    total += x if isinstance(x, (int, float)) and not isinstance(x, bool) else 0.0
        return float(cnt) if tgt is None else total

# Sheet: mô hình một sheet Excel, phục vụ vừa cho LLM đọc (dump văn bản) vừa cho code truy vấn.
# Thuộc tính chính: grid (lưới giá trị), tables (các bảng phát hiện được, mỗi bảng có tiêu đề riêng), _raw (nội
# dung gốc từng ô).
class Sheet:
    # Khởi tạo Sheet: đọc giá trị -> đọc nội dung gốc -> tính công thức thiếu -> xử lý ô gộp -> tách bảng -> phân
    # tích tiêu đề.
    def __init__(self, name, ws, ws_f=None, engine=None):
        self.name = name
        # Kích thước vùng thực sự có dữ liệu/công thức.
        nrows, ncols = self._used_bounds(ws_f if ws_f is not None else ws)
        # Lưới giá trị (đã làm sạch), ban đầu toàn None.
        grid = [[None] * ncols for _ in range(nrows)]
        if nrows and ncols:
            for ri, row in enumerate(ws.iter_rows(min_row=1, max_row=nrows, max_col=ncols, values_only=True)):
                for ci, v in enumerate(row):
                    grid[ri][ci] = clean_cell(v)
        # _raw: nội dung GỐC của từng ô, khoá (hàng, cột) từ 0 - công thức giữ nguyên dạng '=...'. Dùng để tra cứu
        # theo địa chỉ ô.
        self._raw = {}
        try:
            src = ws_f if ws_f is not None else ws
            for ri, row in enumerate(src.iter_rows(min_row=1, max_row=nrows, max_col=ncols, values_only=True)):
                for ci, v in enumerate(row):
                    v = getattr(v, "text", v)  # công thức mảng
                    if v is None or (isinstance(v, str) and v.strip() == ""):
                        continue
                    self._raw[(ri, ci)] = v
        except Exception:
            pass
        # Ô chưa có giá trị tính sẵn nhưng có công thức -> tự tính: FormulaEngine -> _mini_eval -> _eval_sum (thận
        # trọng dần).
        if ws_f is not None and nrows and ncols:
            try:
                for ri, frow in enumerate(ws_f.iter_rows(min_row=1, max_row=nrows, max_col=ncols, values_only=True)):
                    for ci, fval in enumerate(frow):
                        if grid[ri][ci] is not None:
                            continue
                        fval = getattr(fval, "text", fval)  # công thức mảng
                        if isinstance(fval, str) and fval.startswith("="):
                            coord = f"{get_column_letter(ci + 1)}{ri + 1}"
                            ev = XFAIL
                            if engine is not None:
                                try:
                                    ev = engine.value(self.name, ri + 1, ci + 1)
                                except Exception:
                                    ev = XFAIL
                            if ev is XFAIL:  # bộ tính chưa hỗ trợ công thức này -> fallback cũ (thận trọng)
                                ev = _mini_eval(ws_f, grid, coord)
                                if ev is None:
                                    ev = _eval_sum(fval, grid)
                            elif isinstance(ev, str):
                                ev = clean_cell(ev)
                            if ev is not None:
                                grid[ri][ci] = ev
            except Exception:
                pass
        self.nrows, self.ncols = nrows, ncols
        # XỬ LÝ Ô GỘP: ô gộp RỘNG (một hàng, >= 4 cột hoặc >= 80% bề rộng bảng) coi là TIÊU ĐỀ -> ghi vào `wide`;
        # ô gộp thường -> NHÂN BẢN giá trị ra mọi ô trong vùng gộp.
        wide = []
        for rng in ws.merged_cells.ranges:
            r0, c0, r1, c1 = rng.min_row - 1, rng.min_col - 1, rng.max_row - 1, rng.max_col - 1
            if r0 >= nrows or c0 >= ncols:
                continue
            width = c1 - c0 + 1
            if width >= max(WIDE_MERGE_MIN_COLS, WIDE_MERGE_RATIO * ncols) and r1 == r0:
                wide.append((r0, c0))
                continue
            val = grid[r0][c0]
            for r in range(r0, min(r1, nrows - 1) + 1):
                for c in range(c0, min(c1, ncols - 1) + 1):
                    grid[r][c] = val
        # Lưu lưới cuối cùng. Phía dưới: tách sheet thành các bảng và phân tích tiêu đề từng bảng.
        self.grid = grid
        # spans: danh sách khoảng hàng của từng bảng; mặc định cả sheet là một bảng.
        spans = [(0, nrows)]
        if MULTI_TABLE and nrows:
            try:
                spans = self._split_tables(grid, nrows)
            except Exception:
                spans = [(0, nrows)]
        # Phân tích tiêu đề cho từng bảng (xem _analyze).
        self.tables = [self._analyze(grid, ncols, a, b, wide) for (a, b) in spans]
        # Giữ lại thuộc tính của BẢNG ĐẦU TIÊN (header_end, titles, title_rows, headers) để tương thích với code
        # cũ.
        t0 = self.tables[0]
        self.header_end, self.titles, self.title_rows, self.headers = t0["he"], t0["titles"], t0["title_rows"], t0["headers"]
    # _used_bounds(ws): trả (số hàng, số cột) của vùng thực sự có dữ liệu bằng cách quét mỗi hàng từ phải sang
    # trái.
    @staticmethod
    def _used_bounds(ws):
        """(nrows, ncols) của vùng thực sự có dữ liệu/công thức, không áp trần cứng."""
        last_r = last_c = 0
        for ri, row in enumerate(ws.iter_rows(values_only=True), 1):
            for ci in range(len(row) - 1, -1, -1):
                v = row[ci]
                if v is None or (isinstance(v, str) and v == ""):
                    continue
                last_r = ri
                if ci + 1 > last_c:
                    last_c = ci + 1
                break
        return last_r, last_c
    # _split_tables(grid, nrows): tách sheet thành các BẢNG xếp chồng dọc, ngăn cách bởi từ 2 hàng trống liên tiếp
    # trở lên.
    # Khối chỉ có chữ (tiêu đề) được gộp vào khối dữ liệu kế tiếp; dưới 2 bảng thì coi cả sheet là một bảng.
    @staticmethod
    def _split_tables(grid, nrows):
        """Tách sheet thành các bảng xếp chồng, ngăn cách bởi >= 2 dòng trống."""
        # empty[i] = True nếu hàng i hoàn toàn trống.
        empty = [all(v is None for v in row) for row in grid]
        # Bước 1: tìm các khối hàng liên tiếp (cho phép tối đa 1 hàng trống ở giữa; gặp 2 hàng trống thì ngắt
        # khối).
        blocks, r = [], 0
        while r < nrows:
            if empty[r]:
                r += 1
                continue
            s, e, gap = r, r, 0
            while r < nrows:
                if empty[r]:
                    gap += 1
                    if gap >= 2:
                        break
                else:
                    gap, e = 0, r
                r += 1
            blocks.append((s, e + 1))
        # Bước 2: duyệt các khối; khối có hàng dạng dữ liệu thì thành một bảng, khối chỉ toàn tiêu đề được 'treo'
        # (pending) để gộp vào bảng kế tiếp.
        spans, pending = [], None
        for (s, e) in blocks:
            if any(_data_like(grid[i]) for i in range(s, e)):
                spans.append([pending if pending is not None else s, e])
                pending = None
            elif pending is None:
                pending = s
        # Dưới 2 bảng -> coi cả sheet là MỘT bảng.
        if len(spans) < 2:
            return [(0, nrows)]
        # Bảng đầu bắt đầu từ hàng 0 và bảng cuối kéo dài hết sheet.
        spans[0][0] = 0
        spans[-1][1] = nrows
        return [tuple(sp) for sp in spans]
    # _analyze(...): dò TIÊU ĐỀ cho khối hàng [a, b). Hàng đầu tiên có dạng dữ liệu là mốc; các hàng phía trên là
    # tiêu đề.
    # Trả dict: start, end, he (hàng dữ liệu đầu), titles, title_rows, headers ({chữ cột: 'tầng1 > tầng2 > ...'}).
    @staticmethod
    def _analyze(grid, ncols, a, b, wide):
        """Dò tiêu đề cho khối dòng [a, b): dòng đầu có dạng dữ liệu, các dòng trước đó là tiêu đề."""
        # he: chỉ số hàng dữ liệu đầu tiên (mặc định = hàng đầu của khối, tức không có tiêu đề).
        he = a
        for r in range(a, b):
            if _data_like(grid[r]):
                he = r
                break
        # Quá nhiều hàng tiêu đề -> coi như KHÔNG có tiêu đề (tránh nhận nhầm).
        if he - a > MAX_HEADER_ROWS:
            he = a
        # titles: các dòng tiêu đề/ghi chú/đơn vị không thuộc riêng cột nào, dạng (số hàng, nội dung).
        titles = []
        # Ô gộp rộng nằm trong vùng tiêu đề -> là tiêu đề chung; xoá khỏi lưới để không lẫn vào header cột.
        for (r0, c0) in wide:
            if a <= r0 < he and grid[r0][c0] is not None:
                titles.append((r0 + 1, fmt_val(grid[r0][c0])))
                grid[r0][c0] = None
        # title_rows: các hàng tiêu đề chỉ có nội dung ở cột A (ghi chú/đơn vị) - không dùng để tạo header cột.
        title_rows = set()
        for r in range(a, he):
            filled = [c for c in range(ncols) if grid[r][c] is not None]
            if filled and all(c == 0 for c in filled):
                title_rows.add(r)
                titles.append((r + 1, fmt_val(grid[r][0])))
        # headers: với mỗi cột, ghép các tầng tiêu đề (từ trên xuống, bỏ trùng liên tiếp) bằng ' > '. Ví dụ 'Doanh
        # thu > 2023 > Quý 1'.
        headers = {}
        for c in range(ncols):
            parts = []
            for r in range(a, he):
                if r in title_rows:
                    continue
                v = grid[r][c]
                if v is None:
                    continue
                t = fmt_val(v)
                if not parts or parts[-1] != t:
                    parts.append(t)
            if parts:
                headers[get_column_letter(c + 1)] = " > ".join(parts)
        return {"start": a, "end": b, "he": he, "titles": titles, "title_rows": title_rows, "headers": headers}
    # v(addr): giá trị ô theo địa chỉ, ví dụ sh.v('B5'). Ngoài vùng dữ liệu -> None.
    def v(self, addr):
        col, row = coordinate_from_string(addr.replace("$", "").upper())
        r, c = row - 1, column_index_from_string(col) - 1
        if 0 <= r < self.nrows and 0 <= c < self.ncols:
            return self.grid[r][c]
        return None
    # raw(addr): nội dung GỐC của ô (công thức giữ nguyên chuỗi '=...'); ô trống -> None.
    def raw(self, addr):
        """Nội dung gốc của ô: chuỗi công thức đúng như viết (có '=') nếu là công thức, ngược lại giá trị gốc. Trống → None."""
        col, row = coordinate_from_string(addr.replace("$", "").upper())
        return self._raw.get((row - 1, column_index_from_string(col) - 1))
    # formula(addr): trả công thức đúng như viết trong ô (vd '=SUM(A1:A3)'); ô không phải công thức -> None.
    def formula(self, addr):
        """Công thức đúng như viết trong ô (vd '=SUM(A1:A3)'); None nếu ô không phải công thức."""
        v = self.raw(addr)
        return v if isinstance(v, str) and v.startswith("=") else None
    # raw_row(r): dict {chữ cột: nội dung gốc} của các ô KHÔNG trống ở hàng r.
    def raw_row(self, r):
        return {get_column_letter(c + 1): v for (rr, c), v in sorted(self._raw.items(), key=lambda kv: kv[0][1]) if rr == r - 1}
    # raw_col(letter): dict {số hàng: nội dung gốc} của các ô không trống ở cột `letter`.
    def raw_col(self, letter):
        ci = column_index_from_string(letter.upper()) - 1
        return {rr + 1: v for (rr, c), v in sorted(self._raw.items()) if c == ci}
    # row(r): dict {chữ cột: giá trị} của các ô không trống ở hàng r. Ví dụ sh.row(5) -> {'A': 'Táo', 'B': 10}.
    def row(self, r):
        if not (1 <= r <= self.nrows):
            return {}
        return {get_column_letter(c + 1): v for c, v in enumerate(self.grid[r - 1]) if v is not None}
    # col(letter): dict {số hàng: giá trị} của các ô không trống ở cột. Ví dụ sh.col('B') -> {5: 10, 6: 20}.
    def col(self, letter):
        c = column_index_from_string(letter.upper()) - 1
        if not (0 <= c < self.ncols):
            return {}
        return {r + 1: self.grid[r][c] for r in range(self.nrows) if self.grid[r][c] is not None}
    
    # find_rows(text, col='A', exact=False): tìm các SỐ HÀNG có nhãn khớp `text` trong cột `col` (so sánh qua
    # norm: bỏ dấu, bỏ khoảng trắng).
    # Trả danh sách khớp CHÍNH XÁC nếu có; nếu không mới trả danh sách khớp MỘT PHẦN (trừ khi exact=True).
    def find_rows(self, text, col="A", exact=False):
        n = norm(text)
        nj = n.replace(" ", "")
        exact_hits, partial_hits = [], []
        for r, val in self.col(col).items():
            vn = norm(val)
            if not vn:
                continue
            if vn == n or vn.replace(" ", "") == nj:
                exact_hits.append(r)
            elif n in vn or nj in vn.replace(" ", ""):
                partial_hits.append(r)
        if exact:
            return exact_hits
        return exact_hits if exact_hits else partial_hits

    # find_cols(text): tìm các CHỮ CỘT có đường dẫn tiêu đề chứa `text` (xét mọi bảng trong sheet, không trùng
    # lặp).
    def find_cols(self, text):
        n = norm(text)
        nj = n.replace(" ", "")
        out = []
        for t in self.tables:
            for letter, path in t["headers"].items():
                pn = norm(path)
                if (n in pn or nj in pn.replace(" ", "")) and letter not in out:
                    out.append(letter)
        return out
    # _dump_table(t): sinh các dòng văn bản mô tả MỘT bảng: TITLE/UNITS, COLUMN HEADERS, HEADER ROWS, ROWS.
    def _dump_table(self, t):
        lines = []
        seen = set()
        tl = []
        # Phần TITLE / UNITS: tiêu đề chung, ghi chú, đơn vị (bỏ trùng và bỏ tiêu đề đã trùng với header cột).
        for r, title in t["titles"]:
            if title and title not in seen and title not in t["headers"].values():
                seen.add(title)
                tl.append(f"  R{r}: {title}")
        if tl:
            lines.append("TITLE / UNITS:")
            lines += tl
        # Phần COLUMN HEADERS: mỗi dòng 'chữ cột: đường dẫn tiêu đề'.
        if t["headers"]:
            lines.append("COLUMN HEADERS (letter: header path):")
            for letter, p in t["headers"].items():
                lines.append(f"  {letter}: {p}")
        # Phần HEADER ROWS: các ô tiêu đề dạng thô kèm địa chỉ ô (để LLM tra theo địa chỉ khi cần).
        if DUMP_HEADER_CELLS:
            hl = []
            for r in range(t["start"], t["he"]):
                if r in t["title_rows"]:
                    continue
                cells = [f"{get_column_letter(c + 1)}={fmt_val(v)}" for c, v in enumerate(self.grid[r]) if v is not None]
                if cells:
                    hl.append(f"R{r + 1} | " + " | ".join(cells))
            if hl:
                lines.append("HEADER ROWS (same cells as above, by address):")
                lines += hl
        # Phần ROWS: từng hàng dữ liệu dạng 'R<hàng> | A=... | B=...' (ô trống được bỏ qua).
        lines.append("ROWS:")
        for r in range(t["he"], t["end"]):
            cells = [f"{get_column_letter(c + 1)}={fmt_val(v)}" for c, v in enumerate(self.grid[r]) if v is not None]
            if cells:
                lines.append(f"R{r + 1} | " + " | ".join(cells))
        return lines
    # dump(): văn bản mô tả CẢ sheet gửi cho LLM. Sheet có nhiều bảng thì thêm cảnh báo: cùng một chữ cột có thể
    # mang nghĩa khác nhau ở mỗi bảng.
    def dump(self):
        lines = [f"=== SHEET: {self.name} (rows 1-{self.nrows}, columns A-{get_column_letter(max(self.ncols, 1))}) ==="]
        if len(self.tables) == 1:
            lines += self._dump_table(self.tables[0])
        else:
            lines.append(f"NOTE: this sheet holds {len(self.tables)} separate tables stacked vertically; "
                         "each table has its own COLUMN HEADERS, so the same column letter may mean different things per table.")
            for k, t in enumerate(self.tables, 1):
                lines.append(f"--- TABLE {k} (rows {t['start'] + 1}-{t['end']}) ---")
                lines += self._dump_table(t)
        return "\n".join(lines)
# Bộ nhớ đệm các workbook đã nạp (theo đường dẫn file) + khoá chống xung đột khi nhiều luồng cùng nạp.
_wb_cache, _wb_lock = {}, threading.Lock()

# load_workbook_sheets(path): nạp file .xlsx, trả về (dict {tên sheet: Sheet}, văn bản dump toàn workbook).
# Kết quả được cache nên các câu hỏi cùng một file không phải nạp lại.
def load_workbook_sheets(path):
    with _wb_lock:
        if path in _wb_cache:
            return _wb_cache[path]
        # Mở lần 1: lấy GIÁ TRỊ đã tính sẵn (data_only=True).
        wb_v = openpyxl.load_workbook(path, data_only=True)
        # Mở lần 2: lấy CÔNG THỨC (data_only=False); nếu lỗi thì wb_f = None (vẫn dùng được giá trị).
        try:
            wb_f = openpyxl.load_workbook(path, data_only=False)
        except Exception:
            wb_f = None
        sheets = {}
        # Tạo bộ tính công thức xuyên sheet (cần cả hai workbook); lỗi -> không dùng engine.
        try:
            engine = FormulaEngine(wb_v, wb_f) if wb_f is not None else None
        except Exception:
            engine = None
        # Dựng đối tượng Sheet cho từng sheet; bỏ qua sheet hoàn toàn trống.
        for ws in wb_v.worksheets:
            wsf = wb_f[ws.title] if (wb_f is not None and ws.title in wb_f.sheetnames) else None
            sh = Sheet(ws.title, ws, wsf, engine)
            if any(v is not None for row in sh.grid for v in row):
                sheets[ws.title] = sh
        # Ghép bản dump của mọi sheet thành một văn bản duy nhất để đưa cho LLM.
        text = "\n\n".join(s.dump() for s in sheets.values())
        _wb_cache[path] = (sheets, text)
        return _wb_cache[path]

# =============================================================================
# 3. SANDBOX CHẠY CODE DO LLM SINH
# =============================================================================
# Chỉ cho phép code của LLM import các module an toàn này.
_SAFE_MODULES = {"datetime", "math", "re", "statistics", "itertools", "collections", "functools"}

# _safe_import: thay thế hàm __import__ trong sandbox - chặn mọi module ngoài danh sách an toàn (open file, os,
# sys... đều bị cấm).
def _safe_import(name, *a, **k):
    if name.split(".")[0] not in _SAFE_MODULES:
        raise ImportError(f"import {name} khong duoc phep")
    return __import__(name, *a, **k)
# Danh sách hàm/kiểu dựng sẵn được phép dùng trong sandbox (KHÔNG có open, eval, exec, getattr, __import__
# gốc...).
_BUILTIN_NAMES = ["abs", "all", "any", "bool", "dict", "enumerate", "filter", "float", "int", "len", "list", "map", "max",
                  "min", "range", "round", "set", "sorted", "str", "sum", "tuple", "zip", "isinstance", "reversed",
                  "divmod", "pow", "repr", "next", "iter", "slice", "frozenset", "Exception", "ValueError", "KeyError",
                  "IndexError", "TypeError", "ZeroDivisionError", "StopIteration", "True", "False", "None"]

# run_code(code, sheets, timeout): chạy code do LLM sinh ra và trả (answer, lỗi).
# Code được cung cấp sẵn: SHEETS, sh (sheet đầu tiên), num, norm, to_date, dt, math, re, statistics.
# Chạy trong một thread daemon để có thể đặt timeout.
# Lưu ý: đây là sandbox 'mềm' (hạn chế builtins + module), không phải cách ly an ninh tuyệt đối.
def run_code(code, sheets, timeout=None):
    import builtins as _b
    # Dựng bộ builtins an toàn từ danh sách cho phép.
    safe = {n: getattr(_b, n) for n in _BUILTIN_NAMES if hasattr(_b, n)}
    safe["__import__"] = _safe_import
    # Vô hiệu hoá print (code LLM in gì cũng bị bỏ).
    safe["print"] = lambda *a, **k: None
    # `sh` mặc định = sheet đầu tiên trong dict.
    first = next(iter(sheets.values()))
    # Môi trường chạy: những đối tượng mà code của LLM được phép dùng; `answer` là biến kết quả.
    env = {"__builtins__": safe, "SHEETS": sheets, "sh": first, "num": num, "norm": norm, "to_date": to_date,
           "dt": dt, "math": math, "re": re, "statistics": statistics, "answer": None}
    res = {}
    # target(): chạy code, bắt MỌI lỗi, lưu đáp án hoặc thông báo lỗi vào res.
    def target():
        try:
            exec(compile(code, "<llm_code>", "exec"), env)
            res["answer"] = env.get("answer")
        except BaseException as e:  # noqa
            res["err"] = f"{type(e).__name__}: {e}"
    # Chạy trong thread daemon; join(timeout) để giới hạn thời gian.
    th = threading.Thread(target=target, daemon=True)
    th.start()
    th.join(timeout if timeout is not None else SANDBOX_TIMEOUT)
    # Quá thời gian mà code vẫn chạy -> bỏ kết quả, báo Timeout.
    if th.is_alive():
        return None, "Timeout: code chay qua lau"
    return res.get("answer"), res.get("err")

# =============================================================================
# 4. CHUẨN HOÁ ĐÁP ÁN (Exact Match) — format lấy từ QUERY đa ngôn ngữ
# =============================================================================
# Từ chỉ số bằng chữ (tiếng Anh/Việt) -> số.
_WORD_NUM = {
    "zero": 0, "one": 1, "two": 2, "three": 3, "four": 4, "five": 5, "six": 6,
    "seven": 7, "eight": 8, "nine": 9, "ten": 10,
    "một": 1, "mot": 1, "hai": 2, "ba": 3, "bốn": 4, "bon": 4, "năm": 5, "nam": 5,
}
# Từ/chữ số chỉ số chữ số thập phân.
_PLACES = {"one": 1, "two": 2, "three": 3, "four": 4, "five": 5,
           "1": 1, "2": 2, "3": 3, "4": 4, "5": 5, "0": 0}

# _nfc(s): chuẩn hoá Unicode về dạng NFC (tiếng Việt có thể ở dạng dấu tách rời NFD).
def _nfc(s):
    """Chuẩn hoá Unicode về NFC (tiếng Việt có thể ở dạng NFD / dấu tách rời)."""
    return unicodedata.normalize("NFC", s) if isinstance(s, str) else s

# _parse_place_token(tok): đổi một từ/số ('two', 'hai', '3') thành số chữ số thập phân; không nhận ra -> None.
def _parse_place_token(tok):
    tok = (tok or "").strip().lower()
    if tok in _PLACES:
        return _PLACES[tok]
    if tok in _WORD_NUM:
        return _WORD_NUM[tok]
    if tok.isdigit():
        return int(tok)
    return None

# _wants_integer(q): câu hỏi có yêu cầu đáp án là SỐ NGUYÊN không ('integer', 'số nguyên', '整数', '정수').
def _wants_integer(q):
    if not q:
        return False
    q = _nfc(q)  
    ql = q.lower()
    if re.search(r"\binteger\b|whole\s+number|số\s+nguyên", ql):
        return True
    if re.search(r"整数|整數|정수", q):
        return True
    return False
# -----------------------------------------------------------------------------
# -----------------------------------------------------------------------------
# Dấu ngăn cách khi đáp án gồm nhiều giá trị (mặc định ', ').
MULTI_SEP = os.environ.get("MULTI_SEP", ", ")
# Tên kiểu làm tròn nội bộ -> hằng số kiểu làm tròn của Decimal.
_MODE_DEC = {"half_up": ROUND_HALF_UP, "up": ROUND_UP, "down": ROUND_DOWN, "ceil": ROUND_CEILING,
             "floor": ROUND_FLOOR, "half_even": ROUND_HALF_EVEN}
# Mô tả từng kiểu làm tròn bằng tiếng Anh (đưa vào prompt cho LLM).
_MODE_LABEL = {"half_up": "round half up / nearest (Excel ROUND)", "up": "round up, away from zero (Excel ROUNDUP)",
               "down": "round down / truncate toward zero (Excel ROUNDDOWN)", "ceil": "ceiling (toward +inf)",
               "floor": "floor (toward -inf)", "half_even": "banker's rounding (half to even)"}
# _fn_mode(word): tạo regex nhận 'floor'/'ceiling' CHỈ khi dùng như kiểu làm tròn (floor(x), use floor...),
# không nhận danh từ thường như 'floor area', 'ceiling height'.
def _fn_mode(word):
    """'floor'/'ceiling' chỉ là kiểu làm tròn khi dùng như hàm/động từ (floor(x), use floor, floor the result...);
    không bắt danh từ thường như "floor area", "which floor", "ceiling height"."""
    not_noun = r"(?!\s*(?:area|plan|price|space|level|number|height|limit|rate|\d))"
    return (rf"\b{word}\s*\(|"
            rf"\b(?:use|using|apply|applying|take|taking|via|then)\s+(?:the\s+)?(?:excel\s+)?{word}\b{not_noun}|"
            rf"\bround(?:ed|ing|s)?\s+(?:\w+\s+){{0,5}}(?:using|with|by|via)\s+(?:the\s+)?{word}\b{not_noun}|"
            rf"\b{word}\s+(?:the\s+)?(?:result|value|number|average|mean|sum|total|figure|it)\b|"
            rf"\b{word}\s+to\s+(?:the\s+)?(?:nearest|\d|an?\s+integer|whole|integer)")

# Mẫu regex nhận diện KIỂU LÀM TRÒN trong câu hỏi (đa ngôn ngữ EN/VI/ZH/JA/KO).
# Thứ tự quan trọng: mẫu cụ thể đặt trước, 'half_up' (làm tròn chung chung) đặt cuối cùng.
_MODE_PATTERNS = [
    ("half_even", r"banker|half[\s\-]*(?:to[\s\-]*)?even|round(?:ed|ing)?\s+to\s+even|làm\s*tròn\s*chẵn|银行家|銀行家|짝수"),
    ("ceil", _fn_mode(r"ceil(?:ing)?") + r"|向上取整|向上舍入"),
    ("floor", _fn_mode("floor") + r"|向下取整|向下舍入"),
    ("up", r"round(?:ed|ing|s)?\s+up\b|\broundup\b|làm\s*tròn\s*lên|tròn\s*lên|进一法|進一法|切り上げ|(?<!반)올림"),
    ("down", r"round(?:ed|ing|s)?\s+down\b|\brounddown\b|làm\s*tròn\s*xuống|tròn\s*xuống|truncat\w*|"
             r"cắt\s*(?:bỏ|bớt)|截断|截斷|截尾|舍去|去尾|切り捨て|切捨て|내림|버림"),
    ("half_up", r"round(?:ed|ing|s)?\b|làm\s*tròn|四舍五入|四捨五入|반올림|丸め"),
]
# Biên dịch sẵn các regex ở trên.
_MODE_RE = [(m, re.compile(p, re.I)) for m, p in _MODE_PATTERNS]
# Chữ số Hán -> số.
_CN_NUM = {"一": 1, "二": 2, "两": 2, "兩": 2, "三": 3, "四": 4, "五": 5}
# Mẫu regex cho số viết bằng chữ tiếng Anh hoặc chữ số.
_NUMW = r"zero|one|two|three|four|five|six|\d{1,2}"
# Từ chỉ số bổ sung: tiếng Việt, tiếng Hàn, tiếng Pháp.
_EXTRA_NUM = {"nhất": 1, "nhat": 1, "첫째": 1, "둘째": 2, "셋째": 3, "넷째": 4,
              "une": 1, "un": 1, "deux": 2, "trois": 3, "quatre": 4, "cinq": 5}

# _pl(tok): đổi token thành số chữ số thập phân, hiểu cả chữ Hán, Hàn, Pháp, Việt.
def _pl(tok):
    tok = (tok or "").strip().lower()
    if tok in _CN_NUM:
        return _CN_NUM[tok]
    if tok in _EXTRA_NUM:
        return _EXTRA_NUM[tok]
    return _parse_place_token(tok)
# Danh sách (regex, hàm đổi) nhận diện SỐ CHỮ SỐ THẬP PHÂN trong câu hỏi, nhiều ngôn ngữ.
# Quy ước: số dương = số chữ số thập phân; số âm = làm tròn về hàng chục/trăm/nghìn; 0 = số nguyên.
_PLACE_PATTERNS = [
    (re.compile(r"nearest\s+(tenth|hundredth|thousandth)", re.I), lambda t: {"tenth": 1, "hundredth": 2, "thousandth": 3}[t.lower()]),
    (re.compile(r"nearest\s+(ten|hundred|thousand)\b", re.I), lambda t: -{"ten": 1, "hundred": 2, "thousand": 3}[t.lower()]),
    (re.compile(r"hàng\s*(chục|trăm|nghìn|ngàn)", re.I), lambda t: -{"chục": 1, "trăm": 2, "nghìn": 3, "ngàn": 3}[t.lower()]),
    (re.compile(r"\b(" + _NUMW + r")[\s\-]*decimals?\b", re.I), _pl),
    (re.compile(r"decimal\s+places?\s*[:=]?\s*(\d+)", re.I), _pl),
    (re.compile(r"\b(\d+)\s*d\.?\s*p\.?\b", re.I), _pl),
    (re.compile(r"(một|hai|ba|bốn|bon|năm|nam|\d+)(?=\s*chữ\s*số\s*thập\s*phân)", re.I), _pl),
    (re.compile(r"làm\s*tròn\s*(?:đến|tới|về)?\s*(một|hai|ba|bốn|bon|\d+)\b", re.I), _pl),
    (re.compile(r"(\d+)\s*số\s*lẻ", re.I), _pl),
    (re.compile(r"thập\s*phân\s*[:=]?\s*(\d+)", re.I), _pl),
    (re.compile(r"保留\s*([一二两兩三四五\d]+)\s*位"), _pl),
    (re.compile(r"(\d+)\s*位小数"), _pl),
    (re.compile(r"四舍五入.*?(\d+)\s*位"), _pl),
    (re.compile(r"小数点以下\s*(\d+)\s*桁"), _pl),
    (re.compile(r"小数第\s*(\d+)\s*位"), _pl),
    (re.compile(r"소수점\s*(\d+)\s*자리"), _pl),
    (re.compile(r"(\d+)\s*자리\s*소수"), _pl),
    # --- bổ sung nhiều cách diễn đạt khác ---
    (re.compile(r"\b(" + _NUMW + r")[\s\-]*(?:digits?|numbers?|figures?)\s+(?:after|behind|following|to\s+the\s+right\s+of)\s+(?:the\s+)?(?:decimal|comma|dot|period|point)", re.I), _pl),
    (re.compile(r"(?:nearest|to|precision\s+of|increments?\s+of)\s+0\.(0*)1(?!\d)", re.I), lambda t: len(t) + 1),
    (re.compile(r"nearest\s+(10{1,3})\b", re.I), lambda t: -(len(t) - 1)),
    (re.compile(r"(?:format|pattern|as)\s*[:：]?\s*[\"'“]?[#,]*0\.(0+)", re.I), lambda t: len(t)),
    (re.compile(r"(một|hai|ba|bốn|bon|năm|nam|\d+)\s*chữ\s*số\s*(?:sau|đằng\s*sau)\s*dấu\s*(?:phẩy|thập\s*phân|chấm)", re.I), _pl),
    (re.compile(r"chữ\s*số\s*thập\s*phân\s*thứ\s*(nhất|nhat|hai|ba|bốn|bon)", re.I), _pl),
    (re.compile(r"(một|hai|ba|bốn|bon|\d+)\s*số\s*thập\s*phân", re.I), _pl),
    (re.compile(r"小数点[后後]\s*第?\s*([一二两兩三四五\d]+)\s*位"), _pl),
    (re.compile(r"([一二两兩三四五\d]+)\s*位小数"), _pl),
    (re.compile(r"精确到\s*0\.(0*)1(?!\d)"), lambda t: len(t) + 1),
    (re.compile(r"小数(?:点(?:以下)?)?第?\s*([一二三四五\d]+)\s*[位桁]"), _pl),
    (re.compile(r"소수(?:점)?\s*(?:이하\s*)?(첫째|둘째|셋째|넷째)\s*자리"), _pl),
    (re.compile(r"소수점\s*(?:이하\s*)?(\d+)\s*자리"), _pl),
    (re.compile(r"\b(\d+|une?|deux|trois|quatre|cinq)\s*décimales?\b", re.I), _pl),
    (re.compile(r"\b(integers?|whole\s+numbers?)\b|(số\s+nguyên)|(整数|整數|(?<![上下])取整|정수)", re.I), lambda t: 0),
]

# _scan(patterns, s): chạy nhiều regex trên chuỗi s, trả [(đầu, cuối, giá trị)] KHÔNG chồng lấn (regex đứng trước
# thắng), sắp theo vị trí xuất hiện.
def _scan(patterns, s):
    """Trả về [(start, end, giá trị)] không chồng nhau, theo thứ tự xuất hiện."""
    taken = []
    for rx, fn in patterns:
        for m in rx.finditer(s):
            a, b = m.span()
            if any(a < e and s_ < b for s_, e, _ in taken):
                continue
            g = next((x for x in m.groups() if x), "") if m.groups() else ""
            try:
                val = fn(g)
            except Exception:
                val = None
            if val is None:
                continue
            taken.append((a, b, val))
    return sorted(taken)

# _modes_in(s): danh sách các kiểu làm tròn xuất hiện trong s, theo thứ tự.
def _modes_in(s):
    return [v for _, _, v in _scan([(rx, (lambda m: (lambda _t: m))(m)) for m, rx in _MODE_RE], s)]

# _places_in(s): danh sách các số chữ số thập phân xuất hiện trong s, theo thứ tự.
def _places_in(s):
    return [v for _, _, v in _scan(_PLACE_PATTERNS, s)]
# Regex tách câu hỏi thành các mệnh đề (theo dấu phẩy, chấm phẩy, 'and', 'và', 'then', '和'...).
_CLAUSE_SPLIT = re.compile(
    r"\s*(?:[;,，、；]|\band\b|\bas well as\b|\bthen\b|\bvà\b|\bcũng như\b|和|与|與|及|と|및)\s*", re.I)

# rounding_specs(q): trả tuple các (kiểu làm tròn, số chữ số) theo thứ tự trong câu hỏi; () nếu không yêu cầu làm
# tròn.
# Ví dụ: 'average, round to 2 decimals' -> (('half_up', 2),); 'round up the average' -> (('up', 0),).
@lru_cache(maxsize=4096)
def rounding_specs(q):
    """Danh sách (mode, places) theo thứ tự trong câu hỏi. () nếu không yêu cầu làm tròn."""
    if not q:
        return ()
    # gm/gp: tập các kiểu làm tròn / số chữ số xuất hiện trong CẢ câu.
    gm, gp = set(_modes_in(q)), set(_places_in(q))
    # Nếu cả câu chỉ có MỘT kiểu (hoặc một số chữ số) thì dùng làm giá trị mặc định cho các mệnh đề còn thiếu.
    d_mode = next(iter(gm)) if len(gm) == 1 else None
    d_places = next(iter(gp)) if len(gp) == 1 else None
    # specs: kết quả (mode, places) thu được.
    specs = []
    # Xét từng mệnh đề của câu hỏi.
    for c in _CLAUSE_SPLIT.split(q):
        if not c or not c.strip():
            continue
        ms, ps = _modes_in(c), _places_in(c)
        if not ms and not ps:
            continue
        if not ps and d_places is None:
            if all(m == "half_up" for m in ms):
                continue  # "round the average" -> không đủ thông tin, giữ hành vi cũ
            ps = [0]    # "round up the average" -> ROUNDUP về số nguyên
        for i in range(max(len(ms), len(ps), 1)):
            mode = (ms[i] if i < len(ms) else ms[-1]) if ms else (d_mode or "half_up")
            places = (ps[i] if i < len(ps) else ps[-1]) if ps else d_places
            specs.append((mode, places))
    # Bỏ trùng nhưng giữ nguyên thứ tự.
    return tuple(dict.fromkeys(specs))

# _fmt_round(v, places, mode): làm tròn đúng kiểu yêu cầu và GIỮ ĐỦ số chữ số thập phân (3 -> '3.00' nếu
# places=2).
# Ví dụ: _fmt_round(2.675, 2) -> '2.68' (round() của Python cho 2.67).
def _fmt_round(v, places, mode="half_up"):
    """Làm tròn và giữ đúng số chữ số thập phân theo yêu cầu."""
    places = 0 if places is None else int(places)
    # Dùng Decimal tạo từ chuỗi 15 chữ số có nghĩa để tránh sai số nhị phân của float.
    try:
        r = Decimal(f"{float(v):.15g}").quantize(
            Decimal(1).scaleb(-places),
            rounding=_MODE_DEC.get(mode, ROUND_HALF_UP)
        )
    except (InvalidOperation, ValueError, OverflowError):
        return f"{round(float(v), places):.{max(places, 0)}f}"
    # Tránh kết quả '-0'.
    if r == 0:
        r = abs(r)
    # Không có phần thập phân -> trả số nguyên; ngược lại định dạng đủ số chữ số.
    if places <= 0:
        return str(int(r))
    return f"{r:.{places}f}"

# _post_multi(ans, specs): áp NHIỀU yêu cầu làm tròn lên đáp án. Trả None nếu không áp dụng được (để rơi về luồng
# xử lý thường).
def _post_multi(ans, specs):
    """Nhiều yêu cầu làm tròn. Trả None nếu không áp dụng được (để rơi về luồng cũ)."""
    if ans is None or isinstance(ans, bool):
        return None
    # Chuẩn hoá đầu vào thành danh sách các giá trị (list / số / chuỗi có <FINAL_ANSWER> hoặc nhiều giá trị phân
    # tách bằng ; | ,).
    if isinstance(ans, (list, tuple)):
        items = list(ans)
    elif isinstance(ans, (int, float)):
        items = [ans]
    else:
        s = re.sub(r"\s+", " ", str(ans)).strip()
        m = re.search(r"<FINAL_ANSWER>\s*(.*?)\s*</FINAL_ANSWER>", s, flags=re.S | re.I)
        if m:
            s = m.group(1).strip()
        s = _PREFIX_JUNK.sub("", s).strip()
        items = [p for p in re.split(r"\s*[;|]\s*|,\s+", s) if p != ""]
    # Đổi từng phần tử sang float; có phần tử không đổi được (hoặc NaN/Inf) -> trả None.
    vals = []
    for x in items:
        v = float(x) if isinstance(x, (int, float)) and not isinstance(x, bool) else _try_float(x)
        if v is None or math.isnan(v) or math.isinf(v):
            return None
        vals.append(v)
    # 1 giá trị + nhiều yêu cầu -> làm tròn giá trị đó theo TỪNG yêu cầu;
    # số giá trị = số yêu cầu -> ghép đôi theo thứ tự; còn lại -> None.
    if len(vals) == 1:
        pairs = [(vals[0], sp) for sp in specs]
    elif len(vals) == len(specs):
        pairs = list(zip(vals, specs))
    else:
        return None
    return MULTI_SEP.join(_fmt_round(v, p, m) for v, (m, p) in pairs)

# multi_round_hint(query): nếu câu hỏi có từ 2 yêu cầu làm tròn trở lên, sinh đoạn GHI CHÚ đưa vào prompt
# bảo LLM trả số đầy đủ độ chính xác (không tự làm tròn), việc làm tròn sẽ do hậu xử lý đảm nhiệm.
def multi_round_hint(query):
    specs = rounding_specs(query)
    if len(specs) < 2:
        return ""
    lines = []
    for i, (m, p) in enumerate(specs, 1):
        pl = "integer" if p == 0 else (f"{p} decimal places" if p > 0 else f"nearest 10^{-p}")
        lines.append(f"  {i}. {_MODE_LABEL[m]}, {pl}")
    return ("\n\nOUTPUT FORMAT NOTE (overrides rule 7): the question asks for several rounded outputs, in this order:\n"
            + "\n".join(lines) +
            "\nDo NOT round or format yourself. Set `answer` to the FULL-PRECISION unrounded number when all outputs "
            "are the same quantity; if the outputs are DIFFERENT quantities, set `answer` to a Python list of "
            "full-precision numbers, one per output, in the order above. The post-processor applies the roundings and joins them.")
# Regex bỏ tiền tố thừa ở đáp án như 'The answer is:', 'Kết quả:', 'Đáp án:'.
_PREFIX_JUNK = re.compile(
    r"^(?:the\s+answer\s+is\s*[:：]?|(?:final\s*answer|answer|result|đáp\s*án|kết\s*quả)\s*[:：])\s*",
    re.I,
)

# _try_float(s): chuỗi -> float theo luật dấu thập phân giống num(); nan/none/null hoặc không phải số -> None.
def _try_float(s):
    if s is None:
        return None
    t = str(s).strip().replace("\u00a0", "").replace("\u202f", "").replace(" ", "")
    if not t or re.fullmatch(r"(?i)nan|none|null", t):
        return None
    if "," in t and "." in t:
        if t.rfind(",") > t.rfind("."):
            t = t.replace(".", "").replace(",", ".")
        else:
            t = t.replace(",", "")
    elif "," in t:
        parts = t.split(",")
        if len(parts) == 2 and len(parts[1]) != 3:
            t = t.replace(",", ".")
        else:
            t = t.replace(",", "")
    try:
        return float(t)
    except ValueError:
        return None
# Ánh xạ ký hiệu ngày trong câu hỏi (yyyy, yy, mm, dd) -> mã strftime.
_DATE_TOKEN = {"yyyy": "%Y", "yy": "%y", "mm": "%m", "dd": "%d"}
# Regex bắt mẫu định dạng ngày dạng 3 phần, ví dụ 'dd/mm/yyyy' hoặc 'yyyy-mm-dd'.
_DATE_TRIPLE_RE = re.compile(r"(?<![a-z])(yyyy|yy|mm|dd)\s*([/\-.])\s*(yyyy|yy|mm|dd)\s*\2\s*(yyyy|yy|mm|dd)(?![a-z])", re.I)

# _date_format_from_query(q): trả chuỗi strftime theo định dạng ngày mà câu hỏi yêu cầu; mặc định ISO
# (YYYY-MM-DD).
def _date_format_from_query(q):
    """Chuỗi strftime theo yêu cầu trong query; mặc định ISO YYYY-MM-DD."""
    if not q:
        return "%Y-%m-%d"
    q = _nfc(q) 
    if re.search(r"yyyy\s*年\s*mm\s*月\s*dd\s*日?", q, re.I):
        return "%Y年%m月%d日"
    m = _DATE_TRIPLE_RE.search(q)
    if m:
        toks = [m.group(i).lower() for i in (1, 3, 4)]
        if len(set(toks)) == 3:
            return m.group(2).join(_DATE_TOKEN[t] for t in toks)
    ql = q.lower()
    if re.search(r"ngày\s*[/\-.]\s*tháng\s*[/\-.]\s*năm", ql):
        return "%d/%m/%Y"
    if re.search(r"tháng\s*[/\-.]\s*ngày\s*[/\-.]\s*năm", ql):
        return "%m/%d/%Y"
    return "%Y-%m-%d"

# _fmt_date_answer(d, q): định dạng ngày theo yêu cầu trong câu hỏi; chỉ thêm ' HH:MM:SS' nếu giờ khác 00:00:00.
def _fmt_date_answer(d, q):
    has_time = isinstance(d, dt.datetime) and (d.hour, d.minute, d.second) != (0, 0, 0)
    if not isinstance(d, dt.datetime):
        d = dt.datetime(d.year, d.month, d.day)
    out = d.strftime(_date_format_from_query(q))
    return out + d.strftime(" %H:%M:%S") if has_time else out
# Regex nhận chuỗi ngày ISO lúc nửa đêm (2024-01-05 00:00:00) để rút gọn về chỉ còn ngày.
_ISO_MIDNIGHT = re.compile(r"^(\d{4}-\d{2}-\d{2})[ T]00:00:00(?:\.0+)?$")
# -----------------------------------------------------------------------------
# -----------------------------------------------------------------------------
# Regex bắt chú thích dạng '(1)' ở cuối nhãn.
_FOOTNOTE_TAIL = re.compile(r"\s*\(\d+\)\s*$")

# _lab_key(s): khoá chuẩn hoá NHẸ của nhãn (NFC, gom khoảng trắng, cắt hai đầu).
def _lab_key(s):
    return re.sub(r"\s+", " ", unicodedata.normalize("NFC", str(s))).strip()

# _lab_loose(k): khoá NỚI LỎNG: bỏ dấu câu ở cuối + không phân biệt hoa/thường (để khôi phục dấu chấm bị mất).
def _lab_loose(k):
    """Khoá nới lỏng: bỏ dấu câu cuối + không phân biệt hoa/thường (để khôi phục dấu chấm bị mất)."""
    return k.rstrip(" .,;:").casefold()

# sheet_labels(sh): thu mọi NHÃN chữ của một sheet, trả (tập nhãn chính xác, {khoá lỏng: tập nhãn gốc}, tập từ
# viết tắt có dấu chấm). Có cache trên đối tượng Sheet.
def sheet_labels(sh):
    """(tập nhãn chính xác, {khoá nới lỏng: {nhãn gốc}}) của 1 sheet; cache trên đối tượng Sheet."""
    cached = getattr(sh, "_labels", None)
    if cached is not None:
        return cached
    exact, loose, dotted = set(), {}, set()
    def add(v):
        if isinstance(v, str):
            k = _lab_key(v)
            if k and _try_float(k) is None:
                exact.add(k)
                loose.setdefault(_lab_loose(k), set()).add(k)
                dotted.update(w for w in k.split(" ") if len(w) > 1 and w.endswith(".") and not w[:-1].isdigit())
    for row in sh.grid:
        for v in row:
            add(v)
    for t in sh.tables:
        for _, title in t["titles"]:
            add(title)
    sh._labels = (exact, loose, dotted)
    return sh._labels

# build_labels(sheets): gộp nhãn của tất cả sheet trong workbook.
def build_labels(sheets):
    exact, loose, dotted = set(), {}, set()
    for sh in sheets.values():
        e, l, d = sheet_labels(sh)
        exact |= e
        dotted |= d
        for k, v in l.items():
            loose.setdefault(k, set()).update(v)
    return exact, loose, dotted

# _label_restore(s, q, labels): nếu chuỗi s khớp một nhãn trong workbook (chính xác, hoặc khớp lỏng và DUY NHẤT)
# thì trả đúng NHÃN GỐC (giữ dấu chấm/phẩy); ngược lại None. Bỏ chú thích (1) ở cuối.
def _label_restore(s, q, labels):
    """Nếu chuỗi s khớp một nhãn trong workbook → trả đúng nhãn gốc (giữ dấu chấm/phẩy); ngược lại None."""
    if not labels or not isinstance(s, str):
        return None
    k = _lab_key(s)
    if not k or _try_float(k) is not None:
        return None
    exact, loose = labels[0], labels[1]
    hit = k if k in exact else None
    if hit is None:
        c = loose.get(_lab_loose(k))
        if c and len(c) == 1 and k.casefold() not in _lab_key(q or "").casefold():
            hit = next(iter(c))
    if hit is None:
        return None
    return _FOOTNOTE_TAIL.sub("", hit) or hit

# _split_labels(s, q, labels): tách 'a, b, c' thành các phần, nhưng GHÉP LẠI những mảnh thật ra là MỘT nhãn có dấu
# phẩy bên trong.
# Ví dụ nếu 'Smith, John' là một nhãn thật thì 'Smith, John, Doe' -> ['Smith, John', 'Doe'].
def _split_labels(s, q, labels, max_join=6):
    """Tách 'a, b, c' thành các phần, nhưng ghép lại các mảnh cùng tạo thành MỘT nhãn có dấu phẩy"""
    toks = re.split(r",\s+", s.strip())
    if not labels:
        return toks
    out, i, n = [], 0, len(toks)
    while i < n:
        merged = None
        for j in range(min(n, i + max_join), i + 1, -1):
            cand = ", ".join(toks[i:j])
            if _label_restore(cand, q, labels) is not None:
                merged = (cand, j)
                break
        if merged:
            out.append(merged[0])
            i = merged[1]
        else:
            out.append(toks[i])
            i += 1
    return out

# _is_abbrev_token(tok, dotted): token cuối có phải TỪ VIẾT TẮT có dấu chấm không (L.L.C., U.S., chữ cái đầu 'J.')
# để không xoá nhầm dấu chấm cuối của nó.
def _is_abbrev_token(tok, dotted=()):
    """Token cuối là từ viết tắt có dấu chấm? Nhận biết tổng quát bằng cấu trúc (L.L.C., U.S., chữ cái đầu 'J.')"""
    t = tok.strip("()[]\"'“”")
    if not t.endswith("."):
        return False
    body = t[:-1]
    if re.search(r"[^\W\d_]\.[^\W\d_]", t):                    # có dấu chấm xen giữa: L.L.C. / U.S. / e.g.
        return True
    if len(body) == 1 and body.isalpha() and body.isupper():       # chữ cái đầu: J.
        return True
    return t in dotted
# Regex nhận yêu cầu 'không kèm dấu %' (bằng tiếng Anh, Việt, Trung, Nhật, Hàn).
_NO_PCT = re.compile(r"(?:without|no|omit)\s+(?:the\s+)?%|(?:không|khong)\s+(?:kèm|có|co|kem)\s+(?:dấu\s+)?%|不(?:带|含|要)\s*%|%\s*なし|%\s*없이")

# _post_scalar(x, q, spec, labels): chuẩn hoá MỘT giá trị đáp án theo yêu cầu của câu hỏi
# (ngày, nhãn gốc, làm tròn, Yes/No, bỏ %, bỏ tiền tố...).
def _post_scalar(x, q, spec=None, labels=None):
    q = _nfc(q)
    # Giá trị rỗng -> chuỗi rỗng.
    if x is None:
        return ""
    # bool -> 'Yes'/'No'.
    if isinstance(x, bool):
        return "Yes" if x else "No"
    # Ngày -> định dạng theo yêu cầu trong câu hỏi.
    if isinstance(x, (dt.datetime, dt.date)):
        return _fmt_date_answer(x, q)
    # Chuỗi ngày ISO lúc nửa đêm -> rút gọn thành ngày theo định dạng yêu cầu.
    if isinstance(x, str):
        _m = _ISO_MIDNIGHT.match(x.strip())
        if _m:
            return _fmt_date_answer(dt.datetime.strptime(_m.group(1), "%Y-%m-%d"), q)
    # Chuỗi khớp một nhãn trong workbook -> trả đúng nhãn gốc.
    if isinstance(x, str):
        _hit = _label_restore(x, q, labels)
        if _hit is not None:
            return _hit
    # Lấy yêu cầu làm tròn từ câu hỏi (mặc định: half_up, chưa rõ số chữ số).
    _sp = rounding_specs(q)
    mode, places = spec if spec is not None else (_sp[0] if _sp else ("half_up", None))
    want_int = _wants_integer(q)
    ql = (q or "").lower()
    # Đáp án là số: có số chữ số thập phân yêu cầu thì làm tròn; ngược lại, số nguyên thì in dạng nguyên, số lẻ
    # thì in %.10g.
    if isinstance(x, (int, float)):
        v = float(x)
        if math.isnan(v) or math.isinf(v):
            return ""
        if places is not None:
            return _fmt_round(v, places, mode)
        if want_int or abs(v - round(v)) < 1e-12:
            return str(int(round(v)))
        return f"{v:.10g}"
    # Đáp án dạng chuỗi: gom khoảng trắng, bỏ tiền tố như 'Answer:'.
    s = re.sub(r"\s+", " ", str(x)).strip()
    s = _PREFIX_JUNK.sub("", s).strip()
    # Bỏ dấu chấm cuối câu (vd '42.' -> '42'), nhưng GIỮ dấu chấm của từ viết tắt.
    if re.fullmatch(r"[-+]?\d+\.", s):
        s = s[:-1]
    elif s.endswith(".") and not re.search(r"\d+\.\d", s):
        if not _is_abbrev_token(s.split(" ")[-1], labels[2] if labels else ()):  # giữ dấu chấm của từ viết tắt
            s = s[:-1].strip()
    # Đáp án Yes/No/True/False: viết theo đúng cách câu hỏi dùng ('yes or no' hay 'true/false').
    if re.fullmatch(r"(?i)yes|no|true|false", s):
        if re.search(r"yes\s*(or|/)\s*no", ql) or "yes or no" in ql:
            return "Yes" if s.lower() in ("yes", "true") else "No"
        if re.search(r"\btrue\b|\bfalse\b", ql):
            return "True" if s.lower() in ("yes", "true") else "False"
    # Câu hỏi yêu cầu không kèm % -> bỏ dấu % ở cuối.
    if _NO_PCT.search(ql) \
            and re.fullmatch(r"[-+]?[\d.,]+\s*%", s):
        s = s.rstrip("%").strip()
    # Chuỗi là số (và không phải mã bắt đầu bằng 0 như '007') -> làm tròn / định dạng như số.
    v = _try_float(s)
    if v is not None and not re.match(r"^[-+]?0\d", s):
        if places is not None:
            return _fmt_round(v, places, mode)
        if not want_int and re.fullmatch(r"[-+]?\d+\.\d{1,6}", s) and s.endswith("0"):
            return s
        if want_int or abs(v - round(v)) < 1e-12:
            return str(int(round(v)))
        return f"{v:.10g}" if abs(v - round(v)) >= 1e-12 else str(int(round(v)))
    # Không phải số -> là chuỗi nhãn; chỉ bỏ chú thích '(1)' ở cuối.
    return re.sub(r"\s*\(\d+\)\s*$", "", s)

# postprocess(ans, q, labels): CỔNG CHUẨN HOÁ CUỐI CÙNG của đáp án.
# Nhiều yêu cầu làm tròn -> _post_multi | danh sách -> xử lý từng phần tử, nối bằng ', ' | chuỗi -> bóc
# <FINAL_ANSWER>, khớp nhãn, tách theo dấu phẩy | còn lại -> _post_scalar.
def postprocess(ans, q, labels=None):
    # Không có đáp án -> chuỗi rỗng.
    if ans is None:
        return ""
    # Từ 2 yêu cầu làm tròn trở lên -> xử lý riêng bằng _post_multi.
    specs = rounding_specs(q)
    if len(specs) >= 2:
        out = _post_multi(ans, specs)
        if out is not None:
            return out
    # Đáp án là danh sách -> chuẩn hoá từng phần tử rồi nối bằng ', '.
    if isinstance(ans, (list, tuple)):
        return ", ".join(p for p in (_post_scalar(a, q, labels=labels) for a in ans) if p != "")
    # Đáp án là chuỗi: bỏ % nếu cần, bóc <FINAL_ANSWER>, thử khớp CẢ chuỗi là một nhãn (có thể chứa dấu phẩy), nếu
    # có dấu phẩy thì tách từng phần.
    if isinstance(ans, str):
        if _NO_PCT.search(_nfc(q).lower()) and re.fullmatch(r"\s*[-+]?[\d.,]+\s*%\s*", ans):  # "không kèm %" → bỏ % trước khi khớp nhãn
            ans = ans.strip().rstrip("%").strip()
        m = re.search(r"<FINAL_ANSWER>\s*(.*?)\s*</FINAL_ANSWER>", ans, flags=re.S | re.I)
        if m:
            ans = m.group(1).strip()
        whole = _label_restore(ans, q, labels)  # cả chuỗi là MỘT nhãn (có thể chứa dấu phẩy)
        if whole is not None:
            return whole
        if re.search(r",\s+", ans):
            parts = [_post_scalar(p, q, labels=labels) for p in _split_labels(ans, q, labels)]
            return ", ".join(p for p in parts if p != "")
    return _post_scalar(ans, q, labels=labels)

# =============================================================================
# 4b. TRA CỨU Ô THEO ĐỊA CHỈ (câu hỏi "ô D5 chứa gì / công thức nào")
# =============================================================================
# Regex nhận biết câu hỏi hỏi về CÔNG THỨC (công thức / 公式 / 数式 / 수식 / formule...).
_FORMULA_Q = re.compile(r"formula|công\s*thức|cong\s*thuc|公式|数式|數式|수식|fórmula|formule", re.I)
# Regex bắt địa chỉ ô trong câu hỏi (vd 'cell D5', 'ô D5', '单元格 D5', 'Sheet1!D5'); loại trừ vùng A1:B2.
_QREF_RE = re.compile(
    r"(?:\bcells?\b|\bô\b|单元格|單元格|儲存格|储存格|セル|셀|!)\s*[:：]?\s*['\"“‘]?\$?([A-Za-z]{1,3})\$?(\d{1,7})(?![\w])"
    r"(?!\s*:\s*\$?[A-Za-z]{1,3}\$?\d)", re.I)

# _query_refs(query): trả danh sách địa chỉ ô nêu trong câu hỏi, dạng [('D', 5), ...]; không trùng lặp, bỏ cột >
# XFD hoặc hàng < 1.
def _query_refs(query):
    out = []
    for m in _QREF_RE.finditer(query or ""):
        col, row = m.group(1).upper(), int(m.group(2))
        try:
            if column_index_from_string(col) > 16384 or row < 1:
                continue
        except Exception:
            continue
        if (col, row) not in out:
            out.append((col, row))
    return out

# _query_sheet_names(query, sheets): các tên sheet được nhắc tới trong câu hỏi.
def _query_sheet_names(query, sheets):
    ql = (query or "").lower()
    return [k for k in sheets if k and k.lower() in ql]

# _formulaish(t): chấm điểm mức độ 'giống công thức': 6 = công thức thật ('=...'); 2/1 = văn bản trông như công
# thức; 0 = không.
def _formulaish(t):
    """0 = không giống công thức; 6 = công thức thật ('=...'); 2/1 = chuỗi chữ trông như công thức."""
    if not isinstance(t, str):
        return 0
    s_ = t.strip()
    if s_.startswith("="):
        return 6
    if "=" in s_ and re.search(r"[\w\)\]]\s*[-+*/^]\s*[\w\(\[]", s_):
        return 2
    if re.search(r"(?:\(\w{1,3}\)|\b[A-Za-z]{1,3}\$?\d+\b)\s*[-+*/^]\s*(?:\(\w{1,3}\)|\b[A-Za-z]{1,3}\$?\d+\b)", s_):
        return 1
    return 0

# _label_match(t, letter): ô có bắt đầu bằng nhãn trùng chữ cột không (vd '(d)', 'D:', 'd = ...') - kiểu cột đánh
# nhãn (a), (b), (c)...
def _label_match(t, letter):
    """Ô bắt đầu bằng nhãn trùng chữ cột, vd '(d)', '(d) = ...', 'D:' (kiểu cột đánh nhãn (a),(b),(c)...)."""
    return isinstance(t, str) and bool(re.match(rf"^\s*\(?{re.escape(letter)}\)?\s*(?:[=:.\-–—]|$)", t, re.I))

# _fmt_raw(v, n=90): chuỗi hiển thị gọn: bỏ xuống dòng, cắt tối đa n ký tự (thêm '…'), bọc bằng repr.
def _fmt_raw(v, n=90):
    t = str(v).replace("\n", " ")
    return repr(t if len(t) <= n else t[:n] + "…")

# _empty_cell_candidates(...): khi ô được hỏi TRỐNG, xếp hạng các ô cùng hàng theo mức khớp (giống công thức /
# nhãn trùng chữ cột).
# Trả danh sách (điểm, địa chỉ, nội dung) giảm dần theo điểm.
def _empty_cell_candidates(sh, col, row, formula_intent):
    """Khi ô được hỏi TRỐNG: xếp hạng các ô cùng dòng theo mức khớp (công thức/nhãn trùng chữ cột)."""
    res = []
    for letter, v in sh.raw_row(row).items():
        sc = (_formulaish(v) if formula_intent else 0) + (3 if _label_match(v, col) else 0)
        if sc > 0:
            res.append((sc, f"{letter}{row}", v))
    res.sort(key=lambda x: (-x[0], column_index_from_string(re.match(r"[A-Z]+", x[1]).group(0))))
    return res

# cell_lookup_text(query, sheets): dựng khối bằng chứng 'CELL LOOKUP' đưa vào prompt cho mọi câu hỏi nêu địa chỉ ô
# cụ thể.
# Đọc nội dung thật của ô từ file (nguồn đáng tin cậy nhất).
def cell_lookup_text(query, sheets):
    """Khối bằng chứng đưa vào prompt cho mọi câu hỏi có nêu địa chỉ ô cụ thể ('cell D5')."""
    # Tính năng bị tắt -> không thêm gì vào prompt.
    if not CELL_FALLBACK:
        return ""
    # Lấy các địa chỉ ô trong câu hỏi; không có ô nào thì thôi.
    refs = _query_refs(query)
    if not refs:
        return ""
    # Sheet được nhắc tên trong câu hỏi; nếu không có thì lấy tối đa vài sheet đầu.
    names = _query_sheet_names(query, sheets) or list(sheets)[:4]
    # fq: câu hỏi có đang hỏi về CÔNG THỨC không.
    fq = bool(_FORMULA_Q.search(query))
    # Bắt đầu khối văn bản bằng chứng.
    lines = ["CELL LOOKUP (exact addresses named in the question, read from the raw file; authoritative):"]
    # Tối đa 3 ô x 4 sheet để prompt không quá dài.
    for col, row in refs[:3]:
        for name in names[:4]:
            sh = sheets[name]
            addr = f"{col}{row}"
            v = sh.raw(addr)
            # Ô TRỐNG: liệt kê các ô không trống cùng hàng/cùng cột và các ứng viên gần đúng nhất (có điểm) để LLM
            # tham khảo.
            if v is None:
                lines.append(f"- {name}!{addr}: EMPTY (no value, no formula)")
                rc = sh.raw_row(row)
                if rc:
                    lines.append(f"  non-empty cells in row {row}: " + " | ".join(f"{k}={_fmt_raw(x)}" for k, x in list(rc.items())[:25]))
                cc = sh.raw_col(col)
                if cc:
                    lines.append(f"  non-empty cells in column {col}: " + " | ".join(f"{col}{k}={_fmt_raw(x)}" for k, x in list(cc.items())[:8]))
                cands = _empty_cell_candidates(sh, col, row, fq)
                if cands:
                    lines.append("  closest matches in the same row (score): " +
                                 " ; ".join(f"{a}={_fmt_raw(x)} [{sc}]" for sc, a, x in cands[:4]))
            # Ô là CÔNG THỨC: ghi công thức kèm giá trị đã tính.
            elif isinstance(v, str) and v.startswith("="):
                lines.append(f"- {name}!{addr}: FORMULA {_fmt_raw(v, 300)}  (computed value: {_fmt_raw(sh.v(addr))})")
            else:
                lines.append(f"- {name}!{addr}: {type(v).__name__.upper()} {_fmt_raw(v, 300)}")
    return "\n".join(lines) + "\n\n"

# try_direct_cell(query, sheets): nếu câu hỏi CHỈ hỏi công thức của ĐÚNG MỘT ô nêu rõ địa chỉ -> trả lời thẳng,
# không cần LLM.
# Trả (đáp án, nhãn nguồn) hoặc None nếu không chắc chắn.
def try_direct_cell(query, sheets):
    """Câu hỏi chỉ đọc công thức của MỘT ô nêu rõ địa chỉ → trả lời thẳng, không cần LLM."""
    # Chỉ áp dụng khi tính năng bật VÀ câu hỏi hỏi về công thức.
    if not CELL_FALLBACK or not _FORMULA_Q.search(query or ""):
        return None
    refs = _query_refs(query)
    # Phải có đúng 1 địa chỉ ô; nhiều hơn thì để LLM xử lý.
    if len(refs) != 1:
        return None
    names = _query_sheet_names(query, sheets)
    # Mơ hồ về sheet (nhiều sheet được nhắc, hoặc workbook nhiều sheet mà không nêu tên) -> để LLM xử lý.
    if len(names) > 1 or (not names and len(sheets) > 1):
        return None
    sh = sheets[names[0]] if names else next(iter(sheets.values()))
    col, row = refs[0]
    addr = f"{col}{row}"
    v = sh.raw(addr)
    # Ô là công thức thật -> trả nguyên công thức.
    if isinstance(v, str) and v.startswith("="):
        return v, "formula"
    # Ô là văn bản trông như công thức -> trả nguyên văn bản.
    if isinstance(v, str) and _formulaish(v) >= 2:
        return v, "text-formula"
    if v is None:  # ô được hỏi trống → ô cùng dòng khớp rõ ràng nhất (phải vượt trội, không đoán khi mơ hồ)
        c = _empty_cell_candidates(sh, col, row, True)
        if c and c[0][0] >= 3 and (len(c) == 1 or c[0][0] - c[1][0] >= 2):
            return str(c[0][2]), f"empty->{c[0][1]}"
    return None

# _Done: ngoại lệ dùng để THOÁT SỚM khỏi solve_one khi đã có đáp án (không phải lỗi thật).
class _Done(Exception):
    pass
# SYSTEM_PROMPT: lời nhắc hệ thống - hướng dẫn LLM viết một đoạn Python gán biến `answer`.
# Gồm: cách đọc bản dump, API Python có sẵn, quy tắc ô theo địa chỉ, quy tắc đa ngôn ngữ, quy tắc cấu trúc bảng,
# quy tắc định dạng đáp án.
# (Nội dung giữ nguyên tiếng Anh vì LLM hiểu tốt hơn; sửa nội dung này sẽ ảnh hưởng trực tiếp đến kết quả.)
SYSTEM_PROMPT = """You answer ONE question about an Excel workbook by writing a SHORT Python snippet. The snippet is executed; its variable `answer` is the final answer. Never compute in your head. Never translate labels.

WHAT YOU SEE
Text dump of every sheet: TITLE / UNITS, COLUMN HEADERS ("letter: path" with " > "), ROWS as "R<row> | A=val | B=val ...". Empty cells omitted. Merged cells expanded.
Header rows are shown both as header paths and as "HEADER ROWS" (raw cells by address). If the question names an exact cell (e.g. "cell D5"), a "CELL LOOKUP" block lists what that address really holds.

PYTHON API (do not import)
sh, SHEETS["name"], sh.v("B5"), sh.row(r), sh.col("B"),
sh.find_rows(text, col="A", exact=False), sh.find_cols(text),
num(x) → float|None (None for blank/-/n/a/#N/A/… — NEVER treat missing as 0),
to_date(x), modules: math, re, statistics, dt.
sh.raw("B5") → original content of the cell exactly as stored (a formula keeps its "=..." text; empty → None),
sh.formula("B5") → formula text exactly as written, or None if the cell is not a formula, sh.raw_row(r), sh.raw_col("B").

CELL-ADDRESS QUESTIONS ("what is in / which formula is in cell X")
- Read the exact address with sh.raw(); never answer from a different cell just because it looks similar.
- "Formula in cell X": real formula → sh.formula() exactly as written (keep "="); text cell → its COMPLETE text, never a header-path fragment or a part of it.
- If the addressed cell is EMPTY, use the CELL LOOKUP candidates: take the same-row/column cell whose content is what the question asks for (formula-like text; a label equal to the column letter such as "(d)") and return that cell's COMPLETE text. If nothing fits, answer with the empty cell's value (empty string).

MULTILINGUAL — sheet AND question may be EN/VI/ZH/JA/KO/FR
- Understand the question in any language; map its meaning to the sheet's labels even when the languages differ.
- Label answers: copy EXACT sheet string (keep its punctuation, e.g. abbreviation dots and commas inside names), do not translate.
- Format requests may be multilingual — obey them:
  "two decimal places" / "2 chữ số thập phân" / "保留2位小数" / "소수점 2자리" / "小数点以下2桁".

STRUCTURE
1. Match FULL header path + row label (not one ambiguous word like "Total").
2. Blocks: use the block the question names; else overall/TOTAL.
3. of which / trong đó / 其中 = parts of parent (do not double-count). Total rows are aggregates — use an existing Total, do not re-sum for the same quantity. For rank/max/min/count among categories, exclude Total/aggregate rows and columns.
4. Sums, counts, ranks, max/min over a group: use lowest-level columns, not aggregate siblings. AVERAGE / MEAN "across all X" (all levels, all categories, all years...) is different: take EVERY data cell on the axis the question names, exactly as the table lists them, subtotal columns/rows included (do not drop, merge or re-group cells by your own judgment), and take their simple arithmetic mean. Exclude a cell only if the question explicitly says to exclude totals/subtotals, or it is the single overall grand total of that whole axis. Never mix cells from different blocks or dimensions unless the question asks for it.
5. Orientation: answer may be column header text OR row label.
6. Missing values: ignore in sum/avg/min/max (`num(v) is not None`).
6b. If a sheet is split into "TABLE 1, TABLE 2 ...", each table has its own COLUMN HEADERS; the same column letter can mean different things per table. Use the table the question refers to; use R<row> numbers to stay inside that table's row range.

ANSWER FORMAT (Exact Match) — read the question carefully
7. Rounding is MANDATORY when asked:
   - Compute full precision, THEN format.
   - N decimal places → answer = f"{x:.Nf}" (show trailing zeros: 3 → "3.00").
   - Integer / số nguyên / 整数 → int, no decimal point.
   - Set `answer` to the formatted value, not a raw float.
   - EXCEPTION: if the message ends with an "OUTPUT FORMAT NOTE" (several roundings requested), follow that note instead.
8. Yes/No as the question requests. Labels: exact sheet string. EXCEPTION: if the question itself lists the items to output, write them exactly as worded in the QUESTION, not the sheet's longer label.
9. No commentary, no unit suffix unless required.
10. NEVER rescale units. Even if the title says "in thousands/millions", return the number exactly as stored in the cell, unless the question explicitly asks to convert.

Assign `answer`. Return exactly one ```python block, no explanation."""
# REPAIR_PROMPT: lời nhắc gửi khi code lỗi hoặc đáp án rỗng/không hợp lệ - kèm code cũ và mô tả lỗi để LLM sửa.
REPAIR_PROMPT = """The previous code failed or produced an empty/invalid answer.
Problem: {problem}
Code:
```
{code}
```
Common causes: the label sits in a DIFFERENT sheet/table, is spelled differently, or is in another column than assumed. Search ALL sheets (SHEETS.items()), try shorter keywords, scan every column. Never set `answer` to an error message or placeholder text.
Fix it. Return exactly one ```python block, same rules."""
# DIRECT_SYSTEM: lời nhắc cho phương án CUỐI CÙNG - hỏi LLM trả lời thẳng (không viết code) khi mọi phiếu đều thất
# bại.
DIRECT_SYSTEM = ("You answer questions about an Excel sheet dump. Think privately, then output ONLY the final answer in "
                 "the exact format the question asks for. No explanation, no markdown.")
# JUDGE_SYSTEM: lời nhắc cho TRỌNG TÀI - so sánh các đáp án ứng viên cùng code sinh ra chúng; giữ đáp án đa số trừ
# khi thấy lỗi rõ ràng.
# Kết thúc bằng một dòng JSON {'choice': chỉ số}.
JUDGE_SYSTEM = """You are an arbitrator for Excel analysis answers. Several candidate answers exist for the same question; each comes with the Python code that produced it.
Vote counts are only a hint: a majority can still be wrong if every vote read the wrong column/row.
Cross-check against the sheet dump (each cell shows its column letter):
- Does each code use the cell/column/row whose label matches the question (multi-level headers, value type, units)?
- Is the filter / ranking / arithmetic logic correct for the question?
- Is the format correct?
- Never prefer an answer that rescales units (e.g. x1000 because the title says "in thousands") unless the question explicitly asks for it.
- Keep the majority answer unless you find a clear error in its column/row/logic/format.
Give a very short rationale (under 80 words), then a final line that is ONLY JSON: {"choice": <0-based candidate index>}"""
# Cờ cho biết có gửi tham số TẮT reasoning hay không (dùng dict để có thể sửa bên trong hàm).
_reasoning_ok = {"v": REASONING_OFF}

# Budget: bộ đếm token, chi phí, số lần gọi LLM của MỘT câu hỏi (có khoá vì nhiều phiếu chạy song song).
class Budget:
    # Khởi tạo các bộ đếm về 0.
    def __init__(self):
        self.tin = self.tout = 0
        self.cost = 0.0
        self.calls = 0
        self._lock = threading.Lock()
    # add(tin, tout, cost): cộng dồn token vào/ra và chi phí, an toàn đa luồng.
    def add(self, tin, tout, cost):
        with self._lock:
            self.tin += tin
            self.tout += tout
            self.cost += cost
            self.calls += 1

# llm_chat(messages, bud, deadline, temperature, max_tokens): gọi API chat của LLM, trả về nội dung văn bản.
# Có thử lại khi lỗi, kiểm soát hạn chót và trần chi phí, tự ghi nhận token/chi phí vào `bud`.
def llm_chat(messages, bud, deadline, temperature=0, max_tokens=None):
    # Lấy khoá API từ biến môi trường; thiếu khoá -> báo lỗi.
    key = os.environ.get("OPENROUTER_API_KEY") or os.environ.get("OPENAI_API_KEY") or ""
    if not key:
        raise RuntimeError("Thieu OPENROUTER_API_KEY")
    last = None
    tok_limit = max_tokens or MAX_TOKENS
    # Thử lại tối đa LLM_RETRIES lần.
    for attempt in range(LLM_RETRIES):
        # Còn dưới 10 giây tới hạn chót -> dừng hẳn (ném TimeoutError).
        remaining = deadline - time.time()
        if remaining < 10:
            raise TimeoutError("het ngan sach thoi gian")
        # Chi phí của câu hỏi đã vượt trần -> dừng.
        if bud.cost >= MAX_COST_PER_Q:
            raise RuntimeError("vuot ngan sach chi phi cau hoi")
        # Nội dung yêu cầu gửi tới API.
        body = {"model": MODEL, "messages": messages, "temperature": temperature, "max_tokens": tok_limit}
        # Nếu cấu hình tắt reasoning -> gửi kèm tham số tắt.
        if _reasoning_ok["v"]:
            body["reasoning"] = {"enabled": False}
        # Gửi yêu cầu HTTP; lỗi mạng -> chờ (tăng dần) rồi thử lại.
        try:
            r = requests.post(API_URL, headers={"Authorization": f"Bearer {key}", "Content-Type": "application/json"},
                              json=body, timeout=min(HTTP_TIMEOUT, remaining))
        except requests.RequestException as e:
            last = e
            time.sleep(1.5 * (attempt + 1))
            continue
        # Bị giới hạn tốc độ hoặc máy chủ quá tải -> chờ lâu hơn rồi thử lại.
        if r.status_code in (429, 502, 503):
            last = RuntimeError(f"HTTP {r.status_code}")
            time.sleep(2 * (attempt + 1))
            continue
        # Mã lỗi HTTP khác -> chờ 1 giây rồi thử lại.
        if r.status_code != 200:
            last = RuntimeError(f"HTTP {r.status_code}: {r.text[:300]}")
            time.sleep(1)
            continue
        # Đọc phản hồi JSON; lỗi định dạng -> thử lại.
        try:
            j = r.json()
        except Exception as e:
            last = e
            continue
        # Ghi nhận số token và chi phí của lần gọi này vào Budget.
        usage = j.get("usage") or {}
        tin = int(usage.get("prompt_tokens") or 0)
        tout = int(usage.get("completion_tokens") or 0)
        cost = tin * PRICE_IN / 1e6 + tout * PRICE_OUT / 1e6
        bud.add(tin, tout, cost)
        # Lấy nội dung câu trả lời của mô hình.
        msg = (j.get("choices") or [{}])[0].get("message") or {}
        text = msg.get("content") or ""
        # Xoá phần 'suy nghĩ' <think>...</think> (nếu có) rồi trả nội dung sạch.
        return re.sub(r"<think>.*?</think>", "", text, flags=re.S).strip()
    raise RuntimeError(f"LLM loi sau nhieu lan thu: {last}")

# extract_code(text): lấy code Python từ câu trả lời của LLM - chọn khối ```python DÀI NHẤT;
# không có khối nào mà văn bản có chữ 'answer' thì coi cả văn bản là code; còn lại trả chuỗi rỗng.
def extract_code(text):
    m = re.findall(r"```(?:python|py)?\s*\n(.*?)```", text, flags=re.S | re.I)
    if m:
        return max(m, key=len).strip()
    return text.strip() if "answer" in text else ""

# =============================================================================
# 6. GIẢI MỘT CÂU HỎI (bỏ phiếu + trọng tài)
# =============================================================================
# resolve_file(data_dir, fname): tìm file theo tên trong thư mục (rồi đệ quy các thư mục con); không thấy ->
# FileNotFoundError.
def resolve_file(data_dir, fname):
    p = os.path.join(data_dir, fname)
    if os.path.exists(p):
        return p
    for root, _, files in os.walk(data_dir):
        if fname in files:
            return os.path.join(root, fname)
    raise FileNotFoundError(fname)
# Regex nhận diện đáp án là THÔNG BÁO LỖI do mô hình tự viết ('Error', 'not found', 'lỗi', 'không tìm thấy',
# '无法'...), nhiều ngôn ngữ.
_FAIL_RE = re.compile(
    r"^\W*(?:(?:error|exception|not\s+found|n/a|cannot|can't|unable|no\s+data|insufficient"
    r"|missing\s+(?:data|value|label|row|column))\b"
    r"|lỗi|không\s+tìm\s+thấy|thiếu\s+dữ\s+liệu|错误|錯誤|无法|無法|未找到|エラー|見つかりません|오류)",
    re.I)

# _bad(ans, sheets): đáp án có KHÔNG DÙNG ĐƯỢC không? Có nếu: rỗng/NaN/'None'/nhiều dòng (văn xuôi) hoặc là chuỗi
# báo lỗi.
# Ngoại lệ: chuỗi trông như báo lỗi (vd 'N/A') nhưng đúng là giá trị có thật trong workbook thì giữ lại.
def _bad(ans, sheets=None):
    """Đáp án không dùng được: rỗng/NaN/nhiều dòng (văn xuôi), hoặc là chuỗi báo lỗi do mô hình tự viết.
    Ngoại lệ: chuỗi trông như báo lỗi (vd "N/A") nhưng đúng là giá trị có thật trong workbook thì được giữ."""
    if ans is None or (isinstance(ans, float) and math.isnan(ans)):
        return True
    t = str(ans).strip()
    if t in ("", "None", "nan") or "\n" in t:
        return True
    if REJECT_ERROR_ANSWERS and isinstance(ans, str) and _FAIL_RE.match(t):
        if sheets and _lab_key(t) in build_labels(sheets)[0]:
            return False
        return True
    return False

# vote_key(s): khoá để SO SÁNH các phiếu: số -> làm tròn 6 chữ số; danh sách số -> tuple; chữ -> hạ chữ thường +
# gom khoảng trắng.
# Nhờ vậy '3.0' và '3' được tính là cùng một phiếu.
def vote_key(s):
    """Khoá so sánh đáp án: số → giá trị làm tròn 6 chữ số; list số; chữ → lower + gộp khoảng trắng."""
    s = str(s).strip()
    try:
        return ("n", round(float(s), 6))
    except ValueError:
        pass
    parts = s.split(",")
    if len(parts) > 1 and all(re.fullmatch(r"\s*[-+]?\d+(\.\d+)?\s*", p) for p in parts):
        return ("l", tuple(round(float(p), 6) for p in parts))
    return ("s", re.sub(r"\s+", " ", s.lower()))

# gen_candidate(...): sinh MỘT 'phiếu' = LLM viết code -> chạy -> (nếu lỗi) sửa 1 lần.
# Trả dict {answer, code, raw, path ('code'/'repair'/'none'), err}.
def gen_candidate(user, sheets, query, bud, deadline, temperature, user_full=None, labels=None):
    """Một bản: sinh code → chạy → (nếu lỗi) sửa 1 lần. Trả dict {answer, code, path, err, raw}."""
    # Hội thoại gửi LLM: lời nhắc hệ thống + câu hỏi kèm dữ liệu.
    msgs = [{"role": "system", "content": SYSTEM_PROMPT}, {"role": "user", "content": user}]
    # code: code đã sinh | raw: đáp án thô | problem: mô tả lỗi (None nếu thành công).
    code, raw, problem = "", None, "no attempt"
    user_retry = user_full if (FULL_DUMP_ON_RETRY and user_full) else user  # lần repair dùng dump đầy đủ
    # Tối đa 2 lượt: lượt 1 = sinh code, lượt 2 = sửa lỗi (repair).
    for attempt in range(2):
        text = llm_chat(msgs, bud, deadline, temperature)
        code = extract_code(text)
        raw = None
        if not code:
            problem = "No Python code block found."
        else:
            # Chạy code trong sandbox.
            raw, err = run_code(code, sheets)
            # Có lỗi khi chạy, hoặc đáp án rỗng/không hợp lệ -> ghi lại vấn đề; ngược lại problem = None.
            problem = err if err else ("`answer` is empty, invalid, or an error message (a needed value was not found)" if _bad(raw, sheets) else None)
        # Thành công -> chuẩn hoá đáp án (postprocess) và trả về phiếu.
        if not problem:
            return {"answer": postprocess(raw, query, labels), "code": code, "raw": str(raw),
                    "path": "code" if attempt == 0 else "repair", "err": None}
        # Chuẩn bị lượt sửa lỗi: gửi lại hội thoại cũ + code cũ + mô tả vấn đề (REPAIR_PROMPT).
        msgs = [{"role": "system", "content": SYSTEM_PROMPT}, {"role": "user", "content": user_retry},
                {"role": "assistant", "content": text},
                {"role": "user", "content": REPAIR_PROMPT.format(problem=problem, code=code or "(none)")}]
    # Cả hai lượt đều thất bại -> trả phiếu RỖNG kèm lý do lỗi.
    return {"answer": "", "code": code, "raw": None if raw is None else str(raw), "path": "none", "err": problem}

# judge_votes(...): TRỌNG TÀI - khi các phiếu lệch nhau, nhờ LLM chọn ứng viên đúng nhất dựa trên code và bảng dữ
# liệu.
# Trả đáp án được chọn, hoặc None nếu không chọn được.
def judge_votes(query, dump, cands, bud, deadline, trace=None):
    """Trọng tài: khi phiếu lệch, nhờ LLM chọn ứng viên đúng nhất dựa trên code + bảng."""
    # seen: các khoá phiếu đã gặp | body_parts: nội dung mô tả từng ứng viên gửi cho trọng tài.
    seen, body_parts = set(), []
    unique = []
    # Gom các đáp án KHÁC NHAU (bỏ trùng theo vote_key), kèm số phiếu và code sinh ra đáp án đó.
    for c in cands:
        if not c.get("answer"):
            continue
        k = vote_key(c["answer"])
        if k in seen:
            continue
        seen.add(k)
        unique.append(c)
        n = sum(1 for x in cands if x.get("answer") and vote_key(x["answer"]) == k)
        body_parts.append(
            f"--- Candidate {len(unique) - 1} | answer: {c['answer']!r} | {n} vote(s) ---\n"
            f"{c.get('code') or ''}"
        )
    # Dưới 2 đáp án khác nhau thì không cần trọng tài.
    if len(unique) < 2:
        return None
    body = "\n\n".join(body_parts)
    view = dump
    # Nhờ LLM làm trọng tài (nhiệt độ 0 cho ổn định).
    text = llm_chat(
        [{"role": "system", "content": JUDGE_SYSTEM},
         {"role": "user", "content": f"{view}\n\nQUESTION: {query}\n\n{body}"}],
        bud, deadline, temperature=0,
    )
    if trace is not None:
        trace["text"] = text
    # Đọc chỉ số 'choice' cuối cùng trong câu trả lời; hợp lệ -> trả đáp án của ứng viên đó.
    ms = re.findall(r'"choice"\s*:\s*(\d+)', text)
    if ms and 0 <= int(ms[-1]) < len(unique):
        return unique[int(ms[-1])]["answer"]
    return None
# ROUTE_SYSTEM: lời nhắc cho bước CHỌN SHEET liên quan (routing) - nên chọn dư còn hơn thiếu; trả JSON {'sheets':
# [...]}.
ROUTE_SYSTEM = ("You select which sheets of an Excel workbook are needed to answer a question. You see a catalog "
                "(sheet name, size, column headers, sample rows). Include EVERY sheet that might be needed, including "
                "sheets needed for lookups/joins/units; when unsure, include more. Reply with ONLY JSON: "
                '{"sheets": ["<exact sheet name>", ...]}')

# _catalog(sheets): tạo DANH MỤC THU GỌN của workbook (tên sheet, kích thước, số bảng, tối đa 14 tiêu đề, 2 dòng
# mẫu) để LLM chọn sheet mà không phải đọc toàn bộ dữ liệu.
def _catalog(sheets):
    parts = []
    for name, s in sheets.items():
        t = s.tables[0] if s.tables else None
        hdr = list((t["headers"] if t else {}).values())[:14]
        samples = []
        for r in range(t["he"] if t else 0, s.nrows):
            cells = [f"{get_column_letter(c + 1)}={fmt_val(v)[:40]}" for c, v in enumerate(s.grid[r]) if v is not None]
            if cells:
                samples.append(f"R{r + 1} | " + " | ".join(cells[:8]))
            if len(samples) >= 2:
                break
        parts.append(f"- {name!r}: {s.nrows} rows x {s.ncols} cols"
                     + (f", {len(s.tables)} tables" if len(s.tables) > 1 else "")
                     + "\n  headers: " + " ; ".join(h[:80] for h in hdr)
                     + "\n  " + "\n  ".join(samples))
    return "\n".join(parts)

# route_sheets(...): nhờ LLM chọn các sheet cần dùng. Trả danh sách tên sheet (giữ thứ tự gốc) hoặc None
# nếu lỗi / rỗng / chọn đủ tất cả sheet (không rút gọn được gì).
def route_sheets(query, sheets, bud, deadline):
    """Chọn các sheet liên quan. Trả list tên sheet (giữ thứ tự gốc) hoặc None nếu lỗi/rỗng/không giảm được gì."""
    try:
        text = llm_chat([{"role": "system", "content": ROUTE_SYSTEM},
                         {"role": "user", "content": f"CATALOG:\n{_catalog(sheets)}\n\nQUESTION: {query}"}],
                        bud, deadline, temperature=0)
        m = re.search(r"\{.*\}", text, flags=re.S)
        names = json.loads(m.group(0)).get("sheets") if m else None
        if not isinstance(names, list):
            return None
        lower = {k.lower().strip(): k for k in sheets}
        picked = {lower[str(n).lower().strip()] for n in names if str(n).lower().strip() in lower}
        chosen = [k for k in sheets if k in picked]
        return chosen if chosen and len(chosen) < len(sheets) else None
    except Exception:
        return None

# solve_one(q, data_dir, deadline): ĐIỀU PHỐI việc giải MỘT câu hỏi: nạp file -> (đường tắt ô) -> routing -> dựng
# prompt -> bỏ phiếu -> trọng tài -> hậu xử lý.
# Trả bản ghi `rec` (đáp án, đường đi, token, chi phí, thời gian...).
def solve_one(q, data_dir, deadline):
    t0 = time.time()
    qid, fname = q["id"], q.get("file_name") or q.get("file name")
    query = q["query"]
    # rec: bản ghi kết quả và dữ liệu gỡ lỗi của câu hỏi này.
    rec = {"id": qid, "file": fname, "query": query, "path": "none", "answer": "", "raw": None, "code": "", "err": None,
           "votes": [], "judge": None}
    # Bộ đếm token/chi phí riêng cho câu hỏi này.
    bud = Budget()
    # Toàn bộ xử lý nằm trong try để lỗi của một câu không làm sập các câu khác.
    try:
        # Nạp workbook (có cache): sheets = dict các Sheet, dump = văn bản mô tả dữ liệu.
        sheets, dump = load_workbook_sheets(resolve_file(data_dir, fname))
        if not sheets:
            raise RuntimeError("file khong co du lieu")
        # ĐƯỜNG TẮT: câu hỏi chỉ đọc công thức của một ô cụ thể -> trả lời thẳng từ file, không gọi LLM.
        dc = try_direct_cell(query, sheets)
        if dc is not None:  # đọc thẳng công thức của đúng ô được hỏi, không qua LLM
            rec.update(answer=dc[0], raw=dc[0], path="cell:" + dc[1], err=None)
            raise _Done()
        full_dump = dump  # bản đầy đủ, dùng lại khi repair/direct
        labels = build_labels(sheets)  # nhãn gốc của mọi ô chữ, dùng để giữ nguyên dấu chấm/phẩy khi chuẩn hoá đáp án
        # ROUTING: workbook nhiều sheet và dump dài -> nhờ LLM chọn các sheet liên quan để rút ngắn prompt (tiết
        # kiệm token, ít nhiễu hơn).
        if ROUTING_ON and len(sheets) > 1 and len(dump) > ROUTE_MIN_CHARS:
            chosen = route_sheets(query, sheets, bud, deadline)
            if chosen:
                routed = "\n\n".join(sheets[k].dump() for k in chosen)
                if len(routed) < ROUTE_SHRINK_RATIO * len(dump):  # chỉ dùng khi thật sự giảm đáng kể
                    dump = routed
                    # Đưa các sheet được chọn lên ĐẦU dict (nên `sh` mặc định là sheet đầu tiên được chọn);
                    # các sheet khác vẫn được giữ lại để code vẫn truy cập được qua SHEETS.
                    sheets = {**{k: sheets[k] for k in chosen}, **{k: v for k, v in sheets.items() if k not in chosen}}
                    rec["routed"] = chosen
        # Khối 'CELL LOOKUP' nếu câu hỏi nêu địa chỉ ô cụ thể (rỗng nếu không có).
        hint = cell_lookup_text(query, sheets)
        # Nội dung gửi LLM: tên file + dữ liệu (dump) + gợi ý tra cứu ô + câu hỏi (+ ghi chú làm tròn nhiều giá
        # trị nếu có).
        user = f"FILE: {fname}\n\n{dump}\n\n{hint}QUESTION: {query}" + multi_round_hint(query)
        # Bản dùng dump ĐẦY ĐỦ (không routing) cho lượt sửa lỗi / trả lời trực tiếp; None nếu dump không bị rút
        # gọn.
        user_full = (f"FILE: {fname}\n\n{full_dump}\n\n{hint}QUESTION: {query}" + multi_round_hint(query)) if dump is not full_dump else None
        # cands: danh sách các phiếu | need: số phiếu tối thiểu để coi là đa số (VOTES=3 -> cần 2).
        cands, need = [], VOTES // 2 + 1
        # _try(i): sinh phiếu thứ i. Phiếu đầu (i=0) dùng nhiệt độ 0 cho ổn định, các phiếu sau dùng VOTE_TEMP để
        # đa dạng. Trả (phiếu, None) hoặc (None, lỗi).
        def _try(i):
            try:
                return gen_candidate(user, sheets, query, bud, deadline, 0 if i == 0 else VOTE_TEMP, user_full, labels), None
            except Exception as e:  # noqa
                return None, e
        # _agreed(): đã có đáp án nào đạt đủ số phiếu đa số chưa?
        def _agreed():
            counts = Counter(vote_key(x["answer"]) for x in cands if x["answer"])
            return bool(counts) and counts.most_common(1)[0][1] >= need
        # Lượt đầu: sinh song song tối đa 2 phiếu (nếu bật PARALLEL_VOTES), nếu không thì chỉ 1 phiếu.
        first = min(VOTES, 2) if PARALLEL_VOTES else 1
        with ThreadPoolExecutor(max_workers=first) as vex:
            first_res = [f.result() for f in [vex.submit(_try, i) for i in range(first)]]
        # Gom kết quả lượt đầu: phiếu lỗi thì ghi lại lỗi, phiếu tốt thì thêm vào danh sách.
        for c, e in first_res:
            if e is not None:
                rec["err"] = f"{type(e).__name__}: {e}"
                continue
            cands.append(c)
            rec["votes"].append(c["answer"])
        # Không phiếu nào sinh được -> báo lỗi.
        if not cands:
            raise RuntimeError(rec["err"] or "khong co phieu nao")
        # i: số thứ tự phiếu kế tiếp cần sinh.
        i = first
        # Sinh thêm phiếu cho tới khi: đã đủ đa số, hoặc đủ VOTES phiếu hợp lệ, hoặc hết số lượt tối đa (VOTES x
        # VOTE_ATTEMPT_FACTOR).
        while i < VOTES * VOTE_ATTEMPT_FACTOR:
            if _agreed():
                break
            if i >= VOTES and sum(1 for x in cands if x["answer"]) >= VOTES:
                break
            c, e = _try(i)
            if e is not None:
                rec["err"] = f"{type(e).__name__}: {e}"
                break
            cands.append(c)
            rec["votes"].append(c["answer"])
            i += 1
        rec["cand_errs"] = [x.get("err") for x in cands if not x["answer"]]  # lý do phiếu rỗng, xem trong debug_log
        # Ghi toàn bộ các phiếu (đáp án, code, lỗi) vào bản ghi để gỡ lỗi sau.
        if LOG_CANDS:
            rec["cands"] = [{"answer": x["answer"], "raw": x.get("raw"), "err": x.get("err"),
                             "code": x.get("code") or ""} for x in cands]
        # Chỉ xét các phiếu có đáp án (khác rỗng).
        valid = [x for x in cands if x["answer"]]
        if valid:
            # Đếm số phiếu theo vote_key và chọn đáp án có NHIỀU PHIẾU NHẤT (best = phiếu đại diện, final = đáp
            # án).
            counts = Counter(vote_key(x["answer"]) for x in valid)
            top_key = counts.most_common(1)[0][0]
            best = next(x for x in valid if vote_key(x["answer"]) == top_key)
            final = best["answer"]
            path = f"vote{len(cands)}"
            n_distinct = len({vote_key(x["answer"]) for x in valid})
            # Gọi TRỌNG TÀI khi: bật JUDGE, có >= 2 đáp án khác nhau, chưa đủ đa số (hoặc JUDGE_OVERRIDE bật),
            # còn đủ thời gian và chi phí còn dưới ngưỡng cho phép.
            if JUDGE_ON and n_distinct >= 2 and (JUDGE_OVERRIDE or counts.most_common(1)[0][1] < need) and time.time() < deadline - JUDGE_TIME_MARGIN and bud.cost < MAX_COST_PER_Q * JUDGE_COST_RATIO:
                try:
                    jt = {}
                    picked = judge_votes(query, dump, cands, bud, deadline, jt)
                    if picked:
                        rec["judge"] = {"from": final, "to": picked, "why": jt.get("text")}
                        final = picked
                        path = f"judge{len(cands)}"
                except Exception as e:
                    rec["judge"] = {"error": f"{type(e).__name__}: {e}"}
            final = postprocess(final, query, labels)  # ép format lần cuối theo query
            rec.update(answer=final, code=best["code"], raw=best["raw"], path=path, err=None)
        # Không có phiếu hợp lệ nào -> phương án cuối ('direct'): hỏi thẳng LLM đáp án (dùng dump đầy đủ).
        else:
            if cands:
                rec.update(code=cands[-1]["code"], err=cands[-1]["err"])
            text = llm_chat([{"role": "system", "content": DIRECT_SYSTEM}, {"role": "user", "content": (user_full if FULL_DUMP_ON_RETRY and user_full else user)}],
                            bud, deadline)
            # Lấy DÒNG CUỐI của câu trả lời (bỏ rào ```), coi là đáp án.
            raw = re.sub(r"^```\w*|```$", "", text.strip()).strip().splitlines()[-1] if text.strip() else None
            rec.update(answer=postprocess(raw, query, labels), path="direct")
    # Thoát sớm có chủ đích (đường tắt ô) - không phải lỗi.
    except _Done:
        pass
    # Lỗi bất kỳ: ghi vào bản ghi và log, KHÔNG làm sập chương trình.
    except Exception as e:
        rec["err"] = f"{type(e).__name__}: {e}"
        log(f"[{qid}] LOI: {rec['err']}")
    # Ghi thống kê: token vào/ra, chi phí, số lần gọi LLM, thời gian xử lý.
    rec.update(tok_in=bud.tin, tok_out=bud.tout, cost=round(bud.cost, 6), calls=bud.calls, secs=round(time.time() - t0, 1))
    # In một dòng log tóm tắt cho câu hỏi này.
    log(f"[{qid}] {rec['path']:8s} {rec['secs']:5.1f}s tok={bud.tin}/{bud.tout} ${bud.cost:.5f} votes={rec['votes']} -> {rec['answer']!r}")
    return rec

# write_submission(path, ids, answers): ghi file kết quả CSV (id, answer) theo thứ tự `ids`; câu chưa có đáp án để
# trống.
# Ghi NGUYÊN TỬ (ghi file tạm rồi đổi tên) nên file kết quả không bao giờ ở trạng thái hỏng dở.
def write_submission(path, ids, answers):
    tmp = path + ".tmp"
    # Ghi vào file tạm (.tmp) trước.
    with open(tmp, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["id", "answer"])
        for i in ids:
            w.writerow([i, answers.get(i, "")])
    # Đổi tên file tạm thành file thật (thao tác nguyên tử).
    os.replace(tmp, path)

# find_queries(data_dir): tìm file queries.json ở thư mục gốc, rồi tìm đệ quy trong các thư mục con.
def find_queries(data_dir):
    p = os.path.join(data_dir, "queries.json")
    if os.path.exists(p):
        return p
    for root, _, files in os.walk(data_dir):
        if "queries.json" in files:
            return os.path.join(root, "queries.json")
    raise FileNotFoundError("khong tim thay queries.json")

# run(data_dir, output, queries_path): luồng chính - đọc câu hỏi, giải song song, ghi kết quả dần dần, ghi log gỡ
# lỗi và in thống kê.
def run(data_dir, output, queries_path=None):
    # Tạo sẵn thư mục chứa file đầu ra.
    os.makedirs(os.path.dirname(os.path.abspath(output)), exist_ok=True)
    qpath = queries_path if queries_path else find_queries(data_dir)
    # Đọc danh sách câu hỏi từ queries.json.
    queries = json.load(open(qpath, encoding="utf-8"))
    ids = [q["id"] for q in queries]
    # answers: {id: đáp án} | records: các bản ghi chi tiết để ghi log.
    answers, records = {}, []
    # Ghi sẵn file kết quả RỖNG ngay từ đầu để luôn có đầu ra, kể cả khi chương trình bị ngắt giữa chừng.
    write_submission(output, ids, answers)
    # Hạn chót tuyệt đối của cả chương trình.
    deadline = START + TIME_BUDGET
    # Giải các câu hỏi SONG SONG bằng WORKERS luồng.
    with ThreadPoolExecutor(max_workers=WORKERS) as ex:
        futs = [ex.submit(solve_one, q, data_dir, deadline) for q in queries]
        # Thu kết quả theo thứ tự nộp bài; mỗi khi xong một câu thì GHI LẠI file kết quả ngay (không mất kết quả
        # nếu bị ngắt).
        for f in futs:
            try:
                rec = f.result()
                answers[rec["id"]] = rec["answer"]
                records.append(rec)
            except Exception as e:
                log(f"loi khong mong doi: {e}\n{traceback.format_exc()}")
            write_submission(output, ids, answers)
    # Ghi file kết quả lần cuối.
    write_submission(output, ids, answers)
    # Ghi log gỡ lỗi debug_log.jsonl (mỗi dòng một bản ghi JSON); lỗi ghi log thì bỏ qua.
    try:
        if DEBUG_LOG:
            with open(DEBUG_LOG, "w", encoding="utf-8") as f:
                for r in records:
                    f.write(json.dumps(r, ensure_ascii=False) + "\n")
    except Exception:
        pass
    # Tính và in thống kê tổng kết: số câu, thời gian, token trung bình mỗi câu, tổng chi phí.
    tin = sum(r.get("tok_in", 0) for r in records)
    tout = sum(r.get("tok_out", 0) for r in records)
    n = max(len(records), 1)
    log(f"Xong {len(answers)}/{len(ids)} cau trong {time.time() - START:.1f}s | token TB/cau: vao {tin // n}, ra {tout // n} | tong ${sum(r.get('cost', 0) for r in records):.4f}")

# main(): điểm vào dòng lệnh - đọc tham số, áp dụng khoá API/mô hình/URL, rồi chạy run().
def main():
    # Cho phép ghi đè MODEL và API_URL (biến toàn cục) theo tham số dòng lệnh.
    global MODEL, API_URL
    # Khai báo các tham số dòng lệnh.
    ap = argparse.ArgumentParser()
    # 6 tham số mà hệ thống chấm truyền vào
    ap.add_argument("--xlsx-dir", help="thư mục chứa các file .xlsx")
    ap.add_argument("--queries", help="đường dẫn tới queries.json")
    ap.add_argument("--output", "-o", default="submission.csv", help="đường dẫn file submission.csv")
    ap.add_argument("--openrouter-key", help="OpenRouter API key")
    ap.add_argument("--openrouter-model", help="tên model, ví dụ qwen/qwen3.7-flash")
    ap.add_argument("--openrouter-base-url", help="base URL, ví dụ https://openrouter.ai/api/v1")
    ap.add_argument("--input", "-i", help="(cũ) thư mục chứa queries.json và xlsx")
    ap.add_argument("--api-key", help="(cũ) OpenRouter API key")
    ap.add_argument("--dump", help="in ra dump của một file xlsx rồi thoát")
    # parse_known_args: tham số lạ chỉ bị cảnh báo chứ không làm chương trình dừng.
    a, unknown = ap.parse_known_args()
    if unknown:
        log(f"Bo qua tham so khong nhan dien: {unknown}")
    # Khoá API: ưu tiên --openrouter-key, nếu không có thì dùng --api-key (kiểu cũ).
    key = a.openrouter_key or a.api_key
    if key:
        os.environ["OPENROUTER_API_KEY"] = key
    if a.openrouter_model:
        MODEL = a.openrouter_model
    # Chuẩn hoá base URL thành đường dẫn đầy đủ tới /chat/completions.
    if a.openrouter_base_url:
        base = a.openrouter_base_url.rstrip("/")
        API_URL = base if base.endswith("/chat/completions") else base + "/chat/completions"
    # Chế độ gỡ lỗi: chỉ in bản dump của một file xlsx rồi thoát (để xem LLM nhìn thấy gì).
    if a.dump:
        print(load_workbook_sheets(a.dump)[1])
        return
    # Thư mục dữ liệu: ưu tiên --xlsx-dir, nếu không có thì --input (kiểu cũ); thiếu cả hai -> báo lỗi, thoát mã
    # 1.
    data_dir = a.xlsx_dir or a.input
    if not data_dir:
        log("LOI: thieu --xlsx-dir (hoac --input)")
        sys.exit(1)
    # Chạy chương trình chính; mọi lỗi nghiêm trọng -> ghi log và thoát với mã 1.
    try:
        run(data_dir, a.output, a.queries)
    except (RuntimeError, FileNotFoundError) as e:
        log(f"LOI: {e}")
        sys.exit(1)
    except Exception as e:
        log(f"LOI khong mong doi: {e}\n{traceback.format_exc()}")
        sys.exit(1)
# Điểm vào: chỉ chạy main() khi file được chạy trực tiếp (không chạy khi bị import).
if __name__ == "__main__":
    main()