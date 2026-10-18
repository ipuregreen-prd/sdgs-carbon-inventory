"""
碳足跡排放係數資料庫 每日定時爬蟲程式
資料來源：中華民國環境部 產品碳足跡資訊網
URL: https://cfp.moenv.gov.tw/WebPage/WebSites/CoefficientDB.aspx
目標檔案：C:\\0-Learning\\0-綠色公益基金會\\程式\\sdgs-carbon-inventory\\data.xlsx
"""

import os
import sys
import time
import re
import logging
from datetime import datetime
import requests
from lxml import html
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# 設定 stdout 編碼為 UTF-8 (避免 Windows 終端亂碼)
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

# 設定日誌
LOG_FORMAT = '%(asctime)s [%(levelname)s] %(message)s'
logging.basicConfig(
    level=logging.INFO,
    format=LOG_FORMAT,
    handlers=[
        logging.StreamHandler(sys.stdout),
        logging.FileHandler(
            os.path.join(os.path.dirname(os.path.abspath(__file__)), 'crawler.log'),
            encoding='utf-8'
        )
    ]
)
logger = logging.getLogger('CarbonCrawler')

BASE_URL = 'https://cfp.moenv.gov.tw/WebPage/WebSites/CoefficientDB.aspx'
DEFAULT_EXCEL_PATH = r'C:\0-Learning\0-綠色公益基金會\程式\sdgs-carbon-inventory\data.xlsx'

DEFAULT_HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36',
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8',
    'Accept-Language': 'zh-TW,zh;q=0.9,en-US;q=0.8,en;q=0.7',
    'Referer': BASE_URL,
}


def fetch_initial_page(session: requests.Session, max_retries: int = 3, retry_delay: float = 5.0):
    """
    獲取首頁內容與 ASP.NET ViewState 相關表單欄位
    """
    for attempt in range(1, max_retries + 1):
        try:
            logger.info(f"正在連線至碳足跡排放係數資料庫首頁... (第 {attempt}/{max_retries} 次)")
            resp = session.get(BASE_URL, headers=DEFAULT_HEADERS, timeout=30)
            if resp.status_code == 429:
                wait_time = retry_delay * attempt * 2
                logger.warning(f"遭遇 429 請求過於頻繁限制，等待 {wait_time:.1f} 秒後重試...")
                time.sleep(wait_time)
                continue
            resp.raise_for_status()
            
            doc = html.fromstring(resp.content.decode('utf-8', errors='ignore'))
            
            # 抓取所有 input 隱藏欄位 (ViewState, ViewStateGenerator, EventValidation 等)
            form_inputs = {}
            for inp in doc.xpath('//form//input'):
                name = inp.get('name')
                val = inp.get('value', '')
                if name and 'btn_export' not in name:
                    form_inputs[name] = val
                    
            if '__VIEWSTATE' not in form_inputs:
                raise ValueError("無法在首頁中解析出 __VIEWSTATE 欄位")
                
            # 同步解析首頁上的分類層級清單 (大類、中類、小類)
            categories = parse_category_tree(doc)
            
            return form_inputs, categories
        except Exception as e:
            logger.warning(f"獲取首頁失敗: {e}")
            if attempt < max_retries:
                time.sleep(retry_delay * attempt)
            else:
                raise


def parse_category_tree(doc):
    """
    解析首頁中的分類樹狀結構 (四大類別、中類、小類與 cc_id)
    """
    categories = []
    try:
        # collapse1 包含完整分類結構卡片
        cards = doc.xpath('//*[@id="collapse1"]//div[contains(@class, "card")]')
        for card in cards:
            hdr = card.xpath('.//*[contains(@class, "card-header")]')
            cat_name = hdr[0].text_content().strip() if hdr else ''
            
            links = card.xpath('.//a[@cc_id]')
            for a in links:
                cc_id = a.get('cc_id', '').strip()
                item_name = a.text_content().strip()
                if item_name:
                    categories.append({
                        '中類': cat_name,
                        '小類': item_name,
                        '分類代碼(cc_id)': cc_id
                    })
    except Exception as e:
        logger.warning(f"解析分類架構時發生小錯誤 (不影響主數據抓取): {e}")
    return categories


def fetch_all_coefficients(session: requests.Session, form_inputs: dict, max_retries: int = 3, retry_delay: float = 5.0):
    """
    送出查詢請求以獲取全部碳足跡排放係數
    """
    # 稍微停頓以模擬正常使用者行為，避免觸發防爬蟲頻率限制
    time.sleep(1.5)
    
    post_data = dict(form_inputs)
    # ASP.NET AJAX UpdatePanel 查詢參數設定
    post_data['ctl00$ContentPlaceHolder1$CoefficientData_Tab$ctl00'] = 'ctl00$ContentPlaceHolder1$CoefficientData_Tab$ctl01|ctl00$ContentPlaceHolder1$CoefficientData_Tab$btnSearch'
    post_data['__ASYNCPOST'] = 'true'
    post_data['ctl00$ContentPlaceHolder1$CoefficientData_Tab$txt_keyword'] = ''
    post_data['ctl00$ContentPlaceHolder1$CoefficientData_Tab$btnSearch'] = '查詢'
    
    ajax_headers = dict(DEFAULT_HEADERS)
    ajax_headers.update({
        'X-MicrosoftAjax': 'Delta=true',
        'X-Requested-With': 'XMLHttpRequest'
    })
    
    for attempt in range(1, max_retries + 1):
        try:
            logger.info(f"正在發送查詢請求以擷取全部係數數據... (第 {attempt}/{max_retries} 次)")
            resp = session.post(BASE_URL, data=post_data, headers=ajax_headers, timeout=60)
            if resp.status_code == 429:
                wait_time = retry_delay * attempt * 2
                logger.warning(f"遭遇 429 頻率限制，等待 {wait_time:.1f} 秒後重試...")
                time.sleep(wait_time)
                continue
            resp.raise_for_status()
            
            # 解析 ASP.NET AJAX Delta 回應
            tokens = resp.text.split('|')
            panel_html = None
            for i in range(len(tokens) - 3):
                if tokens[i] == 'updatePanel' and 'ContentPlaceHolder1_CoefficientData_Tab' in tokens[i+1]:
                    panel_html = tokens[i+2]
                    break
                    
            if not panel_html:
                raise ValueError("未在 AJAX 回應中找到含有數據的 updatePanel 區塊")
                
            doc_panel = html.fromstring(panel_html)
            table_list = doc_panel.xpath('//table[@id="ContentPlaceHolder1_CoefficientData_Tab_sgvList_Tab23"]')
            if not table_list:
                raise ValueError("未找到係數資料表格 (sgvList_Tab23)")
                
            table = table_list[0]
            rows = table.xpath('.//tr')
            
            records = []
            crawl_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            
            for idx, r in enumerate(rows[1:], start=1):
                tds = r.xpath('.//td')
                if len(tds) < 5:
                    continue
                
                # 欄位1: 碳係數名稱與編號
                link = tds[0].xpath('.//a')
                cd_id = link[0].get('cd_id', '').strip() if link else ''
                name = tds[0].text_content().strip()
                
                # 欄位2: 生產區域名稱
                region = tds[1].text_content().strip()
                
                # 欄位3: 係數數值與單位 (如 "6.058E-01 kgCO₂e")
                raw_val = tds[2].text_content().strip()
                val_match = re.match(r'^([\d\.\+\-Ee]+)\s*(.*)$', raw_val)
                if val_match:
                    num_str = val_match.group(1)
                    unit = val_match.group(2).strip()
                    try:
                        num_float = float(num_str)
                    except ValueError:
                        num_float = None
                else:
                    num_str = raw_val
                    unit = ''
                    num_float = None
                
                # 欄位4: 宣告單位 (如 "度(kwh)")
                declared_unit = tds[3].text_content().strip()
                
                # 欄位5: 公告年份 (如 "2025")
                year = tds[4].text_content().strip()
                
                records.append({
                    '序號': idx,
                    '係數編號': cd_id,
                    '碳係數名稱': name,
                    '生產區域': region,
                    '數值': num_float,
                    '數值(科學記號)': num_str,
                    '碳排單位': unit,
                    '宣告單位': declared_unit,
                    '完整數值與單位': f"{raw_val} / {declared_unit}" if declared_unit else raw_val,
                    '公告年份': year,
                    '數據抓取時間': crawl_time
                })
                
            logger.info(f"成功擷取到 {len(records)} 筆碳足跡排放係數資料！")
            return records
            
        except Exception as e:
            logger.warning(f"擷取係數數據失敗: {e}")
            if attempt < max_retries:
                time.sleep(retry_delay * attempt)
            else:
                raise


def export_to_excel(records: list, categories: list, target_path: str):
    """
    將爬取到的數據匯出為結構化、排版美觀的 Excel 活頁簿
    """
    logger.info(f"正在將資料匯出至 Excel: {target_path}")
    
    os.makedirs(os.path.dirname(os.path.abspath(target_path)), exist_ok=True)
    
    wb = openpyxl.Workbook()
    # 移除預設工作表
    default_sheet = wb.active
    wb.remove(default_sheet)
    
    # 樣式定義 (綠色主題，呼應綠色公益基金會與永續碳盤查風格)
    header_fill = PatternFill(start_color="1E4620", end_color="1E4620", fill_type="solid")  # 深森林綠
    sub_header_fill = PatternFill(start_color="2D6A4F", end_color="2D6A4F", fill_type="solid")
    accent_fill = PatternFill(start_color="E8F5E9", end_color="E8F5E9", fill_type="solid")  # 淡綠斑馬紋
    
    header_font = Font(name="微軟正黑體", size=11, bold=True, color="FFFFFF")
    data_font = Font(name="微軟正黑體", size=10)
    bold_font = Font(name="微軟正黑體", size=10, bold=True)
    
    thin_border = Border(
        left=Side(style='thin', color='D0D7DE'),
        right=Side(style='thin', color='D0D7DE'),
        top=Side(style='thin', color='D0D7DE'),
        bottom=Side(style='thin', color='D0D7DE')
    )
    
    align_center = Alignment(horizontal="center", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center")
    align_right = Alignment(horizontal="right", vertical="center")
    
    # -------------------------------------------------------------
    # 工作表 1: 碳足跡排放係數清冊
    # -------------------------------------------------------------
    ws1 = wb.create_sheet(title="碳足跡排放係數清冊")
    ws1.views.sheetView[0].showGridLines = True
    
    headers1 = [
        "序號", "係數編號", "碳係數名稱", "生產區域",
        "係數數值(數值)", "係數數值(科學記號)", "碳排單位", "宣告單位",
        "完整數值與單位", "公告年份", "數據抓取時間"
    ]
    ws1.append(headers1)
    
    # 樣式化標頭
    ws1.row_dimensions[1].height = 28
    for col_idx, col_name in enumerate(headers1, start=1):
        cell = ws1.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = align_center
        cell.border = thin_border
        
    for row_idx, rec in enumerate(records, start=2):
        ws1.row_dimensions[row_idx].height = 20
        row_values = [
            rec['序號'],
            rec['係數編號'],
            rec['碳係數名稱'],
            rec['生產區域'],
            rec['數值'],
            rec['數值(科學記號)'],
            rec['碳排單位'],
            rec['宣告單位'],
            rec['完整數值與單位'],
            rec['公告年份'],
            rec['數據抓取時間']
        ]
        ws1.append(row_values)
        
        # 斑馬紋與對齊格式
        is_even = (row_idx % 2 == 0)
        row_fill = accent_fill if is_even else None
        
        for col_idx, val in enumerate(row_values, start=1):
            cell = ws1.cell(row=row_idx, column=col_idx)
            cell.font = data_font
            cell.border = thin_border
            if row_fill:
                cell.fill = row_fill
                
            # 特殊欄位對齊與格式
            if col_idx in [1, 2, 4, 10, 11]:  # 序號, 係數編號, 生產區域, 公告年份, 抓取時間
                cell.alignment = align_center
            elif col_idx in [5]:  # 浮點數值
                cell.alignment = align_right
                if isinstance(val, (int, float)):
                    cell.number_format = '#,##0.000000'
            elif col_idx in [6]:  # 科學記號字串
                cell.alignment = align_right
            else:
                cell.alignment = align_left
                
    # 啟用自動篩選
    ws1.auto_filter.ref = ws1.dimensions
    # 凍結第一列
    ws1.freeze_panes = "A2"
    
    # -------------------------------------------------------------
    # 工作表 2: 分類架構表
    # -------------------------------------------------------------
    if categories:
        ws2 = wb.create_sheet(title="係數分類架構")
        ws2.views.sheetView[0].showGridLines = True
        headers2 = ["序號", "中類", "小類", "分類代碼(cc_id)"]
        ws2.append(headers2)
        ws2.row_dimensions[1].height = 26
        
        for col_idx in range(1, len(headers2) + 1):
            cell = ws2.cell(row=1, column=col_idx)
            cell.fill = sub_header_fill
            cell.font = header_font
            cell.alignment = align_center
            cell.border = thin_border
            
        for row_idx, cat in enumerate(categories, start=2):
            ws2.row_dimensions[row_idx].height = 19
            row_vals = [row_idx - 1, cat['中類'], cat['小類'], cat['分類代碼(cc_id)']]
            ws2.append(row_vals)
            is_even = (row_idx % 2 == 0)
            row_fill = accent_fill if is_even else None
            for col_idx, val in enumerate(row_vals, start=1):
                cell = ws2.cell(row=row_idx, column=col_idx)
                cell.font = data_font
                cell.border = thin_border
                if row_fill:
                    cell.fill = row_fill
                cell.alignment = align_center if col_idx in [1, 4] else align_left
                
        ws2.auto_filter.ref = ws2.dimensions
        ws2.freeze_panes = "A2"
        
    # -------------------------------------------------------------
    # 工作表 3: 摘要與更新日誌
    # -------------------------------------------------------------
    ws3 = wb.create_sheet(title="更新日誌與說明")
    ws3.views.sheetView[0].showGridLines = True
    ws3.row_dimensions[1].height = 26
    ws3.append(["項目", "詳細內容"])
    ws3.cell(row=1, column=1).fill = header_fill
    ws3.cell(row=1, column=1).font = header_font
    ws3.cell(row=1, column=2).fill = header_fill
    ws3.cell(row=1, column=2).font = header_font
    
    summary_data = [
        ("資料來源", "中華民國環境部 產品碳足跡資訊網 (CFP)"),
        ("來源網址", BASE_URL),
        ("最近抓取時間", datetime.now().strftime('%Y-%m-%d %H:%M:%S')),
        ("有效係數總筆數", f"{len(records):,} 筆"),
        ("涵蓋公告年份", f"{min((r['公告年份'] for r in records if r['公告年份']), default='-')} ~ {max((r['公告年份'] for r in records if r['公告年份']), default='-')}"),
        ("使用單位", "綠色公益基金會 (SDGs 碳盤查計畫)"),
        ("執行狀態", "每日定時排程爬蟲自動同步完成"),
    ]
    for row_idx, (k, v) in enumerate(summary_data, start=2):
        ws3.row_dimensions[row_idx].height = 22
        c1 = ws3.cell(row=row_idx, column=1, value=k)
        c2 = ws3.cell(row=row_idx, column=2, value=v)
        c1.font = bold_font
        c2.font = data_font
        c1.border = thin_border
        c2.border = thin_border
        c1.alignment = align_left
        c2.alignment = align_left
        
    # 自動調整欄寬
    for ws in [ws1, ws2 if categories else None, ws3]:
        if ws is None:
            continue
        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                val_str = str(cell.value or '')
                # 計算全形與半形寬度
                length = sum(2 if ord(ch) > 127 else 1 for ch in val_str)
                if length > max_len:
                    max_len = length
            ws.column_dimensions[col_letter].width = max(max_len + 3, 12)
            
    # 儲存 Excel
    wb.save(target_path)
    logger.info(f"Excel 匯出成功！檔案已儲存至: {target_path} (大小: {os.path.getsize(target_path):,} bytes)")


def crawl_and_export(target_path: str = DEFAULT_EXCEL_PATH):
    """
    執行單次爬取與匯出完整流程
    """
    logger.info("=" * 60)
    logger.info("開始執行環境部碳足跡排放係數資料庫爬蟲作業...")
    start_time = time.time()
    
    session = requests.Session()
    form_inputs, categories = fetch_initial_page(session)
    records = fetch_all_coefficients(session, form_inputs)
    export_to_excel(records, categories, target_path)
    
    elapsed = time.time() - start_time
    logger.info(f"爬蟲作業全部順利完成！共抓取 {len(records)} 筆資料，耗時 {elapsed:.2f} 秒。")
    logger.info("=" * 60)
    return records


if __name__ == '__main__':
    # 支援直接執行單次爬取
    target = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_EXCEL_PATH
    crawl_and_export(target)
