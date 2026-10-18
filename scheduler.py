"""
碳足跡排放係數資料庫 每日定時排程服務 (Scheduler)
每天定時自動呼叫 crawler.py 抓取最新係數並更新 data.xlsx
"""

import os
import sys
import time
import argparse
import logging
from datetime import datetime, timedelta

# 設定 stdout 編碼為 UTF-8
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

# 載入爬蟲主程式
from crawler import crawl_and_export, DEFAULT_EXCEL_PATH

LOG_FORMAT = '%(asctime)s [%(levelname)s] %(message)s'
logging.basicConfig(
    level=logging.INFO,
    format=LOG_FORMAT,
    handlers=[
        logging.StreamHandler(sys.stdout),
        logging.FileHandler(
            os.path.join(os.path.dirname(os.path.abspath(__file__)), 'scheduler.log'),
            encoding='utf-8'
        )
    ]
)
logger = logging.getLogger('CarbonScheduler')


def calculate_seconds_until(target_time_str: str) -> tuple[float, datetime]:
    """
    計算距離下一個目標時間點 (格式 "HH:MM") 的秒數與時間
    """
    now = datetime.now()
    target_hour, target_minute = map(int, target_time_str.split(':'))
    
    target_dt = now.replace(hour=target_hour, minute=target_minute, second=0, microsecond=0)
    if target_dt <= now:
        target_dt += timedelta(days=1)
        
    seconds_until = (target_dt - now).total_seconds()
    return seconds_until, target_dt


def run_scheduler(target_time_str: str = "03:00", run_now: bool = False, excel_path: str = DEFAULT_EXCEL_PATH):
    """
    定時循環執行爬蟲作業
    """
    logger.info("=" * 60)
    logger.info("啟動碳足跡排放係數每日定時排程服務")
    logger.info(f"每日預定執行時間: {target_time_str}")
    logger.info(f"匯出目標路徑: {excel_path}")
    logger.info("=" * 60)
    
    # 若指定立即執行一次
    if run_now:
        logger.info("【立即執行】正在執行首次爬蟲資料擷取與匯出...")
        try:
            crawl_and_export(excel_path)
        except Exception as e:
            logger.error(f"首次執行失敗: {e}", exc_info=True)
            
    while True:
        seconds_until, next_dt = calculate_seconds_until(target_time_str)
        hours = int(seconds_until // 3600)
        minutes = int((seconds_until % 3600) // 60)
        secs = int(seconds_until % 60)
        
        logger.info(f"下一次預計執行時間: {next_dt.strftime('%Y-%m-%d %H:%M:%S')} (倒數約 {hours} 小時 {minutes} 分 {secs} 秒)")
        
        # 休眠至預計時間 (採用每小時或整段休眠)
        try:
            time.sleep(seconds_until)
        except KeyboardInterrupt:
            logger.info("收到中斷訊號 (Ctrl+C)，排程服務已正常停止。")
            break
            
        logger.info(f"到達預定執行時間: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}，開始抓取最新資料...")
        try:
            crawl_and_export(excel_path)
            logger.info("本次每日定時任務執行成功！")
        except Exception as e:
            logger.error(f"本次每日定時任務執行發生異常: {e}", exc_info=True)
            # 發生異常時稍作等待，避免連續重試
            time.sleep(60)


def main():
    parser = argparse.ArgumentParser(description="碳足跡排放係數資料庫 每日定時排程服務")
    parser.add_argument("--time", type=str, default="03:00", help="每日定時執行時間 (HH:MM，預設 03:00)")
    parser.add_argument("--run-now", action="store_true", help="啟動時立即執行一次抓取")
    parser.add_argument("--excel-path", type=str, default=DEFAULT_EXCEL_PATH, help="Excel 儲存路徑")
    
    args = parser.parse_args()
    run_scheduler(target_time_str=args.time, run_now=args.run_now, excel_path=args.excel_path)


if __name__ == '__main__':
    main()
