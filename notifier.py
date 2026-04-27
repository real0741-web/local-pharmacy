#!/usr/bin/env python3
"""
로컬 창고약국 — 주문 알림 & 자동 출력 프로그램
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
약국 PC에서 실행하면 새 주문 도착 시:
  1. 화면 팝업 알림 (최상단 표시)
  2. 영수증 프린터 자동 출력

실행 방법:  설치_실행.bat 더블클릭
종료:       Ctrl + C
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
"""

import requests
import time
import threading
import json
import os
import sys
from datetime import datetime

# ════════════════════════════════════════════════════
#   ▼▼▼  약국 설치 시 여기만 수정하세요  ▼▼▼
# ════════════════════════════════════════════════════

# 프린터 이름
#   - 비워두면 Windows 기본 프린터 자동 사용
#   - 프린터 이름 확인: 윈도우 → 설정 → 프린터 및 스캐너
#   예시: 'BIXOLON SRP-350III'  /  'EPSON TM-T88V'  /  'Star TSP143'
PRINTER_NAME = ''

PHARMACY_NAME = '로컬 창고약국'   # 영수증 상단 약국명
POLL_SECONDS  = 3                  # 주문 확인 주기 (초)

# 영수증 용지 너비
#   58mm 프린터  → 32
#   80mm 프린터  → 48
RECEIPT_WIDTH = 32

# ntfy 토픽 — index.html 의 NTFY_TOPIC 과 반드시 일치해야 합니다
NTFY_TOPIC = 'wooripharmacy_orders'

# ════════════════════════════════════════════════════
#   ▲▲▲  수정 끝  ▲▲▲
# ════════════════════════════════════════════════════

_ntfy_since = str(int(time.time()))   # 프로그램 시작 이후 주문만 수신


# ─── 주문 수신 (ntfy.sh 폴링) ──────────────────────────────────────

def get_new_orders():
    """ntfy.sh 에서 새 주문을 폴링합니다."""
    global _ntfy_since
    try:
        resp = requests.get(
            f'https://ntfy.sh/{NTFY_TOPIC}/json',
            params={'poll': '1', 'since': _ntfy_since},
            timeout=10
        )
        orders = []
        for line in resp.text.strip().split('\n'):
            line = line.strip()
            if not line:
                continue
            try:
                msg = json.loads(line)
                if msg.get('event') == 'message' and msg.get('message'):
                    _ntfy_since = msg['id']
                    orders.append(msg['message'])
            except Exception:
                pass
        return orders
    except Exception:
        return []


# ─── 화면 팝업 ─────────────────────────────────────────────────────

def show_popup(order_text):
    """새 주문 팝업을 화면 최상단에 띄웁니다."""
    try:
        import tkinter as tk
        from tkinter import messagebox

        root = tk.Tk()
        root.withdraw()
        root.attributes('-topmost', True)
        root.lift()

        lines = order_text.split('\n')
        short = '\n'.join(lines[:10])
        if len(lines) > 10:
            short += '\n...'

        messagebox.showinfo('🔔  새 주문이 들어왔어요!', short, parent=root)
        root.destroy()
    except Exception as e:
        print(f'  [팝업 오류] {e}')


# ─── 영수증 출력 ────────────────────────────────────────────────────

def build_receipt(order_text):
    """영수증 텍스트를 만듭니다."""
    now = datetime.now().strftime('%Y-%m-%d  %H:%M:%S')
    W = RECEIPT_WIDTH

    lines_in = order_text.strip().split('\n')
    out = [
        '',
        PHARMACY_NAME.center(W),
        now.center(W),
        '─' * W,
    ]

    for l in lines_in:
        clean = l.encode('ascii', 'ignore').decode().strip()
        if not clean:
            clean = l.strip()
        out.append(clean)

    out += [
        '─' * W,
        '카운터에서 결제해 주세요'.center(W),
        '─' * W,
        '', '', '',   # 자동 컷팅 여백
    ]
    return '\n'.join(out)


def print_receipt(order_text):
    """열감지 영수증 프린터(ESC/POS RAW)로 출력합니다."""
    try:
        import win32print

        printer_name = PRINTER_NAME or win32print.GetDefaultPrinter()
        receipt = build_receipt(order_text)
        data = receipt.encode('cp949', errors='replace')

        hPrinter = win32print.OpenPrinter(printer_name)
        try:
            hJob = win32print.StartDocPrinter(hPrinter, 1, ('Order', None, 'RAW'))
            win32print.StartPagePrinter(hPrinter)
            win32print.WritePrinter(hPrinter, data)
            win32print.EndPagePrinter(hPrinter)
            win32print.EndDocPrinter(hPrinter)
        finally:
            win32print.ClosePrinter(hPrinter)

        ts = datetime.now().strftime('%H:%M:%S')
        print(f'  [{ts}] 출력 완료 → {printer_name}')

    except ImportError:
        print('  ⚠️  pywin32 미설치 → pip install pywin32')
    except Exception as e:
        import traceback
        print(f'  ⚠️  출력 오류: {e}')
        traceback.print_exc()


# ─── 주문 처리 ─────────────────────────────────────────────────────

def handle_order(text):
    """새 주문: 팝업(별도 스레드) + 출력"""
    ts = datetime.now().strftime('%H:%M:%S')
    print(f'\n  [{ts}] ★ 새 주문 수신!')
    for line in text.split('\n')[:6]:
        print(f'       {line}')

    threading.Thread(target=show_popup, args=(text,), daemon=True).start()

    print(f'  [{ts}] 🖨️  인쇄 시도 중...')
    try:
        print_receipt(text)
    except Exception as e:
        import traceback
        print(f'  ⚠️  인쇄 오류: {e}')
        traceback.print_exc()


# ─── 자동 시작 등록 ────────────────────────────────────────────────

def register_startup():
    """Windows 로그인 시 자동 실행 등록 (최초 1회)"""
    try:
        import winreg
        notifier_path = os.path.abspath(__file__)
        python_exe    = sys.executable
        command       = f'"{python_exe}" "{notifier_path}"'
        reg_key = r'SOFTWARE\Microsoft\Windows\CurrentVersion\Run'
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, reg_key, 0, winreg.KEY_ALL_ACCESS) as key:
            try:
                existing, _ = winreg.QueryValueEx(key, 'LocalPharmacyNotifier')
                if existing == command:
                    return
            except FileNotFoundError:
                pass
            winreg.SetValueEx(key, 'LocalPharmacyNotifier', 0, winreg.REG_SZ, command)
            print('  ✅ 시작 프로그램 등록 완료 (다음 로그인부터 자동 시작)')
    except Exception as e:
        print(f'  ⚠️  자동 시작 등록 실패: {e}')


# ─── 메인 ──────────────────────────────────────────────────────────

def main():
    print()
    print('=' * 44)
    print('    로컬 창고약국  주문 알림 프로그램')
    print('=' * 44)

    register_startup()

    try:
        import requests  # noqa
    except ImportError:
        print('\n  ❌ requests 미설치 → pip install requests')
        return

    try:
        import win32print
        printer = PRINTER_NAME or win32print.GetDefaultPrinter()
        print(f'\n  🖨️  프린터: {printer}')
    except ImportError:
        print('\n  ⚠️  pywin32 미설치 — 출력 비활성화 (pip install pywin32)')

    print(f'  📡 ntfy 채널 확인 중...')
    print(f'      topic: {NTFY_TOPIC}')
    try:
        r = requests.get(
            f'https://ntfy.sh/{NTFY_TOPIC}/json',
            params={'poll': '1', 'since': str(int(time.time()) - 1)},
            timeout=5
        )
        if r.status_code == 200:
            print(f'  ✅ ntfy 연결 성공!')
        else:
            print(f'  ⚠️  ntfy 응답 이상: {r.status_code}')
    except Exception as e:
        print(f'  ❌ ntfy 연결 실패: {e}')
        return

    print(f'\n  ✅ 준비 완료! {POLL_SECONDS}초마다 새 주문을 확인합니다.')
    print('  종료하려면 Ctrl+C 를 누르세요.\n')
    print('-' * 44)

    while True:
        try:
            orders = get_new_orders()
            for order_text in orders:
                handle_order(order_text)
        except KeyboardInterrupt:
            print('\n\n  프로그램을 종료합니다.')
            break
        except Exception as e:
            ts = datetime.now().strftime('%H:%M:%S')
            print(f'  [{ts}] 오류: {e}')
        time.sleep(POLL_SECONDS)


if __name__ == '__main__':
    main()
