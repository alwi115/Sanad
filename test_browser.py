from pathlib import Path
from http.server import ThreadingHTTPServer, SimpleHTTPRequestHandler
from functools import partial
import threading
from playwright.sync_api import sync_playwright
server=ThreadingHTTPServer(('127.0.0.1',0),partial(SimpleHTTPRequestHandler,directory='/mnt/data/sanad-tool'))
threading.Thread(target=server.serve_forever,daemon=True).start()
p=f'http://127.0.0.1:{server.server_port}/index.html'
with sync_playwright() as pw:
    browser=pw.chromium.launch(headless=True,executable_path='/usr/bin/chromium',args=['--no-sandbox','--allow-file-access-from-files'])
    page=browser.new_page(viewport={'width':1360,'height':920},device_scale_factor=1)
    errors=[]
    page.on('pageerror',lambda e: errors.append(str(e)))
    page.set_content(Path("/mnt/data/sanad-tool/index.html").read_text())
    page.add_style_tag(content=Path("/mnt/data/sanad-tool/styles.css").read_text())
    page.add_script_tag(content=Path("/mnt/data/sanad-tool/vendor/jszip.min.js").read_text())
    page.add_script_tag(content=Path("/mnt/data/sanad-tool/app.js").read_text())
    assert page.title() == 'سَنَد | مساعد المحاسب', page.title()
    assert page.locator('#statOperations').inner_text()=='0'
    page.locator('#demoBtn').click()
    assert page.locator('#reviewTotal').inner_text()=='7',page.locator('#reviewTotal').inner_text()
    assert page.locator('#reviewDuplicates').inner_text()=='1'
    assert page.locator('#reviewFlags').inner_text()=='3',page.locator('#reviewFlags').inner_text()
    page.locator('[data-view="match"]').click()
    assert page.locator('#matchSuggested').inner_text()=='2'
    assert page.locator('#matchApproved').inner_text()=='0'
    page.locator('[data-pair-action="approve"]').first.click()
    assert page.locator('#matchApproved').inner_text()=='1'
    page.locator('[data-view="home"]').click()
    page.locator('[data-company="shop"]').click()
    assert page.locator('#statOperations').inner_text()=='0', 'company data leak'
    page.locator('[data-company="chips"]').click()
    assert page.locator('#statApproved').inner_text()=='1', 'company state lost'
    page.locator('[data-view="report"]').click()
    assert page.locator('#reportStats .report-stat').count()==6
    page.screenshot(path='/mnt/data/sanad-tool/preview-desktop.png',full_page=True)
    page.set_viewport_size({'width':390,'height':844})
    page.screenshot(path='/mnt/data/sanad-tool/preview-mobile.png',full_page=True)
    # paste import test on a separate company
    page.locator('[data-view="home"]').click()
    page.locator('[data-company="kitchens"]').click()
    page.locator('[data-view="import"]').click()
    page.locator('#pasteToggle').click()
    page.locator('#pasteInput').fill('التاريخ\tالبيان\tالمبلغ\tرقم المستند\tنوع العملية\n2026-09-12\tاختبار\t12.500\tABC-1\tمصروف')
    page.locator('#pasteRead').click()
    assert not page.locator('#mappingPanel').is_hidden()
    assert page.locator('#map-date').input_value()=='0'
    assert page.locator('#map-amount').input_value()=='2'
    page.locator('#confirmImport').click()
    assert page.locator('#reviewTotal').inner_text()=='1'
    assert page.locator('#reviewFlags').inner_text()=='0'
    assert page.evaluate("Number.isNaN(parseAmount('abc123'))") is True
    assert page.evaluate("parseAmount('١٬٢٣٤٫٥٠٠ ر.ع')") == 1234.5
    assert page.evaluate("normalizeDate('12/09/2026')") == '2026-09-12'
    assert page.evaluate("parseDelimited('\"a,b\",x\\n1,2', ',')[0][0]") == 'a,b'
    # Offline XLSX: multi-sheet selection + numeric Excel dates + signed expense mapping.
    page.locator('[data-view="home"]').click()
    page.locator('[data-company="shop"]').click()
    page.locator('[data-view="import"]').click()
    page.locator('#ledgerFile').set_input_files('/tmp/sanad-test.xlsx')
    page.locator('#mappingPanel').wait_for(state='visible',timeout=10000)
    assert page.locator('#sheetSelect option').count() == 2, page.locator('#toast').inner_text()
    page.locator('#sheetSelect').select_option('1')
    assert page.locator('#map-amount').input_value() == '2'
    page.locator('#confirmImport').click()
    assert page.locator('#reviewTotal').inner_text() == '2'
    assert page.locator('#reviewFlags').inner_text() == '0'
    assert '-125.750' in page.locator('#reviewTable').inner_text()
    # CSV export is actual downloadable and protects formula-starting text.
    page.locator('[data-view="report"]').click()
    with page.expect_download() as download_info:
        page.locator('#exportReport').click()
    download=download_info.value
    assert download.suggested_filename.endswith('.csv')
    assert page.evaluate("safeCsvCell('=1+1')") == "\"'=1+1\""
    assert errors==[],errors
    print('PASS: app loads; demo 7 operations/1 duplicate/3 flagged; 2 suggested matches; approval; four-company isolation; report; mobile view; clipboard import; XLSX multi-sheet with serial dates; CSV download; formula guard; no browser JS errors')
    browser.close()
