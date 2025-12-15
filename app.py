# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import os
import glob
import time
import traceback
import tempfile
import zipfile
import shutil
from io import BytesIO

# Selenium 및 드라이버 관리 모듈
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service 
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.remote.remote_connection import RemoteConnection 
from webdriver_manager.chrome import ChromeDriverManager


# =========================
# 1. 상태 관리 및 유틸리티
# =========================
if 'log_messages' not in st.session_state:
    st.session_state.log_messages = []
if 'zip_download_data' not in st.session_state:
    st.session_state.zip_download_data = None
if 'is_running' not in st.session_state:
    st.session_state.is_running = False

def append_log(text: str):
    """로그 메시지를 세션 상태에 추가합니다."""
    st.session_state.log_messages.append(text)

def clear_log():
    """로그와 다운로드 데이터를 초기화합니다."""
    st.session_state.log_messages = []
    st.session_state.zip_download_data = None


@st.cache_resource(ttl=3600)
def get_chrome_driver_path():
    """크롬 드라이버 경로를 한 번만 설치/가져옵니다."""
    try:
        path = ChromeDriverManager().install()
        return path
    except Exception as e:
        return 'chromedriver' 

# =========================
# 2. Selenium 작업 함수 (V28: PDF 대기 시간 30초)
# =========================
def run_selenium_process(uploaded_file_bytes: bytes, log_placeholder):
    """
    엑셀 파일을 처리하고 우체국 등기 조회를 수행하는 핵심 로직입니다.
    """
    st.session_state.is_running = True
    clear_log()
    driver = None
    successful_files = []
    
    # [V27] 드라이버 타임아웃 설정 (5분)
    RemoteConnection.set_timeout(300) 
    
    with tempfile.TemporaryDirectory() as temp_save_dir:
        
        def log_and_update(message):
            """로그를 기록하고 Streamlit UI를 즉시 업데이트합니다."""
            append_log(message)
            log_placeholder.code('\n'.join(st.session_state.log_messages), language='text')
            time.sleep(0.1) 

        try:
            log_and_update(f"임시 저장 폴더 생성: {os.path.basename(temp_save_dir)}")
            
            # 엑셀 로드
            df = pd.read_excel(BytesIO(uploaded_file_bytes))

            if "등기번호" not in df.columns:
                log_and_update("엑셀에 '등기번호' 컬럼이 없습니다.")
                return

            # 크롬 옵션 설정
            options = Options()
            options.add_experimental_option("prefs", {
                "printing.print_preview_sticky_settings.appState": '{"recentDestinations": [{"id": "Save as PDF", "origin": "local"}], "selectedDestinationId": "Save as PDF", "version": 2}',
                "savefile.default_directory": temp_save_dir
            })
            options.add_argument("--kiosk-printing")
            options.add_argument("--headless") 
            
            # --- 클라우드 안정화 및 DevToolsActivePort 에러 회피 최종 옵션 ---
            options.add_argument("--no-sandbox") 
            options.add_argument("--disable-dev-shm-usage") 
            options.add_argument("--disable-gpu")
            options.add_argument("--window-size=1920,1080")
            options.add_argument("--remote-debugging-pipe") 
            options.add_argument("--user-data-dir=/tmp/user-data")
            options.add_argument("--data-path=/tmp/data-path")
            options.add_argument("--disk-cache-dir=/tmp/cache-dir")
            options.add_argument("--disable-extensions")
            options.add_argument("--disable-application-cache")
            options.add_argument("--disable-logging")

            # 드라이버 실행
            try:
                # 로컬과 클라우드 환경 분리 (V25 로직 기반)
                if 'chrome' in st.secrets and 'BIN' in st.secrets['chrome']:
                    # 1. 클라우드 환경: secrets에 설정된 경로 사용
                    options.binary_location = st.secrets['chrome']['BIN']
                    log_and_update(f"Chromium BIN 경로 사용: {st.secrets['chrome']['BIN']}")
                    driver = webdriver.Chrome(options=options)
                else:
                    # 2. 로컬 환경: webdriver-manager 사용
                    driver_path = get_chrome_driver_path()
                    service = Service(driver_path) 
                    driver = webdriver.Chrome(service=service, options=options)
                    log_and_update("로컬 환경: webdriver-manager 드라이버 경로 사용")

            except Exception as e:
                # secrets 에러가 발생해도 로컬 테스트를 계속 진행할 수 있도록 처리
                if "StreamlitSecretNotFoundError" in str(e):
                    driver_path = get_chrome_driver_path()
                    service = Service(driver_path)
                    driver = webdriver.Chrome(service=service, options=options)
                    log_and_update("로컬 환경 (Secrets 에러 무시): webdriver-manager 드라이버 경로 사용")
                else:
                    raise e
                    
            log_and_update("Chrome 드라이버 세션 시작 시도 완료.")
            driver.maximize_window()
            wait = WebDriverWait(driver, 20) 
            total = len(df)
            
            for i, row in df.iterrows():
                tracking_number = str(row["등기번호"]).strip()
                if not tracking_number: continue

                log_and_update(f"[{i+1}/{total}] 조회 시도: {tracking_number}")

                before_files = set(glob.glob(os.path.join(temp_save_dir, "*.pdf")))
                
                # V26: time.sleep(1.5) 제거
                driver.get("https://service.epost.go.kr/trace.RetrieveDomRigiTraceList.comm")
                # driver.get() 성공을 위해 명시적 대기가 필요할 수 있음
                wait.until(EC.presence_of_element_located((By.ID, "sid1")))


                try:
                    # Selenium 핵심 로직
                    input_box = driver.find_element(By.ID, "sid1") 
                    input_box.clear()
                    input_box.send_keys(tracking_number)
                    
                    try:
                        form_elem = driver.find_element(By.ID, "frmDomRigiTrace")
                        driver.execute_script("arguments[0].submit();", form_elem)
                    except:
                        input_box.send_keys(Keys.RETURN)

                    print_btn = wait.until(EC.element_to_be_clickable((By.ID, "btnPrint")))
                    print_btn.click()
                    
                    # ----------------------------------------------------
                    # [V28 수정] 파일 생성 감지 대기 시간을 10초 -> 30초로 증가
                    # ----------------------------------------------------
                    after_files = set(glob.glob(os.path.join(temp_save_dir, "*.pdf")))
                    start_time = time.time()
                    
                    # 최대 30초까지 파일이 새로 생성되기를 명시적으로 기다립니다.
                    while time.time() - start_time < 30: 
                        current_files = set(glob.glob(os.path.join(temp_save_dir, "*.pdf")))
                        new_files = list(current_files - after_files)
                        if new_files:
                            break
                        time.sleep(0.5) 
                    
                    # 파일 저장 및 이름 변경 (기존 로직)
                    if new_files:
                        latest_file = max(new_files, key=os.path.getctime)
                        new_name = os.path.join(temp_save_dir, f"{tracking_number}.pdf")
                        
                        try:
                            time.sleep(0.5) 
                            shutil.move(latest_file, new_name) 
                            log_and_update(f"→ 저장 완료: {tracking_number}.pdf")
                            successful_files.append(new_name)
                        except Exception as e:
                            log_and_update(f"→ 파일명 변경 에러: {e}")
                    else:
                        log_and_update(f"→ PDF 생성 안됨 (Timecheck)")

                except Exception as e:
                    log_and_update(f"→ 오류 발생! 상세 에러: {e}")
                    continue

            # 작업 완료 후 ZIP 파일 생성
            zip_buffer = BytesIO()
            zip_file_name = f"epost_tracking_results_{time.strftime('%Y%m%d_%H%M%S')}.zip"
            with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                for file_path in successful_files:
                    if os.path.exists(file_path): 
                        zip_file.write(file_path, os.path.basename(file_path))
            
            st.session_state.zip_download_data = {
                'name': zip_file_name, 
                'data': zip_buffer.getvalue(),
                'count': len(successful_files)
            }
            log_and_update("ZIP 파일 생성 완료.")
            
        except Exception as e:
            log_and_update("치명적 예외 발생:\n" + traceback.format_exc())
            st.session_state.zip_download_data = {'count': 0, 'error': True}
            
        finally:
            if driver:
                driver.quit()
            log_and_update("크롬 드라이버 종료")
            st.session_state.is_running = False
            log_and_update("작업 완료")


# =========================
# 3. Streamlit UI
# =========================
def main():
    st.set_page_config(page_title="우체국 등기 조회 웹앱", layout="centered")
    st.title("📮 우체국 등기 조회 자동화 (Streamlit)")
    
    st.info("💡 **배포 환경:** 클라우드에서는 브라우저 화면이 보이지 않습니다. 작업이 완료될 때까지 잠시 기다려 주세요.")
    st.warning("⚠️ **주의:** '작업 시작' 버튼을 누르면 Selenium 작업이 완료될 때까지 **화면이 잠깁니다.** 작업 중에는 브라우저를 닫지 마세요.")
    st.markdown("---")

    is_running = st.session_state.is_running

    # 1. 입력: 엑셀 파일 업로드 
    uploaded_file = st.file_uploader(
        "**1. 등기번호 엑셀 파일 업로드** (컬럼명: '등기번호')",
        type=["xlsx", "xls"],
        disabled=is_running 
    )
    st.markdown("---")
    
    # 2. 버튼
    col1, col2 = st.columns([1, 1])
    with col1:
        start_button = st.button("🚀 작업 시작", 
                                 type="primary", 
                                 disabled=(uploaded_file is None) or is_running)
    with col2:
        if not is_running:
             st.button("🔄 상태 초기화", on_click=clear_log)
    
    st.subheader("로그")
    log_placeholder = st.empty()
    log_placeholder.code('\n'.join(st.session_state.log_messages), language='text')

    # '시작' 버튼 클릭 이벤트 처리
    if start_button and uploaded_file:
        # Streamlit Spinner를 사용하여 작업 중임을 사용자에게 알림
        with st.spinner('Selenium 작업 진행 중... (클라우드 환경에서는 시간이 다소 소요될 수 있습니다)'):
            run_selenium_process(uploaded_file.read(), log_placeholder) 
        
        # 작업이 끝나면 (Spinner 종료) 다운로드 섹션을 보여주기 위해 RERUN
        st.rerun() 

    # 4. 결과 출력 및 다운로드 버튼 표시
    download_data = st.session_state.zip_download_data
    
    if download_data and not is_running:
        st.subheader("✅ 작업 결과")
        
        if download_data['count'] > 0:
            st.download_button(
                label=f"⬇️ {download_data['count']}개 PDF 파일 전체 다운로드 (ZIP)",
                data=download_data['data'],
                file_name=download_data['name'],
                mime="application/zip",
            )
            
            st.success(f"총 **{download_data['count']}개**의 등기번호 조회가 완료되었습니다. ZIP 파일을 다운로드하세요.")
            
        elif 'error' in download_data:
             st.error("치명적인 오류로 작업이 종료되었습니다. 로그를 확인해 주세요.")
        else:
             st.error("엑셀에 '등기번호'가 없거나, 조회된 PDF 파일이 없습니다. 로그를 확인해 주세요.")
            

if __name__ == '__main__':
    main()