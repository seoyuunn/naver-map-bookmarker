import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, ElementClickInterceptedException
from webdriver_manager.chrome import ChromeDriverManager
import traceback

def naver_map_bookmark_automation(excel_file_path):
    """
    네이버 지도에서 엑셀 파일의 장소 정보를 검색하고 즐겨찾기에 추가하는 함수
    
    Args:
        excel_file_path (str): 장소 정보가 담긴 엑셀 파일 경로
    """
    # 엑셀 파일 읽기
    try:
        df = pd.read_excel(excel_file_path)
        print(f"엑셀 파일을 성공적으로 읽었습니다. 총 {len(df)}개의 항목이 있습니다.")
        # 처음 5개 행 출력하여 데이터 확인
        print("데이터 샘플:")
        print(df.head())
    except Exception as e:
        print(f"엑셀 파일을 읽는 중 오류가 발생했습니다: {e}")
        return
    
    # 상호명과 주소 컬럼이 있는지 확인
    required_columns = ['상호명', '주소']
    missing_columns = [col for col in required_columns if col not in df.columns]
    
    if missing_columns:
        print(f"엑셀 파일에 다음 필수 컬럼이 없습니다: {', '.join(missing_columns)}")
        print(f"현재 컬럼: {', '.join(df.columns)}")
        return
    
    # 크롬 드라이버 설정
    chrome_options = Options()
    chrome_options.add_argument("--start-maximized")  # 브라우저 창 최대화
    # chrome_options.add_argument("--headless")  # 필요시 헤드리스 모드 활성화
    
    # 크롬 드라이버 초기화
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
    wait = WebDriverWait(driver, 20)  # 최대 20초로 증가
    
    try:
        # 네이버 로그인 페이지 직접 접속
        driver.get("https://nid.naver.com/nidlogin.login")
        print("네이버 로그인 페이지로 이동했습니다.")
        
        # 로그인 대기 시간 증가
        print("🔐 네이버 로그인을 진행해주세요. 로그인 완료 후 엔터 키를 눌러주세요.")
        input("로그인 완료 후 엔터 키를 눌러주세요...")
        
        # 네이버 지도 페이지로 이동
        driver.get("https://map.naver.com")
        print("네이버 지도 페이지로 이동했습니다.")
        time.sleep(5)  # 페이지 로딩 대기
        
        # 각 장소에 대해 검색 및 즐겨찾기 추가
        success_count = 0
        failure_count = 0
        
        for index, row in df.iterrows():
            store_name = row['상호명']
            address = row['주소']
            search_query = f"{store_name} {address}"
            
            try:
                print(f"\n처리 중 ({index+1}/{len(df)}): {store_name} - {address}")
                
                # 검색창 찾기 및 검색어 입력 (여러 가능한 셀렉터 시도)
                try:
                    search_box = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "input.input_search")))
                except TimeoutException:
                    try:
                        search_box = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "input.search_input")))
                    except TimeoutException:
                        search_box = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[contains(@class, 'search')]")))
                
                # 기존 검색어 지우기
                search_box.clear()
                time.sleep(1)
                
                # 검색어 입력 및 검색
                search_box.send_keys(search_query)
                time.sleep(1)
                search_box.send_keys(Keys.ENTER)
                print(f"검색어 입력 완료: {search_query}")
                
                # 검색 결과 로딩 대기
                time.sleep(3)
                
                # 첫 번째 검색 결과 항목 클릭 (여러 셀렉터 시도)
                try:
                    first_result = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "li.search_item, li.item_search")))
                    first_result.click()
                    print("첫 번째 검색 결과 클릭 완료")
                except TimeoutException:
                    try:
                        first_result = wait.until(EC.element_to_be_clickable((By.XPATH, "//li[contains(@class, 'search') or contains(@class, 'item')][1]")))
                        first_result.click()
                        print("대체 셀렉터로 첫 번째 검색 결과 클릭 완료")
                    except:
                        print("첫 번째 검색 결과를 찾을 수 없습니다. 다음 항목으로 넘어갑니다.")
                        failure_count += 1
                        continue
                
                # 상세 정보 페이지 로딩 대기
                time.sleep(3)
                
                # 즐겨찾기 버튼 클릭 (여러 셀렉터 시도)
                try:
                    bookmark_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.btn_save, button.btn_bookmark")))
                    bookmark_button.click()
                    print("즐겨찾기 버튼 클릭 완료")
                except TimeoutException:
                    try:
                        # 대체 셀렉터 시도
                        bookmark_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(@class, 'save') or contains(@class, 'bookmark') or contains(@class, 'favorite')]")))
                        bookmark_button.click()
                        print("대체 셀렉터로 즐겨찾기 버튼 클릭 완료")
                    except Exception as e:
                        print(f"즐겨찾기 버튼을 찾을 수 없습니다: {e}")
                        failure_count += 1
                        continue
                
                print(f"✅ 성공: {store_name}")
                success_count += 1
                
                # 작업 간 딜레이
                time.sleep(2)
                
            except Exception as e:
                print(f"❌ 실패: {store_name} - 예상치 못한 오류")
                print(traceback.format_exc())  # 자세한 오류 정보 출력
                failure_count += 1
                
            # 작업 간 딜레이
            time.sleep(2)
        
        # 결과 요약
        print("\n===== 처리 결과 =====")
        print(f"총 항목: {len(df)}")
        print(f"성공: {success_count}")
        print(f"실패: {failure_count}")
        
    except Exception as e:
        print(f"전체 프로세스 실행 중 오류 발생: {e}")
        print(traceback.format_exc())  # 자세한 오류 정보 출력
    
    finally:
        # 종료 전 사용자 입력 기다리기
        input("\n처리가 완료되었습니다. 브라우저를 닫으려면 엔터 키를 누르세요...")
        # 브라우저 종료
        driver.quit()

if __name__ == "__main__":
    # 직접 경로 지정 (원하는 경로로 수정)
    excel_file_path = "/Users/yunseo/github/naver-map-bookmarker/seongnam.xlsx"
    # 또는 사용자 입력으로 경로 받기
    # excel_file_path = input("장소 정보가 있는 엑셀 파일의 경로를 입력하세요: ")
    naver_map_bookmark_automation(excel_file_path)