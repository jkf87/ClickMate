# 필요한 라이브러리 임포트
import os  # 파일/디렉토리 조작
import json  # JSON 파일 처리
import time  # 시간 지연
import threading  # 멀티스레딩 
import keyboard  # 키보드 입력 감지
import win32api  # 윈도우 API 접근
import pyautogui  # 마우스/키보드 자동화
import numpy as np  # 배열 처리
from PIL import Image  # 이미지 처리
import cv2  # 컴퓨터 비전
from datetime import datetime  # 날짜/시간 처리

class AutomationScenario:
    """화면 자동화 시나리오를 생성하고 실행하는 클래스"""

    def __init__(self):
        # 시나리오 저장을 위한 딕셔너리
        self.scenarios = {}
        # 시나리오 실행 상태
        self.running = True
        # 시나리오 파일 저장 폴더
        self.scenario_folder = "scenarios"
        # 시나리오 폴더 생성
        os.makedirs(self.scenario_folder, exist_ok=True)
        
    def get_mouse_position(self):
        """마우스 우클릭 위치를 감지하여 좌표 반환"""
        # 마우스 좌/우 버튼 상태 초기화
        state_left = win32api.GetKeyState(0x01)
        state_right = win32api.GetKeyState(0x02)
        
        while True:
            # 우클릭 상태 변화 감지
            new_state_right = win32api.GetKeyState(0x02)
            if new_state_right != state_right:
                state_right = new_state_right
                if new_state_right < 0:  # 우클릭 감지
                    x, y = pyautogui.position()
                    return x, y
                    
            # ESC 키로 취소
            if keyboard.is_pressed('esc'):
                return None
                
            time.sleep(0.001)

    def capture_screen_region(self):
        """화면의 특정 영역을 캡처"""
        print("화면 영역을 지정하기 위해 마우스 오른쪽 버튼으로 왼쪽 상단 지점을 선택하세요.")
        top_left = self.get_mouse_position()
        if not top_left:
            return None
        print(f"선택된 좌표: {top_left}")
        
        print("이제 오른쪽 하단 지점을 선택하세요.")
        bottom_right = self.get_mouse_position()
        if not bottom_right:
            return None
        print(f"선택된 좌표: {bottom_right}")
        
        # 선택된 영역 캡처
        x1, y1 = top_left
        x2, y2 = bottom_right
        screenshot = pyautogui.screenshot(region=(min(x1, x2), min(y1, y2), 
                                                abs(x2-x1), abs(y2-y1)))
        return screenshot

    def create_scenario(self):
        """새로운 자동화 시나리오 생성"""
        scenario_name = input("시나리오 이름을 입력하세요: ").strip()
        if not scenario_name:
            print("시나리오 이름은 비워둘 수 없습니다.")
            return
            
        # 시나리오 데이터 구조 초기화
        scenario_data = {"triggers": []}
        
        while True:
            print("\n1. 트리거-액션 쌍 추가")
            print("2. 저장 후 종료")
            choice = input("선택하세요: ").strip()
            
            if choice == "1":
                # 트리거 이미지 캡처
                print("\n트리거 영역을 캡처합니다...")
                trigger_img = self.capture_screen_region()
                if trigger_img is None:
                    print("캡처가 취소되었습니다.")
                    continue
                
                # 캡처된 이미지 저장
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                img_filename = f"trigger_{timestamp}.png"
                img_path = os.path.join(self.scenario_folder, img_filename)
                trigger_img.save(img_path)
                
                # 클릭 위치 지정
                click_positions = []
                print("\n클릭할 위치들을 순서대로 오른쪽 클릭으로 지정하세요. (ESC로 종료)")
                while True:
                    pos = self.get_mouse_position()
                    if pos is None:
                        break
                    click_positions.append(pos)
                    print(f"클릭 위치 추가됨: {pos}")
                
                # 트리거-액션 쌍 추가
                scenario_data["triggers"].append({
                    "image": img_filename,
                    "clicks": click_positions
                })
                
            elif choice == "2":
                # 시나리오를 JSON 파일로 저장
                scenario_path = os.path.join(self.scenario_folder, f"{scenario_name}.json")
                with open(scenario_path, 'w', encoding='utf-8') as f:
                    json.dump(scenario_data, f, ensure_ascii=False, indent=2)
                print(f"\n시나리오가 저장되었습니다: {scenario_path}")
                break
                
    def load_scenarios(self):
        """저장된 시나리오 파일들을 로드"""
        scenarios = {}
        if not os.path.exists(self.scenario_folder):
            return scenarios
            
        # 모든 JSON 파일 로드
        for filename in os.listdir(self.scenario_folder):
            if filename.endswith('.json'):
                try:
                    filepath = os.path.join(self.scenario_folder, filename)
                    with open(filepath, 'r', encoding='utf-8') as f:
                        scenario_data = json.load(f)
                        scenarios[filename[:-5]] = scenario_data
                except Exception as e:
                    print(f"시나리오 로드 중 오류 발생: {filename} - {str(e)}")
        return scenarios

    def monitor_trigger(self, trigger_img_path, click_positions):
        """트리거 이미지를 모니터링하고 일치하면 클릭 수행"""
        # 트리거 이미지 로드
        trigger_img = Image.open(os.path.join(self.scenario_folder, trigger_img_path))
        trigger_array = np.array(trigger_img)
        
        while self.running:
            try:
                # 현재 화면 캡처
                screenshot = pyautogui.screenshot()
                screenshot_array = np.array(screenshot)
                
                # 이미지 매칭 수행
                result = cv2.matchTemplate(
                    screenshot_array, 
                    trigger_array, 
                    cv2.TM_CCOEFF_NORMED
                )
                
                min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(result)
                
                # 매칭률이 70% 이상이면 클릭 수행
                if max_val > 0.7:
                    for pos in click_positions:
                        if not self.running:
                            break
                        pyautogui.click(x=pos[0], y=pos[1])
                        time.sleep(0.5)
                    time.sleep(1)  # 다음 감지까지 대기
                    
            except Exception as e:
                print(f"모니터링 중 오류 발생: {str(e)}")
                time.sleep(1)
                
            time.sleep(0.1)

    def run_scenario(self, scenario_name):
        """선택된 시나리오 실행"""
        if scenario_name not in self.scenarios:
            print("시나리오를 찾을 수 없습니다.")
            return
            
        scenario_data = self.scenarios[scenario_name]
        threads = []
        
        print(f"\n시나리오 '{scenario_name}' 실행 중... (Ctrl+C로 중지)")
        self.running = True
        
        try:
            # 각 트리거에 대해 모니터링 스레드 생성
            for trigger in scenario_data["triggers"]:
                img_path = trigger["image"]
                click_positions = trigger["clicks"]
                
                thread = threading.Thread(
                    target=self.monitor_trigger,
                    args=(img_path, click_positions)
                )
                thread.daemon = True
                thread.start()
                threads.append(thread)
            
            # Ctrl+C 입력 대기
            keyboard.wait('ctrl+c')
            
        except KeyboardInterrupt:
            pass
        finally:
            # 모든 스레드 종료
            self.running = False
            for thread in threads:
                thread.join()
            print("\n시나리오 실행이 중지되었습니다.")

    def main_menu(self):
        """메인 메뉴 UI"""
        print("\n=== 클릭메이트 (ClickMate) v1.0 ===")
        print("화면 자동화 도우미 프로그램")
        
        while True:
            print("\n=== 메인 메뉴 ===")
            print("1. 새 시나리오 생성")
            print("2. 시나리오 실행")
            print("3. 종료")
            
            choice = input("선택하세요: ").strip()
            
            if choice == "1":
                self.create_scenario()
                
            elif choice == "2":
                # 저장된 시나리오 목록 표시
                self.scenarios = self.load_scenarios()
                if not self.scenarios:
                    print("저장된 시나리오가 없습니다.")
                    continue
                    
                print("\n=== 저장된 시나리오 목록 ===")
                for idx, name in enumerate(self.scenarios.keys(), 1):
                    print(f"{idx}. {name}")
                    
                # 시나리오 선택 및 실행
                try:
                    selection = int(input("\n실행할 시나리오 번호를 입력하세요: "))
                    if 1 <= selection <= len(self.scenarios):
                        scenario_name = list(self.scenarios.keys())[selection-1]
                        self.run_scenario(scenario_name)
                    else:
                        print("잘못된 번호입니다.")
                except ValueError:
                    print("올바른 숫자를 입력하세요.")
                    
            elif choice == "3":
                print("\n클릭메이트를 종료합니다. 이용해 주셔서 감사합니다.")
                break
                
            else:
                print("잘못된 선택입니다.")

# 프로그램 시작점
if __name__ == "__main__":
    automation = AutomationScenario()
    automation.main_menu() 