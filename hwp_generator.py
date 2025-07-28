import win32com.client
import os
import calendar
import json
import sys
from datetime import datetime
import glob

class HwpGenerator:
    def __init__(self):
        """HWP 파일 생성기 - 같은 폴더의 파일들 사용"""
        # 현재 스크립트 위치 기준으로 파일 경로 설정
        script_dir = os.path.dirname(os.path.abspath(__file__))
        self.template_path = os.path.join(script_dir, "청년이룸출근부.hwp")
        self.data_file = os.path.join(script_dir, "청년이룸출근부_데이터.json")
        self.output_dir = script_dir  # 현재 폴더
        self.hwp = None
        
    def initialize_hwp(self):
        """한글 오피스 객체 초기화"""
        try:
            print("HWP 객체 초기화 시작...")
            self.hwp = win32com.client.gencache.EnsureDispatch("HWPFrame.HwpObject")
            print("HWP 객체 생성 완료")
            self.hwp.RegisterModule("FilePathCheckDLL", "SecurityModule")
            print("HWP 모듈 등록 완료")
            return True
        except Exception as e:
            print(f"HWP 초기화 실패: {e}")
            return False
    
    def open_template(self):
        """템플릿 파일 열기"""
        try:
            print(f"템플릿 파일 열기: {self.template_path}")
            print(f"파일 존재 여부: {os.path.exists(self.template_path)}")
            self.hwp.Open(self.template_path)
            print("템플릿 파일 열기 성공")
            return True
        except Exception as e:
            print(f"템플릿 파일 열기 실패: {e}")
            return False
    
    def find_and_replace(self, replace_dict):
        """문서 내 텍스트 검색 및 바꾸기"""
        print(f"치환할 항목들: {replace_dict}")
        
        for find_text, replace_text in replace_dict.items():
            try:
                print(f"치환 시도: {find_text} -> {replace_text}")
                
                # 문서 맨 위로 이동
                self.hwp.HAction.Run("MoveTop")
                
                # 찾기/바꾸기 설정
                self.hwp.HAction.GetDefault("RepeatFind", self.hwp.HParameterSet.HFindReplace.HSet)
                self.hwp.HParameterSet.HFindReplace.FindString = find_text
                self.hwp.HParameterSet.HFindReplace.ReplaceString = replace_text
                self.hwp.HParameterSet.HFindReplace.ReplaceMode = 1  # 모두 바꾸기
                self.hwp.HParameterSet.HFindReplace.IgnoreMessage = 1  # 메시지 창 숨김
                
                # 치환 실행
                result = self.hwp.HAction.Execute("AllReplace", self.hwp.HParameterSet.HFindReplace.HSet)
                print(f"치환 결과: {result}")
                
            except Exception as e:
                print(f"텍스트 치환 실패 ({find_text} -> {replace_text}): {e}")
    
    def remove_excess_days(self, end_day):
        """endday 이후의 일자들을 지우기"""
        print(f"endday 이후 일자 제거 시작: {end_day}일 이후")
        
        try:
            # 31일부터 endday까지 역순으로 처리
            for day in range(31, end_day, -1):
                day_str = str(day)
                print(f"일자 {day_str} 제거 시도")
                
                # 문서 맨 위로 이동
                self.hwp.HAction.Run("MoveTop")
                
                # 찾기/바꾸기 설정
                self.hwp.HAction.GetDefault("RepeatFind", self.hwp.HParameterSet.HFindReplace.HSet)
                self.hwp.HParameterSet.HFindReplace.FindString = day_str
                self.hwp.HParameterSet.HFindReplace.ReplaceString = ""  # 빈 문자열로 치환 (삭제)
                self.hwp.HParameterSet.HFindReplace.ReplaceMode = 1  # 모두 바꾸기
                self.hwp.HParameterSet.HFindReplace.IgnoreMessage = 1  # 메시지 창 숨김
                
                # 치환 실행
                result = self.hwp.HAction.Execute("AllReplace", self.hwp.HParameterSet.HFindReplace.HSet)
                print(f"일자 {day_str} 제거 결과: {result}")
                
        except Exception as e:
            print(f"일자 제거 실패: {e}")
    
    def is_non_working_day(self, day, year, month, working_saturdays, substitute_holidays):
        """근무하지 않는 날인지 판단"""
        try:
            # 해당 날짜의 요일 계산
            date_obj = datetime(year, month, day)
            day_of_week = date_obj.weekday()  # 0=월요일, 6=일요일
            
            # 일요일인 경우 (6)
            if day_of_week == 6:
                return True
            
            # 토요일인 경우 (5) - 근무하는 토요일이 아닌 경우만
            if day_of_week == 5:
                # 해당 날짜가 근무하는 토요일인지 확인
                return day not in working_saturdays
            
            # 대체 휴무일인 경우
            date_key = f"{year}-{month:02d}-{day:02d}"
            if date_key in substitute_holidays:
                return True
            
            return False
            
        except Exception as e:
            print(f"근무일 판단 실패 (날짜: {day}): {e}")
            return False
    
    def get_non_working_days(self, year, month, overtime_schedules, substitute_holidays):
        """근무하지 않는 날짜 목록 반환"""
        non_working_days = []
        
        try:
            # 해당 월의 마지막 날 계산
            last_day = calendar.monthrange(year, month)[1]
            
            for day in range(1, last_day + 1):
                if self.is_non_working_day(day, year, month, overtime_schedules, substitute_holidays):
                    non_working_days.append(day)
                    print(f"근무하지 않는 날 발견: {month}월 {day}일")
            
            return non_working_days
            
        except Exception as e:
            print(f"근무하지 않는 날 계산 실패: {e}")
            return []
    
    def add_diagonal_lines(self, year, month, last_day, overtime_schedules, substitute_holidays):
        """근무하지 않는 날에 빗금 추가"""
        print(f"빗금 추가 시작: {year}년 {month}월")
        
        try:
            # 근무하지 않는 날짜 목록 가져오기
            non_working_days = self.get_non_working_days(year, month, overtime_schedules, substitute_holidays)
            
            if not non_working_days:
                print("근무하지 않는 날이 없습니다.")
                return
            
            print(f"근무하지 않는 날 발견: {non_working_days}")
            
            # 리스트를 set으로 변환하여 효율적인 비교
            non_working_days_set = set(non_working_days)
            print(f"근무하지 않는 날 set: {non_working_days_set}")
            
            # "%일1"과 "%일2" 헤더를 찾아서 각 열을 순회
            self.add_diagonal_lines_to_columns(non_working_days_set)
                    
        except Exception as e:
            print(f"빗금 추가 실패: {e}")
    
    def add_diagonal_lines_with_days(self, year, month, last_day, non_working_days):
        """주어진 근무하지 않는 날 목록으로 빗금 추가"""
        print(f"빗금 추가 시작: {year}년 {month}월")
        
        try:
            if not non_working_days:
                print("근무하지 않는 날이 없습니다.")
                return
            
            print(f"근무하지 않는 날 발견: {non_working_days}")
            
            # 리스트를 set으로 변환하여 효율적인 비교
            non_working_days_set = set(non_working_days)
            print(f"근무하지 않는 날 set: {non_working_days_set}")
            
            # "%일1"과 "%일2" 헤더를 찾아서 각 열을 순회
            self.add_diagonal_lines_to_columns(non_working_days_set)
                    
        except Exception as e:
            print(f"빗금 추가 실패: {e}")
    
    def find_text(self, search_text):
        """특정 문자열을 문서에서 찾음"""
        self.hwp.HAction.Run("MoveTop")
        self.hwp.HAction.GetDefault("RepeatFind", self.hwp.HParameterSet.HFindReplace.HSet)
        self.hwp.HParameterSet.HFindReplace.FindString = search_text
        self.hwp.HParameterSet.HFindReplace.Direction = 1  # 아래 방향 검색
        return self.hwp.HAction.Execute("RepeatFind", self.hwp.HParameterSet.HFindReplace.HSet)

    def select_cell(self):
        """현재 위치의 셀을 지정"""
        self.hwp.HAction.Run("TableCellBlock")

    def find_and_select_cell(self, search_text):
        """특정 문자열을 찾아 해당 셀을 지정"""
        if self.find_text(search_text):
            self.select_cell()
            return True
        return False

    def move_cell(self, direction, steps=1):
        """현재 지정된 셀을 상하좌우로 주어진 횟수만큼 이동"""
        directions = {
            "up": "UppperCell",
            "down": "LowerCell",
            "left": "LeftCell",
            "right": "RightCell"
        }
        if direction not in directions:
            raise ValueError("Invalid direction. Use 'up', 'down', 'left', or 'right'.")
        if not isinstance(steps, int) or steps < 1:
            raise ValueError("Steps must be a positive integer.")
        for _ in range(steps):
            self.hwp.HAction.Run("Table" + directions[direction])
        self.hwp.HAction.Run("TableCellBlock")  # 이동 후 현재 셀 다시 지정

    def diagonal_cell(self):
        """현재 지정된 셀에 대각선을 긋는 함수"""
        self.hwp.HAction.Run("TableCellBorderDiagonalUp")
    
    def add_diagonal_line_for_day_in_table(self, day):
        """표에서 특정 날짜의 행에 빗금 추가"""
        try:
            day_str = str(day)
            print(f"표에서 날짜 {day_str}에 빗금 추가 시도")
            
            # 문서 맨 위로 이동
            self.hwp.HAction.Run("MoveTop")
            
            # 여러 가능한 헤더 텍스트 시도
            header_texts = ["월/일", "%일1", "%일2", "일1", "일2"]
            
            found_header = False
            for header_text in header_texts:
                try:
                    print(f"헤더 '{header_text}' 검색 중...")
                    
                    # 문서 맨 위로 이동
                    self.hwp.HAction.Run("MoveTop")
                    
                    # 헤더 찾기
                    self.hwp.HAction.GetDefault("RepeatFind", self.hwp.HParameterSet.HFindReplace.HSet)
                    self.hwp.HParameterSet.HFindReplace.FindString = header_text
                    self.hwp.HParameterSet.HFindReplace.ReplaceMode = 0  # 찾기만
                    self.hwp.HParameterSet.HFindReplace.IgnoreMessage = 1
                    
                    # 헤더 찾기
                    found_header = self.hwp.HAction.Execute("Find", self.hwp.HParameterSet.HFindReplace.HSet)
                    
                    if found_header:
                        print(f"헤더 '{header_text}' 찾음, 날짜 {day_str} 검색 중...")
                        break
                        
                except Exception as header_error:
                    print(f"헤더 '{header_text}' 검색 실패: {header_error}")
                    continue
            
            if found_header:
                print(f"'월/일' 헤더 찾음, 날짜 {day_str} 검색 중...")
                
                # 첫 번째 "월/일" 아래 16칸에서 날짜 찾기
                if day <= 16:
                    # 첫 번째 "월/일" 아래에서 날짜 찾기
                    self.find_date_in_first_column(day_str)
                else:
                    # 두 번째 "월/일" 찾기
                    found_second_header = self.hwp.HAction.Execute("Find", self.hwp.HParameterSet.HFindReplace.HSet)
                    if found_second_header:
                        print(f"두 번째 '월/일' 헤더 찾음, 날짜 {day_str} 검색 중...")
                        self.find_date_in_second_column(day_str)
                    else:
                        print(f"두 번째 '월/일' 헤더를 찾을 수 없음")
                        
            else:
                print(f"'월/일' 헤더를 찾을 수 없음")
                
        except Exception as e:
            print(f"날짜 {day} 빗금 추가 실패: {e}")
    
    def find_date_in_first_column(self, day_str):
        """첫 번째 "월/일" 열에서 날짜 찾기 (1-16일)"""
        try:
            # 첫 번째 "월/일" 아래 16칸에서 날짜 찾기
            for row in range(16):
                try:
                    # 아래로 이동
                    self.hwp.HAction.Run("MoveDown")
                    
                    # find_text 함수를 사용해서 날짜 찾기
                    found_date = self.find_text(day_str)
                    
                    if found_date:
                        print(f"첫 번째 열에서 날짜 {day_str} 찾음, 셀 지정 중...")
                        self.select_cell()
                        print(f"우측 5칸에 빗금 추가 중...")
                        self.add_diagonal_lines_to_cells()
                        return True
                        
                except Exception as row_error:
                    print(f"행 {row} 검색 실패: {row_error}")
                    continue
            
            print(f"첫 번째 열에서 날짜 {day_str}를 찾을 수 없음")
            return False
            
        except Exception as e:
            print(f"첫 번째 열 검색 실패: {e}")
            return False
    
    def find_date_in_second_column(self, day_str):
        """두 번째 "월/일" 열에서 날짜 찾기 (17일 이후)"""
        try:
            # 두 번째 "월/일" 아래에서 날짜 찾기
            for row in range(31):  # 최대 31일까지
                try:
                    # 아래로 이동
                    self.hwp.HAction.Run("MoveDown")
                    
                    # find_text 함수를 사용해서 날짜 찾기
                    found_date = self.find_text(day_str)
                    
                    if found_date:
                        print(f"두 번째 열에서 날짜 {day_str} 찾음, 셀 지정 중...")
                        self.select_cell()
                        print(f"우측 5칸에 빗금 추가 중...")
                        self.add_diagonal_lines_to_cells()
                        return True
                        
                except Exception as row_error:
                    print(f"행 {row} 검색 실패: {row_error}")
                    continue
            
            print(f"두 번째 열에서 날짜 {day_str}를 찾을 수 없음")
            return False
            
        except Exception as e:
            print(f"두 번째 열 검색 실패: {e}")
            return False
    
    def add_diagonal_lines_to_columns(self, non_working_days_set):
        """%일1과 %일2 열을 찾아서 근무하지 않는 날에 빗금 추가"""
        try:
            # %일1과 %일2 헤더 찾기
            column_headers = ["%일1", "%일2"]
            
            for header in column_headers:
                print(f"헤더 '{header}' 검색 중...")
                
                # 문서 맨 위로 이동
                self.hwp.HAction.Run("MoveTop")
                
                # 헤더 찾기
                found_header = self.find_text(header)
                
                if found_header:
                    print(f"헤더 '{header}' 찾음, 열 순회 시작...")
                    self.select_cell()
                    
                    # 해당 열을 순회하면서 근무하지 않는 날 찾기
                    self.process_column_for_diagonal_lines(header, non_working_days_set)
                    
                    # 다음 헤더 검색을 위해 잠시 대기
                    import time
                    time.sleep(0.1)
                else:
                    print(f"헤더 '{header}'를 찾을 수 없음")
                    
        except Exception as e:
            print(f"열별 빗금 추가 실패: {e}")
    
    def process_column_for_diagonal_lines(self, header, non_working_days_set):
        """특정 열을 순회하면서 근무하지 않는 날에 빗금 추가"""
        try:
            print(f"열 '{header}'에서 근무하지 않는 날 검색 시작...")
            
            # 헤더를 찾아서 셀 지정
            self.find_and_select_cell(header)
            
            # 각 열의 설정
            if header == "%일1":
                max_rows = 16
                print(f"열 '%일1': 16칸 순회 시작")
            elif header == "%일2":
                max_rows = 15
                print(f"열 '%일2': 15칸 순회 시작")
            else:
                print(f"알 수 없는 헤더: {header}")
                return
            
            # 각 행을 순회
            for row in range(max_rows):
                try:
                    print(f"행 {row+1} 처리 중...")
                    
                    # 아래로 이동
                    self.move_cell("down", 1)
                    
                    # 셀 선택
                    self.select_cell()
                    
                    # 셀 내용 초기화 (참고 코드와 동일)
                    self.hwp.HAction.Run("TableCellInput")
                    
                    # 셀 내용 읽기 (참고 코드와 동일한 방식)
                    self.hwp.HAction.Run("TableCellInput")
                    self.hwp.InitScan(0, 2)
                    text_tuple = self.hwp.GetText()
                    
                    print(f"행 {row+1}: 읽은 텍스트 = '{text_tuple}'")
                    
                    try:
                        # 참고 코드와 동일한 방식으로 두 번째 요소 사용
                        if isinstance(text_tuple, tuple) and len(text_tuple) > 1:
                            cell_date = int(text_tuple[1])
                        else:
                            print(f"행 {row+1}: 유효한 텍스트가 없음 - 건너뜀")
                            continue
                        
                        print(f"행 {row+1}: 추출된 날짜 = {cell_date}")
                        
                        # 근무하지 않는 날인지 확인
                        if cell_date in non_working_days_set:
                            print(f"열 '{header}'에서 근무하지 않는 날 {cell_date} 발견, 빗금 추가 중...")
                            self.apply_diagonal_to_weekend()
                        else:
                            print(f"행 {row+1}: 날짜 {cell_date}는 근무하는 날입니다. (빗금 추가 안함)")
                            
                    except Exception as date_error:
                        print(f"행 {row+1}: 날짜 처리 실패: {date_error}")
                        continue
                        
                except Exception as row_error:
                    print(f"행 {row+1} 처리 실패: {row_error}")
                    continue
                    
            print(f"열 '{header}' 처리 완료")
                    
        except Exception as e:
            print(f"열 '{header}' 처리 실패: {e}")
    

    
    def add_diagonal_lines_to_cells(self):
        """현재 위치에서 우측 5칸에 빗금 추가"""
        try:
            # 현재 위치에서 5개 셀에 빗금 추가
            for col in range(5):  # 출근, 결근, 지각/조퇴, 근로자, 담당 열
                try:
                    # 셀 선택 (우측으로 이동)
                    self.move_cell("right", 1)
                    
                    # 셀에 빗금 추가 (좌하단에서 우상단으로)
                    self.diagonal_cell()
                    
                except Exception as col_error:
                    print(f"열 {col} 빗금 추가 실패: {col_error}")
                    continue
                    
        except Exception as e:
            print(f"빗금 추가 실패: {e}")
    
    def apply_diagonal_to_weekend(self):
        """현재 셀에서 5칸 오른쪽으로 이동하며 대각선 적용"""
        try:
            print("우측 5칸에 빗금 추가 시작...")
            for col in range(5):  # 출근, 결근, 지각/조퇴, 근로자, 담당 열
                try:
                    # 우측으로 이동
                    self.move_cell("right", 1)
                    print(f"열 {col+1}에 빗금 추가 중...")
                    # 셀에 빗금 추가 (좌하단에서 우상단으로)
                    self.diagonal_cell()
                    
                except Exception as col_error:
                    print(f"열 {col+1} 빗금 추가 실패: {col_error}")
                    continue
            
            # 원래 위치로 복귀
            self.move_cell("left", 5)
            print("빗금 추가 완료, 원래 위치로 복귀")
            
        except Exception as e:
            print(f"빗금 추가 실패: {e}")
    

    
    def save_file(self, filename):
        """파일 저장"""
        try:
            output_path = os.path.join(self.output_dir, filename)
            print(f"저장 경로: {output_path}")
            
            # HWP 파일 형식으로 저장
            if not filename.endswith('.hwp'):
                output_path = output_path + '.hwp'
            
            self.hwp.SaveAs(output_path)
            print(f"파일 저장 완료: {filename}")
            return True
        except Exception as e:
            print(f"파일 저장 실패: {e}")
            return False
    
    def close_hwp(self):
        """한글 오피스 종료"""
        if self.hwp:
            self.hwp.Quit()
    
    def generate_replace_dict(self, manager_name, target_date):
        """치환할 데이터 딕셔너리 생성"""
        try:
            target_date_obj = datetime.strptime(target_date, "%Y-%m")
            year = target_date_obj.year
            month = target_date_obj.month
            last_day = calendar.monthrange(year, month)[1]
            
            return {
                "%Name": manager_name,
                "%Year": str(year),
                "%Month": str(month),
                "%Endday": str(last_day),
                "%일1": "월/일",
                "%일2": "월/일"
            }
        except Exception as e:
            print(f"치환 딕셔너리 생성 실패: {e}")
            return {}
    
    def get_manager_specific_data(self, manager_name, data):
        """특정 매니저의 개별 데이터 추출"""
        try:
            print(f"=== 매니저 {manager_name} 데이터 추출 시작 ===")
            
            # 전체 데이터에서 매니저별 데이터 찾기
            managers = data.get('managers', [])
            print(f"전체 매니저 목록: {managers}")
            
            # 매니저가 존재하는지 확인
            manager_exists = False
            for manager in managers:
                if isinstance(manager, dict) and manager.get('name') == manager_name:
                    manager_exists = True
                    break
            
            if not manager_exists:
                print(f"매니저 {manager_name}이 목록에 없습니다.")
                return {}, {}
            
            # 전체 overtimeSchedules와 substituteHolidays에서 해당 매니저의 데이터만 추출
            overtime_schedules = data.get('overtimeSchedules', {})
            substitute_holidays = data.get('substituteHolidays', {})
            
            # 매니저별 overtime 스케줄 추출 (근무하는 토요일들)
            manager_working_saturdays = set()  # 근무하는 토요일 날짜들
            for date, schedule_list in overtime_schedules.items():
                for schedule in schedule_list:
                    if isinstance(schedule, dict) and schedule.get('name') == manager_name:
                        date_obj = datetime.strptime(date, "%Y-%m-%d")
                        weekday = date_obj.weekday()  # 5 = 토요일
                        if weekday == 5:  # 토요일
                            manager_working_saturdays.add(date_obj.day)
            
            # 매니저별 대체휴무 추출
            manager_substitute = {}
            for date, holiday_list in substitute_holidays.items():
                for holiday in holiday_list:
                    if isinstance(holiday, dict) and holiday.get('name') == manager_name:
                        manager_substitute[date] = "대체휴무"
            
            print(f"매니저 {manager_name}의 개별 데이터:")
            print(f"- 근무하는 토요일: {manager_working_saturdays}")
            print(f"- 개별 대체휴무: {manager_substitute}")
            
            return manager_working_saturdays, manager_substitute
                
        except Exception as e:
            print(f"매니저 {manager_name} 데이터 추출 실패: {e}")
            return set(), {}
    
    def get_manager_non_working_days(self, manager_name, year, month, data):
        """특정 매니저의 근무하지 않는 날 계산"""
        try:
            print(f"매니저 {manager_name}의 근무하지 않는 날 계산 시작...")
            
            # 매니저별 개별 데이터 추출
            manager_working_saturdays, manager_substitute = self.get_manager_specific_data(manager_name, data)
            
            # 해당 월의 마지막 날 계산
            last_day = calendar.monthrange(year, month)[1]
            non_working_days = []
            
            for day in range(1, last_day + 1):
                if self.is_non_working_day(day, year, month, manager_working_saturdays, manager_substitute):
                    non_working_days.append(day)
                    print(f"매니저 {manager_name}: 근무하지 않는 날 발견 - {month}월 {day}일")
            
            print(f"매니저 {manager_name}의 근무하지 않는 날 목록: {non_working_days}")
            return non_working_days
            
        except Exception as e:
            print(f"매니저 {manager_name} 근무하지 않는 날 계산 실패: {e}")
            return []
    
    def generate_hwp_for_manager(self, manager_name, target_date, data):
        """특정 매니저의 HWP 파일 생성"""
        try:
            # 파일명 생성
            target_date_obj = datetime.strptime(target_date, "%Y-%m")
            year_suffix = str(target_date_obj.year)[-2:]
            month_str = str(target_date_obj.month).zfill(2)
            filename = f"청년이룸출근부_{manager_name}_{year_suffix}{month_str}.hwp"
            
            # 템플릿 파일 열기
            if not self.open_template():
                return False
            
            # last_day 계산
            year = target_date_obj.year
            month = target_date_obj.month
            last_day = calendar.monthrange(year, month)[1]
            
            # 매니저별 근무하지 않는 날 계산
            non_working_days = self.get_manager_non_working_days(manager_name, year, month, data)
            
            # 근무하지 않는 날에 빗금 추가 (텍스트 치환 전에 수행)
            self.add_diagonal_lines_with_days(year, month, last_day, non_working_days)
            
            # 치환 데이터 생성
            replace_dict = self.generate_replace_dict(manager_name, target_date)
            
            # 텍스트 치환
            self.find_and_replace(replace_dict)
            
            # endday 이후 일자 제거
            self.remove_excess_days(last_day)
            
            # 파일 저장
            success = self.save_file(filename)
            
            return success
            
        except Exception as e:
            print(f"매니저 {manager_name}의 HWP 파일 생성 실패: {e}")
            return False
    
    def generate_hwp_for_all_managers(self, managers, target_date, data):
        """모든 매니저의 HWP 파일 생성"""
        if not self.initialize_hwp():
            return 0
        
        success_count = 0
        
        try:
            for manager in managers:
                manager_name = manager.get('name', manager) if isinstance(manager, dict) else manager
                print(f"매니저 {manager_name}의 HWP 파일 생성 중...")
                if self.generate_hwp_for_manager(manager_name, target_date, data):
                    success_count += 1
                else:
                    print(f"매니저 {manager_name}의 파일 생성 실패")
        finally:
            self.close_hwp()
        
        return success_count

def load_data():
    """데이터 파일 로드"""
    try:
        # 현재 스크립트 위치 기준으로 파일 경로 설정
        script_dir = os.path.dirname(os.path.abspath(__file__))
        data_file_path = os.path.join(script_dir, "청년이룸출근부_데이터.json")
        
        if not os.path.exists(data_file_path):
            print(f"청년이룸출근부_데이터.json 파일을 찾을 수 없습니다.")
            print(f"찾는 경로: {data_file_path}")
            return None
        
        with open(data_file_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
            print("데이터 파일 로드 완료")
            return data
                
    except Exception as e:
        print(f"데이터 파일 로드 실패: {e}")
        return None

def main():
    """메인 실행 함수"""
    print("=== 청년이룸 HWP 생성기 ===")
    
    # 데이터 로드
    data = load_data()
    if not data:
        print("데이터를 로드할 수 없습니다. 프로그램을 종료합니다.")
        return
    
    # 데이터에서 정보 추출
    managers = data.get('managers', [])
    target_date = data.get('targetDate', '2025-03')
    
    print(f"로드된 데이터:")
    print(f"- 매니저 수: {len(managers)}명")
    print(f"- 대상 월: {target_date}")
    
    # HWP 생성기 초기화
    generator = HwpGenerator()
    
    # 모든 매니저의 HWP 파일 생성
    print(f"\n총 {len(managers)}명의 매니저에 대해 HWP 파일을 생성합니다...")
    success_count = generator.generate_hwp_for_all_managers(managers, target_date, data)
    
    print(f"\n생성 완료: {success_count}/{len(managers)}개 파일")

if __name__ == "__main__":
    main() 