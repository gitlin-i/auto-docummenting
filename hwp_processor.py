import win32com.client
from eroom import MetaData, generate_replace_dict, EroomManagerSchedule
import os

class HwpProcessor:

    def __init__(self, meta_data):
        """
        HWP 문서 자동 처리를 담당하는 클래스

        :param meta_data: MetaData 객체
        """
        self.meta_data : MetaData = meta_data
        self.hwp = self._initialize_hwp()

    def _initialize_hwp(self):
        """한글 오피스 객체를 초기화하고 보안 모듈을 등록"""
        try:
            hwp = win32com.client.gencache.EnsureDispatch("HWPFrame.HwpObject")
            hwp.RegisterModule("FilePathCheckDLL", "SecurityModule")  # 보안 경고 방지
            return hwp
        except Exception as e:
            raise Exception(f"HWP 초기화 실패: {e}")

    def open_file(self):
        """HWP 파일 열기"""
        try:
            file_path = os.path.join(self.meta_data.default_file_path, self.meta_data.input_file)
            self.hwp.Open(file_path)
        except Exception as e:
            raise Exception(f"파일을 열 수 없습니다: {e}")

    def find_and_replace(self, replace_dict):
        """
        문서 내 여러 개의 텍스트 검색 및 바꾸기

        :param replace_dict: 키-값 형식의 치환 데이터 (예: {"%Name": "홍길동"})
        """

        for find_text, replace_text in replace_dict.items():
            self.hwp.HAction.Run("MoveTop")
            self.hwp.HAction.GetDefault("RepeatFind", self.hwp.HParameterSet.HFindReplace.HSet)
            
            self.hwp.HParameterSet.HFindReplace.FindString = find_text
            self.hwp.HParameterSet.HFindReplace.ReplaceString = replace_text
            
            self.hwp.HParameterSet.HFindReplace.ReplaceMode = 1  # 모두 바꾸기
            self.hwp.HParameterSet.HFindReplace.IgnoreMessage = 1  # 메시지 창 숨김
            
            # self.hwp.HAction.Execute("RepeatFind", self.hwp.HParameterSet.HFindReplace.HSet)
            self.hwp.HAction.Execute("AllReplace", self.hwp.HParameterSet.HFindReplace.HSet)
            # 메시지 창 자동 확인 처리 추가

    def find_text(self, search_text):
        """특정 문자열을 문서에서 찾음"""
        self.hwp.HAction.Run("MoveTop")
        self.hwp.HAction.GetDefault("RepeatFind", self.hwp.HParameterSet.HFindReplace.HSet)
        self.hwp.HParameterSet.HFindReplace.FindString = search_text
        self.hwp.HParameterSet.HFindReplace.Direction = 1  # 아래 방향 검색
        return self.hwp.HAction.Execute("RepeatFind", self.hwp.HParameterSet.HFindReplace.HSet)

    def select_cell(self):
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

    def apply_diagonal_to_weekend(self):
        """현재 셀에서 5칸 오른쪽으로 이동하며 대각선 적용"""
        for _ in range(5):
            self.move_cell("right", 1)
            self.diagonal_cell()
        self.move_cell("left", 5)

    def mark_day_off(self, weekends: set):
        """주말을 찾아 해당 셀에 대각선 표시"""

        for day_label in ["%일1", "%일2"]:
            self.find_and_select_cell(day_label)
            for _ in range(16):
                self.move_cell("down", 1)
                self.select_cell()
                self.hwp.HAction.Run("TableCellInput")
                self.hwp.InitScan(0,2)
                text = self.hwp.GetText()
                try:
                    cell_date = int(text[1])
                except ValueError:
                    continue
                if cell_date in weekends:
                    self.apply_diagonal_to_weekend()

                self.hwp.HAction.Run("MoveTop")
    def remove_invalid_days(self):
        """달의 말일을 기준으로 존재하지 않는 날짜를 제거"""
        invalid_days = self.meta_data.get_invalid_days()
        
        for day_label in ["%일2"]:
            self.find_and_select_cell(day_label)
            for _ in range(16):
                self.move_cell("down", 1)
                self.select_cell()
                self.hwp.HAction.Run("TableCellInput")
                self.hwp.InitScan(0,2)
                text = self.hwp.GetText()
                try:
                    cell_date = int(text[1])
                except ValueError:
                    continue
                if cell_date in invalid_days:
                    self.hwp.HAction.Run("TableDeleteCell")
            # self.hwp.HAction.Run("MoveTop")
    def save_file(self):
        """파일 저장"""
        try:
            output_path = os.path.join(self.meta_data.default_file_path, self.meta_data.output_file_name)
            self.hwp.SaveAs(output_path)
        except Exception as e:
            raise Exception(f"파일 저장 실패: {e}")

    def close(self):
        """HWP 종료"""
        self.hwp.Quit()


def modify_hwp_file(meta_data: MetaData, sc:EroomManagerSchedule):
    """HWP 파일을 열고 지정된 단어를 변경한 후 저장"""
    processor = None
    replace_dict = generate_replace_dict(meta_data, sc)
    try:
        processor = HwpProcessor(meta_data)
        processor.open_file()
        processor.mark_day_off(sc.get_day_off(meta_data))
        processor.remove_invalid_days()
        processor.find_and_replace(replace_dict)
        processor.save_file()
        print(f"파일이 성공적으로 저장되었습니다: {meta_data.output_file_name}")
    except Exception as e:
        print(f"오류 발생: {e}")
    finally:
        if processor:
            processor.close()

