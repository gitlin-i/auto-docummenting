
from time import sleep
import win32com.client
from eroom import MetaData,generate_replace_dict, EroomManagerSchedule
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
            "up": "UpCell",
            "down": "DownCell",
            "left": "LeftCell",
            "right": "RightCell"
        }
        if direction not in directions:
            raise ValueError("Invalid direction. Use 'up', 'down', 'left', or 'right'.")
        if not isinstance(steps, int) or steps < 1:
            raise ValueError("Steps must be a positive integer.")
        for _ in range(steps):
            self.hwp.HAction.Run("Table"+directions[direction])
        self.hwp.HAction.Run("TableCellBlock")  # 이동 후 현재 셀 다시 지정

    def diagonal_cell(self):
        """현재 지정된 셀에 대각선을 긋는 함수"""
        self.hwp.HAction.Run("TableCellBorderDiagonalUp")

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
    def test_functions(self):
        """테스트 함수: 특정 셀을 찾고 이동 및 대각선 처리 실행"""
        test_text = "%일1"
        if self.find_and_select_cell(test_text):
            print(f"'{test_text}' 셀을 찾고 선택하였습니다.")
            self.move_cell("right",3)

            print("오른쪽으로 3 칸 이동하였습니다.")
            self.hwp.HAction.Run("TableCellBorderDiagonalUp")
            print("현재 칸을 대각선 처리하였습니다.")
        else:
            print(f"'{test_text}' 셀을 찾을 수 없습니다.")



def modify_hwp_file(meta_data:MetaData, replace_dict):
    """HWP 파일을 열고 지정된 단어를 변경한 후 저장"""
    processor = None
    try:
        processor = HwpProcessor(meta_data)
        processor.open_file()
        processor.diagonal_cell()
        processor.find_and_replace(replace_dict)
        processor.save_file()
        print(f"파일이 성공적으로 저장되었습니다: {meta_data.output_file_name}")
    except Exception as e:
        print(f"오류 발생: {e}")
    finally:
        if processor:
            processor.close()


default_file_path = "C:/Users/pc/Desktop/project/"
# 파일 경로 및 검색/교체 문자열 설정
manager_name = "박석진"
input_file = "청년이룸출근부.hwp"  # 현재 경로에 있는 파일
output_file_name = "청년이룸출근부{}.hwp".format("_"+ manager_name)

         
if __name__ == '__main__':
    meta_data = MetaData(default_file_path,input_file,output_file_name,"2025-02")
    sc = EroomManagerSchedule(manager_name,"2025-02-10","2025-02-15")
    replace_dict = generate_replace_dict(meta_data,sc)
    modify_hwp_file(meta_data,replace_dict)
