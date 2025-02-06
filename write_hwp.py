
import win32com.client

class HwpProcessor:
    def __init__(self, default_file_path):
        self.default_file_path = default_file_path
        self.hwp = self._initialize_hwp()

    def _initialize_hwp(self):
        """한글 오피스 객체를 초기화하고 보안 모듈을 등록"""
        try:
            hwp = win32com.client.gencache.EnsureDispatch("HWPFrame.HwpObject")
            hwp.RegisterModule("FilePathCheckDLL", "SecurityModule")  # 보안 경고 방지
            return hwp
        except Exception as e:
            raise Exception(f"HWP 초기화 실패: {e}")

    def open_file(self, file_name):
        """HWP 파일 열기"""
        try:
            self.hwp.Open(self.default_file_path + file_name)
        except Exception as e:
            raise Exception(f"파일을 열 수 없습니다: {e}")

    def find_and_replace(self, find_text, replace_text):
        """문서 내 텍스트 검색 및 바꾸기"""
        self.hwp.HAction.Run("MoveTop")  # 문서 맨 앞으로 이동
        self.hwp.HAction.GetDefault("RepeatFind", self.hwp.HParameterSet.HFindReplace.HSet)
        
        self.hwp.HParameterSet.HFindReplace.FindString = find_text
        self.hwp.HParameterSet.HFindReplace.ReplaceString = replace_text
        self.hwp.HParameterSet.HFindReplace.ReplaceMode = 1  # 모두 바꾸기
        self.hwp.HParameterSet.HFindReplace.IgnoreMessage = 1  # 메시지 창 숨김
        self.hwp.HAction.Execute("RepeatFind", self.hwp.HParameterSet.HFindReplace.HSet)
        self.hwp.HAction.Execute("AllReplace", self.hwp.HParameterSet.HFindReplace.HSet)

    def save_file(self, output_file_name):
        """파일 저장"""
        try:
            self.hwp.SaveAs(self.default_file_path + output_file_name)
        except Exception as e:
            raise Exception(f"파일 저장 실패: {e}")

    def close(self):
        """HWP 종료"""
        self.hwp.Quit()


def modify_hwp_file(default_file_path, input_file_name, output_file_name, find_text, replace_text):
    try:
        processor = HwpProcessor(default_file_path)
        processor.open_file(input_file_name)
        processor.find_and_replace(find_text, replace_text)
        processor.save_file(output_file_name)
        processor.close()
        
        print(f"파일이 성공적으로 저장되었습니다: {output_file_name}")
    
    except Exception as e:
        print(f"오류 발생: {e}")



default_file_path = "C:/Users/pc/Desktop/project/"
# 파일 경로 및 검색/교체 문자열 설정
manager_name = "박석진"
input_file = "청년이룸출근부.hwp"  # 현재 경로에 있는 파일
output_file_name = "청년이룸출근부{}.hwp".format("_"+ manager_name)
find_text = "%Name"  # 예: "안녕하세요"
replace_text = "박석진"  # 예: "반갑습니다"

modify_hwp_file(default_file_path, input_file, output_file_name, find_text, replace_text)
