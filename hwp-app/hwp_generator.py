import win32com.client
import os
import calendar
from datetime import datetime
import json

class HwpGenerator:
    def __init__(self, template_path, output_dir):
        """
        HWP 파일 생성기
        
        :param template_path: 템플릿 HWP 파일 경로
        :param output_dir: 출력 디렉토리 경로
        """
        self.template_path = template_path
        self.output_dir = output_dir
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
            print(f"오류 타입: {type(e)}")
            return False
    
    def open_template(self):
        """템플릿 파일 열기"""
        try:
            print(f"템플릿 파일 열기 시도: {self.template_path}")
            print(f"파일 존재 여부: {os.path.exists(self.template_path)}")
            print(f"절대 경로: {os.path.abspath(self.template_path)}")
            
            # 현재 작업 디렉토리 확인
            print(f"현재 작업 디렉토리: {os.getcwd()}")
            
            # 템플릿 파일 열기 - 다른 방법 시도
            try:
                # 방법 1: 기본 Open
                self.hwp.Open(self.template_path)
                print("템플릿 파일 열기 성공 (방법1)")
            except Exception as open_error:
                print(f"방법 1 실패: {open_error}")
                try:
                    # 방법 2: 절대 경로로 열기
                    abs_path = os.path.abspath(self.template_path)
                    self.hwp.Open(abs_path)
                    print("템플릿 파일 열기 성공 (방법2)")
                except Exception as open_error2:
                    print(f"방법 2 실패: {open_error2}")
                    # 방법 3: 새 문서로 열기
                    self.hwp.Open(self.template_path, "HWPX")
                    print("템플릿 파일 열기 성공 (방법3)")
            
            # 문서 내용 확인
            try:
                # 여러 방법으로 문서 내용 확인
                content = None
                
                # 방법 1: GetTextFile
                try:
                    content = self.hwp.GetTextFile("", "")
                    print("문서 내용 가져오기 성공 (GetTextFile)")
                except:
                    pass
                
                # 방법 2: HGetText
                if not content:
                    try:
                        content = self.hwp.HGetText("", "")
                        print("문서 내용 가져오기 성공 (HGetText)")
                    except:
                        pass
                
                # 방법 3: GetText
                if not content:
                    try:
                        content = self.hwp.GetText()
                        print("문서 내용 가져오기 성공 (GetText)")
                    except:
                        pass
                
                print(f"문서 내용 길이: {len(content) if content else 0}")
                if content and len(str(content)) > 10:
                    print(f"문서 내용 일부: {str(content)[:100]}...")
                else:
                    print("문서 내용을 가져올 수 없습니다.")
                    
            except Exception as content_error:
                print(f"문서 내용 확인 실패: {content_error}")
            
            return True
        except Exception as e:
            print(f"템플릿 파일 열기 실패: {e}")
            print(f"오류 타입: {type(e)}")
            return False
    
    def find_and_replace(self, replace_dict):
        """문서 내 텍스트 검색 및 바꾸기"""
        print(f"치환할 항목들: {replace_dict}")
        
        # 먼저 문서 내용을 확인
        try:
            doc_text = self.hwp.GetTextFile("", "")
            print(f"원본 문서 내용: {doc_text}")
        except Exception as text_error:
            print(f"문서 내용 가져오기 실패: {text_error}")
            return
        
        for find_text, replace_text in replace_dict.items():
            try:
                print(f"치환 시도: {find_text} -> {replace_text}")
                
                # 문서 맨 위로 이동
                self.hwp.HAction.Run("MoveTop")
                
                # 찾기/바꾸기 설정 - 다른 방법 시도
                try:
                    # 방법 1: 기본 Find/Replace
                    self.hwp.HAction.GetDefault("RepeatFind", self.hwp.HParameterSet.HFindReplace.HSet)
                    
                    self.hwp.HParameterSet.HFindReplace.FindString = find_text
                    self.hwp.HParameterSet.HFindReplace.ReplaceString = replace_text
                    self.hwp.HParameterSet.HFindReplace.ReplaceMode = 1  # 모두 바꾸기
                    self.hwp.HParameterSet.HFindReplace.IgnoreMessage = 1  # 메시지 창 숨김
                    
                    # 치환 실행
                    result = self.hwp.HAction.Execute("AllReplace", self.hwp.HParameterSet.HFindReplace.HSet)
                    print(f"치환 결과 (방법1): {result}")
                    
                    if not result:
                        # 방법 2: 단순 Find/Replace
                        print("방법 1 실패, 방법 2 시도...")
                        self.hwp.HAction.GetDefault("Find", self.hwp.HParameterSet.HFindReplace.HSet)
                        self.hwp.HParameterSet.HFindReplace.FindString = find_text
                        self.hwp.HParameterSet.HFindReplace.ReplaceString = replace_text
                        result = self.hwp.HAction.Execute("Find", self.hwp.HParameterSet.HFindReplace.HSet)
                        print(f"치환 결과 (방법2): {result}")
                        
                except Exception as method_error:
                    print(f"치환 메서드 오류: {method_error}")
                    # 방법 3: 직접 텍스트 치환
                    print("방법 3: 직접 텍스트 치환 시도...")
                    try:
                        # 현재 문서의 텍스트를 가져와서 치환
                        doc_text = self.hwp.GetTextFile("", "")
                        if doc_text:
                            new_text = doc_text.replace(find_text, replace_text)
                            self.hwp.SetTextFile("", "", new_text)
                            print("직접 텍스트 치환 성공")
                    except Exception as direct_error:
                        print(f"직접 텍스트 치환 실패: {direct_error}")
                
            except Exception as e:
                print(f"텍스트 치환 실패 ({find_text} -> {replace_text}): {e}")
                print(f"오류 타입: {type(e)}")
    
    def save_file(self, filename):
        """파일 저장"""
        try:
            # 절대 경로로 변환
            output_path = os.path.abspath(os.path.join(self.output_dir, filename))
            print(f"저장 경로: {output_path}")
            print(f"출력 디렉토리 존재 여부: {os.path.exists(self.output_dir)}")
            
            # HWP 파일 형식으로 저장 (확장자 명시)
            if not filename.endswith('.hwp'):
                output_path = output_path + '.hwp'
                print(f"수정된 저장 경로: {output_path}")
            
            # 현재 열린 문서 저장
            self.hwp.Save()
            print(f"파일 저장 완료: {filename}")
            print(f"저장된 파일 존재 여부: {os.path.exists(output_path)}")
            return True
        except Exception as e:
            print(f"파일 저장 실패: {e}")
            print(f"오류 타입: {type(e)}")
            return False
    
    def close_hwp(self):
        """한글 오피스 종료"""
        if self.hwp:
            self.hwp.Quit()
    
    def generate_replace_dict(self, manager_name, target_date):
        """
        치환할 데이터 딕셔너리 생성
        
        :param manager_name: 매니저 이름
        :param target_date: 대상 날짜 (YYYY-MM 형식)
        :return: 치환 딕셔너리
        """
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
    
    def generate_hwp_for_manager(self, manager_name, target_date):
        """
        특정 매니저의 HWP 파일 생성
        
        :param manager_name: 매니저 이름
        :param target_date: 대상 날짜 (YYYY-MM 형식)
        :return: 성공 여부
        """
        try:
            # 파일명 생성 (청년이룸출근부_매니저이름_YYMM.hwp)
            target_date_obj = datetime.strptime(target_date, "%Y-%m")
            year_suffix = str(target_date_obj.year)[-2:]  # YY
            month_str = str(target_date_obj.month).zfill(2)  # MM
            filename = f"청년이룸출근부_{manager_name}_{year_suffix}{month_str}.hwp"
            
            # 템플릿 파일을 직접 복사
            import shutil
            template_path = os.path.abspath(self.template_path)
            output_path = os.path.abspath(os.path.join(self.output_dir, filename))
            
            print(f"템플릿 파일 복사: {template_path} -> {output_path}")
            shutil.copy2(template_path, output_path)
            print(f"파일 복사 완료: {filename}")
            
            # 복사된 파일을 HWP로 열기
            try:
                self.hwp.Open(output_path)
                print(f"복사된 파일 열기 성공: {filename}")
                
                # 치환 데이터 생성
                replace_dict = self.generate_replace_dict(manager_name, target_date)
                
                # 텍스트 치환
                self.find_and_replace(replace_dict)
                
                # 파일 저장
                success = self.save_file(filename)
                
                return success
                
            except Exception as open_error:
                print(f"복사된 파일 열기 실패: {open_error}")
                return False
            
        except Exception as e:
            print(f"매니저 {manager_name}의 HWP 파일 생성 실패: {e}")
            return False
    
    def generate_hwp_for_all_managers(self, managers, target_date):
        """
        모든 매니저의 HWP 파일 생성
        
        :param managers: 매니저 리스트
        :param target_date: 대상 날짜 (YYYY-MM 형식)
        :return: 성공한 파일 수
        """
        if not self.initialize_hwp():
            return 0
        
        success_count = 0
        
        try:
            for manager in managers:
                print(f"매니저 {manager}의 HWP 파일 생성 중...")
                if self.generate_hwp_for_manager(manager, target_date):
                    success_count += 1
                else:
                    print(f"매니저 {manager}의 파일 생성 실패")
        finally:
            self.close_hwp()
        
        return success_count

def main():
    """메인 실행 함수"""
    # 설정
    template_path = "hwp-app/청년이룸출근부.hwp"  # 템플릿 파일 경로
    output_dir = "hwp-app"  # 출력 디렉토리 (hwp-app 폴더)
    
    # 매니저 리스트 (예시)
    managers = [
        "김단아",
        "문주원", 
        "박석진",
        "서은영"
    ]
    
    # 대상 날짜 (예시)
    target_date = "2025-03"  # YYYY-MM 형식
    
    # HWP 생성기 초기화
    generator = HwpGenerator(template_path, output_dir)
    
    # 모든 매니저의 HWP 파일 생성
    print(f"총 {len(managers)}명의 매니저에 대해 HWP 파일을 생성합니다...")
    success_count = generator.generate_hwp_for_all_managers(managers, target_date)
    
    print(f"생성 완료: {success_count}/{len(managers)}개 파일")

if __name__ == "__main__":
    main() 