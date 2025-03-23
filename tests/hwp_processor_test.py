import pytest
import os
import sys
from tempfile import NamedTemporaryFile
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
from hwp_processor import HwpProcessor, MetaData

@pytest.fixture
def test_hwp_processor():
    """테스트용 HwpProcessor 인스턴스를 생성하는 픽스처"""
    temp_hwp = NamedTemporaryFile(delete=True, suffix=".hwp")  # 임시 HWP 파일 생성
    meta_data = MetaData(
        default_file_path=os.path.dirname(temp_hwp.name),
        input_file=os.path.basename(temp_hwp.name),
        output_file_name="test_output.hwp",
        target_date="2025-02"
    )

    processor = HwpProcessor(meta_data)
    yield processor  # 테스트 실행

    # 테스트 끝난 후 정리 (파일 닫기 및 삭제)
    processor.close()
    try:
        os.remove(temp_hwp.name)
        os.remove(os.path.join(meta_data.default_file_path, meta_data.output_file_name)) 
    except PermissionError:
        print("파일이 사용 중이라 삭제할 수 없습니다. 프로세스를 확인하세요.")

def test_open_and_save_file(test_hwp_processor):
    """HWP 파일 열고 저장하는 기능 테스트"""
    processor = test_hwp_processor
    processor.open_file()
    processor.save_file()

    assert os.path.exists(os.path.join(processor.meta_data.default_file_path, processor.meta_data.output_file_name))



# def test_find_and_replace(test_hwp_processor):
#     """HWP 문서 내 텍스트 치환 기능 테스트"""
#     processor = test_hwp_processor
#     processor.open_file()
    
#     replace_dict = {"%Name": "테스트"}
#     processor.find_and_replace(replace_dict)
    
#     assert processor.find_text("테스트")  # 치환된 텍스트가 존재해야 함

# def test_mark_day_off(test_hwp_processor):
#     """주말(토, 일) 처리 테스트"""
#     processor = test_hwp_processor
#     processor.open_file()
    
#     weekends = {3, 4, 10, 11}  # 예제 주말 데이터
#     processor.mark_day_off(weekends)
    
#     # 특정 주말 날짜가 있는지 확인하는 로직 추가 가능
#     assert True  # 실제 검증하려면 문서 구조를 분석해야 함

# def test_remove_invalid_days(test_hwp_processor):
#     """존재하지 않는 날짜 삭제 테스트"""
#     processor = test_hwp_processor
#     processor.open_file()
    
#     processor.remove_invalid_days()
    
#     # 특정 삭제된 날짜가 존재하지 않는지 확인 (실제 검증 로직 추가 필요)
#     assert True  # HWP 내부 내용을 확인할 방법이 있다면 보강
