import pytest
from datetime import datetime
import os
import sys

sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
from eroom import MetaData, EroomManagerSchedule, generate_replace_dict

# 기본 테스트 설정
def setup_test_directory():
    test_path = "./tests"
    os.makedirs(test_path, exist_ok=True)
    return test_path

# MetaData 테스트
@pytest.fixture
def valid_metadata():
    path = setup_test_directory()
    return MetaData(path, "template.hwp", "output.hwp", "2025-03")

@pytest.fixture
def invalid_metadata_path():
    return "./invalid_path"

@pytest.fixture
def invalid_metadata_date():
    return "2025-13"  # 잘못된 월 값

# 유효한 MetaData 객체 생성 테스트
def test_metadata_initialization(valid_metadata):
    assert valid_metadata.default_file_path == "./tests"
    assert valid_metadata.input_file == "template.hwp"
    assert valid_metadata.output_file_name == "output.hwp"
    assert valid_metadata.target_date == "2025-03"

# 경로 검증 실패 테스트
def test_metadata_invalid_path(invalid_metadata_path):
    with pytest.raises(FileNotFoundError):
        MetaData(invalid_metadata_path, "template.hwp", "output.hwp", "2025-03")

# 날짜 검증 실패 테스트
def test_metadata_invalid_date():
    with pytest.raises(ValueError):
        MetaData("./tests", "template.hwp", "output.hwp", "2025-13")

# 주말 날짜 계산 테스트
def test_get_weekends(valid_metadata):
    weekends = valid_metadata.get_weekends()
    assert isinstance(weekends, set)
    assert all(isinstance(day, int) for day in weekends)

# EroomManagerSchedule 테스트
@pytest.fixture
def valid_schedule():
    return EroomManagerSchedule("홍길동", "2025-03-10", "2025-03-15")

def test_get_day_off(valid_metadata, valid_schedule):
    days_off = valid_schedule.get_day_off(valid_metadata)
    expected_days_off = valid_metadata.get_weekends()
    expected_days_off.discard(15)  # 주말 제외
    expected_days_off.add(10)  # 대체 공휴일 포함
    assert days_off == expected_days_off
    assert isinstance(days_off, set)



# generate_replace_dict 테스트
def test_generate_replace_dict(valid_metadata, valid_schedule):
    replace_dict = generate_replace_dict(valid_metadata, valid_schedule)
    assert replace_dict["%Name"] == "홍길동"
    assert replace_dict["%Year"] == "2025"
    assert replace_dict["%Month"] == "3"
    assert "%일1" in replace_dict
