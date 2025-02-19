
from datetime import datetime
import os
import calendar
class EroomManagerSchedule:
    def __init__(self, name, substitute_holiday, saturday_workday):
        """
        청년이룸 매니저의 근무 일정을 관리하는 클래스
        
        :param name: 매니저 이름
        :param substitute_holiday: 대체휴무일 (YYYY-MM-DD 형식의 문자열)
        :param saturday_workday: 토요일 근무일 (YYYY-MM-DD 형식의 문자열)
        """
        self.name = name  # 매니저 이름
        self.substitute_holiday = substitute_holiday  # 대체휴무일
        self.saturday_workday = saturday_workday  # 토요일 근무일

    def __repr__(self):
        return f"EroomManagerSchedule(name={self.name}, substitute_holiday={self.substitute_holiday}, saturday_workday={self.saturday_workday})"
    
    def to_dict(self):
        """객체를 딕셔너리 형태로 변환"""
        return {
            "name": self.name,
            "substitute_holiday": self.substitute_holiday,
            "saturday_workday": self.saturday_workday
        }



class PublicHoliday:
    def __init__(self, date):
        """
        공휴일 정보를 저장하는 클래스

        :param date: 공휴일 날짜 (YYYY-MM-DD 형식의 문자열)
        """
        self.date = self._validate_date(date)  # 날짜 검증 후 저장

    def _validate_date(self, date):
        """YYYY-MM-DD 형식의 날짜인지 검증"""
        try:
            return datetime.strptime(date, "%Y-%m-%d").date()
        except ValueError:
            raise ValueError("날짜 형식이 올바르지 않습니다. YYYY-MM-DD 형식이어야 합니다.")

    def __repr__(self):
        return f"PublicHoliday(date={self.date})"

    def to_dict(self):
        """객체를 딕셔너리 형태로 변환"""
        return {"date": self.date.strftime("%Y-%m-%d")}
    
class MetaData:
    def __init__(self, default_file_path, input_file, output_file_name, target_date):
        """
        HWP 자동 생성에 필요한 메타데이터를 관리하는 클래스

        :param default_file_path: 기본 파일 경로 (디렉토리 경로)
        :param input_file: 템플릿으로 사용할 파일 이름
        :param output_file_name: 출력할 파일 이름
        :param target_date: 자동 생성 기준 연월 (YYYY-MM 형식)
        """
        self.default_file_path = self._validate_path(default_file_path)
        self.input_file = input_file
        self.output_file_name = output_file_name
        self.target_date = self._validate_date(target_date)

    def _validate_path(self, path):
        """경로가 유효한지 검사 (존재하지 않으면 예외 발생)"""
        if not os.path.exists(path):
            raise FileNotFoundError(f"경로가 존재하지 않습니다: {path}")
        return path

    def _validate_date(self, date):
        """YYYY-MM 형식의 날짜인지 검증"""
        try:
            return datetime.strptime(date, "%Y-%m").strftime("%Y-%m")
        except ValueError:
            raise ValueError("날짜 형식이 올바르지 않습니다. YYYY-MM 형식이어야 합니다.")

    def __repr__(self):
        return (f"MetaData(default_file_path={self.default_file_path}, input_file={self.input_file}, "
                f"output_file_name={self.output_file_name}, target_date={self.target_date})")

    def to_dict(self):
        """객체를 딕셔너리 형태로 변환"""
        return {
            "default_file_path": self.default_file_path,
            "input_file": self.input_file,
            "output_file_name": self.output_file_name,
            "target_date": self.target_date
        }
    def get_weekends(self):
        """지정된 월의 주말 날짜 목록을 반환"""
        year, month = map(int, self.target_date.split('-'))
        return {day for day in range(1, calendar.monthrange(year, month)[1] + 1)
                if datetime(year, month, day).weekday() in [5, 6]}
    def get_invalid_days(self):
        """해당 월의 존재하지 않는 날짜 목록을 반환"""
        year, month = map(int, self.target_date.split('-'))
        last_day = calendar.monthrange(year, month)[1]
        return {29, 30, 31} - set(range(1, last_day + 1))



def generate_replace_dict(metadata:MetaData, eroom_manager_schedule:EroomManagerSchedule):
    """
    MetaData와 EroomManagerSchedule을 기반으로 한 replace_dict 생성 함수

    :param metadata: MetaData 객체
    :param eroom_manager_schedule: EroomManagerSchedule 객체
    :return: 치환할 데이터를 담은 딕셔너리
    """
    try:
        # target_date에서 연도와 월 추출
        target_date = datetime.strptime(metadata.target_date, "%Y-%m")
        year = target_date.year
        month = target_date.month
        last_day = calendar.monthrange(year, month)[1]  # 해당 월의 마지막 날 계산

        return {
            
            "%Name": eroom_manager_schedule.name,  # 매니저 이름
            "%Year": str(year),  # 연도
            "%Month": str(month),  # 월 (숫자 형태)
            "%Endday": str(last_day),  # 해당 월의 마지막 날
            "%일1" : "월/일",
            "%일2": "월/일"
        }
    except Exception as e:
        raise ValueError(f"replace_dict 생성 중 오류 발생: {e}")
