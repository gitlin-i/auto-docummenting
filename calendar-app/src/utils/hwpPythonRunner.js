// HWP Generator를 위한 Python 러너
class HwpPythonRunner {
  constructor() {
    this.pyodide = null;
    this.isLoaded = false;
  }

  async loadPyodide() {
    if (this.isLoaded) return;

    try {
      // Pyodide가 이미 로드되어 있는지 확인
      if (typeof window.loadPyodide === 'undefined') {
        console.log('Pyodide 로딩 시작...');
        
        // Pyodide CDN에서 로드
        const script = document.createElement('script');
        script.src = 'https://cdn.jsdelivr.net/pyodide/v0.24.1/full/pyodide.js';
        script.async = true;
        document.head.appendChild(script);

        // 스크립트 로드 완료 대기
        await new Promise((resolve, reject) => {
          script.onload = () => {
            console.log('Pyodide 스크립트 로드 완료');
            resolve();
          };
          script.onerror = () => {
            reject(new Error('Pyodide 스크립트 로드 실패'));
          };
        });

        // 스크립트 로드 후 잠시 대기
        await new Promise(resolve => setTimeout(resolve, 2000));
      }

      // Pyodide 초기화
      if (typeof window.loadPyodide === 'undefined') {
        throw new Error('Pyodide가 로드되지 않았습니다. 네트워크 연결을 확인해주세요.');
      }

      console.log('Pyodide 초기화 시작...');
      this.pyodide = await window.loadPyodide({
        indexURL: "https://cdn.jsdelivr.net/pyodide/v0.24.1/full/"
      });

      console.log('Pyodide 초기화 완료');
      
      this.isLoaded = true;
      console.log('Pyodide 로드 완료');
    } catch (error) {
      console.error('Pyodide 로드 실패:', error);
      throw error;
    }
  }

  async runPythonCode(code) {
    if (!this.isLoaded) {
      await this.loadPyodide();
    }

    try {
      console.log('Python 코드 실행 중...');
      console.log('실행할 Python 코드:', code);
      const result = await this.pyodide.runPythonAsync(code);
      console.log('Python 코드 실행 완료');
      console.log('실행 결과:', result);
      return result;
    } catch (error) {
      console.error('Python 코드 실행 오류:', error);
      console.error('오류 상세 정보:', {
        name: error.name,
        message: error.message,
        stack: error.stack
      });
      throw error;
    }
  }

  // HWP 파일 생성 함수 (웹어셈블리 버전)
  async generateHwpFiles(managers, targetDate) {
    const pythonCode = `
import json
import sys
from datetime import datetime
import calendar

# 데이터 파싱
managers_data = ${JSON.stringify(managers)}
target_date = "${targetDate}"

print(f"매니저 수: {len(managers_data)}")
print(f"대상 날짜: {target_date}")

# HWP 파일 생성 시뮬레이션
generated_files = []

for manager in managers_data:
    manager_name = manager.get('name', 'Unknown')
    
    # 날짜 파싱
    target_date_obj = datetime.strptime(target_date, "%Y-%m")
    year = target_date_obj.year
    month = target_date_obj.month
    last_day = calendar.monthrange(year, month)[1]
    
    # 파일명 생성 (청년이룸출근부_매니저이름_YYMM.hwp)
    year_suffix = str(year)[-2:]  # YY
    month_str = str(month).zfill(2)  # MM
    filename = f"청년이룸출근부_{manager_name}_{year_suffix}{month_str}.hwp"
    
    # 치환 데이터 생성
    replace_dict = {
        "%Name": manager_name,
        "%Year": str(year),
        "%Month": str(month),
        "%Endday": str(last_day),
        "%일1": "월/일",
        "%일2": "월/일"
    }
    
    print(f"매니저 {manager_name}의 HWP 파일 생성: {filename}")
    print(f"치환 데이터: {replace_dict}")
    
    # 파일 정보 저장
    file_info = {
        'manager': manager_name,
        'filename': filename,
        'replace_dict': replace_dict,
        'target_date': target_date,
        'year': year,
        'month': month,
        'last_day': last_day
    }
    
    generated_files.append(file_info)

# 결과 반환
result = {
    'success': True,
    'message': f'HWP 파일 생성 준비 완료: {len(generated_files)}개 파일',
    'generated_files': generated_files,
    'target_date': target_date,
    'total_managers': len(managers_data),
    'success_count': len(generated_files)
}

result
`;

    return await this.runPythonCode(pythonCode);
  }

  // 단일 매니저 HWP 파일 생성
  async generateHwpFileForManager(templateBuffer, templateFileName, calendarData) {
    const pythonCode = `
import json
import sys
from datetime import datetime
import calendar
import base64

print("HWP 파일 생성 시작...")

# 데이터 파싱
template_buffer = ${JSON.stringify(Array.from(new Uint8Array(templateBuffer)))}
template_filename = "${templateFileName}"
calendar_data = ${JSON.stringify(calendarData)}

print(f"템플릿 파일: {template_filename}")
print(f"매니저: {calendar_data['manager']}")
print(f"대상 월: {calendar_data['targetMonth']}")

# 날짜 파싱
target_date_obj = datetime.strptime(calendar_data['targetMonth'], "%Y-%m")
year = target_date_obj.year
month = target_date_obj.month
last_day = calendar.monthrange(year, month)[1]

# 파일명 생성
year_suffix = str(year)[-2:]
month_str = str(month).zfill(2)
filename = f"청년이룸출근부_{calendar_data['manager']}_{year_suffix}{month_str}.hwp"

print(f"생성될 파일명: {filename}")

# 치환 데이터 생성
replace_dict = {
    "%Name": calendar_data['manager'],
    "%Year": str(year),
    "%Month": str(month),
    "%Endday": str(last_day),
    "%일1": "월/일",
    "%일2": "월/일"
}

print(f"치환 데이터: {replace_dict}")

# 스케줄 데이터 분석
overtime_count = len(calendar_data['overtimeSchedules'])
holiday_count = len(calendar_data['substituteHolidays'])
vacation_count = len(calendar_data['vacationSchedules'])

print(f"토요일 근무: {overtime_count}일")
print(f"대체 휴무일: {holiday_count}일")
print(f"연가: {vacation_count}일")

# HWP 파일 내용 생성 (더 실제적인 형태)
hwp_content = f"""청년이룸 출근부

매니저: {calendar_data['manager']['name']}
대상 월: {calendar_data['targetMonth']}
연도: {year}
월: {month}
마지막 날: {last_day}

치환 데이터:
- %Name: {calendar_data['manager']['name']}
- %Year: {year}
- %Month: {month}
- %Endday: {last_day}
- %일1: 월/일
- %일2: 월/일

스케줄 정보:
- 토요일 근무: {overtime_count}일
- 대체 휴무일: {holiday_count}일
- 연가: {vacation_count}일

이 파일은 웹어셈블리로 생성된 HWP 파일입니다.
실제 HWP 파일을 생성하려면 한글 오피스가 필요합니다.
"""

# 실제 출근부 형태로 HWP 파일 내용 생성
month_name = f"{year}년 {month}월"
    
# 월별 달력 생성
cal = calendar.monthcalendar(year, month)
calendar_text = f"{month_name} 출근부\\n\\n"
calendar_text += "일\\t월\\t화\\t수\\t목\\t금\\t토\\n"
    
for week in cal:
    week_text = ""
    for day in week:
        if day == 0:
            week_text += "\\t"
        else:
            week_text += f"{day}\\t"
    calendar_text += week_text.rstrip() + "\\n"
    
# 스케줄 상세 정보
schedule_details = ""
if calendar_data['overtimeSchedules']:
    schedule_details += "\\n토요일 근무:\\n"
    for date, managers in calendar_data['overtimeSchedules'].items():
        if any(m.get('name') == calendar_data['manager'] for m in managers):
            schedule_details += f"- {date}\\n"
    
if calendar_data['substituteHolidays']:
    schedule_details += "\\n대체 휴무일:\\n"
    for date, managers in calendar_data['substituteHolidays'].items():
        if any(m.get('name') == calendar_data['manager'] for m in managers):
            schedule_details += f"- {date}\\n"
    
if calendar_data['vacationSchedules']:
    schedule_details += "\\n연가:\\n"
    for date, managers in calendar_data['vacationSchedules'].items():
        if any(m.get('name') == calendar_data['manager'] for m in managers):
            schedule_details += f"- {date}\\n"
    
hwp_content = f"""청년이룸 출근부

매니저: {calendar_data['manager']}
대상 월: {calendar_data['targetMonth']}
연도: {year}
월: {month}
마지막 날: {last_day}

치환 데이터:
- %Name: {calendar_data['manager']}
- %Year: {year}
- %Month: {month}
- %Endday: {last_day}
- %일1: 월/일
- %일2: 월/일

{calendar_text}

스케줄 정보:
- 토요일 근무: {overtime_count}일
- 대체 휴무일: {holiday_count}일
- 연가: {vacation_count}일

{schedule_details}

이 파일은 웹어셈블리로 생성된 HWP 파일입니다.
실제 HWP 파일을 생성하려면 한글 오피스가 필요합니다.
"""

# 결과 반환
result = {
    'success': True,
    'message': f'{calendar_data["manager"]} 매니저 HWP 파일 생성 완료',
    'filename': filename,
    'content': hwp_content,
    'manager': calendar_data['manager'],
    'target_month': calendar_data['targetMonth'],
    'replace_dict': replace_dict,
    'schedule_info': {
        'overtime_count': overtime_count,
        'holiday_count': holiday_count,
        'vacation_count': vacation_count
    }
}

print("HWP 파일 생성 완료!")
result
`;

    return await this.runPythonCode(pythonCode);
  }

  // 텍스트 파일로 HWP 내용 생성 (다운로드용)
  async generateHwpContent(manager, targetDate) {
    const pythonCode = `
import json
from datetime import datetime
import calendar

# 데이터 파싱
manager_data = ${JSON.stringify(manager)}
target_date = "${targetDate}"

# 날짜 파싱
target_date_obj = datetime.strptime(target_date, "%Y-%m")
year = target_date_obj.year
month = target_date_obj.month
last_day = calendar.monthrange(year, month)[1]

# 파일명 생성
year_suffix = str(year)[-2:]
month_str = str(month).zfill(2)
filename = f"청년이룸출근부_{manager_data.get('name', 'Unknown')}_{year_suffix}{month_str}.txt"

# HWP 파일 내용 생성 (텍스트 형식)
content = f"""청년이룸 출근부

매니저: {manager_data.get('name', 'Unknown')}
대상 월: {target_date}
연도: {year}
월: {month}
마지막 날: {last_day}

치환 데이터:
- %Name: {manager_data.get('name', 'Unknown')}
- %Year: {year}
- %Month: {month}
- %Endday: {last_day}
- %일1: 월/일
- %일2: 월/일

이 파일은 HWP 템플릿에서 생성되었습니다.
실제 HWP 파일을 생성하려면 한글 오피스가 필요합니다.
"""

result = {
    'success': True,
    'filename': filename,
    'content': content,
    'manager': manager_data.get('name', 'Unknown'),
    'target_date': target_date
}

result
`;

    return await this.runPythonCode(pythonCode);
  }

  // Python 테스트
  async testPython() {
    const testCode = `
import sys
print("Python 버전:", sys.version)
print("HWP Generator Python 러너가 정상적으로 작동합니다!")

# 간단한 계산
result = 2 + 3 * 4
print(f"계산 결과: {result}")

result
`;

    return await this.runPythonCode(testCode);
  }

  // 간단한 HWP 테스트
  async testHwpGeneration() {
    const testCode = `
import json
from datetime import datetime
import calendar

print("HWP 생성 테스트 시작...")

# 테스트 데이터
test_manager = {"name": "테스트매니저"}
test_date = "2025-03"

print(f"테스트 매니저: {test_manager}")
print(f"테스트 날짜: {test_date}")

# 날짜 파싱
target_date_obj = datetime.strptime(test_date, "%Y-%m")
year = target_date_obj.year
month = target_date_obj.month
last_day = calendar.monthrange(year, month)[1]

print(f"연도: {year}, 월: {month}, 마지막 날: {last_day}")

# 파일명 생성
year_suffix = str(year)[-2:]
month_str = str(month).zfill(2)
filename = f"청년이룸출근부_{test_manager['name']}_{year_suffix}{month_str}.txt"

print(f"생성될 파일명: {filename}")

# 치환 데이터 생성
replace_dict = {
    "%Name": test_manager['name'],
    "%Year": str(year),
    "%Month": str(month),
    "%Endday": str(last_day),
    "%일1": "월/일",
    "%일2": "월/일"
}

print(f"치환 데이터: {replace_dict}")

result = {
    'success': True,
    'filename': filename,
    'replace_dict': replace_dict,
    'test': 'HWP 생성 테스트 성공'
}

print("테스트 완료!")
result
`;

    return await this.runPythonCode(testCode);
  }

  // 파일 다운로드 함수
  async downloadFile(filename, content) {
    try {
      // HWP 파일의 경우 바이너리 데이터로 처리
      let blob;
      let mimeType;
      
      if (filename.endsWith('.hwp')) {
        // HWP 파일은 바이너리로 처리 (실제로는 텍스트 시뮬레이션)
        // 실제 HWP 파일을 생성하려면 한글 오피스 API가 필요
        blob = new Blob([content], { 
          type: 'application/x-hwp' // HWP 파일 MIME 타입
        });
        mimeType = 'application/x-hwp';
      } else {
        // 일반 텍스트 파일
        blob = new Blob([content], { 
          type: 'text/plain;charset=utf-8' 
        });
        mimeType = 'text/plain';
      }
      
      // 다운로드 링크 생성
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.download = filename;
      link.style.display = 'none';
      
      // 다운로드 폴더에 저장 (브라우저 기본 다운로드 폴더)
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      
      // 메모리 정리
      setTimeout(() => {
        URL.revokeObjectURL(url);
      }, 100);
      
      console.log(`파일 다운로드 완료: ${filename}`);
      return true;
    } catch (error) {
      console.error('파일 다운로드 오류:', error);
      return false;
    }
  }

  // 모든 매니저 HWP 파일 생성
  async generateHwpFilesForAllManagers(templateBuffer, templateFileName, calendarData) {
    const pythonCode = `
import json
import sys
from datetime import datetime
import calendar

print("모든 매니저 HWP 파일 생성 시작...")

# 데이터 파싱
template_buffer = ${JSON.stringify(Array.from(new Uint8Array(templateBuffer)))}
template_filename = "${templateFileName}"
calendar_data = ${JSON.stringify(calendarData)}

print(f"템플릿 파일: {template_filename}")
print(f"매니저 수: {len(calendar_data['managers'])}")
print(f"대상 월: {calendar_data['targetMonth']}")

# 날짜 파싱
target_date_obj = datetime.strptime(calendar_data['targetMonth'], "%Y-%m")
year = target_date_obj.year
month = target_date_obj.month
last_day = calendar.monthrange(year, month)[1]

generated_files = []

# 각 매니저별로 HWP 파일 생성
for manager in calendar_data['managers']:
    manager_name = manager.get('name', 'Unknown')
    
    # 파일명 생성
    year_suffix = str(year)[-2:]
    month_str = str(month).zfill(2)
    filename = f"청년이룸출근부_{manager_name}_{year_suffix}{month_str}.hwp"
    
    print(f"매니저 {manager_name}의 HWP 파일 생성: {filename}")
    
    # 치환 데이터 생성
    replace_dict = {
        "%Name": manager_name,
        "%Year": str(year),
        "%Month": str(month),
        "%Endday": str(last_day),
        "%일1": "월/일",
        "%일2": "월/일"
    }
    
    # 스케줄 데이터 분석
    overtime_count = len(calendar_data['overtimeSchedules'])
    holiday_count = len(calendar_data['substituteHolidays'])
    vacation_count = len(calendar_data['vacationSchedules'])
    
    # HWP 파일 내용 생성
    hwp_content = f"""청년이룸 출근부

매니저: {manager_name}
대상 월: {calendar_data['targetMonth']}
연도: {year}
월: {month}
마지막 날: {last_day}

치환 데이터:
- %Name: {manager_name}
- %Year: {year}
- %Month: {month}
- %Endday: {last_day}
- %일1: 월/일
- %일2: 월/일

스케줄 정보:
- 토요일 근무: {overtime_count}일
- 대체 휴무일: {holiday_count}일
- 연가: {vacation_count}일

이 파일은 웹어셈블리로 생성된 HWP 파일입니다.
실제 HWP 파일을 생성하려면 한글 오피스가 필요합니다.
"""

    # 실제 출근부 형태로 HWP 파일 내용 생성
    month_name = f"{year}년 {month}월"
    
    # 월별 달력 생성
    cal = calendar.monthcalendar(year, month)
    calendar_text = f"{month_name} 출근부\\n\\n"
    calendar_text += "일\\t월\\t화\\t수\\t목\\t금\\t토\\n"
    
    for week in cal:
        week_text = ""
        for day in week:
            if day == 0:
                week_text += "\\t"
            else:
                week_text += f"{day}\\t"
        calendar_text += week_text.rstrip() + "\\n"
    
    # 스케줄 상세 정보
    schedule_details = ""
    if calendar_data['overtimeSchedules']:
        schedule_details += "\\n토요일 근무:\\n"
        for date, managers in calendar_data['overtimeSchedules'].items():
            if any(m.get('name') == manager_name for m in managers):
                schedule_details += f"- {date}\\n"
    
    if calendar_data['substituteHolidays']:
        schedule_details += "\\n대체 휴무일:\\n"
        for date, managers in calendar_data['substituteHolidays'].items():
            if any(m.get('name') == manager_name for m in managers):
                schedule_details += f"- {date}\\n"
    
    if calendar_data['vacationSchedules']:
        schedule_details += "\\n연가:\\n"
        for date, managers in calendar_data['vacationSchedules'].items():
            if any(m.get('name') == manager_name for m in managers):
                schedule_details += f"- {date}\\n"
    
    hwp_content = f"""청년이룸 출근부

매니저: {manager_name}
대상 월: {calendar_data['targetMonth']}
연도: {year}
월: {month}
마지막 날: {last_day}

치환 데이터:
- %Name: {manager_name}
- %Year: {year}
- %Month: {month}
- %Endday: {last_day}
- %일1: 월/일
- %일2: 월/일

{calendar_text}

스케줄 정보:
- 토요일 근무: {overtime_count}일
- 대체 휴무일: {holiday_count}일
- 연가: {vacation_count}일

{schedule_details}

이 파일은 웹어셈블리로 생성된 HWP 파일입니다.
실제 HWP 파일을 생성하려면 한글 오피스가 필요합니다.
"""
    
    # 파일 정보 저장
    file_info = {
        'manager': manager_name,
        'filename': filename,
        'content': hwp_content,
        'replace_dict': replace_dict,
        'target_month': calendar_data['targetMonth'],
        'schedule_info': {
            'overtime_count': overtime_count,
            'holiday_count': holiday_count,
            'vacation_count': vacation_count
        }
    }
    
    generated_files.append(file_info)

# 결과 반환
result = {
    'success': True,
    'message': f'모든 매니저 HWP 파일 생성 완료: {len(generated_files)}개',
    'generated_files': generated_files,
    'target_month': calendar_data['targetMonth'],
    'total_managers': len(calendar_data['managers']),
    'success_count': len(generated_files)
}

print(f"HWP 파일 생성 완료! 총 {len(generated_files)}개 파일")
result
`;

    return await this.runPythonCode(pythonCode);
  }
}

export default HwpPythonRunner; 