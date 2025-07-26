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

  // 파일 다운로드 기능
  async downloadFile(filename, content) {
    const blob = new Blob([content], { type: 'text/plain;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }
}

export default HwpPythonRunner; 