// Pyodide를 사용한 Python 실행기
class PythonRunner {
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

      console.log('Pyodide 초기화 완료, 패키지 로딩 중...');
      
      // 필요한 Python 패키지 설치
      await this.pyodide.loadPackage(['numpy', 'pandas']);
      
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
      const result = await this.pyodide.runPythonAsync(code);
      console.log('Python 코드 실행 완료');
      return result;
    } catch (error) {
      console.error('Python 코드 실행 오류:', error);
      throw error;
    }
  }

  // HWP 파일 생성 함수
  async generateHwpFile(data) {
    const pythonCode = `
import json
import sys
from datetime import datetime

# 데이터 파싱
data = ${JSON.stringify(data)}

managers = data.get('managers', [])
overtime_schedules = data.get('overtimeSchedules', {})
substitute_holidays = data.get('substituteHolidays', {})
vacation_schedules = data.get('vacationSchedules', {})
target_month = data.get('targetMonth', '2025-03')

print(f"매니저 수: {len(managers)}")
print(f"토요일 근무 일정: {len(overtime_schedules)}개")
print(f"대체 휴무일: {len(substitute_holidays)}개")
print(f"연가 일정: {len(vacation_schedules)}개")
print(f"대상 월: {target_month}")

# 여기서 실제 HWP 파일 생성 로직을 구현할 수 있습니다
# 현재는 데이터 검증만 수행

result = {
    'success': True,
    'message': f'HWP 파일 생성 준비 완료: {target_month}',
    'data': {
        'managers': managers,
        'overtime_schedules': overtime_schedules,
        'substitute_holidays': substitute_holidays,
        'vacation_schedules': vacation_schedules,
        'target_month': target_month
    }
}

result
`;

    return await this.runPythonCode(pythonCode);
  }

  // 간단한 Python 코드 테스트
  async testPython() {
    const testCode = `
import sys
print("Python 버전:", sys.version)
print("웹어셈블리에서 Python이 실행되고 있습니다!")

# 간단한 계산
result = 2 + 3 * 4
print(f"계산 결과: {result}")

result
`;

    return await this.runPythonCode(testCode);
  }
}

export default PythonRunner; 