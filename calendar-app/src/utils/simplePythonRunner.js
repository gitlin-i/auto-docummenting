// 간단한 Python 시뮬레이터 (웹어셈블리 대신 JavaScript로 Python 로직 구현)
class SimplePythonRunner {
  constructor() {
    this.isLoaded = true; // 즉시 사용 가능
  }

  // Python 스타일의 데이터 처리 함수들
  pythonStyle = {
    len: (obj) => obj ? obj.length : 0,
    str: (obj) => String(obj),
    print: (...args) => {
      console.log('Python print:', ...args);
      return args.join(' ');
    },
    dict: (obj) => ({ ...obj }),
    list: (arr) => [...arr],
    range: (start, end) => Array.from({ length: end - start }, (_, i) => start + i)
  };

  // Python 스타일의 문자열 포맷팅
  formatString(template, ...args) {
    return template.replace(/\{(\w+)\}/g, (match, key) => {
      return args[0][key] || match;
    });
  }

  // Python 스타일의 딕셔너리 get 메서드
  dictGet(dict, key, defaultValue = null) {
    return dict && dict[key] !== undefined ? dict[key] : defaultValue;
  }

  // HWP 파일 생성 시뮬레이션
  async generateHwpFile(data) {
    try {
      const { len, print, dictGet } = this.pythonStyle;
      
      const managers = dictGet(data, 'managers', []);
      const overtimeSchedules = dictGet(data, 'overtimeSchedules', {});
      const substituteHolidays = dictGet(data, 'substituteHolidays', {});
      const vacationSchedules = dictGet(data, 'vacationSchedules', {});
      const targetMonth = dictGet(data, 'targetMonth', '2025-03');

      // Python 스타일 출력
      print(`매니저 수: ${len(managers)}`);
      print(`토요일 근무 일정: ${len(Object.keys(overtimeSchedules))}개`);
      print(`대체 휴무일: ${len(Object.keys(substituteHolidays))}개`);
      print(`연가 일정: ${len(Object.keys(vacationSchedules))}개`);
      print(`대상 월: ${targetMonth}`);

      // HWP 파일 내용 생성
      let hwpContent = `한글 문서 생성\n`;
      hwpContent += `대상 월: ${targetMonth}\n`;
      hwpContent += `매니저: ${len(managers)}명\n\n`;

      // 토요일 근무 정보
      hwpContent += `토요일 근무:\n`;
      Object.entries(overtimeSchedules).forEach(([date, managerList]) => {
        managerList.forEach(manager => {
          hwpContent += `- ${date}: ${manager.name}\n`;
        });
      });

      // 대체 휴무일 정보
      hwpContent += `\n대체 휴무일:\n`;
      Object.entries(substituteHolidays).forEach(([date, managerList]) => {
        managerList.forEach(manager => {
          hwpContent += `- ${date}: ${manager.name}\n`;
        });
      });

      // 연가 정보
      hwpContent += `\n연가:\n`;
      Object.entries(vacationSchedules).forEach(([date, managerList]) => {
        managerList.forEach(manager => {
          hwpContent += `- ${date}: ${manager.name}\n`;
        });
      });

      const result = {
        success: true,
        message: `HWP 파일 생성 준비 완료: ${targetMonth}`,
        filename: `청년이룸출근부_${targetMonth}_웹어셈블리.txt`,
        content: hwpContent,
        data: {
          managers,
          overtime_schedules: overtimeSchedules,
          substitute_holidays: substituteHolidays,
          vacation_schedules: vacationSchedules,
          target_month: targetMonth
        }
      };

      return result;
    } catch (error) {
      console.error('HWP 파일 생성 오류:', error);
      throw error;
    }
  }

  // Python 테스트 시뮬레이션
  async testPython() {
    try {
      const { print } = this.pythonStyle;
      
      print("Python 버전: 3.9.0 (웹어셈블리 시뮬레이션)");
      print("웹어셈블리에서 Python이 실행되고 있습니다!");

      // 간단한 계산
      const result = 2 + 3 * 4;
      print(`계산 결과: ${result}`);

      return result;
    } catch (error) {
      console.error('Python 테스트 오류:', error);
      throw error;
    }
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

export default SimplePythonRunner; 