import React, { useState } from 'react';
import './DataExporter.css';

const DataExporter = ({ managers, savedData, currentDate }) => {
  const [exportedData, setExportedData] = useState(null);
  const [message, setMessage] = useState('');

  // 데이터를 JSON 파일로 저장
  const exportDataForPython = () => {
    try {
      const exportData = {
        managers: managers,
        calendarData: {
          overtimeSchedules: savedData?.overtimeSchedules || {},
          substituteHolidays: savedData?.substituteHolidays || {},
          vacationSchedules: savedData?.vacationSchedules || {}
        },
        currentDate: {
          year: currentDate.getFullYear(),
          month: currentDate.getMonth() + 1,
          targetMonth: `${currentDate.getFullYear()}-${String(currentDate.getMonth() + 1).padStart(2, '0')}`
        },
        exportTime: new Date().toISOString(),
        version: "1.0"
      };

      // JSON 파일 생성
      const dataStr = JSON.stringify(exportData, null, 2);
      const dataBlob = new Blob([dataStr], { type: 'application/json' });
      
      // 파일명 생성
      const timestamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);
      const filename = `calendar_data_${timestamp}.json`;
      
      // 다운로드 링크 생성
      const url = URL.createObjectURL(dataBlob);
      const link = document.createElement('a');
      link.href = url;
      link.download = filename;
      link.style.display = 'none';
      
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      
      // 메모리 정리
      setTimeout(() => {
        URL.revokeObjectURL(url);
      }, 100);
      
      setExportedData(exportData);
      setMessage(`✅ 데이터가 성공적으로 저장되었습니다!\n📁 파일명: ${filename}\n💡 이 파일을 Python 스크립트와 같은 폴더에 저장하세요.`);
      
    } catch (error) {
      setMessage(`❌ 데이터 저장 실패: ${error.message}`);
    }
  };

  // Python 스크립트 실행 안내
  const showPythonInstructions = () => {
    const instructions = `
🚀 Python 스크립트 실행 방법:

1. 위에서 저장한 JSON 파일을 Python 스크립트와 같은 폴더에 저장
2. 명령 프롬프트(cmd)에서 다음 명령어 실행:

   python hwp_generator.py calendar_data_YYYY-MM-DDTHH-MM-SS.json

3. 또는 Python 스크립트를 직접 실행:

   python hwp_generator.py

   (스크립트가 자동으로 최신 JSON 파일을 찾아서 실행)

📝 저장된 데이터:
- 매니저 수: ${managers?.length || 0}명
- 토요일 근무: ${Object.keys(savedData?.overtimeSchedules || {}).length}일
- 대체 휴무일: ${Object.keys(savedData?.substituteHolidays || {}).length}일
- 연가: ${Object.keys(savedData?.vacationSchedules || {}).length}일
- 대상 월: ${currentDate.getFullYear()}년 ${currentDate.getMonth() + 1}월
    `;
    
    setMessage(instructions);
  };

  return (
    <div className="data-exporter-container">
      <h2 className="data-exporter-title">📊 데이터 내보내기</h2>
      
      <div className="export-section">
        <button
          onClick={exportDataForPython}
          className="export-btn"
          disabled={!managers || managers.length === 0}
        >
          💾 Python용 데이터 저장
        </button>
        
        <button
          onClick={showPythonInstructions}
          className="instructions-btn"
        >
          📋 실행 방법 보기
        </button>
      </div>

      {/* 저장된 데이터 미리보기 */}
      {exportedData && (
        <div className="data-preview">
          <h3>저장된 데이터 미리보기:</h3>
          <div className="preview-content">
            <p><strong>매니저:</strong> {exportedData.managers.map(m => m.name).join(', ')}</p>
            <p><strong>대상 월:</strong> {exportedData.currentDate.targetMonth}</p>
            <p><strong>저장 시간:</strong> {new Date(exportedData.exportTime).toLocaleString()}</p>
          </div>
        </div>
      )}

      {/* 메시지 표시 */}
      {message && (
        <div className="exporter-message">
          <pre>{message}</pre>
        </div>
      )}

      {/* 사용법 안내 */}
      <div className="exporter-info">
        <h3>💡 사용법:</h3>
        <ol>
          <li>웹에서 데이터를 입력하고 저장하세요</li>
          <li>"Python용 데이터 저장" 버튼을 클릭하세요</li>
          <li>다운로드된 JSON 파일을 Python 스크립트 폴더에 저장하세요</li>
          <li>명령 프롬프트에서 Python 스크립트를 실행하세요</li>
        </ol>
        <p><strong>장점:</strong> 웹에서 데이터 입력 → 파일 저장 → 수동 실행</p>
        <p><strong>단점:</strong> 자동화되지 않음, 수동 단계 필요</p>
      </div>
    </div>
  );
};

export default DataExporter; 