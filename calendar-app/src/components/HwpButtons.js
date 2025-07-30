import React, { useState, useRef } from 'react';
import './HwpButtons.css';
import HwpPythonRunner from '../utils/hwpPythonRunner';

const HwpButtons = ({ managers, savedData, currentDate }) => {
  const [isGenerating, setIsGenerating] = useState(false);
  const [message, setMessage] = useState('');
  const [pythonRunner] = useState(new HwpPythonRunner());
  // 템플릿 파일 상태 추가
  const [templateFile, setTemplateFile] = useState(null);
  const [templateFileName, setTemplateFileName] = useState('');
  // input ref 추가
  const fileInputRef = useRef();
  // 매니저 선택 상태
  const [selectedManager, setSelectedManager] = useState('');

  // 템플릿 파일 업로드 핸들러
  const handleTemplateUpload = (e) => {
    const file = e.target.files[0];
    if (file) {
      setTemplateFile(file);
      setTemplateFileName(file.name);
      setMessage(`템플릿 파일 업로드 완료: ${file.name}`);
    }
  };

  // 버튼 클릭 시 input 클릭
  const handleUploadButtonClick = () => {
    if (fileInputRef.current) {
      fileInputRef.current.click();
    }
  };

  // 매니저 지정 생성
  const handleSingleManagerPrint = async () => {
    if (!templateFile) {
      setMessage('템플릿 파일을 먼저 업로드하세요.');
      return;
    }
    if (!selectedManager) {
      setMessage('매니저를 선택하세요.');
      return;
    }
    setIsGenerating(true);
    setMessage('HWP 파일을 생성하고 있습니다...');
    
    try {
      const targetMonth = `${currentDate.getFullYear()}-${String(currentDate.getMonth() + 1).padStart(2, '0')}`;
      
      // 템플릿 파일을 ArrayBuffer로 변환
      const templateBuffer = await templateFile.arrayBuffer();
      
      // 달력 데이터 수집
      const calendarData = {
        manager: selectedManager,
        targetMonth: targetMonth,
        year: currentDate.getFullYear(),
        month: currentDate.getMonth() + 1,
        overtimeSchedules: savedData?.overtimeSchedules || {},
        substituteHolidays: savedData?.substituteHolidays || {},
        vacationSchedules: savedData?.vacationSchedules || {},
        paidHolidays: savedData?.paidHolidays || []
      };
      
      console.log('전달할 데이터:', calendarData);
      
      // Pyodide로 HWP 파일 생성
      const result = await pythonRunner.generateHwpFileForManager(
        templateBuffer,
        templateFileName,
        calendarData
      );
      
      if (result && result.success) {
        setMessage(`✅ ${selectedManager} 매니저 HWP 파일 생성 완료!`);
        console.log('HWP 파일 생성 결과:', result);
        
        // 파일 다운로드
        if (result.filename && result.content) {
          const downloadSuccess = await pythonRunner.downloadFile(result.filename, result.content);
          if (downloadSuccess) {
            setMessage(`✅ ${selectedManager} 매니저 HWP 파일이 다운로드되었습니다!\\n📁 브라우저 다운로드 폴더를 확인하세요.`);
          } else {
            setMessage(`❌ 파일 다운로드 실패`);
          }
        }
      } else {
        setMessage(`❌ HWP 파일 생성 실패: ${result?.message || '알 수 없는 오류'}`);
      }
    } catch (error) {
      console.error('HWP 파일 생성 오류:', error);
      setMessage(`❌ 오류: ${error.message}`);
    } finally {
      setIsGenerating(false);
    }
  };

  // 전체 매니저 생성
  const handleAllManagersPrint = async () => {
    if (!templateFile) {
      setMessage('템플릿 파일을 먼저 업로드하세요.');
      return;
    }
    if (!managers || managers.length === 0) {
      setMessage('매니저를 먼저 추가하세요.');
      return;
    }
    setIsGenerating(true);
    setMessage('모든 매니저 HWP 파일을 생성하고 있습니다...');
    
    try {
      const targetMonth = `${currentDate.getFullYear()}-${String(currentDate.getMonth() + 1).padStart(2, '0')}`;
      
      // 템플릿 파일을 ArrayBuffer로 변환
      const templateBuffer = await templateFile.arrayBuffer();
      
      // 달력 데이터 수집
      const calendarData = {
        managers: managers,
        targetMonth: targetMonth,
        year: currentDate.getFullYear(),
        month: currentDate.getMonth() + 1,
        overtimeSchedules: savedData?.overtimeSchedules || {},
        substituteHolidays: savedData?.substituteHolidays || {},
        vacationSchedules: savedData?.vacationSchedules || {},
        paidHolidays: savedData?.paidHolidays || []
      };
      
      console.log('전달할 데이터:', calendarData);
      
      // Pyodide로 모든 매니저 HWP 파일 생성
      const result = await pythonRunner.generateHwpFilesForAllManagers(
        templateBuffer,
        templateFileName,
        calendarData
      );
      
      if (result && result.success) {
        setMessage(`✅ 모든 매니저 HWP 파일 생성 완료! (${result.generated_files?.length || 0}개)`);
        console.log('HWP 파일 생성 결과:', result);
        
        // 각 파일 다운로드
        if (result.generated_files) {
          let downloadCount = 0;
          for (const fileInfo of result.generated_files) {
            if (fileInfo.filename && fileInfo.content) {
              const downloadSuccess = await pythonRunner.downloadFile(fileInfo.filename, fileInfo.content);
              if (downloadSuccess) {
                downloadCount++;
              }
            }
          }
          setMessage(`✅ ${downloadCount}개 HWP 파일이 다운로드되었습니다!\\n📁 브라우저 다운로드 폴더를 확인하세요.`);
        }
      } else {
        setMessage(`❌ HWP 파일 생성 실패: ${result?.message || '알 수 없는 오류'}`);
      }
    } catch (error) {
      console.error('HWP 파일 생성 오류:', error);
      setMessage(`❌ 오류: ${error.message}`);
    } finally {
      setIsGenerating(false);
    }
  };

  return (
    <div className="hwp-buttons-container">
      <h2 className="hwp-title">웹어셈블리 HWP 생성</h2>
      {/* 템플릿 업로드 UI */}
      <div className="template-upload-section">
        <input
          type="file"
          accept=".hwp"
          ref={fileInputRef}
          style={{ display: 'none' }}
          onChange={handleTemplateUpload}
        />
        <button
          type="button"
          className="hwp-button"
          style={{ padding: '10px 18px', fontSize: '0.95rem', marginBottom: '6px' }}
          onClick={handleUploadButtonClick}
        >
          📂 템플릿 파일 업로드
        </button>
        {templateFileName && (
          <div style={{ fontSize: '0.9rem', color: '#333', marginTop: '2px' }}>
            선택됨: {templateFileName}
          </div>
        )}
      </div>

      {/* 매니저 선택 UI */}
      <div style={{ marginBottom: '18px', textAlign: 'center' }}>
        <div style={{ fontSize: '0.9rem', fontWeight: '600', marginBottom: '8px', color: '#333' }}>
          매니저 선택:
        </div>
        <div className="manager-list" style={{ display: 'flex', flexWrap: 'wrap', gap: '8px', justifyContent: 'center' }}>
          {managers && managers.map((manager, idx) => (
            <button
              key={manager.name || idx}
              className={`manager-item ${selectedManager === manager.name ? 'selected' : ''}`}
              onClick={() => setSelectedManager(manager.name)}
              style={{
                padding: '6px 12px',
                borderRadius: '20px',
                border: selectedManager === manager.name ? '2px solid #4ECDC4' : '1px solid #ddd',
                background: selectedManager === manager.name ? '#4ECDC4' : '#f8f9fa',
                color: selectedManager === manager.name ? 'white' : '#333',
                cursor: 'pointer',
                fontSize: '0.85rem',
                fontWeight: selectedManager === manager.name ? '600' : '500',
                transition: 'all 0.2s ease'
              }}
            >
              {manager.name}
            </button>
          ))}
        </div>
        {selectedManager && (
          <div style={{ fontSize: '0.8rem', color: '#666', marginTop: '6px' }}>
            선택됨: {selectedManager}
          </div>
        )}
      </div>

      <div className="hwp-buttons">
        <button
          className="hwp-button single-print-button"
          onClick={handleSingleManagerPrint}
          disabled={isGenerating}
        >
          <div className="button-icon">{isGenerating ? '⏳' : '👤'}</div>
          <span className="button-text">{isGenerating ? '생성 중...' : 'HWP 매니저 지정 생성'}</span>
        </button>
        <button
          className="hwp-button print-all-button"
          onClick={handleAllManagersPrint}
          disabled={isGenerating}
        >
          <div className="button-icon">{isGenerating ? '⏳' : '👥'}</div>
          <span className="button-text">{isGenerating ? '생성 중...' : 'HWP 매니저 모두 생성'}</span>
        </button>
      </div>

      {/* 메시지 표시 */}
      {message && (
        <div className="hwp-message">
          {message}
        </div>
      )}

      {/* 다운로드 폴더 안내 */}
      <div className="download-info">
        <p>💡 <strong>다운로드 폴더:</strong> 브라우저 설정의 기본 다운로드 폴더에 저장됩니다.</p>
        <p>📁 <strong>파일명:</strong> 청년이룸출근부_매니저이름_YYMM.hwp</p>
      </div>

      <div className="hwp-info">
        <p>🚀 웹어셈블리 HWP 생성 사용법:</p>
        <ol>
          <li>템플릿 파일을 업로드하세요</li>
          <li>매니저를 선택하거나 모두 생성 버튼을 누르세요</li>
          <li>생성된 HWP 파일이 다운로드됩니다</li>
        </ol>
        <p>💡 웹어셈블리로 브라우저에서 직접 Python을 실행하여 HWP 파일을 생성합니다!</p>
      </div>
    </div>
  );
};

export default HwpButtons; 