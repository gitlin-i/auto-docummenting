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
      // 실제 Pyodide 연동 시 templateFile, selectedManager, 날짜 등 전달
      // 아래는 예시
      const targetMonth = `${currentDate.getFullYear()}-${String(currentDate.getMonth() + 1).padStart(2, '0')}`;
      // await pythonRunner.generateHwpFileForManager(templateFile, selectedManager, targetMonth);
      setMessage(`✅ ${selectedManager} 매니저 HWP 파일 생성 완료 (시뮬레이션)`);
    } catch (error) {
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
      // 실제 Pyodide 연동 시 templateFile, managers, 날짜 등 전달
      // await pythonRunner.generateHwpFiles(templateFile, managers, targetMonth);
      setMessage(`✅ 모든 매니저 HWP 파일 생성 완료 (시뮬레이션)`);
    } catch (error) {
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

      {message && (
        <div className={`message ${message.includes('✅') ? 'success' : message.includes('❌') ? 'error' : 'info'}`}>
          {message}
        </div>
      )}

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