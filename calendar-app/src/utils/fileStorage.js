// File System Access API를 사용한 로컬 파일 저장/로드 유틸리티

class FileStorage {
  constructor() {
    this.fileHandle = null;
    this.data = {
      managers: [],
      overtimeSchedules: {},
      substituteHolidays: {},
      vacationSchedules: {}
    };
  }

  // 파일 핸들 초기화
  async initializeFile() {
    try {
      // 기존 파일이 있는지 확인
      const options = {
        types: [{
          description: 'JSON Files',
          accept: {
            'application/json': ['.json'],
          },
        }],
      };

      // 파일 선택 다이얼로그 표시
      const [fileHandle] = await window.showOpenFilePicker(options);
      this.fileHandle = fileHandle;
      
      // 기존 데이터 로드
      await this.loadData();
      return true;
    } catch (error) {
      console.log('파일이 선택되지 않았거나 새 파일을 생성해야 합니다.');
      return false;
    }
  }

  // 새 파일 생성
  async createNewFile() {
    try {
      const options = {
        suggestedName: '청년이룸출근부_데이터.json',
        types: [{
          description: 'JSON Files',
          accept: {
            'application/json': ['.json'],
          },
        }],
      };

      this.fileHandle = await window.showSaveFilePicker(options);
      
      // 파일 정보 출력
      console.log('생성된 파일 정보:', {
        name: this.fileHandle.name,
        kind: this.fileHandle.kind,
        // File System Access API는 보안상 전체 경로를 직접 제공하지 않음
      });
      
      await this.saveData();
      return true;
    } catch (error) {
      console.error('파일 생성 실패:', error);
      return false;
    }
  }

  // 데이터 저장
  async saveData() {
    if (!this.fileHandle) {
      console.error('파일 핸들이 없습니다.');
      return false;
    }

    try {
      const writable = await this.fileHandle.createWritable();
      await writable.write(JSON.stringify(this.data, null, 2));
      await writable.close();
      console.log('데이터가 성공적으로 저장되었습니다.');
      return true;
    } catch (error) {
      console.error('데이터 저장 실패:', error);
      return false;
    }
  }

  // 데이터 로드
  async loadData() {
    if (!this.fileHandle) {
      console.error('파일 핸들이 없습니다.');
      return false;
    }

    try {
      const file = await this.fileHandle.getFile();
      const contents = await file.text();
      const loadedData = JSON.parse(contents);
      
      // 기존 데이터 마이그레이션 (id 필드 제거)
      this.data = this.migrateData(loadedData);
      
      console.log('데이터가 성공적으로 로드되었습니다.');
      return true;
    } catch (error) {
      console.error('데이터 로드 실패:', error);
      return false;
    }
  }

  // 데이터 마이그레이션 (id 필드 제거)
  migrateData(loadedData) {
    const migratedData = {
      managers: [],
      overtimeSchedules: {},
      substituteHolidays: {},
      vacationSchedules: {}
    };

    // 매니저 데이터 마이그레이션
    if (loadedData.managers) {
      migratedData.managers = loadedData.managers.map(manager => {
        // id 필드가 있으면 제거하고 name만 유지
        if (manager.id !== undefined) {
          const { id, ...managerWithoutId } = manager;
          return managerWithoutId;
        }
        return manager;
      });
    }

    // 스케줄 데이터 마이그레이션
    const migrateScheduleData = (scheduleData) => {
      const migrated = {};
      Object.keys(scheduleData).forEach(dateKey => {
        migrated[dateKey] = scheduleData[dateKey].map(manager => {
          // id 필드가 있으면 제거하고 name만 유지
          if (manager.id !== undefined) {
            const { id, ...managerWithoutId } = manager;
            return managerWithoutId;
          }
          return manager;
        });
      });
      return migrated;
    };

    if (loadedData.overtimeSchedules) {
      migratedData.overtimeSchedules = migrateScheduleData(loadedData.overtimeSchedules);
    }
    if (loadedData.substituteHolidays) {
      migratedData.substituteHolidays = migrateScheduleData(loadedData.substituteHolidays);
    }
    if (loadedData.vacationSchedules) {
      migratedData.vacationSchedules = migrateScheduleData(loadedData.vacationSchedules);
    }

    return migratedData;
  }

  // 매니저 데이터 업데이트
  updateManagers(managers) {
    this.data.managers = managers;
  }

  // 달력 데이터 업데이트
  updateCalendarData(overtimeSchedules, substituteHolidays, vacationSchedules) {
    this.data.overtimeSchedules = overtimeSchedules;
    this.data.substituteHolidays = substituteHolidays;
    this.data.vacationSchedules = vacationSchedules;
  }

  // 전체 데이터 가져오기
  getData() {
    return this.data;
  }

  // 파일 핸들 가져오기
  getFileHandle() {
    return this.fileHandle;
  }

  // 파일이 초기화되었는지 확인
  isInitialized() {
    return this.fileHandle !== null;
  }
}

export default FileStorage; 