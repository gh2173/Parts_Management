// login.js - 로그인 인증 서비스
const bcrypt = require('bcrypt');
const Store = require('electron-store');
const ExcelJS = require('exceljs');

class AuthService {
  constructor(ftpService) {
    this.ftpService = ftpService;
    this.sessionStore = new Store({ name: 'user-session' });
    this.isAuthenticated = false;
    this.currentUser = null;
    
    // 앱 시작시 세션 복원
    this.restoreSession();
  }

  // 세션 복원
  restoreSession() {
    try {
      const savedAuth = this.sessionStore.get('isAuthenticated', false);
      const savedUser = this.sessionStore.get('currentUser', null);
      
      if (savedAuth && savedUser) {
        this.isAuthenticated = true;
        this.currentUser = savedUser;
        console.log('세션 복원됨:', savedUser.username);
      }
    } catch (error) {
      console.error('세션 복원 중 오류:', error);
    }
  }

  // 관리자 등록
  async registerAdmin(userData) {
    try {
      console.log('관리자 등록 요청:', userData.username);
      
      // 최신 파일 다운로드
      await this.ftpService.downloadFile();
      
      // 엑셀 파일 로드
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.readFile(this.ftpService.localTempFile);
      
      // 세 번째 시트 (관리자 계정 정보) 가져오기 또는 생성
      let adminSheet = workbook.getWorksheet(3);
      if (!adminSheet) {
        console.log('관리자_계정 시트 생성');
        adminSheet = workbook.addWorksheet('관리자_계정');
        
        // 헤더 추가 (A: ID, B: PW, C: PW찾기코드)
        const headerRow = adminSheet.addRow(['ID', 'PW', 'PW찾기코드']);
        headerRow.font = { bold: true };
      }
      
      // 기존 사용자 중복 체크 (A열에서 ID 확인)
      const userExists = this.checkUserExistsInAdminSheet(adminSheet, userData.username);
      if (userExists) {
        return { success: false, error: '이미 존재하는 사용자명입니다.' };
      }
      
      // 관리자 계정 데이터를 세 번째 시트의 A, B, C열에 추가
      adminSheet.addRow([
        userData.username,        // A열: ID
        userData.password,        // B열: PW
        userData.recoveryCode     // C열: PW찾기코드 (사용자 입력)
      ]);
      
      // 파일 저장 및 업로드
      await workbook.xlsx.writeFile(this.ftpService.localTempFile);
      await this.ftpService.uploadFile();
      
      console.log('관리자 등록 완료:', userData.username);
      return { 
        success: true, 
        message: '관리자가 성공적으로 등록되었습니다.'
      };
      
    } catch (error) {
      console.error('관리자 등록 중 오류:', error);
      return { success: false, error: error.message };
    }
  }

  // 사용자 로그인
  async login(loginData) {
    try {
      console.log('로그인 시도:', loginData.username);
      
      // 최신 파일 다운로드
      await this.ftpService.downloadFile();
      
      // 엑셀 파일 로드
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.readFile(this.ftpService.localTempFile);
      
      // 세 번째 시트 (관리자 계정) 가져오기
      const adminSheet = workbook.getWorksheet(3);
      if (!adminSheet) {
        return { success: false, error: '관리자 계정 정보를 찾을 수 없습니다.' };
      }
      
      // 사용자 찾기 및 비밀번호 확인
      const userInfo = this.validateAdminUser(adminSheet, loginData);
      
      if (!userInfo) {
        return { success: false, error: '사용자명 또는 비밀번호가 올바르지 않습니다.' };
      }
      
      // 세션 정보 저장
      this.isAuthenticated = true;
      this.currentUser = userInfo;
      this.sessionStore.set('isAuthenticated', true);
      this.sessionStore.set('currentUser', userInfo);
      
      console.log('로그인 성공:', userInfo.username);
      return { success: true, user: userInfo };
      
    } catch (error) {
      console.error('로그인 중 오류:', error);
      return { success: false, error: error.message };
    }
  }

  // 로그아웃
  async logout() {
    try {
      this.isAuthenticated = false;
      this.currentUser = null;
      this.sessionStore.clear();
      
      console.log('로그아웃 완료');
      return { success: true };
    } catch (error) {
      console.error('로그아웃 중 오류:', error);
      return { success: false, error: error.message };
    }
  }

  // 인증 상태 확인
  checkAuthStatus() {
    return {
      isAuthenticated: this.isAuthenticated,
      user: this.currentUser
    };
  }

  // 권한 확인 미들웨어
  requireAuth() {
    if (!this.isAuthenticated) {
      throw new Error('로그인이 필요합니다.');
    }
    return true;
  }

  // 사용자 존재 여부 확인 (기존 네 번째 시트용)
  checkUserExists(userSheet, username) {
    let userExists = false;
    userSheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return; // 헤더 건너뛰기
      
      const existingUsername = row.getCell(2).value;
      if (existingUsername && existingUsername.toString() === username) {
        userExists = true;
      }
    });
    return userExists;
  }

  // 관리자 시트에서 사용자 존재 여부 확인 (세 번째 시트 A열 확인)
  checkUserExistsInAdminSheet(adminSheet, username) {
    let userExists = false;
    adminSheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return; // 헤더 건너뛰기
      
      const existingUsername = row.getCell(1).value; // A열: ID
      if (existingUsername && existingUsername.toString() === username) {
        userExists = true;
      }
    });
    return userExists;
  }

  // 사용자 검증 (기존 네 번째 시트용)
  async validateUser(userSheet, loginData) {
    let userInfo = null;
    
    for (let rowNumber = 2; rowNumber <= userSheet.rowCount; rowNumber++) {
      const row = userSheet.getRow(rowNumber);
      const username = row.getCell(2).value;
      const hashedPassword = row.getCell(3).value;
      
      if (username && username.toString() === loginData.username) {
        const isValidPassword = await bcrypt.compare(loginData.password, hashedPassword);
        
        if (isValidPassword) {
          userInfo = {
            id: row.getCell(1).value,
            username: username.toString(),
            role: row.getCell(4).value || 'user',
            lastLogin: new Date().toISOString()
          };
          break;
        }
      }
    }
    
    return userInfo;
  }

  // 관리자 사용자 검증 (세 번째 시트용)
  validateAdminUser(adminSheet, loginData) {
    let userInfo = null;
    
    adminSheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return; // 헤더 건너뛰기
      
      const username = row.getCell(1).value; // A열: ID
      const password = row.getCell(2).value; // B열: PW
      const recoveryCode = row.getCell(3).value; // C열: PW찾기코드
      
      if (username && username.toString() === loginData.username) {
        // 평문 비밀번호 비교
        if (password && password.toString() === loginData.password) {
          userInfo = {
            id: username.toString(),
            username: username.toString(),
            role: 'admin',
            recoveryCode: recoveryCode ? recoveryCode.toString() : '',
            lastLogin: new Date().toISOString()
          };
        }
      }
    });
    
    return userInfo;
  }

  // 마지막 로그인 시간 업데이트
  async updateLastLogin(userSheet, username) {
    userSheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return; // 헤더 건너뛰기
      
      const existingUsername = row.getCell(2).value;
      if (existingUsername && existingUsername.toString() === username) {
        row.getCell(6).value = new Date().toISOString().split('T')[0];
      }
    });
  }

  // 비밀번호 찾기
  async findPassword(findData) {
    try {
      console.log('비밀번호 찾기 요청:', findData.username);
      
      // 최신 파일 다운로드
      await this.ftpService.downloadFile();
      
      // 엑셀 파일 로드
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.readFile(this.ftpService.localTempFile);
      
      // 세 번째 시트 (관리자 계정) 가져오기
      const adminSheet = workbook.getWorksheet(3);
      if (!adminSheet) {
        return { success: false, error: '관리자 계정 정보를 찾을 수 없습니다.' };
      }
      
      // 사용자 찾기
      let foundUser = null;
      adminSheet.eachRow((row, rowNumber) => {
        if (rowNumber === 1) return; // 헤더 건너뛰기
        
        const username = row.getCell(1).value; // A열: ID
        const password = row.getCell(2).value; // B열: PW
        const recoveryCode = row.getCell(3).value; // C열: PW찾기코드
        
        if (username && username.toString() === findData.username) {
          if (recoveryCode && recoveryCode.toString() === findData.recoveryCode) {
            foundUser = {
              username: username.toString(),
              password: password ? password.toString() : '',
              recoveryCode: recoveryCode.toString()
            };
          }
        }
      });
      
      if (foundUser) {
        console.log('비밀번호 찾기 성공:', foundUser.username);
        return { 
          success: true, 
          password: foundUser.password,
          message: '비밀번호를 찾았습니다.'
        };
      } else {
        return { success: false, error: 'ID 또는 PW찾기코드가 올바르지 않습니다.' };
      }
      
    } catch (error) {
      console.error('비밀번호 찾기 중 오류:', error);
      return { success: false, error: error.message };
    }
  }

  // 비밀번호 변경
  async changePassword(changeData) {
    try {
      console.log('비밀번호 변경 요청:', changeData.username);
      
      // 최신 파일 다운로드
      await this.ftpService.downloadFile();
      
      // 엑셀 파일 로드
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.readFile(this.ftpService.localTempFile);
      
      // 세 번째 시트 (관리자 계정) 가져오기
      const adminSheet = workbook.getWorksheet(3);
      if (!adminSheet) {
        return { success: false, error: '관리자 계정 정보를 찾을 수 없습니다.' };
      }
      
      // 사용자 찾기 및 비밀번호 변경
      let userFound = false;
      adminSheet.eachRow((row, rowNumber) => {
        if (rowNumber === 1) return; // 헤더 건너뛰기
        
        const username = row.getCell(1).value; // A열: ID
        const recoveryCode = row.getCell(3).value; // C열: PW찾기코드
        
        if (username && username.toString() === changeData.username) {
          if (recoveryCode && recoveryCode.toString() === changeData.recoveryCode) {
            // B열(PW)에 새 비밀번호 저장
            row.getCell(2).value = changeData.newPassword;
            userFound = true;
            console.log(`비밀번호 변경 완료 - 행 ${rowNumber}`);
          }
        }
      });
      
      if (!userFound) {
        return { success: false, error: 'ID 또는 PW찾기코드가 올바르지 않습니다.' };
      }
      
      // 파일 저장 및 업로드
      await workbook.xlsx.writeFile(this.ftpService.localTempFile);
      await this.ftpService.uploadFile();
      
      console.log('비밀번호 변경 완료:', changeData.username);
      return { success: true, message: '비밀번호가 성공적으로 변경되었습니다.' };
      
    } catch (error) {
      console.error('비밀번호 변경 중 오류:', error);
      return { success: false, error: error.message };
    }
  }
}

// FTP 서비스 클래스
class FTPService {
  constructor(downloadFileFromFTP, uploadFileToFTP, localTempFile) {
    this.downloadFile = downloadFileFromFTP;
    this.uploadFile = uploadFileToFTP;
    this.localTempFile = localTempFile;
  }
}

module.exports = { AuthService, FTPService };
