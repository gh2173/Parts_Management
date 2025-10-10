const { app, BrowserWindow, ipcMain } = require('electron');
const path = require('path');
const ExcelJS = require('exceljs');
const fs = require('fs');
const os = require('os');
const ftp = require('basic-ftp');
const fse = require('fs-extra');

// 로그인 서비스 import
const { AuthService, FTPService } = require('./login');

// 메인 윈도우 전역 변수로 선언 (가비지 컬렉션 방지)
let mainWindow;

// 인증 서비스 인스턴스
let authService;

// =============== FTP 서버 설정 (운영용) ===============
const ftpConfig = {
  host: "192.168.223.225",
  user: "vega",
  password: "vegagcc",
  secure: false
};

// =============== FTP 파일 경로 (운영용) ===============
const ftpFilePath = "/untitle/장비솔루션_파트재고.xlsx";
const localTempDir = path.join(os.tmpdir(), 'electron-app-temp');
const localTempFile = path.join(localTempDir, '장비솔루션_파트재고.xlsx');

// =============== 임시 디렉토리 생성 (운영용) ===============
if (!fs.existsSync(localTempDir)) {
  fse.ensureDirSync(localTempDir);
}

// =============== FTP 서버 함수들 (운영용) ===============
async function downloadFileFromFTP() {
  const client = new ftp.Client();
  client.ftp.verbose = false; // 디버깅 메시지 끄기

  try {
    console.log('FTP 서버 연결 중...');
    await client.access(ftpConfig);
    console.log('FTP 서버 연결 성공');

    console.log(`파일 다운로드 중: ${ftpFilePath}`);
    await client.downloadTo(localTempFile, ftpFilePath);
    console.log(`파일 다운로드 완료: ${localTempFile}`);

    return true;
  } catch (error) {
    console.error('FTP 다운로드 오류:', error);
    throw error;
  } finally {
    client.close();
  }
}

// =============== FTP 업로드 함수 (운영용) ===============
async function uploadFileToFTP() {
  const client = new ftp.Client();
  client.ftp.verbose = false; // 디버깅 메시지 끄기

  try {
    console.log('FTP 서버 연결 중...');
    await client.access(ftpConfig);
    console.log('FTP 서버 연결 성공');

    console.log(`파일 업로드 중: ${ftpFilePath}`);
    await client.uploadFrom(localTempFile, ftpFilePath);
    console.log(`파일 업로드 완료: ${ftpFilePath}`);

    return true;
  } catch (error) {
    console.error('FTP 업로드 오류:', error);
    throw error;
  } finally {
    client.close();
  }
}

// 브라우저 창 생성
function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1200,
    height: 800,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      nodeIntegration: false,
      contextIsolation: true
    }
  });

  // 인증 상태 확인 후 적절한 페이지 로드
  checkAndLoadPage();

  // 개발 도구 열기 (필요하면 주석 해제)
  // mainWindow.webContents.openDevTools();
}

// 페이지 로드 확인
async function checkAndLoadPage() {
  try {
    const authStatus = authService.checkAuthStatus();
    
    if (authStatus.isAuthenticated) {
      mainWindow.loadFile('재고목록.html');
    } else {
      mainWindow.loadFile('login.html');
    }
  } catch (error) {
    console.error('페이지 로드 확인 중 오류:', error);
    mainWindow.loadFile('login.html');
  }
}

// 앱이 준비되면 창 생성
app.whenReady().then(async () => {
  // 앱 시작 시 FTP에서 최신 파일 다운로드
  try {
    await downloadFileFromFTP();
    console.log('초기 파일 다운로드 완료');
  } catch (error) {
    console.error('초기 파일 다운로드 실패:', error);
  }

  // FTP 서비스 및 인증 서비스 초기화
  const ftpService = new FTPService(downloadFileFromFTP, uploadFileToFTP, localTempFile);
  authService = new AuthService(ftpService);
  
  createWindow();

  app.on('activate', function () {
    // macOS에서는 창이 모두 닫혀도 앱이 종료되지 않음
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});

// 모든 창이 닫히면 앱 종료 (Windows와 Linux)
app.on('window-all-closed', function () {
  if (process.platform !== 'darwin') app.quit();
});

// 부품 등록 처리 핸들러 수정 (권한 보호 추가)
ipcMain.handle('register-part', requireAuth(async (event, partData) => {
    try {
      console.log('부품 등록 요청 받음:', partData);
      
      // 최신 파일 다운로드
      await downloadFileFromFTP();
      
      // 엑셀 파일 로드
      const workbook = new ExcelJS.Workbook();
      try {
        await workbook.xlsx.readFile(localTempFile);
      } catch (error) {
        console.error('엑셀 파일을 읽을 수 없습니다:', error);
        return { success: false, error: '엑셀 파일을 읽을 수 없습니다.' };
      }
      
      // 첫 번째 시트 가져오기
      const sheet = workbook.getWorksheet(1);
      
      // 시트 열 헤더 확인하기
      console.log('시트 열 헤더:', sheet.getRow(1).values);
        // 명시적으로 데이터 배열을 만들어서 추가 (13개 컬럼)
      const rowData = [
        partData.location,    // 위치 (1열)
        partData.name,        // 품명 (2열)
        partData.company,     // 구매처 (3열)
        partData.category || '기타', // 카테고리 (4열)
        parseInt(partData.minStock || 20), // 최소 재고 (5열)
        partData.inDate || new Date().toISOString().split('T')[0], // 입고일 (6열)
        0, // 현재 재고 (7열) - 항상 0
        partData.purchasePrice || '', // 구매금액 (8열)
        '', // 설비군 (9열) - 항상 빈 값 (출고 시 입력)
        '', // BOARD명 (10열) - 항상 빈 값 (출고 시 입력)
        '', // S/N (11열) - 항상 빈 값 (출고 시 입력)
        partData.operator || '', // 작업자 (12열)
        parseFloat(partData.unitPrice || 0) // 단가 (13열)
      ];
      
      console.log('추가할 행 데이터:', rowData);
      
      // 새 행 추가
      const newRow = sheet.addRow(rowData);
      
      // 추가된 행 확인
      console.log('추가된 행:', newRow.values);
      
      // 파일 저장
      await workbook.xlsx.writeFile(localTempFile);
      
      // FTP 서버로 업로드
      await uploadFileToFTP();
        console.log('부품 등록 완료');
      return { success: true };
    } catch (error) {
      console.error('부품 등록 중 오류 발생:', error);
      return { success: false, error: error.message };
    }
  }));

// main.js에 추가 - 품명 중복 체크 API
ipcMain.handle('check-duplicate-part', async (event, partName) => {
  try {
    console.log('품명 중복 체크 요청:', partName);
    
    // 최신 파일 다운로드
    await downloadFileFromFTP();
    
    // 엑셀 파일 로드
    const workbook = new ExcelJS.Workbook();
    try {
      await workbook.xlsx.readFile(localTempFile);
    } catch (error) {
      console.error('엑셀 파일을 읽을 수 없습니다:', error);
      return { success: false, error: '엑셀 파일을 읽을 수 없습니다.' };
    }
    
    // 첫 번째 시트 가져오기
    const sheet = workbook.getWorksheet(1);
    
    // 품명 열(2열)을 검색하여 중복 체크
    let isDuplicate = false;
    let duplicateInfo = null;
    
    // 2번째 행부터 데이터 검색 (1번째 행은 헤더)
    for (let i = 2; i <= sheet.rowCount; i++) {
      const row = sheet.getRow(i);
      const existingName = row.getCell(2).value; // 품명은 2열에 있음
      
      if (existingName && existingName.toString().toLowerCase() === partName.toLowerCase()) {
        isDuplicate = true;
        duplicateInfo = {
          location: row.getCell(1).value, // 위치
          name: existingName,             // 품명
          company: row.getCell(3).value,  // 구매처
          category: row.getCell(4).value  // 카테고리
        };
        break;
      }
    }
    
    return {
      success: true,
      isDuplicate,
      duplicateInfo
    };
  } catch (error) {
    console.error('품명 중복 체크 중 오류 발생:', error);
    return { success: false, error: error.message };
  }
});

// main.js에 추가 - 기존 DB 데이터 가져오기 API
ipcMain.handle('get-db-data', async (event) => {
  try {
    // 최신 파일 다운로드
    await downloadFileFromFTP();
    
    // 엑셀 파일 로드
    const workbook = new ExcelJS.Workbook();
    try {
      await workbook.xlsx.readFile(localTempFile);
    } catch (error) {
      console.error('엑셀 파일을 읽을 수 없습니다:', error);
      return { success: false, error: '엑셀 파일을 읽을 수 없습니다.' };
    }
    
    // 첫 번째 시트 가져오기
    const sheet = workbook.getWorksheet(1);
    
    if (!sheet) {
      return { success: false, error: '시트를 찾을 수 없습니다.' };
    }
      // 데이터 추출
    const locations = [];
    const names = [];
    const companies = [];
    const categories = [];
    const purchasePrices = [];
    const equipmentGroups = [];
    const boardNames = [];
    const serialNumbers = [];
    const operators = [];
    
    // 2번째 행부터 데이터 추출 (1번째 행은 헤더)
    for (let i = 2; i <= sheet.rowCount; i++) {
      const row = sheet.getRow(i);
      
      // 셀 값이 존재하는 경우에만 추가 (빈 셀, null, undefined 제외)
      const locationValue = row.getCell(1).value;
      const nameValue = row.getCell(2).value;
      const companyValue = row.getCell(3).value;
      const categoryValue = row.getCell(4).value;
      const purchasePriceValue = row.getCell(8).value;
      const equipmentGroupValue = row.getCell(9).value;
      const boardNameValue = row.getCell(10).value;
      const serialNumberValue = row.getCell(11).value;
      const operatorValue = row.getCell(12).value;
      
      if (locationValue) locations.push(String(locationValue));
      if (nameValue) names.push(String(nameValue));
      if (companyValue) companies.push(String(companyValue));
      if (categoryValue) categories.push(String(categoryValue));
      if (purchasePriceValue) purchasePrices.push(String(purchasePriceValue));
      if (equipmentGroupValue) equipmentGroups.push(String(equipmentGroupValue));
      if (boardNameValue) boardNames.push(String(boardNameValue));
      if (serialNumberValue) serialNumbers.push(String(serialNumberValue));
      if (operatorValue) operators.push(String(operatorValue));
    }
    
    // 중복 제거 및 정렬
    const uniqueLocations = [...new Set(locations)].filter(Boolean).sort();
    const uniqueNames = [...new Set(names)].filter(Boolean).sort();
    const uniqueCompanies = [...new Set(companies)].filter(Boolean).sort();
    const uniqueCategories = [...new Set(categories)].filter(Boolean).sort();
    const uniquePurchasePrices = [...new Set(purchasePrices)].filter(Boolean).sort();
    const uniqueEquipmentGroups = [...new Set(equipmentGroups)].filter(Boolean).sort();
    const uniqueBoardNames = [...new Set(boardNames)].filter(Boolean).sort();
    const uniqueSerialNumbers = [...new Set(serialNumbers)].filter(Boolean).sort();
    const uniqueOperators = [...new Set(operators)].filter(Boolean).sort();
    
    console.log('추출된 데이터 수:', {
      locations: uniqueLocations.length,
      names: uniqueNames.length,
      companies: uniqueCompanies.length,
      categories: uniqueCategories.length,
      purchasePrices: uniquePurchasePrices.length,
      equipmentGroups: uniqueEquipmentGroups.length,
      boardNames: uniqueBoardNames.length,
      serialNumbers: uniqueSerialNumbers.length,
      operators: uniqueOperators.length
    });
    
    return {
      success: true,
      data: {
        locations: uniqueLocations,
        names: uniqueNames,
        companies: uniqueCompanies,
        categories: uniqueCategories,
        purchasePrices: uniquePurchasePrices,
        equipmentGroups: uniqueEquipmentGroups,
        boardNames: uniqueBoardNames,
        serialNumbers: uniqueSerialNumbers,
        operators: uniqueOperators
      }
    };
  } catch (error) {
    console.error('데이터 가져오기 중 오류 발생:', error);
    return { success: false, error: error.message };
  }
});

ipcMain.handle('get-excel-data', async () => {
  try {
    // 최신 파일 다운로드
    await downloadFileFromFTP();
      
    // 엑셀 파일 로드
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(localTempFile);

    const sheet = workbook.getWorksheet(1);
    if (!sheet) {
      return { success: false, error: '시트를 찾을 수 없습니다.' };
    }    // 데이터 추출
    const data = [];
    sheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return; // 첫 번째 행(헤더)은 건너뜀

      data.push({
        location: row.getCell(1).value || '', // 위치
        name: row.getCell(2).value || '', // 품명
        company: row.getCell(3).value || '', // 구매처
        category: row.getCell(4).value || '', // 카테고리
        minStock: row.getCell(5).value || 0, // 최소 재고
        inDate: row.getCell(6).value || '', // 입고일
        currentStock: row.getCell(7).value || 0, // 현재 재고
        purchasePrice: row.getCell(8).value || '', // 구매금액 (8열)
        equipmentGroup: row.getCell(9).value || '', // 설비군 (9열)
        boardName: row.getCell(10).value || '', // BOARD명 (10열)
        serialNumber: row.getCell(11).value || '', // S/N (11열)
        operator: row.getCell(12).value || '', // 작업자 (12열)
        unitPrice: row.getCell(13).value || 0 // 단가 (13열)
      });
    });

    return { success: true, data };
  } catch (error) {
    console.error('엑셀 데이터 추출 중 오류 발생:', error);
    return { success: false, error: error.message };
  }
});

// 재고 업데이트 처리 핸들러 (권한 보호 추가)
ipcMain.handle('update-stock', requireAuth(async (event, data) => {
  try {
    console.log('재고 업데이트 요청 데이터:', data);
    
    // 최신 파일 다운로드
    await downloadFileFromFTP();
      
    // 파일 존재 확인
    if (!fs.existsSync(localTempFile)) {
      console.error('엑셀 파일이 존재하지 않습니다:', localTempFile);
      return { success: false, error: '엑셀 파일이 존재하지 않습니다.' };
    }
    
    // 엑셀 파일 로드
    const workbook = new ExcelJS.Workbook();
    try {
      await workbook.xlsx.readFile(localTempFile);
    } catch (fileError) {
      console.error('엑셀 파일 읽기 오류:', fileError);
      return { success: false, error: '엑셀 파일을 읽을 수 없습니다: ' + fileError.message };
    }

    // 첫 번째 시트 (재고 관리)
    const stockSheet = workbook.getWorksheet(1);
    if (!stockSheet) {
      return { success: false, error: '재고 시트를 찾을 수 없습니다.' };
    }    // 재고 업데이트만 처리 (이력은 별도 API에서 처리)

    // 부품 찾기 및 재고 업데이트
    let updated = false;
    let partLocation = '';
    let partCategory = data.category || '기타'; // 카테고리 정보 초기화
    
    console.log('시트 행 수:', stockSheet.rowCount);
    
    stockSheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return; // 헤더 행 건너뛰기
      
      const nameCell = row.getCell(2);
      const name = nameCell.value;
      
      // 문자열로 변환하여 비교 (대소문자 구분 없이)
      if (name && name.toString().toLowerCase() === data.name.toString().toLowerCase()) {
        console.log(`일치하는 품명 찾음 - 행 ${rowNumber}`);
        
        // 위치 정보 가져오기
        partLocation = row.getCell(1).value || '';
        
        // 카테고리 정보 가져오기 (제공되지 않은 경우)
        if (!data.category) {
          partCategory = row.getCell(4).value || '기타';
        }
        
        // 현재 재고값 가져오기
        const currentStockCell = row.getCell(7);
        const oldStock = currentStockCell.value || 0;
        console.log(`기존 재고: ${oldStock}, 새 재고: ${data.currentStock}`);
        
        // 현재 재고 업데이트 (7번 셀)
        currentStockCell.value = parseInt(data.currentStock);
        updated = true;
        
        console.log(`재고 업데이트 완료 - 행 ${rowNumber}`);
      }
    });    if (!updated) {
      console.error('해당 부품을 찾을 수 없습니다:', data.name);
      return { success: false, error: '해당 부품을 찾을 수 없습니다.' };
    }

    // 변경사항 저장 (이력은 별도 API에서 처리)
    try {
      await workbook.xlsx.writeFile(localTempFile);
      console.log('엑셀 파일 저장 성공');
        // FTP 서버에 업로드
      await uploadFileToFTP();
      console.log('FTP 서버에 파일 업로드 성공');
      
    } catch (saveError) {
      console.error('파일 저장/업로드 오류:', saveError);
      return { success: false, error: '변경사항을 저장할 수 없습니다: ' + saveError.message };
    }

    // 로그 기록
    console.log(`재고 업데이트 성공: ${data.name}, ${data.action}, ${data.amount}개, 새 재고: ${data.currentStock}, 카테고리: ${partCategory}`);    return { success: true };
  } catch (error) {
    console.error('재고 업데이트 중 오류 발생:', error);
    return { success: false, error: error.message };
  }
}));

// 입출고 이력 가져오기 API
ipcMain.handle('get-history-data', async () => {
  try {
    // 최신 파일 다운로드
    await downloadFileFromFTP();
      
    // 엑셀 파일 로드
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(localTempFile);

    // 두 번째 시트 (입출고 이력)
    const historySheet = workbook.getWorksheet(2);
    if (!historySheet) {
      return { success: false, error: '입출고 이력 시트를 찾을 수 없습니다.' };
    }

    // 헤더 구조에 따라 위치 정보 포함 여부 확인 (한 번만 실행)
    const headerRow = historySheet.getRow(1);
    const hasLocationColumn = headerRow.getCell(3).value === '위치';
    
    console.log('Header row cells:', {
      cell1: headerRow.getCell(1).value,
      cell2: headerRow.getCell(2).value,
      cell3: headerRow.getCell(3).value,
      cell4: headerRow.getCell(4).value,
      hasLocationColumn: hasLocationColumn
    });
    
    // 데이터 추출
    const historyData = [];
    historySheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return; // 첫 번째 행(헤더)은 건너뜀
      
      // 모든 셀이 비어있는 행은 건너뛰기
      const isEmpty = !row.getCell(1).value && !row.getCell(2).value && !row.getCell(3).value;
      if (isEmpty) return;
        let rowData;      if (hasLocationColumn) {
        // 새로운 형식 (위치, 구매처, 구매금액, 설비군, Board명, S/N, 작업자 포함): 번호, 품명, 위치, 구분, 수량, 날짜, 카테고리, 구매처, 구매금액, 설비군, Board명, S/N, 작업자, 비고
        rowData = {
          id: row.getCell(1).value || rowNumber - 1,  // 번호
          excelRowNumber: rowNumber,                  // Excel 행 번호 추가
          name: row.getCell(2).value || '',           // 품명
          location: row.getCell(3).value || '',       // 위치
          type: row.getCell(4).value || '',           // 구분 (입고/출고)
          amount: row.getCell(5).value || 0,          // 수량
          date: row.getCell(6).value || '',           // 날짜
          category: row.getCell(7).value || '기타',   // 카테고리
          company: row.getCell(8).value || '',        // 구매처
          purchasePrice: row.getCell(9).value || '',  // 구매금액
          equipmentGroup: row.getCell(10).value || '', // 설비군
          boardName: row.getCell(11).value || '',     // Board명
          serialNumber: row.getCell(12).value || '',  // S/N
          operator: row.getCell(13).value || ''       // 작업자
        };
      } else {
        // 기존 형식 (위치 없음): 번호, 품명, 구분, 수량, 날짜, 카테고리, 비고
        rowData = {
          id: row.getCell(1).value || rowNumber - 1,  // 번호
          excelRowNumber: rowNumber,                  // Excel 행 번호 추가
          name: row.getCell(2).value || '',           // 품명
          location: '-',                              // 위치 정보 없음을 명시적으로 표시
          type: row.getCell(3).value || '',           // 구분 (입고/출고)
          amount: row.getCell(4).value || 0,          // 수량
          date: row.getCell(5).value || '',           // 날짜
          category: row.getCell(6).value || '기타',   // 카테고리
          company: '',                                // 구매처 없음
          purchasePrice: '',                          // 구매금액 없음
          equipmentGroup: '',                         // 설비군 없음
          boardName: '',                              // Board명 없음
          serialNumber: '',                           // S/N 없음
          operator: ''                                // 작업자 없음
        };
      }
      
      console.log(`Row ${rowNumber} data:`, rowData);
      historyData.push(rowData);
    });

    // 최신 이력이 맨 위에 오도록 정렬
    historyData.reverse();

    return { success: true, data: historyData };
  } catch (error) {
    console.error('입출고 이력 조회 중 오류 발생:', error);
    return { success: false, error: error.message };
  }
});

// IP 주소 가져오기 핸들러
ipcMain.handle('get-client-ip', async () => {
  try {
    // 로컬 IP 주소 가져오기
    const networkInterfaces = os.networkInterfaces();
    let ipAddress = 'localhost';
    
    // 모든 네트워크 인터페이스에서 IPv4 주소 찾기
    for (const interfaceName in networkInterfaces) {
      const interfaces = networkInterfaces[interfaceName];
      for (const iface of interfaces) {
        // IPv4 주소이면서 내부 네트워크 주소인 경우 (localhost가 아닌)
        if (iface.family === 'IPv4' && !iface.internal) {
          ipAddress = iface.address;
          break;
        }
      }
      if (ipAddress !== 'localhost') break;
    }
    
    return ipAddress;
  } catch (error) {
    console.error('IP 주소를 가져오는 중 오류:', error);
    return 'IP 확인 불가';
  }
});

// 채팅 메시지 저장 핸들러 - 최소 버전
ipcMain.handle('save-chat-message', async (event, chatData) => {
  try {
    console.log('채팅 메시지 저장 시작');
    
    // 최신 파일 다운로드
    await downloadFileFromFTP();
      
    // 1. 독립적인 새로운 워크북 생성
    const workbook = new ExcelJS.Workbook();
    
    // 2. 기존 엑셀 파일 읽기
    await workbook.xlsx.readFile(localTempFile);
    console.log('엑셀 파일 읽기 성공');
    
    // 3. 세 번째 시트 가져오기 (인덱스 2)
    let chatSheet;
    if (workbook.worksheets.length >= 3) {
      chatSheet = workbook.worksheets[2];
      console.log('세 번째 시트 사용:', chatSheet.name);
    } else {
      // 시트가 없으면 생성
      chatSheet = workbook.addWorksheet('채팅이력');
      console.log('새 채팅이력 시트 생성됨');
      
      // 첫 번째 행에 헤더 추가
      chatSheet.addRow(['메시지', '날짜 및 시간', 'IP 주소']);
    }
    
    // 4. 데이터를 새 행에 추가 (배열로)
    chatSheet.addRow([
      chatData.message, 
      chatData.dateTime, 
      chatData.ipAddress || 'IP 없음'
    ]);
    console.log('새 행 추가됨');
    
    // 5. 파일 저장
    await workbook.xlsx.writeFile(localTempFile);
    console.log('파일 저장 완료');
    
    // FTP 서버에 업로드
    await uploadFileToFTP();
    console.log('FTP 서버에 파일 업로드 성공');
    
    return { success: true, message: '채팅 메시지가 저장되었습니다.' };
  } catch (error) {
    console.error('채팅 메시지 저장 중 오류:', error);
    
    // 사용자가 엑셀 파일을 열어두었을 때 발생하는 오류 메시지
    if (error.message.includes('EBUSY') || 
        error.message.includes('process') || 
        error.message.includes('being used')) {
      return { 
        success: false, 
        error: '엑셀 파일이 다른 프로그램에서 열려있습니다. 파일을 닫고 다시 시도해주세요.' 
      };
    }
    
    return { success: false, error: error.message };
  }
});

// 채팅 히스토리 가져오기 핸들러
ipcMain.handle('get-chat-history', async () => {
  try {
    console.log('채팅 히스토리 가져오기 시작');
    
    // 최신 파일 다운로드
    await downloadFileFromFTP();
      
    // 엑셀 파일 존재 확인
    if (!fs.existsSync(localTempFile)) {
      console.error('엑셀 파일이 존재하지 않습니다:', localTempFile);
      return [];
    }
    
    // 엑셀 파일 읽기
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(localTempFile);
    
    // 세 번째 시트(채팅 이력) 가져오기
    let chatSheet;
    if (workbook.worksheets.length >= 3) {
      chatSheet = workbook.worksheets[2]; // 세 번째 시트
      console.log('세 번째 시트 사용:', chatSheet.name);
    } else {
      console.log('채팅 이력 시트가 없습니다.');
      return [];
    }
    
    // 데이터 추출
    const chatData = [];
    
    // 첫 번째 행은 헤더로 간주
    let headerRow = true;
    
    // 각 행 순회
    chatSheet.eachRow((row, rowNumber) => {
      // 헤더 행 건너뛰기
      if (headerRow) {
        headerRow = false;
        return;
      }
      
      // 행 데이터 추출
      const message = row.getCell(1).value || '';
      const dateTime = row.getCell(2).value || '';
      const ipAddress = row.getCell(3).value || '';
      
      chatData.push({
        message: message.toString(),
        dateTime: dateTime.toString(),
        ipAddress: ipAddress.toString()
      });
    });
    
    console.log(`채팅 이력 ${chatData.length}개를 불러왔습니다.`);
    return chatData;
    
  } catch (error) {
    console.error('채팅 히스토리 가져오기 중 오류:', error);
    throw error;
  }
});

// 입출고 이력 추가 API (위치 정보 포함)
ipcMain.handle('add-history-data', async (event, data) => {
  try {
    console.log('입출고 이력 추가 요청:', JSON.stringify(data, null, 2));
    console.log('구매처:', data.company);
    console.log('구매금액:', data.purchasePrice);
    
    // 최신 파일 다운로드
    await downloadFileFromFTP();
    
    // 엑셀 파일 로드
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(localTempFile);
    
    // 첫 번째 시트 (재고 데이터)에서 위치 정보 찾기
    const stockSheet = workbook.getWorksheet(1);
    let partLocation = data.location || ''; // 직접 입력된 위치가 있으면 사용
    let partCategory = data.category || '기타';
    
    console.log('입력된 위치 정보:', partLocation);
    
    // 품명으로 재고 시트에서 위치 정보 찾기
    if (!partLocation && data.name) {
      console.log('품명으로 위치 정보 찾는 중:', data.name);
      stockSheet.eachRow((row, rowNumber) => {
        if (rowNumber === 1) return; // 헤더 행 건너뛰기
        
        const nameCell = row.getCell(2);
        const name = nameCell.value;
        
        if (name && name.toString().toLowerCase() === data.name.toString().toLowerCase()) {
          partLocation = row.getCell(1).value || ''; // 위치 정보
          console.log('찾은 위치 정보:', partLocation);
          if (!data.category) {
            partCategory = row.getCell(4).value || '기타'; // 카테고리
          }
        }
      });
    }
    
    console.log('최종 사용할 위치 정보:', partLocation);
    
    // 두 번째 시트 (입출고 이력) 가져오기 또는 생성
    let historySheet = workbook.getWorksheet(2);    if (!historySheet) {
      console.log('입출고 이력 시트 생성');
      historySheet = workbook.addWorksheet('입출고이력');
        // 헤더 추가 (위치, 구매처, 구매금액, 설비군, Board명, S/N, 작업자 포함)
      const headerRow = historySheet.addRow([
        '번호', '품명', '위치', '구분', '수량', '날짜', '카테고리', '구매처', '구매금액', '설비군', 'Board명', 'S/N', '작업자', '비고'
      ]);
      
      headerRow.font = { bold: true };
      headerRow.commit();
    } else {
      // 기존 시트가 있는 경우 헤더 확인 및 업데이트
      const headerRow = historySheet.getRow(1);
      const currentHeader3 = headerRow.getCell(3).value;
      const currentHeader10 = headerRow.getCell(10).value;
      
      console.log('기존 헤더 3번째 컬럼:', currentHeader3);
      console.log('기존 헤더 10번째 컬럼:', currentHeader10);
        // 위치 컬럼이 없거나 설비군 컬럼이 없는 경우 헤더 업데이트
      if (currentHeader3 !== '위치' || currentHeader10 !== '설비군') {
        console.log('헤더를 위치, 구매처, 구매금액, 설비군, Board명, S/N, 작업자 포함 형식으로 업데이트');
        headerRow.getCell(1).value = '번호';
        headerRow.getCell(2).value = '품명';
        headerRow.getCell(3).value = '위치';
        headerRow.getCell(4).value = '구분';
        headerRow.getCell(5).value = '수량';
        headerRow.getCell(6).value = '날짜';
        headerRow.getCell(7).value = '카테고리';
        headerRow.getCell(8).value = '구매처';
        headerRow.getCell(9).value = '구매금액';
        headerRow.getCell(10).value = '설비군';
        headerRow.getCell(11).value = 'Board명';
        headerRow.getCell(12).value = 'S/N';
        headerRow.getCell(13).value = '작업자';
        headerRow.getCell(14).value = '비고';
        headerRow.font = { bold: true };
        headerRow.commit();
      }
    }
      // 새 행 번호 계산
    const newRowNumber = historySheet.rowCount;
      // 이력 데이터 준비 (위치, 구매처, 구매금액, 설비군, Board명, S/N, 작업자 포함)
    const historyRowData = [
      newRowNumber,                    // 번호
      data.name || '',                 // 품명
      partLocation || '',              // 위치
      data.type || '',                 // 구분 (입고/출고)
      parseInt(data.amount) || 0,      // 수량
      data.date || '',                 // 날짜
      partCategory,                    // 카테고리
      data.company || '',              // 구매처
      data.purchasePrice || '',        // 구매금액
      data.equipmentGroup || '',       // 설비군
      data.boardName || '',            // Board명
      data.serialNumber || '',         // S/N
      data.operator || '',             // 작업자
      ''                              // 비고
    ];
    
    console.log('추가할 이력 데이터 (위치, 설비군, Board명, S/N, 작업자 포함):', historyRowData);
    
    // 행 추가
    const newHistoryRow = historySheet.addRow(historyRowData);
    console.log('이력 행 추가 완료, 행 번호:', newHistoryRow.number);
    
    // 추가된 데이터 검증
    console.log('추가된 행의 각 셀 값:');
    newHistoryRow.eachCell((cell, colNumber) => {
      console.log(`  열 ${colNumber}: "${cell.value}"`);
    });
    
    // 파일 저장
    await workbook.xlsx.writeFile(localTempFile);
    console.log('엑셀 파일 저장 성공');
    
    // FTP 서버에 업로드
    await uploadFileToFTP();
    console.log('FTP 서버에 파일 업로드 성공');
    
    return { success: true, message: '입출고 이력이 성공적으로 추가되었습니다.' };
    
  } catch (error) {
    console.error('입출고 이력 추가 중 오류 발생:', error);
    return { success: false, error: error.message };
  }
});

// 입출고 이력 수정 API
ipcMain.handle('update-history-data', async (event, data) => {
  try {
    console.log('입출고 이력 수정 요청:', data);

    // 최신 파일 다운로드
    await downloadFileFromFTP();

    // 엑셀 파일 로드
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(localTempFile);

    // 두 번째 시트 (입출고 이력)
    const historySheet = workbook.getWorksheet(2);
    if (!historySheet) {
      return { success: false, error: '입출고 이력 시트를 찾을 수 없습니다.' };
    }

    // 첫 번째 시트 (재고 관리)
    const stockSheet = workbook.getWorksheet(1);
    if (!stockSheet) {
      return { success: false, error: '재고 시트를 찾을 수 없습니다.' };
    }

    // rowIndex는 Excel 행 번호
    const excelRowNumber = data.rowIndex;
    const row = historySheet.getRow(excelRowNumber);

    if (!row) {
      return { success: false, error: '해당 행을 찾을 수 없습니다.' };
    }

    // 기존 데이터 저장 (재고 계산용)
    const oldName = row.getCell(2).value || '';
    const oldType = row.getCell(4).value || '';
    const oldAmount = parseInt(row.getCell(5).value) || 0;

    // 새 데이터
    const newName = data.name || '';
    const newType = data.type || '';
    const newAmount = parseInt(data.amount) || 0;

    // 재고 업데이트 로직
    // 1. 기존 입출고 기록을 취소 (역산)
    if (oldName) {
      let oldStockRow = null;
      let oldStockRowNumber = 0;

      stockSheet.eachRow((stockRow, rowNum) => {
        if (rowNum === 1) return; // 헤더 건너뛰기
        const stockName = stockRow.getCell(2).value || '';
        if (stockName === oldName) {
          oldStockRow = stockRow;
          oldStockRowNumber = rowNum;
        }
      });

      if (oldStockRow) {
        const currentStock = parseInt(oldStockRow.getCell(7).value) || 0;
        // 기존 기록 취소 (입고였으면 빼고, 출고였으면 더함)
        if (oldType === '입고') {
          oldStockRow.getCell(7).value = currentStock - oldAmount;
        } else if (oldType === '출고') {
          oldStockRow.getCell(7).value = currentStock + oldAmount;
        }
      }
    }

    // 2. 새 입출고 기록 적용
    if (newName) {
      let newStockRow = null;
      let newStockRowNumber = 0;

      stockSheet.eachRow((stockRow, rowNum) => {
        if (rowNum === 1) return; // 헤더 건너뛰기
        const stockName = stockRow.getCell(2).value || '';
        if (stockName === newName) {
          newStockRow = stockRow;
          newStockRowNumber = rowNum;
        }
      });

      if (newStockRow) {
        const currentStock = parseInt(newStockRow.getCell(7).value) || 0;
        // 새 기록 적용
        if (newType === '입고') {
          newStockRow.getCell(7).value = currentStock + newAmount;
        } else if (newType === '출고') {
          newStockRow.getCell(7).value = currentStock - newAmount;
        }
      }
    }

    // 데이터 업데이트 (열 순서: 번호, 품명, 위치, 구분, 수량, 날짜, 카테고리, 구매처, 구매금액, 설비군, Board명, S/N, 작업자, 비고)
    // 번호는 유지 (getCell(1))
    row.getCell(2).value = newName;
    row.getCell(3).value = data.location || '';
    row.getCell(4).value = newType;
    row.getCell(5).value = newAmount;
    row.getCell(6).value = data.date || '';
    row.getCell(7).value = data.category || '';
    row.getCell(8).value = data.company || '';
    row.getCell(9).value = parseFloat(data.purchasePrice) || 0;
    row.getCell(10).value = data.equipmentGroup || '';
    row.getCell(11).value = data.boardName || '';
    row.getCell(12).value = data.serialNumber || '';
    row.getCell(13).value = data.operator || '';

    row.commit();

    // 파일 저장
    await workbook.xlsx.writeFile(localTempFile);
    await uploadFileToFTP();

    console.log('입출고 이력 및 재고 수정 완료:', excelRowNumber);
    return { success: true, message: '입출고 이력 및 재고가 수정되었습니다.' };

  } catch (error) {
    console.error('입출고 이력 수정 중 오류 발생:', error);
    return { success: false, error: error.message };
  }
});

// 입출고 이력 삭제 API
ipcMain.handle('delete-history-data', async (event, rowIndex) => {
  try {
    console.log('입출고 이력 삭제 요청:', rowIndex);

    // 최신 파일 다운로드
    await downloadFileFromFTP();

    // 엑셀 파일 로드
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(localTempFile);

    // 두 번째 시트 (입출고 이력)
    const historySheet = workbook.getWorksheet(2);
    if (!historySheet) {
      return { success: false, error: '입출고 이력 시트를 찾을 수 없습니다.' };
    }

    // 첫 번째 시트 (재고 관리)
    const stockSheet = workbook.getWorksheet(1);
    if (!stockSheet) {
      return { success: false, error: '재고 시트를 찾을 수 없습니다.' };
    }

    // rowIndex는 Excel 행 번호
    const excelRowNumber = rowIndex;
    const row = historySheet.getRow(excelRowNumber);

    if (!row) {
      return { success: false, error: '해당 행을 찾을 수 없습니다.' };
    }

    // 삭제할 데이터 저장 (재고 계산용)
    const deleteName = row.getCell(2).value || '';
    const deleteType = row.getCell(4).value || '';
    const deleteAmount = parseInt(row.getCell(5).value) || 0;

    // 재고 업데이트 (삭제 기록 취소)
    if (deleteName) {
      let stockRow = null;

      stockSheet.eachRow((sRow, rowNum) => {
        if (rowNum === 1) return; // 헤더 건너뛰기
        const stockName = sRow.getCell(2).value || '';
        if (stockName === deleteName) {
          stockRow = sRow;
        }
      });

      if (stockRow) {
        const currentStock = parseInt(stockRow.getCell(7).value) || 0;
        // 삭제된 기록 취소 (입고 삭제하면 재고 감소, 출고 삭제하면 재고 증가)
        if (deleteType === '입고') {
          stockRow.getCell(7).value = currentStock - deleteAmount;
        } else if (deleteType === '출고') {
          stockRow.getCell(7).value = currentStock + deleteAmount;
        }
      }
    }

    // 행 삭제
    historySheet.spliceRows(excelRowNumber, 1);

    // 번호 재정렬
    let rowNum = 1;
    historySheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return; // 헤더 건너뛰기
      row.getCell(1).value = rowNum++;
    });

    // 파일 저장
    await workbook.xlsx.writeFile(localTempFile);
    await uploadFileToFTP();

    console.log('입출고 이력 및 재고 삭제 완료:', excelRowNumber);
    return { success: true, message: '입출고 이력이 삭제되고 재고가 업데이트되었습니다.' };

  } catch (error) {
    console.error('입출고 이력 삭제 중 오류 발생:', error);
    return { success: false, error: error.message };
  }
});

// 재고 수정 핸들러
ipcMain.handle('update-stock-item', async (event, data) => {
  try {
    // 최신 파일 다운로드
    await downloadFileFromFTP();

    // 엑셀 파일 로드
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(localTempFile);

    const stockSheet = workbook.getWorksheet(1); // 재고 시트
    const historySheet = workbook.getWorksheet(2); // 입출고 이력 시트

    if (!stockSheet) {
      return { success: false, error: '재고 시트를 찾을 수 없습니다.' };
    }

    // 재고 시트에서 해당 행 업데이트
    const stockRow = stockSheet.getRow(data.rowNumber);
    stockRow.getCell(1).value = data.location;
    stockRow.getCell(2).value = data.name;
    stockRow.getCell(3).value = data.company;
    stockRow.getCell(4).value = data.category;
    stockRow.getCell(5).value = data.minStock;
    stockRow.getCell(6).value = data.inDate;
    stockRow.getCell(7).value = data.currentStock;
    stockRow.getCell(8).value = data.purchasePrice;
    stockRow.getCell(9).value = data.equipmentGroup;
    stockRow.getCell(10).value = data.boardName;
    stockRow.getCell(11).value = data.serialNumber;
    stockRow.getCell(12).value = data.operator;
    stockRow.getCell(13).value = data.unitPrice;
    stockRow.commit();

    // 품명이 변경된 경우, 입출고 이력의 모든 관련 레코드 품명 업데이트
    if (historySheet && data.oldName !== data.name) {
      historySheet.eachRow((row, rowNumber) => {
        if (rowNumber === 1) return; // 헤더 건너뛰기

        const historyName = row.getCell(2).value;
        if (historyName === data.oldName) {
          row.getCell(2).value = data.name;
          row.commit();
        }
      });
    }

    // 엑셀 파일 저장
    await workbook.xlsx.writeFile(localTempFile);

    // FTP 업로드
    await uploadFileToFTP();

    return { success: true, message: '재고가 성공적으로 수정되었습니다.' };
  } catch (error) {
    console.error('재고 수정 중 오류:', error);
    return { success: false, error: error.message };
  }
});

// 재고 삭제 핸들러
ipcMain.handle('delete-stock-item', async (event, rowNumber, name) => {
  try {
    // 최신 파일 다운로드
    await downloadFileFromFTP();

    // 엑셀 파일 로드
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(localTempFile);

    const stockSheet = workbook.getWorksheet(1); // 재고 시트
    const historySheet = workbook.getWorksheet(2); // 입출고 이력 시트

    if (!stockSheet) {
      return { success: false, error: '재고 시트를 찾을 수 없습니다.' };
    }

    // 재고 시트에서 해당 행 삭제
    stockSheet.spliceRows(rowNumber, 1);

    // 입출고 이력에서 해당 품명의 모든 레코드 삭제
    if (historySheet) {
      const rowsToDelete = [];

      historySheet.eachRow((row, rowNumber) => {
        if (rowNumber === 1) return; // 헤더 건너뛰기

        const historyName = row.getCell(2).value;
        if (historyName === name) {
          rowsToDelete.push(rowNumber);
        }
      });

      // 역순으로 삭제 (인덱스 변경 방지)
      for (let i = rowsToDelete.length - 1; i >= 0; i--) {
        historySheet.spliceRows(rowsToDelete[i], 1);
      }
    }

    // 엑셀 파일 저장
    await workbook.xlsx.writeFile(localTempFile);

    // FTP 업로드
    await uploadFileToFTP();

    return { success: true, message: '재고 및 관련 이력이 성공적으로 삭제되었습니다.' };
  } catch (error) {
    console.error('재고 삭제 중 오류:', error);
    return { success: false, error: error.message };
  }
});

// 로그인 관련 핸들러들
ipcMain.handle('register-admin', async (event, userData) => {
  return await authService.registerAdmin(userData);
});

ipcMain.handle('user-login', async (event, loginData) => {
  return await authService.login(loginData);
});

// 비밀번호 검증 핸들러
ipcMain.handle('verify-password', async (event, password) => {
  return await authService.verifyPassword(password);
});

ipcMain.handle('user-logout', async (event) => {
  return await authService.logout();
});

ipcMain.handle('check-auth-status', async (event) => {
  return authService.checkAuthStatus();
});

ipcMain.handle('find-password', async (event, findData) => {
  return await authService.findPassword(findData);
});

ipcMain.handle('change-password', async (event, changeData) => {
  return await authService.changePassword(changeData);
});

// 권한 확인 미들웨어 함수
function requireAuth(handler) {
  return async (event, ...args) => {
    try {
      authService.requireAuth();
      return await handler(event, ...args);
    } catch (error) {
      return { success: false, error: error.message };
    }
  };
}