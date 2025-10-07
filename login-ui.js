// login-ui.js - 로그인 UI 스크립트

// 페이지 로드 시 모든 이벤트 리스너 등록
window.addEventListener('DOMContentLoaded', async () => {
    console.log('DOM 로드 완료');
    
    // 요소들 확인
    const loginForm = document.getElementById('loginForm');
    const adminSetupButton = document.getElementById('adminSetupButton');
    const usernameInput = document.getElementById('username');
    const passwordInput = document.getElementById('password');
    const adminModal = document.getElementById('adminModal');
    const adminForm = document.getElementById('adminForm');
    const adminCancelButton = document.getElementById('adminCancelButton');
    const forgotPasswordButton = document.getElementById('forgotPasswordButton');
    const forgotPasswordModal = document.getElementById('forgotPasswordModal');
    const forgotPasswordForm = document.getElementById('forgotPasswordForm');
    const forgotPasswordCancelButton = document.getElementById('forgotPasswordCancelButton');
    const changePasswordModal = document.getElementById('changePasswordModal');
    const changePasswordForm = document.getElementById('changePasswordForm');
    const changePasswordCancelButton = document.getElementById('changePasswordCancelButton');
    
    console.log('loginForm:', loginForm);
    console.log('adminSetupButton:', adminSetupButton);
    console.log('adminModal:', adminModal);
    
    // 로그인 폼 이벤트 리스너
    if (loginForm) {
        loginForm.addEventListener('submit', async (e) => {
            e.preventDefault();
            
            const employeeId = document.getElementById('username').value.trim();
            const password = document.getElementById('password').value;
            const errorDiv = document.getElementById('errorMessage');
            const loginButton = document.getElementById('loginButton');
            
            // 입력 값 검증
            if (!employeeId || !password) {
                errorDiv.textContent = '사번과 비밀번호를 모두 입력해주세요.';
                return;
            }
            
            // 로딩 상태
            loginButton.disabled = true;
            loginButton.textContent = '로그인 중...';
            errorDiv.textContent = '';
            document.querySelector('.login-container').classList.add('loading');
            
            try {
                const result = await window.partsAPI.userLogin({ 
                    username: employeeId,  // 사번을 username으로 전달
                    password 
                });
                
                if (result.success) {
                    // 로그인 성공 시 메인 화면으로 이동
                    window.location.href = '재고목록.html';
                } else {
                    errorDiv.textContent = result.error;
                }
            } catch (error) {
                errorDiv.textContent = '로그인 중 오류가 발생했습니다.';
                console.error('로그인 오류:', error);
            } finally {
                loginButton.disabled = false;
                loginButton.textContent = '로그인';
                document.querySelector('.login-container').classList.remove('loading');
            }
        });
    }

    // 관리자 계정 생성 버튼 클릭 - 모달 열기
    if (adminSetupButton) {
        console.log('관리자 계정 생성 버튼 이벤트 리스너 등록');
        adminSetupButton.addEventListener('click', () => {
            console.log('관리자 계정 생성 버튼 클릭됨 - 모달 열기');
            adminModal.style.display = 'block';
            document.getElementById('adminEmployeeId').focus();
        });
    } else {
        console.error('adminSetupButton을 찾을 수 없습니다!');
    }

    // 모달 취소 버튼
    if (adminCancelButton) {
        adminCancelButton.addEventListener('click', () => {
            closeAdminModal();
        });
    }

    // 모달 배경 클릭 시 닫기
    if (adminModal) {
        adminModal.addEventListener('click', (e) => {
            if (e.target === adminModal) {
                closeAdminModal();
            }
        });
    }

    // 관리자 계정 생성 폼 제출
    if (adminForm) {
        adminForm.addEventListener('submit', async (e) => {
            e.preventDefault();
            await handleAdminCreate();
        });
    }

    // 비밀번호 찾기 버튼 클릭
    if (forgotPasswordButton) {
        forgotPasswordButton.addEventListener('click', () => {
            console.log('비밀번호 찾기 버튼 클릭됨');
            forgotPasswordModal.style.display = 'block';
            document.getElementById('forgotEmployeeId').focus();
        });
    }

    // 비밀번호 찾기 모달 취소 버튼
    if (forgotPasswordCancelButton) {
        forgotPasswordCancelButton.addEventListener('click', () => {
            closeForgotPasswordModal();
        });
    }

    // 비밀번호 찾기 모달 배경 클릭 시 닫기
    if (forgotPasswordModal) {
        forgotPasswordModal.addEventListener('click', (e) => {
            if (e.target === forgotPasswordModal) {
                closeForgotPasswordModal();
            }
        });
    }

    // 비밀번호 찾기 폼 제출
    if (forgotPasswordForm) {
        forgotPasswordForm.addEventListener('submit', async (e) => {
            e.preventDefault();
            await handleForgotPassword();
        });
    }

    // 비밀번호 변경 모달 취소 버튼
    if (changePasswordCancelButton) {
        changePasswordCancelButton.addEventListener('click', () => {
            closeChangePasswordModal();
        });
    }

    // 비밀번호 변경 모달 배경 클릭 시 닫기
    if (changePasswordModal) {
        changePasswordModal.addEventListener('click', (e) => {
            if (e.target === changePasswordModal) {
                closeChangePasswordModal();
            }
        });
    }

    // 비밀번호 변경 폼 제출
    if (changePasswordForm) {
        changePasswordForm.addEventListener('submit', async (e) => {
            e.preventDefault();
            await handleChangePassword();
        });
    }

    // Enter 키 이벤트 처리
    if (usernameInput) {
        usernameInput.addEventListener('keypress', (e) => {
            if (e.key === 'Enter') {
                passwordInput?.focus();
            }
        });
    }

    if (passwordInput) {
        passwordInput.addEventListener('keypress', (e) => {
            if (e.key === 'Enter') {
                loginForm?.dispatchEvent(new Event('submit'));
            }
        });
    }

    // 기존 세션 확인
    try {
        const authStatus = await window.partsAPI.checkAuthStatus();
        
        // 이미 로그인되어 있다면 메인 페이지로 리다이렉트
        if (authStatus.isAuthenticated) {
            window.location.href = '재고목록.html';
        }
    } catch (error) {
        console.error('인증 상태 확인 오류:', error);
    }
});

// 모달 닫기 함수
function closeAdminModal() {
    const adminModal = document.getElementById('adminModal');
    const adminForm = document.getElementById('adminForm');
    const adminErrorMessage = document.getElementById('adminErrorMessage');
    
    adminModal.style.display = 'none';
    adminForm.reset();
    adminErrorMessage.textContent = '';
}

// 비밀번호 찾기 모달 닫기 함수
function closeForgotPasswordModal() {
    const forgotPasswordModal = document.getElementById('forgotPasswordModal');
    const forgotPasswordForm = document.getElementById('forgotPasswordForm');
    const forgotPasswordResult = document.getElementById('forgotPasswordResult');
    
    forgotPasswordModal.style.display = 'none';
    forgotPasswordForm.reset();
    forgotPasswordResult.textContent = '';
}

// 비밀번호 변경 모달 닫기 함수
function closeChangePasswordModal() {
    const changePasswordModal = document.getElementById('changePasswordModal');
    const changePasswordForm = document.getElementById('changePasswordForm');
    const changePasswordError = document.getElementById('changePasswordError');
    
    changePasswordModal.style.display = 'none';
    changePasswordForm.reset();
    changePasswordError.textContent = '';
}

// 비밀번호 찾기 처리 함수
async function handleForgotPassword() {
    console.log('비밀번호 찾기 처리 시작');
    
    const employeeId = document.getElementById('forgotEmployeeId').value.trim();
    const recoveryCode = document.getElementById('forgotRecoveryCode').value.trim();
    const resultDiv = document.getElementById('forgotPasswordResult');
    const findButton = document.getElementById('forgotPasswordFindButton');
    
    // 입력 값 검증
    if (!employeeId) {
        resultDiv.textContent = '사번을 입력해주세요.';
        resultDiv.style.color = '#e74c3c';
        return;
    }
    
    if (!recoveryCode) {
        resultDiv.textContent = 'PW찾기코드를 입력해주세요.';
        resultDiv.style.color = '#e74c3c';
        return;
    }
    
    // 로딩 상태
    findButton.disabled = true;
    findButton.textContent = '찾는 중...';
    resultDiv.textContent = '';
    
    try {
        console.log('비밀번호 찾기 API 호출 시작');
        const result = await window.partsAPI.findPassword({
            username: employeeId,
            recoveryCode: recoveryCode
        });
        
        console.log('비밀번호 찾기 결과:', result);
        
        if (result.success) {
            // 현재 비밀번호 표시하고 변경 모달로 이동
            document.getElementById('currentPasswordDisplay').textContent = result.password;
            
            // 비밀번호 찾기 모달 닫기
            closeForgotPasswordModal();
            
            // 비밀번호 변경 모달 열기
            document.getElementById('changePasswordModal').style.display = 'block';
            document.getElementById('newPassword').focus();
            
            // 변경 모달에 사용자 정보 저장
            window.currentUserForChange = {
                username: employeeId,
                recoveryCode: recoveryCode
            };
        } else {
            resultDiv.textContent = result.error;
            resultDiv.style.color = '#e74c3c';
        }
    } catch (error) {
        console.error('비밀번호 찾기 오류:', error);
        resultDiv.textContent = '비밀번호 찾기 중 오류가 발생했습니다.';
        resultDiv.style.color = '#e74c3c';
    } finally {
        findButton.disabled = false;
        findButton.textContent = '비밀번호 찾기';
    }
}

// 비밀번호 변경 처리 함수
async function handleChangePassword() {
    console.log('비밀번호 변경 처리 시작');
    
    const newPassword = document.getElementById('newPassword').value;
    const confirmNewPassword = document.getElementById('confirmNewPassword').value;
    const errorDiv = document.getElementById('changePasswordError');
    const saveButton = document.getElementById('changePasswordSaveButton');
    
    // 입력 값 검증
    if (!newPassword || newPassword.length < 4) {
        errorDiv.textContent = '새 비밀번호는 최소 4자 이상이어야 합니다.';
        return;
    }
    
    if (newPassword !== confirmNewPassword) {
        errorDiv.textContent = '새 비밀번호가 일치하지 않습니다.';
        return;
    }
    
    // 현재 비밀번호와 같은지 확인
    const currentPassword = document.getElementById('currentPasswordDisplay').textContent;
    if (newPassword === currentPassword) {
        errorDiv.textContent = '현재 비밀번호와 다른 비밀번호를 입력해주세요.';
        return;
    }
    
    // 로딩 상태
    saveButton.disabled = true;
    saveButton.textContent = '변경 중...';
    errorDiv.textContent = '';
    
    try {
        console.log('비밀번호 변경 API 호출 시작');
        const result = await window.partsAPI.changePassword({
            username: window.currentUserForChange.username,
            recoveryCode: window.currentUserForChange.recoveryCode,
            newPassword: newPassword
        });
        
        console.log('비밀번호 변경 결과:', result);
        
        if (result.success) {
            alert('비밀번호가 성공적으로 변경되었습니다.\n새 비밀번호로 로그인해주세요.');
            
            // 비밀번호 변경 모달 닫기
            closeChangePasswordModal();
            
            // 로그인 폼에 사번 자동 입력
            document.getElementById('username').value = window.currentUserForChange.username;
            document.getElementById('password').focus();
            
            // 임시 저장된 사용자 정보 제거
            window.currentUserForChange = null;
        } else {
            errorDiv.textContent = result.error;
        }
    } catch (error) {
        console.error('비밀번호 변경 오류:', error);
        errorDiv.textContent = '비밀번호 변경 중 오류가 발생했습니다.';
    } finally {
        saveButton.disabled = false;
        saveButton.textContent = '비밀번호 변경';
    }
}

// 관리자 계정 생성 처리 함수
async function handleAdminCreate() {
    console.log('관리자 계정 생성 처리 시작');
    
    const employeeId = document.getElementById('adminEmployeeId').value.trim();
    const password = document.getElementById('adminPassword').value;
    const confirmPassword = document.getElementById('adminConfirmPassword').value;
    const recoveryCode = document.getElementById('adminRecoveryCode').value.trim();
    const errorDiv = document.getElementById('adminErrorMessage');
    const createButton = document.getElementById('adminCreateButton');
    
    // 입력 값 검증
    if (!employeeId) {
        errorDiv.textContent = '사번을 입력해주세요.';
        return;
    }
    
    if (!password || password.length < 4) {
        errorDiv.textContent = '비밀번호는 최소 4자 이상이어야 합니다.';
        return;
    }
    
    if (password !== confirmPassword) {
        errorDiv.textContent = '비밀번호가 일치하지 않습니다.';
        return;
    }
    
    if (!recoveryCode) {
        errorDiv.textContent = 'PW찾기코드를 입력해주세요.';
        return;
    }
    
    // 로딩 상태
    createButton.disabled = true;
    createButton.textContent = '생성 중...';
    errorDiv.textContent = '';
    
    try {
        console.log('관리자 계정 생성 API 호출 시작');
        const result = await window.partsAPI.registerAdmin({
            username: employeeId,  // 사번을 username으로 전달
            password,
            recoveryCode,          // PW찾기코드 추가
            role: 'admin'
        });
        
        console.log('관리자 계정 생성 결과:', result);
        
        if (result.success) {
            alert('관리자 계정이 성공적으로 생성되었습니다.\n로그인해주세요.');
            
            // 모달 닫기
            closeAdminModal();
            
            // 입력 필드에 자동으로 사번 채우기
            document.getElementById('username').value = employeeId;
            document.getElementById('password').focus();
        } else {
            errorDiv.textContent = result.error;
        }
    } catch (error) {
        console.error('계정 생성 오류:', error);
        errorDiv.textContent = '계정 생성 중 오류가 발생했습니다.';
    } finally {
        createButton.disabled = false;
        createButton.textContent = '생성';
    }
}
