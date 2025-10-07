// preload.js
const { contextBridge, ipcRenderer } = require('electron');

// preload.js 파일에 추가
contextBridge.exposeInMainWorld('partsAPI', {
    getDbData: () => ipcRenderer.invoke('get-db-data'),
    registerPart: (data) => ipcRenderer.invoke('register-part', data),
    checkDuplicatePart: (name) => ipcRenderer.invoke('check-duplicate-part', name),    getExcelData: () => ipcRenderer.invoke('get-excel-data'),
    updateStock: (data) => ipcRenderer.invoke('update-stock', data),
    getHistoryData: () => ipcRenderer.invoke('get-history-data'),
    addHistoryData: (data) => ipcRenderer.invoke('add-history-data', data), // 새로 추가

    // 채팅 메시지 저장 함수 추가
    saveChatMessage: async (chatData) => {
        try {
            const result = await ipcRenderer.invoke('save-chat-message', chatData);
            return { success: true, data: result };
        } catch (error) {
            return { success: false, error: error.message };
        }
    },

    // IP 주소 가져오기 함수
    getClientIP: async () => {
        try {
            const result = await ipcRenderer.invoke('get-client-ip');
            return { success: true, data: result };
        } catch (error) {
            return { success: false, error: error.message };
        }
    },    // 채팅 히스토리 가져오기 API
    getChatHistory: async () => {
        try {
            const result = await ipcRenderer.invoke('get-chat-history');
            return { success: true, data: result };
        } catch (error) {
            return { success: false, error: error.message };
        }
    },

    // 로그인 관련 API 추가
    registerAdmin: async (userData) => {
        try {
            const result = await ipcRenderer.invoke('register-admin', userData);
            return result;
        } catch (error) {
            return { success: false, error: error.message };
        }
    },

    userLogin: async (loginData) => {
        try {
            const result = await ipcRenderer.invoke('user-login', loginData);
            return result;
        } catch (error) {
            return { success: false, error: error.message };
        }
    },

    userLogout: async () => {
        try {
            const result = await ipcRenderer.invoke('user-logout');
            return result;
        } catch (error) {
            return { success: false, error: error.message };
        }
    },

    checkAuthStatus: async () => {
        try {
            const result = await ipcRenderer.invoke('check-auth-status');
            return result;
        } catch (error) {
            return { success: false, error: error.message };
        }
    },

    findPassword: async (findData) => {
        try {
            const result = await ipcRenderer.invoke('find-password', findData);
            return result;
        } catch (error) {
            return { success: false, error: error.message };
        }
    },

    changePassword: async (changeData) => {
        try {
            const result = await ipcRenderer.invoke('change-password', changeData);
            return result;
        } catch (error) {
            return { success: false, error: error.message };
        }
    }
});