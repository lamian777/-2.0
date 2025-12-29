const { app, BrowserWindow, dialog } = require('electron');
const path = require('path');
const fs = require('fs');
const { enable } = require('@electron/remote/main');

// 保持对window对象的全局引用，避免JavaScript对象被垃圾回收时，窗口被自动关闭
let mainWindow;

function createWindow() {
  // 创建浏览器窗口
  mainWindow = new BrowserWindow({
    width: 1000,
    height: 800,
    webPreferences: {
      nodeIntegration: true, // 允许在渲染进程中使用Node.js API
      contextIsolation: false, // 禁用上下文隔离
      enableRemoteModule: true // 启用remote模块
    },
    icon: path.join(__dirname, 'build/icon.png')
  });
  
  // 启用remote模块
  enable(mainWindow.webContents);

  // 加载应用的index.html
  mainWindow.loadFile('index.html');

  // 设置窗口标题
  mainWindow.setTitle('骑缝章工具');

  // 打开开发者工具（开发时使用，发布时可注释掉）
  // mainWindow.webContents.openDevTools();

  // 当window被关闭时，触发下面的事件
  mainWindow.on('closed', function () {
    mainWindow = null;
  });
}

// 初始化remote模块
require('@electron/remote/main').initialize();

// 当Electron完成初始化并准备创建浏览器窗口时，调用这个方法
app.whenReady().then(createWindow);

// 当所有窗口关闭时退出应用
app.on('window-all-closed', function () {
  // 在macOS上，用户通常希望点击dock图标重新打开应用窗口
  if (process.platform !== 'darwin') {
    app.quit();
  }
});

app.on('activate', function () {
  // 在macOS上，当dock图标被点击且没有其他窗口打开时，通常会重新创建一个窗口
  if (mainWindow === null) {
    createWindow();
  }
});

// 在这里可以添加其他的主进程代码