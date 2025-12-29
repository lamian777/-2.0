// 骑缝章工具 - 主脚本文件
// 功能：用于在PDF文档上添加骑缝章，支持印章上传、PDF上传、印章切割和PDF生成

// 全局变量
let stampDataUrl = null;     // 印章图片的DataURL
let pdfData = null;          // PDF文件的二进制数据
let pdfPageCount = 1;        // PDF文件的页数
let stampSlices = [];        // 印章切片的DataURL数组
let generatedPdfUrl = null;  // 生成的PDF文件的预览URL
let originalPdfFileName = null; // 原始PDF文件名

// 调试函数
/**
 * 输出调试日志
 * @param {string} message - 调试信息
 */
function debugLog(message) {
    console.log('骑缝章工具调试:', message);
}

// DOM元素
const stampFileInput = document.getElementById('stampFile');
const pdfFileInput = document.getElementById('pdfFile');
const generateBtn = document.getElementById('generateBtn');
const stampImage = document.getElementById('stampImage');
const stampPreviewContainer = document.getElementById('stampPreviewContainer');
const pdfPreviewContainer = document.getElementById('pdfPreviewContainer');
const pdfFrame = document.getElementById('pdfFrame');
const statusMessage = document.getElementById('statusMessage');
const resultPreviewContainer = document.getElementById('resultPreviewContainer');
const resultFrame = document.getElementById('resultFrame');
const stampPositionSelect = document.getElementById('stampPosition');
const downloadBtn = document.getElementById('downloadBtn');
const downloadContainer = document.getElementById('downloadContainer');
const progressContainer = document.getElementById('progressContainer');
const progressFill = document.getElementById('progressFill');
const progressText = document.getElementById('progressText');
const stampUploadHint = document.getElementById('stampUploadHint');
const pdfUploadHint = document.getElementById('pdfUploadHint');
const step1Status = document.getElementById('step1Status');
const step2Status = document.getElementById('step2Status');
const step3Status = document.getElementById('step3Status');
const step4Status = document.getElementById('step4Status');

/**
 * 初始化应用程序
 * 设置事件监听器和初始状态
 */
function init() {
    // 绑定事件监听器
    stampFileInput.addEventListener('change', handleStampUpload);
    pdfFileInput.addEventListener('change', handlePdfUpload);
    generateBtn.addEventListener('click', generateStampedPdf);
    downloadBtn.addEventListener('click', downloadPdf);
    
    // 初始化步骤状态
    updateStepStatus('step1', 'pending');
    updateStepStatus('step2', 'pending');
    updateStepStatus('step3', 'pending');
    updateStepStatus('step4', 'pending');
}

/**
 * 处理印章上传
 * @param {Event} e - 文件上传事件
 */
function handleStampUpload(e) {
    const file = e.target.files[0];
    if (!file) return;
    
    // 验证文件类型
    if (file.type !== 'image/png') {
        showStatus('请上传PNG格式的印章图片', 'error');
        updateStepStatus('step1', 'error');
        return;
    }
    
    // 更新状态并显示加载消息
    updateStepStatus('step1', 'in-progress');
    showStatus('正在加载印章图片...', 'loading');
    
    // 读取文件
    const reader = new FileReader();
    reader.onload = (event) => {
        stampDataUrl = event.target.result;
        
        // 显示预览
        stampImage.src = stampDataUrl;
        stampPreviewContainer.classList.remove('hidden');
        stampUploadHint.classList.add('hidden');
        
        // 更新状态
        showStatus('印章图片加载成功', 'success');
        updateStepStatus('step1', 'completed');
        
        // 检查是否可以切割印章
        checkGenerateButtonState();
    };
    
    reader.onerror = () => {
        showStatus('印章图片加载失败', 'error');
        updateStepStatus('step1', 'error');
    };
    
    reader.readAsDataURL(file);
}

/**
 * 处理PDF上传
 * @param {Event} e - 文件上传事件
 */
function handlePdfUpload(e) {
    const file = e.target.files[0];
    if (!file) return;
    
    // 保存原始文件名
    originalPdfFileName = file.name;
    
    // 验证文件类型
    if (file.type !== 'application/pdf') {
        showStatus('请上传PDF格式的文件', 'error');
        updateStepStatus('step2', 'error');
        return;
    }
    
    // 更新状态并显示加载消息
    updateStepStatus('step2', 'in-progress');
    showStatus('正在加载PDF文件...', 'loading');
    
    // 读取文件
    const reader = new FileReader();
    reader.onload = async (event) => {
        pdfData = new Uint8Array(event.target.result);
        
        try {
            // 预览PDF
            const pdfUrl = URL.createObjectURL(file);
            pdfFrame.src = pdfUrl;
            pdfPreviewContainer.classList.remove('hidden');
            pdfUploadHint.classList.add('hidden');
            
            // 获取PDF页数
            const pdf = await PDFLib.PDFDocument.load(pdfData);
            pdfPageCount = pdf.getPageCount();
            
            // 更新状态
            showStatus(`PDF文件加载成功，共${pdfPageCount}页`, 'success');
            updateStepStatus('step2', 'completed');
            
            // 如果已经上传了印章，自动切割印章
            if (stampDataUrl) {
                await sliceStamp();
            }
            
            // 检查是否可以生成PDF
            checkGenerateButtonState();
        } catch (error) {
            debugLog('PDF加载失败:', error);
            showStatus('PDF文件加载失败', 'error');
            updateStepStatus('step2', 'error');
        }
    };
    
    reader.onerror = () => {
        showStatus('PDF文件加载失败', 'error');
        updateStepStatus('step2', 'error');
    };
    
    reader.readAsArrayBuffer(file);
}

/**
 * 切割印章为多个切片
 * @returns {Promise<void>} - Promise对象
 */
async function sliceStamp() {
    // 检查必要条件
    if (!stampDataUrl || pdfPageCount < 1) {
        debugLog('切割印章失败：缺少印章图片或PDF页数信息');
        return;
    }
    
    // 更新状态并显示消息
    updateStepStatus('step3', 'in-progress');
    showStatus('正在切割印章...', 'loading');
    
    try {
        // 使用requestAnimationFrame来处理复杂的DOM操作，提高性能
        await new Promise(resolve => {
            requestAnimationFrame(async () => {
                // 创建或重用canvas元素
                let canvas = document.getElementById('cachedCanvas');
                let ctx = null;
                
                if (!canvas) {
                    canvas = document.createElement('canvas');
                    canvas.id = 'cachedCanvas';
                    canvas.style.display = 'none';
                    document.body.appendChild(canvas);
                }
                ctx = canvas.getContext('2d');
                
                // 加载印章图片
                const stampImageObj = new Image();
                stampImageObj.crossOrigin = 'anonymous';
                
                try {
                    await new Promise((imgResolve, imgReject) => {
                        stampImageObj.onload = imgResolve;
                        stampImageObj.onerror = imgReject;
                        stampImageObj.src = stampDataUrl;
                    });
                } catch (error) {
                    throw new Error('印章图片加载失败: ' + error.message);
                }
                
                // 计算每个切片的宽度
                const stampWidth = stampImageObj.width;
                const stampHeight = stampImageObj.height;
                const sliceWidth = Math.ceil(stampWidth / pdfPageCount);
                
                // 设置canvas大小
                canvas.width = stampWidth;
                canvas.height = stampHeight;
                
                // 绘制印章到canvas
                ctx.clearRect(0, 0, canvas.width, canvas.height);
                ctx.drawImage(stampImageObj, 0, 0);
                
                // 切割印章
                stampSlices = [];
                const slicePromises = [];
                
                // 创建切片画布
                let sliceCanvas = document.getElementById('cachedSliceCanvas');
                let sliceCtx = null;
                
                if (!sliceCanvas) {
                    sliceCanvas = document.createElement('canvas');
                    sliceCanvas.id = 'cachedSliceCanvas';
                    sliceCanvas.style.display = 'none';
                    document.body.appendChild(sliceCanvas);
                }
                sliceCanvas.width = sliceWidth;
                sliceCanvas.height = stampHeight;
                sliceCtx = sliceCanvas.getContext('2d');
                
                // 生成所有切片
                for (let i = 0; i < pdfPageCount; i++) {
                    slicePromises.push(new Promise((resolve) => {
                        // 计算切片区域
                        const x = i * sliceWidth;
                        const y = 0;
                        const width = sliceWidth;
                        const height = stampHeight;
                        
                        // 创建切片图片
                        sliceCtx.clearRect(0, 0, sliceCanvas.width, sliceCanvas.height);
                        sliceCtx.drawImage(canvas, x, y, width, height, 0, 0, width, height);
                        
                        const sliceDataUrl = sliceCanvas.toDataURL('image/png');
                        stampSlices.push(sliceDataUrl);
                        resolve();
                    }));
                }
                
                // 等待所有切片处理完成
                await Promise.all(slicePromises);
                
                // 更新状态（不再显示切片预览）
                updateStepStatus('step3', 'completed');
                showStatus('印章切割成功', 'success');
                resolve();
            });
        });
    } catch (error) {
        debugLog('印章切割失败:', error);
        updateStepStatus('step3', 'error');
        showStatus('印章切割失败: ' + error.message, 'error');
    }
}

/**
 * 检查生成按钮的状态，根据当前条件启用或禁用
 */
function checkGenerateButtonState() {
    const hasSlices = stampSlices.length > 0;
    const hasPdf = pdfData !== null;
    const hasStamp = stampDataUrl !== null;
    
    debugLog(`检查生成按钮状态 - 有切片: ${hasSlices}, 有PDF: ${hasPdf}, 有印章: ${hasStamp}`);
    
    // 如果有PDF和印章，但没有切片，自动切割印章
    if (hasPdf && hasStamp && !hasSlices) {
        debugLog('检测到有PDF和印章但没有切片，自动切割印章');
        sliceStamp();
        return;
    }
    
    // 更新生成按钮状态
    generateBtn.disabled = !(hasSlices && hasPdf && hasStamp);
    
    // 更新步骤4状态
    if (hasSlices && hasPdf && hasStamp) {
        updateStepStatus('step4', 'in-progress');
    } else {
        updateStepStatus('step4', 'pending');
    }
}

/**
 * 显示状态消息
 * @param {string} message - 消息内容
 * @param {string} type - 消息类型：success, error, loading
 */
function showStatus(message, type) {
    statusMessage.textContent = message;
    statusMessage.className = `status ${type}`;
    statusMessage.classList.remove('hidden');
    
    // 3秒后自动隐藏成功消息
    if (type === 'success') {
        setTimeout(() => {
            hideStatus();
        }, 3000);
    }
}

/**
 * 隐藏状态消息
 */
function hideStatus() {
    statusMessage.classList.add('hidden');
}

/**
 * 更新步骤状态指示器
 * @param {string} stepId - 步骤ID（step1, step2, step3, step4）
 * @param {string} status - 状态：pending, in-progress, completed, error
 */
function updateStepStatus(stepId, status) {
    const statusElement = document.getElementById(`${stepId}Status`);
    if (!statusElement) return;
    
    // 移除所有状态类
    statusElement.className = 'step-status';
    
    // 添加新状态类
    if (status) {
        statusElement.classList.add(status);
    }
}

/**
 * 更新进度条
 * @param {number} progress - 进度百分比（0-100）
 */
function updateProgress(progress) {
    const percentage = Math.min(100, Math.max(0, progress));
    progressFill.style.width = `${percentage}%`;
    progressText.textContent = `${percentage}%`;
}

/**
 * 生成带有骑缝章的PDF文件
 */
async function generateStampedPdf() {
    // 检查必要条件
    if (!stampSlices.length) {
        showStatus('请先上传印章图片并切割印章', 'error');
        return;
    }
    
    if (!pdfData) {
        showStatus('请先上传PDF文件', 'error');
        return;
    }
    
    // 更新状态并显示消息
    updateStepStatus('step4', 'in-progress');
    showStatus('正在生成带有骑缝章的PDF...', 'loading');
    generateBtn.disabled = true;
    
    // 显示进度条
    progressContainer.classList.remove('hidden');
    updateProgress(0);
    
    // 设置按钮为加载状态
    const originalButtonText = generateBtn.innerHTML;
    generateBtn.innerHTML = `<span class="loading-spinner"></span><span class="button-text">生成中...</span>`;
    
    try {
        // 加载PDF文档
        debugLog('开始加载PDF文档');
        const pdfDoc = await PDFLib.PDFDocument.load(pdfData);
        const pages = pdfDoc.getPages();
        
        // 预加载所有切片图片，提高性能
        const sliceImagePromises = stampSlices.map((slice, index) => {
            return createImageFromDataUrl(slice)
                .then(image => ({ image, index }))
                .catch(error => {
                    throw new Error(`加载切片图片${index + 1}失败: ${error.message}`);
                });
        });
        
        // 等待所有切片图片加载完成
        const loadedSlices = await Promise.all(sliceImagePromises);
        
        // 按照页码排序
        loadedSlices.sort((a, b) => a.index - b.index);
        
        // 实际印章大小（42mm直径，转换为点：1mm = 2.83465点）
        const actualStampDiameter = 42 * 2.83465; // 约119点
        
        // 添加骑缝章到每一页
        debugLog(`开始添加骑缝章到${pages.length}页PDF中`);
        for (let i = 0; i < pages.length; i++) {
            // 更新进度
            const progress = Math.round((i / pages.length) * 50) + 10;
            updateProgress(progress);
            
            const page = pages[i];
            const { width, height } = page.getSize();
            
            // 确保在范围内有切片
            if (i < loadedSlices.length) {
                const sliceData = loadedSlices[i];
                
                // 按照实际印章大小计算比例
                const sliceImageObj = sliceData.image;
                
                // 计算缩放比例，保持原始比例
                const scaleFactor = actualStampDiameter / Math.max(sliceImageObj.width, sliceImageObj.height);
                
                // 计算印章位置和大小
                const scaledWidth = sliceImageObj.width * scaleFactor;
                const scaledHeight = sliceImageObj.height * scaleFactor;
                
                // 获取用户选择的印章位置（垂直位置）
                const stampPosition = stampPositionSelect.value;
                let xPosition;
                let yPosition;
                
                // 印章贴着最右边
                xPosition = width - scaledWidth;
                
                // 根据选择的位置计算垂直坐标（从上往下的1/3、1/2、2/3处）
                switch (stampPosition) {
                    case '1/3':
                        yPosition = height - (height / 3) - (scaledHeight / 2); // 从上往下1/3处
                        break;
                    case '1/2':
                        yPosition = height / 2 - (scaledHeight / 2); // 垂直居中（1/2处）
                        break;
                    case '2/3':
                        yPosition = (height / 3) - (scaledHeight / 2); // 从上往下2/3处
                        break;
                    default:
                        yPosition = height / 2 - (scaledHeight / 2); // 默认垂直居中
                }
                
                // 添加印章到页面
                const slicePngBytes = await pdfDoc.embedPng(sliceData.image.src);
                page.drawImage(slicePngBytes, {
                    x: xPosition,
                    y: yPosition,
                    width: scaledWidth,
                    height: scaledHeight
                });
            }
        }
        
        // 更新进度
        updateProgress(90);
        
        // 保存生成的PDF
        debugLog('开始保存生成的PDF');
        const generatedPdfBytes = await pdfDoc.save();
        
        // 更新进度
        updateProgress(100);
        
        // 显示预览
        generatedPdfUrl = URL.createObjectURL(new Blob([generatedPdfBytes], { type: 'application/pdf' }));
        resultFrame.src = generatedPdfUrl;
        resultPreviewContainer.classList.remove('hidden');
        
        // 显示下载按钮
        downloadContainer.classList.remove('hidden');
        
        // 更新状态
        showStatus('PDF生成成功！', 'success');
        updateStepStatus('step4', 'completed');
        
        // 恢复按钮状态
        generateBtn.disabled = false;
        generateBtn.innerHTML = originalButtonText;
        
        // 隐藏进度条
        setTimeout(() => {
            progressContainer.classList.add('hidden');
        }, 1000);
        
    } catch (error) {
        debugLog('生成PDF失败:', error);
        showStatus('PDF生成失败: ' + error.message, 'error');
        updateStepStatus('step4', 'error');
        
        // 恢复按钮状态
        generateBtn.disabled = false;
        generateBtn.innerHTML = originalButtonText;
        
        // 隐藏进度条
        progressContainer.classList.add('hidden');
    }
}

/**
 * 从DataURL创建Image对象
 * @param {string} dataUrl - 图片的DataURL
 * @returns {Promise<Image>} - Promise对象，解析为Image对象
 */
function createImageFromDataUrl(dataUrl) {
    return new Promise((resolve, reject) => {
        const img = new Image();
        img.crossOrigin = 'anonymous';
        img.onload = () => resolve(img);
        img.onerror = reject;
        img.src = dataUrl;
    });
}

/**
 * 将Base64字符串转换为Uint8Array
 * @param {string} base64 - Base64字符串
 * @returns {Uint8Array} - Uint8Array对象
 */
function base64ToUint8Array(base64) {
    const binaryString = window.atob(base64);
    const binaryLen = binaryString.length;
    const bytes = new Uint8Array(binaryLen);
    for (let i = 0; i < binaryLen; i++) {
        bytes[i] = binaryString.charCodeAt(i);
    }
    return bytes;
}

/**
 * 下载生成的PDF文件
 */
function downloadPdf() {
    if (!generatedPdfUrl) return;
    
    // 检查是否在Electron环境中
    if (window.require) {
        try {
            const { dialog } = window.require('@electron/remote');
            const fs = window.require('fs');
            
            // 生成文件名：原文件名（已盖章）
            let defaultFileName = '带骑缝章的文档.pdf';
            if (originalPdfFileName) {
                // 移除原文件名的.pdf后缀
                const nameWithoutExt = originalPdfFileName.replace(/\.pdf$/i, '');
                defaultFileName = `${nameWithoutExt}（已盖章）.pdf`;
            }
            
            // 显示保存文件对话框
            dialog.showSaveDialog({
                title: '保存带骑缝章的PDF',
                defaultPath: defaultFileName,
                filters: [
                    { name: 'PDF文件', extensions: ['pdf'] }
                ]
            }).then(result => {
                if (!result.canceled && result.filePath) {
                    // 将PDF数据保存到文件
                    fetch(generatedPdfUrl)
                        .then(res => res.arrayBuffer())
                        .then(buffer => {
                            fs.writeFileSync(result.filePath, Buffer.from(buffer));
                            showStatus('PDF文件保存成功！', 'success');
                        })
                        .catch(error => {
                            debugLog('保存文件失败:', error);
                            showStatus('PDF文件保存失败: ' + error.message, 'error');
                        });
                }
            }).catch(error => {
                debugLog('保存文件对话框失败:', error);
                showStatus('保存文件对话框失败: ' + error.message, 'error');
            });
        } catch (error) {
            debugLog('使用Electron API保存失败:', error);
            // 如果Electron API失败，使用浏览器下载
            browserDownload();
        }
    } else {
        // 如果不在Electron环境中，使用浏览器下载
        browserDownload();
    }
}

/**
 * 使用浏览器下载生成的PDF文件
 */
function browserDownload() {
    if (!generatedPdfUrl) return;
    
    // 创建下载链接
    const a = document.createElement('a');
    a.href = generatedPdfUrl;
    
    // 生成文件名：原文件名（已盖章）
    let downloadFileName = '带骑缝章的文档.pdf';
    if (originalPdfFileName) {
        // 移除原文件名的.pdf后缀
        const nameWithoutExt = originalPdfFileName.replace(/\.pdf$/i, '');
        downloadFileName = `${nameWithoutExt}（已盖章）.pdf`;
    }
    
    a.download = downloadFileName;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    
    showStatus('PDF文件下载成功！', 'success');
}

// 图片缩放函数
/**
 * 缩放图片到指定尺寸
 * @param {Image} image - 图片对象
 * @param {number} maxWidth - 最大宽度
 * @param {number} maxHeight - 最大高度
 * @returns {Promise<string>} - Promise对象，解析为缩放后的图片DataURL
 */
async function resizeImage(image, maxWidth, maxHeight) {
    return new Promise((resolve, reject) => {
        try {
            // 计算缩放比例
            const scale = Math.min(maxWidth / image.width, maxHeight / image.height, 1);
            
            // 计算新尺寸
            const width = Math.round(image.width * scale);
            const height = Math.round(image.height * scale);
            
            // 创建canvas
            const canvas = document.createElement('canvas');
            canvas.width = width;
            canvas.height = height;
            
            // 绘制图片
            const ctx = canvas.getContext('2d');
            ctx.drawImage(image, 0, 0, width, height);
            
            // 获取DataURL
            const dataUrl = canvas.toDataURL('image/png');
            resolve(dataUrl);
        } catch (error) {
            reject(error);
        }
    });
}

// 页面加载完成后初始化应用
window.addEventListener('DOMContentLoaded', init);