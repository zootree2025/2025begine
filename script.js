document.addEventListener('DOMContentLoaded', function() {
    // 顏色預覽功能
    const fontColor = document.getElementById('fontColor');
    const bgColor = document.getElementById('bgColor');
    const fontColorPreview = document.getElementById('fontColorPreview');
    const bgColorPreview = document.getElementById('bgColorPreview');
    
    fontColor.addEventListener('input', function() {
        fontColorPreview.style.backgroundColor = this.value;
    });
    
    bgColor.addEventListener('input', function() {
        bgColorPreview.style.backgroundColor = this.value;
    });
    
    // 初始化顏色預覽
    fontColorPreview.style.backgroundColor = fontColor.value;
    bgColorPreview.style.backgroundColor = bgColor.value;
    
    // 轉換按鈕點擊事件
    document.getElementById('convertBtn').addEventListener('click', function() {
        const fileInput = document.getElementById('txtFile');
        const file = fileInput.files[0];
        
        if (!file) {
            alert('請選擇一個 TXT 文件');
            return;
        }
        
        const reader = new FileReader();
        reader.onload = function(e) {
            const content = e.target.result;
            createPPT(content, fontColor.value, bgColor.value);
        };
        reader.readAsText(file, 'UTF-8');
    });
    
    // 創建 PPT 函數
    function createPPT(content, fontColor, bgColor) {
        try {
            // 检查 PptxGenJS 是否可用
            if (typeof pptxgen !== 'function') {
                throw new Error('PptxGenJS 库未正确加载，请刷新页面重试');
            }
            
            // 使用 PptxGenJS 創建 PPT
            const pptx = new pptxgen();
            
            // 設置默認幻燈片屬性
            pptx.defineLayout({ name: 'LAYOUT_16x9', width: 10, height: 5.625 });
            pptx.layout = 'LAYOUT_16x9';
            
            // 分割內容為幻燈片
            const slides = content.split(/\n{3,}/); // 兩個以上空白行分隔
            
            slides.forEach(slideContent => {
                if (slideContent.trim() === '') return;
                
                // 創建新幻燈片
                const slide = pptx.addSlide();
                
                // 設置背景顏色
                slide.background = { color: bgColor };
                
                // 處理換行符 /N
                const processedContent = slideContent.replace(/\/N/g, '\n');
                
                // 添加文本框
                slide.addText(processedContent, {
                    x: 0.5,
                    y: 0.5,
                    w: '90%',
                    h: '80%',
                    color: fontColor,
                    fontSize: 24,
                    align: 'left',
                    valign: 'top'
                });
            });
            
            // 保存 PPT
            pptx.writeFile({ fileName: 'presentation.pptx' })
                .then(() => {
                    document.getElementById('result').classList.remove('hidden');
                })
                .catch(err => {
                    console.error('PPT 生成失敗:', err);
                    alert('PPT 生成失敗: ' + err.message);
                });
        } catch (error) {
            console.error('生成 PPT 時發生錯誤:', error);
            alert('生成 PPT 時發生錯誤: ' + error.message);
        }
    }
});