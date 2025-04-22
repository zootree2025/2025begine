document.addEventListener('DOMContentLoaded', function() {
    const fontColor = document.getElementById('fontColor');
    const bgColor = document.getElementById('bgColor');
    const fontColorPreview = document.getElementById('fontColorPreview');
    const bgColorPreview = document.getElementById('bgColorPreview');

    fontColor.addEventListener('input', function () {
        fontColorPreview.style.backgroundColor = this.value;
    });

    bgColor.addEventListener('input', function () {
        bgColorPreview.style.backgroundColor = this.value;
    });

    fontColorPreview.style.backgroundColor = fontColor.value;
    bgColorPreview.style.backgroundColor = bgColor.value;

    document.getElementById('convertBtn').addEventListener('click', function () {
        const fileInput = document.getElementById('txtFile');
        const file = fileInput.files[0];

        if (!file) {
            alert('請選擇一個 TXT 文件');
            return;
        }

        const reader = new FileReader();
        reader.onload = function (e) {
            const content = e.target.result;
            createPPT(content, fontColor.value, bgColor.value);
        };
        reader.readAsText(file, 'UTF-8');
    });

    function createPPT(content, fontColor, bgColor) {
        try {
            if (typeof PptxGenJS === 'undefined') {
            throw new Error('PptxGenJS 库未正確加載，請刷新頁面重試');
            }

            const pptx = new PptxGenJS();

            pptx.defineLayout({ name: 'LAYOUT_16x9', width: 10, height: 5.625 });
            pptx.layout = 'LAYOUT_16x9';

            // 修改重點：使用兼容不同系統換行符的正則表達式
            const slides = content.split(/(\r?\n){2,}/); 

            slides.forEach(slideContent => {
                if (slideContent.trim() === '') return;

                const slide = pptx.addSlide();
                slide.background = { color: bgColor };

                const processedContent = slideContent.replace(/\/N/g, '\n');

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