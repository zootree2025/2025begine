document.getElementById('convertBtn').addEventListener('click', function () {
    const txtFile = document.getElementById('txtFile').files[0];
    const fontColor = document.getElementById('fontColor').value;
    const bgColor = document.getElementById('bgColor').value;
    const slideRatio = document.getElementById('slideRatio').value;

    if (!txtFile) {
        alert('請選擇一個 TXT 檔案');
        return;
    }

    const pptx = new PptxGenJS();

    // 設定版面比例
    if (slideRatio === '16:9') {
        pptx.layout = 'LAYOUT_16x9';
    } else if (slideRatio === '4:3') {
        pptx.layout = 'LAYOUT_4x3';
    }

    const reader = new FileReader();
    reader.onload = function (e) {
        const text = e.target.result;
        const slides = text.split(/\n\s*\n/); // 使用空白行分割投影片

        slides.forEach(slideText => {
            const slide = pptx.addSlide();
            slide.background = { fill: bgColor };
            slide.color = fontColor;

            const lines = slideText.split('/N'); // 使用 /N 分割換行
            lines.forEach((line, idx) => {
                slide.addText(line, {
                    x: 0.5,
                    y: 0.5 + idx * 0.5,
                    fontSize: 18,
                    color: fontColor
                });
            });
        });

        pptx.writeFile({ fileName: 'presentation.pptx' });
    };

    reader.readAsText(txtFile, 'UTF-8');
});
