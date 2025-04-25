document.getElementById('submitBtn').addEventListener('click', () => {
  const input = document.getElementById('input').value;

  // 模擬輸入檔案處理
  const inputFileContent = `Input: ${input}`;

  // 模擬輸出檔案處理
  const outputFileContent = `Output: You entered "${input}"`;

  // 顯示結果
  document.getElementById('result').textContent = outputFileContent;

  // 模擬文件存取 (在瀏覽器中下載)
  downloadFile('input.txt', inputFileContent);
  downloadFile('output.txt', outputFileContent);
});

function downloadFile(filename, content) {
  const blob = new Blob([content], { type: 'text/plain' });
  const link = document.createElement('a');
  link.href = URL.createObjectURL(blob);
  link.download = filename;
  link.click();
}
