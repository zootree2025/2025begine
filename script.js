// DOM Elements
const input = document.getElementById('input');
const submitBtn = document.getElementById('submitBtn');
const resultDiv = document.getElementById('result');

// Event Listener for Button
submitBtn.addEventListener('click', () => {
  const inputValue = input.value;
  if (inputValue.trim() === '') {
    resultDiv.textContent = 'Please enter a value!';
  } else {
    resultDiv.textContent = `You entered: ${inputValue}`;
  }
});
