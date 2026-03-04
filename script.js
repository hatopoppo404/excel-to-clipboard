'use strict';

const dropZone = document.getElementById('drop-zone');

dropZone.addEventListener('dragover', (e) => {
    e.preventDefault(); // ブラウザの標準動作（ファイルを開く）を止める
    dropZone.style.borderColor = 'white'; // ヒント：重なった時に枠の色を変えると親切！
});