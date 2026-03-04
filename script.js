'use strict';

const dropZone = document.getElementById('drop-zone');

dropZone.addEventListener('dragover', (e) => {
    e.preventDefault(); // ブラウザの標準動作（ファイルを開く）を止める
    dropZone.style.backgroundColor = 'rgba(255, 255, 255, 0.3)'; // 重なった時に枠の色を変える
});

// ドラッグキャンセル
dropZone.addEventListener('dragleave', (e) => {
    e.preventDefault(); // ここでもブラウザの動作を止める
    dropZone.style.backgroundColor = 'rgba(255, 255, 255, 0.1)'; // 💡 元の透明に戻す
});

// drop = ファイルをパッと離した瞬間
dropZone.addEventListener('drop', (e) => {
    e.preventDefault(); // ここでもブラウザの動作を止める
    dropZone.style.backgroundColor = 'rgba(255, 255, 255, 0.1)';

    // 投げ込まれたファイルたちの中から、1番目のファイルを取り出す
    const file = e.dataTransfer.files[0];

    // ファイルがExcelファイルかどうかをチェック
    const allowedExtensions = ['.xlsx', '.xls', '.xlsb'];
    const extension = file.name.substring(file.name.lastIndexOf('.')).toLowerCase();
    if (!allowedExtensions.includes(extension)) return alert('Excelファイルをドロップしてください');

    const reader = new FileReader();
    reader.onload = (event) => {

        const data = event.target.result; // ファイルのバイナリデータを取得
        const workbook = XLSX.read(data, { type: 'array' }); // SheetJSで「ブック」として読み込む
        const worksheet = workbook.Sheets[workbook.SheetNames[0]]; // そのシート名を使って「シート」そのものを取り出す

        // 全データを「配列の配列」として取得
        // これで rows[0][0] みたいに座標でアクセスできる
        const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

        const resultText = formatData(rows);

        // // F. コンソールで中身を確認！
        // console.log("読み込み成功！データの中身:", rows);
        // G. クリップボードにコピーする
        navigator.clipboard.writeText(resultText)
            .then(() => alert('Excelの内容を整形してクリップボードにコピーしました'))
            .catch(err => alert('コピーに失敗しました: ' + err));
    };

    // 読取実行！
    reader.readAsArrayBuffer(file);

});

btn.addEventListener('click', async () => {
    const pastedText = await navigator.clipboard.readText();

    // 1. 行に分割する（ダブルクォーテーション内の改行は無視する正規表現）
    // [^"] は「" 以外の文字」、(?:"[^"]*")* は「" で囲まれた中身」を意味するよ
    const rowMatches = pastedText.match(/(?:(?:"[^"]*")|[^"\r\n])+/g);

    if (!rowMatches) return;

    const rows = rowMatches.map(rowLine => {
        // 2. 各行をタブで分割する（ここでもセル内のタブを考慮）
        // セル内改行がある場合、セル自体が " " で囲まれているのでそれを剥がす
        const cells = rowLine.split('\t').map(cell => {
            let content = cell.trim();
            // 先頭と末尾に " があったら削除し、内部の ""（Excel特有の回避）を " に戻す
            if (content.startsWith('"') && content.endsWith('"')) {
                content = content.slice(1, -1).replace(/""/g, '"');
            }
            return content;
        });
        return cells;
    });

    // --- 【ここまで】 ---

    const finalString = formatData(rows); // ★呼び出し
    await navigator.clipboard.writeText(finalString)
        .then(() => alert('クリップボードの内容を整形しました'))
        .catch(err => alert('コピーに失敗しました: ' + err));
});


function formatData(rows) {
    let hdrIndex = rows.findIndex(row => row.includes("品目")); // 「品目」がある行を探す
    if (hdrIndex === -1) return alert("見出しが見つかりませんでした");
    console.log(rows);

    const headerRow = rows[hdrIndex];
    const lastColIdx = (headerRow[headerRow.length - 1] == "") ? headerRow.length - 1 : headerRow.length; // 最後の列のインデックスを取得
    const targetColNames = [
        "品目",
        "注文番号/\n伝票番号(オーダ)",
        "作業手順\n番号",
        "順序\n番号",
        "明細\n番号",
        "計画数量",
        "保管場所"
    ];

    const colIndices = targetColNames.map(name => {
        return headerRow.findIndex(cell => cell && cell.replace(/\s/g, '') === name.replace(/\s/g, '')); // 空白を無視して比較
    });
    colIndices.push(lastColIdx);
    console.log(colIndices);
    const resultText = rows
        .filter(row => ![undefined, ""].includes(row[lastColIdx])) // 空でない行だけ
        .map(row => {
            // 配列をタブ（\t）でつなげて一行の文字列にする
            return colIndices
                .map(idx => row[idx]) // 各インデックスの値を配列化
                .join('\t');                // それをタブで結合
        })
        .join('\n'); // 最後に行同士を改行（\n）でつなぐ

    console.log(resultText);
    return resultText;
}