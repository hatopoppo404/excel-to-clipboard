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

        let hdrIndex = rows.findIndex(row => row.includes("品目")); // 「品目」がある行を探す
        if (hdrIndex === -1) return alert("見出しが見つかりませんでした");

        const headerRow = rows[hdrIndex]; // 見出し行そのもの
        const pnColIdx = headerRow.indexOf("品目");
        const poColIdx = headerRow.indexOf("注文番号/\n伝票番号(オーダ)");
        const tejunColIdx = headerRow.indexOf("作業手順\n番号");
        const junjoColIdx = headerRow.indexOf("順序\n番号");
        const meisaiColIdx = headerRow.indexOf("明細\n番号");
        const qtyColIdx = headerRow.indexOf("計画数量");
        const whColIdx = headerRow.indexOf("保管場所");
        const lastColIdx = headerRow.length;

        const resultText = rows.slice(hdrIndex + 1)
            .filter(row => row[lastColIdx] !== undefined && row[lastColIdx] !== "") // 品目が空でない行だけ
            .map(row => {
                // 列の順番を自由に入れ替えて、新しい配列を作る
                const rearranged = [
                    row[pnColIdx],
                    row[poColIdx],
                    row[tejunColIdx],
                    row[junjoColIdx],
                    row[meisaiColIdx],
                    row[qtyColIdx],
                    row[whColIdx],
                    row[lastColIdx]
                ];

                // 配列をタブ（\t）でつなげて一行の文字列にする
                return rearranged.join('\t');
            })
            .join('\n'); // 最後に行同士を改行（\n）でつなぐ


        // // F. コンソールで中身を確認！
        // console.log("読み込み成功！データの中身:", rows);
        // G. クリップボードにコピーする
        navigator.clipboard.writeText(resultText)
            .then(() => alert('Excelの内容をクリップボードにコピーしました！'))
            .catch(err => alert('コピーに失敗しました: ' + err));
    };

    // 読取実行！
    reader.readAsArrayBuffer(file);

});