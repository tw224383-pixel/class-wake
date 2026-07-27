const express = require('express');
const multer = require('multer');
const xlsx = require('xlsx');
const path = require('path');

const app = express();
const port = 3000;

// エクセルファイルはメモリ上にのみ保持し、ディスクには保存しない
const storage = multer.memoryStorage();
const upload = multer({ storage: storage });

// JSONのパース制限を少し大きくする（エクスポート時にクラス全データを送るため）
app.use(express.json({ limit: '50mb' }));
app.use(express.urlencoded({ extended: true, limit: '50mb' }));

// 静的ファイルの提供（フロントエンド）
app.use(express.static(path.join(__dirname, 'public')));

// 1. エクセルアップロードとデータ解析エンドポイント
app.post('/api/upload', upload.single('file'), (req, res) => {
    try {
        if (!req.file) {
            return res.status(400).json({ error: 'ファイルがありません' });
        }
        
        // メモリ上のバッファから直接エクセルを読み込む
        const workbook = xlsx.read(req.file.buffer, { type: 'buffer' });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        
        // JSONデータに変換
        const data = xlsx.utils.sheet_to_json(worksheet, { defval: '' });
        
        res.json({ success: true, data: data });
    } catch (error) {
        console.error(error);
        res.status(500).json({ error: 'ファイルの読み込みに失敗しました' });
    }
});

// 2. 結果のエクスポートエンドポイント
app.post('/api/export', (req, res) => {
    try {
        const { classesData } = req.body;
        
        if (!classesData || !Array.isArray(classesData)) {
            return res.status(400).json({ error: '無効なデータ形式です' });
        }

        const workbook = xlsx.utils.book_new();
        const allExportData = [];
        
        // クラスごとの処理
        classesData.forEach((clsObj, index) => {
            const className = `${index + 1}組`;
            if (clsObj.students && Array.isArray(clsObj.students)) {
                // 1. このクラスの生徒データ（内部プロパティ除外）
                let clsStudents = clsObj.students.map(s => {
                    const { id, okNames, ngNames, ...exportStudent } = s;
                    // okNames, ngNamesなどは必要なら文字列化するか除外
                    // フロント側でそのまま送ってくるので、必要なものだけ抽出する
                    return {
                        '決定クラス': className,
                        '氏名': s.name,
                        'ふりがな': s.kana,
                        '性別': s.gender,
                        '点数': s.score,
                        '50m/100m走': s.sprintTime || '',
                        'ピアノ': s.piano ? '〇' : '',
                        'リーダー': s.leader ? '〇' : '',
                        '配慮事項': s.notes || '',
                        '保護者注意': s.warning ? '〇' : '',
                        '旧クラス': s.oldClass || '',
                        '同じクラスにしない人': s.ngNames ? s.ngNames.join(', ') : '',
                        '同じクラスにする人': s.okNames ? s.okNames.join(', ') : ''
                    };
                });
                
                // 全体リスト用に追加
                allExportData.push(...clsStudents);
                
                // 2. クラスごと（五十音ソート）
                let byKana = [...clsStudents].sort((a, b) => a['ふりがな'].localeCompare(b['ふりがな'], 'ja'));
                let wsKana = xlsx.utils.json_to_sheet(byKana);
                xlsx.utils.book_append_sheet(workbook, wsKana, `${className}(五十音)`);
                
                // 3. クラスごと（男女別五十音ソート）
                let byGenderKana = [...clsStudents].sort((a, b) => {
                    // 男を先に、女を後に
                    let ga = a['性別'] === '男' ? 0 : (a['性別'] === '女' ? 1 : 2);
                    let gb = b['性別'] === '男' ? 0 : (b['性別'] === '女' ? 1 : 2);
                    if (ga !== gb) return ga - gb;
                    return a['ふりがな'].localeCompare(b['ふりがな'], 'ja');
                });
                let wsGenderKana = xlsx.utils.json_to_sheet(byGenderKana);
                xlsx.utils.book_append_sheet(workbook, wsGenderKana, `${className}(男女別)`);
            }
        });
        
        // 全体シート（クラス名順、五十音）
        allExportData.sort((a, b) => {
            if (a['決定クラス'] !== b['決定クラス']) return a['決定クラス'].localeCompare(b['決定クラス'], 'ja');
            return a['ふりがな'].localeCompare(b['ふりがな'], 'ja');
        });
        const wsAll = xlsx.utils.json_to_sheet(allExportData);
        // 全体シートを一番前に挿入するため、直接Sheetsに追加するか、順番にappendするかですが
        // book_append_sheetで末尾に追加しても問題ありません。ここでは全体を最後にします。
        // もし一番前にしたい場合は直接 workbook.SheetNames.unshift() 等を行います。
        xlsx.utils.book_append_sheet(workbook, wsAll, "全体結果");

        
        // バッファとして書き出し
        const buffer = xlsx.write(workbook, { type: 'buffer', bookType: 'xlsx' });
        
        // クライアントへダウンロードとして送信
        res.setHeader('Content-Disposition', 'attachment; filename="class_result.xlsx"');
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.send(buffer);
    } catch (error) {
        console.error(error);
        res.status(500).json({ error: 'エクセルの作成に失敗しました' });
    }
});

// 3. テンプレートのエクスポートエンドポイント
app.get('/api/template', (req, res) => {
    try {
        const templateData = [
            {
                '氏名': '山田 太郎',
                '性別': '男',
                '点数': 80,
                'ピアノ': '',
                'リーダー': '〇',
                '配慮事項': '視力が弱いため前方の席希望',
                '保護者注意': '',
                '旧クラス': '3-1',
                '同じクラスにしない人': '鈴木 一郎, 田中 次郎',
                '同じクラスにする人': ''
            },
            {
                '氏名': '佐藤 花子',
                '性別': '女',
                '点数': 85,
                'ピアノ': '〇',
                'リーダー': '',
                '配慮事項': '',
                '保護者注意': '',
                '旧クラス': '3-2',
                '同じクラスにしない人': '鈴木 一郎',
                '同じクラスにする人': '伊藤 美咲'
            }
        ];
        
        const worksheet = xlsx.utils.json_to_sheet(templateData);
        const workbook = xlsx.utils.book_new();
        xlsx.utils.book_append_sheet(workbook, worksheet, "名簿テンプレート");
        
        const buffer = xlsx.write(workbook, { type: 'buffer', bookType: 'xlsx' });
        
        res.setHeader('Content-Disposition', 'attachment; filename="template.xlsx"');
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.send(buffer);
    } catch (error) {
        console.error(error);
        res.status(500).json({ error: 'テンプレートの作成に失敗しました' });
    }
});

app.listen(port, () => {
    console.log(`システムが起動しました: http://localhost:${port}`);
});
