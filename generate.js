const xlsx = require('xlsx');

const lastNames = ['佐藤', '鈴木', '高橋', '田中', '伊藤', '渡辺', '山本', '中村', '小林', '加藤', '吉田', '山田', '佐々木', '山口', '松本'];
const lastKana = ['さとう', 'すずき', 'たかはし', 'たなか', 'いとう', 'わたなべ', 'やまもと', 'なかむら', 'こばやし', 'かとう', 'よしだ', 'やまだ', 'ささき', 'やまぐち', 'まつもと'];

const firstNamesM = ['太郎', '健太', '翔太', '蓮', '大輝', '陽翔', '颯太', '樹', '湊', '大和'];
const firstKanaM = ['たろう', 'けんた', 'しょうた', 'れん', 'だいき', 'はると', 'そうた', 'いつき', 'みなと', 'やまと'];

const firstNamesF = ['花子', '結衣', '陽菜', '莉子', '葵', '結菜', '凛', '美咲', 'さくら', '結愛'];
const firstKanaF = ['はなこ', 'ゆい', 'ひな', 'りこ', 'あおい', 'ゆいな', 'りん', 'みさき', 'さくら', 'ゆあ'];

const students = [];

for (let i = 1; i <= 100; i++) {
    const isMale = Math.random() < 0.5;
    const lastNameIdx = Math.floor(Math.random() * lastNames.length);
    const firstNameIdx = Math.floor(Math.random() * (isMale ? firstNamesM.length : firstNamesF.length));
    
    const name = `${lastNames[lastNameIdx]} ${isMale ? firstNamesM[firstNameIdx] : firstNamesF[firstNameIdx]}`;
    const kana = `${lastKana[lastNameIdx]} ${isMale ? firstKanaM[firstNameIdx] : firstKanaF[firstNameIdx]}`;
    
    const isPiano = Math.random() < 0.05; // 5% chance
    const isLeader = Math.random() < 0.15; // 15% chance
    const isWarning = Math.random() < 0.08; // 8% chance
    
    const oldClass = Math.floor(Math.random() * 3) + 1; // 1, 2, or 3
    
    const score = Math.floor(Math.random() * 41) + 60;
    
    // 50m走のタイム (7.0 〜 11.0秒)
    const sprint = isMale ? (Math.random() * 3 + 7.0).toFixed(1) : (Math.random() * 3 + 7.5).toFixed(1);
    
    students.push({
        '氏名': name,
        'ふりがな': kana,
        '性別': isMale ? '男' : '女',
        '点数': score,
        '50m走': sprint,
        'ピアノ': isPiano ? '〇' : '',
        'リーダー': isLeader ? '〇' : '',
        '配慮事項': isWarning ? 'アレルギーあり、視力低下' : '',
        '保護者注意': isWarning ? '〇' : '',
        '旧クラス': `${oldClass}組`,
        '同じクラスにしない人': '',
        '同じクラスにする人': ''
    });
}

// NGとOKをランダムに少し設定する
// 100人のうち10人くらいに設定
for (let i = 0; i < 15; i++) {
    const p1 = Math.floor(Math.random() * 100);
    const p2 = Math.floor(Math.random() * 100);
    if (p1 !== p2) {
        students[p1]['同じクラスにしない人'] = students[p2]['氏名'];
    }
}
for (let i = 0; i < 5; i++) {
    const p1 = Math.floor(Math.random() * 100);
    const p2 = Math.floor(Math.random() * 100);
    if (p1 !== p2) {
        students[p1]['同じクラスにする人'] = students[p2]['氏名'];
    }
}

const workbook = xlsx.utils.book_new();
const worksheet = xlsx.utils.json_to_sheet(students);
xlsx.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
xlsx.writeFile(workbook, 'test_100_students.xlsx');

console.log('test_100_students.xlsx was generated successfully.');
