const fs = require('fs').promises;
const path = require('path');
const xlsx = require('xlsx');
const today = new Date();

const argFilePath = process.argv[2];
const tgtFilePath = argFilePath || String.raw`C:\Users\jfe106005\Box\02_PPP事業部\03_PPP_管理G\B_損益\2026年度\a_見直し\【1】年度計画\所別資料\24E_鈴鹿\2026年度\ENG\EJ24E 鈴鹿）2026損益管理表.xlsm`;
const book = xlsx.readFile(tgtFilePath);
const outputDir = String.raw`C:\Users\jfe106005\projects\PLExtraction\csv_output`;
const detailSheet = book.Sheets['月次計画(明細)_001'];
const summarySheet = book.Sheets['データ取込'];
const header = ['no', 'WBS', 'WBS Name', 'Date', 'Review','Status', 'Title', 'Value'];
const currentYear = today.getFullYear();
const currentMonth = String(today.getMonth() + 1).padStart(2, '0');
const currentDay = today.getDate();

const siteName = summarySheet['D8']?.v || '工場不明';
const fiscalYear = summarySheet['B8']?.v || '0000';

const startRow = 13;
const endRow = 2021;
const wbsColumn = 'B';
const columns = [
    {
        review : '年度計画',
        data : {
            '04' : 'I',
            '05' : 'J',
            '06' : 'K',
            '07' : 'L', 
            '08' : 'M',
            '09' : 'N',
            '10' : 'O',
            '11' : 'P',
            '12' : 'Q',
            '01' : 'R',
            '02' : 'S',
            '03' : 'T'
        }
    },
    {
        review : '7月見直し',
        data : {
            '05' : 'Z',
            '06' : 'AA',
            '07' : 'AB', 
            '08' : 'AC',
            '09' : 'AD',
            '10' : 'AE',
            '11' : 'AF',
            '12' : 'AG',
            '01' : 'AH',
            '02' : 'AI',
            '03' : 'AJ'
        },
    },
    {
        review : '10月見直し',
        data : {
            '07' : 'AR',
            '08' : 'AS',
            '09' : 'AT',
            '10' : 'AU',
            '11' : 'AV',
            '12' : 'AW',
            '01' : 'AX',
            '02' : 'AY',
            '03' : 'AZ'
        },
    },
    {
        review : '1月見直し',
        data : {
            '11' : 'BL',
            '12' : 'BM',
            '01' : 'BN',
            '02' : 'BO',
            '03' : 'BP'
        },
    },
    {
        review : '推定実績',
        data : {
            '01' : 'CD',
            '02' : 'CE',
            '03' : 'CF'
        },
    }
];

const strWbs = (row) => {
    const buf = detailSheet[`${wbsColumn}${row}`]?.v;
    return buf ? String(buf) : null;
}

const floatValue = (c, r) => {
    const buf = Math.round(detailSheet[`${c}${r}`]?.v * 1000);
    return buf ? parseFloat(buf/1000) : null;
}

const strDate = (month) => {
    const m = String(month).padStart(2, '0');
    const year = parseInt(month) >= 4 ? fiscalYear : fiscalYear + 1;
    return `${year}/${m}/01`;
}

const offsetMonth = (m) => {
    if (m >= 4) {
        return m;
    } else {
        return m + 12;
    }
}

const normalizeDate = dateStr => dateStr.replace(/(\d{4})(\d{2})(\d{2})/, "$1/$2/$3");

async function main() {

    const csvWbsList = await fs.readFile(String.raw`C:\Users\jfe106005\projects\PLExtraction\WBS.csv`, "utf-8");
    const wbsArray = csvWbsList.toString().trim().split(/\r?\n/).map(o => o.split(','));

    const csvActual = await fs.readFile(String.raw`C:\Users\jfe106005\Box\01_jfe106005\VDPoutput\\${siteName}実績明細表.csv`, "utf-8");
    const actualArray = csvActual.toString().trim().split(/\r?\n/).map(o => o.split(',')).slice(1);

    const intWbsNumber = (wbsCode) => {
        const wbsEntry = wbsArray.find(d => d[1] === wbsCode);
        return wbsEntry && wbsEntry[0] ? parseInt(wbsEntry[0]) : null;
    }

    const strWbsName = (wbsCode) => {
        const wbsEntry = wbsArray.find(d => d[1] === wbsCode);
        return wbsEntry && wbsEntry[2] ? wbsEntry[2] : null;
    }
    
    function getItem(data, review) {
        const result = [];
        const title = review;
        for(let r = startRow; r <= endRow; r++) {
            const wbs = strWbs(r);
            if (wbs === null) continue;

            const dataByWBS = Object.entries(data).map(([m, c]) => {
                const value = floatValue(c, r);
                if (value === null) {
                    return null;
                } else {
                    return [intWbsNumber(wbs), wbs, strWbsName(wbs), strDate(m), review, '予算', title, value]
                }
            })
            if (dataByWBS !== null) result.push(...dataByWBS);
        }
        return result.filter(o => o !== null);
    }       

    function pushActual(result) {
        actualArray.forEach(o => result.push([intWbsNumber(o[7]), o[7], strWbsName(o[7]), normalizeDate(o[13]), '最終実績', o[15], o[17], parseFloat(o[20]/1000)]));
    }

    function plExtraction() {
        const result = [header];
        columns.map(o => getItem(o.data, o.review)).forEach(byType => result.push(...byType));
        pushActual(result);
        return result;
    }

    const plResult = plExtraction();
    const r0 = plResult.filter(o => o[5] === '実績' && o[3] < `${currentYear}/${currentMonth}/01`);
    const r1 = r0.filter(o => o[3] < `${currentYear}/05/01`).map(o => [...o.slice(0, 4), '7月見直し', ...o.slice(5)]); // 7月見直し前の実績
    const r2 = r0.filter(o => o[3] <`${currentYear}/07/01`).map(o => [...o.slice(0, 4), '10月見直し', ...o.slice(5)]); // 10月見直し前の実績
    const r3 = r0.filter(o => o[3] < `${currentYear}/11/01`).map(o => [...o.slice(0, 4), '1月見直し', ...o.slice(5)]); // 1月見直し前の実績
    const r4 = r0.filter(o => o[3] < `${currentYear + 1}/01/01`).map(o => [...o.slice(0, 4), '推定実績', ...o.slice(5)]); // 推定実績前の実績

    plResult.push(...r1, ...r2, ...r3, ...r4);
        
    function toCsv(arr){
        return arr.map(o => o.join(',')).join('\n');
    }

    function log() {
        console.log(toCsv(columns.map(o => getItem(o.data, o.review))[0]));
        console.log(columns.map(o => getItem(o.data, o.review)));
    };

    await fs.writeFile(
        String.raw`${outputDir}\PLExtract_${siteName})${fiscalYear}.csv`,
        toCsv(plResult),
        "utf-8"
    );
}

//log();
main().catch(err => {
    console.error('エラー:', err);
    process.exit(0);
});
