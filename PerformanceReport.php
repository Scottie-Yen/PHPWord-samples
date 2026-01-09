<?php

ini_set('display_errors', 1);
ini_set('display_startup_errors', 1);
error_reporting(E_ALL);

// api
function httpRequest($api, $data_string) {

    $ch = curl_init($api);
    curl_setopt($ch, CURLOPT_POST, 1);
    curl_setopt($ch, CURLOPT_CUSTOMREQUEST, "POST");
    curl_setopt($ch, CURLOPT_POSTFIELDS, $data_string);
    curl_setopt($ch, CURLOPT_RETURNTRANSFER, true);
    curl_setopt($ch, CURLOPT_HTTPHEADER, array(
        'Content-Type: application/json',
        'Content-Length: ' . strlen($data_string),
		'authorization: authorization ' . $_GET["token"]
    ));
    $result = curl_exec($ch);
    curl_close($ch);
  
    return json_decode($result, true);
}

// 包裝百分比
function toPercent($a, $b = null) {
    // 轉成浮點數
    $a = (float)$a;

    // 單一參數：直接轉百分比（無正負號）
    if ($b === null) {
        return number_format($a * 100, 2) . '%';
    }

    // 兩個參數：相減後加上正負號
    $b = (float)$b;
    $delta = ($a - $b) * 100;                // 轉成百分比
    return sprintf('%+.2f%%', $delta);       // 帶正負號與百分比
}

// 包裝件數
function toCount($type, $a, $b = null) {
    $str = ''; 

    if ($type === 'a') {
        $str = '件';
    } else {
        $str = '元';
    }

    if ($b === null) {
        if ($type === 'a') {
            return $a . $str;
        } else {
            return number_format($a) . $str;
        }
    }

    $a = (float)$a;
    $b = (float)$b;
    $delta = $a - $b; 
    if ($type === 'a') {
        return sprintf('%+d'.$str, $delta);
    } else {
        $formattedNum = number_format($delta);

        if ($delta >= 0) {
            return '+' . $formattedNum . $str; // 加上 + 號
        } else {
            return $formattedNum . $str; // 加上 - 號
        }
    } 
}



use PhpOffice\PhpWord\Element\Section;
use PhpOffice\PhpWord\Element\Table;
use PhpOffice\PhpWord\SimpleType\JcTable;
use PhpOffice\PhpWord\SimpleType\Jc;
use PhpOffice\PhpWord\Shared\Converter;

/**
 * 固定樣式的「業績狀況」表格
 * - 表頭：灰底 D9D9D9、置中，文字「業績狀況」
 * - 表格：邊框 6、黑色、cellMargin 60、右對齊、寬度 50%
 * - 儲存格：內容置中
 *
 * @param PhpOffice\PhpWord\PhpWord $phpWord
 * @param Section $section
 * @param string $styleName 樣式名稱
 * @param int $w       表格寬度（用於計算每列欄位寬度）
 * @param int $rowNum  每列欄位數（用於表頭 gridSpan）
 * @param array $info  列資料（例如：[ ['業績','目標',...], ['服務費','目標',...] ]）
 * @return Table
 */
function renderPerformanceTableFixed(
    \PhpOffice\PhpWord\PhpWord $phpWord,
    Section $section,
    string $styleName,
    string $title,
    int $w,
    int $rowNum,
    array $info
): Table {
    // 註冊樣式（固定）
    $phpWord->addTableStyle($styleName, [
        'borderSize' => 6,
        'borderColor'=> '000000',
        'cellMargin' => 60,
        'alignment' => \PhpOffice\PhpWord\SimpleType\JcTable::END,
    ]);

    // 建立表格
    $tbl = $section->addTable($styleName);

    // 表頭（固定樣式）
    $titleSize = 12;
    $fontSize = 12;
    if ($title === '業績狀況') {
        $fontSize = 10;
    }
    if ($title === '開案狀況') {
        $fontSize = 10;
    }

    $row  = $tbl->addRow(null);
    $cell = $row->addCell($w, ['gridSpan' => $rowNum, 'bgColor' => 'D9D9D9', 'valign' => 'center']);
    $cell->addText($title, ['size' => $titleSize,], ['alignment' => Jc::CENTER, 'spaceBefore' => 0, 'spaceAfter'  => 0, 'lineHeight'  => 1.0]);

    // 內容列（固定置中）
    foreach ($info as $r) {
        $row = $tbl->addRow(null);
        foreach ($r as $text) {
            $row->addCell($w/$rowNum, ['valign' => 'center'])
                ->addText((string)$text, ['size' => $fontSize,], ['alignment' => Jc::CENTER, 'spaceBefore' => 0, 'spaceAfter'  => 0, 'lineHeight'  => 1.0]);
        }
    }

    return $tbl;
}

// 國外新申請案整理陣列
function NewForeignApplicationCaseSort(array $rows, int $chunkSize = 8): array {
    $labels = [];
    $counts = [];

    foreach ($rows as $r) {
        $labels[] = $r['CountryName'];
        $counts[] = $r['Count'];
    }

    // 依大小切塊
    $labelChunks = array_chunk($labels, $chunkSize);
    $countChunks = array_chunk($counts, $chunkSize);

    // 交錯合併成 [labels1, counts1, labels2, counts2, ...]
    $result = [];
    $n = count($labelChunks);
    for ($i = 0; $i < $n; $i++) {
        $result[] = $labelChunks[$i];
        $result[] = $countChunks[$i] ?? [];
    }
    return $result;
}

// POST
$data='';

if(isset($_GET["id"]) && isset($_GET["number"]) && isset($_GET["type"]) && isset($_GET["year"])){
	$data = array(
		"function" => $_GET["function"],
        "id" => $_GET["id"],
        "name" => $_GET["name"],
        "number" => (int)$_GET["number"],
        "type" => $_GET["type"],
        "year" => $_GET["year"],
	);
}else{
	echo "Error : 參數錯誤";
}

$host = $_SERVER['HTTP_HOST']; // 例如 "192.168.0.52:8080" 或 "example.com"
if ($host === '192.168.0.52' || $host === '192.168.0.52:8080') {
    $data = httpRequest('http://192.168.0.52:8000/report/select-salesrperformancereport', json_encode($data));
} elseif ($host === '192.168.0.188' || $host === '192.168.0.188:8080') {
    $data = httpRequest('http://192.168.0.186:8000/report/select-salesrperformancereport', json_encode($data));
}
// $data = httpRequest('http://192.168.0.52:8000/report/select-salesrperformancereport', json_encode($data));
// $data = httpRequest('http://192.168.0.186:8000/report/select-salesrperformancereport', json_encode($data));
$Data = $data['Data'];
// echo json_encode($data, JSON_UNESCAPED_UNICODE);

// New Word document
include_once 'Sample_Header.php';
$phpWord = new \PhpOffice\PhpWord\PhpWord();
// use PhpOffice\PhpWord\Shared\Converter;
// use PhpOffice\PhpWord\SimpleType\Jc;

$w = 10500;

// New portrait section
$section = $phpWord->addSection(array(
	'headerHeight' => 300,
    'marginTop' => 2000,
	'marginLeft' => 0,
    'marginRight' => 750,
	'footerHeight' => 0,
));

// Add header for all other pages
$header = $section->addHeader();
$header->addImage('images/巨群信頭_專利商標所 CH (Top).png', 
array(
	'width' => 600, 
	'height' => 'auto', 
));

// ===== 1) 文件標題 =====
$section->addText(
    '業績報告',
    ['name' => 'Microsoft JhengHei', 'size' => 26, 'bold' => true],
    ['alignment' => Jc::CENTER, 'spaceAfter' => 100, 'indentation' => [ 'left'  => Converter::cmToTwip(1)]],
    
);

// ===== 2) 基本資訊：業務員 / 統計時間 =====
$section->addText('業務員： '. $data['Data']['Name'], ['name' => 'Microsoft JhengHei', 'size' => 12], ['spaceAfter' => 0, 'indentation' => [ 'left'  => Converter::cmToTwip(1.25)]]);
$Statistical = '';

if ($_GET["type"] === 'Year') {
    $Statistical = '統計時間：' . $_GET["year"] . ' 年度';
} else if ($_GET["type"] === 'Season') {
    $Statistical = '統計時間：' . $_GET["year"] . ' 年 第 '. $_GET["number"] .' 季';
} else {
    $Statistical = '統計時間：' . $_GET["year"] . ' 年 '. $_GET["number"] .' 月';
}

$section->addText($Statistical , ['name' => 'Microsoft JhengHei', 'size' => 12], ['spaceAfter' => 300, 'indentation' => [ 'left'  => Converter::cmToTwip(1.25)]]);



// ===== 3) 業績狀況： =====
$Per = $Data['Performance'];
$Info = [
    ['業績', '目標',  number_format($Per['CaseSale']['Target']), '實達', number_format($Per['CaseSale']['Sale']), '達成率', toPercent($Per['CaseSale']['Percent']), '去年同期', toPercent($Per['CaseSale']['Percent'], $Per['CaseSale']['YoYPercent'])],
    ['服務費', '目標',  number_format($Per['Incasesale']['Target']), '實達', number_format($Per['Incasesale']['Sale']), '達成率', toPercent($Per['Incasesale']['Percent']), '去年同期', toPercent($Per['Incasesale']['Percent'], $Per['Incasesale']['YoYPercent'])],
];

renderPerformanceTableFixed($phpWord, $section, 'PerformanceTable', '業績狀況', $w, 9, $Info);
$section->addTextBreak(1);

// ===== 4) 開案狀況： =====
$Pro = $Data['ProjectStatus'];
$Info = [
    ['專利新案', 
    '台灣件數',  toCount('a', $Pro['PatentNewCase']['TaiwanCount']), 
    '台灣服務費', toCount('b', $Pro['PatentNewCase']['TaiwanIncasesale']), '去年同期', toCount('a', $Pro['PatentNewCase']['TaiwanCount'], $Pro['PatentNewCase']['TaiwanLastYearCount']), toCount('b', $Pro['PatentNewCase']['TaiwanIncasesale'], $Pro['PatentNewCase']['TaiwanLastYearIncasesale']), 
    ],
    ['', 
    '國外件數', toCount('a', $Pro['PatentNewCase']['ForeignCount']),
    '國外服務費', toCount('b', $Pro['PatentNewCase']['ForeignIncasesale']), '去年同期', toCount('a', $Pro['PatentNewCase']['ForeignCount'], $Pro['PatentNewCase']['ForeignLastYearCount']), toCount('b', $Pro['PatentNewCase']['ForeignIncasesale'], $Pro['PatentNewCase']['ForeignLastYearIncasesale'])
    ],
    ['商標新案', 
    '台灣件數',  toCount('a', $Pro['TrademarkNewCase']['TaiwanCount']),
    '台灣服務費', toCount('b', $Pro['TrademarkNewCase']['TaiwanIncasesale']), '去年同期', toCount('a', $Pro['TrademarkNewCase']['TaiwanCount'], $Pro['TrademarkNewCase']['TaiwanLastYearCount']), toCount('b', $Pro['TrademarkNewCase']['TaiwanIncasesale'], $Pro['TrademarkNewCase']['TaiwanLastYearIncasesale']), 
    ],
    ['',
    '國外件數', toCount('a', $Pro['TrademarkNewCase']['ForeignCount']),
    '國外服務費', toCount('b', $Pro['TrademarkNewCase']['ForeignIncasesale']), '去年同期', toCount('a', $Pro['TrademarkNewCase']['ForeignCount'], $Pro['TrademarkNewCase']['ForeignLastYearCount']), toCount('b', $Pro['TrademarkNewCase']['ForeignIncasesale'], $Pro['TrademarkNewCase']['ForeignLastYearIncasesale'])
    ],
    ['專利年費', 
    '台灣件數',  toCount('a', $Pro['PatentAnnuity']['TaiwanCount']),
    '台灣服務費', toCount('b', $Pro['PatentAnnuity']['TaiwanIncasesale']), '去年同期', toCount('a', $Pro['PatentAnnuity']['TaiwanCount'], $Pro['PatentAnnuity']['TaiwanLastYearCount']), toCount('b', $Pro['PatentAnnuity']['TaiwanIncasesale'], $Pro['PatentAnnuity']['TaiwanLastYearIncasesale']), 
    ],
    ['', 
    '國外件數', toCount('a', $Pro['PatentAnnuity']['ForeignCount']),
    '國外服務費', toCount('b', $Pro['PatentAnnuity']['ForeignIncasesale']), '去年同期', toCount('a', $Pro['PatentAnnuity']['ForeignCount'], $Pro['PatentAnnuity']['ForeignLastYearCount']), toCount('b', $Pro['PatentAnnuity']['ForeignIncasesale'], $Pro['PatentAnnuity']['ForeignLastYearIncasesale'])
    ],
    ['專利答辯', 
    '台灣件數',  toCount('a', $Pro['PatentAnswer']['TaiwanCount']), 
    '台灣服務費', toCount('b', $Pro['PatentAnswer']['TaiwanIncasesale']), '去年同期', toCount('a', $Pro['PatentAnswer']['TaiwanCount'], $Pro['PatentAnswer']['TaiwanLastYearCount']), toCount('b', $Pro['PatentAnswer']['TaiwanIncasesale'], $Pro['PatentAnswer']['TaiwanLastYearIncasesale']), 
    ],
    ['', 
    '國外件數', toCount('a', $Pro['PatentAnswer']['ForeignCount']),
    '國外服務費', toCount('b', $Pro['PatentAnswer']['ForeignIncasesale']), '去年同期', toCount('a', $Pro['PatentAnswer']['ForeignCount'], $Pro['PatentAnswer']['ForeignLastYearCount']), toCount('b', $Pro['PatentAnswer']['ForeignIncasesale'], $Pro['PatentAnswer']['ForeignLastYearIncasesale'])
    ],
    ['專利領證', 
    '台灣件數',  toCount('a', $Pro['PatentIssuance']['TaiwanCount']),
    '台灣服務費', toCount('b', $Pro['PatentIssuance']['TaiwanIncasesale']), '去年同期', toCount('a', $Pro['PatentIssuance']['TaiwanCount'], $Pro['PatentIssuance']['TaiwanLastYearCount']), toCount('b', $Pro['PatentIssuance']['TaiwanIncasesale'], $Pro['PatentIssuance']['TaiwanLastYearIncasesale']), 
    ],
    ['', 
    '國外件數', toCount('a', $Pro['PatentIssuance']['ForeignCount']),
    '國外服務費', toCount('b', $Pro['PatentIssuance']['ForeignIncasesale']), '去年同期', toCount('a', $Pro['PatentIssuance']['ForeignCount'], $Pro['PatentIssuance']['ForeignLastYearCount']), toCount('b', $Pro['PatentIssuance']['ForeignIncasesale'], $Pro['PatentIssuance']['ForeignLastYearIncasesale'])
    ],
    ['商標領證', 
    '台灣件數',  toCount('a', $Pro['TrademarkIssuance']['TaiwanCount']),
    '台灣服務費', toCount('b', $Pro['TrademarkIssuance']['TaiwanIncasesale']), '去年同期', toCount('a', $Pro['TrademarkIssuance']['TaiwanCount'], $Pro['TrademarkIssuance']['TaiwanLastYearCount']), toCount('b', $Pro['TrademarkIssuance']['TaiwanIncasesale'], $Pro['TrademarkIssuance']['TaiwanLastYearIncasesale']), 
    ],
    ['', 
    '國外件數', toCount('a', $Pro['TrademarkIssuance']['ForeignCount']),
    '國外服務費', toCount('b', $Pro['TrademarkIssuance']['ForeignIncasesale']), '去年同期', toCount('a', $Pro['TrademarkIssuance']['ForeignCount'], $Pro['TrademarkIssuance']['ForeignLastYearCount']), toCount('b', $Pro['TrademarkIssuance']['ForeignIncasesale'], $Pro['TrademarkIssuance']['ForeignLastYearIncasesale'])
    ],
    ['商標延展', 
    '台灣件數',  toCount('a', $Pro['TrademarkExtend']['TaiwanCount']),
    '台灣服務費', toCount('b', $Pro['TrademarkExtend']['TaiwanIncasesale']), '去年同期', toCount('a', $Pro['TrademarkExtend']['TaiwanCount'], $Pro['TrademarkExtend']['TaiwanLastYearCount']), toCount('b', $Pro['TrademarkExtend']['TaiwanIncasesale'], $Pro['TrademarkExtend']['TaiwanLastYearIncasesale']),
    ],
    ['', 
    '國外件數', toCount('a', $Pro['TrademarkExtend']['ForeignCount']),
    '國外服務費', toCount('b', $Pro['TrademarkExtend']['ForeignIncasesale']), '去年同期', toCount('a', $Pro['TrademarkExtend']['ForeignCount'], $Pro['TrademarkExtend']['ForeignLastYearCount']), toCount('b', $Pro['TrademarkExtend']['ForeignIncasesale'], $Pro['TrademarkExtend']['ForeignLastYearIncasesale'])
    ],
];

if ($_GET["DeptName"] === '國際部') {
    $Info[] = ['植物品種權', 
    '台灣件數',  toCount('a', $Pro['PlantVariety']['TaiwanCount']),
    '台灣服務費', toCount('b', $Pro['PlantVariety']['TaiwanIncasesale']), '去年同期', toCount('a', $Pro['PlantVariety']['TaiwanCount'], $Pro['PlantVariety']['TaiwanLastYearCount']), toCount('b', $Pro['PlantVariety']['TaiwanIncasesale'], $Pro['PlantVariety']['TaiwanLastYearIncasesale']),
    ];
    $Info[] = ['', 
    '國外件數', toCount('a', $Pro['PlantVariety']['ForeignCount']),
    '國外服務費', toCount('b', $Pro['PlantVariety']['ForeignIncasesale']), '去年同期', toCount('a', $Pro['PlantVariety']['ForeignCount'], $Pro['PlantVariety']['ForeignLastYearCount']), toCount('b', $Pro['PlantVariety']['ForeignIncasesale'], $Pro['PlantVariety']['ForeignLastYearIncasesale'])
    ];
}

renderPerformanceTableFixed($phpWord, $section, 'ProjectStatus', '開案狀況', $w, 8, $Info);
$section->addTextBreak(1);

// ===== 5) 國外專利新申請案： ===== 
$ForPa = $Data['ForeignNewCase']['Patent'];
$Info = NewForeignApplicationCaseSort($ForPa);

if (count($Info) === 0) {
    $Info = [['查無資料'], ['0件']];
}
renderPerformanceTableFixed($phpWord, $section, 'ForeignPatentNewCase', '國外專利新申請案', $w, count($Info[0]), $Info);
$section->addTextBreak(1);

// ===== 6) 國外商標新申請案： =====
$ForTr = $Data['ForeignNewCase']['Trademark'];
$Info = NewForeignApplicationCaseSort($ForTr);

if (count($Info) === 0) {
    $Info = [['查無資料'], ['0件']];
}
renderPerformanceTableFixed($phpWord, $section, 'ForeignTrademarkNewCase', '國外商標新申請案', $w, count($Info[0]), $Info);
$section->addTextBreak(1);

// ===== 7) 新客戶： =====
$NewClient = $Data['Performance'];
$Info = [
    ['業績', number_format($NewClient['NewClientSale']), '服務費', number_format($NewClient['NewIncasesale'])],
];
renderPerformanceTableFixed($phpWord, $section, 'NewClientSale', '新客戶', $w, 4, $Info);
$section->addTextBreak(1);

// ===== 8) 截至目前未收款金額： =====
$Rcv = $Data['Performance']['RcvbleBalance'];
$Abn = $Data['Performance']['AbnormalIncasesale'];

$phpWord->addTableStyle('RcvbleBalance', [
    'borderSize' => 6,
    'borderColor'=> '000000',
    'cellMargin' => 60,
    'alignment' => \PhpOffice\PhpWord\SimpleType\JcTable::END
]);

// 建立表格
$tbl = $section->addTable('RcvbleBalance');

// 表頭（固定樣式）
$row  = $tbl->addRow(null);
$cell = $row->addCell($w, ['gridSpan' => 1, 'bgColor' => 'D9D9D9', 'valign' => 'center']);
$cell->addText('截至目前未收款金額', ['size' => 12,], ['alignment' => Jc::CENTER, 'spaceBefore' => 0, 'spaceAfter'  => 0, 'lineHeight'  => 1.0]);

$row = $tbl->addRow(null);
$row->addCell(1, ['valign' => 'center'])->addText((string)number_format($Rcv), ['size' => 12, ], ['alignment' => Jc::CENTER, 'spaceBefore' => 0, 'spaceAfter'  => 0, 'lineHeight'  => 1.0]);

$row  = $tbl->addRow(null);
$cell = $row->addCell($w, ['gridSpan' => 1, 'bgColor' => 'D9D9D9', 'valign' => 'center']);

$AbnormalDay = '90';
if ($_GET["DeptName"] === '國際部') {
    $AbnormalDay = '150';
}

$cell->addText($AbnormalDay.'天異常收款金額', ['size' => 12,], ['alignment' => Jc::CENTER, 'spaceBefore' => 0, 'spaceAfter'  => 0, 'lineHeight'  => 1.0]);

$row = $tbl->addRow(null);
$row->addCell(1, ['valign' => 'center'])->addText((string)number_format($Abn), ['size' => 12, ], ['alignment' => Jc::CENTER, 'spaceBefore' => 0, 'spaceAfter'  => 0, 'lineHeight'  => 1.0]);
$section->addTextBreak(1);

// ===== 9) 銷案件數與金額： =====

$phpWord->addTableStyle('NumberAndAmountOfSales', [
    'borderSize' => 6,
    'borderColor'=> '000000',
    'cellMargin' => 60,
    'alignment' => \PhpOffice\PhpWord\SimpleType\JcTable::END
]);
// 建立表格
$tbl = $section->addTable('NumberAndAmountOfSales');

// 表頭（固定樣式）
$row  = $tbl->addRow(400);
$cell = $row->addCell($w, ['gridSpan' => 1, 'bgColor' => 'D9D9D9', 'valign' => 'center']);
$cell->addText('銷案件數與金額', ['size' => 12,], ['alignment' => Jc::CENTER, 'spaceBefore' => 0, 'spaceAfter'  => 0, 'lineHeight'  => 1.0]);

$row = $tbl->addRow(1000);
$row->addCell(1, [])->addText('(自行填寫)', ['size' => 12], []);
$section->addTextBreak(1);

// ===== 10) 本月(季)(年)業績前十名客戶： / 若國際部顯示 本月(季)(年)A級客戶： =====
$SaleTop = [];
$Info = [];
$title = '';
$num = 0;

if ($_GET["DeptName"] === '國際部') {
    $SaleTop = $Data['InternationalClientList'];
    $Info = [['', '名稱', '專利數', '商標數', '植物品種權', '業績', '服務費']];
    $title = '本月(季)(年)A級客戶';
    $num = 7;

    if ($SaleTop !== null) {
        foreach ($SaleTop as $i => $r) {
            $Info[] = [$i + 1, $r['CustomerCName'], $r['PatentCount'], $r['TrademarkCount'], $r['PlantVarietyCount'], number_format($r['Sale']), number_format($r['Incasesale'])];
        }
    }

    if (count($Info) === 1) {
        $Info[] = ['', '查無資料', 0, 0, 0, 0, 0];
    }

} else {
    $SaleTop = $Data['SaleTopList'];
    $Info = [['', '名稱', '專利數', '商標數', '業績', '服務費']];
    $title = '本月(季)(年)業績前十名客戶';
    $num = 6;

    if ($SaleTop !== null) {
        foreach ($SaleTop as $i => $r) {
            $Info[] = [$i + 1, $r['CustomerCName'], $r['PatentCount'], $r['TrademarkCount'], number_format($r['Sale']), number_format($r['Incasesale'])];
        }
    }

    if (count($Info) === 1) {
        $Info[] = ['', '查無資料', 0, 0, 0, 0];
    }
}

renderPerformanceTableFixed($phpWord, $section, 'SaleTopList', $title, $w, $num, $Info);
$section->addTextBreak(1);

// ===== 11) S/A級客戶狀況： / 若國際部顯示 本月(季)(年)A級代理人： =====

if ($_GET["DeptName"] === '國際部') {
    $SaleTop = $Data['InternationalForSalesList'];
    $Info = [['', '名稱', '專利數', '商標數', '植物品種權', '業績', '服務費']];
    $title = '本月(季)(年)A級代理人';
    $num = 7;

    if ($SaleTop !== null) {
        foreach ($SaleTop as $i => $r) {
            $Info[] = [$i + 1, $r['ForSales'], $r['PatentCount'], $r['TrademarkCount'], $r['PlantVarietyCount'], number_format($r['Sale']), number_format($r['Incasesale'])];
        }
    }

    if (count($Info) === 1) {
        $Info[] = ['', '查無資料', 0, 0, 0, 0, 0];
    }

    renderPerformanceTableFixed($phpWord, $section, 'CustomerLevelList', $title, $w, $num, $Info);
} else {
    $phpWord->addTableStyle('CustomerLevelList', [
        'borderSize' => 6,
        'borderColor'=> '000000',
        'cellMargin' => 60,
        'alignment' => \PhpOffice\PhpWord\SimpleType\JcTable::END
    ]);
    // 建立表格
    $tbl = $section->addTable('CustomerLevelList');


    // 表頭（固定樣式）
    $row  = $tbl->addRow(null);
    $cell = $row->addCell($w, ['gridSpan' => 8, 'bgColor' => 'D9D9D9', 'valign' => 'center']);
    $cell->addText('S/A級客戶狀況', ['size' => 12,], ['alignment' => Jc::CENTER, 'spaceBefore' => 0, 'spaceAfter'  => 0, 'lineHeight'  => 1.0]);

    $Customer = [];
    if ($Data['CustomerLevelList'] !== null) {
        $Customer = $Data['CustomerLevelList'];
    }
    $TextList = ['今年業績', '去年業績', '今年服務費', '去年服務費', '今年專利', '去年專利', '今年商標', '去年商標'];
    $Fields = ['Incasesale', 'IncasesaleLastYear', 'PatentCount', 'PatentLastYearCount', 'Sale', 'SaleLastYear', 'TrademarkCount', 'TrademarkLastYearCount'];
    foreach ($Customer as $r) {
        $row = $tbl->addRow(null);
        $row->addCell($w, ['gridSpan' => 8, 'valign' => 'center'])
            ->addText($r['CustomerCName'], ['size' => 12,], ['alignment' => Jc::CENTER, 'spaceBefore' => 0, 'spaceAfter'  => 0, 'lineHeight'  => 1.0]);
        $row = $tbl->addRow(null);
        foreach ($TextList as $t) {
            $row->addCell($w/8, ['valign' => 'center'])
                ->addText($t, ['size' => 12,], ['alignment' => Jc::CENTER, 'spaceBefore' => 0, 'spaceAfter'  => 0, 'lineHeight'  => 1.0]);
        }
        $row = $tbl->addRow(null);
        foreach ($Fields as $f) {
            $row->addCell($w/8, ['valign' => 'center'])
                ->addText((string)number_format($r[$f]), ['size' => 12,], ['alignment' => Jc::CENTER, 'spaceBefore' => 0, 'spaceAfter'  => 0, 'lineHeight'  => 1.0]);
        }
    }

    if (count($Customer) === 0) {
        $row = $tbl->addRow(null);
        $row->addCell($w, ['gridSpan' => 8, 'valign' => 'center'])
            ->addText('查無資料', ['size' => 12,], ['alignment' => Jc::CENTER, 'spaceBefore' => 0, 'spaceAfter'  => 0, 'lineHeight'  => 1.0]);
    }
}



if($data['Code'] === 200) {
    // 檔名
    if ( $_GET["type"] === 'Year') {
        $fileName = '業績報告-' . $_GET["name"] . $_GET["year"] . '年' . date('YmdHis');
    } else {
        $fileName = '業績報告-' . $_GET["name"] . $_GET["year"] . '年' . $_GET["number"] . (  $_GET["type"] === 'Season' ? '季' : '月'). date('YmdHis') ;
    }
    
	// Save file
	// echo write($phpWord, basename(__FILE__, '.php'), $writers);
    date_default_timezone_set('Asia/Taipei');
	write($phpWord, $fileName, $writers);
	// if (!CLI) {
	// 	include_once 'Sample_Footer.php';
	// }


	echo "<script language='javascript' type ='text/javascript'>"; 
	echo "window.location.href = 'results/" . $fileName . ".docx';";
	echo 'document.getElementById("print").innerHTML = "檔案已下載完成。";';
	echo "</script>"; 
} else {
	echo "<script language='javascript' type ='text/javascript'>"; 

    $text = 'document.getElementById("print").innerHTML = "檔案下載失敗。';
    $text.= '<br> 未取得資料API錯誤。 <br>';
    
    echo $text . '";';
	echo "</script>"; 
}