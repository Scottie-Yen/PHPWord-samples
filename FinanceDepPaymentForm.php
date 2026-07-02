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

// 寬度百分比計算 (總寬度基準為 11000 TWIP)
function width($percent) {
    return $percent / 100 * 11000;
}

// 計算約計
function Approximately ($fee, $rate) {
    return number_format((float) round((float)$fee * (float)$rate));
}

// 計算營業稅5%
function BusinessTax($fee) {
    return number_format((float) round((float)$fee * 0.05));
}

// 計算合計
function Total($fee, $rate, $inCaseSale2) {
    return number_format((float) round((float)$fee * (float)$rate) + (float)$inCaseSale2 + (float) round((float)$inCaseSale2 * 0.05));
}


use PhpOffice\PhpWord\Element\Section;
use PhpOffice\PhpWord\Element\Table;
use PhpOffice\PhpWord\SimpleType\JcTable;
use PhpOffice\PhpWord\SimpleType\Jc;
use PhpOffice\PhpWord\Shared\Converter;

// POST
$data='';

if(isset($_GET["rcvAcntNum"]) && isset($_GET["type"])){
    $data = array(
        "rcvAcntNum" => $_GET["rcvAcntNum"], // 應收款單號
        "type" => $_GET["type"],             // 案件類型
    );
}else{
    echo "Error : 參數錯誤";
}

// 參數接收
$rcvAcntNum = $_GET["rcvAcntNum"];      // 應收款單號
$documentType = $_GET["documentType"];  // 單據類型
$currency = $_GET["ReplaceCurrency"];   // 幣別
$amount = $_GET["ReplaceAmount"];       // 國內案:合記 國外案(合):代收代付費用約計
$rate = $_GET["ReplaceExchangeRate"];   // 匯率
$rateDate = $_GET["ReplaceAmountDate"]; // 匯率日期
$rcvbleDue = $_GET["RcvbleDue"];  // 請款日期
$inCaseSale2 = $_GET["InCaseSale2"]; // 本所服務費

// 國外案(分):規費外幣金額
if(isset($_GET["ForeignCurrencyAmount"])){
    $foreignCurrencyAmount = $_GET["ForeignCurrencyAmount"];
}
// 國外案(分):代理人費用外幣金額
if(isset($_GET["ReplaceAmount"])){
    $replaceAmount = $_GET["ReplaceAmount"];
}

$host = $_SERVER['HTTP_HOST']; 
if ($host === '192.168.0.52' || $host === '192.168.0.52:8080') {
    $data = httpRequest('http://192.168.0.52:8000/report/select-generatepaymentrequest', json_encode($data));
} elseif ($host === '192.168.0.188' || $host === '192.168.0.188:8080') {
    $data = httpRequest('http://192.168.0.186:8000/report/select-generatepaymentrequest', json_encode($data));
}
$Data = $data['Data'];

// New Word document
include_once 'Sample_Header.php';
$phpWord = new \PhpOffice\PhpWord\PhpWord();

// New portrait section
$section = $phpWord->addSection(array(
    'headerHeight' => 300,
    'marginTop' => 2000,
    // 🌟 修正：給予微小的左右邊界 (約0.7cm)，防止 WPS 或 Mac Pages 強制縮排導致跑版
    'marginLeft' => 400,
    'marginRight' => 400,
    'footerHeight' => 0,
));

// ==========================================
// 頁首 (Header) 設定
// ==========================================
$header = $section->addHeader();

// 🌟 終極解法：使用「絕對定位」，直接對齊「實體紙張邊緣」，無視任何 Margin 限制！
$headerImageStyle = array(
    'width'            => 600,  // PhpWord 中給定寬度即會等比例縮放
    'positioning'      => 'absolute', // 啟用絕對定位
    'posHorizontal'    => 'left',     // 靠左
    'posHorizontalRel' => 'page',     // 基準點：整張實體紙張 (無視左右邊界)
    'posVertical'      => 'top',      // 靠上
    'posVerticalRel'   => 'page',     // 基準點：整張實體紙張 (無視上下邊界)
);

if ($documentType === 'domestic') {
    if($_GET["type"] === 'Patent') {
        $header->addImage('images/巨群信頭_專利請款單.png', $headerImageStyle);
    }
    if($_GET["type"] === 'TM') {
        $header->addImage('images/巨群信頭_商標請款單.png', $headerImageStyle);
    }
} else if ($documentType !== 'domestic') {
    if($_GET["type"] === 'Patent') {
        $header->addImage('images/巨群信頭_智財諮詢專利請款單.png', $headerImageStyle);
    }
    if($_GET["type"] === 'TM') {
        $header->addImage('images/巨群信頭_智財諮詢商標請款單.png', $headerImageStyle);
    }
} else {
    $header->addImage('images/巨群信頭_國外請款單.png', $headerImageStyle);
}

// ==========================================
// 頁尾 (Footer) 設定
// ==========================================
$footer = $section->addFooter();

// 🌟 頁尾一樣使用絕對定位，對齊紙張的「左下角」
$footerImageStyle = array(
    'width'            => 600, 
    'positioning'      => 'absolute',
    'posHorizontal'    => 'left',
    'posHorizontalRel' => 'page', // 基準點：整張實體紙張的左邊
    'posVertical'      => 'bottom', // 靠下
    'posVerticalRel'   => 'page', // 基準點：整張實體紙張的最底端
);

if ($documentType === 'domestic') {
    if ($_GET["type"] === 'Patent') {
        $footer->addImage('images/巨群信頭_巨群智財諮詢 Confidential 永豐1.png', $footerImageStyle);
    }
    if ($_GET["type"] === 'TM') {
        $footer->addImage('images/巨群信頭_巨群智財諮詢 Confidential 永豐2.png', $footerImageStyle);
    }
} else {
    $footer->addImage('images/巨群信頭_巨群智財諮詢 Confidential 富邦.png', $footerImageStyle);
}

// ===== 1) 文件標題 =====
$section->addText(
    '請款單',
    ['name' => 'Microsoft JhengHei', 'size' => 26, 'bold' => true, 'underline' => 'double'],
    ['alignment' => Jc::CENTER, 'spaceAfter' => 0]
);

// ===== 2) 客戶資訊 =====
$section->addText(
    $Data['CustomerCName'],
    ['name' => 'Microsoft JhengHei', 'size' => 14, 'bold' => true],
    ['alignment' => Jc::START, 'spaceAfter' => 0, 'indentation' => [ 'left'  => Converter::cmToTwip(1)], 'lineHeight' => 0.7]
);
$section->addText(
    $Data['ContactAddress'],
    ['name' => 'Microsoft JhengHei', 'size' => 14, 'bold' => true],
    ['alignment' => Jc::START, 'spaceAfter' => 100, 'indentation' => [ 'left'  => Converter::cmToTwip(1)], 'lineHeight' => 0.7]
);

$titleStyle = ['name' => 'Microsoft JhengHei', 'size' => 10, 'color' => '000000', 'bold' => true];
$fontStyle = ['name' => 'Microsoft JhengHei', 'size' => 10, 'color' => '000000'];
$CLeft = ['alignment' => \PhpOffice\PhpWord\SimpleType\Jc::START, 'spaceAfter' => 0, 'lineHeight' => 1.0, 'indentation' => [ 'left'  => Converter::cmToTwip(1)]];

// 🌟 修正：基本資訊表格，單位改為 WIDTH_TWIP，總寬度 11000
$infoTable = $section->addTable([
    'borderSize' => 0,
    'borderColor' => 'FFFFFF',
    'cellMargin' => 0,
    'unit' => \PhpOffice\PhpWord\Style\Table::WIDTH_TWIP, // 改用絕對寬度 TWIP
    'width' => 11000 // 總長度 11000
]);
$infoTable->addRow();
$infoTable->addCell(5000)->addText('貴司編號：'.$Data['CltFileID'], $fontStyle, $CLeft);
$infoTable->addCell(5000)->addText('應收款單號：'.$rcvAcntNum, $fontStyle, $CLeft);

$infoTable->addRow();
$infoTable->addCell(5000)->addText('本所編號：'.$Data['FileId'], $fontStyle, $CLeft);
$infoTable->addCell(5000)->addText('請款日期：'.$rcvbleDue, $fontStyle, $CLeft);

$infoTable->addRow();
$infoTable->addCell(9000)->addText('申請案號：'.$Data['AppCaseld'], $fontStyle, $CLeft);

$subtitle = '';
if ($_GET["type"] === 'Patent') {
 $subtitle = '專利名稱：';
} else if ($_GET["type"] === 'TM') {
 $subtitle = '商標名稱：';
}
$infoTable->addRow();
$infoTable->addCell(11000, ['gridSpan' => 2])->addText($subtitle.$Data['CTitle'], $fontStyle, ['alignment' => \PhpOffice\PhpWord\SimpleType\Jc::START, 'spaceAfter' => 100, 'lineHeight' => 1.0, 'indentation' => [ 'left'  => Converter::cmToTwip(1)]]);

// ===== 3) 表格資訊 =====
// 🌟 修正：主表格樣式，單位改為 WIDTH_TWIP，總寬度 11000
$tableStyle = [
    'borderSize' => 6, 
    'borderColor' => '000000', 
    'cellMargin' => 80,
    'unit' => \PhpOffice\PhpWord\Style\Table::WIDTH_TWIP, // 改用絕對寬度 TWIP
    'width' => 11000, // 總長度 11000
    'alignment' => Jc::CENTER
];
$TCenter = ['alignment' => Jc::CENTER, 'spaceAfter' => 0, 'lineHeight' => 1.0];
$TLeft = ['alignment' => \PhpOffice\PhpWord\SimpleType\Jc::START, 'spaceAfter' => 0, 'lineHeight' => 1.0];
$TRight = ['alignment' => \PhpOffice\PhpWord\SimpleType\Jc::END, 'spaceAfter' => 0, 'lineHeight' => 1.0];

$hideRightLine = ['borderRightColor' => 'FFFFFF', 'borderRightSize' => 0];
$hideLeftLine = ['borderLeftColor' => 'FFFFFF', 'borderLeftSize' => 0];
$hideTitleLine = ['borderLeftColor' => 'FFFFFF', 'borderLeftSize' => 0, 'borderBottomColor' => 'FFFFFF', 'borderBottomSize' => 0, 'borderTopColor' => 'FFFFFF', 'borderTopSize' => 0];

$phpWord->addTableStyle('InvoiceTable', $tableStyle);
// 2. 建立表格
$table = $section->addTable('InvoiceTable');

if ($documentType === 'domestic') {
    // 國內案
    $table->addRow();
    $table->addCell(width(65), ['gridSpan' => 2])->addText('項目', $titleStyle, $TCenter);
    $table->addCell(width(20), ['gridSpan' => 2])->addText('規費', $titleStyle, $TCenter);
    $table->addCell(width(15), ['gridSpan' => 2])->addText('服務費', $titleStyle, $TCenter);

    $table->addRow();
    $table->addCell(width(65), ['gridSpan' => 2])->addText($Data['WhatFor'], $fontStyle, $TLeft);
    $table->addCell(width(5), $hideRightLine)->addText('NTD', $fontStyle, $TLeft);
    $table->addCell(width(15), $hideLeftLine)->addText(number_format($Data['OfficialFee']), $fontStyle, $TRight);
    $table->addCell(width(5), $hideRightLine)->addText('NTD', $fontStyle, $TLeft);
    $table->addCell(width(10), $hideLeftLine)->addText(number_format($inCaseSale2), $fontStyle, $TRight);

    $table->addRow();
    $table->addCell(width(53), ['borderLeftColor' => 'FFFFFF', 'borderLeftSize' => 0, 'borderBottomColor' => 'FFFFFF', 'borderBottomSize' => 0])->addText('');
    $table->addCell(width(12))->addText('小計', $fontStyle, $TCenter);
    $table->addCell(width(5), $hideRightLine)->addText('NTD', $fontStyle, $TLeft);
    $table->addCell(width(15), $hideLeftLine)->addText(number_format($Data['OfficialFee']), $fontStyle, $TRight);
    $table->addCell(width(5), $hideRightLine)->addText('NTD', $fontStyle, $TLeft);
    $table->addCell(width(10), $hideLeftLine)->addText(number_format($inCaseSale2), $fontStyle, $TRight);

    $table->addRow();
    $table->addCell(width(53), $hideTitleLine)->addText('');
    $table->addCell(width(12))->addText('合計', $fontStyle, $TCenter);
    $table->addCell(width(20), ['gridSpan' => 2, 'borderRightColor' => 'FFFFFF', 'borderRightSize' => 0])->addText('NTD', $fontStyle, $TLeft);
    $table->addCell(width(15), ['gridSpan' => 2, 'borderLeftColor' => 'FFFFFF', 'borderLeftSize' => 0])->addText(number_format($Data['OfficialFee'] + $inCaseSale2), $fontStyle, $TRight);

} else if ($documentType === 'foreignJoin') {
    // 國外案合
    $table->addRow();
    $table->addCell(width(65), ['gridSpan' => 2])->addText('項目', $titleStyle, $TCenter);
    $table->addCell(width(20), ['gridSpan' => 2])->addText('代收代付費用', $titleStyle, $TCenter);
    $table->addCell(width(15), ['gridSpan' => 2])->addText('本所服務費', $titleStyle, $TCenter);

    $table->addRow();
    $table->addCell(width(65), ['gridSpan' => 2])->addText($Data['WhatFor'], $fontStyle, $TLeft);
    $table->addCell(width(5), $hideRightLine)->addText($currency, $fontStyle, $TLeft);
    $table->addCell(width(15), $hideLeftLine)->addText((number_format((float)$amount)), $fontStyle, $TRight);
    $table->addCell(width(5), $hideRightLine)->addText('NTD', $fontStyle, $TLeft);
    $table->addCell(width(10), $hideLeftLine)->addText(number_format($inCaseSale2), $fontStyle, $TRight);

    $table->addRow();
    $table->addCell(width(53), ['borderLeftColor' => 'FFFFFF', 'borderLeftSize' => 0, 'borderBottomColor' => 'FFFFFF', 'borderBottomSize' => 0])->addText('');
    $table->addCell(width(12))->addText('小計', $fontStyle, $TCenter);
    $table->addCell(width(5), $hideRightLine)->addText($currency, $fontStyle, $TLeft);
    $table->addCell(width(15), $hideLeftLine)->addText((number_format((float)$amount)), $fontStyle, $TRight);
    $table->addCell(width(5), $hideRightLine)->addText('NT$', $fontStyle, $TLeft);
    $table->addCell(width(10), $hideLeftLine)->addText(number_format($inCaseSale2), $fontStyle, $TRight);

    $table->addRow();
    $table->addCell(width(53), $hideTitleLine)->addText('');
    $table->addCell(width(12))->addText('匯率', $fontStyle, $TCenter);
    $table->addCell(width(9))->addText($rate, $fontStyle, $TRight);
    $table->addCell(width(11))->addText('('. $rateDate .')', $fontStyle, $TLeft);
    $table->addCell(width(15), ['gridSpan' => 2])->addText('----', $fontStyle, $TCenter);

    $table->addRow();
    $table->addCell(width(53), $hideTitleLine)->addText('');
    $table->addCell(width(12))->addText('約計', $fontStyle, $TCenter);
    $table->addCell(width(5), $hideRightLine)->addText('NTD', $fontStyle, $TLeft);
    $table->addCell(width(15), $hideLeftLine)->addText(Approximately((float)$amount, (float)$rate), $fontStyle, $TRight);
    $table->addCell(width(5), $hideRightLine)->addText('NTD', $fontStyle, $TLeft);
    $table->addCell(width(10), $hideLeftLine)->addText(number_format($inCaseSale2), $fontStyle, $TRight);

    $table->addRow();
    $table->addCell(width(53), $hideTitleLine)->addText('');
    $table->addCell(width(12))->addText('營業稅5%', $fontStyle, $TCenter);
    $table->addCell(width(20), ['gridSpan' => 2])->addText('----', $fontStyle, $TCenter);
    $table->addCell(width(5), $hideRightLine)->addText('NTD', $fontStyle, $TLeft);
    $table->addCell(width(10), $hideLeftLine)->addText(BusinessTax($inCaseSale2), $fontStyle, $TRight);

    $table->addRow();
    $table->addCell(width(53), $hideTitleLine)->addText('');
    $table->addCell(width(12))->addText('合計', $fontStyle, $TCenter);
    $table->addCell(width(20), ['gridSpan' => 2, 'borderRightColor' => 'FFFFFF', 'borderRightSize' => 0])->addText('NTD', $fontStyle, $TLeft);
    $table->addCell(width(15), ['gridSpan' => 2, 'borderLeftColor' => 'FFFFFF', 'borderLeftSize' => 0])->addText(Total($amount, $rate, $inCaseSale2), $fontStyle, $TRight);

} else {
    // 國外案分
    $table->addRow();
    $table->addCell(width(27.5))->addText('項目', $titleStyle, $TCenter);
    $table->addCell(width(27.5), ['gridSpan' => 2])->addText('說明', $titleStyle, $TCenter);
    $table->addCell(width(15), ['gridSpan' => 2])->addText('規費', $titleStyle, $TCenter);
    $table->addCell(width(15), ['gridSpan' => 2])->addText('代理人費用', $titleStyle, $TCenter);
    $table->addCell(width(15), ['gridSpan' => 2])->addText('本所服務費', $titleStyle, $TCenter);

    $table->addRow();
    $table->addCell(width(27.5))->addText($Data['WhatFor'], $fontStyle, $TLeft);
    $table->addCell(width(27.5), ['gridSpan' => 2])->addText('', $fontStyle, $TLeft);
    $table->addCell(width(5), $hideRightLine)->addText($currency, $fontStyle, $TLeft);
    $table->addCell(width(10), $hideLeftLine)->addText(number_format($foreignCurrencyAmount), $fontStyle, $TRight);
    $table->addCell(width(5), $hideRightLine)->addText($currency, $fontStyle, $TLeft);
    $table->addCell(width(10), $hideLeftLine)->addText(number_format($replaceAmount), $fontStyle, $TRight);
    $table->addCell(width(5), $hideRightLine)->addText('NTD', $fontStyle, $TLeft);
    $table->addCell(width(10), $hideLeftLine)->addText((number_format($inCaseSale2)), $fontStyle, $TRight);

    $table->addRow();
    $table->addCell(width(27.5), ['borderSize' => 0, 'borderColor' => 'FFFFFF'])->addText('');
    $table->addCell(width(15.5), ['borderSize' => 0, 'borderColor' => 'FFFFFF'])->addText('');
    $table->addCell(width(12))->addText('小計', $fontStyle, $TCenter);
    $table->addCell(width(5), $hideRightLine)->addText($currency, $fontStyle, $TLeft);
    $table->addCell(width(25), ['borderLeftColor' => 'FFFFFF', 'borderLeftSize' => 0, 'gridSpan' => 3])->addText(number_format($foreignCurrencyAmount+$replaceAmount), $fontStyle, $TRight);
    $table->addCell(width(5), $hideRightLine)->addText('NTD', $fontStyle, $TLeft);
    $table->addCell(width(10), $hideLeftLine)->addText((number_format($inCaseSale2)), $fontStyle, $TRight);

    $table->addRow();
    $table->addCell(width(27.5), ['borderSize' => 0, 'borderColor' => 'FFFFFF'])->addText('');
    $table->addCell(width(15.5), ['borderSize' => 0, 'borderColor' => 'FFFFFF'])->addText('');
    $table->addCell(width(12))->addText('匯率', $fontStyle, $TCenter);
    $table->addCell(width(12))->addText($rate, $fontStyle, $TRight);
    $table->addCell(width(18), ['gridSpan' => 3])->addText('('. $rateDate .')', $fontStyle, $TLeft);
    $table->addCell(width(15), ['borderLeftColor' => 'FFFFFF', 'borderLeftSize' => 0, 'gridSpan' => 2])->addText('1.00', $fontStyle, $TCenter);

    $table->addRow();
    $table->addCell(width(27.5), ['borderSize' => 0, 'borderColor' => 'FFFFFF'])->addText('');
    $table->addCell(width(15.5), ['borderSize' => 0, 'borderColor' => 'FFFFFF'])->addText('');
    $table->addCell(width(12))->addText('約計', $fontStyle, $TCenter);
    $table->addCell(width(5), $hideRightLine)->addText('NTD', $fontStyle, $TLeft);
    $table->addCell(width(25), ['borderLeftColor' => 'FFFFFF', 'borderLeftSize' => 0, 'gridSpan' => 3])->addText(Approximately(((float)$foreignCurrencyAmount+(float)$replaceAmount), (float)$rate), $fontStyle, $TRight);
    $table->addCell(width(5), $hideRightLine)->addText('NTD', $fontStyle, $TLeft);
    $table->addCell(width(10), $hideLeftLine)->addText(number_format($inCaseSale2), $fontStyle, $TRight);

    $table->addRow();
    $table->addCell(width(27.5), ['borderSize' => 0, 'borderColor' => 'FFFFFF'])->addText('');
    $table->addCell(width(15.5), ['borderSize' => 0, 'borderColor' => 'FFFFFF'])->addText('');
    $table->addCell(width(12))->addText('營業稅5%', $fontStyle, $TCenter);
    $table->addCell(width(30), ['gridSpan' => 4])->addText('----', $fontStyle, $TCenter);
    $table->addCell(width(5), $hideRightLine)->addText('NTD', $fontStyle, $TLeft);
    $table->addCell(width(10), $hideLeftLine)->addText(BusinessTax($inCaseSale2), $fontStyle, $TRight);

    $table->addRow();
    $table->addCell(width(27.5), ['borderSize' => 0, 'borderColor' => 'FFFFFF'])->addText('');
    $table->addCell(width(15.5), ['borderSize' => 0, 'borderColor' => 'FFFFFF'])->addText('');
    $table->addCell(width(12))->addText('合計', $fontStyle, $TCenter);
    $table->addCell(width(5), $hideRightLine)->addText('NTD', $fontStyle, $TLeft);
    $table->addCell(width(40), ['gridSpan' => 2, 'borderLeftColor' => 'FFFFFF', 'borderLeftSize' => 0, 'gridSpan' => 5])->addText(Total($foreignCurrencyAmount+$replaceAmount, $rate, $inCaseSale2), $fontStyle, $TRight);
}

// ----------------------------------------------------
// 第 2 頁：代收代付國外收據 (只有國外案顯示)
// ----------------------------------------------------
if ($documentType !== 'domestic') {
    $section->addPageBreak();
    
    // ===== 1) 文件標題 =====
    $section->addText(
        '代收代付國外收據',
        ['name' => 'Microsoft JhengHei', 'size' => 18,],
        ['alignment' => Jc::CENTER, 'spaceAfter' => 0, 'lineHeight' => 1]
    );
    $section->addText(
        '(正本)',
        ['name' => 'Microsoft JhengHei', 'size' => 14,],
        ['alignment' => Jc::CENTER, 'spaceAfter' => 0, 'lineHeight' => 0.7]
    );

    // ===== 2) 客戶資訊 =====
    $section->addText(
        $Data['CustomerCName'],
        ['name' => 'Microsoft JhengHei', 'size' => 14],
        ['alignment' => Jc::START, 'spaceAfter' => 0, 'indentation' => [ 'left'  => Converter::cmToTwip(1)], 'lineHeight' => 0.7]
    );

    // 🌟 修正：收據的客戶資訊表，單位改為 WIDTH_TWIP，總寬度 11000
    $infoTable = $section->addTable([
        'borderSize' => 0,
        'borderColor' => 'FFFFFF',
        'cellMargin' => 0,
        'unit' => \PhpOffice\PhpWord\Style\Table::WIDTH_TWIP, // 改用絕對寬度 TWIP
        'width' => 11000 // 總長度 11000
    ]);

    $DLeft = ['alignment' => \PhpOffice\PhpWord\SimpleType\Jc::START, 'spaceAfter' => 0, 'lineHeight' => 1, 'indentation' => [ 'left'  => Converter::cmToTwip(1)]];

    $infoTable->addRow();
    $infoTable->addCell(width(60))->addText('', $fontStyle, $DLeft);
    $infoTable->addCell(width(10))->addText('客戶編號', $fontStyle, ['alignment' => 'distribute', 'lineHeight' => 1, 'spaceAfter' => 0]);
    $infoTable->addCell(width(0.5))->addText('：', $fontStyle, ['lineHeight' => 1, 'spaceAfter' => 0]);
    $infoTable->addCell(width(25))->addText($Data['CltFileID'], $fontStyle, ['alignment' => \PhpOffice\PhpWord\SimpleType\Jc::START, 'spaceAfter' => 0, 'lineHeight' => 1]);

    $infoTable->addRow();
    $infoTable->addCell(width(60))->addText('', $fontStyle, $DLeft);
    $infoTable->addCell(width(10))->addText('收據單號', $fontStyle, ['alignment' => 'distribute', 'lineHeight' => 1, 'spaceAfter' => 0]);
    $infoTable->addCell(width(0.5))->addText('：', $fontStyle, ['lineHeight' => 1, 'spaceAfter' => 0]);
    $infoTable->addCell(width(25))->addText($rcvAcntNum, $fontStyle, ['alignment' => \PhpOffice\PhpWord\SimpleType\Jc::START, 'spaceAfter' => 0, 'lineHeight' => 1]);

    $infoTable->addRow();
    $infoTable->addCell(width(60))->addText('', $fontStyle, $DLeft);
    $infoTable->addCell(width(10))->addText('日期', $fontStyle, ['alignment' => 'distribute', 'lineHeight' => 1.0 ]);
    $infoTable->addCell(width(0.5))->addText('：', $fontStyle, ['lineHeight' => 1.0]);
    $infoTable->addCell(width(25))->addText($rcvbleDue, $fontStyle, ['alignment' => \PhpOffice\PhpWord\SimpleType\Jc::START, 'lineHeight' => 1.0]);

    $section->addTextBreak(1, ['size' => 3]);

    // ===== 3) 收據表格資訊 =====
    // 🌟 修正：收據主表格，單位改為 WIDTH_TWIP，總寬度 11000
    $table2Style = [
        'borderColor' => '000000',
        'borderSize'  => 15,
        'borderStyle' => 'thickThinSmallGap', 
        'insideRowBorderBorderStyle' => 'single',
        'insideRowBorderSize' => 1,
        'insideColBorderBorderStyle' => 'single',
        'insideColBorderSize' => 1,
        'cellMargin' => 80,
        'unit' => \PhpOffice\PhpWord\Style\Table::WIDTH_TWIP, // 改用絕對寬度 TWIP
        'width' => 11000, // 總長度 11000
        'alignment' => Jc::CENTER
    ];
    $font2Style = ['name' => 'Microsoft JhengHei', 'size' => 12, 'color' => '000000'];

    $phpWord->addTableStyle('ReceiptTable', $table2Style);
    $table = $section->addTable('ReceiptTable');
    
    $table->addRow();
    $table->addCell(width(60))->addText('項目', $font2Style, $TCenter);
    $table->addCell(width(40))->addText('金額(NTD)', $font2Style, $TCenter);

    $table->addRow();
    $cell = $table->addCell(width(60));
    $cell->addTextBreak(1);

    if ($documentType === 'domestic') {
        $cell->addText('代收代付國外規費('.$currency.number_format($Data['OfficialFee']).'*'.$rate.')', $font2Style, $TLeft);
        $cell->addText('代收代付國外代理人費用('.$currency.number_format($Data['AgentFee']).'*'.$rate.')', $font2Style, $TLeft);
        $cell->addTextBreak(1);
        $cell = $table->addCell(width(40));
        $cell->addTextBreak(1);
        $cell->addText(number_format((int)$Data['OfficialFee']*(float)$rate), $font2Style, $TRight);
        $cell->addText(number_format((int)$Data['AgentFee']*(float)$rate), $font2Style, $TRight);
        $cell->addTextBreak(1);

        $table->addRow();
        $table->addCell(width(60))->addText('合計', $font2Style, $TLeft);
        $table->addCell(width(40))->addText(number_format((int)$Data['OfficialFee']*(float)$rate+(int)$Data['AgentFee']*(float)$rate), $font2Style, $TRight);
    } else if ($documentType === 'foreignJoin') {
        $cell->addText('代收代付國外代理人費用('.$currency.(float)$amount.'*'.$rate.')', $font2Style, $TLeft);
        $cell->addTextBreak(1);
        $cell = $table->addCell(width(40));
        $cell->addTextBreak(1);
        $cell->addText(number_format((float)$amount*(float)$rate), $font2Style, $TRight);
        $cell->addTextBreak(1);

        $table->addRow();
        $table->addCell(width(60))->addText('合計', $font2Style, $TLeft);
        $table->addCell(width(40))->addText(number_format((float)$amount*(float)$rate), $font2Style, $TRight);
    } else if ($documentType === 'foreignSeparate') {
        $cell->addText('代收代付國外規費('.$currency.((float)$foreignCurrencyAmount).'*'.$rate.')', $font2Style, $TLeft);
        $cell->addText('代收代付國外代理人費用('.$currency.((float)$replaceAmount).'*'.$rate.')', $font2Style, $TLeft);
        $cell->addTextBreak(1);
        $cell = $table->addCell(width(40));
        $cell->addTextBreak(1);

        $foreignTotal = round((float)$foreignCurrencyAmount*(float)$rate);
        $replaceTotal = round((float)$replaceAmount*(float)$rate);
        $cell->addText(number_format($foreignTotal), $font2Style, $TRight);
        $cell->addText(number_format($replaceTotal), $font2Style, $TRight);
        $cell->addTextBreak(1);

        $table->addRow();
        $table->addCell(width(60))->addText('合計', $font2Style, $TLeft);
        $table->addCell(width(40))->addText(number_format($foreignTotal+$replaceTotal), $font2Style, $TRight);
    }
    
    $table->addRow();
    $cell = $table->addCell(width(100), ['gridSpan' => 2]);
    $cell->addText($Data['WhatFor'], $font2Style, $TLeft);
    $cell->addText($subtitle.$Data['CTitle'], $font2Style, $TLeft);
    $cell->addText('申請案號：'.$Data['AppCaseld'], $font2Style, $TLeft);
    $cell->addText('本所編號：'.$Data['FileId'], $font2Style, $TLeft);
    $cell->addText('客戶內部編號：'.$Data['CltFileID'], $font2Style, $TLeft);
}

if($data['Code'] === 200) {
    // 檔名
    $fileName = '財務部請款單_'. $Data['CustomerCName'] . ' ' . date('YmdHis');
    
    date_default_timezone_set('Asia/Taipei');
    write($phpWord, $fileName, $writers);

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