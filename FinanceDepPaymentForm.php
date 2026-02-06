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

// 寬度百分比計算
function width($percent) {
    return $percent / 100 * 11000;
}

// 計算約計
function Approximately ($fee, $rate) {
    return number_format((int) ceil((int)$fee * (float)$rate));
}

// 計算營業稅5%
function BusinessTax($fee) {
    return number_format((int) ceil((int)$fee * 0.05));
}

// 計算合計
function Total($fee, $rate, $inCaseSale2) {
    return number_format((int) ceil((int)$fee * (float)$rate) + (int)$inCaseSale2 + (int) ceil((int)$inCaseSale2 * 0.05));
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

// 國外案(分):規費外幣金額
if(isset($_GET["ForeignCurrencyAmount"])){
    $foreignCurrencyAmount = $_GET["ForeignCurrencyAmount"];
}
// 國外案(分):代理人費用外幣金額
if(isset($_GET["ReplaceAmount"])){
    $replaceAmount = $_GET["ReplaceAmount"];
}



$host = $_SERVER['HTTP_HOST']; // 例如 "192.168.0.52:8080" 或 "example.com"
if ($host === '192.168.0.52' || $host === '192.168.0.52:8080') {
    $data = httpRequest('http://192.168.0.52:8000/report/select-generatepaymentrequest', json_encode($data));
} elseif ($host === '192.168.0.188' || $host === '192.168.0.188:8080') {
    $data = httpRequest('http://192.168.0.186:8000/report/select-generatepaymentrequest', json_encode($data));
}
$Data = $data['Data'];
// echo json_encode($data, JSON_UNESCAPED_UNICODE);

// New Word document
include_once 'Sample_Header.php';
$phpWord = new \PhpOffice\PhpWord\PhpWord();

// New portrait section
$section = $phpWord->addSection(array(
	'headerHeight' => 300,
    'marginTop' => 2000,
	'marginLeft' => 0,
    'marginRight' => 0,
	'footerHeight' => 0,
));

// Add header for all other pages
$header = $section->addHeader();

if ($documentType === 'domestic') {

    if($_GET["type"] === 'Patent') {
        $header->addImage('images/巨群信頭_專利請款單.png', 
        array(
            'width' => 600, 
            'height' => 'auto', 
        ));
    }

    if($_GET["type"] === 'TM') {
        $header->addImage('images/巨群信頭_商標請款單.png', 
        array(
            'width' => 600, 
            'height' => 'auto', 
        ));
    }
    
} else if ($documentType !== 'domestic') {

    if($_GET["type"] === 'Patent') {
        $header->addImage('images/巨群信頭_智財諮詢專利請款單.png', 
        array(
            'width' => 600, 
            'height' => 'auto', 
        ));
    }

    if($_GET["type"] === 'TM') {
        $header->addImage('images/巨群信頭_智財諮詢商標請款單.png', 
        array(
            'width' => 600, 
            'height' => 'auto', 
        ));
    }

} else {
    $header->addImage('images/巨群信頭_國外請款單.png', 
    array(
        'width' => 600, 
        'height' => 'auto', 
    ));
}

// Add footer
$footer = $section->addFooter();
if ($documentType === 'domestic') {

    if ($_GET["type"] === 'Patent') {
        $footer->addImage('images/巨群信頭_巨群智財諮詢 Confidential 永豐1.png', 
        array(
            'width' => 600, 
            'height' => 'auto', 
        ));
    }

    if ($_GET["type"] === 'TM') {
        $footer->addImage('images/巨群信頭_巨群智財諮詢 Confidential 永豐2.png', 
        array(
            'width' => 600, 
            'height' => 'auto', 
        ));
    }

    
} else {
    $footer->addImage('images/巨群信頭_巨群智財諮詢 Confidential 富邦.png', 
    array(
        'width' => 600, 
        'height' => 'auto', 
    ));
}


// ===== 1) 文件標題 =====
$section->addText(
    '請款單',
    ['name' => 'Microsoft JhengHei', 'size' => 26, 'bold' => true, 'underline' => 'double'],
    ['alignment' => Jc::CENTER, 'spaceAfter' => 0,],
);

// ===== 2) 客戶資訊 =====
$section->addText(
    $Data['CustomerCName'],
    ['name' => 'Microsoft JhengHei', 'size' => 14, 'bold' => true],
    ['alignment' => Jc::START, 'spaceAfter' => 0, 'indentation' => [ 'left'  => Converter::cmToTwip(1)], 'lineHeight' => 0.7,],
);
$section->addText(
    $Data['ContactAddress'],
    ['name' => 'Microsoft JhengHei', 'size' => 14, 'bold' => true],
    ['alignment' => Jc::START, 'spaceAfter' => 100, 'indentation' => [ 'left'  => Converter::cmToTwip(1)], 'lineHeight' => 0.7,],
);

$titleStyle = ['name' => 'Microsoft JhengHei', 'size' => 10, 'color' => '000000', 'bold' => true];
$fontStyle = ['name' => 'Microsoft JhengHei', 'size' => 10, 'color' => '000000'];
$CLeft = ['alignment' => \PhpOffice\PhpWord\SimpleType\Jc::START, 'spaceAfter' => 0, 'lineHeight' => 1.0, 'indentation' => [ 'left'  => Converter::cmToTwip(1)]];

$infoTable = $section->addTable([
    'borderSize' => 0,      // 邊框大小為 0
    'borderColor' => 'FFFFFF', // 邊框顏色設為白色 (隱形)
    'cellMargin' => 0,      // 儲存格無內距
    'unit' => \PhpOffice\PhpWord\Style\Table::WIDTH_PERCENT, // 設定寬度單位為百分比
    'width' => 100 * 50     // 設定寬度 100% (50 是 PHPWord 的百分比係數)
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
$infoTable->addCell(12500, ['gridSpan' => 2])->addText($subtitle.$Data['CTitle'], $fontStyle, ['alignment' => \PhpOffice\PhpWord\SimpleType\Jc::START, 'spaceAfter' => 100, 'lineHeight' => 1.0, 'indentation' => [ 'left'  => Converter::cmToTwip(1)]]);

// ===== 3) 表格資訊 =====
// 定義表格框線樣式
$tableStyle = [
    'borderSize' => 6, 
    'borderColor' => '000000', 
    'cellMargin' => 80,
    'unit' => \PhpOffice\PhpWord\Style\Table::WIDTH_PERCENT,
    'width' => 100 * 50,
    'alignment' => Jc::CENTER,
    'indent' => new \PhpOffice\PhpWord\ComplexType\TblWidth(
        \PhpOffice\PhpWord\Shared\Converter::cmToTwip(2)
    )
];
$TCenter = ['alignment' => Jc::CENTER, 'spaceAfter' => 0, 'lineHeight' => 1.0];
$TLeft = ['alignment' => \PhpOffice\PhpWord\SimpleType\Jc::START, 'spaceAfter' => 0, 'lineHeight' => 1.0];
$TRight = ['alignment' => \PhpOffice\PhpWord\SimpleType\Jc::END, 'spaceAfter' => 0, 'lineHeight' => 1.0];

$hideRightLine = ['borderRightColor' => 'FFFFFF', 'borderRightSize' => 0,];
$hideLeftLine = ['borderLeftColor' => 'FFFFFF', 'borderLeftSize' => 0,];
$hideTitleLine = ['borderLeftColor' => 'FFFFFF', 'borderLeftSize' => 0, 'borderBottomColor' => 'FFFFFF', 'borderBottomSize' => 0, 'borderTopColor' => 'FFFFFF', 'borderTopSize' => 0,];

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
    $table->addCell(width(10), $hideLeftLine)->addText(number_format($Data['InCaseSale2']), $fontStyle, $TRight);

    $table->addRow();
    $table->addCell(width(53), ['borderLeftColor' => 'FFFFFF', 'borderLeftSize' => 0, 'borderBottomColor' => 'FFFFFF', 'borderBottomSize' => 0,])->addText('');
    $table->addCell(width(12))->addText('小計', $fontStyle, $TCenter);
    $table->addCell(width(5), $hideRightLine)->addText('NTD', $fontStyle, $TLeft);
    $table->addCell(width(15), $hideLeftLine)->addText(number_format($Data['OfficialFee']), $fontStyle, $TRight);
    $table->addCell(width(5), $hideRightLine)->addText('NTD', $fontStyle, $TLeft);
    $table->addCell(width(10), $hideLeftLine)->addText(number_format($Data['InCaseSale2']), $fontStyle, $TRight);

    $table->addRow();
    $table->addCell(width(53), $hideTitleLine)->addText('');
    $table->addCell(width(12))->addText('合計', $fontStyle, $TCenter);
    $table->addCell(width(20), ['gridSpan' => 2, 'borderRightColor' => 'FFFFFF', 'borderRightSize' => 0,])->addText('NTD', $fontStyle, $TLeft);
    $table->addCell(width(15), ['gridSpan' => 2, 'borderLeftColor' => 'FFFFFF', 'borderLeftSize' => 0,])->addText(number_format($Data['OfficialFee'] + $Data['InCaseSale2']), $fontStyle, $TRight);

} else if ($documentType === 'foreignJoin') {
    // 國外案合
    $table->addRow();
    $table->addCell(width(65), ['gridSpan' => 2])->addText('項目', $titleStyle, $TCenter);
    $table->addCell(width(20), ['gridSpan' => 2])->addText('代收代付費用', $titleStyle, $TCenter);
    $table->addCell(width(15), ['gridSpan' => 2])->addText('本所服務費', $titleStyle, $TCenter);

    $table->addRow();
    $table->addCell(width(65), ['gridSpan' => 2])->addText($Data['WhatFor'], $fontStyle, $TLeft);
    $table->addCell(width(5), $hideRightLine)->addText($currency, $fontStyle, $TLeft);
    $table->addCell(width(15), $hideLeftLine)->addText(number_format($amount), $fontStyle, $TRight);
    $table->addCell(width(5), $hideRightLine)->addText('NTD', $fontStyle, $TLeft);
    $table->addCell(width(10), $hideLeftLine)->addText(number_format($Data['InCaseSale2']), $fontStyle, $TRight);

    $table->addRow();
    $table->addCell(width(53), ['borderLeftColor' => 'FFFFFF', 'borderLeftSize' => 0, 'borderBottomColor' => 'FFFFFF', 'borderBottomSize' => 0,])->addText('');
    $table->addCell(width(12))->addText('小計', $fontStyle, $TCenter);
    $table->addCell(width(5), $hideRightLine)->addText($currency, $fontStyle, $TLeft);
    $table->addCell(width(15), $hideLeftLine)->addText(number_format($amount), $fontStyle, $TRight);
    $table->addCell(width(5), $hideRightLine)->addText('NT$', $fontStyle, $TLeft);
    $table->addCell(width(10), $hideLeftLine)->addText(number_format($Data['InCaseSale2']), $fontStyle, $TRight);

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
    $table->addCell(width(15), $hideLeftLine)->addText(Approximately($amount, $rate), $fontStyle, $TRight);
    $table->addCell(width(5), $hideRightLine)->addText('NTD', $fontStyle, $TLeft);
    $table->addCell(width(10), $hideLeftLine)->addText(number_format($Data['InCaseSale2']), $fontStyle, $TRight);

    $table->addRow();
    $table->addCell(width(53), $hideTitleLine)->addText('');
    $table->addCell(width(12))->addText('營業稅5%', $fontStyle, $TCenter);
    $table->addCell(width(20), ['gridSpan' => 2])->addText('----', $fontStyle, $TCenter);
    $table->addCell(width(5), $hideRightLine)->addText('NTD', $fontStyle, $TLeft);
    $table->addCell(width(10), $hideLeftLine)->addText(BusinessTax($Data['InCaseSale2']), $fontStyle, $TRight);

    $table->addRow();
    $table->addCell(width(53), $hideTitleLine)->addText('');
    $table->addCell(width(12))->addText('合計', $fontStyle, $TCenter);
    $table->addCell(width(20), ['gridSpan' => 2, 'borderRightColor' => 'FFFFFF', 'borderRightSize' => 0,])->addText('NTD', $fontStyle, $TLeft);
    $table->addCell(width(15), ['gridSpan' => 2, 'borderLeftColor' => 'FFFFFF', 'borderLeftSize' => 0,])->addText(Total($amount, $rate, $Data['InCaseSale2']), $fontStyle, $TRight);

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
    $table->addCell(width(10), $hideLeftLine)->addText(number_format($Data['InCaseSale2']), $fontStyle, $TRight);

    $table->addRow();
    $table->addCell(width(27.5), ['borderSize' => 0, 'borderColor' => 'FFFFFF'])->addText('');
    $table->addCell(width(15.5), ['borderSize' => 0, 'borderColor' => 'FFFFFF'])->addText('');
    $table->addCell(width(12))->addText('小計', $fontStyle, $TCenter);
    $table->addCell(width(5), $hideRightLine)->addText($currency, $fontStyle, $TLeft);
    $table->addCell(width(25), ['borderLeftColor' => 'FFFFFF', 'borderLeftSize' => 0, 'gridSpan' => 3])->addText(number_format($foreignCurrencyAmount+$replaceAmount), $fontStyle, $TRight);
    $table->addCell(width(5), $hideRightLine)->addText('NTD', $fontStyle, $TLeft);
    $table->addCell(width(10), $hideLeftLine)->addText(number_format($Data['InCaseSale2']), $fontStyle, $TRight);

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
    $table->addCell(width(25), ['borderLeftColor' => 'FFFFFF', 'borderLeftSize' => 0, 'gridSpan' => 3])->addText(Approximately($foreignCurrencyAmount+$replaceAmount, $rate), $fontStyle, $TRight);
    $table->addCell(width(5), $hideRightLine)->addText('NTD', $fontStyle, $TLeft);
    $table->addCell(width(10), $hideLeftLine)->addText(number_format($Data['InCaseSale2']), $fontStyle, $TRight);

    $table->addRow();
    $table->addCell(width(27.5), ['borderSize' => 0, 'borderColor' => 'FFFFFF'])->addText('');
    $table->addCell(width(15.5), ['borderSize' => 0, 'borderColor' => 'FFFFFF'])->addText('');
    $table->addCell(width(12))->addText('營業稅5%', $fontStyle, $TCenter);
    $table->addCell(width(30), ['gridSpan' => 4])->addText('----', $fontStyle, $TCenter);
    $table->addCell(width(5), $hideRightLine)->addText('NTD', $fontStyle, $TLeft);
    $table->addCell(width(10), $hideLeftLine)->addText(BusinessTax($Data['InCaseSale2']), $fontStyle, $TRight);

    $table->addRow();
    $table->addCell(width(27.5), ['borderSize' => 0, 'borderColor' => 'FFFFFF'])->addText('');
    $table->addCell(width(15.5), ['borderSize' => 0, 'borderColor' => 'FFFFFF'])->addText('');
    $table->addCell(width(12))->addText('合計', $fontStyle, $TCenter);
    $table->addCell(width(5), $hideRightLine)->addText('NTD', $fontStyle, $TLeft);
    $table->addCell(width(40), ['gridSpan' => 2, 'borderLeftColor' => 'FFFFFF', 'borderLeftSize' => 0, 'gridSpan' => 5])->addText(Total($foreignCurrencyAmount+$replaceAmount, $rate, $Data['InCaseSale2']), $fontStyle, $TRight);
}


if ($documentType !== 'domestic') {
    $section->addPageBreak();
    
    // ===== 1) 文件標題 =====
    $section->addText(
        '代收代付國外收據',
        ['name' => 'Microsoft JhengHei', 'size' => 18,],
        ['alignment' => Jc::CENTER, 'spaceAfter' => 0, 'lineHeight' => 0.7],
    );
    $section->addText(
        '(正本)',
        ['name' => 'Microsoft JhengHei', 'size' => 14,],
        ['alignment' => Jc::CENTER, 'spaceAfter' => 0, 'lineHeight' => 0.7],
    );

    // ===== 2) 客戶資訊 =====
    $section->addText(
        $Data['CustomerCName'],
        ['name' => 'Microsoft JhengHei', 'size' => 14],
        ['alignment' => Jc::START, 'spaceAfter' => 0, 'indentation' => [ 'left'  => Converter::cmToTwip(1)], 'lineHeight' => 0.7,],
    );

    $infoTable = $section->addTable([
        'borderSize' => 0,      // 邊框大小為 0
        'borderColor' => 'FFFFFF', // 邊框顏色設為白色 (隱形)
        'cellMargin' => 0,      // 儲存格無內距
        'unit' => \PhpOffice\PhpWord\Style\Table::WIDTH_PERCENT, // 設定寬度單位為百分比
        'width' => 100 * 50     // 設定寬度 100% (50 是 PHPWord 的百分比係數)
    ]);

    $DLeft = ['alignment' => \PhpOffice\PhpWord\SimpleType\Jc::START, 'spaceAfter' => 0, 'lineHeight' => 0.7, 'indentation' => [ 'left'  => Converter::cmToTwip(1)]];

    $infoTable->addRow();
    $infoTable->addCell(width(60))->addText('', $fontStyle, $DLeft);
    $infoTable->addCell(width(10))->addText('客戶編號', $fontStyle, ['alignment' => 'distribute', 'lineHeight' => 0.7, 'spaceAfter' => 0]);
    $infoTable->addCell(width(0.5))->addText('：', $fontStyle, ['lineHeight' => 0.7, 'spaceAfter' => 0]);
    $infoTable->addCell(width(25))->addText($Data['CltFileID'], $fontStyle, ['alignment' => \PhpOffice\PhpWord\SimpleType\Jc::START, 'spaceAfter' => 0, 'lineHeight' => 0.7,]);

    $infoTable->addRow();
    $infoTable->addCell(width(60))->addText('', $fontStyle, $DLeft);
    $infoTable->addCell(width(10))->addText('收據單號', $fontStyle, ['alignment' => 'distribute', 'lineHeight' => 0.7, 'spaceAfter' => 0]);
    $infoTable->addCell(width(0.5))->addText('：', $fontStyle, ['lineHeight' => 0.7, 'spaceAfter' => 0]);
    $infoTable->addCell(width(25))->addText($rcvAcntNum, $fontStyle, ['alignment' => \PhpOffice\PhpWord\SimpleType\Jc::START, 'spaceAfter' => 0, 'lineHeight' => 0.7,]);

    $infoTable->addRow();
    $infoTable->addCell(width(60))->addText('', $fontStyle, $DLeft);
    $infoTable->addCell(width(10))->addText('日期', $fontStyle, ['alignment' => 'distribute', 'lineHeight' => 1.0, ]);
    $infoTable->addCell(width(0.5))->addText('：', $fontStyle, ['lineHeight' => 1.0,]);
    $infoTable->addCell(width(25))->addText($rcvbleDue, $fontStyle, ['alignment' => \PhpOffice\PhpWord\SimpleType\Jc::START, 'lineHeight' => 1.0,]);

    // ===== 2) 表格資訊 =====
    $table2Style = [
        // 1. 設定邊框顏色
        'borderColor' => '000000',
        
        // 2. 設定邊框大小 (建議設大一點，例如 12 或 18，才看得出粗細差別)
        'borderSize'  => 15,
        
        // 3. 【關鍵】設定邊框樣式為 "粗-細"
        'borderStyle' => 'thickThinSmallGap', 
        
        // 4. (選用) 如果你希望裡面還是普通的單線，要分開設定 inside
        'insideRowBorderBorderStyle' => 'single',
        'insideRowBorderSize' => 1,
        'insideColBorderBorderStyle' => 'single',
        'insideColBorderSize' => 1,
        'cellMargin' => 80,
        'unit' => \PhpOffice\PhpWord\Style\Table::WIDTH_PERCENT,
        'width' => 100 * 50,
        'alignment' => Jc::CENTER,
        'indent' => new \PhpOffice\PhpWord\ComplexType\TblWidth(
            \PhpOffice\PhpWord\Shared\Converter::cmToTwip(2)
        )
    ];
    $font2Style = ['name' => 'Microsoft JhengHei', 'size' => 12, 'color' => '000000'];

    $phpWord->addTableStyle('ReceiptTable', $table2Style);
    $table = $section->addTable('ReceiptTable');
    $table->addRow();
    $table->addCell(width(60))->addText('項目', $font2Style, $TCenter);
    $table->addCell(width(60))->addText('金額(NTD)', $font2Style, $TCenter);

    // $table->addRow();
    // $table->addCell(width(60))->addText('代收代付國外代理人服務費('.$currency.$amount.'*'.$rate.')', $font2Style, $TLeft);
    // $table->addCell(width(40))->addText(number_format((int)$amount*(float)$rate), $font2Style, $TRight);

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
        $cell->addText('代收代付國外代理人費用('.$currency.number_format($amount).'*'.$rate.')', $font2Style, $TLeft);
        $cell->addTextBreak(1);
        $cell = $table->addCell(width(40));
        $cell->addTextBreak(1);
        $cell->addText(number_format((int)$amount*(float)$rate), $font2Style, $TRight);
        $cell->addTextBreak(1);

        $table->addRow();
        $table->addCell(width(60))->addText('合計', $font2Style, $TLeft);
        $table->addCell(width(40))->addText(number_format((int)$amount*(float)$rate), $font2Style, $TRight);
    } else if ($documentType === 'foreignSeparate') {
        $cell->addText('代收代付國外規費('.$currency.number_format($foreignCurrencyAmount).'*'.$rate.')', $font2Style, $TLeft);
        $cell->addText('代收代付國外代理人費用('.$currency.number_format($replaceAmount).'*'.$rate.')', $font2Style, $TLeft);
        $cell->addTextBreak(1);
        $cell = $table->addCell(width(40));
        $cell->addTextBreak(1);
        $cell->addText(number_format((int)$foreignCurrencyAmount*(float)$rate), $font2Style, $TRight);
        $cell->addText(number_format((int)$replaceAmount*(float)$rate), $font2Style, $TRight);
        $cell->addTextBreak(1);

        $table->addRow();
        $table->addCell(width(60))->addText('合計', $font2Style, $TLeft);
        $table->addCell(width(40))->addText(number_format((int)$foreignCurrencyAmount*(float)$rate+(int)$replaceAmount*(float)$rate), $font2Style, $TRight);
    }
    

    $table->addRow();
    $cell = $table->addCell(width(100), ['gridSpan' => 2]);
    $cell->addText($Data['WhatFor'], $font2Style, $TLeft);
    $cell->addText($subtitle.$Data['CTitle'], $font2Style, $TLeft);
    $cell->addText('申請案號：'.$Data['AppCaseld'], $font2Style, $TLeft);
    $cell->addText('本所編號：'.$Data['FileId'], $font2Style, $TLeft);
    $cell->addText('客戶內部編號：'.$Data['CltFileID'], $font2Style, $TLeft);

    // $transferTable = $section->addTable([
    //     'borderSize' => 0,      // 邊框大小為 0
    //     'borderColor' => 'FFFFFF', // 邊框顏色設為白色 (隱形)
    //     'cellMargin' => 0,      // 儲存格無內距
    //     'unit' => \PhpOffice\PhpWord\Style\Table::WIDTH_PERCENT, // 設定寬度單位為百分比
    //     'width' => 100 * 50     // 設定寬度 100% (50 是 PHPWord 的百分比係數)
    // ]);
    // $font3Style = ['name' => 'Microsoft JhengHei', 'size' => 12, 'color' => '000000'];

    // $transferTable->addRow();
    // $transferTable->addCell(width(30), ['gridSpan' => 2])->addText('煩請  貴客戶匯至下列帳戶：', $font3Style, ['alignment' => Jc::START, 'spaceAfter' => 0, 'indentation' => [ 'left'  => Converter::cmToTwip(1)], 'spaceBefore' => Converter::cmToTwip(0.75),]);

    // $transferTable->addRow();
    // $transferTable->addCell(width(5))->addText('銀行', $font3Style, ['alignment' => 'distribute', 'indentation' => [ 'left'  => Converter::cmToTwip(1)], 'lineHeight' => 0.7,]);
    // $transferTable->addCell(width(10))->addText('： 台北富邦銀行(八德分行)', $font3Style, ['alignment' => \PhpOffice\PhpWord\SimpleType\Jc::START, 'spaceAfter' => 0, 'lineHeight' => 0.7,]);
    
    // $transferTable->addRow();
    // $transferTable->addCell(width(5))->addText('銀行代號', $font3Style, ['alignment' => 'distribute', 'indentation' => [ 'left'  => Converter::cmToTwip(1)], 'lineHeight' => 0.7,]);
    // $transferTable->addCell(width(10))->addText('： 012', $font3Style, ['alignment' => \PhpOffice\PhpWord\SimpleType\Jc::START, 'spaceAfter' => 0, 'lineHeight' => 0.7,]);

    // $transferTable->addRow();
    // $transferTable->addCell(width(5))->addText('戶名', $font3Style, ['alignment' => 'distribute', 'indentation' => [ 'left'  => Converter::cmToTwip(1)], 'lineHeight' => 0.7,]);
    // $transferTable->addCell(width(10))->addText('： 巨群智財諮詢有限公司', $font3Style, ['alignment' => \PhpOffice\PhpWord\SimpleType\Jc::START, 'spaceAfter' => 0, 'lineHeight' => 0.7,]);

    // $transferTable->addRow();
    // $transferTable->addCell(width(5))->addText('帳號', $font3Style, ['alignment' => 'distribute', 'indentation' => [ 'left'  => Converter::cmToTwip(1)], 'lineHeight' => 0.7,]);
    // $transferTable->addCell(width(10))->addText('： 34010209427-7', $font3Style, ['alignment' => \PhpOffice\PhpWord\SimpleType\Jc::START, 'spaceAfter' => 0, 'lineHeight' => 0.7,]);
}




if($data['Code'] === 200) {
    // 檔名
    $fileName = '財務部請款單_'. $Data['CustomerCName'] . ' ' . date('YmdHis');
    
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