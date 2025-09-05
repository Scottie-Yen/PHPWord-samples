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
		// 'authorization: authorization ' . $_GET["token"]
    ));
    $result = curl_exec($ch);
    curl_close($ch);
  
    return json_decode($result, true);
}

// POST
$data='';

if(isset($_GET["FileID"])){
	$data = array(
		"FileID" => $_GET["FileID"],
        "PageList" => array('info', 'receivable1')
	);
}else{
	echo "Error : No FileID";
}

$data = httpRequest('http://192.168.0.52:8000/report/select-patentinfopublic', json_encode($data));
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
	'footerHeight' => 0
));

// Add header for all other pages
$header = $section->addHeader();
$header->addImage('images/巨群信頭-四所_四所 Head_0.png', 
array(
	'width' => 600, 
	'height' => 'auto', 
));

// Add footer
$footer = $section->addFooter();
$footer->addImage('images/巨群信頭-四所_四所 Bottom_0.png', 
array(
	'width' => 600, 
	'height' => 'auto', 
));
// $footer->addPreserveText('第 {PAGE} 頁，共 {NUMPAGES} 頁', null, array('alignment' => \PhpOffice\PhpWord\SimpleType\Jc::CENTER));


// Header Start
$HeaderStyle = array('borderColor' => '999999', 'cellMarginLeft' => 200, 'alignment' => \PhpOffice\PhpWord\SimpleType\Jc::END);
$phpWord->addTableStyle('Header Row Style', $HeaderStyle);
$HeaderTable = $section->addTable('Header Row Style');
$HeaderSpan = array('gridSpan' => 3, 'vMerge' => 'restart');
$AlignRight = array('align' => 'right');
$AlignCenter = array('align' => 'center');
$fontStyle = array();
$paragraphStyle = array('indent' => 1.2);

// Table樣式
$TitleStyle = array(
	// 'borderBottomSize' => 12,
    // 'borderBottomColor' => 'black',
    // 'borderTopSize' => 12,
    // 'borderTopColor' => 'black',
    // 'borderRightSize' => 12,
    // 'borderRightColor' => 'black',
    // 'borderLeftSize' => 12,
    // 'borderLeftColor' => 'black',
	'alignment' => \PhpOffice\PhpWord\SimpleType\Jc::CENTER ,
	'leftFromText'  => 1000,
	'cellMarginLeft' => 20, 
	'cellMarginRight' => 400,
);

$NullStyle = array(
    'alignment' => \PhpOffice\PhpWord\SimpleType\Jc::CENTER,
);

$phpWord->addTableStyle('NullTableStyle', $NullStyle);
$NullTable = $section->addTable('NullTableStyle');

$phpWord->addTableStyle('Title Row Style', $TitleStyle);
$TitleTable = $section->addTable('Title Row Style');
// 字體樣式
$CellCenter = array('vMerge' => 'restart', 'valign' => 'center');

// 左邊欄位樣式
$DashedStyle = array( 
    'vMerge' => 'restart',
    'valign' => 'center',
    'name' => 'DFKai-SB', 'size' => 12, 'hint' => 'eastAsia'
);

// 冒號樣式
$SubtStyle = array(
	'vMerge' => 'restart',
	'valign' => 'center',
);

// 內容使用底線
$DashedContentStyle = array(
	'gridSpan' => 3,
	'vMerge' => 'restart',
	'valign' => 'center',
	'borderBottomSize' => 1,
);

$ImageStyle = array(
	'vMerge' => 'restart',
	'valign' => 'center',
	'borderTopSize' => 1,
	'borderRightSize' => 1,
	'borderLeftSize' => 1,
	'borderBottomSize' => 1,
);

$fontBiaokai12 = ['name' => 'DFKai-SB', 'size' => 12, 'hint' => 'eastAsia']; // 標楷體 12pt
$width = 3600;
$TitleWidth = 2000;
$TableWidth = 250;


// 字體置中對其
$TextStart = array('alignment' => \PhpOffice\PhpWord\SimpleType\Jc::START);
$TextRun = array('alignment' => \PhpOffice\PhpWord\SimpleType\Jc::CENTER);
$TextEnd = array('alignment' => \PhpOffice\PhpWord\SimpleType\Jc::END);
$TextRunTop = $TextEnd;                // $TextRun 是你原本的對齊設定
$TextRunTop['spaceBefore'] = 50;      // 200 twips ≈ 0.14 in；數值可調
$TextRunTop['spaceAfter']  = 0;

$LineHeight = 700;

// Header start
if($data['Code'] == 200 && $Data['PatentInfo']['FileID'] && !empty($Data['Receivable'])) {
    $TitleRow = $NullTable->addRow(4000, ['exactHeight' => true]);
    $TitleRow->addCell(1250, [])->addTextRun($TextRun)->addText('');
    $TitleRow->addCell(4500, $ImageStyle)->addTextRun($TextRun)->addText('');

    $TitleRow = $NullTable->addRow(1170, ['exactHeight' => true]);
    $TitleRow->addCell(1250, [])->addTextRun($TextRun)->addText('');

	$TitleRow = $TitleTable->addRow($LineHeight, ['exactHeight' => true]);
    $TitleRow->addCell($TableWidth, [])->addTextRun([])->addText('');
    $leftCell = $TitleRow->addCell($TitleWidth, $DashedStyle);
    $leftCell->addText('客戶名稱', $DashedStyle, $TextRunTop);
    $leftCell->addText('(CLIENT)', $DashedStyle, $TextEnd);
    $TitleRow->addCell(100, $SubtStyle)->addTextRun($TextStart)->addText('：');
	$TitleRow->addCell($width, $DashedContentStyle)->addTextRun($TextRun)->addText($Data['PatentInfo']['CltCName'], $fontBiaokai12);

    $TitleRow = $TitleTable->addRow($LineHeight, ['exactHeight' => true]);
    $TitleRow->addCell($TableWidth, [])->addTextRun([])->addText('');
    $leftCell = $TitleRow->addCell($TitleWidth, $DashedStyle);
    $leftCell->addText('客戶內部編號', $DashedStyle, $TextRunTop);
    $leftCell->addText('(CLIENT.NO.)', $DashedStyle, $TextEnd);
    $TitleRow->addCell(100, $SubtStyle)->addTextRun($TextStart)->addText('：');
	$TitleRow->addCell($width, $DashedContentStyle)->addTextRun($TextRun)->addText($Data['PatentInfo']['CltFileID'], $fontBiaokai12);

    $TitleRow = $TitleTable->addRow($LineHeight, ['exactHeight' => true]);
    $TitleRow->addCell($TableWidth, [])->addTextRun([])->addText('');
    $leftCell = $TitleRow->addCell($TitleWidth, $DashedStyle);
    $leftCell->addText('申請國家', $DashedStyle, $TextRunTop);
    $leftCell->addText('(COUNTRY)', $DashedStyle, $TextEnd);
    $TitleRow->addCell(100, $SubtStyle)->addTextRun($TextStart)->addText('：');
	$TitleRow->addCell($width, $DashedContentStyle)->addTextRun($TextRun)->addText($Data['PatentInfo']['CountryName'], $fontBiaokai12);

    $TitleRow = $TitleTable->addRow($LineHeight, ['exactHeight' => true]);
    $TitleRow->addCell($TableWidth, [])->addTextRun([])->addText('');
    $leftCell = $TitleRow->addCell($TitleWidth, $DashedStyle);
    $leftCell->addText('案 別', $DashedStyle, $TextRunTop);
    $leftCell->addText('(CASE)', $DashedStyle, $TextEnd);
    $TitleRow->addCell(100, $SubtStyle)->addTextRun($TextStart)->addText('：');
	$TitleRow->addCell($width, $DashedContentStyle)->addTextRun($TextRun)->addText($Data['Receivable'][0]['WhatFor'], $fontBiaokai12);

	$TitleRow = $TitleTable->addRow($LineHeight, ['exactHeight' => true]);
    $TitleRow->addCell($TableWidth, [])->addTextRun([])->addText('');
    $leftCell = $TitleRow->addCell($TitleWidth, $DashedStyle);
    $leftCell->addText('申請日期', $DashedStyle, $TextRunTop);
    $leftCell->addText('(APPLDAY.)', $DashedStyle, $TextEnd);
    $TitleRow->addCell(100, $SubtStyle)->addTextRun($TextStart)->addText('：');
	$TitleRow->addCell($width, $DashedContentStyle)->addTextRun($TextRun)->addText($Data['PatentInfo']['AppDate'], $fontBiaokai12);

	$TitleRow = $TitleTable->addRow($LineHeight, ['exactHeight' => true]);
    $TitleRow->addCell($TableWidth, [])->addTextRun([])->addText('');
    $leftCell = $TitleRow->addCell($TitleWidth, $DashedStyle);
    $leftCell->addText('申請案號', $DashedStyle, $TextRunTop);
    $leftCell->addText('(APPL.NO.)', $DashedStyle, $TextEnd);
    $TitleRow->addCell(100, $SubtStyle)->addTextRun($TextStart)->addText('：');
	$TitleRow->addCell($width, $DashedContentStyle)->addTextRun($TextRun)->addText($Data['PatentInfo']['AppCaseld'], $fontBiaokai12);

	$TitleRow = $TitleTable->addRow($LineHeight, ['exactHeight' => true]);
    $TitleRow->addCell($TableWidth, [])->addTextRun([])->addText('');
    $leftCell = $TitleRow->addCell($TitleWidth, $DashedStyle);
    $leftCell->addText('申請名稱', $DashedStyle, $TextRunTop);
    $leftCell->addText('(TITLE)', $DashedStyle, $TextEnd);
    $TitleRow->addCell(100, $SubtStyle)->addTextRun($TextStart)->addText('：');
	$TitleRow->addCell($width, $DashedContentStyle)->addTextRun($TextRun)->addText($Data['PatentInfo']['CTitle'], $fontBiaokai12);

    $TitleRow = $TitleTable->addRow($LineHeight, ['exactHeight' => true]);
    $TitleRow->addCell($TableWidth, [])->addTextRun([])->addText('');
    $leftCell = $TitleRow->addCell($TitleWidth, $DashedStyle);
    $leftCell->addText('本所編號', $DashedStyle, $TextRunTop);
    $leftCell->addText('(CO.NO.)', $DashedStyle, $TextEnd);
    $TitleRow->addCell(100, $SubtStyle)->addTextRun($TextStart)->addText('：');
	$TitleRow->addCell($width, $DashedContentStyle)->addTextRun($TextRun)->addText($Data['PatentInfo']['FileID'], $fontBiaokai12);

    $TitleRow = $TitleTable->addRow($LineHeight, ['exactHeight' => true]);
    $TitleRow->addCell($TableWidth, [])->addTextRun([])->addText('');
    $leftCell = $TitleRow->addCell($TitleWidth, $DashedStyle);
    $leftCell->addText('服務人員', $DashedStyle, $TextRunTop);
    $leftCell->addText('(PECEPTIONIST)', $DashedStyle, $TextEnd);
    $TitleRow->addCell(100, $SubtStyle)->addTextRun($TextStart)->addText('：');
	$TitleRow->addCell($width, $DashedContentStyle)->addTextRun($TextRun)->addText($Data['PatentInfo']['PastSales'], $fontBiaokai12);
}


if($data['Code'] == 200 && $Data['PatentInfo']['FileID'] && !empty($Data['Receivable'])) {
	// Save file
	// echo write($phpWord, basename(__FILE__, '.php'), $writers);
    date_default_timezone_set('Asia/Taipei');
	write($phpWord, date('YmdHis'), $writers);
	if (!CLI) {
		include_once 'Sample_Footer.php';
	}


	echo "<script language='javascript' type ='text/javascript'>"; 
	echo "window.location.href = 'results/" . date('YmdHis') . ".docx';";
	echo 'document.getElementById("print").innerHTML = "檔案已下載完成。";';
	echo "</script>"; 
} else {
	echo "<script language='javascript' type ='text/javascript'>"; 

    $text = 'document.getElementById("print").innerHTML = "檔案下載失敗。';
	// echo 'document.getElementById("print").innerHTML = "檔案下載失敗。";';
    if ($data['Code'] !== 200) {
        $text.= '<br> 未取得資料API錯誤。 <br>';
    } else if (!$Data['PatentInfo']['FileID']) {
        $text.= '<br> 未取得PatentInfo陣列資料。 <br>';
    } else if (!empty($Data['Receivable'])) {
        $text.= '<br> 未取得Receivable陣列資料。 <br>';
    }

    echo $text . '";';
	echo "</script>"; 
}
?>