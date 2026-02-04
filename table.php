<?php
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

// POST
$data='';

if(isset($_GET["seq"])){
	$data = array(
		"Seq" => (int)$_GET["seq"]
	);
}else{
	echo "Error";
}

$data = httpRequest('http://192.168.0.52:8000/law/select-contract-requestpayment', json_encode($data));
$Data = $data['Data'];
// echo json_encode($data, JSON_UNESCAPED_UNICODE);


// 將起始日期和最終合併成文字
function ContractDate ($ContractSDate, $ContractEDate) {
	$i = 0;
	$j = 0;
	$countS = 0;
	$countE = 0;
	$SDate = '';
	$EDate = '';
	while(isset($ContractSDate[$i])){
		if($ContractSDate[$i]!='-'){
			$SDate = $SDate.$ContractSDate[$i];
		}else if($countS==0){
			$SDate = $SDate.'年';
			$countS++;
		}else if($countS==1){
			$SDate = $SDate.'月';
			$countS++;
		}else if($countS==2){
			$SDate = $SDate.'日';
			break;
		}	
		$i++;
	}
	while(isset($ContractEDate[$j])){
		if($ContractEDate[$j]!='-'){
			$EDate = $EDate.$ContractEDate[$j];
		}else if($countE==0){
			$EDate = $EDate.'年';
			$countE++;
		}else if($countE==1){
			$EDate = $EDate.'月';
			$countE++;
		}else if($countE==2){
			$EDate = $EDate.'日';
			break;
		}	
		$j++;
	}
	return $SDate.'-'.$EDate;
}

// 將2022-07-29 轉成 2022年07月29日
function ContractSEDate ($date) {
	$i = 0;
	$countS = 0;
	$PrintDate = '';
	while(isset($date[$i])){
		if($date[$i]!='-'){
			$PrintDate = $PrintDate.$date[$i];
		}else if($countS==0){
			$PrintDate = $PrintDate.'年';
			$countS++;
		}else if($countS==1){
			$PrintDate = $PrintDate.'月';
			$countS++;
		}
		$i++;
	}
	return $PrintDate.'日';
}

// 計算總共累計時數 個人取Hours欄位 不分取WorkHours欄位
function TotalHours($HoursData, $type){
	$i = 0;
	$total = 0;
	$field = '';

	if (isset($HoursData[$i]["Hours"]) && $type=='P') {
		$field = "Hours";
	} else if (isset($HoursData[$i]["WorkHours"])  && $type=='All') {
		$field = "WorkHours";
	}

	while(isset($HoursData[$i][$field])){
		$total=$total+(float)$HoursData[$i][$field];
		$i++;
	}
	return $total;
}

// 不分報表 計算各律師時數
function PersonalHours($HoursData, $Name){
	$i = 0;
	$total = 0;
	while(isset($HoursData[$i]['WorkHours'])){
		if($HoursData[$i]['Name'] == $Name) {
			$total=$total+(float)$HoursData[$i]['WorkHours'];
		}
		$i++;
	}
	return $total;
}

// 月日期範圍以-區分
function Period($WorkDate){
	$dateList = [];
	$count = 0;
	$i = 0;

	while(isset($WorkDate[$i])){
		if($WorkDate[$i] == '-'){
			$count++;
		}else{
			if(isset($dateList[$count])){
				$dateList[$count] = $dateList[$count].$WorkDate[$i];
			}else{
				$dateList[$count] = $WorkDate[$i];
			}
		}
		$i++;
	}
	return $dateList[0].'年'.$dateList[1].'月'.$dateList[2].'日至'.$dateList[3].'年'.$dateList[4].'月'.$dateList[5].'日';
}

// Discount
function Discount($TotalCost, $CasePrice, $ContractDiscount, $OverDiscount, $num){
	// CasePrice=0 , ContractDiscount=9 , OverDiscount=0
    if($TotalCost > $CasePrice){
        if($num){
			// $Discount = $ContractDiscount<($OverDiscount==10?0:$OverDiscount)?$ContractDiscount:$OverDiscount;
			$Discount = $ContractDiscount<$OverDiscount?$ContractDiscount:$OverDiscount;

			$DiscountList = ['一', '二', '三', '四', '五', '六', '七', '八', '九'];

			if($Discount==10){
				return "無折扣";
			} else {
				return $DiscountList[$Discount-1].'折';
			}

        }else{
            $a = (float)$ContractDiscount/10;
			$b = (float)$OverDiscount/10;
            return $a<$b?$a:$b;
        }
    }else{
		if($num){
			switch ($ContractDiscount) {
				case 1:
					return "一折";
				case 1.5:
					return "一五折";
				case 2:
					return "二折";
				case 2.5:
					return "二五折";
				case 3:
					return "三折";
				case 3.5:
					return "三五折";
				case 4:
					return "四折";
				case 4.5:
					return "四五折";
				case 5:
					return "五折";
				case 5.5:
					return "五五折";
				case 6:
					return "六折";
				case 6.5:
					return "六五折";
				case 7:
					return "七折";
				case 7.5:
					return "七五折";
				case 8:
					return "八折";
				case 8.5:
					return "八五折";
				case 9:
					return "九折";
				case 9.5:
					return "九五折";
				case 10:
					return "無折扣";
				default:
					return "Error";
			}
		}else{
			return (float)$ContractDiscount/10;
		}
	}
}

// position
function Position($Position){
    $tmp = '';
    if(strlen($Position)>=12){
        for($i=0;$i<strlen($Position);$i++){
            if($i>5){
                $tmp = $tmp.$Position[$i];
            }
        }
        return $tmp;
    }else{
        return $Position;
    }
}

// 計算總金額 個人則取時數及時間
function TotalPrice($Data, $Name){
    $i = 0;
    $totalprice = 0;
    while(isset($Data["HoursData"][$i]["Name"])){
		if($Data['PaymentType'] == "P"){
			$totalprice = $totalprice+$Data["HoursData"][$i]["Pay"]*$Data["HoursData"][$i]["Hours"];
		}
		if($Data['PaymentType'] == "All" && $Data["HoursData"][$i]["Name"]==$Name ){
			$totalprice = $totalprice+$Data["HoursData"][$i]["Pay"]*$Data["HoursData"][$i]["WorkHours"];
		}
        $i++;
    }
    return $totalprice;
}

// New Word document
include_once 'Sample_Header.php';
$phpWord = new \PhpOffice\PhpWord\PhpWord();

// New portrait section
$section = $phpWord->addSection(array(
	'headerHeight' => 0,
    'marginTop' => -2000,
	'marginLeft' => 0,
	'footerHeight' => 1200
));

// Add header for all other pages
$header = $section->addHeader();
$header->addImage('images/巨群信頭_法律所 CH.png', 
array(
	'width' => 600, 
	'height' => 'auto', 
));

// Add footer
$footer = $section->addFooter();
$footer->addPreserveText('第 {PAGE} 頁，共 {NUMPAGES} 頁', null, array('alignment' => \PhpOffice\PhpWord\SimpleType\Jc::CENTER));


// Header Start
$HeaderStyle = array('borderColor' => '999999', 'cellMarginLeft' => 200, 'alignment' => \PhpOffice\PhpWord\SimpleType\Jc::END);
$phpWord->addTableStyle('Header Row Style', $HeaderStyle);
$HeaderTable = $section->addTable('Header Row Style');
$width = 3333;
$HeaderSpan = array('gridSpan' => 3, 'vMerge' => 'restart');
$AlignRight = array('align' => 'right');
$AlignCenter = array('align' => 'center');
$fontStyle = array();
$paragraphStyle = array('indent' => 1.2);

// 8/3淘汰原表格
// $HeaderRow = $HeaderTable->addRow();
// $HeaderRow->addCell($width)->addText('當事人名稱：', $fontStyle, $AlignRight);
// $HeaderRow->addCell($width, $HeaderSpan)->addText(' '.$Data['ClientName']);

// $HeaderRow = $HeaderTable->addRow();
// $HeaderRow->addCell($width)->addText('地址：', $fontStyle, $AlignRight);
// $HeaderRow->addCell($width, $HeaderSpan)->addText(' '.$Data['CompanyAddress']);

// $HeaderRow = $HeaderTable->addRow();
// $HeaderRow->addCell($width)->addText('法定代理人：', $fontStyle, $AlignRight);
// $HeaderRow->addCell($width, $HeaderSpan)->addText(' '.$Data['CompanyOwner']);

// $HeaderRow = $HeaderTable->addRow();
// $HeaderRow->addCell($width)->addText('案別：', $fontStyle, $AlignRight);
// $HeaderRow->addCell($width, $HeaderSpan)->addText(' '.$Data['CaseName']);

// $HeaderRow = $HeaderTable->addRow();
// $HeaderRow->addCell($width)->addText('顧問期間：', $fontStyle, $AlignRight);
// $HeaderRow->addCell($width)->addText(' '.ContractSEDate($Data['ContractSDate']));
// $HeaderRow->addCell($width)->addText('至', $fontStyle, $AlignCenter);
// $HeaderRow->addCell($width)->addText(' '.ContractSEDate($Data['ContractEDate']));
// // 原所有字串一併輸出
// // $HeaderRow->addCell($width)->addText(' '.ContractDate($Data['ContractSDate'], $Data['ContractEDate']));

// $HeaderRow = $HeaderTable->addRow();
// $HeaderRow->addCell($width)->addText('本所編號：', $fontStyle, $AlignRight);
// $HeaderRow->addCell($width, $HeaderSpan)->addText(' '.$Data['CaseNumber']);
// // 取消這欄
// // $HeaderRow->addCell($width)->addText('時　　數　　　總　　計');
// // $HeaderRow->addCell($width)->addText(' '.TotalHours($Data['HoursData']));


// Table樣式
$TitleStyle = array(
	'borderBottomSize' => 12,
    'borderBottomColor' => 'black',
    'borderTopSize' => 12,
    'borderTopColor' => 'black',
    'borderRightSize' => 12,
    'borderRightColor' => 'black',
    'borderLeftSize' => 12,
    'borderLeftColor' => 'black',
	'alignment' => \PhpOffice\PhpWord\SimpleType\Jc::END ,
	'leftFromText'  => 1000,
	'cellMarginLeft' => 400, 
	'cellMarginRight' => 400,
);

$phpWord->addTableStyle('Title Row Style', $TitleStyle);
$TitleTable = $section->addTable('Title Row Style');
// 字體樣式
$CellCenter = array('vMerge' => 'restart', 'valign' => 'center');

// 標題除最後一行的虛線格式
$DashedStyle = array( 
	// 'borderRightStyle' => 'dashed', 
	// 'borderRightSize' => 2,
	// 'borderBottomStyle' => 'dashed', 
	'borderBottomSize' => 2,
);

// 內容除最後一行的虛線格式
$DashedContentStyle = array(
	'gridSpan' => 3,
	'vMerge' => 'restart',
	'valign' => 'center',
	// 'borderBottomStyle' => 'dashed', 
	'borderBottomSize' => 2,
);

// 顧問期間內容用格式
$DashedContentSingleStyle = array(
	'gridSpan' => 1,
	'vMerge' => 'restart',
	'valign' => 'center',
	// 'borderBottomStyle' => 'dashed', 
	'borderBottomSize' => 2,
);

// 內容最後一行的虛線格式
$DashedContentBottomStyle = array(
	'gridSpan' => 3,
	'vMerge' => 'restart',
	'valign' => 'center',
);

// 字體置中對其
$TextStart = array('alignment' => \PhpOffice\PhpWord\SimpleType\Jc::START);
$TextRun = array('alignment' => \PhpOffice\PhpWord\SimpleType\Jc::CENTER);
$TextEnd = array('alignment' => \PhpOffice\PhpWord\SimpleType\Jc::END);

// Header start
if($data['Code'] == 200) {
	$TitleRow = $TitleTable->addRow();
	$TitleRow->addCell($width, $DashedStyle)->addTextRun($TextEnd)->addText('當事人名稱：', $DashedStyle);
	$TitleRow->addCell($width, $DashedContentStyle)->addTextRun($TextStart)->addText($Data['ClientName']);

	$TitleRow = $TitleTable->addRow();
	$TitleRow->addCell($width, $DashedStyle)->addTextRun($TextEnd)->addText('地址：');
	$TitleRow->addCell($width, $DashedContentStyle)->addTextRun($TextStart)->addText($Data['CompanyAddress']);

	$TitleRow = $TitleTable->addRow();
	$TitleRow->addCell($width, $DashedStyle)->addTextRun($TextEnd)->addText('法定代理人：');
	$TitleRow->addCell($width, $DashedContentStyle)->addTextRun($TextStart)->addText($Data['CompanyOwner']);

	$TitleRow = $TitleTable->addRow();
	$TitleRow->addCell($width, $DashedStyle)->addTextRun($TextEnd)->addText('案別：');
	$TitleRow->addCell($width, $DashedContentStyle)->addTextRun($TextStart)->addText($Data['CaseName']);

	$TitleRow = $TitleTable->addRow();
	$TitleRow->addCell($width, $DashedStyle)->addTextRun($TextEnd)->addText('顧問期間：');
	$TitleRow->addCell(2750, $DashedContentSingleStyle)->addTextRun($TextRun)->addText(ContractSEDate($Data['ContractSDate']));
	$TitleRow->addCell(1250, $DashedContentSingleStyle)->addTextRun($TextRun)->addText('至');
	$TitleRow->addCell(2750, $DashedContentSingleStyle)->addTextRun($TextRun)->addText(ContractSEDate($Data['ContractEDate']));

	$TitleRow = $TitleTable->addRow();
	$TitleRow->addCell($width)->addTextRun($TextEnd)->addText('本所編號：');
	$TitleRow->addCell($width, $DashedContentBottomStyle)->addTextRun($TextStart)->addText($Data['CaseNumber']);
}
// Header End

// Details Start 分成個人及不分類表格
$i = 0;
if ($Data['PaymentType'] == "P") {
	while(1) {
		if(isset($Data['HoursData'][$i])) {
			$HoursData = $Data['HoursData'][$i];
			$section->addTextBreak(1);
			$section->addText($HoursData['Name'].' '.$HoursData['Position'], $fontStyle , $paragraphStyle);

			// Table樣式
			$DetailsStyle = array(
				'borderSize' => 12, 
				'borderColor' => 'black',
				'alignment' => \PhpOffice\PhpWord\SimpleType\Jc::END ,
				'leftFromText'  => 1000,
				'cellMarginLeft' => 400, 
				'cellMarginRight' => 400
			);
			$phpWord->addTableStyle('Details Row Style', $DetailsStyle);
			$DetailsTable = $section->addTable('Details Row Style');
			// 字體樣式
			$CellCenter = array('vMerge' => 'restart', 'valign' => 'center');
			// 字體置中對其
			$TextRun = array('alignment' => \PhpOffice\PhpWord\SimpleType\Jc::CENTER);

			$DetailsRow = $DetailsTable->addRow();
			$DetailsRow->addCell($width, $CellCenter)->addTextRun($TextRun)->addText('期間');
			$DetailsRow->addCell($width, array('gridSpan' => 2, 'vMerge' => 'restart', 'valign' => 'center'))->addTextRun($TextRun)->addText(Period($Data['WorkDate']));

			$DetailsRow = $DetailsTable->addRow();
			$DetailsRow->addCell($width)->addTextRun($TextRun)->addText('日期');
			$DetailsRow->addCell($width)->addTextRun($TextRun)->addText('合計時數');
			// $DetailsRow->addCell($width)->addTextRun($TextRun)->addText('案別');
			$DetailsRow->addCell($width)->addTextRun($TextRun)->addText('內容');
			
			$j = 0;
			while(1) {
				if(isset($HoursData['HoursList'][$j])) {
					$HoursList = $HoursData['HoursList'][$j];
					$DetailsRow = $DetailsTable->addRow();
					$DetailsRow->addCell($width)->addTextRun($TextRun)->addText($HoursList['WorkDate']);
					$DetailsRow->addCell($width)->addTextRun($TextRun)->addText($HoursList['WorkHours']);
					// $DetailsRow->addCell($width)->addTextRun($TextRun)->addText($HoursList['CaseName']);
					$DetailsRow->addCell($width)->addText($HoursList['JobDescription']);
					$j++;
				} else {
					break;
				}
			}
			$DetailsRow = $DetailsTable->addRow();
			$DetailsRow->addCell($width, $CellCenter)->addTextRun($TextRun)->addText('共計時數');
			$DetailsRow->addCell($width, array('gridSpan' => 2, 'vMerge' => 'restart', 'valign' => 'center'))->addTextRun($TextRun)->addText($HoursData['Hours']);

			$section->addText('●'.Period($Data['WorkDate']).$HoursData['Name'].$HoursData['Position'].'之服務時數共計'.$HoursData['Hours'].'小時。', $fontStyle , $paragraphStyle);

			$i++;
		} else {
			break;
		}
	}
}

if ($Data['PaymentType'] == "All") {
	$section->addTextBreak(1);

	// Table樣式
	$DetailsStyle = array(
		'borderSize' => 12, 
		'borderColor' => 'black',
		'alignment' => \PhpOffice\PhpWord\SimpleType\Jc::END ,
		'leftFromText'  => 1000,
		'cellMarginLeft' => 400, 
		'cellMarginRight' => 400
	);
	$phpWord->addTableStyle('Details Row Style', $DetailsStyle);
	$DetailsTable = $section->addTable('Details Row Style');
	// 字體樣式
	$CellCenter = array('vMerge' => 'restart', 'valign' => 'center');
	// 字體置中對其
	$TextRun = array('alignment' => \PhpOffice\PhpWord\SimpleType\Jc::CENTER);

	$DetailsRow = $DetailsTable->addRow();
	$DetailsRow->addCell(2000, $CellCenter)->addTextRun($TextRun)->addText('期間');
	$DetailsRow->addCell(2000, array('gridSpan' => 3, 'vMerge' => 'restart', 'valign' => 'center'))->addTextRun($TextRun)->addText(Period($Data['WorkDate']));

	$DetailsRow = $DetailsTable->addRow();
	$DetailsRow->addCell(2500)->addTextRun($TextRun)->addText('日期');
	$DetailsRow->addCell(2000)->addTextRun($TextRun)->addText('負責人');
	$DetailsRow->addCell(2500)->addTextRun($TextRun)->addText('合計時數');
	// $DetailsRow->addCell(2000)->addTextRun($TextRun)->addText('案別');
	$DetailsRow->addCell(3000)->addTextRun($TextRun)->addText('內容');

	while(1) {
		if(isset($Data['HoursData'][$i])) {
			$HoursList = $Data['HoursData'][$i];
			$DetailsRow = $DetailsTable->addRow();
			$DetailsRow->addCell(2500)->addTextRun($TextRun)->addText($HoursList['WorkDate']);
			$DetailsRow->addCell(2000)->addTextRun($TextRun)->addText($HoursList['Name']);
			$DetailsRow->addCell(2500)->addTextRun($TextRun)->addText($HoursList['WorkHours']);
			// $DetailsRow->addCell(2000)->addTextRun($TextRun)->addText($HoursList['CaseName']);
			$DetailsRow->addCell(3000)->addText($HoursList['JobDescription']);
			$i++;
		} else {
			break;
		}
	}
	$DetailsRow = $DetailsTable->addRow();
	$DetailsRow->addCell(2000, $CellCenter)->addTextRun($TextRun)->addText('共計時數');
	$DetailsRow->addCell(2000, array('gridSpan' => 3, 'vMerge' => 'restart', 'valign' => 'center'))->addTextRun($TextRun)->addText(TotalHours($Data['HoursData'], $Data['PaymentType']));

	$section->addText('●'.Period($Data['WorkDate']).'服務時數共計'.TotalHours($Data['HoursData'], $Data['PaymentType']).'小時。', $fontStyle , $paragraphStyle);
}
// Details End

// 表格後結尾 start
if($data['Code'] == 200) {
	if($Data['PaymentType'] == 'P'){
		// Tatal Start
		$section->addTextBreak(1);
		$phpWord->addTitleStyle(1, array('bold' => true), array('spaceAfter' => 240));
		$section->addText('●'.Period($Data['WorkDate']).'之服務時數總計'.TotalHours($Data['HoursData'], $Data['PaymentType']).'小時', $fontStyle , $paragraphStyle);
		// Tatal End
	}

	// Count Start
	$DiscountText = Discount(TotalPrice($Data, ''),$Data["CasePrice"],$Data["ContractDiscount"],$Data["OverDiscount"],1);
	$DiscountNum = Discount(TotalPrice($Data, ''),$Data["CasePrice"],$Data["ContractDiscount"],$Data["OverDiscount"],0);
	$section->addTextBreak(1);

	if ($DiscountText === "無折扣") {
		$section->addText('●工作時數以下述計算後：', $fontStyle , $paragraphStyle);

	} else {
		$section->addText('●工作時數以下述計算後，再優惠'.$DiscountText.'折抵：', $fontStyle , $paragraphStyle);
	}
}
$i = 0;
if($Data['PaymentType'] == 'P'){
	while(1) {
		if(isset($Data['HoursData'][$i])) {
			$HoursData = $Data['HoursData'][$i];
			$section->addText($HoursData['Name'].$HoursData['Position'].'： '.$HoursData['Pay'].'元/時計算', $fontStyle , $paragraphStyle);
			$i++;
		} else {
			break;
		}
	}
}
if($Data['PaymentType'] == 'All'){
	$array = [];
	while(1) {
		if(isset($Data['HoursData'][$i]['Name'])) {
			$HoursData = $Data['HoursData'][$i];
			if (!in_array($HoursData['Name'], $array)) {
				$section->addText($HoursData['Name'].$HoursData['Position'].'： '.$HoursData['Pay'].'元/時計算', $fontStyle , $paragraphStyle);
				array_push($array, $HoursData['Name']);
			}
			$i++;
		} else {
			break;
		}
	}
}
// 表格後結尾 end

// 各律師費用計算 start
$i = 0;
if($Data['PaymentType'] == 'P'){
	while(1) {
		if(isset($Data['HoursData'][$i])) {
			$HoursData = $Data['HoursData'][$i];
			$CountStyle = array(
				'borderColor' => '999999', 
				'alignment' => \PhpOffice\PhpWord\SimpleType\Jc::END ,			
				'cellMarginLeft' => 200, 
				'cellMarginRight' => 200
			);
			$phpWord->addTableStyle('Count Row Style', $CountStyle);
			$CountTable = $section->addTable('Count Row Style');
			$CountWidth = 1428;

			$CountRow = $CountTable->addRow();
			$CountRow->addCell($CountWidth, $CellCenter)->addTextRun($TextRun)->addText($HoursData['Name'].Position($HoursData['Position']));
			$CountRow->addCell($CountWidth, $CellCenter)->addTextRun($TextRun)->addText($HoursData['Hours']);
			$CountRow->addCell($CountWidth, $CellCenter)->addTextRun($TextRun)->addText('時*');
			$CountRow->addCell($CountWidth, $CellCenter)->addTextRun($TextRun)->addText($HoursData['Pay']);
			$CountRow->addCell($CountWidth, $CellCenter)->addTextRun($TextRun)->addText('元/時');
			$CountRow->addCell($CountWidth, $CellCenter)->addTextRun($TextRun)->addText($HoursData['Hours']*$HoursData['Pay']);
			$CountRow->addCell($CountWidth, $CellCenter)->addTextRun($TextRun)->addText('元');
			$i++;
		} else {
			break;
		}
	}
	$CountRow = $CountTable->addRow();
	$CountRow->addCell($CountWidth, $CellCenter)->addTextRun($TextRun)->addText('合計');
	$CountRow->addCell($CountWidth, $CellCenter)->addTextRun($TextRun)->addText(TotalPrice($Data, $HoursData['Name']));
	$CountRow->addCell($CountWidth, $CellCenter)->addTextRun($TextRun)->addText('元;');

	if ($DiscountText === "無折扣") {
		$CountRow->addCell($CountWidth, array('gridSpan' => 2, 'vMerge' => 'restart', 'valign' => 'center'))->addTextRun($TextRun)->addText('');
	} else {
		$CountRow->addCell($CountWidth, array('gridSpan' => 2, 'vMerge' => 'restart', 'valign' => 'center'))->addTextRun($TextRun)->addText($DiscountText.'優惠');
	}
	
	$CountRow->addCell($CountWidth, $CellCenter)->addTextRun($TextRun)->addText(TotalPrice($Data, $HoursData['Name'])*$DiscountNum);
	$CountRow->addCell($CountWidth, $CellCenter)->addTextRun($TextRun)->addText('元');
}

if($Data['PaymentType'] == 'All'){
	$array = [];
	$CountPersonalTotalPrice = 0;
	while(1) {
		if(isset($Data['HoursData'][$i]['Name'])) {
			$HoursData = $Data['HoursData'][$i];
			if (!in_array($HoursData['Name'], $array)) {
				$CountStyle = array(
					'borderColor' => '999999', 
					'alignment' => \PhpOffice\PhpWord\SimpleType\Jc::END ,			
					'cellMarginLeft' => 200, 
					'cellMarginRight' => 200
				);
				$phpWord->addTableStyle('Count Row Style', $CountStyle);
				$CountTable = $section->addTable('Count Row Style');
				$CountWidth = 1428;
	
				$CountRow = $CountTable->addRow();
				$Counthours = PersonalHours($Data['HoursData'], $HoursData['Name']);
				$CountPersonalTotalPrice = $CountPersonalTotalPrice + $Counthours*$HoursData['Pay'];
				$CountRow->addCell($CountWidth, $CellCenter)->addTextRun($TextRun)->addText($HoursData['Name'].Position($HoursData['Position']));
				$CountRow->addCell($CountWidth, $CellCenter)->addTextRun($TextRun)->addText($Counthours);
				$CountRow->addCell($CountWidth, $CellCenter)->addTextRun($TextRun)->addText('時*');
				$CountRow->addCell($CountWidth, $CellCenter)->addTextRun($TextRun)->addText($HoursData['Pay']);
				$CountRow->addCell($CountWidth, $CellCenter)->addTextRun($TextRun)->addText('元/時');
				$CountRow->addCell($CountWidth, $CellCenter)->addTextRun($TextRun)->addText($Counthours*$HoursData['Pay']);
				$CountRow->addCell($CountWidth, $CellCenter)->addTextRun($TextRun)->addText('元');
				array_push($array, $HoursData['Name']);
			}
			$i++;
		} else {
			break;
		}
	}
	$CountRow = $CountTable->addRow();
	$CountRow->addCell($CountWidth, $CellCenter)->addTextRun($TextRun)->addText('合計');
	$CountRow->addCell($CountWidth, $CellCenter)->addTextRun($TextRun)->addText($CountPersonalTotalPrice);
	$CountRow->addCell($CountWidth, $CellCenter)->addTextRun($TextRun)->addText('元;');

	if ($DiscountText === "無折扣") {
		$CountRow->addCell($CountWidth, array('gridSpan' => 2, 'vMerge' => 'restart', 'valign' => 'center'))->addTextRun($TextRun)->addText('');
	} else {
		$CountRow->addCell($CountWidth, array('gridSpan' => 2, 'vMerge' => 'restart', 'valign' => 'center'))->addTextRun($TextRun)->addText($DiscountText.'優惠');
	}
	
	$CountRow->addCell($CountWidth, $CellCenter)->addTextRun($TextRun)->addText($CountPersonalTotalPrice*$DiscountNum);
	$CountRow->addCell($CountWidth, $CellCenter)->addTextRun($TextRun)->addText('元');
}
// 各律師費用計算 End

// 最後總結算 start
if($data['Code'] == 200) {
	// End
	$TotalCost = 0;
	if ($Data['PaymentType'] == 'P') {
		$TotalCost = TotalPrice($Data, '')*$DiscountNum;
	}
	if ($Data['PaymentType'] == 'All') {
		$TotalCost = $CountPersonalTotalPrice*$DiscountNum;
	}
	$section->addTextBreak(1);
	$TextEnd = $section->addTextRun($paragraphStyle);
	$TextEnd->addText('●'.Period($Data['WorkDate']).'之'.$Data['AppointType'].'時數總計為'.TotalHours($Data['HoursData'], $Data['PaymentType']).'小時，經上述計算，截至止，', array('bgColor'=>'C0C0C0'));
	$TextEnd->addText('顧問費用累計'.$TotalCost.'元。', array('underline' => 'single', 'bgColor'=>'C0C0C0'));
	// End

	// Save file
	// echo write($phpWord, basename(__FILE__, '.php'), $writers);
	write($phpWord, '時數報表', $writers);
	if (!CLI) {
		include_once 'Sample_Footer.php';
	}


	echo "<script language='javascript' type ='text/javascript'>"; 
	echo "window.location.href = 'results/時數報表.docx';";
	echo 'document.getElementById("print").innerHTML = "時數報表已下載完成。";';
	echo "</script>"; 
} else {
	echo "<script language='javascript' type ='text/javascript'>"; 
	echo 'document.getElementById("print").innerHTML = "時數報表下載失敗。";';
	echo "</script>"; 
}
// 最後總結算 end


?>