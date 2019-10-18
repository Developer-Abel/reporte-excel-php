<?php
require "config.php";
if (!isset($_SESSION['usuario'])) {
	header("Location: signin.php");
	exit;
}

error_reporting(E_ALL);
ini_set('display_errors', '1');

require "db/conexion.php";
require "models/solicitud.php";

$ObjS = new Solicitud();

require_once 'PHPExcel/PHPExcel.php';
$objPHPExcel = new PHPExcel();

$objPHPExcel->getProperties()->setCreator("Brb")
	->setLastModifiedBy("Brb")
	->setTitle("datos")
	->setSubject("Datos")
	->setDescription("Reporte General")
	->setKeywords("")
	->setCategory("");

$objPHPExcel->setActiveSheetIndex(0)
	->setCellValue('A2', 'Titulo (Dr./Dra.)')
	->setCellValue('B2', 'Nombre')
	->setCellValue('C2', 'Apellido Paterno')
	->setCellValue('D2', 'Apellido Materno')
	->setCellValue('E2', 'Universidad de egreso de la especialidad')
	->setCellValue('F2', 'Institución de residencia de la especialidad ')
	->setCellValue('G2', 'Institución dónde labora')
	->setCellValue('H2', 'Hospital privado dónde labora')
	->setCellValue('I2', 'R.F.C.')
	->setCellValue('J2', 'CURP')
	->setCellValue('K2', 'Cédula profesional de médico general')
	->setCellValue('L1', 'fecha de nacimiento')
	->setCellValue('L2', 'Día')
	->setCellValue('M2', 'Mes')
	->setCellValue('N2', 'Año')
	->setCellValue('O2', 'Nacionalidad')
	->setCellValue('P2', 'Estado donde radica')
	->setCellValue('Q2', 'Municipio')
	->setCellValue('R2', 'Género Femenino = F Masculino = M')
	->setCellValue('S1', 'Fecha elaboración certificado')
	->setCellValue('S2', 'Día')
	->setCellValue('T2', 'Mes')
	->setCellValue('U2', 'Año')
	->setCellValue('V1', 'Valido de')
	->setCellValue('V2', 'Día')
	->setCellValue('W2', 'Mes')
	->setCellValue('X2', 'Año') 
	->setCellValue('Y1', 'Valido a')
	->setCellValue('Y2', 'Día')
	->setCellValue('Z2', 'Mes')
	->setCellValue('AA2', 'Año') 
	->setCellValue('AB2', 'No. Certificado')
	->setCellValue('AC2', 'Libro')
	->setCellValue('AD2', 'Foja')
	->setCellValue('AE2', 'Título ') 
	->setCellValue('AE1', 'Presidente del consejo')
	->setCellValue('AF2', 'Presidente')
	->setCellValue('AG1', 'Responsable del consejo')
	->setCellValue('AG2', 'Título')
	->setCellValue('AH2', 'Responsable')
	->setCellValue('AI2', 'Costo')
	->setCellValue('AJ2', 'Email')
	->setCellValue('AK2', 'Observaciones')
	->setCellValue('AL2', 'Cédula de la especialidad')
;

$solicitudes = $ObjS->ReporteAll();
$i = 3;


foreach ($solicitudes as $solicitud) {
	$objPHPExcel->setActiveSheetIndex(0)
		->setCellValue('A' . $i, $solicitud->prefijo)
		->setCellValue('B' . $i, $solicitud->nombre)
		->setCellValue('C' . $i, $solicitud->apellidop)
		->setCellValue('D' . $i, $solicitud->apellidom)
		->setCellValue('E' . $i, $solicitud->institucion_academica)
		->setCellValue('F' . $i, $solicitud->institucion_academica)
		->setCellValue('G' . $i, $solicitud->hospital)
		->setCellValue('H' . $i, $solicitud->hospital)
		->setCellValue('I' . $i, $solicitud->rfc)
		->setCellValue('J' . $i, $solicitud->curp)
		->setCellValue('K' . $i, $solicitud->cedpro)
		->setCellValue('L' . $i, $solicitud->dianac)
		->setCellValue('M' . $i, $solicitud->mesnac)
		->setCellValue('N' . $i, $solicitud->añonac)
		->setCellValue('O' . $i, $solicitud->nacionalidad)
		->setCellValue('P' . $i, $solicitud->estado)
		->setCellValue('Q' . $i, $solicitud->delomun)
		->setCellValue('R' . $i, $solicitud->sexo)
		->setCellValue('S' . $i, $solicitud->diare)
		->setCellValue('T' . $i, $solicitud->mesre)
		->setCellValue('U' . $i, $solicitud->añore)
		->setCellValue('V' . $i, $solicitud->diaini)
		->setCellValue('W' . $i, $solicitud->mesini)
		->setCellValue('X' . $i, $solicitud->añoini)
		->setCellValue('Y' . $i, $solicitud->diafin)
		->setCellValue('Z' . $i, $solicitud->mesfin)
		->setCellValue('AA' . $i, $solicitud->añofin)
		->setCellValue('AB' . $i, $solicitud->certificado)
		->setCellValue('AC' . $i, $solicitud->libro)
		->setCellValue('AD' . $i, $solicitud->foja)
		->setCellValue('AE' . $i, "Dr")
		->setCellValue('AF' . $i, $solicitud->presidente)
		->setCellValue('AG' . $i, "Dr")
		->setCellValue('AH' . $i, $solicitud->responsable)
		->setCellValue('AI' . $i, "$ 9,690.00")
		->setCellValue('AJ' . $i, $solicitud->email)
		->setCellValue('AK' . $i, $solicitud->especialidad)
		->setCellValue('AL' . $i, $solicitud->especialidad);
		
		$objPHPExcel->getActiveSheet()->getStyle("A".$i.":AL".$i)->applyFromArray(centrar());
		$objPHPExcel->getActiveSheet()->getStyle("A".$i.":AL".$i)->getAlignment()->setWrapText(true);
	$i++;
}


//exit;



for ($f = 'A'; $f !== 'AM'; $f++) {
	$objPHPExcel->setActiveSheetIndex(0)
		->getColumnDimension($f)->setAutoSize(TRUE);
}



$objPHPExcel->getActiveSheet()->setTitle('Reporte general ');

$objPHPExcel->setActiveSheetIndex(0);

/**************************************************/

//color relleno
$objPHPExcel->getActiveSheet()->getStyle( 'A2:c2' )->applyFromArray(colorRelleno('2F75B5'));//AZUL
$objPHPExcel->getActiveSheet()->getStyle( 'D2' )->applyFromArray(colorRelleno('F2F2F2') );//BLANCO
$objPHPExcel->getActiveSheet()->getStyle( 'E2:F2' )->applyFromArray(colorRelleno('2F75B5')); 
$objPHPExcel->getActiveSheet()->getStyle( 'G2:H2' )->applyFromArray(colorRelleno('FFC000'));//AMARILLO
$objPHPExcel->getActiveSheet()->getStyle( 'I2:AB2' )->applyFromArray(colorRelleno('2F75B5')); 
$objPHPExcel->getActiveSheet()->getStyle( 'AC2:AD2' )->applyFromArray(colorRelleno('F2F2F2'));
$objPHPExcel->getActiveSheet()->getStyle( 'AE2:AI2' )->applyFromArray(colorRelleno('2F75B5'));
$objPHPExcel->getActiveSheet()->getStyle( 'AJ2' )->applyFromArray(colorRelleno('F2F2F2'));
$objPHPExcel->getActiveSheet()->getStyle( 'AK2' )->applyFromArray(colorRelleno('000000'));
$objPHPExcel->getActiveSheet()->getStyle( 'AL2' )->applyFromArray(colorRelleno('FFC000'));

$objPHPExcel->getActiveSheet()->getStyle( 'L1:N1' )->applyFromArray(colorRelleno('2F75B5'));
$objPHPExcel->getActiveSheet()->getStyle( 'S1:AA1' )->applyFromArray(colorRelleno('2F75B5'));
$objPHPExcel->getActiveSheet()->getStyle( 'AE1:AH1' )->applyFromArray(colorRelleno('2F75B5'));
//color letra
$objPHPExcel->getActiveSheet()->getStyle('A2:c2')->applyFromArray(colorLetra('FFFFFF'));//BALNCO
$objPHPExcel->getActiveSheet()->getStyle('D2')->applyFromArray(colorLetra('000000'));//NEGRO
$objPHPExcel->getActiveSheet()->getStyle('E2:F2')->applyFromArray(colorLetra('FFFFFF'));
$objPHPExcel->getActiveSheet()->getStyle('G2:H2')->applyFromArray(colorLetra('000000'));
$objPHPExcel->getActiveSheet()->getStyle('I2:AB2')->applyFromArray(colorLetra('FFFFFF'));
$objPHPExcel->getActiveSheet()->getStyle('AC2:AD2')->applyFromArray(colorLetra('000000'));
$objPHPExcel->getActiveSheet()->getStyle('AE2:AI2')->applyFromArray(colorLetra('FFFFFF'));
$objPHPExcel->getActiveSheet()->getStyle('AJ2')->applyFromArray(colorLetra('000000'));
$objPHPExcel->getActiveSheet()->getStyle('AK2')->applyFromArray(colorLetra('FFFFFF'));
$objPHPExcel->getActiveSheet()->getStyle('AJ2')->applyFromArray(colorLetra('000000'));

$objPHPExcel->getActiveSheet()->getStyle('L1:N1')->applyFromArray(colorLetra('FFFFFF'));
$objPHPExcel->getActiveSheet()->getStyle('S1:AA1')->applyFromArray(colorLetra('FFFFFF'));
$objPHPExcel->getActiveSheet()->getStyle('AE1:AH1')->applyFromArray(colorLetra('FFFFFF'));

 //alto fila
$objPHPExcel->getActiveSheet()->getRowDimension('1')->setRowHeight(45); 
$objPHPExcel->getActiveSheet()->getRowDimension('2')->setRowHeight(45);

//ancho
$objPHPExcel->getActiveSheet()->getColumnDimension('A')->setAutoSize(false); 
$objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth('10');
$objPHPExcel->getActiveSheet()->getColumnDimension('E')->setAutoSize(false); 
$objPHPExcel->getActiveSheet()->getColumnDimension('E')->setWidth('50');
$objPHPExcel->getActiveSheet()->getColumnDimension('F')->setAutoSize(false); 
$objPHPExcel->getActiveSheet()->getColumnDimension('F')->setWidth('50');
$objPHPExcel->getActiveSheet()->getColumnDimension('G')->setAutoSize(false); 
$objPHPExcel->getActiveSheet()->getColumnDimension('G')->setWidth('50');
$objPHPExcel->getActiveSheet()->getColumnDimension('H')->setAutoSize(false); 
$objPHPExcel->getActiveSheet()->getColumnDimension('H')->setWidth('50');
$objPHPExcel->getActiveSheet()->getColumnDimension('K')->setAutoSize(false); 
$objPHPExcel->getActiveSheet()->getColumnDimension('K')->setWidth('18');
$objPHPExcel->getActiveSheet()->getColumnDimension('O')->setAutoSize(false); 
$objPHPExcel->getActiveSheet()->getColumnDimension('O')->setWidth('15');
$objPHPExcel->getActiveSheet()->getColumnDimension('R')->setAutoSize(false); 
$objPHPExcel->getActiveSheet()->getColumnDimension('R')->setWidth('15');
$objPHPExcel->getActiveSheet()->getColumnDimension('AB')->setAutoSize(false); 
$objPHPExcel->getActiveSheet()->getColumnDimension('AB')->setWidth('15');
$objPHPExcel->getActiveSheet()->getColumnDimension('AE')->setAutoSize(false); 
$objPHPExcel->getActiveSheet()->getColumnDimension('AE')->setWidth('10');
$objPHPExcel->getActiveSheet()->getColumnDimension('AG')->setAutoSize(false); 
$objPHPExcel->getActiveSheet()->getColumnDimension('AG')->setWidth('10');
$objPHPExcel->getActiveSheet()->getColumnDimension('AI')->setAutoSize(false); 
$objPHPExcel->getActiveSheet()->getColumnDimension('AI')->setWidth('25');


//centrar
$objPHPExcel->getActiveSheet()->getStyle("A2:AL2")->applyFromArray(centrar());
$objPHPExcel->getActiveSheet()->getStyle("A1:AL1")->applyFromArray(centrar());

//combinar celdas
$objPHPExcel->getActiveSheet()->mergeCells('L1:N1');
$objPHPExcel->getActiveSheet()->mergeCells('S1:U1');
$objPHPExcel->getActiveSheet()->mergeCells('V1:X1');
$objPHPExcel->getActiveSheet()->mergeCells('Y1:AA1');
$objPHPExcel->getActiveSheet()->mergeCells('AE1:AF1');
$objPHPExcel->getActiveSheet()->mergeCells('AG1:AH1');

//ajustar al texto
$objPHPExcel->getActiveSheet()->getStyle('A1:AL1')->getAlignment()->setWrapText(true);
$objPHPExcel->getActiveSheet()->getStyle('A2')->getAlignment()->setWrapText(true);
$objPHPExcel->getActiveSheet()->getStyle('E2')->getAlignment()->setWrapText(true);
$objPHPExcel->getActiveSheet()->getStyle('F2')->getAlignment()->setWrapText(true);
$objPHPExcel->getActiveSheet()->getStyle('G2')->getAlignment()->setWrapText(true);
$objPHPExcel->getActiveSheet()->getStyle('H2')->getAlignment()->setWrapText(true);
$objPHPExcel->getActiveSheet()->getStyle('K2')->getAlignment()->setWrapText(true);
$objPHPExcel->getActiveSheet()->getStyle('O2')->getAlignment()->setWrapText(true);
$objPHPExcel->getActiveSheet()->getStyle('R2')->getAlignment()->setWrapText(true);
$objPHPExcel->getActiveSheet()->getStyle('AB2')->getAlignment()->setWrapText(true);
$objPHPExcel->getActiveSheet()->getStyle('AE2')->getAlignment()->setWrapText(true);
$objPHPExcel->getActiveSheet()->getStyle('AG2')->getAlignment()->setWrapText(true);
$objPHPExcel->getActiveSheet()->getStyle('AI2')->getAlignment()->setWrapText(true);

/**************************************************/
$objPHPExcel->getActiveSheet(0)->freezePaneByColumnAndRow(0, 2);
header('Content-Type: application/vnd.ms-excel; charset=utf-8');
header('Content-Disposition: attachment;filename="Reporte general.xlsx"');
header('Cache-Control: max-age=0');

$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
$objWriter->save('php://output');

function colorLetra($colorH){
	$styleArray = array(
	        'font'  => array(
	        'bold'  => true,
	        'color' => array('rgb' => $colorH),
	        'size'  => 11,
	        'name'  => 'calibri'
	    )
    );
    return $styleArray;
}

function centrar(){
	$style = array( 'alignment' => array( 
	'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER, 
	'vertical' => PHPExcel_Style_Alignment::VERTICAL_CENTER
		) 
	);
	return $style;
}
function centrar2(){
	$style = array( 'alignment' => array( 
	'vertical' => PHPExcel_Style_Alignment::VERTICAL_CENTER
		) 
	);
	return $style;
}

function colorRelleno($color){
	$color = array(
        'fill' => array(
            'type' => PHPExcel_Style_Fill::FILL_SOLID,
            'color' => array('rgb' => $color)
	        )
	    );
    return $color;
}

?>