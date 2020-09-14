<?php
	//Incluimos librería y archivo de conexión
	require 'Classes/PHPExcel.php';
	require 'conexion.php';
	
	//Consulta
	$sql = "SELECT trani.quantidade, tran.data, tran.tipo_transacao, cli.nome, ven.nome, ite.quantidade, ite.valor FROM transacao_item trani INNER JOIN transacao tran ON trani.id_transacao = tran.id_trancasao INNER JOIN cliente cli ON cli.id_cliente = trani.id_transacaoitem INNER JOIN vendedor ven ON ven.id_vendedor = trani.id_transacaoitem INNER JOIN item ite ON trani.id_item = ite.id_item";
	$resultado = $mysqli->query($sql);
	$fila = 7; //Establecemos en que fila inciara a imprimir los datos
	
	$gdImage = imagecreatefrompng('images/logo.png');//Logotipo
	
	//Objeto de PHPExcel
	$objPHPExcel  = new PHPExcel();
	
	//Propiedades de Documento
	$objPHPExcel->getProperties()->setCreator("Ansony Martinez")->setDescription("Relatorio de vendas");
	
	//Establecemos la pestaña activa y nombre a la pestaña
	$objPHPExcel->setActiveSheetIndex(0);
	$objPHPExcel->getActiveSheet()->setTitle("Relatorio");
	
	$objDrawing = new PHPExcel_Worksheet_MemoryDrawing();
	$objDrawing->setName('Logotipo');
	$objDrawing->setDescription('Logotipo');
	$objDrawing->setImageResource($gdImage);
	$objDrawing->setRenderingFunction(PHPExcel_Worksheet_MemoryDrawing::RENDERING_PNG);
	$objDrawing->setMimeType(PHPExcel_Worksheet_MemoryDrawing::MIMETYPE_DEFAULT);
	$objDrawing->setHeight(100);
	$objDrawing->setCoordinates('A1');
	$objDrawing->setWorksheet($objPHPExcel->getActiveSheet());
	
	$estiloTituloReporte = array(
    'font' => array(
	'name'      => 'Arial',
	'bold'      => true,
	'italic'    => false,
	'strike'    => false,
	'size' =>13
    ),
    'fill' => array(
	'type'  => PHPExcel_Style_Fill::FILL_SOLID
	),
    'borders' => array(
	'allborders' => array(
	'style' => PHPExcel_Style_Border::BORDER_NONE
	)
    ),
    'alignment' => array(
	'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
	'vertical' => PHPExcel_Style_Alignment::VERTICAL_CENTER
    )
	);
	
	$estiloTituloColumnas = array(
    'font' => array(
	'name'  => 'Arial',
	'bold'  => true,
	'size' =>10,
	'color' => array(
	'rgb' => 'FFFFFF'
	)
    ),
    'fill' => array(
	'type' => PHPExcel_Style_Fill::FILL_SOLID,
	'color' => array('rgb' => '538DD5')
    ),
    'borders' => array(
	'allborders' => array(
	'style' => PHPExcel_Style_Border::BORDER_THIN
	)
    ),
    'alignment' =>  array(
	'horizontal'=> PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
	'vertical'  => PHPExcel_Style_Alignment::VERTICAL_CENTER
    )
	);
	
	$estiloInformacion = new PHPExcel_Style();
	$estiloInformacion->applyFromArray( array(
    'font' => array(
	'name'  => 'Arial',
	'color' => array(
	'rgb' => '000000'
	)
    ),
    'fill' => array(
	'type'  => PHPExcel_Style_Fill::FILL_SOLID
	),
    'borders' => array(
	'allborders' => array(
	'style' => PHPExcel_Style_Border::BORDER_THIN
	)
    ),
	'alignment' =>  array(
	'horizontal'=> PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
	'vertical'  => PHPExcel_Style_Alignment::VERTICAL_CENTER
    )
	));
	
	$objPHPExcel->getActiveSheet()->getStyle('A1:E4')->applyFromArray($estiloTituloReporte);
	$objPHPExcel->getActiveSheet()->getStyle('A6:O6')->applyFromArray($estiloTituloColumnas);
	
	$objPHPExcel->getActiveSheet()->setCellValue('B3', 'RELATORIO DE VENDAS');
	$objPHPExcel->getActiveSheet()->mergeCells('B3:D3');
	
	$objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth(10);
	$objPHPExcel->getActiveSheet()->setCellValue('A6', 'cliente');
	$objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth(30);
	$objPHPExcel->getActiveSheet()->setCellValue('B6', 'vendedor');
	$objPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth(10);
	$objPHPExcel->getActiveSheet()->setCellValue('C6', 'JAN');
	$objPHPExcel->getActiveSheet()->getColumnDimension('D')->setWidth(10);
	$objPHPExcel->getActiveSheet()->setCellValue('D6', 'FEV');
	$objPHPExcel->getActiveSheet()->getColumnDimension('E')->setWidth(10);
	$objPHPExcel->getActiveSheet()->setCellValue('E6', 'MAR');
	$objPHPExcel->getActiveSheet()->getColumnDimension('F')->setWidth(10);
	$objPHPExcel->getActiveSheet()->setCellValue('E6', 'ABR');
	$objPHPExcel->getActiveSheet()->getColumnDimension('G')->setWidth(10);
	$objPHPExcel->getActiveSheet()->setCellValue('E6', 'MAI');
	$objPHPExcel->getActiveSheet()->getColumnDimension('H')->setWidth(10);
	$objPHPExcel->getActiveSheet()->setCellValue('E6', 'JUN');
	$objPHPExcel->getActiveSheet()->getColumnDimension('I')->setWidth(10);
	$objPHPExcel->getActiveSheet()->setCellValue('E6', 'JUL');
	$objPHPExcel->getActiveSheet()->getColumnDimension('J')->setWidth(10);
	$objPHPExcel->getActiveSheet()->setCellValue('E6', 'AGOSTO');
	$objPHPExcel->getActiveSheet()->getColumnDimension('K')->setWidth(10);
	$objPHPExcel->getActiveSheet()->setCellValue('E6', 'SET');
	$objPHPExcel->getActiveSheet()->getColumnDimension('L')->setWidth(10);
	$objPHPExcel->getActiveSheet()->setCellValue('E6', 'OUT');
	$objPHPExcel->getActiveSheet()->getColumnDimension('N')->setWidth(10);
	$objPHPExcel->getActiveSheet()->setCellValue('E6', 'NOV');
	$objPHPExcel->getActiveSheet()->getColumnDimension('M')->setWidth(10);
	$objPHPExcel->getActiveSheet()->setCellValue('E6', 'DEZ');
	$objPHPExcel->getActiveSheet()->getColumnDimension('O')->setWidth(10);
	$objPHPExcel->getActiveSheet()->setCellValue('E6', 'TOTAL');
	//Recorremos los resultados de la consulta y los imprimimos
	while($rows = $resultado->fetch_assoc()){
		
		$objPHPExcel->getActiveSheet()->setCellValue('A'.$fila, $rows['nome']);
		$objPHPExcel->getActiveSheet()->setCellValue('B'.$fila, $rows['nome']);
		$objPHPExcel->getActiveSheet()->setCellValue('C'.$fila, $rows['data']);
		$objPHPExcel->getActiveSheet()->setCellValue('D'.$fila, $rows['existencia']);
		$objPHPExcel->getActiveSheet()->setCellValue('E'.$fila, $rows['existencia']);
		$objPHPExcel->getActiveSheet()->setCellValue('F'.$fila, $rows['existencia']);
		$objPHPExcel->getActiveSheet()->setCellValue('G'.$fila, $rows['existencia']);
		$objPHPExcel->getActiveSheet()->setCellValue('H'.$fila, $rows['existencia']);
		$objPHPExcel->getActiveSheet()->setCellValue('I'.$fila, $rows['existencia']);
		$objPHPExcel->getActiveSheet()->setCellValue('J'.$fila, $rows['existencia']);
		$objPHPExcel->getActiveSheet()->setCellValue('K'.$fila, $rows['existencia']);
		$objPHPExcel->getActiveSheet()->setCellValue('L'.$fila, $rows['existencia']);
		$objPHPExcel->getActiveSheet()->setCellValue('N'.$fila, $rows['existencia']);
		$objPHPExcel->getActiveSheet()->setCellValue('M'.$fila, $rows['existencia']);
		$objPHPExcel->getActiveSheet()->setCellValue('O'.$fila, '=C'.$fila.'*D'.$fila);
		
		$fila++; //Sumamos 1 para pasar a la siguiente fila
	}
	
	
	header("Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
	header('Content-Disposition: attachment;filename="Productos.xlsx"');
	header('Cache-Control: max-age=0');
	
	$writer->save('php://output');
	
?>