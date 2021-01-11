<?php
require_once 'Classes/PHPExcel.php';

$fechaActual= date('d-m-Y');
$objetoExcel = new PHPExcel();
$columnas= array("A","B","C","D","E","F");
$cabeceras= array("#","ID","Fecha","Libras","Pies Cúbicos","Total");


$objetoExcel->getProperties()
        ->setCreator("TechTeks")
        //->setLastModifiedBy("Códigos de Programación")
        ->setTitle("Reporte")
        //->setSubject("Documento de prueba")
        ->setDescription("Documento generado con catorce")
        //->setKeywords("excel phpexcel php")
        ->setCategory("Reporte");
$objetoExcel->setActiveSheetIndex(0);
$objetoExcel->getActiveSheet()->setTitle('Hoja 1');

// columnas[A,B,C...]

// Ajusta el tamaño de las celdas cuando hay mucho texto horizontal
foreach(range('A','F') as $columnaID) {
    $objetoExcel->getActiveSheet()->getColumnDimension($columnaID)
        ->setAutoSize(true);
}



// Agregando estilo------------------------------------------------------------------------------------
// Estilo textos
$styleArray = array(
    'font'  => array(
        'bold'  => true,
        'color' => array('rgb' => '30475e'),
        'size'  => 13,
        'name'  => 'Arial'
    ));

$objetoExcel->getDefaultStyle()->applyFromArray($styleArray);



// Estilo celdas
$objetoExcel->getActiveSheet()->getDefaultStyle()->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

// terminar de editar, no funciona con getDefaultStyle()
$objetoExcel->getActiveSheet()->getStyle("A1:E1")->applyFromArray(
    array(
        'borders' => array(
            'allborders' => array(
                'style' => PHPExcel_Style_Border::BORDER_THICK,
                'color' => array('rgb' => 'bdbbbb')
            )
        )
    )
);

$objetoExcel->getActiveSheet()->getStyle("A1:E1")->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
$objetoExcel->getActiveSheet()->getStyle("A1:E1")->getFill()->getStartColor()->setARGB('00f0f0f0');



//Cabeceras

for($i=0; $i<=count($cabeceras)-1; $i++)
      {     
      $objetoExcel->getActiveSheet()->setCellValue($columnas[$i]."1",$cabeceras[$i]);
    }




$fechaTransac= array("2020-11-27","2020-12-22","2021-01-02");
$datos = array(
            array(
                array(1,"2020-11-27", 1, 0.0, 3.0, 30)),

            array(
                array(2,"2020-11-27", 3, 0.0, 35.0, 430.41),
                array(3,"2020-11-27", 4, 0.0, 10.0, 138.50)),

            array(
                array(4,"2020-11-27", 2, 12.0, 5.0, 83.00)),
        );

/*
foreach($datos as $datos2)
    {
    foreach($datos2 as $datos3)
        {
        foreach($datos3 as $dato)
            {
            echo $fechaTransac[$i];
            echo "$dato ";
            $i++;
            }
        echo "<br>";
        }
    echo "<br>";
    }
*/

$contador=2;
// for para la fecha
for($i=0; $i<count($datos); $i++)
      {
        $objetoExcel->getActiveSheet()->mergeCells("A"."$contador".":F"."$contador")->setCellValue("A"."$contador",$fechaTransac[$i]);


        $objetoExcel->getActiveSheet()->getStyle("A"."$contador".":F"."$contador")->
        getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);

        $objetoExcel->getActiveSheet()->getStyle("A"."$contador".":F"."$contador")->
        getFill()->getStartColor()->setARGB('00ccccccc');
        $contador++;


    //For para el bloque de datos de cada fecha
        for ($f=0; $f < count($datos[$i]); $f++) {

            //for para cada celda de datos
            for ($g=0; $g < count($datos[$i][$f]); $g++) {
                $objetoExcel->getActiveSheet()->
                setCellValue($columnas[$g]."$contador",$datos[$i][$f][$g]);

        $objetoExcel->getActiveSheet()->getStyle("A"."$contador".":F"."$contador")->
        getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);

        $objetoExcel->getActiveSheet()->getStyle("A"."$contador".":F"."$contador")->
        getFill()->getStartColor()->setARGB('00f0f0f0');

            }
            $contador++;
    }
}


/* usar en caso que se necesite un link
$objetoExcel->getActiveSheet()->setCellValue('E26', 'www.phpexcel.net');
$objetoExcel->getActiveSheet()->getCell('E26')->getHyperlink()->setUrl('http://www.phpexcel.net');
*/

//-----------------------------------------------------------------------------------------------------

header('Content-Type: application/vnd.ms-excel');
header("Content-Disposition: attachment;filename=Excel ".$fechaActual.".xls");
header('Cache-Control: max-age=0');
    
$objWriter = PHPExcel_IOFactory::createWriter($objetoExcel, 'Excel5');
$objWriter->save('php://output');
?>