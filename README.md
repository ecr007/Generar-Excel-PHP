# Generar Excel PHP

## Dependencia

Codigo de ejemplo de la version: ```"phpoffice/phpexcel": "^1.8"```

Link: https://github.com/PHPOffice/PHPExcel
Link v2: https://github.com/PHPOffice/PhpSpreadsheet


```php
/**
 *
 * Esta funcion se usara para genrar archivos de excel
 * 
 * @param Array $dbInfo [Debe tener los resultados en orden segun los titulos]
 * @param Array $titulars [Nombre, Apellido, Edad, ... x]
 * @param String $excelTitle
 * @param String $rutaDestino
 * @return File
 */
function generateExcel($dbInfo,$titulars,$excelTitle,$rutaDestino)
{
	$abc = ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z'];

	$data = [];
	$i = 0;

	foreach ($dbInfo as $key) {
		
		$set = [];
		$info = array_keys($key);

		foreach ($info as $keyCh => $valueCh) {
			array_push($set, $key[$valueCh]);
		}

		$data[$i] = $set;
		$i++;
	}

	$file = new \PHPExcel();

	$file->getProperties()
	->setCreator('MBE')
	->setTitle('Ebox Web Excel Generator')
	->setLastModifiedBy($_SESSION['LOGIN_ADMIN']['loggedemail'])
	->setDescription('Ebox Web Excel Generator')
	->setSubject('Ebox Web Excel Generator')
	->setKeywords('excel ebox office generate mbe')
	->setCategory('all');

	$page = $file->getSheet(0);
	$page->setTitle($excelTitle);

	for ($i=0; $i < count($titulars); $i++) { 
		$page->setCellValue($abc[$i].'1',$titulars[$i]);
	}

	$page->fromArray($data, ' ', 'A2');


	$header = 'A1:'.$abc[count($titulars)-1].'1';
	$page->getStyle($header)->getFill()->setFillType(\PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setARGB('00ffff00');
	
	$style = [
    	'font' => array('bold' => true),
    	'alignment' => array('horizontal' => \PHPExcel_Style_Alignment::HORIZONTAL_CENTER),
    ];

	$page->getStyle($header)->applyFromArray($style);

	for ($col = ord('A'); $col <= ord($abc[count($titulars)-1]); $col++){
    	$page->getColumnDimension(chr($col))->setAutoSize(true);
	}

	$writer = \PHPExcel_IOFactory::createWriter($file, 'Excel2007');
            
	$writer->save($rutaDestino);
}
```
