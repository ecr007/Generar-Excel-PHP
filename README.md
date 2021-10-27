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

## Nuevo

```php
public function download($id)
    {
        
        $record = Event::find($id);

        if (is_null($record)) {
            return back()->with('error',__('msj.str_not_found'));
        }
        
        $user_without_badge = [];

        if(count($record->users) > 0){

            foreach ($record->users as $guest) {
                
                $exists = Storage::disk(env('APP_DISK'))->exists('images/badges/'.$record->badge->folder.'/'.$guest->user['email'].'.png');

                if(!$exists){
                    array_push($user_without_badge, $guest->user);
                }
            }
        }

        if (count($user_without_badge) > 0) {

            $abc = ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z'];

            $titles = ['Fullname','Email'];

            // Redefinir invitados con sus url
            $dbInfo = [];
            
            foreach ($user_without_badge as $key) {

                $item = [
                    'fullname' => $key->firstname.' '.$key->lastname,
                    'email' => $key->email,
                ];

                array_push($dbInfo, $item);
            }

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

            $file = new Spreadsheet();

            $file->getProperties()
            ->setCreator('GathR')
            ->setTitle('GathR Excel Generator')
            ->setLastModifiedBy(Auth::user()->email)
            ->setDescription('GathR Excel Generator')
            ->setSubject('GathR Excel Generator')
            ->setKeywords('gathe office generate invite')
            ->setCategory('all');

            $page = $file->getSheet(0);
            $title = Str::slug("Missing guest badges");
            $page->setTitle($title);

            for ($i=0; $i < count($titles); $i++) { 
                $page->setCellValue($abc[$i].'1',$titles[$i]);
            }

            $page->fromArray($data, ' ', 'A2');


            $header = 'A1:'.$abc[count($titles)-1].'1';
            $page->getStyle($header)->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('00ffff00');
            
            $style = [
                'font' => array('bold' => true),
                'alignment' => array('horizontal' => Alignment::HORIZONTAL_CENTER),
            ];

            $page->getStyle($header)->applyFromArray($style);

            for ($col = ord('A'); $col <= ord($abc[count($titles)-1]); $col++){
                $page->getColumnDimension(chr($col))->setAutoSize(true);
            }

            // Redirect output to a client’s web browser (Xlsx)
            header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
            header('Content-Disposition: attachment;filename="'.$title.'.xlsx"');
            header('Cache-Control: max-age=0');
            // If you're serving to IE 9, then the following may be needed
            header('Cache-Control: max-age=1');

            // If you're serving to IE over SSL, then the following may be needed
            header('Expires: Mon, 26 Jul 1997 05:00:00 GMT'); // Date in the past
            header('Last-Modified: ' . gmdate('D, d M Y H:i:s') . ' GMT'); // always modified
            header('Cache-Control: cache, must-revalidate'); // HTTP/1.1
            header('Pragma: public'); // HTTP/1.0

            $writer = IOFactory::createWriter($file, 'Xlsx');
            $writer->save('php://output');
            exit;
        }

        return back()->with('error',"All guests now have badges.");
    }
```
