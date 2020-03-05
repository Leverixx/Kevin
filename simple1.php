<!doctype html>
<html>
    <head>
        <meta charset="utf-8">
        <title>Test</title>
    </head>

<body>

<h1>PHP-Code wurde ausgeführt...check Email und "Termine.html"</h1>

<?php


require 'vendor/autoload.php';


//PHPspreadsheet-Klassen werden geladen (wahrscheinlich sind nicht alle notwendig)
use PhpOffice\PhpSpreadsheet\Writer\IWriter;
use PhpOffice\PhpSpreadsheet\Reader\IReader; 
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Writer\Xls;
use Symfony\Component\HttpFoundation\StreamedResponse;
use PhpOffice\PhpSpreadsheet\Writer as Writer;
use PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter;



//Excel-Datei wird aus Wiki-Eintrag eingelesen
$url = "http://192.168.123.75/wiki/images/b/b4/Termine.xlsx";
$filecontent = file_get_contents($url);
$tmpfname = tempnam(sys_get_temp_dir(), "tmpxls");
file_put_contents($tmpfname, $filecontent);



//Excel-Datei wird geladen
$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($tmpfname);
$sheet = $spreadsheet->getActiveSheet();



//aktuelles Datum wird im selben Format wie in der Excel-Tabelle angezeigt
$t = time();
$date = (date("d/m/Y", $t));
echo "Aktuelles Datum: ";
echo $date;




$hrow = $sheet->getHighestRow();//Anzahl der Reihen in der Excel-Datei werden in $hrow gespeichert

$hcoloumn = $sheet->getHighestColumn();//Anzahl der Spalten werden in $hcoloumn gespeichert





//Reihen, bei denen das Datum nicht dem aktuellen Datum entspricht, werden herausgefiltert
//PROBLEM: "removeRow($row)" führt zu fehlerhafter Filterung der Excel-Datei
for($row = 2; $row <= $hrow; $row++){ 

        
        $cellVal = $sheet->getCell('A'.$row)->getFormattedValue(); //Spalte "A" wird durchiteriert, Zellwerte in $cellVal gespeichert
        $cellVal = "0".$cellVal; //Formatierung des Datums
        $cellVal = date("d/m/Y", strtotime($cellVal));//Formatierung des Datums
      
        //Formatiertes Datum wird in die jeweilige Zelle gespeichert
        $sheet
        ->getCell('A'.$row)
        ->setValue($cellVal);
        
        
        
                if($cellVal !== $date){//Überprüfung ob das Datum in der jeweiligen Zelle ungleich dem aktuellen Datum ist
                
                    $sheet->getRowDimension($row)->setVisible(false); //Reihen, bei denen das Datum != $date werden ausgeblendet
                    //$sheet->removeRow($row);  //PROBLEM: Siehe oben
                    echo "<br>";
                    echo "row " . $row . " wird geloescht";
                    echo "<br>";
                } 
                
                    else{
                        echo "<br>";
                        echo "Datum gleich in row: ".$row . ".";
                        echo "<br>";
                
                    }

                //Beispiel für weitere Filterung
                $cellVal2 = $sheet->getCell('B'.$row)->getFormattedValue();//Spalte "B" wird durchiteriert, Zellwerte in $cellVal2 gespeichert
                
                if($cellVal2 == "Sarah Borjans"){
                
                    echo "Zeile " . $row . "enthält den Namen";
                    $sheet->getRowDimension($row)->setVisible(false);
                //$sheet->removeRow($row); //PROBLEM: Siehe oben
            }
    }







//gefilterte Excel-Tabelle wird als HTML-Datei in "Termine.html" gespeichert
$writer = new \PhpOffice\PhpSpreadsheet\Writer\HTML($spreadsheet); 
$writer->save("Termine.html"); 


















//HTML-Datei wird per E-Mail verschickt
use PHPMailer\PHPMailer\PHPMailer;
use PHPMailer\PHPMailer\Exception;

require 'C:\xampp\htdocs\phpspreadsheet\vendor\autoload.php'; 

$mail = new PHPMailer(TRUE);

/* Open the try/catch block. */
try {
    
    $mail->isSMTP();
    $mail->SMTPOptions = array(
        'ssl' => array(
        'verify_peer' => false,
        'verify_peer_name' => false,
        'allow_self_signed' => true
        )
        );
   /* Set the mail sender. */
   $mail->setFrom('kevinfichter30@gmail.com', 'Kevin');   

   /* Add a recipient. */
   $mail->addAddress('Kevin.Fichter@bbw-azubi.eu', 'Kevin');  

   /* Set the subject. */
   $mail->Subject = 'This is a sample Subject';

   /* Sets the body content type to HTML */
   $mail->isHTML(TRUE); 

   $mail->msgHTML(file_get_contents('Termine.html'), __DIR__);  

   $mail->Host = '192.168.136.17';
   
   $mail->Port = 25;

   /* Attach file */     
   $mail->addAttachment('Termine.html', 'Termine');

   /* send the mail. */
   $mail->send();
}
  



        catch (Exception $e)
        {
        /* PHPMailer exception. */
        echo $e->errorMessage();
        }
        catch (\Exception $e)
        {
        /* PHP exception */
        echo $e->getMessage();
        }





?>

</body>

</html>





