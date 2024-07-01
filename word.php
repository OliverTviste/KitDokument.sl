<?php

require_once 'vendor/autoload.php';

$document = new \PhpOffice\PhpWord\TemplateProcessor('./temp.docx');

$uploadDir =  __DIR__;
$outputFile = 'review_full.docx';

// name html name Document

$surname = $_POST['surname'];
$name = $_POST['name'];
$patronymic = $_POST['patronymic'];
$citizenship = $_POST['citizenship'];
$BDATE = $_POST['BDATE'];
$BPLACE = $_POST['BPLACE'];
$ADDRESP = $_POST['ADDRESP'];
$ADDRESF = $_POST['ADDRESF'];
$INDEX = $_POST['INDEX'];
$PHONE = $_POST['PHONE'];
$OBRBTK = $_POST['OBRBTK'];
$UHSTT = $_POST['UHSTT'];
$FIRSTOBR = $_POST['FIRSTOBR'];
$SURNAMEF = $_POST['SURNAMEF'];
$NAMEF = $_POST['NAMEF'];
$PATRONYMICF = $_POST['PATRONYMICF'];
$PHONEF = $_POST['PHONEF'];
$SURNAMEM = $_POST['SURNAMEM'];
$NAMEM = $_POST['NAMEM'];
$PATRONYMICM = $_POST['PATRONYMICM'];
$PHONEM = $_POST['PHONEM'];
$PASS = $_POST['PASS'];
$PASN = $_POST['PASN'];
$PASV = $_POST['PASV'];
$PASD = $_POST['PASD'];
$INN = $_POST['INN'];
$SNILS = $_POST['SNILS'];
$DOBR = $_POST['DOBR'];
$ATTESTOTL = $_POST['ATTESTOTL'];
$DOBRD = $_POST['DOBRD'];
$DOBRV = $_POST['DOBRV'];
$DOBRN = $_POST['DOBRN'];
$HOSTEL = $_POST['HOSTEL'];
$DATEZD = $_POST['DATEZD'];
$PROF1 = $_POST['PROF1'];
$PROF2 = $_POST['PROF2'];
$PROF3 = $_POST['PROF3'];
$OBRFORM = $_POST['OBRFORM'];
$OBRPAY = $_POST['OBRPAY'];
$OBRLVL = $_POST['OBRLVL'];


//Document

$document->setValue('surname', $surname);
$document->setValue('name', $name);
$document->setValue('patronymic', $patronymic);
$document->setValue('citizenship', $citizenship);
$document->setValue('BDATE', $BDATE);
$document->setValue('BPLACE', $BPLACE);
$document->setValue('ADDRESP', $ADDRESP);
$document->setValue('ADDRESF', $ADDRESF);
$document->setValue('INDEX', $INDEX);
$document->setValue('PHONE', $PHONE);
$document->setValue('OBRBTK', $OBRBTK);
$document->setValue('UHSTT', $UHSTT);
$document->setValue('FIRSTOBR', $FIRSTOBR);
$document->setValue('SURNAMEF', $SURNAMEF);
$document->setValue('NAMEF', $NAMEF);
$document->setValue('PATRONYMICF', $PATRONYMICF);
$document->setValue('PHONEF', $PHONEF);
$document->setValue('SURNAMEM', $SURNAMEM);
$document->setValue('NAMEM', $NAMEM);
$document->setValue('PATRONYMICM', $PATRONYMICM);
$document->setValue('PHONEM', $PHONEM);
$document->setValue('PASS', $PASS);
$document->setValue('PASN', $PASN);
$document->setValue('PASV', $PASV);
$document->setValue('PASD', $PASD);
$document->setValue('INN', $INN);
$document->setValue('SNILS', $SNILS);
$document->setValue('DOBR', $DOBR);
$document->setValue('ATTESTOTL', $ATTESTOTL);
$document->setValue('DOBRD', $DOBRD);
$document->setValue('DOBRV', $DOBRV);
$document->setValue('DOBRN', $DOBRN);
$document->setValue('HOSTEL', $HOSTEL);
$document->setValue('DATEZD', $DATEZD);
$document->setValue('PROF1', $PROF1);
$document->setValue('PROF2', $PROF2);
$document->setValue('PROF3', $PROF3);
$document->setValue('OBRFORM', $OBRFORM);
$document->setValue('OBRPAY', $OBRPAY);
$document->setValue('OBRLVL', $OBRLVL);

//seve Document

$document->saveAs($outputFile );
//
//sending
//

$to  = "olivertvissthahaha@gmail.com" ; 

$user_email = $_POST['user_email'];

$subject = $surname = $_POST['surname']; 

$message ="Текст сообщения";

$filename = "$surname.docx";

$filepath = "review_full.docx";

$boundary = "--".md5(uniqid(time()));

$mailheaders = "MIME-Version: 1.0;\r\n"; 
$mailheaders .="Content-Type: multipart/mixed; boundary=\"$boundary\"\r\n";

$mailheaders .= "From: $user_email";

$multipart = "--$boundary\r\n"; 
$multipart .= "Content-Type: text/html; charset=windows-1251\r\n";
$multipart .= "Content-Transfer-Encoding: base64\r\n"; 

$fp = fopen($filepath,"r"); 
		if (!$fp) 
		{ 
			print "Не удается открыть файл22"; 
			exit(); 
		} 
$file = fread($fp, filesize($filepath)); 
fclose($fp); 

$message_part = "\r\n--$boundary\r\n"; 
$message_part .= "Content-Type: application/octet-stream; name=\"$filename\"\r\n";  
$message_part .= "Content-Transfer-Encoding: base64\r\n"; 
$message_part .= "Content-Disposition: attachment; filename=\"$filename\"\r\n"; 
$message_part .= chunk_split(base64_encode($file));
$message_part .= "\r\n--$boundary--\r\n";

$multipart .= $message_part;

mail($to,$subject,$multipart,$mailheaders);

?>