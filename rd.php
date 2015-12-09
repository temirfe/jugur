<?php
error_reporting(E_ALL);
ini_set('display_errors', TRUE);
ini_set('display_startup_errors', TRUE);
$username = "temirbek_kivano";
$password = "temirbek85fj";
$hostname = "127.0.0.1";
$date=date('Y-m-d');
$dir=dirname(__FILE__)."/excel/".$date;
$report_dir=dirname(__FILE__)."/report/".$date;
echo phpinfo();
//connection to the database
$dbhandle = mysql_connect($hostname, $username, $password)
    or die("Unable to connect to MySQL");

$selected = mysql_select_db("temirbek_kivano",$dbhandle)
    or die("Could not select examples");

mysql_set_charset("utf8");

$GLOBALS['dbh'] = new PDO('mysql:host=127.0.0.1;dbname=temirbek_kivano;charset=utf8', $username, $password);

//$dbh->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
//saveexcel();
//parsexcelsimple($dir);;
//priceAPI($date);
//dmail($date);
//scantest($dir);
//scandisk($dir);
//importProducts();
//noSku();
//importSkus();
//sku2();
//close the connection
mysql_close($dbhandle);
//dbhtest();
function huyak()
{
    $dbh=$GLOBALS['dbh'];

    /*$stmt = $dbh->prepare("SELECT product_id FROM thead WHERE product_id<>0");
    if ($stmt->execute()) {
        while ($row = $stmt->fetch()) {
            $pid=$row['product_id'];
            $product='';
            $product=$dbh->query("SELECT product_id FROM product WHERE product_id='{$pid}'")->fetch(PDO::FETCH_ASSOC);
            if(!$product) echo $pid.",";
        }
    }*/
}
//huyak();


function dbhtest()
{
    $arr=array();
    $dbh=$GLOBALS['dbh'];

    $stmt = $dbh->prepare("SELECT product_id, title FROM sku WHERE sender='asdf@asdf.com'");
    if ($stmt->execute()) {
        echo 'execute';
        while ($row = $stmt->fetch()) {
            //echo $row['title'];
            // echo ' price: '.$row['price'].'<br />';
            $arr[]=$row['title'];
            echo 'while';
        }
    }
    $result=$dbh->query("SELECT id FROM product WHERE product_id='9000'")->fetch();
    if($result) echo 'arr'; else echo 'no';
}

function scandisk($dir, $sendermail=false)
{
    //mysql_query("INSERT INTO temir (text) VALUES('asdf')");
    $files = scandir($dir, 1);
    if(is_file($file=$dir.'/'.$files[0]))
    {

        //parsexcelsimple($file);
        $info = pathinfo($file);
        if(!$sendermail)
        {
            $sender=explode('--',$info['filename']);
            if(isset($sender[1])) $sendermail=$sender[1]; else $sendermail='';

            /*if($sendermail=='alla-ultra@mail.ru')
            {
                $newname=$dir.'/the--alla-ultra@mail.ru--file.xml';
                rename($file, $newname); //because alla sends xml file with xls extension
                $file=$newname;
            }*/
        }

        $path = pathinfo(realpath($file), PATHINFO_DIRNAME);
        if ($info["extension"] == "xls" || $info["extension"] == "xlsx"  || $info["extension"] == "xml")
        {
            //file from alla-ultra is saved as xls and it's corrupt. to fix it we use convert() in dmail() and skip the corrupted one here
            //file from elenaultra is saved as skip and it's kinda also corrupt
            if (strpos($info["basename"],'alla-ultra@mail.ru--file.xls') === false || strpos($info["basename"],'skip') === false) {
                parsexcel3($file, $sendermail);
                $sendermail='';
                //parsexcelsimple($file);
            }
            else{$sendermail=false;}

        }
        elseif($info["extension"] == "zip")
        {
            $zip = new ZipArchive;
            if ($zip->open($file) === TRUE) {
                $zip->extractTo($path);
                $zip->close();
            } else {
                die("Can't open zip archive");
            }
        }
        elseif($info["extension"] == "rar")
        {
            $rar_file = rar_open($file) or die("Can't open Rar archive");
            $entries = rar_list($rar_file);
            foreach ($entries as $entry) {
                $entry->extract($path);
            }
            rar_close($rar_file);
        }
        unlink($file);
        scandisk($dir,$sendermail);
    }
}

//this one checks the parsing of excel
function parsexcel3($file, $sendermail)
{

    $dbh=$GLOBALS['dbh'];
    $date=date('Y-m-d');
    if (!file_exists($file)) {
        exit('No file yoba');
    }
    /* ------ Здесь мы готовим слова для заголовок столбцов ------*/
    $match_title=array();
    $match_price=array();

    $stmt = $dbh->prepare("SELECT title, price, note FROM thead");
    if ($stmt->execute()) {
        while ($row = $stmt->fetch()) {
            if($row['title'] && !in_array($row['title'],$match_title)) $match_title[]=$row['title'];
            if($row['price'] && !in_array($row['price'],$match_price)) $match_price[]=$row['price'];
        }
    }

    /*-----------------*/

    require_once dirname(__FILE__) . '/Classes/PHPExcel/IOFactory.php';
    $file=mb_convert_encoding($file, 'Windows-1251', 'UTF-8');
    $objReader = PHPExcel_IOFactory::createReaderForFile($file);
    $objReader->setReadDataOnly(true);
    $objPHPExcel=$objReader->load($file);

    $objWorksheet = $objPHPExcel->getActiveSheet();

    $highestRow = $objWorksheet->getHighestRow(); // e.g. 10
    $highestColumn = $objWorksheet->getHighestColumn(); // e.g 'F'
    $columnIndex=PHPExcel_Cell::stringFromColumnIndex($highestColumn);
    $highestColumnIndex = PHPExcel_Cell::columnIndexFromString($highestColumn); // e.g. 5


    for ($row = 1; $row <= $highestRow; ++$row) {
        $title=''; $price='';
        for ($col = 0; $col <= $highestColumnIndex; ++$col) {
            $curval=$objWorksheet->getCellByColumnAndRow($col, $row)->getCalculatedValue();
            //echo "-->".$curval."<--";
            if(!isset($tcolumn) && (in_array($curval,$match_title) || strpos($curval, "Товар/Склад")!== false))
            {$tcolumn=$col;}
            if(!isset($prcolumn) && (in_array($curval,$match_price)))
            {$prcolumn=$col;}
            if (isset($tcolumn) && isset($prcolumn)) {
                if($col==$tcolumn && $curval) $title=$curval;
                elseif($col==$prcolumn && $curval) $price=$curval;
            }
        }
        //echo "<br />";

        if($title && $price)
        {
            echo 'title: '.$title." price:".$price."</br>";
        }
    }
}

function saveexcel()
{
    $date=date("Y-m-d");
    require_once dirname(__FILE__) . '/Classes/PHPExcel/IOFactory.php';
    $arrayData = array(
        array(NULL, 2010, 2011, 2012),
        array('Q1',   12,   15,   21),
        array('Q2',   56,   73,   86),
        array('Q3',   52,   61,   69),
        array('Q4',   30,   32,    0),
    );
    $objPHPExcel = new PHPExcel();
    $objPHPExcel->getActiveSheet()
        ->fromArray(
        $arrayData,  // The data to set
        NULL,        // Array values with this value will not be set
        'C3'         // Top left coordinate of the worksheet range where
    //    we want to set these values (default is A1)
    );
    $objPHPExcel->getActiveSheet()->getStyle('B3:B7')->getFill()
        ->setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        ->getStartColor()->setARGB('FFFF0000');
    $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
    $objWriter->save("report/".$date."/test.xlsx");

}

function NoSku()
{
    $dbh=$GLOBALS['dbh'];
    $db_rows=$dbh->query("SELECT title FROM sku WHERE product_id='0' AND id BETWEEN 1001 AND 2000")->fetchAll(PDO::FETCH_ASSOC);
    $products=array();
    foreach($db_rows as $prod)
    {
        $products[]=array('title'=>$prod['title']);
    }
    //echo count($db_rows);
    $json=json_encode($products);

    header('Content-type: application/json');
    echo $json;
}

function skutoctgry()
{
    $dbh=$GLOBALS['dbh'];
    $db_rows=$dbh->query("SELECT id, title FROM sku WHERE product_id=''")->fetchAll(PDO::FETCH_ASSOC);



    $stmt = $dbh->prepare("SELECT product_id, sku FROM product");
    if ($stmt->execute()) {
        while ($prod = $stmt->fetch()) {
            $products[]=array('product_id'=>$prod['product_id'], 'sku'=>$prod['sku']);
        }
    }

    foreach($products as $producttable)
    {
        $skus=explode(';;',$producttable['sku']);
        $producttable_product_id=$producttable['product_id'];

        foreach($skus as $sku)
        {
            $productsku=strtolower(preg_replace("/\s/", "", $sku));

            foreach($db_rows as $db_row)
            {
                $skutable_row_id=$db_row['id'];
                $linesku=strtolower(preg_replace("/\s/", "", $db_row['title']));
                if($linesku ==$productsku)
                {
                    $dbh->exec("UPDATE sku SET product_id='{$producttable_product_id}' WHERE id='{$skutable_row_id}'");
                }
            }

        }
    }
}

function convert($options = array(), $inputfile = NULL, $outputfile = NULL, &$result = NULL) {
    $ch = curl_init();
    curl_setopt($ch, CURLOPT_FOLLOWLOCATION, true);
    curl_setopt($ch, CURLOPT_URL, "https://api.cloudconvert.org/convert");
    curl_setopt($ch, CURLOPT_POST, true);
    if ($inputfile !== NULL)
        $options = array_merge(array('file' =>  '@' . $inputfile), $options);
    curl_setopt($ch, CURLOPT_POSTFIELDS, http_build_query($options));
    curl_setopt($ch, CURLOPT_RETURNTRANSFER, true);
    curl_setopt($ch, CURLOPT_TIMEOUT, 300);

    // If you have SSL cert errors, try to disable SSL verifyer.
    //curl_setopt($ch, CURLOPT_SSL_VERIFYPEER,false);

    $output = curl_exec($ch);

    $http_status = curl_getinfo($ch, CURLINFO_HTTP_CODE);
    $content_type = curl_getinfo($ch, CURLINFO_CONTENT_TYPE);
    $curlerr = curl_error($ch);
    curl_close($ch);

    if ($curlerr && $result !== NULL) {
        $result = array('error' => $curlerr);
    } elseif (strpos($content_type, "application/json") === 0 && $result !== NULL) {
        $result = @json_decode($output, true);
    } elseif ($http_status == 200 && $outputfile !== NULL) {
        $file = fopen($outputfile, "w+");
        fputs($file, $output);
        fclose($file);
    }

    return $http_status == 200;
}

function importSkus()
{
    $dbh=$GLOBALS['dbh'];
    $string = file_get_contents("http://kivano.kg/product/MyImport");
    //$string = file_get_contents("http://kivano.com.kg/product/productAPI");
    $json_a = json_decode($string, true);

    if ($json_a) {
        $db_rows=$dbh->query("SELECT id, title FROM sku WHERE product_id=0")->fetchAll(PDO::FETCH_ASSOC);

        foreach ($json_a as $key => $value)
        {
            $id = $value['id'];
            $commonsku=$value['commonsku'];

            $skus=explode(';;',$commonsku);
            foreach($skus as $sku)
            {
                $productsku=strtolower(preg_replace("/\s/", "", $sku));

                foreach($db_rows as $db_row)
                {
                    $skutable_row_id=$db_row['id'];
                    $linesku=strtolower(preg_replace("/\s/", "", $db_row['title']));
                    if($linesku ==$productsku)
                    {
                        $dbh->exec("UPDATE sku SET product_id='{$id}' WHERE id='{$skutable_row_id}'");
                    }
                }

            }

            //if(mysql_affected_rows()==0)
            //mysql_query("INSERT INTO product (product_id, sku, category_id) VALUES('{$id}', '{$sku}', '{$category}')");

            //$newads[] = array('post_id' => $post_id, 'full_title' => $title, 'body' => $body, 'imgsrc'=>$value['imgsrc']);
        }
    }
}

function Sku2(){
    $dbh=$GLOBALS['dbh'];
    $db_rows=$dbh->query("SELECT id, title FROM sku WHERE product_id='0'")->fetchAll(PDO::FETCH_ASSOC);
    foreach($db_rows as $db_row)
    {
        $id=$db_row['id'];
        $title=$db_row['title'];
        $sth = $dbh->prepare("SELECT product_id FROM sku WHERE product_id<>'0' AND title=:title");
        $sth->bindParam(':title', $title, PDO::PARAM_STR);
        $sth->execute();
        $row=$sth->fetch(PDO::FETCH_ASSOC);

        if($pid=$row['product_id'])
        {
            $dbh->exec("UPDATE sku SET product_id='{$pid}' WHERE id='{$id}'");
        }
    }
}

function dmail($date)
{
    //mysql_query("INSERT INTO temir (text) VALUES('cron huyak from dmail!')");
    /* connect to gmail with your credentials */
    $hostname = '{imap.gmail.com:993/imap/ssl/novalidate-cert}INBOX';
    $username = 'kivanoprice@gmail.com'; # e.g somebody@gmail.com
    $password = 'kivanokivano';
    $dir=dirname(__FILE__).DIRECTORY_SEPARATOR."excel".DIRECTORY_SEPARATOR.$date.DIRECTORY_SEPARATOR;

    /* try to connect */
    $inbox = imap_open($hostname,$username,$password) or die('Cannot connect to Gmail: ' . imap_last_error());

    /* get all new emails. If set to 'ALL' instead
    * of 'NEW' retrieves all the emails, but can be
    * resource intensive, so the following variable,
    * $max_emails, puts the limit on the number of emails downloaded.
    *
    */
    $emails = imap_search($inbox,'UNSEEN');

    /* useful only if the above search is set to 'ALL' */
    $max_emails = 16;
    /* if any emails found, iterate through each email */
    if($emails) {
        $count = 1;
        /* put the newest emails on top */
        rsort($emails);
        /* for every email... */
        foreach($emails as $email_number)
        {
            /* get information specific to this email */
            //$overview = imap_fetch_overview($inbox,$email_number,0);

            /* get mail message */
            //$message = imap_fetchbody($inbox,$email_number,2);

            /* get sender host */
            $header = imap_headerinfo($inbox, $email_number);
            $fromaddr = $header->from[0]->mailbox."@".$header->from[0]->host;

            /* get mail structure */
            $structure = imap_fetchstructure($inbox, $email_number);
            $attachments = array();
            /* if any attachments found... */
            if(isset($structure->parts) && count($structure->parts))
            {
                for($i = 0; $i < count($structure->parts); $i++)
                {
                    $attachments[$i] = array(
                        'is_attachment' => false,
                        'filename' => '',
                        'name' => '',
                        'attachment' => '',
                        'from'=>$fromaddr
                    );

                    if($structure->parts[$i]->ifdparameters)
                    {
                        foreach($structure->parts[$i]->dparameters as $object)
                        {
                            if(strtolower($object->attribute) == 'filename')
                            {
                                $attachments[$i]['is_attachment'] = true;
                                $attachments[$i]['filename'] = $object->value;
                            }
                        }
                    }

                    if($structure->parts[$i]->ifparameters)
                    {
                        foreach($structure->parts[$i]->parameters as $object)
                        {
                            if(strtolower($object->attribute) == 'name')
                            {
                                $attachments[$i]['is_attachment'] = true;
                                $attachments[$i]['name'] = $object->value;
                            }
                        }
                    }

                    if($attachments[$i]['is_attachment'])
                    {
                        $attachments[$i]['attachment'] = imap_fetchbody($inbox, $email_number, $i+1);

                        /* 4 = QUOTED-PRINTABLE encoding */
                        if($structure->parts[$i]->encoding == 3)
                        {
                            $attachments[$i]['attachment'] = base64_decode($attachments[$i]['attachment']);
                        }
                        /* 3 = BASE64 encoding */
                        elseif($structure->parts[$i]->encoding == 4)
                        {
                            $attachments[$i]['attachment'] = quoted_printable_decode($attachments[$i]['attachment']);
                        }
                    }
                }
            }

            /* iterate through each attachment and save it */
            foreach($attachments as $attachment)
            {
                $alla_ultra=false;
                $elena_ultra=false;
                $b2bintermedia=false;
                $elena_dik=false;
                $rand=rand(1,100);
                if($attachment['is_attachment'] == 1)
                {
                    if($attachment['from']=='elena25@ultra.kg')
                        $filename='file.xlsx';
                    elseif($attachment['from']=='sales@technology.kg') //$attachment['name'] is strange so we just save
                        $filename='file.xls';
                    //$filename = $attachment['name'];
                    elseif(strpos($attachment['name'],'xlsx')!==false)
                        $filename='file.xlsx';
                    elseif(strpos($attachment['name'],'xls')!==false)
                        $filename='file.xls';
                    elseif(strpos($attachment['name'],'rar')!==false)
                        $filename='file.rar';
                    elseif(strpos($attachment['name'],'zip')!==false)
                        $filename='file.zip';
                    else $filename='';
                    if(empty($filename))
                    {
                        if(strpos($attachment['filename'],'xlsx')!==false)
                            $filename='file.xlsx';
                        elseif(strpos($attachment['filename'],'xls')!==false)
                            $filename='file.xls';
                        elseif(strpos($attachment['filename'],'rar')!==false)
                            $filename='file.rar';
                        elseif(strpos($attachment['filename'],'zip')!==false)
                            $filename='file.zip';
                        else $filename='';
                    }
                    //$filename = $attachment['filename'];

                    if($filename)
                    {
                        /* prefix the email number to the filename in case two emails
                    * have the attachment with the same file name.
                    */
                        $savename=$email_number.'-'.$rand."--".$fromaddr.'--'.$filename;
                        $fp = fopen($dir.$savename, "w+");
                        fwrite($fp, $attachment['attachment']);
                        fclose($fp);
                        if($fromaddr=='alla-ultra@mail.ru') $alla_ultra=$savename;
                        elseif($fromaddr=='elena_dik@inbox.ru') $elena_dik=$savename;
                        elseif($fromaddr=='b2b@intermedia.kg' || $fromaddr=='441111@intermedia.kg') $b2bintermedia=$savename;
                    }
                }
                if($alla_ultra)
                {
                    $inputfile = NULL;
                    $outputfile = $dir.$rand."--alla-ultra@mail.ru--f.xlsx";
                    $options = array(
                        "apikey" => "H2E2CI4hEuUzukrI7wKu6zhzVMdp1I78btb86z9zNqrbDSGfacLeCAbjiwdBSIax6TalvxpZnIuF2X9ln6zzNg",
                        "input" => "download",
                        "wait"=>true,
                        "download" => true,
                        "inputformat" => "xls",
                        "outputformat" => "xlsx",
                        "file"=>"http://95.85.45.110/excel/".$date."/".$alla_ultra
                    );
                    convert($options, $inputfile, $outputfile);
                }
                elseif($elena_dik)
                {

                    $inputfile = NULL;
                    $outputfile = $dir.$rand."--elena_dik@inbox.ru--f.xls";
                    $options = array(
                        "apikey" => "H2E2CI4hEuUzukrI7wKu6zhzVMdp1I78btb86z9zNqrbDSGfacLeCAbjiwdBSIax6TalvxpZnIuF2X9ln6zzNg",
                        "input" => "download",
                        "wait"=>true,
                        "download" => true,
                        "inputformat" => "xlsx",
                        "outputformat" => "xls",
                        "file"=>"http://95.85.45.110/excel/".$date."/".$elena_dik
                    );
                    convert($options, $inputfile, $outputfile);

                    @unlink($dir.$elena_dik);
                }
                elseif($b2bintermedia)
                {

                    $inputfile = NULL;
                    $outputfile = $dir.$rand."--b2b@intermedia.kg--f.xlsx";
                    $options = array(
                        "apikey" => "H2E2CI4hEuUzukrI7wKu6zhzVMdp1I78btb86z9zNqrbDSGfacLeCAbjiwdBSIax6TalvxpZnIuF2X9ln6zzNg",
                        "input" => "download",
                        "wait"=>true,
                        "download" => true,
                        "inputformat" => "xls",
                        "outputformat" => "xlsx",
                        "file"=>"http://95.85.45.110/excel/".$date."/".$b2bintermedia
                    );
                    convert($options, $inputfile, $outputfile);

                   @unlink($dir.$b2bintermedia);
                }
            }
            if($count++ >= $max_emails) break;
        }
    }
    /* close the connection */
    imap_close($inbox);
}
function priceAPI($date)
{
    $dbh=$GLOBALS['dbh'];
    //$forum=(int)$_GET['forum'];
    $db_rows=$dbh->query("SELECT product_id, category_id, price, sender, exrate FROM product WHERE changed='{$date}'")->fetchAll(PDO::FETCH_ASSOC);
    //$rows=Yii::app()->db->createCommand("SELECT id, commonsku, category_id FROM Product WHERE id=108")->queryAll();
    $products=array();
    foreach($db_rows as $prod)
    {
        $products[]=array('product_id'=>$prod['product_id'], 'category_id'=>$prod['category_id'], 'price'=>$prod['price'], 'sender'=>$prod['sender'], 'exrate'=>$prod['exrate']);
    }

    $json=json_encode($products);

    header('Content-type: application/json');
    echo $json;
}
function run2(){
    die();
    $dbh=$GLOBALS['dbh'];
    $rows=$dbh->query("SELECT product_id FROM sku WHERE product_id<>0 GROUP BY product_id")->fetchAll(PDO::FETCH_ASSOC);
    $i=1;
    foreach($rows as $row){
        echo $i.')'.$row['product_id'].'<br />';
        $i++;
        /*$prod=$dbh->query("SELECT id FROM product WHERE product_id='{$pid}'")->fetch(PDO::FETCH_ASSOC);
        if(!$prod){
            $dbh->exec("DELETE FROM sku WHERE product_id='{$pid}'");
        }*/
    }
}
?>
