<?php
error_reporting(E_ALL);
ini_set('display_errors', TRUE);
ini_set('display_startup_errors', TRUE);

$username = "kivano";
$password = "bishkek86gq";
$hostname = "127.0.0.1";
$date=date('Y-m-d');
$dir=dirname(__FILE__)."/excel/".$date;
$report_dir=dirname(__FILE__)."/report/".$date;

if(!is_dir($dir))
{
    $oldmask = umask(0);
    mkdir($dir,0777);
    umask($oldmask);
}

if(!is_dir($report_dir))
{
    $oldmask = umask(0);
    mkdir($report_dir,0777);
    umask($oldmask);
}

//connection to the database
//$dbhandle = mysql_connect($hostname, $username, $password) or die("Unable to connect to MySQL");

//$selected = mysql_select_db("temirbek_kivano",$dbhandle) or die("Could not select examples");

//mysql_set_charset("utf8");

$GLOBALS['dbh'] = new PDO('mysql:host=127.0.0.1;dbname=jugur_kivano;charset=utf8', $username, $password);


if(isset($argv))
{
    if($argv[1]=='import')
    {
        importProducts();
    }
    elseif($argv[1]=='checkmail')
    {
        dmail($date);
    }
    elseif($argv[1]=='parse')
    {
        scandisk($dir);
    }
    elseif($argv[1]=='delete')
    {
        deleteDeleted();
    }
    elseif($argv[1]=='sku2')
    {
        sku2();
    }
    elseif($argv[1]=='test')
    {
        $dbh=$GLOBALS['dbh'];
        $dbh->prepare("INSERT INTO temir (text) VALUES ('test esje')")->execute();
    }
}
else{
    //dmail($date);
    //scandisk($dir);
    //sku2();
    //run();
    //run2();
}
//saveexcel();
//parsexcelsimple($dir);
//mysql_query("INSERT INTO temir (text) VALUES('CURDATE()')");

//$rows=mysql_query("SELECT product_id, category_id, price, sender, exrate FROM product WHERE changed='CURDATE()'");

//importProducts();


//priceAPI($date);
if(isset($_GET['param']))
{
    if($_GET['param']=='priceapi' && isset($_GET['date']))
    {
        if(isset($_GET['list']) && $_GET['list']=='all') $list_all=true;
        else $list_all=false; //see explanation on function declaration
        if(isset($_GET['change']) && $_GET['change']=='true') $change=true;
        else $change=false; //see explanation on function declaration
        $pdate=$_GET['date'];
        priceAPI($pdate,$change,$list_all);
    }

    elseif($_GET['param']=='reportdirs')
        showReportDirs();
    elseif($_GET['param']=='reportfiles')
    {
        if(isset($_GET['name']) && $_GET['name'])
            showReportFiles($_GET['name']);
    }
    elseif($_GET['param']=='userlist')
    {
        listUsers();
    }
}

//close the connection
//mysql_close($dbhandle);

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
                $rand=rand(1,100);
                $savename='';
                $to_be_converted=array();
                $extension='';
                $alt_ext='';
                if($attachment['is_attachment'] == 1)
                {
                    if($attachment['name'])
                    {
                        //sometimes cyrillic gets messy
                        mb_internal_encoding("UTF-8");
                        $attach_name=mb_decode_mimeheader($attachment['name']);
                    }
                    else{
                        mb_internal_encoding("UTF-8");
                        $attach_name=mb_decode_mimeheader($attachment['filename']);
                    }
                    //$filename = $attachment['name'];
                    if(strpos($attach_name,'xlsx')!==false){$filename='file.xlsx'; $extension='xlsx'; $alt_ext='xls';}
                    elseif(strpos($attach_name,'xls')!==false){$filename='file.xls'; $extension='xls'; $alt_ext='xlsx';}
                    elseif(strpos($attach_name,'rar')!==false){$filename='file.rar';}
                    elseif(strpos($attach_name,'zip')!==false){$filename='file.zip';}
                    else {$filename='';}

                    if($filename)
                    {
                        /* prefix the email number to the filename in case two emails
                    * have the attachment with the same file name.
                    */
                        $savename=$email_number.'-'.$rand."--".$fromaddr.'--'.$filename;
                        $fp = fopen($dir.$savename, "w+");
                        fwrite($fp, $attachment['attachment']);
                        fclose($fp);
                        $to_be_converted=array(
                            'alla-ultra@mail.ru',
                            'elena_dik@inbox.ru',
                            'kivanokg@gmail.com',
                            'b2b@intermedia.kg',
                            '441111@intermedia.kg'
                        );
                    }
                }
                if($savename && in_array($fromaddr,$to_be_converted) && $extension)
                {

                    $inputfile = NULL;
                    $outputfile = $dir.$rand."--".$fromaddr."--f.".$alt_ext;
                    $options = array(
                        "apikey" => "H2E2CI4hEuUzukrI7wKu6zhzVMdp1I78btb86z9zNqrbDSGfacLeCAbjiwdBSIax6TalvxpZnIuF2X9ln6zzNg",
                        "input" => "download",
                        "wait"=>true,
                        "download" => true,
                        "inputformat" => $extension,
                        "outputformat" => $alt_ext,
                        "file"=>"http://api.temirbek.com/excel/".$date."/".$savename
                    );
                    convert($options, $inputfile, $outputfile);
                    @unlink($dir.$savename);
                }
            }
            if($count++ >= $max_emails) break;
        }
    }
    /* close the connection */
    imap_close($inbox);
    echo "Done";
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
                parsexcel($file, $sendermail);
                //parsexcelsimple($file,$sendermail);
                $sendermail='';
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
            //install rar from here:http://php.net/manual/en/rar.installation.php
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

function parsexcel2($file, $sendermail)
{

    $dbh=$GLOBALS['dbh'];
    $date=date('Y-m-d');
    if (!file_exists($file)) {
        exit('No file yoba');
    }

    /* Нужно перевести сомы в доллары для тех у кого цены в сомах */
    if($sendermail=='alex@meloman.kg')
    {
        $string = file_get_contents("http://kivano.kg/product/exrateapi");
        $exrate_arr=json_decode($string, true);
        $exrate=$exrate_arr['usd'];
    }
    else $exrate=false;

    //$notindbs=array();

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

    //for editing we need another object with formats //except from sigmaplus.kg
    if($sendermail!='sigma@sigmaplus.kg')
    {
        $objReader->setReadDataOnly(false);
        $objPHPExcelReport=$objReader->load($file);
    }

    $objWorksheet = $objPHPExcel->getActiveSheet();

    $highestRow = $objWorksheet->getHighestRow(); // e.g. 10
    $highestColumn = $objWorksheet->getHighestColumn(); // e.g 'F'
    //$columnIndex=PHPExcel_Cell::stringFromColumnIndex($highestColumn);
    $highestColumnIndex = PHPExcel_Cell::columnIndexFromString($highestColumn); // e.g. 5

    $sku_products = $dbh->query("SELECT product_id, title FROM sku WHERE sender='{$sendermail}'")->fetchAll(PDO::FETCH_ASSOC);

    $product_rows=$dbh->query("SELECT product_id, price, changed FROM product")->fetchAll(PDO::FETCH_ASSOC);
    $product_changed=array();
    $product_price=array();
    foreach($product_rows as $product_row)
    {
        $product_changed[$product_row['product_id']]=$product_row['changed'];
        $product_price[$product_row['product_id']]=$product_row['price'];
    }
    $datesec=strtotime($date);

    for ($row = 1; $row <= $highestRow; ++$row) {
        $title=''; $price='';
        for ($col = 0; $col <= $highestColumnIndex; ++$col) {
            $curval=$objWorksheet->getCellByColumnAndRow($col, $row)->getCalculatedValue();

            if($sendermail=='alex@meloman.kg') //Для этого поставщика отдельное условие ($tcolumn=2; $prcolumn=4;)
            {
                if($col==2 && $curval) $title=$curval;
                elseif($col==4 && $curval) $price=$curval;
            }
            else if($sendermail=='elena25@ultra.kg') //Для этого поставщика отдельное условие ($tcolumn=0; $prcolumn=1;)
            {
                if($col==0 && $curval) $title=$curval;
                elseif($col==1 && $curval) $price=$curval;
            }
            else
            {
                if(!isset($tcolumn) && (in_array($curval,$match_title) || strpos($curval, "Товар/Склад")!== false))
                {$tcolumn=$col;}
                if(!isset($prcolumn) && (in_array($curval,$match_price)))
                {$prcolumn=$col;}
                if (isset($tcolumn) && isset($prcolumn)) {
                    if($col==$tcolumn && $curval) $title=$curval;
                    elseif($col==$prcolumn && $curval) $price=$curval;
                }
                elseif($sendermail=='elena_dik@inbox.ru') //laptop prices file from this user doesn't have column headers
                {
                    if($col==0 && $curval)$title=$curval;
                    elseif($col==1 && $curval) $price=$curval;
                }
            }
        }

        if($title && $price)
        {
            //echo 'title: '.$title." price:".$price."</br>";
            if($exrate) $price=$price/$exrate;
            $title2=strtolower(preg_replace("/\s/", "", $title));
            $found_in_db=false;
            $has_id=false;
            if($sku_products)
            {
                foreach($sku_products as $sku_product)
                {
                    $pid=$sku_product['product_id'];
                    $dbtitle=strtolower(preg_replace("/\s/", "", $sku_product['title']));
                    if($title2==$dbtitle)
                    {
                        if($pid)
                        {
                            if(isset($product_changed[$pid]))
                            {
                                $timediff=$datesec-strtotime($product_changed[$pid]);
                                $days=$timediff/(60*60*24);
                            }
                            if(isset($days) && $days<=7) //если цена была импортирована в течение последних 7и дней
                            {
                                if($product_price[$pid]>=$price) //то меняем если предыдущая цена была выше этой (или равна этой чтобы changed оставался актуальным для "наличие")
                                    $dbh->exec("UPDATE product SET price='{$price}', changed='{$date}', sender='{$sendermail}', note='1' WHERE product_id='{$pid}'");
                            }
                            else
                            {
                                $dbh->exec("UPDATE product SET price='{$price}', changed='{$date}', sender='{$sendermail}', note='2' WHERE product_id='{$pid}'");
                            }
                            $has_id=true;
                        }

                        $found_in_db=true;
                    }
                    unset($days);
                }
                if(!$has_id)
                {
                    if($sendermail=='sigma@sigmaplus.kg')
                    {
                        $objPHPExcel->getActiveSheet()->getStyle('A'.$row.':F'.$row)->getFill()
                            ->setFillType(PHPExcel_Style_Fill::FILL_SOLID)
                            ->getStartColor()->setARGB('FFFF0000');
                    }
                    elseif($sendermail=='b2b@intermedia.kg')
                    {
                        $objPHPExcelReport->getActiveSheet()->getStyle($highestColumn.$row)->getFill()
                            ->setFillType(PHPExcel_Style_Fill::FILL_SOLID)
                            ->getStartColor()->setARGB('FFFF0000');
                    }
                    else
                    {
                        $objPHPExcelReport->getActiveSheet()->getStyle('A'.$row.':F'.$row)->getFill()
                            ->setFillType(PHPExcel_Style_Fill::FILL_SOLID)
                            ->getStartColor()->setARGB('FFFF0000');
                    }
                }
                if(!$found_in_db)
                {
                    $stmt = $dbh->prepare("INSERT INTO sku (title, sender) VALUES (:title, :sender)");
                    $stmt->bindParam(':title', $title);
                    $stmt->bindParam(':sender', $sendermail);
                    $stmt->execute();
                }
            }
            else //new supplier
            {
                $sth = $dbh->prepare("SELECT product_id FROM sku WHERE product_id<>'0' AND title=:title");
                $sth->bindParam(':title', $title, PDO::PARAM_STR);
                $sth->execute();
                $row=$sth->fetch(PDO::FETCH_ASSOC);
                if($row['product_id']) $product_id=$row['product_id']; else $product_id=0;

                $stmt = $dbh->prepare("INSERT INTO sku (title, sender, product_id) VALUES (:title, :sender, :pid)");
                $stmt->bindParam(':title', $title, PDO::PARAM_STR);
                $stmt->bindParam(':sender', $sendermail, PDO::PARAM_STR);
                $stmt->bindParam(':pid', $product_id, PDO::PARAM_INT);
                $stmt->execute();
            }
        }
    }
    $rand=rand(1,100);
    if($sendermail!='sigma@sigmaplus.kg')
    {
        $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcelReport, 'Excel2007');
    }
    else $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
    $objWriter->save(dirname(__FILE__)."/report/".$date."/".$rand.'-'.$sendermail.".xlsx");
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

function importProducts()
{
    $dbh=$GLOBALS['dbh'];
    $string = file_get_contents("http://kivano.kg/product/productapi"); //parses products that have been changed or added today
    //$string = file_get_contents("http://kivano.com.kg/product/productAPI");
    $json_a = json_decode($string, true);

    if ($json_a) {
        $ids = array();
        $db_rows=$dbh->query("SELECT id, title FROM sku")->fetchAll(PDO::FETCH_ASSOC);

        foreach ($json_a as $key => $value)
        {
            $id = $value['id'];
            $category = $value['category_id'];
            $commonsku=$value['commonsku'];
            $manual=$value['manual'];
            $result=$dbh->query("SELECT id FROM product WHERE product_id='{$id}'")->fetch();
            if($result)
            {
                $stmt = $dbh->prepare("UPDATE product SET sku=:sku, manual='{$manual}' WHERE product_id='{$id}'");
                $stmt->bindParam(':sku', $commonsku);
                $stmt->execute();
            }
            else
            {
                $stmt = $dbh->prepare("INSERT INTO product (product_id, sku, category_id, manual) VALUES('{$id}', :sku, '{$category}', '{$manual}')");
                $stmt->bindParam(':sku', $commonsku);
                $stmt->execute();
            }
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

function deleteDeleted()
{
    $dbh=$GLOBALS['dbh'];
    $string = file_get_contents("http://kivano.kg/product/deletedapi");
    //$string = file_get_contents("http://kivano.com.kg/product/productAPI");
    $json_a = json_decode($string, true);

    if ($json_a) {
        foreach ($json_a as $key => $value)
        {
            $id = $value['product_id'];
            $dbh->exec("DELETE FROM product WHERE product_id='{$id}'");
            $dbh->exec("DELETE FROM sku WHERE product_id='{$id}'");
        }
    }
}

function priceAPI($date,$change,$list_all)
{
    $dbh=$GLOBALS['dbh'];
    //$forum=(int)$_GET['forum'];
    if($list_all) //show all of rows changed this date
    {
        $db_rows=$dbh->query("SELECT id,product_id, category_id, price, sender, exrate FROM product WHERE changed='{$date}'")->fetchAll(PDO::FETCH_ASSOC);
    }
    else{
        $db_rows=$dbh->query("SELECT id,product_id, category_id, price, sender, exrate FROM product WHERE changed='{$date}' AND exported<>'{$date}'")->fetchAll(PDO::FETCH_ASSOC);
    }

    //$rows=Yii::app()->db->createCommand("SELECT id, commonsku, category_id FROM Product WHERE id=108")->queryAll();
    $products=array();
    foreach($db_rows as $prod)
    {
        $products[]=array('product_id'=>$prod['product_id'], 'category_id'=>$prod['category_id'], 'price'=>$prod['price'], 'sender'=>$prod['sender'], 'exrate'=>$prod['exrate']);
        $rowid=$prod['id'];
        if($change) //change exported field to today's date (which will mean it was exported today and no need to export again on next call)
        {
            $dbh->exec("UPDATE product SET exported='{$date}' WHERE id='{$rowid}'");
        }
    }

    $json=json_encode($products);

    header('Content-type: application/json');
    echo $json;
}

function showReportDirs()
{
    $files = scandir(dirname(__FILE__)."/report/", 1);
    if($files)
    {
        foreach($files as $file)
        {
            if(is_dir($dir=dirname(__FILE__)."/report/".$file) && $file!='.' && $file!='..')
            {
                echo "<a href='jugur.php?param=reportfiles&name=".$file."'>".$file."</a><br />";
            }
        }
    }

}

function showReportFiles($date)
{
    $files = scandir(dirname(__FILE__)."/report/".$date, 1);
    if($files)
    {
        foreach($files as $file)
        {
            if(is_file($dir=dirname(__FILE__)."/report/".$date."/".$file))
            {
                echo "<a href='/report/".$date."/".$file."'>".$file."</a><br />";
            }
        }
    }
}

/* when new user is detected in ParseExcel function it's titles (skus) are inserted to `sku` table with product_id=0.
then below function checks if these skus already exists, if yes then copies their product_id
 */
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
        else //look for product table too
        {
            $title="%".$title."%";
            $sth = $dbh->prepare("SELECT product_id FROM product WHERE sku LIKE :title");
            $sth->bindParam(':title', $title, PDO::PARAM_STR);
            $sth->execute();
            $row=$sth->fetch(PDO::FETCH_ASSOC);

            if($pid=$row['product_id'])
            {
                $dbh->exec("UPDATE sku SET product_id='{$pid}' WHERE id='{$id}'");
            }
        }
    }
}

function listUsers()
{

    $string = file_get_contents("http://kivano.kg/user/excel?key=90364c8c198986e5555279dac17939d0");
    $json_a = json_decode($string, true);
    $date=date("d/m/Y");
    $s=3; $u=3;

    if ($json_a) {
        require_once dirname(__FILE__) . '/Classes/PHPExcel/IOFactory.php';
        $objPHPExcel = new PHPExcel();
        $objPHPExcel->setActiveSheetIndex(0);
        $objPHPExcel->getActiveSheet()->setCellValue('A1', 'Дата: '.$date);
        $objPHPExcel->getActiveSheet()->getStyle('A1')->getFont()->setBold(true);
        $objPHPExcel->getActiveSheet()->setCellValue('A2', 'Подписанные');
        $objPHPExcel->getActiveSheet()->setCellValue('B2', 'Не подписанные');
        $objPHPExcel->getActiveSheet()->getStyle('A2')->getFont()->setBold(true);
        $objPHPExcel->getActiveSheet()->getStyle('B2')->getFont()->setBold(true);
        $objPHPExcel->getActiveSheet()->getStyle('A2:B2')->getFill()
            ->setFillType(PHPExcel_Style_Fill::FILL_SOLID)
            ->getStartColor()->setARGB('579FFF');
        foreach ($json_a as $key => $value)
        {
            if(isset($value['subscribed'])) //comes from User table
            {
                if($value['subscribed']){
                    $objPHPExcel->getActiveSheet()->setCellValue('A'.$s, $value['email']);
                    $s++;
                }
                else{
                    $objPHPExcel->getActiveSheet()->setCellValue('B'.$u, $value['email']);
                    $u++;
                }
            }
            else //comes from `subscribed` table
            {
                $objPHPExcel->getActiveSheet()->setCellValue('A'.$s, $value['email']);
                $s++;
            }
            $objPHPExcel->getActiveSheet()->getColumnDimension('A')->setAutoSize(true);
            $objPHPExcel->getActiveSheet()->getColumnDimension('B')->setAutoSize(true);
        }
        //$objPHPExcel->getActiveSheet()->fromArray($arrayData);
        $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
        $objWriter->save("users.xlsx");
        $url='http://api.temirbek.com/users.xlsx';
        header("Location: $url");
    }
}

function parsexcel($file, $sendermail)
{

    $dbh=$GLOBALS['dbh'];
    $date=date('Y-m-d');
    if (!file_exists($file)) {
        exit('No file yoba');
    }

    /* Нужно перевести сомы в доллары для тех у кого цены в сомах */
    if($sendermail=='alex@meloman.kg')
    {
        $string = file_get_contents("http://kivano.kg/product/exrateapi");
        $exrate_arr=json_decode($string, true);
        $exrate=$exrate_arr['usd'];
    }
    else $exrate=false;

    //$notindbs=array();

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

    //for editing we need another object with formats //except from sigmaplus.kg
    if($sendermail!='sigma@sigmaplus.kg')
    {
        $objReader->setReadDataOnly(false);
        $objPHPExcelReport=$objReader->load($file);
    }

    $objWorksheet = $objPHPExcel->getActiveSheet();

    $highestRow = $objWorksheet->getHighestRow(); // e.g. 10
    $hRow=0; //real highest row that has real value in it, $highestRow counts all modified cells even if they are empty
    $highestColumn = $objWorksheet->getHighestColumn(); // e.g 'F'
    //$columnIndex=PHPExcel_Cell::stringFromColumnIndex($highestColumn);
    $highestColumnIndex = PHPExcel_Cell::columnIndexFromString($highestColumn); // e.g. 5
    $hCol=0; //real highest index that is actually has value

    $sku_products = $dbh->query("SELECT product_id, title FROM sku WHERE sender='{$sendermail}'")->fetchAll(PDO::FETCH_ASSOC);

    $product_rows=$dbh->query("SELECT product_id, price, changed, manual FROM product")->fetchAll(PDO::FETCH_ASSOC);
    $product_changed=array();
    $product_price=array();
    $product_manual=array();
    foreach($product_rows as $product_row)
    {
        $product_changed[$product_row['product_id']]=$product_row['changed'];
        $product_price[$product_row['product_id']]=$product_row['price'];
        $product_manual[$product_row['product_id']]=$product_row['manual'];
    }
    $datesec=strtotime($date);
    $processed=array();
    $repetition=array();
    $manual=array();
    $expensive=array();
    $total=0;
    $notfound=0;
    $title_column_index=0;
    for ($row = 1; $row <= $highestRow; ++$row) {
        $title=''; $price=''; $note=''; $rowFilled=false;
        for ($col = 0; $col <= $highestColumnIndex; ++$col) {
            $curval=$objWorksheet->getCellByColumnAndRow($col, $row)->getCalculatedValue();
            if($curval) {$rowFilled=true; if($hCol<$col) $hCol=$col;}

            if($sendermail=='alex@meloman.kg') //Для этого поставщика отдельное условие ($tcolumn=2; $prcolumn=4;)
            {
                if($col==2 && $curval) {$title=$curval; $title_column_index=2;}
                elseif($col==4 && $curval) $price=$curval;
            }
            else if($sendermail=='elena25@ultra.kg') //Для этого поставщика отдельное условие ($tcolumn=0; $prcolumn=1;)
            {
                if($col==0 && $curval) {$title=$curval;  $title_column_index=0;}
                elseif($col==1 && $curval) $price=$curval;
            }
            else
            {
                if(!isset($tcolumn) && (in_array($curval,$match_title) || strpos($curval, "Товар/Склад")!== false))
                {$tcolumn=$col;  $title_column_index=$col;}
                if(!isset($prcolumn) && (in_array($curval,$match_price)))
                {$prcolumn=$col;}
                if (isset($tcolumn) && isset($prcolumn)) {
                    if($col==$tcolumn && $curval) $title=$curval;
                    elseif($col==$prcolumn && $curval) $price=$curval;
                }
                elseif($sendermail=='elena_dik@inbox.ru') //laptop prices file from this user doesn't have column headers
                {
                    if($col==0 && $curval){$title=$curval; $title_column_index=0;}
                    elseif($col==1 && $curval) $price=$curval;
                }
            }
        }
        if($rowFilled)  $hRow++;

        if($title && $price)
        {
            $total++;
            //echo 'title: '.$title." price:".$price."</br>";
            if($exrate) $price=$price/$exrate;
            $title2=strtolower(preg_replace("/\s/", "", $title));
            $found_in_db=false;
            $has_id=false;
            if($sku_products)
            {
                foreach($sku_products as $sku_product)
                {
                    $pid=$sku_product['product_id'];
                    $dbtitle=strtolower(preg_replace("/\s/", "", $sku_product['title']));
                    if($title2==$dbtitle)
                    {
                        if($pid)
                        {
                            if(isset($product_changed[$pid]))
                            {
                                $timediff=$datesec-strtotime($product_changed[$pid]);
                                $days=$timediff/(60*60*24);
                            }
                            if(isset($days) && $days<=7) //если цена была импортирована в течение последних 7и дней
                            {
                                if($product_price[$pid]>=$price) //то меняем если предыдущая цена была выше этой (или равна этой чтобы changed оставался актуальным для "наличие")
                                    $dbh->exec("UPDATE product SET price='{$price}', changed='{$date}', sender='{$sendermail}', note='1' WHERE product_id='{$pid}'");
                                else{
                                    $expensive[]=$pid;
                                    $note.=' дороже';
                                    $objPHPExcelReport->getActiveSheet()->getStyle('A'.$row.':F'.$row)->getFill()
                                        ->setFillType(PHPExcel_Style_Fill::FILL_SOLID)
                                        ->getStartColor()->setARGB('FFF4A460');
                                }
                            }
                            else
                            {
                                $dbh->exec("UPDATE product SET price='{$price}', changed='{$date}', sender='{$sendermail}', note='2' WHERE product_id='{$pid}'");
                            }
                            $has_id=true;
                            if($product_manual[$pid]) {
                                $manual[]=$pid;
                                $note.=' ручной';
                                $objPHPExcelReport->getActiveSheet()->getStyle('A'.$row.':F'.$row)->getFill()
                                    ->setFillType(PHPExcel_Style_Fill::FILL_SOLID)
                                    ->getStartColor()->setARGB('FFDAA520');
                            }
                            if(in_array($pid,$processed)) {
                                $repetition[]=$pid;
                                $note.=' повтор';
                                $objPHPExcelReport->getActiveSheet()->getStyle('A'.$row.':F'.$row)->getFill()
                                    ->setFillType(PHPExcel_Style_Fill::FILL_SOLID)
                                    ->getStartColor()->setARGB('FFDDA0DD');
                            }
                            else {$processed[]=$pid;}
                            $thecol=$hCol+1;
                            $objPHPExcelReport->getActiveSheet()->setCellValueByColumnAndRow($thecol, $row, $pid.' '.$note);
                        }
                        $found_in_db=true;
                    }
                    unset($days);
                }
                if(!$found_in_db)
                {
                    $stmt = $dbh->prepare("INSERT INTO sku (title, sender) VALUES (:title, :sender)");
                    $stmt->bindParam(':title', $title);
                    $stmt->bindParam(':sender', $sendermail);
                    $stmt->execute();
                }
            }
            else //new supplier
            {
                $sth = $dbh->prepare("SELECT product_id FROM sku WHERE product_id<>'0' AND title=:title");
                $sth->bindParam(':title', $title, PDO::PARAM_STR);
                $sth->execute();
                $row=$sth->fetch(PDO::FETCH_ASSOC);
                if($row['product_id']) $product_id=$row['product_id']; else $product_id=0;

                if($product_id)
                {
                    if(isset($product_changed[$product_id]))
                    {
                        $timediff=$datesec-strtotime($product_changed[$product_id]);
                        $days=$timediff/(60*60*24);
                    }
                    if(isset($days) && $days<=7) //если цена была импортирована в течение последних 7и дней
                    {
                        if($product_price[$product_id]>=$price) //то меняем если предыдущая цена была выше этой (или равна этой чтобы changed оставался актуальным для "наличие")
                            $dbh->exec("UPDATE product SET price='{$price}', changed='{$date}', sender='{$sendermail}', note='3' WHERE product_id='{$product_id}'");
                        else{
                            $expensive[]=$product_id;
                            $note.=' дороже';
                            $objPHPExcelReport->getActiveSheet()->getStyle('A'.$row.':F'.$row)->getFill()
                                ->setFillType(PHPExcel_Style_Fill::FILL_SOLID)
                                ->getStartColor()->setARGB('FFF4A460');
                        }
                    }
                    else
                    {
                        $dbh->exec("UPDATE product SET price='{$price}', changed='{$date}', sender='{$sendermail}', note='4' WHERE product_id='{$product_id}'");
                    }
                    $has_id=true;
                    if(in_array($product_id,$processed)) {
                        $repetition[]=$product_id;
                        $note.=' повтор';
                        $objPHPExcelReport->getActiveSheet()->getStyle('A'.$row.':F'.$row)->getFill()
                            ->setFillType(PHPExcel_Style_Fill::FILL_SOLID)
                            ->getStartColor()->setARGB('FFDDA0DD');
                    }
                    else {$processed[]=$product_id;}
                    $thecol=$hCol+1;
                    $objPHPExcelReport->getActiveSheet()->setCellValueByColumnAndRow($thecol, $row, $product_id.' '.$note);

                }

                $stmt = $dbh->prepare("INSERT INTO sku (title, sender, product_id) VALUES (:title, :sender, :pid)");
                $stmt->bindParam(':title', $title, PDO::PARAM_STR);
                $stmt->bindParam(':sender', $sendermail, PDO::PARAM_STR);
                $stmt->bindParam(':pid', $product_id, PDO::PARAM_INT);
                $stmt->execute();
            }

            if(!$has_id)
            {
                $notfound++;
                if($sendermail=='sigma@sigmaplus.kg')
                {
                    $objPHPExcel->getActiveSheet()->getStyle('A'.$row.':F'.$row)->getFill()
                        ->setFillType(PHPExcel_Style_Fill::FILL_SOLID)
                        ->getStartColor()->setARGB('FFFF0000');
                }
                elseif($sendermail=='b2b@intermedia.kg')
                {
                    $objPHPExcelReport->getActiveSheet()->getStyle($highestColumn.$row)->getFill()
                        ->setFillType(PHPExcel_Style_Fill::FILL_SOLID)
                        ->getStartColor()->setARGB('FFFF0000');
                }
                else
                {
                    $objPHPExcelReport->getActiveSheet()->getStyle('A'.$row.':F'.$row)->getFill()
                        ->setFillType(PHPExcel_Style_Fill::FILL_SOLID)
                        ->getStartColor()->setARGB('FFFF0000');
                }
            }
        }
    }
    //$pids=implode(',',$processed);
    //$objPHPExcelReport->getActiveSheet()->setCellValueByColumnAndRow(7, 1, $pids);

    /*-------------report-------------*/
    //total
    $objPHPExcelReport->getActiveSheet()->setCellValueByColumnAndRow($title_column_index, $hRow+1, 'Всего: '.$total);
    //saved
    $expensives=count($expensive);
    $saved=count($processed)-$expensives;
    $objPHPExcelReport->getActiveSheet()->setCellValueByColumnAndRow($title_column_index, $hRow+2, 'Изменено: '.$saved);
    //not found
    $objPHPExcelReport->getActiveSheet()->setCellValueByColumnAndRow($title_column_index, $hRow+3, 'Не найдено: '.$notfound);
    //expensive
    if($expensives){
        $expen_report_row=$hRow+4;
        $objPHPExcelReport->getActiveSheet()->getStyle('A'.$expen_report_row.':B'.$expen_report_row)->getFill()
            ->setFillType(PHPExcel_Style_Fill::FILL_SOLID)
            ->getStartColor()->setARGB('FFF4A460');
        $objPHPExcelReport->getActiveSheet()->setCellValueByColumnAndRow($title_column_index, $expen_report_row, 'Дороже: '.$expensives);

    }
    //repetition
    if($reps=count($repetition)){

        $rep_report_row=$hRow+5;
        $objPHPExcelReport->getActiveSheet()->getStyle('A'.$rep_report_row.':B'.$rep_report_row)->getFill()
            ->setFillType(PHPExcel_Style_Fill::FILL_SOLID)
            ->getStartColor()->setARGB('FFDDA0DD');
        $objPHPExcelReport->getActiveSheet()->setCellValueByColumnAndRow($title_column_index, $rep_report_row, 'Повторяются: '.$reps);
    }
    //manual
    if($manuals=count($manual))
    {
        $man_report_row=$hRow+6;
        $objPHPExcelReport->getActiveSheet()->getStyle('A'.$man_report_row.':B'.$man_report_row)->getFill()
            ->setFillType(PHPExcel_Style_Fill::FILL_SOLID)
            ->getStartColor()->setARGB('FFDDA0DD');
        $objPHPExcelReport->getActiveSheet()->setCellValueByColumnAndRow($title_column_index, $man_report_row, 'Ручной: '.$manuals);
    }


    $rand=rand(1,100);
    if($sendermail!='sigma@sigmaplus.kg')
    {
        $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcelReport, 'Excel2007');
    }
    else $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
    $objWriter->save(dirname(__FILE__)."/report/".$date."/".$rand.'-'.$sendermail.".xlsx");
}

function run2(){
    $dbh=$GLOBALS['dbh'];
   /* $product_id=34;
    $dbh->exec("INSERT INTO temir (text) VALUES ('hehe {$product_id}')");
    die();*/
    /*$pid=234;
    $dbh->exec("INSERT INTO temir (text) VALUES ('f {$pid}')");*/
    /*$stmt = $dbh->query("SELECT product_id FROM product WHERE sender='sales@intercom.kg' AND changed='2015-12-09'")->fetchAll(PDO::FETCH_ASSOC);
    $i=1;
    $ids="17012,19158,20399,17019,8507,17023,16898,8509,20425,9232,19780,19778,20389,19166,17788,19167,19890,19168,19170,19976,19978,19979,19891,4916,19892,19181,19184,7955,9048,2949,4865,8525,8533,20006,20007,20008,7939,4828,19186,19189,20009,4184,4185,17035,20012,20013,20415,20416,9202,9203,9204,17039,17044,17045,20417,20418,20419,10076,10077,20420,20020,9170,20421,20422,20423,20424,20023,20024,19893,17896,19210,20025,9210,9211,20026,20027,20028,20029,17055,14431,17057,19212,19213,19242,20032,20033,20034,20035,20036,20037,17908,17909,17911,19243,17907,19244,19245,9228,7943,7941,7942,8183,8184,8182,8181,9393,9967,17115,10299,10016,19068,9971,20348,9408,19841,2798,19259,11745,17799,19133,19072,20428,19075,20429,20430,20431,19076,20432,20433,20434,20435,19077,19288,20436,19342,19343,17036,20437,9367,9368,17043";
    $products=array();
    foreach($stmt as $row){
        $ids=str_replace($row['product_id'],'',$ids);
    }
    echo $ids;*/
    /*echo "<pre>";
    print_r($stmt);
    echo "</pre>";*/
    /*$stmt = $dbh->query("SELECT DISTINCT title FROM sku WHERE product_id<>0")->fetchAll(PDO::FETCH_ASSOC);
    echo count($stmt);
echo '<br />';*/

    $sku_products = $dbh->query("SELECT product_id, title FROM sku WHERE product_id<>0 GROUP BY title LIMIT 5")->fetchAll(PDO::FETCH_ASSOC);
    foreach($sku_products as $row){
        echo $row['title'].'<br />';
    }

   /* $stmt = $dbh->query("SELECT id, title FROM sku WHERE product_id=0")->fetchAll(PDO::FETCH_ASSOC);
    foreach($stmt as $row){
        $stm = $dbh->prepare("DELETE FROM sku WHERE title=:title AND id<>'{$row['id']}'");
        $stm->bindParam(':title', $row['title']);
        $stm->execute();
    }*/
}
?>