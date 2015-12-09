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


//connection to the database
$dbhandle = mysql_connect($hostname, $username, $password)
    or die("Unable to connect to MySQL");

$selected = mysql_select_db("jugur_kivano",$dbhandle)
    or die("Could not select examples");

mysql_set_charset("utf8");
/*$match_title=array("Товар","Модель","model","Наименование товаров","Notebooks","НАИМЕНОВАНИЕ","Наименование","Наименование товара");
$match_price=array("ДЛР (usd)","Цена реал (USD  )","price (USD  )","Цена","Мелкооптовая цена","Dealer ", "ЦЕНА","Дилер", "Price", "Дилерская","Длр.");
for($i=0;$i<10;$i++){
    $mt=$match_title[$i];
    $mp=$match_price[$i];
    mysql_query("INSERT INTO thead(title, price) VALUES('{$mt}','{$mp}')");
}*/

if(isset($_POST['field'])){
    $field=$_POST['field'];
    $id=$_POST['id'];
    if(mysql_query("UPDATE thead SET `{$field}`='' WHERE id='{$id}'"))
        echo json_encode(array('msg'=>'success'));
    else
        echo json_encode(array('msg'=>'error'));
    die();
}

if(isset($_GET['title']))
{
    if(isset($_GET['title']) && $_GET['title']) $title=$_GET['title']; else $title='';
    if(isset($_GET['price']) && $_GET['price']) $price=$_GET['price']; else $price='';
    if(isset($_GET['control']) && (int)$_GET['control']==5)
    {
        mysql_query("INSERT INTO thead (title, price) VALUES('{$title}','{$price}')");
    }
}
$rows=mysql_query("SELECT id, title, price, note FROM thead");
$tres=array();
$tpri=array();
$fetched=array();
while($row = mysql_fetch_array($rows)) {
    // Print out the contents of each row into a table
$fetched[]=array('price'=>$row['price'],'title'=>$row['title'],'id'=>$row['id']);
}
?>
<style type="text/css">
    table{border-collapse: collapse;}
    td{ border:1px solid #ccc; text-align: center;}
    .delete{text-decoration: underline;
        cursor: pointer;}
</style>
<div style="margin-left: 50px;">
    <form action="" method="GET">
        <label for='title'>Заголовок столбца наименований:</label>
        <br />
        <input type="text" name="title" id="title" />
        <br />
        <br />
        <label for='price'>Заголовок столбца цен:</label>
        <br />
        <input type="text" name="price" id="price" />
        <br />
        <br />
        <label for='control'>3 плюс 2 равно:</label>
        <br />
        <input type="text" name="control" id="control" />
 <!--       <br />
        <br />
        <label for='note'>Примечание:</label>
        <br />
        <textarea name="note" id="note" cols='30' rows="5"></textarea>-->
        <br />
        <input type="submit" value="Добавить" />
    </form>
</div>
<div style="float:left;width:300px;">
    <table>
        <tr>
            <td>Наименование</td>
            <td></td>
        </tr>
        <?php
        foreach($fetched as $tr)
        {
            if($tr['title']) echo "<tr><td style='white-space: nowrap; padding-right: 5px;'>".$tr['title']."</td><td><span class='js_delete delete' field='title' rid='".$tr['id']."'>удалить</span></td></tr>";
        }
        ?>
    </table>
</div>
<div style="width:40%;">
    <table>
        <tr>
            <td>Цена</td>
            <td></td>
        </tr>
        <?php
        foreach($fetched as $tr)
        {
            if($tr['price']) echo "<tr><td style='white-space: nowrap; padding-right: 5px;'>".$tr['price']."</td><td><span class='js_delete delete' field='price' rid='".$tr['id']."'>удалить</span></td></tr>";
        }
        ?>
    </table>
</div>

<?php
//saveexcel();
//parsexcelsimple($dir);
//mysql_query("INSERT INTO temir (text) VALUES('CURDATE()')");

//$rows=mysql_query("SELECT product_id, category_id, price, sender, exrate FROM product WHERE changed='CURDATE()'");

//close the connection
mysql_close($dbhandle);
?>
<script src="https://ajax.googleapis.com/ajax/libs/jquery/2.1.3/jquery.min.js"></script>
<script type="text/javascript">
    $(document).ready(function(){
        $('.js_delete').click(function(){
            if(confirm('Вы уверены что хотите удалить?')){
                var rid=$(this).attr('rid');
                var field=$(this).attr('field');
                var dis=$(this);
                $.ajax({
                    type: 'POST',
                    dataType: 'json',
                    data: {field:field,id:rid},
                    url: '',
                    success:function(msg){
                        if(msg.msg==='success'){
                            dis.parents('tr').remove();
                        }
                        else alert('Не удалось удалить');
                    }
                });
            }
        });
    });
</script>