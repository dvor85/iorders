<?php
ini_set('display_errors',1);
require_once("../config/cf.php");	
require_once("../config/classes.php");	
require_once("../config/functions.php");
	
$my=new datamysql(MYSQL_HOST,MYSQL_BASE,MYSQL_USER,MYSQL_PASS);
$my->set_charset("cp1251");

define("DELIVERY_INFO_ORDER","delivery_info_order");
define("DELIVERY_ORDERS","delivery_orders");
define("DELIVERY_TIME",3600);


function get_xmlorder($id_order)
{
	global $my;
	$sql1 = sprintf("SELECT * FROM %s where id_order=$id_order",DELIVERY_INFO_ORDER);			
	$order=&$my->query($sql1);
	if ($order) {
		$sql2 = sprintf("SELECT * FROM %s where id_order=$id_order",DELIVERY_ORDERS);			
		$menu=&$my->query($sql2);
		
		$xml=new XMLWriter();
		$xml->openMemory();
		$xml->setIndent(true);
		if ($xml) {			
			$xml->startDocument("1.0","UTF-8");
				$xml->startElement("order");
					$xml->startElement("f_name");
						$xml->text(iconv("windows-1251","utf-8",$order[0]["f_name"]));
					$xml->endElement();
					$xml->startElement("l_name");
						$xml->text(iconv("windows-1251","utf-8",$order[0]["l_name"]));
					$xml->endElement();
					$xml->startElement("m_name");
						$xml->text(iconv("windows-1251","utf-8",$order[0]["m_name"]));
					$xml->endElement();
					$xml->startElement("organization");
						$xml->text(iconv("windows-1251","utf-8",$order[0]["organization"]));
					$xml->endElement();
					$xml->startElement("phone1");
						$xml->text(iconv("windows-1251","utf-8",$order[0]["phone1"]));
					$xml->endElement();
					$xml->startElement("phone2");
						$xml->text(iconv("windows-1251","utf-8",$order[0]["phone2"]));
					$xml->endElement();
					$xml->startElement("town");
						$xml->text(iconv("windows-1251","utf-8",$order[0]["town"]));
					$xml->endElement();
					$xml->startElement("street");
						$xml->text(iconv("windows-1251","utf-8",$order[0]["street"]));
					$xml->endElement();
					$xml->startElement("house");
						$xml->text(iconv("windows-1251","utf-8",$order[0]["house"]));
					$xml->endElement();
					$xml->startElement("building");
						$xml->text(iconv("windows-1251","utf-8",$order[0]["building"]));
					$xml->endElement();
					$xml->startElement("entry");
						$xml->text(iconv("windows-1251","utf-8",$order[0]["entry"]));
					$xml->endElement();
					$xml->startElement("flat");
						$xml->text(iconv("windows-1251","utf-8",$order[0]["flat"]));
					$xml->endElement();
					$xml->startElement("floor");
						$xml->text(iconv("windows-1251","utf-8",$order[0]["floor"]));
					$xml->endElement();
					$xml->startElement("codeentry");
						$xml->text(iconv("windows-1251","utf-8",$order[0]["codeentry"]));
					$xml->endElement();
					$xml->startElement("email");
						$xml->text(iconv("windows-1251","utf-8",$order[0]["email"]));
					$xml->endElement();
					$xml->startElement("adv_info");
						$xml->text(iconv("windows-1251","utf-8",$order[0]["adv_info"]));
					$xml->endElement();
					$xml->startElement("advert");
						$xml->text(iconv("windows-1251","utf-8",$order[0]["advert"]));
					$xml->endElement();
					$xml->startElement("order_summ");
						$xml->text((float)$order[0]["order_summ"]);
					$xml->endElement();
					$xml->startElement("d_type");
						$xml->text($order[0]["d_type"]);
					$xml->endElement();
					$xml->startElement("wait_time");
						$xml->text(date("d.m.Y H:i:s",strtotime($order[0]["wait_time"])+DELIVERY_TIME));
					$xml->endElement();
						$xml->startElement("menu");
						for ($i=0;$i<count($menu);$i++)
						{
							$num=$i+1;
							$xml->startElement("itemnum".$num);
								$xml->startElement("item");
									$xml->writeAttribute("id",$menu[$i]["item_id"]);
									$xml->writeAttribute("code",$menu[$i]["item_code"]);
									$xml->writeAttribute("name",iconv("windows-1251","utf-8",$menu[$i]["item_name"]));
									$xml->writeAttribute("quantity",(float)$menu[$i]["quantity"]);
									$xml->writeAttribute("price",(float)$menu[$i]["price"]);
								$xml->endElement();	
							$xml->endElement();
						}
						$xml->endElement();
				$xml->endElement();	
			$xml->endDocument();
			return array("xml"=>$xml->flush(),
						 "datetime"=>strtotime($order[0]["wait_time"]));
		}		
	}
	return "";	
}

function getListOrders() {
	global $my;
	$s="";
	//$sql = sprintf("SELECT * FROM %s where status=0 and DATE(wait_time)=DATE()",DELIVERY_INFO_ORDER);
	$sql = sprintf("SELECT id_order FROM %s where status=0",DELIVERY_INFO_ORDER);
	$orders=&$my->query($sql);	
	if ($orders) {		
		for ($i=0;$i<count($orders);$i++) {
			$s.=$orders[$i]["id_order"]."\r\n";		
		}
	}
	return $s;
}

if (isset($_REQUEST["getlistorders"])) {
	header("Content-type:text/plain; charset=windows-1251");
	header("Last-Modified: ".gmdate("D, d M Y H:i:s")." GMT");
	header("Pragma: no-cache");
	echo getListOrders();
}

if (isset($_REQUEST["getxmlorder"])&&!empty($_REQUEST["id_order"])) {	
	$id_order=(int)substr(strip_tags(stripslashes(trim($_REQUEST["id_order"]))),0,16);
	$res=get_xmlorder($id_order);	
	if (is_array($res)) {
		$file_name="d_".date("dmYHisB",$res["datetime"]).".xml";
		header("Content-type:text/xml; charset=utf-8");				
		header("Content-Disposition: attachment; filename=\"$file_name\"");
		header('Content-Transfer-Encoding: binary');
		header("Last-Modified: ".gmdate("D, d M Y H:i:s")." GMT");
		header("Pragma: no-cache");
		echo $res["xml"];	
	} else {
		header("Content-type:text/plain; charset=utf-8");
	}
	
}
	


?>