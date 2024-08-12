<?php

require './app/controller.php';
date_default_timezone_set("Asia/Bangkok");

$generate = new xcel;
$item = isset($_GET['item']) ? $_GET['item'] : " ";
$from = isset($_GET['from']) ? round($_GET['from']/1000,0) : time() - 86400;
$to = isset($_GET['to']) ? round($_GET['to']/1000,0) : time();
$parameter = isset($_GET['parameter']) ? $_GET['parameter'] : " ";

switch ($parameter) {
    case "power_meter":
        $generate->power_meter($item, $from, $to);
        break;
        
    case "temperature":
        $generate->temperature($from, $to);
        break;
        
    case "flow_meter":
        $generate->flow_meter($item, $from, $to);
        break;
        
    case "dry_contact":
        $generate->dry_contact($from, $to);
        break;
        
    default:
        die("Unrecognized Parameter");
}