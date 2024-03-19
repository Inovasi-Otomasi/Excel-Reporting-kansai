<?php
require './app/models.php';


class xcel {
    function __construct(){
        $this->models = new config;
    }
    
    public function power_meter($item = "", $from, $to){
        $item = $this->models->get_item($item)[0]->id;
        $sql = "SELECT * FROM power WHERE item_id = $item and ts BETWEEN FROM_UNIXTIME($from) AND FROM_UNIXTIME($to) LIMIT 10000";
        $data = $this->models->get_data($sql);
        $spreadsheet = $this->models->reader("./Report/pm.xlsx");
        $activeWorksheet = $spreadsheet->getSheetByName('PLN');

        $i = 3;

        foreach($data as $dt){
            $i1 = $dt->i1;
            $i2 = $dt->i2;
            $i3 = $dt->i3;
            $in = $dt->i_n;
    
            $va = $dt->v1;
            $vb = $dt->v2;
            $vc = $dt->v3;
            $vn = $dt->vn;
    
            $pf = $dt->pf;
            
            $watt   = $dt->watt;
            $va     = $dt->va;
            $var    = $dt->var;
    
            $kwh    = $dt->wh;
            $kvah   = $dt->vah;
            $kvarh  = $dt->varh;
    
            $activeWorksheet->setCellValue('B'.$i, '=(((('. $dt->ts + 25200 .'/60)/60)+8)/24)+DATE(1970,1,1)'); //time
    
            $activeWorksheet->setCellValue('C'.$i, $watt);
            $activeWorksheet->setCellValue('D'.$i, $var);
            $activeWorksheet->setCellValue('E'.$i, $va);
            
            $activeWorksheet->setCellValue('F'.$i, $i1);
            $activeWorksheet->setCellValue('G'.$i, $i2);
            $activeWorksheet->setCellValue('H'.$i, $i3);
            $activeWorksheet->setCellValue('I'.$i, $in);
            
            $activeWorksheet->setCellValue('J'.$i, $va);
            $activeWorksheet->setCellValue('K'.$i, $vb);
            $activeWorksheet->setCellValue('L'.$i, $vc);
            $activeWorksheet->setCellValue('M'.$i, $vn);
            
            $activeWorksheet->setCellValue('N'.$i, $kvarh);
            $activeWorksheet->setCellValue('O'.$i, $kvah);
            $activeWorksheet->setCellValue('P'.$i, $kwh);
    
            $activeWorksheet->setCellValue('Q'.$i, $pf);
    
            $i++;
        }
    
        $filename = $item .' Energy Report '. date("d M, Y (H-i)",$from) .' - '. date("d M, Y (H-i)", $to);
        $this->models->writer($spreadsheet, $filename);
    }

    public function flow_meter($item = "", $from, $to){
        $item = $this->models->get_item($item)[0]->id;
        $sql = "SELECT * FROM flow WHERE item_id = $item and ts BETWEEN FROM_UNIXTIME($from) AND FROM_UNIXTIME($to) LIMIT 10000";
        $data = $this->models->get_data($sql);
        $spreadsheet = $this->models->reader("./Report/flow.xlsx");
        $activeWorksheet = $spreadsheet->getSheetByName('RawData');

        $i = 3;

        foreach($data as $dt){
            $flow = $dt->flow;
            $volume = $dt->volume;

            $activeWorksheet->setCellValue('B'.$i, '=(((('. $dt->ts + 25200 .'/60)/60)+8)/24)+DATE(1970,1,1)'); //time

            $activeWorksheet->setCellValue('C'.$i, $flow);
            $activeWorksheet->setCellValue('D'.$i, $volume);

            $i++;
        }
    
        $filename = $item .' Flow Report '. date("d M, Y (H-i)",$from) .' - '. date("d M, Y (H-i)", $to);
        $this->models->writer($spreadsheet, $filename);
    }

    public function temperature($from, $to){
        $sql = "SELECT * FROM temperature WHERE ts BETWEEN FROM_UNIXTIME($from) AND FROM_UNIXTIME($to) LIMIT 10000";
        $data = $this->models->get_data($sql);
        $items = $this->models->get_item('-a', 'temperature');
        $spreadsheet = $this->models->reader("./Report/temp.xlsx");
        $activeWorksheet = $spreadsheet->getSheetByName('RawData');
        $i = 3;

        foreach($data as $dt){
            $temp = $dt->temperature;
            
            foreach($items as $item){
                if ($item->id == $dt->item_id) {
                    $name = $item->name; 
                    break;
                }
            }

            $activeWorksheet->setCellValue('B'.$i, '=(((('. $dt->ts + 25200 .'/60)/60)+8)/24)+DATE(1970,1,1)'); //time

            $activeWorksheet->setCellValue('C'.$i, $temp);
            $activeWorksheet->setCellValue('D'.$i, $name);

            $i++;
        }

        $i = 3;
        foreach($items as $item){
            $activeWorksheet->setCellValue('G'.$i, $item->name);
            $i++;
        }
    
        $filename = 'Temperature Report ';
        $this->models->writer($spreadsheet, $filename);
    }

    public function dry_contact($from, $to){
        $sql = "SELECT * FROM status WHERE ts BETWEEN FROM_UNIXTIME($from) AND FROM_UNIXTIME($to) LIMIT 10000";
        $data = $this->models->get_data($sql);
        $items = $this->models->get_item('-a', 'digital');
        $spreadsheet = $this->models->reader("./Report/drycon.xlsx");
        $activeWorksheet = $spreadsheet->getSheetByName('RawData');
        $i = 3;

        foreach($data as $dt){
            $status = $dt->status;
            
            foreach($items as $item){
                if ($item->id == $dt->item_id) {
                    $name = $item->name; 
                    break;
                }
            }

            $activeWorksheet->setCellValue('B'.$i, '=(((('. $dt->ts + 25200 .'/60)/60)+8)/24)+DATE(1970,1,1)'); //time

            $activeWorksheet->setCellValue('C'.$i, $status);
            $activeWorksheet->setCellValue('D'.$i, $name);

            $i++;
        }

        $i = 3;
        foreach($items as $item){
            $activeWorksheet->setCellValue('G'.$i, $item->name);
            $i++;
        }
    
        $filename = 'Dry Contact Report '. date("d M, Y (H-i)",$from) .' - '. date("d M, Y (H-i)", $to);
        $this->models->writer($spreadsheet, $filename);
    }
}
