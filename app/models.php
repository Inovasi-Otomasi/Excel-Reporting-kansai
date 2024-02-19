<?php

require './vendor/autoload.php';
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

class config {
    private function connection() {
        $servername = "192.168.0.161:3306";
        $username = "admin";
        $password = "adminiot123";
        $dbname = "utility_data";

        $conn = new mysqli($servername, $username, $password, $dbname);

        if ($conn->connect_error) {
            die("Connection failed: " . $conn->connect_error);
        }

        return $conn;
    }
    
    public function get_item($item = "", $type = "") {
        
        $servername = "192.168.0.161:3306";
        $username = "admin";
        $password = "adminiot123";
        $dbname = "utility_config";
        
        $conn = new mysqli($servername, $username, $password, $dbname);
        
        if ($conn->connect_error) {
            die("Connection failed: " . $conn->connect_error);
        }

        if($item == "-a"){
            $sql = "SELECT * FROM items WHERE type = '$type'";
        } else {
            $sql = "SELECT id FROM items WHERE name = '$item'";
        }

        $result = $conn->query($sql);
        $conn->close();
        if ($result->num_rows > 0) {
            $data = [];
            $i = 0;
            while($row = $result->fetch_assoc()) {
                $data[$i] = json_decode(json_encode($row), FALSE);;
                $i++;
            }
            return $data;
        } else {
            die("Unlisted Item!");
        }
    } 

    public function get_data($sql){
        $conn = $this->connection();

        $result = $conn->query($sql);
        $conn->close();
        
        if ($result->num_rows > 0) {
            $data = [];
            $i = 0;
            while($row = $result->fetch_assoc()) {
                $dateTime = new DateTime($row["ts"]); 
                $row["ts"] = $dateTime->format('U');
                $data[$i] = json_decode(json_encode($row), FALSE);;
                $i++;
            }
            return $data;
        } else {
            die("No Data Available!");
        }
    }

    public function reader($file_path){
        $reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
        $reader->setIncludeCharts(true);

        $ss = $reader->load($file_path);
        return $ss;
    }

    public function writer($ss, $filename){
        $writer = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($ss, 'Xlsx');
        $writer->setIncludeCharts(true); //penting
        $writer->setPreCalculateFormulas(false); //paling penting
        
        
        header('Content-Type: application/vnd.ms-excel');
        header('Content-Disposition: attachment;filename="'.$filename.'.xlsx"');
        header('Cache-Control: max-age=0');
        $writer->save('php://output'); 
    }

    
}