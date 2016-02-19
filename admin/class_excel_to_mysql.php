<?php
/** Error reporting */
error_reporting(E_ALL);
ini_set('display_errors', TRUE);
ini_set('display_startup_errors', TRUE);
date_default_timezone_set('Asia/Dili');
/** Error reporting */


class Excel2Mysql extends Custom_Filter_For_Excel
{
    public $conn;
    private $excelSheetDataArray;
    private $table;
   // private $duplicateRecordsFoundInDb;




    /*  Create and select Database  */
    private function create_and_select_db($db)
    {
        $createDB = 'CREATE DATABASE IF NOT EXISTS '.$db;

        if ($this->conn->query($createDB) === TRUE)
        {
            echo 'Database "' .$db. '" created successfully. <br>';
            $this->conn->select_db($db);
        }
        else
        {
            echo 'Found error while creating database - ' . $this->conn->error .'<br>';
        }
    }


    /*  Create Table  */
    public function create_table($table)
    {
        $createTable = 'CREATE TABLE IF NOT EXISTS '. $table .' (
            ID INT(6) UNSIGNED AUTO_INCREMENT PRIMARY KEY,
            stockID VARCHAR(12) NOT NULL,
            stockName VARCHAR(50) NOT NULL,
            action VARCHAR(12) NOT NULL,
            entryDate VARCHAR(12) NOT NULL,
            exitDate VARCHAR(12) NOT NULL,
            entryPrice VARCHAR(12) NOT NULL,
            exitPrice VARCHAR(12) NOT NULL,
            targetPrice VARCHAR(12) NOT NULL,
            stopLoss VARCHAR(12) NOT NULL,
            callStartDate TIMESTAMP
        )';

        if ($this->conn->query($createTable) === TRUE)
        {
            echo 'Table "' .$table. '" created successfully.<br>';
        }
        else
        {
            echo 'Found Error while creating table - ' .$this->conn->error.'<br>';
        }

    }


    /*  Get all records form table  */
    public function fetch_records_from_db($table, array $columns)
    {
        $columns = implode(', ', $columns);
        $selectDbRecords = "SELECT $columns
                                FROM $table";

        $getAllRecordsFromDB = $this->conn->query($selectDbRecords);

        if ($this->conn->errno)
        {
            die("Fail Select " . $this->conn->error);
        }
        else
        {
            if($getAllRecordsFromDB->num_rows>0) {
                return $getAllRecordsFromDB->fetch_all(MYSQLI_ASSOC);
            }
            else
            {
                echo "Database is empty.";
            }

        }
        //print_r($getAllRecordsFromDB);

    }


    /* Add records in database table    */
    public function insert_records_in_db($table, array $dataToInsert)
    {
        $colNames = implode(', ', array_keys($dataToInsert));

        $i = 0;
        foreach($dataToInsert as $key => $val)
        {
            $strValue[$i] = "'" .$val. "'";
            $i++;
        }

        $colValues = implode(', ', $strValue);

        $insertRecordsInDb = "INSERT INTO $table ($colNames) VALUES ($colValues)";

        if($this->conn->query($insertRecordsInDb) === TRUE)
        {
            echo "<br>Values inserted successfully. <br>";
        }
        else
        {
            echo "Error - " .$this->conn->error;
        }
    }

    public function get_records_from_excel($inputFileName, $inputFileType, $sheetname)
    {
//        /**  Create an Instance of our Read Filter, passing in the cell range  **/
//        $filterSubset = new Custom_Filter_For_Excel(2,50,range('A','I'));
//
//        $objReader = PHPExcel_IOFactory::createReader($inputFileType);
//        $objReader->setLoadSheetsOnly($sheetname);
//        $objReader->setReadFilter($filterSubset);
//        $objPHPExcel = $objReader->load($inputFileName);

         $excelSheetDataArray = $objPHPExcel->getActiveSheet()->toArray(null,true,true,true);
         //return $objPHPExcel->getActiveSheet()->toArray(null,true,true,true);

            echo '<pre>';
            print_r($excelSheetDataArray);
            echo '</pre>';


        //$this->get_duplicate_records_from_db($excelSheetDataArray, $table);

    }


    public function get_duplicate_records_from_db($sheetData, $table, array $columns)
    {
        $colNames = implode(', ', $columns);

        foreach ($sheetData as $excelrow) {
            //echo $excelrow['A'];
            if(!empty($excelrow['A']))
            {

//                echo '<hr>Excel Records - <pre>';
//                print_r($excelrow);
//                echo '</pre>';


            $selectDuplicateRecordsFromDB = "SELECT $colNames FROM $table WHERE stockID = '".$excelrow['A']."'";

                $duplicateRowsFromDB = $this->conn->query($selectDuplicateRecordsFromDB);

                if($duplicateRowsFromDB === FALSE)
                {
                    trigger_error('Wrong SQL : '. $selectDuplicateRecordsFromDB. '<br><b>Error : </b>' .$this->conn->error, E_USER_ERROR);
                }
                else
                {
                    if($duplicateRowsFromDB->num_rows>0) {
                        $dbRow = $duplicateRowsFromDB->fetch_assoc();
                        //print_r($dbRow);


                        $sheetRow = $this->sanitizeExcelRowArrayForDB($columns, $excelrow);

                        $dataToUpdateInDB = array_diff_assoc($sheetRow, $dbRow);


                        if(count($dataToUpdateInDB)>0)
                        {
                            $colNamesValues = $this->convertArrayIntoSyntaxforMysql($sheetRow);

                            $updateColumnsInDB = "UPDATE $table SET $colNamesValues WHERE stockID = '".$excelrow['A']."'";

                            $updateColumnsValuesInDB = $this->conn->query($updateColumnsInDB);

                            if($updateColumnsValuesInDB === FALSE)
                            {
                                trigger_error('Wrong SQL : <span style="color:#f00;">' .$updateColumnsInDB. '</span> <em>Error in</em>' . $this->conn->error, E_USER_ERROR);
                            }
                            else
                            {
                                if(mysqli_affected_rows($this->conn))
                                {
                                    echo "<br>Record updated successfully.<br>";
                                }
                                else {
                                    echo "The Record you want to updated is no longer exists";
                                }
                            }
                        }
                    }
                    else
                    {
                        $insertNewRecordsFromExcelToDb  = "INSERT INTO $table ($colNames) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)";

                            if ($stmt = $this->conn->prepare($insertNewRecordsFromExcelToDb)) {

                                $sanitizedExcelRow = $this->sanitizeExcelRowArrayForDB($columns, $excelrow);


                                $stmt->bind_param("sssssssss", $stockID, $stockName,$action,$entryDate,$entryPrice,$targetPrice,$stopLoss,$exitDate,$exitPrice);

                                extract($sanitizedExcelRow);
                                $stmt->execute();

                                $noOfRowAffected = count($stmt->affected_rows);
                                if($noOfRowAffected>0)
                                {
                                    echo "New records inserted successfully.";
                                }
                                else
                                {
                                    echo "No new record found to update.";
                                }


                                $stmt->close();
                            }
                            else {
                                echo '<br><b style =color:#f00;>Error - </b>' . $conn->error;
                            }
                    }
                }
            }
        }


    }

    protected function sanitizeExcelRowArrayForDB($columns ,$excelrow)
    {
        $values = array_values($excelrow);
        return (array_combine($columns, $values));
    }

    protected function convertArrayIntoSyntaxforMysql($array)
    {
        $i = 0;
        foreach ($array as $col =>$val) {
            $mysqlColsValues[$i] = $col. ' = "' .$val. '"' ;
            $i++;
        }
        return (implode(' AND ', $mysqlColsValues));
    }

}
?>
