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
    private $dbColumns = array('stockID', 'stockName', 'action', 'entryDate', 'entryPrice', 'targetPrice', 'stopLoss', 'exitDate', 'exitPrice');
    private $postType = 'performance_report';
    //private $table = 'report_performance';

    function __construct() {

    }



    /*  Get all records form table  */
    public function fetch_records_from_db($excelRow)
    {
        global $wpdb;
        $dbColumns = $this->dbColumns;
        $postType = $this->postType;
        $dbColumns = implode(', ', $dbColumns);

        $query = "SELECT meta_key, meta_value
          FROM $wpdb->posts
          LEFT JOIN $wpdb->postmeta
          ON ($wpdb->posts.ID = $wpdb->postmeta.post_id)
          WHERE $wpdb->posts.post_type = '".$postType. "'
          AND $wpdb->posts.post_name = '".$excelRow['stockID']."'
          AND $wpdb->posts.ID = $wpdb->postmeta.post_id
          ORDER BY post_date DESC";

        $mysqlDataArray = $wpdb->get_results($query);

        //print_r($mysqlDataArray);
        
        $dbColumns = explode(', ', $dbColumns);
        
        foreach($mysqlDataArray as $rowArray) {
            //print_r($rowArray);
            //echo $rowArray->meta_key;
            //echo $rowArray->meta_value;

            foreach($dbColumns as $col)
            {
                if(($rowArray->meta_key) == $col)
                {
                    $sanitizedDbMetaRow[$rowArray->meta_key] = $rowArray->meta_value;
            
                }             
            }            
        }
        return $sanitizedDbMetaRow;
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

    public function get_records_from_excel($sheetname, $inputFileName)
    {
        $inputFileType = 'Excel2007';
        $dbColumns = $this->dbColumns;  // Get value in $dbcolumns variable


        /**  Create an Instance of our Read Filter, passing in the cell range  **/
        $filterSubset = new Custom_Filter_For_Excel(2,50,range('A','I'));

        $objReader = PHPExcel_IOFactory::createReader($inputFileType);
        $objReader->setLoadSheetsOnly($sheetname);
        $objReader->setReadFilter($filterSubset);


        // if excel file is added then
        if (!empty($inputFileName))
        {
            $objPHPExcel = $objReader->load($inputFileName);
            $excelSheetDataArray = $objPHPExcel->getActiveSheet()->toArray(null,true,true,true);
            //echo '<pre>';
            //print_r($excelSheetDataArray);
            //echo '</pre>';

            foreach ($excelSheetDataArray as $excelRow) {
                $excelSheetValues = array_values($excelRow);
                //print_r($excelSheetValues);

                $excelRowSanitizedArray = array_combine($dbColumns, $excelSheetValues);

               if(!empty($excelRowSanitizedArray['stockID'])) { // return array which have a 'stockID'
                    return $excelRowSanitizedArray;
               }
            }
        }
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
