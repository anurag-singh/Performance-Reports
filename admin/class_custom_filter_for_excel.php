<?php
/**
 * The file that defines filters for excel sheet's data
 *
 * A class definition that includes attributes and functions used across both the
 * public-facing side of the site and the admin area.
 *
 * @link       https://www.facebook.com/anuragsingh.me
 * @since      1.0.0
 *
 * @package    As_Performance_Reports
 * @subpackage As_Performance_Reports/includes
 */

    class Custom_Filter_For_Excel implements PHPExcel_Reader_IReadFilter
    {
        private $_startRow = 0;
        private $_endRow   = 0;
        private $_columns  = array();

        /**  Get the list of rows and columns to read  */
        public function __construct($startRow, $endRow, $columns) {
            $this->_startRow = $startRow;
            $this->_endRow   = $endRow;
            $this->_columns  = $columns;
        }
        public function readCell($column, $row, $worksheetName = '') {
                //  Only read the rows and columns that were configured
                if ($row >= $this->_startRow && $row <= $this->_endRow) {
                    if (in_array($column,$this->_columns)) {
                        return true;
                    }
                }
                return false;
            }
    }
?>
