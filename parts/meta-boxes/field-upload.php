<?php
/*
Title: Upload Performance Report Excel
Post Type: performance_excel
Order: 110
Meta Box: true
Collapse: false
*/

  // Any field with the scope set to the field name of the upload field will be treated as related
  // data to the upload. Below we see we are setting the post_status and post_title, where the
  // post_status is pulled dynamically on page load, hence the current status of the content is
  // applied. Have fun! ;)
  //
  // NOTE: If the post_status of an attachment is anything but inherit or private it will NOT be
  // shown on the Media page in the admin, but it is in the database and can be found using query_posts
  // or get_posts or get_post etc....
?>

<?php

$attachmentID = get_post_meta($post->ID, 'performance_excel_file', TRUE);
$inputFileName = get_attached_file($attachmentID);


    piklist('field', array(
        'type' => 'file'
        ,'field' => 'performance_excel_file'
        ,'scope' => 'post_meta'
        ,'label' => __('Add Excel File','piklist-demo')
        ,'options' => array(
          'modal_title' => __('Add Excel File','piklist-demo')
          ,'button' => __('Select File','piklist-demo')
    )
    ));

    // echo $inputFileName. '<br>';

/************************************************************************************/
/************************************************************************************/
$sheetname      = 'Equity';

$report = new Excel2Mysql();



$excelRow = $report->get_records_from_excel($sheetname, $inputFileName);

echo 'Excel- <pre>';
print_r($excelRow);
echo '<pre>';
/**********************************************************************************/


$dbRow = $report->fetch_records_from_db($excelRow);

echo 'DB- <pre>';
print_r($excelRow);
echo '<pre>';

/**********************************************************************************/
$duplicate = $report->get_duplicate_records_from_db($excelRow);







