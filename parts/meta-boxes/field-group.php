<?php
/*
Title: Field Groups
Post Type: performance_report
Order: 80
*/


  piklist('field', array(
    'type' => 'group'
    //,'field' => 'performance_report'
    ,'label' => 'Performance Report (Grouped)'
    ,'list' => false
    ,'fields' => array(

      array(
        'type' => 'text'
        ,'field' => 'uni_code'
        ,'label' => 'Unique Code'
        ,'columns' => 6
        ,'attributes' => array(
          'placeholder' => 'Unique Code'
        )
      )

      ,array(
        'type' => 'text'
        ,'field' => 'stocks'
        ,'label' => 'Stocks'
        ,'columns' => 6
        ,'attributes' => array(
          'placeholder' => 'Stocks'
        )
      )
        ,array(
        'type' => 'text'
        ,'field' => 'action'
        ,'label' => 'Buy/Sell'
        ,'columns' => 6
        ,'attributes' => array(
          'placeholder' => 'Action'
        )
      )

      ,array(
        'type' => 'text'
        ,'field' => 'entry_date'
        ,'label' => 'Entry Date'
        ,'columns' => 6
        ,'attributes' => array(
          'placeholder' => 'Entry Date'
        )
      )
        ,array(
        'type' => 'text'
        ,'field' => 'entry_price'
        ,'label' => 'Entry Price'
        ,'columns' => 6
        ,'attributes' => array(
          'placeholder' => 'Entry Price'
        )
      )

      ,array(
        'type' => 'text'
        ,'field' => 'target_price'
        ,'label' => 'Target Price'
        ,'columns' => 6
        ,'attributes' => array(
          'placeholder' => 'Target Price'
        )
      )
        ,array(
        'type' => 'text'
        ,'field' => 'stop_loss'
        ,'label' => 'Stop Loss'
        ,'columns' => 6
        ,'attributes' => array(
          'placeholder' => 'Stop Loss'
        )
      )

      ,array(
        'type' => 'text'
        ,'field' => 'exit_date'
        ,'label' => 'Exit Date'
        ,'columns' => 6
        ,'attributes' => array(
          'placeholder' => 'Exit Date'
        )
      )
      ,array(
        'type' => 'text'
        ,'field' => 'exit_price'
        ,'label' => 'Exit Price'
        ,'columns' => 6
        ,'attributes' => array(
          'placeholder' => 'Exit Price'
        )
      )
    )
  ));


$uni_code = get_post_meta($post->ID, 'uni_code', true);
$stocks = get_post_meta($post->ID, 'stocks', true);
$action = get_post_meta($post->ID, 'action', true);
$entry_date = get_post_meta($post->ID, 'entry_date', true);
$entry_price = get_post_meta($post->ID, 'entry_price', true);
$target_price = get_post_meta($post->ID, 'target_price', true);
$stop_loss = get_post_meta($post->ID, 'stop_loss', true);
$exit_date = get_post_meta($post->ID, 'exit_date', true);
$exit_price = get_post_meta($post->ID, 'exit_price', true);

echo $uni_code.'<br>';
echo $stocks.'<br>';
echo $action.'<br>';
echo $entry_date.'<br>';
echo $entry_price.'<br>';
echo $stop_loss.'<br>';
echo $exit_date.'<br>';
echo $exit_price.'<br>';


?>
