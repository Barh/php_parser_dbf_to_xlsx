<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <title>Parser dbf to xlsx</title>
</head>
<body>
<?php

    // Errors
    ini_set('display_errors', 'on');
    error_reporting(E_ALL); // Записывать(показывать) все ошибки

    // require Parser
    chdir('includes');
    require_once 'Parser.php';

    // show message
    echo '<pre style="border: 1px solid black; color: red; font-size: 2em;">';

    // init Parser
    $parser = new Parser();

    // settings
    $settings = array(
        // log
        'log_to'       => __DIR__.DIRECTORY_SEPARATOR.'log.log',
        // archive
        'archive_from' => 'http://www.cbr.ru/mcirabis/BIK/bik_db_'.date('dmY', time()).'.zip',
        'archive_to'   => __DIR__.DIRECTORY_SEPARATOR.'archive.zip',
        // stat
        'stat_from'    => 'bnkseek.dbf',
        'stat_to'      => __DIR__.DIRECTORY_SEPARATOR.'dbf.dbf',
        // export
        'export_to'    => __DIR__.DIRECTORY_SEPARATOR.'export.xlsx',
        // charset (can not specify)
        'charset_from' => 'cp866',
        // columns (can not specify)
        'columns'      => array('newnum', 'namep'),
    );

    // run Parser
    $parser
        ->setLogger($settings['log_to'])
        ->setCharset($settings['charset_from'])
        ->setColumns($settings['columns'])
        ->getArchive($settings['archive_from'], $settings['archive_to'])
        ->getStat($settings['stat_from'], $settings['stat_to'])
        ->export($settings['export_to']);
    echo '</pre>';
?></body>
</html>