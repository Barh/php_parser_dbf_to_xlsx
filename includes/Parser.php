<?php

    use XBase\Table;

    // PHP XBase (a simple parser for *.dbf files)
    // https://github.com/hisamu/php-xbase
    require_once 'XBase'.DIRECTORY_SEPARATOR.'Table.php';
    require_once 'XBase'.DIRECTORY_SEPARATOR.'Column.php';
    require_once 'XBase'.DIRECTORY_SEPARATOR.'Record.php';

    // Classes from Yii framework (required for AlxdExportXLSX)
    // https://github.com/yiisoft/yii/blob/1.1.17/framework/base/CException.php
    // https://github.com/yiisoft/yii/blob/1.1.17/framework/base/CHttpException.php
    // https://github.com/yiisoft/yii/blob/1.1.17/framework/utils/CFileHelper.php
    require_once 'yii_classes/CException.php';
    require_once 'yii_classes/CHttpException.php';
    require_once 'yii_classes/CFileHelper.php';

    // AlxdExportXLSX (Class for export data to Microsoft Excel in format XLSX)
    // required program zip in UNIX
    // https://github.com/Alxdhere/AlxdExportXLSX
    require_once 'AlxdExportXLSX.php';


    /**
     * Class Parser
     */
    class Parser
    {
        /**
         * Init variables
         * @var string
         */
        private $export_to, $stat_from, $stat_to, $archive_to, $archive_from, $log_to;

        /**
         * Default values (charset, columns)
         * @var string
         */
        private $charset_from = 'cp866';
        private $columns      = array('newnum', 'namep');

        /**
         * Errors text
         * @var array
         */
        private $errors = array(
            'archive_from'    => 'Невозможно получить архив по ссылке: ',
            'archive_no_zip'  => 'Файл, скачанный по ссылке, не zip-архив: ',
            'sequence_stat'   => 'Перед выполнением getStat необходимо выполнить getArchive.',
            'sequence_export' => 'Перед выполнением export необходимо выполнить getStat.',
            'stat_from'       => 'В архиве нет необходимого файла %s:',
            'stat_from_empty' => 'Файл %s пуст.',
            'export_to'       => 'Не удалось создать файл: ',
            'clear'           => 'Невозможно удалить временный файл: ',
        );

        /**
         * Success text
         * @var array
         */
        private $success = 'Файл успешно экспортирован';

        /**
         * Set charset
         * @param string $from encode from
         * @return $this
         */
        public function setCharset($charset_from)
        {
            $this->charset_from = $charset_from;

            return $this;
        }

        /**
         * Set columns
         * @param array $columns columns
         * @return $this
         */
        public function setColumns($columns)
        {
            $this->columns = $columns;

            return $this;
        }

        /**
         * Set logger file name
         * @param string $log_to log to (filename)
         * @return $this
         */
        public function setLogger($log_to)
        {
            // set log filename
            $this->log_to = $log_to;

            // no exists file
            if (!is_file($this->log_to)) {

                // error create file
                if (@file_put_contents($this->log_to, '', LOCK_EX) === false) {
                    echo $this->getMessage();
                    exit;
                }
            }

            return $this;
        }

        /**
         * Get message last error
         * @return string
         */
        private function getMessage()
        {
            $message = error_get_last();
            return 'Line: '.$message['line'].', '.$message['message'];
        }

        /**
         * Insert to log
         * @param bool|string $message message
         * @return bool
         */
        private function insertToLog($message = false)
        {
            // get message
            if (!$message) {
                $message = $this->getMessage();
            }

            // print message
            echo $message;

            // logger enabled append log file
            if ($this->log_to && @file_put_contents($this->log_to, date('Y-m-d H:i:s | ', time()).$message.PHP_EOL, FILE_APPEND | LOCK_EX) === false) {
                // error append file
                echo $this->getMessage();
                exit;
            }
        }

        /**
         * Get archive
         * @param string $archive_from archive from (filename)
         * @param string $archive_to archive to filename)
         * @return $this
         */
        public function getArchive($archive_from, $archive_to)
        {
            // set archive values
            $this->archive_from = $archive_from;
            $this->archive_to   = $archive_to;

            // get headers
            @$h = get_headers($this->archive_from, 1);

            // 200 OK (successful HTTP requests)
            if ( $h && strstr($h[0], '200') !== FALSE ) {
                // error create archive file
                if (@file_put_contents($this->archive_to, file_get_contents($this->archive_from), LOCK_EX) === false) {
                    $this->insertToLog();
                    exit;
                }

                return $this;
            } else { // no successful
                $this->insertToLog($this->errors['archive_from'].$this->archive_from);
                exit;
            }
        }

        /**
         * Get stat
         * @param string $stat_from stat from (filename)
         * @param string $stat_to stat to (filename)
         * @return $this|bool
         */
        public function getStat($stat_from, $stat_to)
        {
            // error sequence
            if (!$this->archive_to) {
                $this->insertToLog($this->errors['sequence_stat']);
                exit;
            }

            // set stat file name
            $this->stat_from = $stat_from;
            $this->stat_to   = $stat_to;

            // open archive
            $zip = zip_open($this->archive_to);

            // zip incorrect
            if ( !is_resource($zip) )
            {
                $this->insertToLog($this->errors['archive_no_zip'].$this->archive_from);
                exit;
            }

            // looking for the required file
            $is_dbf = false;
            while (!$is_dbf && $entry = zip_read($zip)) {
                if (zip_entry_name($entry) == $this->stat_from) {
                    $is_dbf = true;
                }
            }

            // is no required file
            if (!$is_dbf) {
                $this->insertToLog( sprintf($this->errors['stat_from'], $this->stat_from).$this->archive_from);
                exit;
            }

            // open file
            zip_entry_open($zip, $entry, "r");

            // read data file
            $data = zip_entry_read($entry, zip_entry_filesize($entry));

            // is no data in file
            if ( empty($data) )
            {
                $this->insertToLog( sprintf($this->errors['stat_from_empty'], $this->stat_from) );
                exit;
            }

            // error create file
            if (@file_put_contents($this->stat_to, $data, LOCK_EX) === false) {
                $this->insertToLog();
                exit;
            }

            return $this;
        }

        /**
         * Export
         * @param string $export_to export to (filename)
         * @return $this
         */
        public function export($export_to)
        {
            // error sequence
            if (!$this->stat_to) {
                $this->insertToLog($this->errors['sequence_export']);
                exit;
            }

            // set export file name
            $this->export_to = $export_to;

            // open stat dbase
            try {
                $table = new Table($this->stat_to, null, $this->charset_from);
            } catch (Exception $e) {
                $this->insertToLog( 'XBase (Table, Line: '.$e->getLine().'): '.$e->getMessage() );
                exit;
            }

            // open stat dbase
            try {
                // create export file
                $export = new AlxdExportXLSX('export.xlsx', count($this->columns), 2);
                $export->openWriter();
            } catch (Exception $e) {
                $this->insertToLog( 'AlxdExportXLSX (Line: '.$e->getLine().'): '.$e->getMessage() );
            }

            // each rows
            while ($record = $table->nextRecord()) {
                // open row
                $export->resetRow();
                $export->openRow();

                // fill data
                foreach ($this->columns as $v) {
                    $export->appendCellString($record->{$v});
                }

                // close row
                $export->closeRow();
                $export->flushRow();
            }

            // close file
            $export->closeWriter();

            // zip (get xlsx)
            $export->zip();

            $export->getZipFullFileName();

            // error create export file
            if (@rename($export->getZipFullFileName(), $this->export_to) === false) {
                $this->insertToLog($this->errors['export_to'].$this->export_to);
                exit;
            }

            // clear temp files
            $this->clearTempFiles();

            // success
            echo $this->success;

            return $this;
        }

        /**
         * Clear temp files
         */
        private function clearTempFiles()
        {
            // each archive and stat files
            foreach (array('archive', 'stat') as $v) {
                // is
                if (is_file($this->{$v.'_to'})) {
                    // delete
                    if (unlink($this->{$v.'_to'}) === false) {
                        // error
                        $this->insertToLog($this->errors['clear'].$this->{$v.'_to'});
                        exit;
                    }
                }
            }
        }
    }