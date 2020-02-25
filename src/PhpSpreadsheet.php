<?php

namespace OSMAviation\Spreadsheet;

use Illuminate\Support\Facades\File;
use Illuminate\Support\Facades\Storage;
use Illuminate\Support\Str;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Writer\Xls as WriterXls;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx as WriterXlsx;
use PhpOffice\PhpSpreadsheet\Spreadsheet as VendorSpreadsheet;

class PhpSpreadsheet
{
    /**
     * Excel writer
     *
     * @var PhpOffice\PhpSpreadsheet\Writer\Xls | PhpOffice\PhpSpreadsheet\Writer\Xlsx
     */
    protected $writer;

    /**
     * Excel reader
     *
     * @var PhpOffice\PhpSpreadsheet\Writer\Xls | PhpOffice\PhpSpreadsheet\Writer\Xlsx
     */
    protected $reader;

    /**
     * Spreadsheet
     *
     * @var PhpOffice\PhpSpreadsheet\Spreadsheet
     */
    protected $spreadsheet;

    /**
     * Name of file to export
     *
     * @var string
     */
    protected $filename;

    /**
     * Set the instances needed
     */
    public function __construct()
    {
        $this->spreadsheet = new VendorSpreadsheet;
    }

    /**
     * Create a new XLS sheet
     *
     * @param $filename
     * @param $callback
     * @return $this
     */
    public function create($filename, $callback = null)
    {
        $this->filename = $filename;
        if (File::extension($filename) === 'xls') {
            $this->writer = new WriterXls($this->spreadsheet);
        } else {
            $this->writer = new WriterXlsx($this->spreadsheet);
        }
        $callback(new Spreadsheet($this->spreadsheet));
        return $this;
    }


    /**
     * Load Excel file
     * @param $filename
     * @param $cb
     * @return $this
     */
    public function load($filename, $disk = null, $callback = null)
    {

        if (!$callback) {
            $callback = $disk;
            $disk = null;
        }

        $this->filename = $filename;
        $fileType = ucfirst(File::extension($filename));

        if ($disk) {
            if (!File::isDirectory(storage_path('spreadsheetTmp/'))) {
                File::makeDirectory(storage_path('spreadsheetTmp/'));
            }

            $tmpFile = storage_path('spreadsheetTmp/' . str_random(20));
            $content = Storage::disk($disk)->get($filename);
            File::put($tmpFile, $content);
            $filename = $tmpFile;
        }

        $this->spreadsheet = IOFactory::load($filename);
        if ($fileType === 'Xlsx') {
            $this->spreadsheet->setActiveSheetIndex(0);
        }
        $this->writer = IOFactory::createWriter($this->spreadsheet, $fileType);

        File::deleteDirectory(storage_path('spreadsheetTmp/'));

        $callback(new Spreadsheet($this->spreadsheet));

        return $this;
    }

    /**
     * Load Excel file
     * @param $filename
     * @param $cb
     * @return $this
     */
    public function read($filename, $callback)
    {
        $this->filename = $filename;
        $fileType = ucfirst(File::extension($filename));

        $this->reader = IOFactory::createReader($fileType);
        $this->reader = $this->reader->load($filename);
        $callback($this->reader);

        return $this;
    }

    /**
     * Stores the Excel spreadsheet to disk
     *
     * @param $disk
     * @return $this
     */
    public function store($disk, $filename = null)
    {
        if (!$filename) {
            $filename = $this->filename;
        }

        if ($sheet = $this->spreadsheet->getSheetByName('Worksheet')) {
            $sheetIndex = $this->spreadsheet->getIndex($sheet);
            $this->spreadsheet->removeSheetByIndex($sheetIndex);
        }

        // we will temporarily store the file on the local
        // filesystem before moving it to the selected disk

        if (!File::isDirectory(storage_path('spreadsheetTmp/'))) {
            File::makeDirectory(storage_path('spreadsheetTmp/'));
        }
        $tmpFile = storage_path('spreadsheetTmp/' . str_random(20));
        
        if (count($this->spreadsheet->getAllSheets()) > 0) {
            $this->spreadsheet->setActiveSheetIndex(0);
        }

        $this->writer->save($tmpFile);
        $file = Storage::disk($disk)->put($filename, file_get_contents($tmpFile));

        File::delete($tmpFile);
        File::deleteDirectory(storage_path('spreadsheetTmp/'));

        $this->filename = $filename;

        return $this;
    }

    /**
     * Gets the fully qualified spreadsheet path
     *
     * @return string
     */
    public function getFilename()
    {
        return $this->filename;
    }

    /**
     * Associate header keys with array values
     * First row will be as header
     * @param $results
     * @return array
     */
    public function associateKeysWithValues($results)
    {
        $headings = array_shift($results);

        // convert keys to snake case and delete extra symbols except spaces, letters, numbers
        $headings = array_map(function ($key) {
            $key = strtolower($key);
            $key = preg_replace("/[^a-zA-Z0-9\s]/", "", $key);
            return Str::snake($key);
        }, $headings);

        // combine keys with values
        array_walk(
            $results,
            function (&$row) use ($headings) {
                $row = array_combine($headings, $row);
            }
        );

        $data = [
            'headings' => $headings,
            'data' => $results,
        ];

        return $data;
    }
}
