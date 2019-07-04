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
    public function create($filename, $callback)
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
     * Stores the Excel spreadsheet to disk
     *
     * @param $disk
     * @return $this
     */
    public function store($disk)
    {
        $sheetIndex = $this->spreadsheet->getIndex(
            $this->spreadsheet->getSheetByName('Worksheet')
        );
        $this->spreadsheet->removeSheetByIndex($sheetIndex);
        
        if (!File::isDirectory(storage_path('spreadsheetTmp/'))) {
            File::makeDirectory(storage_path('spreadsheetTmp/'));
        }

        $tmpFile = storage_path('spreadsheetTmp/' . str_random(20));
        $this->writer->save($tmpFile);

        Storage::disk($disk)->put($this->filename, file_get_contents($tmpFile));

        File::delete($tmpFile);
        File::deleteDirectory(storage_path('spreadsheetTmp/'));

        return $this;
    }


    /**
     * Load Excel file
     * @param $filename
     * @param $cb
     * @return $this
     */
    public function load($filename, $callback)
    {
        $this->filename = $filename;
        $fileType = ucfirst(File::extension($filename));

        $this->reader = IOFactory::createReader($fileType);
        $this->reader = $this->reader->load($filename);
        $callback($this->reader);

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
