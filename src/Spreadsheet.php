<?php

namespace OSMAviation\Spreadsheet;

use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use PhpOffice\PhpSpreadsheet\Spreadsheet as VendorSpreadsheet;

class Spreadsheet
{
    /**
     * Spreadsheet
     *
     * @var VendorSpreadsheet
     */
    protected $spreadsheet;

    /**
     * Hook it up
     *
     * @param VendorSpreadsheet
     */
    public function __construct(VendorSpreadsheet $spreadsheet)
    {
        $this->spreadsheet = $spreadsheet;
    }

    /**
     * Creates a new sheet / edits an existing one
     *
     * @param string $name
     * @param \Closure $callback
     * @return Spreadsheet
     */
    public function sheet($name, \Closure $callback)
    {
        $sheet = $this->spreadsheet->getSheetByName($name);
        if (!$sheet) {
            $sheet = new Worksheet($this->spreadsheet, $name);
            $this->spreadsheet->addSheet($sheet);
        }
        $callback($sheet);
        return $this;
    }

    /**
     * Gets the current spreadsheet
     *
     * @return \PhpOffice\PhpSpreadsheet\Spreadsheet
     */
    public function getSpreadsheet()
    {
        return $this->spreadsheet;
    }
}
