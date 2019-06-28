<?php

namespace OSMAviation\Spreadsheet\Facades;

use Illuminate\Support\Facades\Facade;

class Spreadsheet extends Facade
{
    public static function getFacadeAccessor()
    {
        return 'spreadsheet';
    }
}