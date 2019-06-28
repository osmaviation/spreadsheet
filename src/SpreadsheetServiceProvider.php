<?php

namespace OSMAviation\Spreadsheet;

use Illuminate\Support\ServiceProvider;

class SpreadsheetServiceProvider extends ServiceProvider
{
    /**
     * Bind class to the service container
     *
     * @return void
     */
    public function register()
    {
        $this->app->bind('spreadsheet', function ($app) {
            return new PhpSpreadsheet();
        });

        $this->app->alias('spreadsheet', PhpSpreadsheet::class);
    }
}
