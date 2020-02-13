<?php

namespace OSMAviation\Spreadsheet;

use Illuminate\Support\ServiceProvider;
use Cache\Adapter\Predis\PredisCachePool;
use Cache\Adapter\Redis\RedisCachePool;
use Cache\Bridge\SimpleCache\SimpleCacheBridge;
use PhpOffice\PhpSpreadsheet\Settings as SpreadsheetSettings;

class SpreadsheetServiceProvider extends ServiceProvider
{
    /**
     * Bind class to the service container
     *
     * @return void
     */
    public function register()
    {
        $this->registerCache();

        $this->app->bind('spreadsheet', function ($app) {
            return new PhpSpreadsheet();
        });

        $this->app->alias('spreadsheet', PhpSpreadsheet::class);
    }

    public function boot()
    {
        $this->registerCache();

        if ($bridge = $this->app['spreadsheet.cacheBridge']) {
            SpreadsheetSettings::setCache($bridge);
        }
    }

    protected function registerCache()
    {
        $this->app->bind('spreadsheet.pool', function ($app) {
            if (!isset($app['redis.connection'])) {
                return null;
            }

            $client = $app['redis.connection']->client();

            if ($client instanceof \Redis) {
                return new RedisCachePool($client);
            }

            if ($client instanceof \Predis\Client) {
                return new PredisCachePool($client);
            }

            return null;
        });

        $this->app->bind('spreadsheet.cacheBridge', function ($app) {
            if (!$app['spreadsheet.pool']) {
                return null;
            }
            return new SimpleCacheBridge($app['spreadsheet.pool']);
        });
    }
}
