<?php

namespace Knackline\ExcelTo;

use Illuminate\Support\ServiceProvider as BaseServiceProvider;

class ServiceProvider extends BaseServiceProvider
{
    public function register()
    {
        $this->app->singleton(ExcelTo::class, function ($app) {
            return new ExcelTo();
        });
    }

    public function boot()
    {
        $this->publishes([
            __DIR__ . '/../config/excel-to.php' => config_path('excel-to.php'),
        ], 'config');
    }
}
