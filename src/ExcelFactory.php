<?php

declare(strict_types=1);

namespace Firezihai\Excel;

use Firezihai\Excel\Driver\PhpOffice;
use Firezihai\Excel\Driver\XlsWriter;

class ExcelFactory
{
    protected $drivers = [];

    public function get($name = 'phpOffice'): ExcelInterface
    {
        if (isset($this->drivers[$name]) && $this->drivers[$name] instanceof ExcelInterface) {
            return $this->drivers[$name];
        }
        switch ($name) {
            case 'xlswriter':
                return $this->drivers[$name] = new XlsWriter();
            case 'phpOffice':
                return $this->drivers[$name] = new PhpOffice();
            default:
                throw new \InvalidArgumentException('表格处理驱动 ' . $name . '不存在');
        }
    }
}
