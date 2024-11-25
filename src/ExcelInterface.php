<?php

declare(strict_types=1);

namespace Firezihai\Excel;

interface ExcelInterface
{
    /**
     * 解析excel.
     * @param ExcelDtoInterface $dto
     */
    public function parse(string $filename, string $dto);

    /**
     * 导出excel.
     */
    public function export(string $dto, array $data);
}
