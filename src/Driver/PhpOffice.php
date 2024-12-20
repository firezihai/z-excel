<?php

declare(strict_types=1);

namespace Firezihai\Excel\Driver;

use Firezihai\Excel\AbstractExcel;
use Firezihai\Excel\ExcelInterface;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Color;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Style\NumberFormat;

class PhpOffice extends AbstractExcel implements ExcelInterface
{
    public function export(string $dto, array $data)
    {
        $this->dto = $dto;
        $annotationMeta = $this->parseDtoAnnotation($dto);
        // 创建一个新的Spreadsheet对象，用于生成Excel文件
        $spreadsheet = new Spreadsheet();

        $activeSheet = $spreadsheet->getActiveSheet();

        // 设置全局公共样式
        if (! empty($annotationMeta['height'])) {
            $activeSheet->getDefaultRowDimension()->setRowHeight($annotationMeta['height']);
        }
        if (! empty($annotationMeta['align'])) {
            $spreadsheet->getDefaultStyle()->getAlignment()
                ->setHorizontal($annotationMeta['align'])->setVertical($annotationMeta['align']);
        }
        $excelHeader = $this->sortHeader($annotationMeta['header']);

        // 填充表头
        $this->fillHeader($spreadsheet, $annotationMeta, $excelHeader);

        try {
            $row = 2;
            foreach ($data as $item) {
                $column = 0;
                foreach ($excelHeader as $header) {
                    $coordinate = $this->getColumnIndex($column) . $row;

                    $activeSheet->setCellValue($coordinate, $this->getFieldValue($item, $header));
                    ++$column;
                }
                ++$row;
            }
        } catch (\RuntimeException $e) {
        }

        $writer = IOFactory::createWriter($spreadsheet, 'Xlsx');
        $filePath = $this->getTmpDir() . time() . '.xlsx';

        $writer->save($filePath);

        $spreadsheet->disconnectWorksheets();

        return $filePath;
    }

    /**
     * 填充表头.
     */
    public function fillHeader(Spreadsheet $spreadsheet, array $annotationMeta, array $newHeader)
    {
        $activeSheet = $spreadsheet->getActiveSheet();

        $headerIndex = 0;
        foreach ($newHeader as $item) {
            $headerColumn = $this->getColumnIndex($headerIndex) . '1';
            $activeSheet->setCellValue($headerColumn, $item['name']);

            $columnDimension = $activeSheet->getColumnDimension($headerColumn[0]);
            if (! empty($item['width'])) {
                $columnDimension->setWidth((float) $item['width']);
            } else {
                $columnDimension->setAutoSize(true);
            }
            if (! empty($item['align'])) {
                $activeSheet->getStyle($headerColumn)->getAlignment()->setHorizontal($item['align'])->setVertical($item['align']);
            }
            if (! empty($item['color'])) {
                $activeSheet->getStyle($headerColumn)->getFont()->setColor(new Color(str_replace('#', '', $item['color'])));
            }

            if ($item['headBgColor'] ?? '') {
                $activeSheet->getStyle($headerColumn)->getFill()
                    ->setFillType(Fill::FILL_SOLID)
                    ->getStartColor()->setARGB(str_replace('#', '', $item['headBgColor']));
            }
            ++$headerIndex;
        }
    }

    public function parse(string $filename, string $dto)
    {
        $reader = IOFactory::createReader(IOFactory::identify($filename));
        $reader->setReadDataOnly(true);
        $sheet = $reader->load($filename);

        $annotationMate = $this->parseDtoAnnotation($dto);
        $header = $annotationMate['header'];
        $fieldMap = [];
        $type = $annotationMate['type'] ?? 'name';
        foreach ($header as $value) {
            $fieldMap[$value[$type]] = $value['field'];
        }
        $endCell = $header ? $this->getColumnIndex(count($header)) : null;
        $data = [];
        $i = 0;
        $fieldCell = [];
        foreach ($sheet->getActiveSheet()->getRowIterator(1, 1) as $row) {
            foreach ($row->getCellIterator('A', $endCell) as $index => $item) {
                $value = $item->getValue();
                $fieldKey = $annotationMate['type'] == 'index' ? ord($index) - 65 : $value;
                // 空跳过
                if (empty($value) || ! isset($fieldMap[$fieldKey])) {
                    continue;
                }
                $fieldCell[ord($index)] = $fieldMap[$fieldKey];
            }
            ++$i;
        }
        foreach ($sheet->getActiveSheet()->getRowIterator(2) as $row) {
            $temp = [];
            foreach ($row->getCellIterator('A', $endCell) as $index => $item) {
                $value = $item->getFormattedValue();
                // 转换时间格式
                if (in_array($item->getStyle()->getNumberFormat()->getFormatCode(), [
                    NumberFormat::FORMAT_DATE_DATETIME,
                    NumberFormat::FORMAT_DATE_DDMMYYYY,
                    NumberFormat::FORMAT_DATE_DMYMINUS,
                    NumberFormat::FORMAT_DATE_DMYSLASH,
                ])) {
                    $value = date('Y-m-d H:i:s', strtotime($value));
                }
                $key = $fieldCell[ord($index)] ?? '';
                if ($key) {
                    $temp[$key] = $value;
                }
            }
            $data[] = $temp;
        }

        return $data;
    }
}
